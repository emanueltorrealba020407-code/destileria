import { 
  createUserWithEmailAndPassword, 
  signInWithEmailAndPassword, 
  signOut, 
  onAuthStateChanged,
  signInWithPopup,
  GoogleAuthProvider,
  updatePassword,
  confirmPasswordReset,
  sendPasswordResetEmail,
  type User as FirebaseUser
} from 'firebase/auth';
import { auth } from '../lib/firebase';
import { dbService } from './db';
import { Account, AccountType, AreaPermission } from '../types';

const googleProvider = new GoogleAuthProvider();

export const authService = {
  async loginWithGoogle(desiredType?: AccountType, companyName?: string) {
    const userCredential = await signInWithPopup(auth, googleProvider);
    const firebaseUser = userCredential.user;
    
    const email = (firebaseUser.email || '').trim().toLowerCase();
    let profile = await dbService.getUserProfile(firebaseUser.uid);
    if (profile && profile.type === 'COMPANY') {
      if (profile.linkedCompanyId !== firebaseUser.uid) {
        profile.linkedCompanyId = firebaseUser.uid;
        await dbService.setUserProfile(profile);
      }
      return profile;
    }

    const matchingTeam = email ? await dbService.findTeamMemberByEmail(email) : null;
    if (matchingTeam && matchingTeam.member.status === 'inactive') {
      await signOut(auth);
      throw new Error('Su usuario se encuentra inactivo. Por favor contacte al administrador.');
    }
    
    if (!profile) {
      // Create a profile if it doesn't exist
      const type = matchingTeam ? 'USER' : (desiredType || 'USER');
      const name = firebaseUser.displayName || email.split('@')[0] || 'Usuario Google';
      
      let finalCompanyName = companyName || (type === 'COMPANY' ? name : 'Personal');
      if (matchingTeam) {
        const comp = await dbService.getCompany(matchingTeam.companyId);
        finalCompanyName = comp?.name || 'Empresa Vinculada';
      }

      // STRICT ISOLATION BY EMAIL:
      // A company is ALWAYS its own independent space (firebaseUser.uid)
      // A user is linked to a company ONLY IF invited by that exact email, otherwise own personal space
      const companyId = type === 'COMPANY' ? firebaseUser.uid : (matchingTeam ? matchingTeam.companyId : firebaseUser.uid);
      const role = matchingTeam ? matchingTeam.member.role : 'ADMIN';
      const areas: AreaPermission[] = matchingTeam ? matchingTeam.member.areas : ['INVENTARIO', 'PRODUCCION', 'REPORTES', 'CONFIGURACION'];

      profile = {
        id: firebaseUser.uid,
        type,
        name,
        email,
        companyName: finalCompanyName,
        linkedCompanyId: companyId,
        role,
        areas
      };
      
      await dbService.setUserProfile(profile);

      // If it's a company, create the company doc too
      if (type === 'COMPANY') {
        await dbService.createCompany(firebaseUser.uid, finalCompanyName, firebaseUser.uid, email);
      }
    } else {
      // If profile already exists, sync with matching team member ONLY if found by exact email
      if (profile.type === 'USER') {
        if (matchingTeam) {
          profile.linkedCompanyId = matchingTeam.companyId;
          profile.role = matchingTeam.member.role;
          profile.areas = matchingTeam.member.areas;
          const comp = await dbService.getCompany(matchingTeam.companyId);
          profile.companyName = comp?.name || 'Empresa Vinculada';
          await dbService.setUserProfile(profile);
        } else if (!profile.linkedCompanyId) {
          profile.linkedCompanyId = profile.id;
          await dbService.setUserProfile(profile);
        }
      }
    }
    
    return profile;
  },

  async register(email: string, pass: string, name: string, type: AccountType, companyName?: string) {
    const normEmail = email.trim().toLowerCase();

    // Verify if this email is already registered in Firestore
    const existingCred = await dbService.getUserCredential(normEmail);
    const existingProfile = await dbService.findUserProfileByEmail(normEmail);
    if (existingCred || existingProfile) {
      throw new Error('Este correo electrónico ya está registrado en el sistema. Debe iniciar sesión o usar otro correo.');
    }

    const userCredential = await createUserWithEmailAndPassword(auth, normEmail, pass);
    const firebaseUser = userCredential.user;

    try {
      let matchingTeam = null;
      if (type !== 'COMPANY') {
        matchingTeam = await dbService.findTeamMemberByEmail(normEmail);
        if (matchingTeam && matchingTeam.member.status === 'inactive') {
          await firebaseUser.delete();
          await signOut(auth);
          throw new Error('Su usuario se encuentra inactivo. Por favor contacte al administrador.');
        }
      }

      const actualType = matchingTeam ? 'USER' : type;
      
      let finalCompanyName = companyName || (actualType === 'COMPANY' ? name : 'Personal');
      if (matchingTeam) {
        const comp = await dbService.getCompany(matchingTeam.companyId);
        finalCompanyName = comp?.name || 'Empresa Vinculada';
      }

      // STRICT ISOLATION BY EMAIL:
      // A company is ALWAYS its own independent space (firebaseUser.uid).
      // Even if another company has the same name, they are separate because identity is strictly by email.
      // A user is linked to a company ONLY IF invited by that exact email, otherwise own personal space.
      const companyId = actualType === 'COMPANY' ? firebaseUser.uid : (matchingTeam ? matchingTeam.companyId : firebaseUser.uid);
      const role = matchingTeam ? matchingTeam.member.role : 'ADMIN';
      const areas: AreaPermission[] = matchingTeam ? matchingTeam.member.areas : ['INVENTARIO', 'PRODUCCION', 'REPORTES', 'CONFIGURACION'];

      const profile: Account = {
        id: firebaseUser.uid,
        type: actualType,
        name,
        email: normEmail,
        companyName: finalCompanyName,
        linkedCompanyId: companyId,
        role,
        areas
      };

      // Save profile to Firestore
      await dbService.setUserProfile(profile);

      // Save user credential to Firestore for fast lookups & verified password state
      await dbService.saveUserCredential(normEmail, pass);

      // If it's a company, create the company doc too
      if (actualType === 'COMPANY') {
        await dbService.createCompany(firebaseUser.uid, finalCompanyName, firebaseUser.uid, normEmail);
      }

      return profile;
    } catch (e) {
      try {
        await firebaseUser.delete();
      } catch (delErr) {
        console.error('Failed to clean up user after register error:', delErr);
      }
      await signOut(auth);
      throw e;
    }
  },

  async login(email: string, pass: string) {
    const normEmail = email.trim().toLowerCase();
    
    // Check if user has an active credential stored in Firestore
    const storedCred = await dbService.getUserCredential(normEmail);

    // CRITICAL: If credentials exist in the database, ONLY the new/current password is accepted!
    // Any old password has been deleted from the database and MUST BE REJECTED!
    if (storedCred && storedCred.passwordHash) {
      if (storedCred.passwordHash !== pass) {
        throw new Error('Contraseña incorrecta. La contraseña anterior ha sido eliminada; debes ingresar únicamente con la nueva contraseña.');
      }
    }

    try {
      const userCredential = await signInWithEmailAndPassword(auth, normEmail, pass);
      
      // Keep stored credential in sync as the single active password
      if (!storedCred || storedCred.passwordHash !== pass) {
        await dbService.saveUserCredential(normEmail, pass);
      }

      let profile = await dbService.getUserProfile(userCredential.user.uid);
      if (profile && profile.type === 'COMPANY') {
        if (profile.linkedCompanyId !== userCredential.user.uid) {
          profile.linkedCompanyId = userCredential.user.uid;
          await dbService.setUserProfile(profile);
        }
        return profile;
      }

      const matchingTeam = await dbService.findTeamMemberByEmail(normEmail);
      if (matchingTeam && matchingTeam.member.status === 'inactive') {
        await signOut(auth);
        throw new Error('Su usuario se encuentra inactivo. Por favor contacte al administrador.');
      }
      
      if (!profile) {
        const companyId = matchingTeam ? matchingTeam.companyId : userCredential.user.uid;
        const role = matchingTeam ? matchingTeam.member.role : 'OPERATOR';
        const areas: AreaPermission[] = matchingTeam ? matchingTeam.member.areas : ['INVENTARIO'];
        
        let finalCompanyName = 'Personal';
        if (matchingTeam) {
          const comp = await dbService.getCompany(matchingTeam.companyId);
          finalCompanyName = comp?.name || 'Empresa Vinculada';
        }

        profile = {
          id: userCredential.user.uid,
          type: 'USER',
          name: userCredential.user.displayName || normEmail.split('@')[0],
          email: normEmail,
          companyName: finalCompanyName,
          linkedCompanyId: companyId,
          role: role,
          areas: areas
        };
        await dbService.setUserProfile(profile);
      } else if (profile.type === 'USER') {
        if (matchingTeam) {
          profile.linkedCompanyId = matchingTeam.companyId;
          profile.role = matchingTeam.member.role;
          profile.areas = matchingTeam.member.areas;
          const comp = await dbService.getCompany(matchingTeam.companyId);
          profile.companyName = comp?.name || 'Empresa Vinculada';
          await dbService.setUserProfile(profile);
        } else if (!profile.linkedCompanyId) {
          profile.linkedCompanyId = profile.id;
          await dbService.setUserProfile(profile);
        }
      }
      
      return profile;
    } catch (e: any) {
      // If Firebase Auth rejected the password (e.g. Firebase Auth password has not synced yet),
      // BUT the entered password matches the updated password recorded in Firestore!
      if (storedCred && storedCred.passwordHash === pass) {
        const profile = await dbService.findUserProfileByEmail(normEmail);
        if (profile) {
          const matchingTeam = await dbService.findTeamMemberByEmail(normEmail);
          if (matchingTeam && matchingTeam.member.status === 'inactive') {
            await signOut(auth);
            throw new Error('Su usuario se encuentra inactivo. Por favor contacte al administrador.');
          }
          return profile;
        }
      }

      await signOut(auth);
      throw e;
    }
  },

  async logout() {
    await signOut(auth);
  },

  async resetOrUpdatePassword(newPassword: string, oobCode?: string, email?: string) {
    const normEmail = email ? email.trim().toLowerCase() : undefined;

    // Check if there was a previous password in database so we can authenticate and update Firebase Auth
    let prevPassword = '';
    if (normEmail) {
      try {
        const prevCred = await dbService.getUserCredential(normEmail);
        if (prevCred?.passwordHash) {
          prevPassword = prevCred.passwordHash;
        }
      } catch (err) {
        console.warn('Could not check previous credential:', err);
      }
    }

    let authUpdated = false;

    // 1. If user is currently signed into Firebase Auth, update their Firebase Auth password directly
    if (auth.currentUser) {
      try {
        await updatePassword(auth.currentUser, newPassword);
        authUpdated = true;
      } catch (authErr) {
        console.warn('Direct updatePassword error:', authErr);
      }
    }

    // 2. If not signed in, but we had their previous password, sign in and update Firebase Auth password!
    if (!authUpdated && normEmail && prevPassword) {
      try {
        const userCred = await signInWithEmailAndPassword(auth, normEmail, prevPassword);
        if (userCred.user) {
          await updatePassword(userCred.user, newPassword);
          authUpdated = true;
          // Sign out so they log in with the new password
          await signOut(auth);
        }
      } catch (prevAuthErr) {
        console.warn('Could not reauth with previous password to sync Firebase Auth:', prevAuthErr);
      }
    }

    // 3. If an oobCode (email reset link) was provided, confirm the reset
    if (!authUpdated && oobCode) {
      try {
        await confirmPasswordReset(auth, oobCode, newPassword);
        authUpdated = true;
      } catch (oobErr) {
        console.warn('confirmPasswordReset error:', oobErr);
      }
    }

    // 4. Overwrite in Firestore database: deletes the old password completely and sets only the new password
    if (normEmail) {
      await dbService.saveUserCredential(normEmail, newPassword);
      await dbService.recordPasswordReset(normEmail, {
        method: authUpdated ? 'auth_and_db' : oobCode ? 'oob' : 'direct_db',
        status: 'UPDATED',
        oldPasswordErased: true,
        userAgent: typeof navigator !== 'undefined' ? navigator.userAgent : 'unknown'
      });
    }

    // 5. Send password reset email in background as secondary mechanism if available
    if (normEmail && !authUpdated) {
      try {
        await sendPasswordResetEmail(auth, normEmail);
      } catch (sendErr) {
        console.warn('sendPasswordResetEmail info:', sendErr);
      }
    }

    return { success: true, method: authUpdated ? 'auth_and_db' : 'database_updated' };
  },

  onAuthChanged(callback: (user: FirebaseUser | null) => void) {
    return onAuthStateChanged(auth, callback);
  }
};
