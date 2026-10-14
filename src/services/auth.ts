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
import { Account, AccountType } from '../types';

const googleProvider = new GoogleAuthProvider();

export const authService = {
  async loginWithGoogle(desiredType?: AccountType, companyName?: string) {
    const userCredential = await signInWithPopup(auth, googleProvider);
    const firebaseUser = userCredential.user;
    
    const email = firebaseUser.email || '';
    let profile = await dbService.getUserProfile(firebaseUser.uid);
    if (profile && profile.type === 'COMPANY') {
      return profile;
    }

    const matchingTeam = email ? await dbService.findTeamMemberByEmail(email) : null;
    if (matchingTeam && matchingTeam.member.status === 'inactive') {
      await signOut(auth);
      throw new Error('Su usuario se encuentra inactivo. Por favor contacte al administrador.');
    }
    
    if (!profile) {
      // Create a profile if it doesn't exist
      const email = firebaseUser.email || '';
      const matchingTeam = await dbService.findTeamMemberByEmail(email);

      const type = matchingTeam ? 'USER' : (desiredType || 'USER');
      const name = firebaseUser.displayName || 'Usuario Google';
      
      let finalCompanyName = companyName || 'Personal';
      if (matchingTeam) {
        const comp = await dbService.getCompany(matchingTeam.companyId);
        finalCompanyName = comp?.name || 'Empresa Vinculada';
      }

      const companyId = type === 'COMPANY' ? firebaseUser.uid : (matchingTeam ? matchingTeam.companyId : '');
      const role = matchingTeam ? matchingTeam.member.role : 'ADMIN';
      const areas = matchingTeam ? matchingTeam.member.areas : [];

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
        await dbService.createCompany(firebaseUser.uid, finalCompanyName, firebaseUser.uid);
      }
    } else {
      // If profile already exists but is empty of company ID, sync with any matching team member
      if (profile.type === 'USER' && (!profile.linkedCompanyId || profile.linkedCompanyId === '')) {
        const email = firebaseUser.email || profile.email;
        const matchingTeam = await dbService.findTeamMemberByEmail(email);
        if (matchingTeam) {
          profile.linkedCompanyId = matchingTeam.companyId;
          profile.role = matchingTeam.member.role;
          profile.areas = matchingTeam.member.areas;
          const comp = await dbService.getCompany(matchingTeam.companyId);
          profile.companyName = comp?.name || 'Empresa Vinculada';
          await dbService.setUserProfile(profile);
        }
      }
    }
    
    return profile;
  },

  async register(email: string, pass: string, name: string, type: AccountType, companyName?: string) {
    const userCredential = await createUserWithEmailAndPassword(auth, email, pass);
    const firebaseUser = userCredential.user;

    try {
      let matchingTeam = null;
      if (type !== 'COMPANY') {
        matchingTeam = await dbService.findTeamMemberByEmail(email);
        if (matchingTeam && matchingTeam.member.status === 'inactive') {
          await firebaseUser.delete();
          await signOut(auth);
          throw new Error('Su usuario se encuentra inactivo. Por favor contacte al administrador.');
        }
      }

      const actualType = matchingTeam ? 'USER' : type;
      
      let finalCompanyName = companyName || 'Personal';
      if (matchingTeam) {
        const comp = await dbService.getCompany(matchingTeam.companyId);
        finalCompanyName = comp?.name || 'Empresa Vinculada';
      }

      const companyId = actualType === 'COMPANY' ? firebaseUser.uid : (matchingTeam ? matchingTeam.companyId : '');
      const role = matchingTeam ? matchingTeam.member.role : 'ADMIN';
      const areas = matchingTeam ? matchingTeam.member.areas : [];

      const profile: Account = {
        id: firebaseUser.uid,
        type: actualType,
        name,
        email,
        companyName: finalCompanyName,
        linkedCompanyId: companyId,
        role,
        areas
      };

      // Save profile to Firestore
      await dbService.setUserProfile(profile);

      // Save user credential to Firestore for fast lookups & verified password state
      await dbService.saveUserCredential(email, pass);

      // If it's a company, create the company doc too
      if (actualType === 'COMPANY') {
        await dbService.createCompany(firebaseUser.uid, companyName || name, firebaseUser.uid);
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
        return profile;
      }

      const matchingTeam = await dbService.findTeamMemberByEmail(normEmail);
      if (matchingTeam && matchingTeam.member.status === 'inactive') {
        await signOut(auth);
        throw new Error('Su usuario se encuentra inactivo. Por favor contacte al administrador.');
      }
      
      if (!profile) {
        const companyId = matchingTeam ? matchingTeam.companyId : '';
        const role = matchingTeam ? matchingTeam.member.role : 'OPERATOR';
        const areas = matchingTeam ? matchingTeam.member.areas : [];
        
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
      } else if (profile.type === 'USER' && (!profile.linkedCompanyId || profile.linkedCompanyId === '')) {
        if (matchingTeam) {
          profile.linkedCompanyId = matchingTeam.companyId;
          profile.role = matchingTeam.member.role;
          profile.areas = matchingTeam.member.areas;
          const comp = await dbService.getCompany(matchingTeam.companyId);
          profile.companyName = comp?.name || 'Empresa Vinculada';
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
