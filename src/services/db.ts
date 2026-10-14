import { 
  collection, 
  collectionGroup,
  doc, 
  setDoc, 
  getDoc, 
  getDocs, 
  query, 
  where, 
  onSnapshot, 
  updateDoc, 
  deleteDoc,
  serverTimestamp,
  type Unsubscribe
} from 'firebase/firestore';
import { db } from '../lib/firebase';
import { handleFirestoreError, OperationType } from './error-handler';
import { Product, Batch, Transaction, AuditEntry, TeamMember, AppNotification, Account, ProductionOrderRecord } from '../types';

function cleanUndefined(obj: any): any {
  if (obj === null || obj === undefined) return null;
  if (Array.isArray(obj)) {
    return obj.map(item => cleanUndefined(item));
  }
  if (typeof obj === 'object') {
    const cleaned: any = {};
    for (const key of Object.keys(obj)) {
      const val = obj[key];
      if (val !== undefined) {
        cleaned[key] = cleanUndefined(val);
      }
    }
    return cleaned;
  }
  return obj;
}

// Generic service to handle company-scoped data
export const dbService = {
  // --- USER PROFILES ---
  async getUserProfile(userId: string): Promise<Account | null> {
    try {
      const docRef = doc(db, 'users', userId);
      const snap = await getDoc(docRef);
      return snap.exists() ? snap.data() as Account : null;
    } catch (e) {
      handleFirestoreError(e, OperationType.GET, `users/${userId}`);
      return null;
    }
  },

  async setUserProfile(profile: Account): Promise<void> {
    try {
      const cleaned = cleanUndefined(profile);
      await setDoc(doc(db, 'users', profile.id), {
        ...cleaned,
        updatedAt: serverTimestamp()
      });
    } catch (e) {
      handleFirestoreError(e, OperationType.WRITE, `users/${profile.id}`);
    }
  },

  async findTeamMemberByEmail(email: string): Promise<{ companyId: string, member: TeamMember } | null> {
    try {
      const q = query(collectionGroup(db, 'team'), where('email', '==', email.trim().toLowerCase()));
      const snap = await getDocs(q);
      if (!snap.empty) {
        const docSnap = snap.docs[0];
        const companyId = docSnap.ref.parent?.parent?.id;
        if (companyId) {
          return {
            companyId,
            member: docSnap.data() as TeamMember
          };
        }
      }
      return null;
    } catch (e) {
      console.error('Error finding team member by email:', e);
      return null;
    }
  },

  // --- COMPANY DATA ---
  async getCompany(companyId: string) {
    try {
      const snap = await getDoc(doc(db, 'companies', companyId));
      return snap.exists() ? snap.data() : null;
    } catch (e) {
      handleFirestoreError(e, OperationType.GET, `companies/${companyId}`);
    }
  },

  async createCompany(companyId: string, name: string, ownerId: string) {
    try {
      await setDoc(doc(db, 'companies', companyId), {
        id: companyId,
        name,
        ownerId,
        createdAt: serverTimestamp()
      });
    } catch (e) {
      handleFirestoreError(e, OperationType.WRITE, `companies/${companyId}`);
    }
  },

  // --- SUB-COLLECTIONS WITH REAL-TIME SYNC ---
  subscribeToCollection<T>(
    companyId: string, 
    subPath: string, 
    callback: (data: T[]) => void
  ): Unsubscribe {
    const path = `companies/${companyId}/${subPath}`;
    return onSnapshot(
      collection(db, path),
      (snap) => {
        const items = snap.docs.map(d => ({ ...d.data() } as T));
        callback(items);
      },
      (error) => handleFirestoreError(error, OperationType.GET, path)
    );
  },

  // Generic write for any sub-collection
  async saveItem<T extends { id: string }>(
    companyId: string, 
    subPath: string, 
    item: T
  ): Promise<void> {
    const path = `companies/${companyId}/${subPath}`;
    try {
      const cleaned = cleanUndefined(item);
      await setDoc(doc(db, path, item.id), {
        ...cleaned,
        updatedAt: serverTimestamp()
      });
    } catch (e) {
      handleFirestoreError(e, OperationType.WRITE, `${path}/${item.id}`);
    }
  },

  async deleteItem(
    companyId: string, 
    subPath: string, 
    itemId: string
  ): Promise<void> {
    const path = `companies/${companyId}/${subPath}`;
    try {
      await deleteDoc(doc(db, path, itemId));
    } catch (e) {
      handleFirestoreError(e, OperationType.DELETE, `${path}/${itemId}`);
    }
  },

  // Specific helpers
  async addAuditLog(companyId: string, log: AuditEntry) {
    return this.saveItem(companyId, 'auditLogs', log);
  },

  async saveBatch(companyId: string, batch: Batch) {
    return this.saveItem(companyId, 'batches', batch);
  },

  async saveTransaction(companyId: string, tx: Transaction) {
    return this.saveItem(companyId, 'transactions', tx);
  },

  async saveProduct(companyId: string, product: Product) {
    return this.saveItem(companyId, 'products', product);
  },

  async saveTeamMember(companyId: string, member: TeamMember) {
    return this.saveItem(companyId, 'team', member);
  },

  async saveNotification(companyId: string, notif: AppNotification) {
    return this.saveItem(companyId, 'notifications', notif);
  },

  async saveProductionOrder(companyId: string, order: ProductionOrderRecord) {
    return this.saveItem(companyId, 'productionOrders', order);
  },

  async recordPasswordReset(email: string, details?: Record<string, any>) {
    try {
      const resetId = `reset_${Date.now()}_${Math.random().toString(36).substring(2, 9)}`;
      await setDoc(doc(db, 'passwordResets', resetId), {
        id: resetId,
        email: email.trim().toLowerCase(),
        timestamp: new Date().toISOString(),
        createdAt: serverTimestamp(),
        details: details || {}
      });
      return resetId;
    } catch (e) {
      console.error('Error saving password reset record:', e);
      // Non-blocking so users are not stranded if logging fails
      return null;
    }
  },

  // Encode email for deterministic document key in userCredentials
  getEmailKey(email: string): string {
    return email.trim().toLowerCase().replace(/[^a-zA-Z0-9]/g, '_');
  },

  async saveUserCredential(email: string, passwordHash: string): Promise<void> {
    try {
      const credId = this.getEmailKey(email);
      // Clean overwrite without merge: completely wipes out old password and replaces with the new password
      await setDoc(doc(db, 'userCredentials', credId), {
        email: email.trim().toLowerCase(),
        passwordHash,
        updatedAt: serverTimestamp()
      });
    } catch (e) {
      console.error('Error saving user credential:', e);
    }
  },

  async deleteUserCredential(email: string): Promise<void> {
    try {
      const credId = this.getEmailKey(email);
      await deleteDoc(doc(db, 'userCredentials', credId));
    } catch (e) {
      console.error('Error deleting user credential:', e);
    }
  },

  async getUserCredential(email: string): Promise<{ email: string; passwordHash: string } | null> {
    try {
      const credId = this.getEmailKey(email);
      const snap = await getDoc(doc(db, 'userCredentials', credId));
      if (snap.exists()) {
        return snap.data() as { email: string; passwordHash: string };
      }
      return null;
    } catch (e) {
      console.error('Error reading user credential:', e);
      return null;
    }
  },

  async findUserProfileByEmail(email: string): Promise<Account | null> {
    try {
      const normEmail = email.trim().toLowerCase();
      const q = query(collection(db, 'users'), where('email', '==', normEmail));
      const snap = await getDocs(q);
      if (!snap.empty) {
        return snap.docs[0].data() as Account;
      }
      return null;
    } catch (e) {
      console.error('Error finding profile by email:', e);
      return null;
    }
  }
};
