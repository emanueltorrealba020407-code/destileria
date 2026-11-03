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
      console.warn(`Error getting user profile users/${userId}:`, e);
      return null;
    }
  },

  async setUserProfile(profile: Account): Promise<void> {
    try {
      const cleaned = cleanUndefined(profile);
      await setDoc(doc(db, 'users', profile.id), {
        ...cleaned,
        updatedAt: serverTimestamp()
      }, { merge: true });
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
      console.warn('Error finding team member by email:', e);
      return null;
    }
  },

  // --- COMPANY DATA ---
  async getCompany(companyId: string) {
    try {
      const snap = await getDoc(doc(db, 'companies', companyId));
      return snap.exists() ? snap.data() : null;
    } catch (e) {
      console.warn(`Error getting company companies/${companyId}:`, e);
      return null;
    }
  },

  async createCompany(companyId: string, name: string, ownerId: string) {
    try {
      await setDoc(doc(db, 'companies', companyId), {
        id: companyId,
        name,
        ownerId,
        createdAt: serverTimestamp()
      }, { merge: true });
    } catch (e) {
      console.warn(`Error creating company companies/${companyId}:`, e);
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
      (error) => {
        console.warn(`Subscription update for ${path}:`, error?.message || error);
      }
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
      }, { merge: true });
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
      console.warn(`Error al eliminar de ${path}/${itemId}:`, e);
      handleFirestoreError(e, OperationType.DELETE, `${path}/${itemId}`);
    }
    // Also defensively check and delete from root collection if it exists there
    try {
      await deleteDoc(doc(db, subPath, itemId));
    } catch {
      // Ignore root collection errors
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

  async hasCompanyProducts(companyId: string): Promise<boolean> {
    try {
      const snap = await getDocs(collection(db, `companies/${companyId}/products`));
      return !snap.empty;
    } catch (e) {
      console.warn('Error checking company products:', e);
      return false;
    }
  },

  async listCompaniesWithStats(): Promise<Array<{ id: string; name: string; productCount: number; batchCount: number }>> {
    try {
      const snap = await getDocs(collection(db, 'companies'));
      const results: Array<{ id: string; name: string; productCount: number; batchCount: number }> = [];
      for (const d of snap.docs) {
        const data = d.data();
        let pCount = 0;
        let bCount = 0;
        try {
          const pSnap = await getDocs(collection(db, `companies/${d.id}/products`));
          pCount = pSnap.size;
          const bSnap = await getDocs(collection(db, `companies/${d.id}/batches`));
          bCount = bSnap.size;
        } catch {
          // ignore
        }
        results.push({
          id: d.id,
          name: data.name || d.id,
          productCount: pCount,
          batchCount: bCount
        });
      }
      return results;
    } catch (e) {
      console.warn('Error listing companies with stats:', e);
      return [];
    }
  },

  // Seed standard complete distillery inventory into company database
  async seedDefaultCompanyData(companyId: string, companyName: string, userName: string): Promise<void> {
    try {
      // 1. Ensure company doc exists
      await this.createCompany(companyId, companyName, companyId);

      // 2. Standard Products (Raw Materials, Packaging, and Finished Goods with recipe & packaging specs)
      const rawMaterials: Product[] = [
        {
          id: 'prod-mat-001',
          name: 'Alcohol Etílico Extra Neutro 96°',
          description: 'Alcohol desnaturalizado de caña para formulación de licores finos',
          category: 'Materia Prima',
          type: 'materia prima',
          unit: 'litro',
          minStock: 500
        },
        {
          id: 'prod-mat-002',
          name: 'Melaza de Caña Pura Concentrada',
          description: 'Melaza virgen de caña de azúcar para fermentación y añejamiento',
          category: 'Materia Prima',
          type: 'materia prima',
          unit: 'kilo',
          minStock: 300
        },
        {
          id: 'prod-mat-003',
          name: 'Esencia y Extracto de Roble con Caramelo',
          description: 'Aroma y colorante natural obtenido de maderas tostadas',
          category: 'Insumos',
          type: 'materia prima',
          unit: 'litro',
          minStock: 25
        },
        {
          id: 'prod-emp-004',
          name: 'Botella de Vidrio 750ml Especial Ormón',
          description: 'Envase de vidrio pesado color ámbar para licores premium',
          category: 'Empaque',
          type: 'materia prima',
          unit: 'unidad',
          minStock: 300
        },
        {
          id: 'prod-emp-005',
          name: 'Tapa Corcho Sintético Hermético 28mm',
          description: 'Corcho sintético dorado de alto sellado al vacío',
          category: 'Empaque',
          type: 'materia prima',
          unit: 'unidad',
          minStock: 300
        },
        {
          id: 'prod-emp-006',
          name: 'Caja de Cartón Corrugado x12 Botellas',
          description: 'Caja reforzada con separadores internos para transporte seguro',
          category: 'Empaque',
          type: 'materia prima',
          unit: 'unidad',
          minStock: 50
        },
        {
          id: 'prod-emp-007',
          name: 'Bulto / Manga Termoencogible x24 Botellas',
          description: 'Empaque de polietileno de alta resistencia para distribución mayorista',
          category: 'Empaque',
          type: 'materia prima',
          unit: 'unidad',
          minStock: 30
        }
      ];

      const finishedProducts: Product[] = [
        {
          id: 'prod-fin-001',
          name: 'Ron Añejo Especial 8 Años 750ml',
          description: 'Ron ultra premium destilado en barricas de roble blanco, sabor suave con notas de caramelo y tabaco.',
          category: 'Licores Terminados',
          type: 'producto terminado',
          unit: 'unidad',
          minStock: 50,
          packaging: {
            packagingType: 'ambos',
            unitsPerBox: 12,
            unitsPerBundle: 24
          },
          recipe: [
            { materialProductId: 'prod-mat-001', materialName: 'Alcohol Etílico Extra Neutro 96°', quantityNeeded: 0.75, unit: 'litro' },
            { materialProductId: 'prod-mat-002', materialName: 'Melaza de Caña Pura Concentrada', quantityNeeded: 0.20, unit: 'kilo' },
            { materialProductId: 'prod-mat-003', materialName: 'Esencia y Extracto de Roble con Caramelo', quantityNeeded: 0.05, unit: 'litro' },
            { materialProductId: 'prod-emp-004', materialName: 'Botella de Vidrio 750ml Especial Ormón', quantityNeeded: 1, unit: 'unidad' },
            { materialProductId: 'prod-emp-005', materialName: 'Tapa Corcho Sintético Hermético 28mm', quantityNeeded: 1, unit: 'unidad' }
          ]
        },
        {
          id: 'prod-fin-002',
          name: 'Aguardiente Artesanal Doble Destilado 750ml',
          description: 'Aguardiente tradicional anisado con agua de manantial y triple filtrado.',
          category: 'Licores Terminados',
          type: 'producto terminado',
          unit: 'unidad',
          minStock: 40,
          packaging: {
            packagingType: 'caja',
            unitsPerBox: 12,
            unitsPerBundle: 0
          },
          recipe: [
            { materialProductId: 'prod-mat-001', materialName: 'Alcohol Etílico Extra Neutro 96°', quantityNeeded: 0.75, unit: 'litro' },
            { materialProductId: 'prod-mat-003', materialName: 'Esencia y Extracto de Roble con Caramelo', quantityNeeded: 0.02, unit: 'litro' },
            { materialProductId: 'prod-emp-004', materialName: 'Botella de Vidrio 750ml Especial Ormón', quantityNeeded: 1, unit: 'unidad' },
            { materialProductId: 'prod-emp-005', materialName: 'Tapa Corcho Sintético Hermético 28mm', quantityNeeded: 1, unit: 'unidad' }
          ]
        },
        {
          id: 'prod-fin-003',
          name: 'Licor Especiado de Ron y Vainilla 750ml',
          description: 'Licor dulce infusionado con vainilla de Papantla y canela de Ceilán.',
          category: 'Licores Terminados',
          type: 'producto terminado',
          unit: 'unidad',
          minStock: 30,
          packaging: {
            packagingType: 'caja',
            unitsPerBox: 12,
            unitsPerBundle: 0
          },
          recipe: [
            { materialProductId: 'prod-mat-001', materialName: 'Alcohol Etílico Extra Neutro 96°', quantityNeeded: 0.70, unit: 'litro' },
            { materialProductId: 'prod-mat-002', materialName: 'Melaza de Caña Pura Concentrada', quantityNeeded: 0.15, unit: 'kilo' },
            { materialProductId: 'prod-mat-003', materialName: 'Esencia y Extracto de Roble con Caramelo', quantityNeeded: 0.08, unit: 'litro' },
            { materialProductId: 'prod-emp-004', materialName: 'Botella de Vidrio 750ml Especial Ormón', quantityNeeded: 1, unit: 'unidad' },
            { materialProductId: 'prod-emp-005', materialName: 'Tapa Corcho Sintético Hermético 28mm', quantityNeeded: 1, unit: 'unidad' }
          ]
        }
      ];

      const allProducts = [...rawMaterials, ...finishedProducts];
      for (const prod of allProducts) {
        await this.saveProduct(companyId, prod);
      }

      // 3. Batches for raw materials
      const initialDate = new Date(Date.now() - 15 * 86400000).toISOString();
      const materialBatches: Batch[] = [
        {
          id: 'batch-mat-001',
          productId: 'prod-mat-001',
          entryDate: initialDate,
          quantity: 2500,
          remainingQuantity: 2200,
          unitCost: 8500,
          supplier: 'Destiladora Nacional S.A.S.',
          description: 'Lote inicial de alcohol etílico alta pureza',
          type: 'PURCHASE',
          performedBy: userName
        },
        {
          id: 'batch-mat-002',
          productId: 'prod-mat-002',
          entryDate: initialDate,
          quantity: 1800,
          remainingQuantity: 1600,
          unitCost: 2400,
          supplier: 'Ingenio Central del Valle',
          description: 'Melaza virgen certificada brix 78°',
          type: 'PURCHASE',
          performedBy: userName
        },
        {
          id: 'batch-mat-003',
          productId: 'prod-mat-003',
          entryDate: initialDate,
          quantity: 60,
          remainingQuantity: 52,
          unitCost: 45000,
          supplier: 'Extractos y Aromas de Occidente',
          description: 'Esencia roble americano tostado medio',
          type: 'PURCHASE',
          performedBy: userName
        },
        {
          id: 'batch-emp-004',
          productId: 'prod-emp-004',
          entryDate: initialDate,
          quantity: 2000,
          remainingQuantity: 1760,
          unitCost: 3800,
          supplier: 'Cristalería Andina Industrial',
          description: 'Botellas 750ml paletizadas',
          type: 'PURCHASE',
          performedBy: userName
        },
        {
          id: 'batch-emp-005',
          productId: 'prod-emp-005',
          entryDate: initialDate,
          quantity: 2000,
          remainingQuantity: 1760,
          unitCost: 850,
          supplier: 'Tapones del Caribe',
          description: 'Corchos sintéticos alta hermeticidad',
          type: 'PURCHASE',
          performedBy: userName
        },
        {
          id: 'batch-emp-006',
          productId: 'prod-emp-006',
          entryDate: initialDate,
          quantity: 300,
          remainingQuantity: 280,
          unitCost: 4200,
          supplier: 'Cartones del Pacífico',
          description: 'Cajas x12 botellas serigrafiadas Ormón',
          type: 'PURCHASE',
          performedBy: userName
        },
        {
          id: 'batch-emp-007',
          productId: 'prod-emp-007',
          entryDate: initialDate,
          quantity: 150,
          remainingQuantity: 140,
          unitCost: 6500,
          supplier: 'Empaques Flexibles del Valle',
          description: 'Manga termoencogible para bultos x24',
          type: 'PURCHASE',
          performedBy: userName
        }
      ];

      // 4. Batches for Finished Products (Produced goods in inventory)
      const prodBatchDate1 = new Date(Date.now() - 10 * 86400000).toISOString();
      const prodBatchDate2 = new Date(Date.now() - 5 * 86400000).toISOString();

      const finishedBatches: Batch[] = [
        {
          id: 'batch-fin-001',
          productId: 'prod-fin-001',
          entryDate: prodBatchDate1,
          quantity: 240, // 20 cajas o 10 bultos
          remainingQuantity: 192, // 48 unidades vendidas
          unitCost: 14250,
          supplier: 'Producción Interna',
          description: 'Lote fabricado a partir de materias primas (20 cajas de 12 unidades / 10 bultos de 24)',
          type: 'ENTRY',
          performedBy: userName
        },
        {
          id: 'batch-fin-002',
          productId: 'prod-fin-002',
          entryDate: prodBatchDate2,
          quantity: 120, // 10 cajas
          remainingQuantity: 96, // 24 unidades vendidas
          unitCost: 11400,
          supplier: 'Producción Interna',
          description: 'Lote destilado y envasado (10 cajas de 12 unidades)',
          type: 'ENTRY',
          performedBy: userName
        },
        {
          id: 'batch-fin-003',
          productId: 'prod-fin-003',
          entryDate: prodBatchDate2,
          quantity: 96, // 8 cajas
          remainingQuantity: 84, // 12 unidades vendidas
          unitCost: 15600,
          supplier: 'Producción Interna',
          description: 'Lote especial aromatizado (8 cajas de 12 unidades)',
          type: 'ENTRY',
          performedBy: userName
        }
      ];

      const allBatches = [...materialBatches, ...finishedBatches];
      for (const batch of allBatches) {
        await this.saveBatch(companyId, batch);
      }

      // 5. Initial Transactions (Entries, Production and Sales)
      const transactions: Transaction[] = [
        {
          id: 'tx-init-001',
          productId: 'prod-fin-001',
          type: 'ENTRY',
          quantity: 240,
          unitCost: 14250,
          date: prodBatchDate1,
          reference: 'ORD-PROD-2026-001',
          performedBy: userName,
          description: 'Entrada de producción por lote de 20 cajas (240 botellas) de Ron Añejo Especial',
          batchId: 'batch-fin-001'
        },
        {
          id: 'tx-init-002',
          productId: 'prod-fin-001',
          type: 'SALE',
          quantity: 48, // 4 cajas
          unitCost: 14250,
          salePrice: 38000,
          date: new Date(Date.now() - 3 * 86400000).toISOString(),
          reference: 'FACT-VTA-001',
          performedBy: userName,
          description: 'Venta mayorista a Distribuidora Los Andes (4 cajas / 48 botellas)'
        },
        {
          id: 'tx-init-003',
          productId: 'prod-fin-002',
          type: 'ENTRY',
          quantity: 120,
          unitCost: 11400,
          date: prodBatchDate2,
          reference: 'ORD-PROD-2026-002',
          performedBy: userName,
          description: 'Entrada de producción por 10 cajas de Aguardiente Artesanal',
          batchId: 'batch-fin-002'
        },
        {
          id: 'tx-init-004',
          productId: 'prod-fin-002',
          type: 'SALE',
          quantity: 24, // 2 cajas
          unitCost: 11400,
          salePrice: 28000,
          date: new Date(Date.now() - 2 * 86400000).toISOString(),
          reference: 'FACT-VTA-002',
          performedBy: userName,
          description: 'Venta comercial Restaurantes y Bares del Centro (2 cajas)'
        },
        {
          id: 'tx-init-005',
          productId: 'prod-fin-003',
          type: 'ENTRY',
          quantity: 96,
          unitCost: 15600,
          date: prodBatchDate2,
          reference: 'ORD-PROD-2026-003',
          performedBy: userName,
          description: 'Entrada de producción por 8 cajas de Licor Especiado',
          batchId: 'batch-fin-003'
        },
        {
          id: 'tx-init-006',
          productId: 'prod-fin-003',
          type: 'SALE',
          quantity: 12, // 1 caja
          unitCost: 15600,
          salePrice: 42000,
          date: new Date(Date.now() - 1 * 86400000).toISOString(),
          reference: 'FACT-VTA-003',
          performedBy: userName,
          description: 'Venta boutique gourmet (1 caja de 12 botellas)'
        }
      ];

      for (const tx of transactions) {
        await this.saveTransaction(companyId, tx);
      }

      // 6. Production Order Records
      const prodOrder: ProductionOrderRecord = {
        id: 'ord-prod-rec-001',
        finishedProductId: 'prod-fin-001',
        finishedProductName: 'Ron Añejo Especial 8 Años 750ml',
        quantityProduced: 240,
        boxesProduced: 20,
        bundlesProduced: 10,
        totalCost: 3420000,
        unitCost: 14250,
        date: prodBatchDate1,
        performedBy: userName,
        consumedMaterials: [
          { materialProductId: 'prod-mat-001', materialName: 'Alcohol Etílico Extra Neutro 96°', quantityUsed: 180, unit: 'litro', cost: 1530000 },
          { materialProductId: 'prod-mat-002', materialName: 'Melaza de Caña Pura Concentrada', quantityUsed: 48, unit: 'kilo', cost: 115200 },
          { materialProductId: 'prod-mat-003', materialName: 'Esencia y Extracto de Roble con Caramelo', quantityUsed: 12, unit: 'litro', cost: 540000 },
          { materialProductId: 'prod-emp-004', materialName: 'Botella de Vidrio 750ml Especial Ormón', quantityUsed: 240, unit: 'unidad', cost: 912000 },
          { materialProductId: 'prod-emp-005', materialName: 'Tapa Corcho Sintético Hermético 28mm', quantityUsed: 240, unit: 'unidad', cost: 204000 },
          { materialProductId: 'prod-emp-006', materialName: 'Caja de Cartón Corrugado x12 Botellas', quantityUsed: 20, unit: 'unidad', cost: 84000 }
        ],
        notes: 'Lote maestro de producción inicial con añejamiento en barricas.'
      };
      await this.saveProductionOrder(companyId, prodOrder);

      // 7. Initial Team Member
      const teamMember: TeamMember = {
        id: 'team-root-admin',
        name: userName || 'Administrador General',
        email: 'admin@destileriaormon.com',
        role: 'ADMIN',
        areas: ['INVENTARIO', 'PRODUCCION', 'REPORTES', 'CONFIGURACION'],
        status: 'active'
      };
      await this.saveTeamMember(companyId, teamMember);

      // 8. Welcome notification
      const welcomeNotification: AppNotification = {
        id: 'notif-welcome',
        type: 'SUCCESS',
        title: 'Sistema de Inventario y Kardex Listo',
        message: 'La base de datos se ha cargado con éxito con el inventario completo, materias primas, productos terminados (Ron, Aguardiente y Licores), recetas y movimientos del Kardex.',
        timestamp: new Date().toISOString(),
        read: false
      };
      await this.saveNotification(companyId, welcomeNotification);

      // 9. Audit log
      const auditLog: AuditEntry = {
        id: 'audit-init-sys',
        timestamp: new Date().toISOString(),
        userId: companyId,
        userName: userName || 'Sistema',
        action: 'SYSTEM',
        entityType: 'CONFIG',
        entityId: companyId,
        details: 'Carga inicial completa de la base de datos: catálogo de productos, recetas, empaques, lotes iniciales y transacciones históricas.'
      };
      await this.addAuditLog(companyId, auditLog);

      console.log('Database successfully seeded for company:', companyId);
    } catch (e) {
      console.error('Error seeding company data:', e);
    }
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
      console.warn('Error saving password reset record:', e);
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
      await setDoc(doc(db, 'userCredentials', credId), {
        email: email.trim().toLowerCase(),
        passwordHash,
        updatedAt: serverTimestamp()
      }, { merge: true });
    } catch (e) {
      console.warn('Error saving user credential:', e);
    }
  },

  async deleteUserCredential(email: string): Promise<void> {
    try {
      const credId = this.getEmailKey(email);
      await deleteDoc(doc(db, 'userCredentials', credId));
    } catch (e) {
      console.warn('Error deleting user credential:', e);
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
      console.warn('Error reading user credential:', e);
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
      console.warn('Error finding profile by email:', e);
      return null;
    }
  }
};
