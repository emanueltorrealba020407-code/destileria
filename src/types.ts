/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

export type UserRole = 'OPERATOR' | 'ADMIN';

export type AccountType = 'COMPANY' | 'USER';

export interface Account {
  id: string;
  type: AccountType;
  companyName: string;
  name: string;
  email: string;
  password?: string;
  linkedCompanyId?: string;
  role?: UserRole;
  areas?: AreaPermission[];
}

export type AreaPermission = 'INVENTARIO' | 'PRODUCCION' | 'REPORTES' | 'CONFIGURACION';

export interface TeamMember {
  id: string;
  name: string;
  email: string;
  role: UserRole;
  areas: AreaPermission[];
  status: 'active' | 'inactive';
}

export interface Batch {
  id: string;
  productId: string;
  entryDate: string; // ISO string
  quantity: number;
  remainingQuantity: number;
  unitCost: number;
  supplier: string;
  description?: string;
  type?: 'ENTRY' | 'PURCHASE';
  performedBy?: string;
}

export interface RecipeItem {
  materialProductId: string; // ID of the raw material or component
  materialName?: string;
  quantityNeeded: number; // Quantity needed per 1 unit of finished product
  unit?: 'unidad' | 'kilo' | 'litro';
}

export interface PackagingConfig {
  packagingType: 'ninguno' | 'caja' | 'bulto' | 'ambos';
  unitsPerBox?: number;   // Cuántas unidades lleva cada caja
  unitsPerBundle?: number; // Cuántas unidades lleva cada bulto
}

export interface Product {
  id: string;
  name: string;
  description?: string;
  category: string;
  type: 'materia prima' | 'producción' | 'producto terminado';
  unit: 'unidad' | 'kilo' | 'litro';
  minStock: number;
  // Receta / consumo de materiales (para productos terminados o en producción)
  recipe?: RecipeItem[];
  // Especificación de empaque (caja o bulto y cuánto lleva en cada una)
  packaging?: PackagingConfig;
}

export interface ProductionOrderRecord {
  id: string;
  finishedProductId: string;
  finishedProductName: string;
  quantityProduced: number; // In base units (litro, kilo, unidad)
  boxesProduced?: number;   // Calculated boxes if applicable
  bundlesProduced?: number; // Calculated bundles if applicable
  totalCost: number;        // Sum of consumed materials
  unitCost: number;         // totalCost / quantityProduced
  date: string;
  performedBy: string;
  consumedMaterials: Array<{
    materialProductId: string;
    materialName: string;
    quantityUsed: number;
    unit: string;
    cost: number;
  }>;
  notes?: string;
}

export interface Transaction {
  id: string;
  productId: string;
  type: 'ENTRY' | 'EXIT' | 'PURCHASE' | 'SALE';
  quantity: number;
  unitCost: number; // For ENTRY, cost of entry. For EXIT, average/fifo cost?
  salePrice?: number; // Price per unit at which the sale was made
  date: string;
  reference: string; // Order number, invoice, etc.
  performedBy?: string;
  description?: string;
  batchId?: string; // Optional: link to a specific batch if entry
}

export interface KardexEntry {
  date: string;
  reference: string;
  type: 'ENTRY' | 'EXIT' | 'PURCHASE' | 'SALE';
  productId?: string;
  productName?: string;
  // Entradas
  entryQty: number;
  entryUnitCost: number;
  entryTotal: number;
  // Salidas
  exitQty: number;
  exitUnitCost: number;
  exitTotal: number;
  salePrice?: number; // Price per unit if it's a sale
  // Saldos
  balanceQty: number;
  balanceUnitCost: number;
  balanceTotal: number;
  batchId?: string;
  performedBy?: string;
}

export interface AuditEntry {
  id: string;
  timestamp: string;
  userId: string;
  userName: string;
  action: 'CREATE' | 'UPDATE' | 'DELETE' | 'SYSTEM';
  entityType: 'PRODUCT' | 'BATCH' | 'TRANSACTION' | 'TEAM' | 'CONFIG' | 'PRODUCTION';
  entityId: string;
  details: string;
  previousValue?: string; // JSON string
  newValue?: string; // JSON string
}

export interface AppNotification {
  id: string;
  type: 'LOW_STOCK' | 'SUCCESS' | 'INFO' | 'WARNING';
  title: string;
  message: string;
  timestamp: string;
  read: boolean;
  linkTab?: string;
  productId?: string;
}
