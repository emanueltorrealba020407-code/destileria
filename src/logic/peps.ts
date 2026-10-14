import { Product, Batch, Transaction, KardexEntry } from '../types';

/**
 * Calculates exit cost and updates batches according to PEPS (FIFO)
 * PEPS: First-In, First-Out. Exits take from the oldest entries first.
 */
export function processExit(
  quantityToExit: number,
  allBatches: Batch[],
  productId: string,
  description?: string,
  type: 'EXIT' | 'SALE' = 'EXIT',
  performedBy?: string,
  salePrice?: number,
  customDate?: string
): { 
  updatedBatches: Batch[], 
  transactions: Transaction[],
  totalCost: number 
} {
  // Sort batches by entry date (oldest first) with deterministic ID tie-breaker
  const productBatches = allBatches
    .filter(b => b.productId === productId && b.remainingQuantity > 0)
    .sort((a, b) => {
      const timeDiff = new Date(a.entryDate).getTime() - new Date(b.entryDate).getTime();
      return timeDiff !== 0 ? timeDiff : a.id.localeCompare(b.id);
    });

  let remainingToProcess = quantityToExit;
  let totalCost = 0;
  const newTransactions: Transaction[] = [];
  const batchesToUpdate = [...allBatches];

  // Resolve transaction timestamp
  const txDate = customDate 
    ? (customDate.includes('T') ? customDate : `${customDate}T${new Date().toTimeString().split(' ')[0]}`)
    : new Date().toISOString();

  for (const batch of productBatches) {
    if (remainingToProcess <= 0) break;

    const takeQty = Math.min(batch.remainingQuantity, remainingToProcess);
    const batchIndex = batchesToUpdate.findIndex(b => b.id === batch.id);
    
    // Update batch remaining quantity with zero-clamp
    const newRemaining = Math.max(0, batch.remainingQuantity - takeQty);
    batchesToUpdate[batchIndex] = {
      ...batch,
      remainingQuantity: newRemaining
    };

    totalCost += takeQty * batch.unitCost;
    remainingToProcess = Math.max(0, remainingToProcess - takeQty);

    newTransactions.push({
      id: crypto.randomUUID(),
      productId: productId,
      batchId: batch.id,
      type: type,
      quantity: takeQty,
      unitCost: batch.unitCost,
      salePrice: salePrice,
      date: txDate,
      reference: type === 'SALE' 
        ? `Costo de Ventas - Lote: ${batch.id.slice(0, 8)}` 
        : `Salida de Inventario - Lote: ${batch.id.slice(0, 8)}`,
      performedBy: performedBy,
      description: description
    });
  }

  if (remainingToProcess > 0) {
    throw new Error('Stock insuficiente para realizar la salida solicitada.');
  }

  return { updatedBatches: batchesToUpdate, transactions: newTransactions, totalCost };
}

export interface BatchConsumption {
  batchId: string;
  entryDate: string;
  supplier: string;
  initialRemaining: number;
  consumedQuantity: number;
  finalRemaining: number;
  unitCost: number;
  subtotalCost: number;
}

export interface ItemStockImpact {
  productId: string;
  productName: string;
  category: string;
  unit: string;
  minStock: number;
  currentStock: number;
  quantityToExit: number;
  resultingStock: number;
  status: 'OPTIMAL' | 'LOW_STOCK' | 'OUT_OF_STOCK';
  salePrice?: number;
  totalSaleRevenue?: number;
  totalCost: number;
  averageUnitCost: number;
  margin?: number;
  marginPercent?: number;
  batchConsumptions: BatchConsumption[];
  description?: string;
  error?: string;
}

export interface TransactionSimulationResult {
  type: 'EXIT' | 'SALE';
  items: ItemStockImpact[];
  totalQuantity: number;
  totalCost: number;
  totalRevenue?: number;
  totalMargin?: number;
  hasLowStockWarning: boolean;
  hasOutOfStockWarning: boolean;
  hasError: boolean;
  warningCount: number;
}

/**
 * Simulates the stock impact and FIFO batch consumption without mutating state.
 */
export function simulateExitImpact(
  itemsToExit: Array<{
    productId: string;
    quantity: number;
    salePrice?: number;
    description?: string;
  }>,
  allBatches: Batch[],
  products: Product[],
  type: 'EXIT' | 'SALE' = 'EXIT'
): TransactionSimulationResult {
  const impacts: ItemStockImpact[] = [];
  let totalQuantity = 0;
  let totalCost = 0;
  let totalRevenue = 0;
  let hasLowStockWarning = false;
  let hasOutOfStockWarning = false;
  let hasError = false;
  let warningCount = 0;

  // Track virtual batch states in case multiple items reference the same product
  const virtualBatches = allBatches.map(b => ({ ...b }));

  for (const item of itemsToExit) {
    const product = products.find(p => p.id === item.productId);
    const productName = product?.name || 'Producto Desconocido';
    const category = product?.category || 'General';
    const unit = product?.unit || 'unidad';
    const minStock = product?.minStock || 0;

    const availableBatches = virtualBatches
      .filter(b => b.productId === item.productId && b.remainingQuantity > 0)
      .sort((a, b) => {
        const timeDiff = new Date(a.entryDate).getTime() - new Date(b.entryDate).getTime();
        return timeDiff !== 0 ? timeDiff : a.id.localeCompare(b.id);
      });

    const currentStock = availableBatches.reduce((acc, b) => acc + b.remainingQuantity, 0);
    const quantityToExit = item.quantity;
    const resultingStock = Math.max(0, currentStock - quantityToExit);

    let status: 'OPTIMAL' | 'LOW_STOCK' | 'OUT_OF_STOCK' = 'OPTIMAL';
    if (resultingStock <= 0) {
      status = 'OUT_OF_STOCK';
      hasOutOfStockWarning = true;
      warningCount++;
    } else if (resultingStock <= minStock) {
      status = 'LOW_STOCK';
      hasLowStockWarning = true;
      warningCount++;
    }

    let remainingToProcess = quantityToExit;
    let itemTotalCost = 0;
    const batchConsumptions: BatchConsumption[] = [];
    let itemError: string | undefined = undefined;

    if (quantityToExit > currentStock) {
      itemError = `Stock insuficiente: Se solicitaron ${quantityToExit} ${unit}s pero solo hay ${currentStock} disponibles.`;
      hasError = true;
    } else {
      for (const batch of availableBatches) {
        if (remainingToProcess <= 0) break;
        const takeQty = Math.min(batch.remainingQuantity, remainingToProcess);
        const initialRemaining = batch.remainingQuantity;
        const finalRemaining = initialRemaining - takeQty;

        batch.remainingQuantity = finalRemaining; // Update virtual batch for multi-item simulation

        const subtotal = takeQty * batch.unitCost;
        itemTotalCost += subtotal;
        remainingToProcess -= takeQty;

        batchConsumptions.push({
          batchId: batch.id,
          entryDate: batch.entryDate,
          supplier: batch.supplier,
          initialRemaining,
          consumedQuantity: takeQty,
          finalRemaining,
          unitCost: batch.unitCost,
          subtotalCost: subtotal
        });
      }
    }

    const averageUnitCost = quantityToExit > 0 ? itemTotalCost / quantityToExit : 0;
    const salePrice = item.salePrice;
    const itemRevenue = (salePrice !== undefined && quantityToExit > 0) ? (salePrice * quantityToExit) : undefined;
    const margin = itemRevenue !== undefined ? itemRevenue - itemTotalCost : undefined;
    const marginPercent = (itemRevenue !== undefined && itemRevenue > 0) ? ((margin || 0) / itemRevenue) * 100 : undefined;

    totalQuantity += quantityToExit;
    totalCost += itemTotalCost;
    if (itemRevenue !== undefined) {
      totalRevenue += itemRevenue;
    }

    impacts.push({
      productId: item.productId,
      productName,
      category,
      unit,
      minStock,
      currentStock,
      quantityToExit,
      resultingStock,
      status,
      salePrice,
      totalSaleRevenue: itemRevenue,
      totalCost: itemTotalCost,
      averageUnitCost,
      margin,
      marginPercent,
      batchConsumptions,
      description: item.description,
      error: itemError
    });
  }

  const totalMargin = type === 'SALE' ? (totalRevenue - totalCost) : undefined;

  return {
    type,
    items: impacts,
    totalQuantity,
    totalCost,
    totalRevenue: type === 'SALE' ? totalRevenue : undefined,
    totalMargin,
    hasLowStockWarning,
    hasOutOfStockWarning,
    hasError,
    warningCount
  };
}

export function generateUnifiedKardex(
  products: Product[],
  batches: Batch[],
  transactions: Transaction[]
): KardexEntry[] {
  const kardex: KardexEntry[] = [];
  
  // Track balances per product
  const productBalances: Record<string, { qty: number, total: number }> = {};
  products.forEach(p => {
    productBalances[p.id] = { qty: 0, total: 0 };
  });

  const events = [
    ...batches.map(b => ({
      date: b.entryDate,
      type: (b.type || 'ENTRY') as any,
      productId: b.productId,
      qty: b.quantity,
      cost: b.unitCost,
      performedBy: b.performedBy,
      batchId: b.id,
      ref: b.description ? `${b.description} (Lote: ${b.id.slice(0, 8)}, ${b.supplier})` : `Lote: ${b.id.slice(0, 8)} (${b.supplier})`
    })),
    ...transactions.map(t => ({
      date: t.date,
      type: t.type,
      productId: t.productId,
      qty: t.quantity,
      cost: t.unitCost,
      salePrice: t.salePrice,
      performedBy: t.performedBy,
      batchId: t.batchId,
      ref: t.description ? `${t.description} (${t.reference})` : t.reference
    }))
  ].sort((a, b) => {
    const timeDiff = new Date(a.date).getTime() - new Date(b.date).getTime();
    if (timeDiff !== 0) return timeDiff;
    // Entries/Purchases should always take precedence over exits at the same timestamp
    const isEntryA = a.type === 'ENTRY' || a.type === 'PURCHASE';
    const isEntryB = b.type === 'ENTRY' || b.type === 'PURCHASE';
    if (isEntryA && !isEntryB) return -1;
    if (!isEntryA && isEntryB) return 1;
    return 0;
  });

  events.forEach(event => {
    const isEntry = event.type === 'ENTRY' || event.type === 'PURCHASE';
    const balance = productBalances[event.productId];
    if (!balance) return; // Skip if product doesn't exist anymore

    if (isEntry) {
      balance.qty += event.qty;
      balance.total += (event.qty * event.cost);
    } else {
      balance.qty -= event.qty;
      balance.total -= (event.qty * event.cost);
    }

    // Precise balance correction for zero/near-zero stock (Floating point guard)
    if (balance.qty <= 1e-9) {
      balance.qty = 0;
      balance.total = 0;
    } else if (balance.total <= 1e-9) {
      balance.total = 0;
    }

    // The Balance Unit Cost in PEPS is the weighted average of the remaining layers
    const balanceUnitCost = balance.qty > 0 ? (balance.total / balance.qty) : 0;
    const productName = products.find(p => p.id === event.productId)?.name || 'Unknown';

    kardex.push({
      date: event.date,
      reference: event.ref,
      type: event.type,
      productId: event.productId,
      batchId: event.batchId,
      productName: productName,
      entryQty: isEntry ? event.qty : 0,
      entryUnitCost: isEntry ? event.cost : 0,
      entryTotal: isEntry ? (event.qty * event.cost) : 0,
      exitQty: !isEntry ? event.qty : 0,
      exitUnitCost: !isEntry ? event.cost : 0,
      exitTotal: !isEntry ? (event.qty * event.cost) : 0,
      salePrice: (event as any).salePrice,
      balanceQty: Number(balance.qty.toFixed(4)),
      balanceUnitCost: Number(balanceUnitCost.toFixed(4)),
      balanceTotal: Number(balance.total.toFixed(2)),
      performedBy: event.performedBy
    });
  });

  return kardex;
}
