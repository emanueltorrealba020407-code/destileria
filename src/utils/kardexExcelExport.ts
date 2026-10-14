import * as XLSX from 'xlsx';
import { format } from 'date-fns';
import { Product, Batch, KardexEntry } from '../types';

export interface ExportKardexOptions {
  companyName?: string;
  userName?: string;
  filters?: {
    product?: string;
    type?: string;
    startDate?: string;
    endDate?: string;
    search?: string;
  };
  specificProduct?: Product | null;
  includeSummarySheet?: boolean;
  includeBatchesSheet?: boolean;
  fileNamePrefix?: string;
}

/**
 * Translates transaction type to clean business Spanish terminology
 */
function formatTransactionType(type: string): string {
  switch (type) {
    case 'ENTRY':
      return 'Entrada Manual';
    case 'EXIT':
      return 'Salida / Ajuste';
    case 'PURCHASE':
      return 'Compra Proveedor';
    case 'SALE':
      return 'Venta Comercial';
    default:
      return type;
  }
}

/**
 * Calculates optimal column widths for Excel sheets
 */
function calculateColumnWidths(data: any[][], customMinWidths?: Record<number, number>): { wch: number }[] {
  const colWidths: number[] = [];

  data.forEach(row => {
    row.forEach((cell, colIdx) => {
      const cellStr = cell !== null && cell !== undefined ? String(cell) : '';
      const cellLen = cellStr.length;
      if (!colWidths[colIdx] || cellLen > colWidths[colIdx]) {
        colWidths[colIdx] = cellLen;
      }
    });
  });

  return colWidths.map((width, idx) => {
    const minW = customMinWidths?.[idx] || 12;
    // Add extra padding for readability and cap maximum reasonable width
    return { wch: Math.min(Math.max(width + 3, minW), 55) };
  });
}

/**
 * Exports a comprehensive, beautifully structured PEPS (FIFO) Kardex Excel workbook
 */
export function exportKardexWorkbook(
  kardexEntries: KardexEntry[],
  products: Product[],
  batches: Batch[],
  options: ExportKardexOptions = {}
): void {
  const company = options.companyName || 'DESTILERÍA ORMÓN';
  const user = options.userName || 'Sistema';
  const now = new Date();
  const emissionDateStr = format(now, 'dd/MM/yyyy HH:mm:ss');
  const fileDateStr = format(now, 'yyyyMMdd_HHmm');
  const isSingleProduct = !!options.specificProduct;

  const workbook = XLSX.utils.book_new();

  // ==========================================
  // 1. SHEET 1: KARDEX VALORADO PEPS
  // ==========================================
  const filterInfo: string[] = [];
  if (options.specificProduct) {
    filterInfo.push(`Producto: ${options.specificProduct.name}`);
  } else if (options.filters?.product) {
    filterInfo.push(`Filtro Producto: "${options.filters.product}"`);
  }
  if (options.filters?.type && options.filters.type !== 'ALL') {
    filterInfo.push(`Tipo: ${formatTransactionType(options.filters.type)}`);
  }
  if (options.filters?.startDate || options.filters?.endDate) {
    filterInfo.push(`Rango: ${options.filters.startDate || 'Inicio'} hasta ${options.filters.endDate || 'Actualidad'}`);
  }
  if (options.filters?.search) {
    filterInfo.push(`Búsqueda: "${options.filters.search}"`);
  }
  const filterText = filterInfo.length > 0 ? filterInfo.join(' | ') : 'Todos los movimientos sin restricción';

  // Calculate totals
  const totalEntriesQty = kardexEntries.reduce((acc, e) => acc + (e.entryQty || 0), 0);
  const totalEntriesVal = kardexEntries.reduce((acc, e) => acc + (e.entryTotal || 0), 0);
  const totalExitsQty = kardexEntries.reduce((acc, e) => acc + (e.exitQty || 0), 0);
  const totalExitsVal = kardexEntries.reduce((acc, e) => acc + (e.exitTotal || 0), 0);
  
  // Calculate final balance
  let finalBalanceQty = 0;
  let finalBalanceTotal = 0;
  if (isSingleProduct) {
    const lastEntry = kardexEntries[kardexEntries.length - 1];
    finalBalanceQty = lastEntry ? lastEntry.balanceQty : 0;
    finalBalanceTotal = lastEntry ? lastEntry.balanceTotal : 0;
  } else {
    // Sum balances of all products
    products.forEach(p => {
      const prodEntries = kardexEntries.filter(e => e.productId === p.id);
      if (prodEntries.length > 0) {
        const last = prodEntries[prodEntries.length - 1];
        finalBalanceQty += last.balanceQty;
        finalBalanceTotal += last.balanceTotal;
      }
    });
  }

  const kardexSheetAOA: any[][] = [
    // Header Banner
    [`${company.toUpperCase()} - SISTEMA DE CONTROL DE INVENTARIOS PEPS (FIFO)`],
    [isSingleProduct 
      ? `TARJETA KARDEX INDIVIDUAL VALORADA - ${options.specificProduct?.name.toUpperCase()}${options.specificProduct?.type === 'producto terminado' ? ' (PRODUCTO TERMINADO)' : ''}`
      : 'LIBRO TARJETA KARDEX VALORADO PEPS • PRODUCTOS TERMINADOS'],
    [`Criterios del Reporte: ${filterText}`],
    [`Fecha y Hora de Emisión: ${emissionDateStr} | Responsable: ${user} | Registros: ${kardexEntries.length}`],
    [
      `RESUMEN DEL PERÍODO -> Entradas: ${totalEntriesQty.toLocaleString()} unid. (Bs. ${totalEntriesVal.toFixed(2)}) | ` +
      `Salidas: ${totalExitsQty.toLocaleString()} unid. (Bs. ${totalExitsVal.toFixed(2)}) | ` +
      `Saldo Valorado Final: ${finalBalanceQty.toLocaleString()} unid. (Bs. ${finalBalanceTotal.toFixed(2)})`
    ],
    [], // Blank spacing row

    // Grouping Header Row (Row 6)
    [
      'DATOS DEL MOVIMIENTO', '', '', '', '',
      'ENTRADAS / INGRESOS (PEPS)', '', '',
      'SALIDAS / EGRESOS (PEPS)', '', '', '',
      'SALDOS VALORADOS EN EXISTENCIA', '', '',
      'AUDITORÍA'
    ],

    // Column Headers (Row 7)
    [
      'Fecha y Hora',
      'Tipo de Operación',
      'Producto',
      'Unidad',
      'Referencia / Lote / Detalle',
      'Cantidad Ingreso',
      'Costo Unit. (Bs)',
      'Total Ingreso (Bs)',
      'Cantidad Salida',
      'Costo Unit. PEPS (Bs)',
      'Total Salida (Bs)',
      'Precio Venta (Bs)',
      'Cantidad Saldo',
      'Costo Unit. Valuado (Bs)',
      'Total Saldo (Bs)',
      'Responsable'
    ]
  ];

  // Data rows
  kardexEntries.forEach(entry => {
    const prod = products.find(p => p.id === entry.productId);
    const dateFormatted = entry.date ? format(new Date(entry.date), 'dd/MM/yyyy HH:mm') : '-';

    kardexSheetAOA.push([
      dateFormatted,
      formatTransactionType(entry.type),
      entry.productName || prod?.name || 'N/A',
      prod?.unit || 'unidad',
      entry.reference || '-',
      entry.entryQty > 0 ? entry.entryQty : '',
      entry.entryQty > 0 ? Number(entry.entryUnitCost.toFixed(2)) : '',
      entry.entryQty > 0 ? Number(entry.entryTotal.toFixed(2)) : '',
      entry.exitQty > 0 ? entry.exitQty : '',
      entry.exitQty > 0 ? Number(entry.exitUnitCost.toFixed(2)) : '',
      entry.exitQty > 0 ? Number(entry.exitTotal.toFixed(2)) : '',
      entry.salePrice ? Number(entry.salePrice.toFixed(2)) : '',
      entry.balanceQty,
      Number(entry.balanceUnitCost.toFixed(2)),
      Number(entry.balanceTotal.toFixed(2)),
      entry.performedBy || 'Sistema'
    ]);
  });

  // Final Summary Row
  kardexSheetAOA.push([
    'TOTALES GENERALES',
    '',
    '',
    '',
    `Total: ${kardexEntries.length} movimientos`,
    totalEntriesQty,
    '-',
    Number(totalEntriesVal.toFixed(2)),
    totalExitsQty,
    '-',
    Number(totalExitsVal.toFixed(2)),
    '-',
    finalBalanceQty,
    '-',
    Number(finalBalanceTotal.toFixed(2)),
    ''
  ]);

  const kardexWorksheet = XLSX.utils.aoa_to_sheet(kardexSheetAOA);

  // Column merges for Kardex Sheet
  kardexWorksheet['!merges'] = [
    // Top headers banner
    { s: { r: 0, c: 0 }, e: { r: 0, c: 15 } },
    { s: { r: 1, c: 0 }, e: { r: 1, c: 15 } },
    { s: { r: 2, c: 0 }, e: { r: 2, c: 15 } },
    { s: { r: 3, c: 0 }, e: { r: 3, c: 15 } },
    { s: { r: 4, c: 0 }, e: { r: 4, c: 15 } },
    // Group headers (Row 6)
    { s: { r: 6, c: 0 }, e: { r: 6, c: 4 } },   // DATOS DEL MOVIMIENTO
    { s: { r: 6, c: 5 }, e: { r: 6, c: 7 } },   // ENTRADAS
    { s: { r: 6, c: 8 }, e: { r: 6, c: 11 } },  // SALIDAS
    { s: { r: 6, c: 12 }, e: { r: 6, c: 14 } }, // SALDOS
    { s: { r: 6, c: 15 }, e: { r: 6, c: 15 } }, // AUDITORÍA
  ];

  // Specific column widths
  kardexWorksheet['!cols'] = [
    { wch: 18 }, // Fecha y Hora
    { wch: 18 }, // Tipo de Operación
    { wch: 28 }, // Producto
    { wch: 10 }, // Unidad
    { wch: 38 }, // Referencia / Detalle
    { wch: 16 }, // Cantidad Entrada
    { wch: 16 }, // Costo Unit. Entrada
    { wch: 18 }, // Total Entrada
    { wch: 16 }, // Cantidad Salida
    { wch: 18 }, // Costo Unit. PEPS Salida
    { wch: 18 }, // Total Salida
    { wch: 16 }, // Precio Venta
    { wch: 16 }, // Cantidad Saldo
    { wch: 20 }, // Costo Unit. Saldo
    { wch: 20 }, // Total Saldo
    { wch: 18 }  // Responsable
  ];

  XLSX.utils.book_append_sheet(workbook, kardexWorksheet, isSingleProduct ? 'Tarjeta Kardex' : 'Tarjeta Kardex PEPS');

  // ==========================================
  // 2. SHEET 2: RESUMEN DE INVENTARIO Y SALDOS
  // ==========================================
  if (options.includeSummarySheet !== false) {
    const summaryProducts = isSingleProduct && options.specificProduct 
      ? [options.specificProduct] 
      : products.sort((a, b) => a.name.localeCompare(b.name));

    const inventorySheetAOA: any[][] = [
      [`${company.toUpperCase()} - RESUMEN CONSOLIDADO DE INVENTARIO`],
      [`Emisión: ${emissionDateStr} | Método de Valuación: PEPS (Primeros en Entrar, Primeros en Salir)`],
      [],
      [
        'Código / ID',
        'Nombre del Producto',
        'Categoría',
        'Tipo de Ítem',
        'Unidad de Medida',
        'Stock Mínimo',
        'Existencia Actual',
        'Estado de Stock',
        'Valor Total en Inventario (Bs)',
        'Costo Promedio PEPS (Bs)',
        'Lotes Activos',
        'Total Entradas Históricas',
        'Total Salidas Históricas'
      ]
    ];

    let grandTotalStock = 0;
    let grandTotalValuation = 0;

    summaryProducts.forEach(p => {
      // Calculate active batches and stock
      const prodBatches = batches.filter(b => b.productId === p.id && b.remainingQuantity > 0);
      const currentStock = prodBatches.reduce((acc, b) => acc + b.remainingQuantity, 0);
      const totalValuation = prodBatches.reduce((acc, b) => acc + (b.remainingQuantity * b.unitCost), 0);
      const avgCost = currentStock > 0 ? totalValuation / currentStock : 0;

      // Status
      let status = 'ÓPTIMO';
      if (currentStock === 0) {
        status = 'AGOTADO (Stock 0)';
      } else if (currentStock <= p.minStock) {
        status = 'STOCK BAJO / CRÍTICO';
      }

      // Historical stats
      const prodEntries = kardexEntries.filter(e => e.productId === p.id);
      const histEntriesQty = prodEntries.reduce((acc, e) => acc + (e.entryQty || 0), 0);
      const histExitsQty = prodEntries.reduce((acc, e) => acc + (e.exitQty || 0), 0);

      grandTotalStock += currentStock;
      grandTotalValuation += totalValuation;

      inventorySheetAOA.push([
        p.id.slice(0, 8),
        p.name,
        p.category || 'General',
        p.type || 'Materia Prima',
        p.unit || 'unidad',
        p.minStock || 0,
        currentStock,
        status,
        Number(totalValuation.toFixed(2)),
        Number(avgCost.toFixed(2)),
        prodBatches.length,
        histEntriesQty,
        histExitsQty
      ]);
    });

    // Summary row
    inventorySheetAOA.push([
      'TOTAL GENERAL',
      `${summaryProducts.length} productos`,
      '-',
      '-',
      '-',
      '-',
      grandTotalStock,
      '-',
      Number(grandTotalValuation.toFixed(2)),
      '-',
      batches.filter(b => b.remainingQuantity > 0).length,
      '-',
      '-'
    ]);

    const inventoryWorksheet = XLSX.utils.aoa_to_sheet(inventorySheetAOA);
    inventoryWorksheet['!merges'] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: 12 } },
      { s: { r: 1, c: 0 }, e: { r: 1, c: 12 } }
    ];
    inventoryWorksheet['!cols'] = calculateColumnWidths(inventorySheetAOA, {
      0: 12,
      1: 28,
      2: 18,
      3: 18,
      4: 12,
      5: 14,
      6: 16,
      7: 20,
      8: 24,
      9: 22,
      10: 14,
      11: 20,
      12: 20
    });

    XLSX.utils.book_append_sheet(workbook, inventoryWorksheet, 'Resumen Inventario');
  }

  // ==========================================
  // 3. SHEET 3: DESGLOSE DE LOTES ACTIVOS (PEPS)
  // ==========================================
  if (options.includeBatchesSheet !== false) {
    const activeBatches = batches
      .filter(b => {
        if (b.remainingQuantity <= 0) return false;
        if (isSingleProduct && options.specificProduct && b.productId !== options.specificProduct.id) return false;
        return true;
      })
      .sort((a, b) => {
        // Group by product name, then by entryDate ASC (FIFO order)
        const prodA = products.find(p => p.id === a.productId)?.name || '';
        const prodB = products.find(p => p.id === b.productId)?.name || '';
        if (prodA !== prodB) return prodA.localeCompare(prodB);
        return new Date(a.entryDate).getTime() - new Date(b.entryDate).getTime();
      });

    const batchesSheetAOA: any[][] = [
      [`${company.toUpperCase()} - DESGLOSE DE LOTES ACTIVOS EN EXISTENCIA (PEPS)`],
      [`Emisión: ${emissionDateStr} | Orden de Consumo: Primeros Lotes en Entrar son los Primeros a Salir`],
      [],
      [
        'ID Lote',
        'Producto',
        'Fecha de Ingreso',
        'Proveedor / Origen',
        'Cantidad Inicial',
        'Saldo Disponible',
        '% Disponible',
        'Costo Unitario (Bs)',
        'Valor Inmovilizado (Bs)',
        'Registrado Por',
        'Detalle / Observaciones'
      ]
    ];

    let totalActiveUnits = 0;
    let totalActiveValuation = 0;

    activeBatches.forEach(b => {
      const prod = products.find(p => p.id === b.productId);
      const percentAvailable = b.quantity > 0 ? Math.round((b.remainingQuantity / b.quantity) * 100) : 0;
      const subtotalValuation = b.remainingQuantity * b.unitCost;

      totalActiveUnits += b.remainingQuantity;
      totalActiveValuation += subtotalValuation;

      batchesSheetAOA.push([
        b.id.slice(0, 8),
        prod?.name || 'N/A',
        format(new Date(b.entryDate), 'dd/MM/yyyy HH:mm'),
        b.supplier || 'N/A',
        b.quantity,
        b.remainingQuantity,
        `${percentAvailable}%`,
        Number(b.unitCost.toFixed(2)),
        Number(subtotalValuation.toFixed(2)),
        b.performedBy || 'Sistema',
        b.description || '-'
      ]);
    });

    batchesSheetAOA.push([
      'TOTAL LOTES ACTIVOS',
      `${activeBatches.length} lotes disponibles`,
      '-',
      '-',
      '-',
      totalActiveUnits,
      '-',
      '-',
      Number(totalActiveValuation.toFixed(2)),
      '-',
      ''
    ]);

    const batchesWorksheet = XLSX.utils.aoa_to_sheet(batchesSheetAOA);
    batchesWorksheet['!merges'] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: 10 } },
      { s: { r: 1, c: 0 }, e: { r: 1, c: 10 } }
    ];
    batchesWorksheet['!cols'] = calculateColumnWidths(batchesSheetAOA, {
      0: 12,
      1: 26,
      2: 18,
      3: 22,
      4: 16,
      5: 16,
      6: 14,
      7: 18,
      8: 22,
      9: 18,
      10: 30
    });

    XLSX.utils.book_append_sheet(workbook, batchesWorksheet, 'Lotes Activos PEPS');
  }

  // Generate file name
  const safeCompanyName = company.replace(/[^a-zA-Z0-9]/g, '_').toLowerCase();
  const specificName = isSingleProduct && options.specificProduct 
    ? `_${options.specificProduct.name.replace(/[^a-zA-Z0-9]/g, '_').toLowerCase()}` 
    : '';
  const prefix = options.fileNamePrefix || `kardex_peps_${safeCompanyName}`;
  const filename = `${prefix}${specificName}_${fileDateStr}.xlsx`;

  XLSX.writeFile(workbook, filename);
}
