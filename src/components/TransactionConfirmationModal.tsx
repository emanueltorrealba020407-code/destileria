import React, { useState } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import { 
  AlertTriangle, 
  CheckCircle2, 
  ShieldAlert, 
  Layers, 
  ArrowRight, 
  DollarSign, 
  Package, 
  Clock, 
  TrendingUp, 
  ChevronDown, 
  ChevronUp, 
  CheckSquare, 
  Square, 
  AlertOctagon, 
  Info, 
  X,
  Loader2,
  Calendar
} from 'lucide-react';
import { TransactionSimulationResult } from '../logic/peps';

interface TransactionConfirmationModalProps {
  isOpen: boolean;
  simulation: TransactionSimulationResult | null;
  movementDate?: string;
  operatorName: string;
  notes?: string;
  isSubmitting?: boolean;
  onConfirm: () => void;
  onCancel: () => void;
  onBackToEdit: () => void;
}

export const TransactionConfirmationModal: React.FC<TransactionConfirmationModalProps> = ({
  isOpen,
  simulation,
  movementDate,
  operatorName,
  notes,
  isSubmitting = false,
  onConfirm,
  onCancel,
  onBackToEdit
}) => {
  const [expandedBatches, setExpandedBatches] = useState<Record<string, boolean>>({});
  const [hasAcknowledged, setHasAcknowledged] = useState(false);
  const [hasAcknowledgedCritical, setHasAcknowledgedCritical] = useState(false);

  if (!isOpen || !simulation) return null;

  const isSale = simulation.type === 'SALE';
  const isMultiItem = simulation.items.length > 1;
  const requiresCriticalAck = simulation.hasOutOfStockWarning || simulation.hasLowStockWarning;
  const canSubmit = hasAcknowledged && (!requiresCriticalAck || hasAcknowledgedCritical) && !simulation.hasError && !isSubmitting;

  const toggleBatchAccordion = (productId: string) => {
    setExpandedBatches(prev => ({
      ...prev,
      [productId]: !prev[productId]
    }));
  };

  return (
    <div className="fixed inset-0 bg-black/80 backdrop-blur-md z-50 flex items-center justify-center p-3 sm:p-6 overflow-y-auto">
      <motion.div
        initial={{ opacity: 0, scale: 0.96, y: 10 }}
        animate={{ opacity: 1, scale: 1, y: 0 }}
        exit={{ opacity: 0, scale: 0.96, y: 10 }}
        transition={{ duration: 0.2 }}
        className="bg-card w-full max-w-4xl rounded-3xl p-6 sm:p-8 shadow-2xl border border-main overflow-hidden flex flex-col max-h-[92vh] my-auto"
      >
        {/* Header */}
        <div className="flex justify-between items-start pb-5 border-b border-main/20">
          <div>
            <div className="flex items-center gap-2 mb-2">
              <span className={`text-[10px] font-black uppercase tracking-[0.2em] px-2.5 py-1 rounded-full border ${
                isSale 
                  ? 'bg-accent/10 text-accent border-accent/20' 
                  : 'bg-amber-500/10 text-amber-500 border-amber-500/20'
              }`}>
                {isSale ? (isMultiItem ? 'Venta Masiva (Múltiples Productos)' : 'Venta Comercial') : (isMultiItem ? 'Salida Masiva de Inventario' : 'Ajuste de Salida / Merma')}
              </span>
              <span className="text-[10px] font-black uppercase tracking-wider text-muted bg-main/5 px-2 py-0.5 rounded">
                Paso 2 de 2: Verificación de Impacto
              </span>
            </div>
            <h2 className="text-2xl sm:text-3xl font-black uppercase tracking-tight text-main">
              Confirmar {isSale ? 'Transacción de Venta' : 'Salida de Inventario'}
            </h2>
            <p className="text-xs text-muted font-bold mt-1">
              Revise cuidadosamente el impacto proyectado en los saldos físicos y la valuación PEPS antes de asentar la operación.
            </p>
          </div>
          <button
            onClick={onCancel}
            disabled={isSubmitting}
            className="p-2 hover:bg-main/5 text-muted hover:text-main rounded-xl transition-all"
            title="Cerrar"
          >
            <X size={20} />
          </button>
        </div>

        {/* Scrollable Content */}
        <div className="flex-1 overflow-y-auto py-5 pr-1 space-y-6 custom-scrollbar">
          
          {/* Metadata banner */}
          <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 bg-main/5 p-4 rounded-2xl border border-main/10 text-xs">
            <div className="flex items-center gap-2.5">
              <Calendar size={16} className="text-accent" />
              <div>
                <div className="text-[9px] font-black text-muted uppercase tracking-wider">Fecha de Movimiento</div>
                <div className="font-bold text-main">{movementDate || new Date().toISOString().split('T')[0]}</div>
              </div>
            </div>
            <div className="flex items-center gap-2.5">
              <Package size={16} className="text-accent" />
              <div>
                <div className="text-[9px] font-black text-muted uppercase tracking-wider">Responsable</div>
                <div className="font-bold text-main truncate max-w-[150px]">{operatorName}</div>
              </div>
            </div>
            <div className="flex items-center gap-2.5">
              <Layers size={16} className="text-accent" />
              <div>
                <div className="text-[9px] font-black text-muted uppercase tracking-wider">Productos Afectados</div>
                <div className="font-bold text-main">{simulation.items.length} {simulation.items.length === 1 ? 'producto' : 'productos'}</div>
              </div>
            </div>
            {notes && (
              <div className="col-span-full pt-2 border-t border-main/10 text-[11px] text-muted">
                <span className="font-black uppercase tracking-wider text-main/80 mr-1.5">Motivo / Ref:</span>
                <span className="italic">{notes}</span>
              </div>
            )}
          </div>

          {/* High Priority Alerts (Low stock or Out of stock) */}
          {simulation.hasOutOfStockWarning && (
            <div className="bg-red-500/10 border border-red-500/30 rounded-2xl p-4 flex items-start gap-3">
              <AlertOctagon size={20} className="text-red-500 shrink-0 mt-0.5" />
              <div>
                <h4 className="text-xs font-black uppercase tracking-widest text-red-500 mb-1">
                  Atención: Existencias Agotadas (Stock 0)
                </h4>
                <p className="text-xs text-muted font-bold leading-relaxed">
                  Uno o más productos quedarán completamente sin existencias tras ejecutar esta operación.
                </p>
              </div>
            </div>
          )}

          {simulation.hasLowStockWarning && !simulation.hasOutOfStockWarning && (
            <div className="bg-amber-500/10 border border-amber-500/30 rounded-2xl p-4 flex items-start gap-3">
              <AlertTriangle size={20} className="text-amber-500 shrink-0 mt-0.5" />
              <div>
                <h4 className="text-xs font-black uppercase tracking-widest text-amber-500 mb-1">
                  Alerta: Nivel de Stock Mínimo Alcanzado
                </h4>
                <p className="text-xs text-muted font-bold leading-relaxed">
                  El stock proyectado descenderá por debajo del umbral de seguridad configurado.
                </p>
              </div>
            </div>
          )}

          {simulation.hasError && (
            <div className="bg-red-600 text-white rounded-2xl p-4 flex items-start gap-3 shadow-lg">
              <ShieldAlert size={22} className="shrink-0 mt-0.5" />
              <div>
                <h4 className="text-xs font-black uppercase tracking-widest mb-1">
                  Error de Validación Operativa
                </h4>
                <p className="text-xs font-bold leading-relaxed opacity-90">
                  No se puede procesar esta transacción porque una o más cantidades exceden el inventario disponible.
                </p>
              </div>
            </div>
          )}

          {/* KPI Summary Cards */}
          <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
            <div className="bg-sidebar p-4 rounded-2xl border border-main">
              <div className="text-[9px] font-black uppercase tracking-widest text-muted mb-1">Total a Descontar</div>
              <div className="text-xl sm:text-2xl font-black text-main">{simulation.totalQuantity}</div>
              <div className="text-[9px] font-bold text-muted uppercase">Unidades físicas</div>
            </div>

            <div className="bg-sidebar p-4 rounded-2xl border border-main">
              <div className="text-[9px] font-black uppercase tracking-widest text-muted mb-1">Costo de Salida (PEPS)</div>
              <div className="text-xl sm:text-2xl font-black text-amber-500">Bs. {simulation.totalCost.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</div>
              <div className="text-[9px] font-bold text-muted uppercase">Valuación contable</div>
            </div>

            {isSale && (
              <>
                <div className="bg-sidebar p-4 rounded-2xl border border-main">
                  <div className="text-[9px] font-black uppercase tracking-widest text-muted mb-1">Total Venta</div>
                  <div className="text-xl sm:text-2xl font-black text-accent">Bs. {(simulation.totalRevenue || 0).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</div>
                  <div className="text-[9px] font-bold text-muted uppercase">Ingreso bruto</div>
                </div>

                <div className="bg-sidebar p-4 rounded-2xl border border-main">
                  <div className="text-[9px] font-black uppercase tracking-widest text-muted mb-1">Margen Bruto Est.</div>
                  <div className={`text-xl sm:text-2xl font-black ${(simulation.totalMargin || 0) >= 0 ? 'text-emerald-500' : 'text-red-500'}`}>
                    Bs. {(simulation.totalMargin || 0).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                  </div>
                  <div className="text-[9px] font-bold text-muted uppercase">
                    {simulation.totalRevenue && simulation.totalRevenue > 0 
                      ? `${(((simulation.totalMargin || 0) / simulation.totalRevenue) * 100).toFixed(1)}% margen` 
                      : '0%'}
                  </div>
                </div>
              </>
            )}
          </div>

          {/* Detailed Product Breakdown */}
          <div className="space-y-4">
            <div className="flex items-center justify-between">
              <h3 className="text-xs font-black uppercase tracking-[0.2em] text-muted flex items-center gap-2">
                <Package size={14} className="text-accent" />
                Desglose del Impacto por Producto ({simulation.items.length})
              </h3>
              <span className="text-[10px] font-bold text-muted">Haga clic en un producto para ver los lotes PEPS</span>
            </div>

            <div className="space-y-3">
              {simulation.items.map((item, idx) => {
                const isExpanded = expandedBatches[item.productId] ?? true;
                const percentLeft = item.currentStock > 0 
                  ? Math.max(0, Math.min(100, Math.round((item.resultingStock / item.currentStock) * 100))) 
                  : 0;

                return (
                  <div 
                    key={item.productId}
                    className={`bg-sidebar rounded-2xl border transition-all overflow-hidden ${
                      item.error ? 'border-red-500/50' : 
                      item.status === 'OUT_OF_STOCK' ? 'border-red-500/30' : 
                      item.status === 'LOW_STOCK' ? 'border-amber-500/30' : 'border-main'
                    }`}
                  >
                    {/* Item Main Row */}
                    <div className="p-4 sm:p-5">
                      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 mb-3">
                        <div>
                          <div className="flex items-center gap-2">
                            <span className="text-[10px] font-black text-accent uppercase tracking-wider">{item.category}</span>
                            <span className="text-muted/40">•</span>
                            <span className="text-[10px] font-bold text-muted uppercase">Mínimo: {item.minStock} {item.unit}s</span>
                          </div>
                          <h4 className="text-base sm:text-lg font-black uppercase text-main">{item.productName}</h4>
                        </div>

                        {/* Status Badge */}
                        <div>
                          {item.error ? (
                            <span className="bg-red-500/20 text-red-500 border border-red-500/30 text-[10px] font-black uppercase tracking-wider px-2.5 py-1 rounded-lg">
                              Stock Insuficiente
                            </span>
                          ) : item.status === 'OUT_OF_STOCK' ? (
                            <span className="bg-red-500/10 text-red-500 border border-red-500/20 text-[10px] font-black uppercase tracking-wider px-2.5 py-1 rounded-lg flex items-center gap-1">
                              <AlertOctagon size={12} /> Agotado (0 {item.unit}s)
                            </span>
                          ) : item.status === 'LOW_STOCK' ? (
                            <span className="bg-amber-500/10 text-amber-500 border border-amber-500/20 text-[10px] font-black uppercase tracking-wider px-2.5 py-1 rounded-lg flex items-center gap-1">
                              <AlertTriangle size={12} /> Stock Crítico
                            </span>
                          ) : (
                            <span className="bg-emerald-500/10 text-emerald-500 border border-emerald-500/20 text-[10px] font-black uppercase tracking-wider px-2.5 py-1 rounded-lg flex items-center gap-1">
                              <CheckCircle2 size={12} /> Stock Óptimo
                            </span>
                          )}
                        </div>
                      </div>

                      {/* Stock Impact Visualizer */}
                      <div className="grid grid-cols-3 gap-2 bg-main/5 p-3 rounded-xl border border-main/10 mb-3 text-center">
                        <div>
                          <div className="text-[9px] font-black uppercase text-muted tracking-wider">Stock Actual</div>
                          <div className="text-sm sm:text-base font-black text-main">{item.currentStock} <span className="text-[10px] font-normal text-muted">{item.unit}s</span></div>
                        </div>

                        <div className="flex flex-col items-center justify-center">
                          <div className="text-[9px] font-black uppercase text-amber-500 tracking-wider">A Descontar</div>
                          <div className="text-sm sm:text-base font-black text-amber-500 flex items-center justify-center gap-1">
                            -{item.quantityToExit} <span className="text-[10px] font-normal">{item.unit}s</span>
                          </div>
                        </div>

                        <div>
                          <div className="text-[9px] font-black uppercase text-muted tracking-wider">Stock Resultante</div>
                          <div className={`text-sm sm:text-base font-black ${
                            item.resultingStock === 0 ? 'text-red-500' : 
                            item.resultingStock <= item.minStock ? 'text-amber-500' : 'text-emerald-500'
                          }`}>
                            {item.resultingStock} <span className="text-[10px] font-normal text-muted">{item.unit}s</span>
                          </div>
                        </div>
                      </div>

                      {/* Stock Meter */}
                      <div className="w-full bg-main/10 h-2 rounded-full overflow-hidden mb-3">
                        <div 
                          className={`h-full transition-all duration-300 ${
                            item.resultingStock === 0 ? 'bg-red-500' :
                            item.resultingStock <= item.minStock ? 'bg-amber-500' : 'bg-emerald-500'
                          }`}
                          style={{ width: `${percentLeft}%` }}
                        />
                      </div>

                      {/* Financial info per item */}
                      <div className="flex flex-wrap items-center justify-between text-xs font-bold pt-2 border-t border-main/10">
                        <div className="text-muted">
                          Costo PEPS: <span className="font-black text-main">Bs. {item.totalCost.toFixed(2)}</span>
                          <span className="text-[10px] text-muted/70 ml-1">(Promedio: Bs. {item.averageUnitCost.toFixed(2)}/u)</span>
                        </div>

                        {isSale && item.salePrice !== undefined && (
                          <div className="text-accent flex items-center gap-2">
                            <span>Venta: <strong className="text-main font-black">Bs. {(item.totalSaleRevenue || 0).toFixed(2)}</strong> (Bs. {item.salePrice.toFixed(2)}/u)</span>
                            <span className={`text-[10px] px-1.5 py-0.5 rounded ${(item.margin || 0) >= 0 ? 'bg-emerald-500/10 text-emerald-500' : 'bg-red-500/10 text-red-500'}`}>
                              Margen: Bs. {(item.margin || 0).toFixed(2)} ({item.marginPercent?.toFixed(1)}%)
                            </span>
                          </div>
                        )}
                      </div>

                      {/* Batch Breakdown Accordion Button */}
                      {item.batchConsumptions.length > 0 && (
                        <button
                          type="button"
                          onClick={() => toggleBatchAccordion(item.productId)}
                          className="mt-3 w-full flex items-center justify-between p-2 rounded-lg bg-main/5 hover:bg-main/10 text-[10px] font-black uppercase tracking-wider text-muted hover:text-main transition-all"
                        >
                          <span className="flex items-center gap-1.5">
                            <Layers size={13} className="text-accent" />
                            Desglose de Lotes PEPS a Consumir ({item.batchConsumptions.length} {item.batchConsumptions.length === 1 ? 'lote' : 'lotes'})
                          </span>
                          {isExpanded ? <ChevronUp size={14} /> : <ChevronDown size={14} />}
                        </button>
                      )}

                      {/* Batch Breakdown Content */}
                      <AnimatePresence>
                        {isExpanded && item.batchConsumptions.length > 0 && (
                          <motion.div
                            initial={{ opacity: 0, height: 0 }}
                            animate={{ opacity: 1, height: 'auto' }}
                            exit={{ opacity: 0, height: 0 }}
                            className="mt-2 space-y-1.5 overflow-hidden"
                          >
                            <div className="grid grid-cols-12 gap-1 text-[9px] font-black uppercase tracking-wider text-muted/70 px-2 py-1 bg-main/5 rounded">
                              <span className="col-span-3">Lote ID</span>
                              <span className="col-span-3">Proveedor</span>
                              <span className="col-span-2 text-center">Extracción</span>
                              <span className="col-span-2 text-center">Saldo Lote</span>
                              <span className="col-span-2 text-right">Subtotal</span>
                            </div>

                            {item.batchConsumptions.map((batch, bIdx) => (
                              <div 
                                key={batch.batchId} 
                                className="grid grid-cols-12 gap-1 items-center text-[10px] font-bold p-2 bg-main/5 rounded-lg border border-main/10"
                              >
                                <div className="col-span-3 truncate">
                                  <span className="text-accent font-black mr-1">#{bIdx + 1}</span>
                                  <span className="font-mono text-main">{batch.batchId.slice(0, 7)}</span>
                                </div>
                                <div className="col-span-3 truncate text-muted text-[9px]">
                                  {batch.supplier}
                                </div>
                                <div className="col-span-2 text-center font-black text-amber-500">
                                  -{batch.consumedQuantity}
                                </div>
                                <div className="col-span-2 text-center text-muted font-mono text-[9px]">
                                  {batch.finalRemaining} de {batch.initialRemaining}
                                </div>
                                <div className="col-span-2 text-right font-black text-main">
                                  Bs. {batch.subtotalCost.toFixed(2)}
                                </div>
                              </div>
                            ))}
                          </motion.div>
                        )}
                      </AnimatePresence>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>

          {/* Operational Verification & Explicit Confirmation Checkboxes */}
          <div className="bg-main/5 p-5 rounded-2xl border border-main/20 space-y-3">
            <h4 className="text-xs font-black uppercase tracking-[0.2em] text-accent flex items-center gap-2">
              <ShieldAlert size={16} />
              Protocolo de Seguridad Operativa
            </h4>

            {/* Checkbox 1: General Verification */}
            <label className="flex items-start gap-3 cursor-pointer select-none group">
              <input
                type="checkbox"
                checked={hasAcknowledged}
                onChange={(e) => setHasAcknowledged(e.target.checked)}
                className="sr-only"
              />
              <div className={`mt-0.5 w-5 h-5 rounded-md border flex items-center justify-center transition-all ${
                hasAcknowledged 
                  ? 'bg-accent border-accent text-white shadow-md shadow-brand-gold-shadow' 
                  : 'bg-card border-main group-hover:border-accent'
              }`}>
                {hasAcknowledged && <CheckSquare size={14} />}
              </div>
              <div className="text-xs font-bold text-main leading-snug">
                He verificado el impacto en stock, los precios y confirmo que las cantidades y lotes PEPS a descontar son correctos.
              </div>
            </label>

            {/* Checkbox 2: Critical stock warning acknowledgment (if low stock / depleted) */}
            {requiresCriticalAck && (
              <label className="flex items-start gap-3 cursor-pointer select-none group pt-2 border-t border-main/10">
                <input
                  type="checkbox"
                  checked={hasAcknowledgedCritical}
                  onChange={(e) => setHasAcknowledgedCritical(e.target.checked)}
                  className="sr-only"
                />
                <div className={`mt-0.5 w-5 h-5 rounded-md border flex items-center justify-center transition-all ${
                  hasAcknowledgedCritical 
                    ? 'bg-red-500 border-red-500 text-white shadow-md' 
                    : 'bg-card border-main group-hover:border-red-500'
                }`}>
                  {hasAcknowledgedCritical && <CheckSquare size={14} />}
                </div>
                <div className="text-xs font-bold text-red-500 leading-snug">
                  Entiendo que esta operación llevará uno o más productos a nivel crítico o agotado (Stock 0).
                </div>
              </label>
            )}
          </div>
        </div>

        {/* Footer Actions */}
        <div className="pt-4 border-t border-main/20 flex flex-col sm:flex-row gap-3">
          <button
            type="button"
            onClick={onBackToEdit}
            disabled={isSubmitting}
            className="flex-1 bg-main/5 hover:bg-main/10 text-muted hover:text-main py-4 rounded-xl font-black border border-main uppercase tracking-widest text-xs transition-all disabled:opacity-50"
          >
            Modificar Datos
          </button>

          <button
            type="button"
            onClick={onConfirm}
            disabled={!canSubmit}
            className={`flex-[2] py-4 rounded-xl font-black uppercase tracking-widest text-xs transition-all flex items-center justify-center gap-2 ${
              canSubmit
                ? 'bg-accent text-white hover:brightness-110 shadow-xl shadow-brand-gold-shadow active:scale-98'
                : 'bg-main/20 text-muted cursor-not-allowed border border-main/10'
            }`}
          >
            {isSubmitting ? (
              <>
                <Loader2 size={16} className="animate-spin" />
                <span>Procesando Transacción...</span>
              </>
            ) : (
              <>
                <CheckCircle2 size={16} />
                <span>Confirmar y Ejecutar {isSale ? 'Venta' : 'Salida'}</span>
              </>
            )}
          </button>
        </div>
      </motion.div>
    </div>
  );
};
