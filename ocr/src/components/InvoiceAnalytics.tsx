import { OCRHistoryEntry } from "../types";
import { formatCurrency } from "../utils";
import { TrendingUp, ShoppingBag, Landmark, PieChart } from "lucide-react";

interface InvoiceAnalyticsProps {
  history: OCRHistoryEntry[];
}

export function InvoiceAnalytics({ history }: InvoiceAnalyticsProps) {
  if (history.length === 0) {
    return (
      <div className="bg-stone-50 border border-stone-200 rounded-2xl p-6 text-center text-stone-500 text-sm">
        No hay datos suficientes para generar análisis. Sube tus primeras facturas para activar las estadísticas de consumo.
      </div>
    );
  }

  // Calculate high-level stats (forcing EUR as base calculation display, but flexible if specified)
  const totalSpend = history.reduce((sum, item) => sum + (item.invoiceData.totalAmount || 0), 0);
  const averageTicket = totalSpend / history.length;
  
  // Outstanding unique vendors
  const vendorsMap: Record<string, number> = {};
  history.forEach(item => {
    const vName = item.invoiceData.vendor?.name || "Proveedor no identificado";
    vendorsMap[vName] = (vendorsMap[vName] || 0) + (item.invoiceData.totalAmount || 0);
  });

  const sortedVendors = Object.entries(vendorsMap)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5); // top 5

  // Spends grouped by month
  const monthsMap: Record<string, number> = {};
  history.forEach(item => {
    const dateStr = item.invoiceData.invoiceDate;
    let monthKey = "Otros";
    if (dateStr && dateStr.includes("-")) {
      const parts = dateStr.split("-");
      if (parts.length >= 2) {
        // e.g. "2026-05"
        monthKey = `${parts[0]}-${parts[1]}`;
      }
    }
    monthsMap[monthKey] = (monthsMap[monthKey] || 0) + (item.invoiceData.totalAmount || 0);
  });

  const sortedMonths = Object.entries(monthsMap)
    .sort((a, b) => a[0].localeCompare(b[0]))
    .slice(-6); // last 6 months

  const maxMonthVal = Math.max(...Object.values(monthsMap), 1);
  const maxVendorVal = Math.max(...Object.values(vendorsMap), 1);

  return (
    <div className="space-y-6" id="analytics-section">
      {/* Mini Stats Badges */}
      <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
        <div className="bg-white p-5 border border-stone-200 rounded-2xl shadow-sm flex items-center gap-4">
          <div className="w-10 h-10 bg-emerald-50 rounded-xl flex items-center justify-center text-emerald-600 shrink-0">
            <TrendingUp className="w-5 h-5" />
          </div>
          <div>
            <span className="text-stone-400 text-[10px] font-semibold uppercase tracking-widest block">Gasto Acumulado</span>
            <span className="text-xl font-bold text-stone-900 font-sans block mt-0.5" id="val-total-gasto">
              {formatCurrency(totalSpend, "EUR")}
            </span>
          </div>
        </div>

        <div className="bg-white p-5 border border-stone-200 rounded-2xl shadow-sm flex items-center gap-4">
          <div className="w-10 h-10 bg-blue-50 rounded-xl flex items-center justify-center text-blue-600 shrink-0">
            <ShoppingBag className="w-5 h-5" />
          </div>
          <div>
            <span className="text-stone-400 text-[10px] font-semibold uppercase tracking-widest block">Ticket Promedio</span>
            <span className="text-xl font-bold text-stone-900 font-sans block mt-0.5">
              {formatCurrency(averageTicket, "EUR")}
            </span>
          </div>
        </div>

        <div className="bg-white p-5 border border-stone-200 rounded-2xl shadow-sm flex items-center gap-4">
          <div className="w-10 h-10 bg-purple-50 rounded-xl flex items-center justify-center text-purple-600 shrink-0">
            <Landmark className="w-5 h-5" />
          </div>
          <div>
            <span className="text-stone-400 text-[10px] font-semibold uppercase tracking-widest block">Nro. de Facturas</span>
            <span className="text-xl font-bold text-stone-900 font-sans block mt-0.5">
              {history.length} {history.length === 1 ? "documento" : "documentos"}
            </span>
          </div>
        </div>
      </div>

      {/* Structured Chart grids */}
      <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
        
        {/* CHART 1: MONTHLY BAR GRAPHS */}
        <div className="bg-white p-5 border border-stone-200 rounded-2xl shadow-sm">
          <span className="text-xs font-bold text-stone-700 uppercase tracking-widest flex items-center gap-2 mb-4">
            <PieChart className="w-4 h-4 text-emerald-600" /> Histórico Mensual de Gastos
          </span>
          
          <div className="flex items-end justify-between h-40 pt-4 gap-3 bg-stone-50/50 p-3 rounded-xl border border-stone-100">
            {sortedMonths.map(([month, amount]) => {
              const heightPercent = Math.max(10, Math.min(100, (amount / maxMonthVal) * 100));
              
              // format label "2026-05" to "Mayo 26" or similar
              const formatLabel = (val: string) => {
                const parts = val.split("-");
                if (parts.length < 2) return val;
                const monthNames = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"];
                const monthIdx = parseInt(parts[1]) - 1;
                return `${monthNames[monthIdx] || parts[1]} '${parts[0].slice(-2)}`;
              };

              return (
                <div key={month} className="flex-1 flex flex-col items-center group h-full justify-end">
                  {/* Tooltip on hover */}
                  <div className="opacity-0 group-hover:opacity-100 absolute bg-stone-900 text-stone-100 text-[10px] py-1 px-2 rounded -translate-y-12 transition-all pointer-events-none whitespace-nowrap z-10 shadow-md">
                    {formatCurrency(amount, "EUR")}
                  </div>
                  
                  {/* Bar block */}
                  <div 
                    style={{ height: `${heightPercent}%` }}
                    className="w-full bg-slate-900 hover:bg-emerald-600 rounded-t-md transition-all duration-300 shadow-sm"
                  />
                  <span className="text-[10px] text-stone-500 font-mono mt-2 text-center block leading-none select-none">
                    {formatLabel(month)}
                  </span>
                </div>
              );
            })}
          </div>
        </div>

        {/* CHART 2: TOP SPENDING BY VENDOR */}
        <div className="bg-white p-5 border border-stone-200 rounded-2xl shadow-sm flex flex-col justify-between">
          <div>
            <span className="text-xs font-bold text-stone-700 uppercase tracking-widest flex items-center gap-2 mb-4">
              <Landmark className="w-4 h-4 text-blue-600" /> Distribución por Proveedor
            </span>

            <div className="space-y-3">
              {sortedVendors.map(([vendor, amount], idx) => {
                const widthPercent = (amount / maxVendorVal) * 100;
                
                const barColors = [
                  "bg-emerald-500",
                  "bg-blue-500",
                  "bg-indigo-500",
                  "bg-amber-500",
                  "bg-stone-500"
                ];

                return (
                  <div key={idx} className="space-y-1">
                    <div className="flex justify-between text-xs text-stone-700 font-sans">
                      <span className="font-medium truncate max-w-[140px] sm:max-w-[200px]">{vendor}</span>
                      <span className="font-semibold text-stone-950">{formatCurrency(amount, "EUR")}</span>
                    </div>
                    <div className="w-full bg-stone-100 rounded-full h-2">
                      <div 
                        style={{ width: `${widthPercent}%` }}
                        className={`${barColors[idx % barColors.length]} h-full rounded-full transition-all duration-500`}
                      />
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
          
          <div className="text-[11px] text-stone-400 font-medium italic text-right mt-4">
            * Valores normalizados representados en base EUR
          </div>
        </div>

      </div>
    </div>
  );
}
