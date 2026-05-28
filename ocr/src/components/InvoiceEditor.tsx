import { useState, useEffect } from "react";
import { InvoiceData, LineItem, TaxLine } from "../types";
import { formatCurrency, downloadJson, exportInvoiceToCSV } from "../utils";
import { 
  Building2, User, Calendar, Percent, 
  FileJson, FileDown, Plus, Trash2, Check, Sparkles, AlertTriangle 
} from "lucide-react";

interface InvoiceEditorProps {
  initialData: InvoiceData;
  onSave: (updatedData: InvoiceData) => void;
  fileName: string;
}

type ActiveTab = "general" | "items" | "taxes" | "json";

export function InvoiceEditor({ initialData, onSave, fileName }: InvoiceEditorProps) {
  const [data, setData] = useState<InvoiceData>(initialData);
  const [activeTab, setActiveTab] = useState<ActiveTab>("general");
  const [showSavedFeedback, setShowSavedFeedback] = useState(false);

  // Sync state if initial data changes (e.g., loaded from history or new OCR)
  useEffect(() => {
    setData(initialData);
  }, [initialData]);

  const handleFieldChange = (section: string | null, field: string, value: any) => {
    setData(prev => {
      const updated = { ...prev };
      if (section) {
        // @ts-ignore
        updated[section] = {
          // @ts-ignore
          ...updated[section],
          [field]: value
        };
      } else {
        // @ts-ignore
        updated[field] = value;
      }
      return updated;
    });
  };

  const handleLineItemChange = (index: number, field: keyof LineItem, value: any) => {
    setData(prev => {
      const updatedItems = [...prev.lineItems];
      let numVal = value;
      if (field === "quantity" || field === "unitPrice" || field === "taxRate") {
        numVal = parseFloat(value) || 0;
      }
      
      updatedItems[index] = {
        ...updatedItems[index],
        [field]: numVal
      };

      // Recalculate amount if price or quantity changed
      if (field === "quantity" || field === "unitPrice") {
        updatedItems[index].amount = updatedItems[index].quantity * updatedItems[index].unitPrice;
      }

      return {
        ...prev,
        lineItems: updatedItems
      };
    });
  };

  const addLineItem = () => {
    setData(prev => ({
      ...prev,
      lineItems: [
        ...prev.lineItems,
        { description: "Nuevo concepto", quantity: 1, unitPrice: 0, taxRate: 21, amount: 0 }
      ]
    }));
  };

  const removeLineItem = (index: number) => {
    setData(prev => ({
      ...prev,
      lineItems: prev.lineItems.filter((_, i) => i !== index)
    }));
  };

  // Automatically recalculate subtotal, taxTotal and totals based on lineItems
  const recalculateTotals = () => {
    const subtotal = data.lineItems.reduce((acc, item) => acc + item.amount, 0);
    
    // Group taxes by tax rate
    const taxesMap: Record<number, { taxable: number, tax: number }> = {};
    data.lineItems.forEach(item => {
      const rate = item.taxRate;
      const amount = item.amount;
      const taxComponent = amount * (rate / 100);
      
      if (!taxesMap[rate]) {
        taxesMap[rate] = { taxable: 0, tax: 0 };
      }
      taxesMap[rate].taxable += amount;
      taxesMap[rate].tax += taxComponent;
    });

    const taxesArray: TaxLine[] = Object.entries(taxesMap).map(([rateStr, val]) => ({
      taxRate: parseFloat(rateStr),
      taxableAmount: Math.round(val.taxable * 100) / 100,
      taxAmount: Math.round(val.tax * 100) / 100
    }));

    const taxTotal = taxesArray.reduce((acc, t) => acc + t.taxAmount, 0);
    const totalAmount = subtotal + taxTotal - (data.discount || 0);

    setData(prev => ({
      ...prev,
      taxes: taxesArray,
      subtotal: Math.round(subtotal * 100) / 100,
      taxTotal: Math.round(taxTotal * 100) / 100,
      totalAmount: Math.round(totalAmount * 100) / 100
    }));

    setShowSavedFeedback(true);
    setTimeout(() => setShowSavedFeedback(false), 2000);
  };

  const triggerSave = () => {
    onSave(data);
    setShowSavedFeedback(true);
    setTimeout(() => setShowSavedFeedback(false), 2000);
  };

  return (
    <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden flex flex-col h-full" id="invoice-editor-container">
      {/* OCR Header Summary Row */}
      <div className="p-4 md:p-6 bg-slate-50/50 border-b border-slate-100 flex flex-wrap items-center justify-between gap-4">
        <div>
          <span className="text-xs font-mono font-bold text-blue-700 bg-blue-50 border border-blue-150 px-2 py-0.5 rounded inline-flex items-center gap-1.5 mb-2">
            <Sparkles className="w-3 h-3 text-blue-600" /> Procesada de forma inteligente
          </span>
          <h2 className="text-lg font-bold text-slate-800 leading-none">
            Factura {data.invoiceNumber || "S/N"}
          </h2>
          <p className="text-slate-500 text-xs mt-1.5 truncate max-w-sm md:max-w-md">
            Archivo: {fileName || "Documento actual"}
          </p>
        </div>

        <div className="flex items-center gap-2">
          <button
            onClick={() => downloadJson(data, `invoice-${data.invoiceNumber || "export"}.json`)}
            className="flex items-center gap-1.5 px-3 py-1.5 text-xs text-slate-705 bg-white border border-slate-200 rounded-md shadow-sm hover:bg-slate-50 transition cursor-pointer"
            title="Descargar JSON"
            id="btn-download-json"
          >
            <FileJson className="w-3.5 h-3.5 text-orange-500" />
            <span className="hidden sm:inline">JSON</span>
          </button>
          
          <button
            onClick={() => exportInvoiceToCSV(data, `invoice-${data.invoiceNumber || "export"}.csv`)}
            className="flex items-center gap-1.5 px-3 py-1.5 text-xs text-slate-705 bg-white border border-slate-200 rounded-md shadow-sm hover:bg-slate-50 transition cursor-pointer"
            title="Exportar conceptos a CSV"
            id="btn-export-csv"
          >
            <FileDown className="w-3.5 h-3.5 text-emerald-600" />
            <span className="hidden sm:inline">CSV</span>
          </button>

          <button
            onClick={triggerSave}
            className="flex items-center gap-1.5 px-4 py-2 bg-blue-600 text-white rounded-md text-sm font-semibold shadow-sm hover:bg-blue-700 transition cursor-pointer"
            id="btn-save-invoice"
          >
            {showSavedFeedback ? (
              <>
                <Check className="w-3.5 h-3.5 text-blue-200" />
                <span>¿Guardado!</span>
              </>
            ) : (
              <span>Guardar cambios</span>
            )}
          </button>
        </div>
      </div>

      {/* Accuracy Warning / Insights */}
      {data.summaryOfAccuracy && (
        <div className="px-6 py-2 bg-amber-50/50 border-b border-amber-100 flex items-start gap-2.5 text-xs text-amber-800">
          <AlertTriangle className="w-3.5 h-3.5 text-amber-600 shrink-0 mt-0.5" />
          <p className="leading-normal font-sans">
            <span className="font-semibold">Nota OCR:</span> {data.summaryOfAccuracy}
          </p>
        </div>
      )}

      {/* Tabs list */}
      <div className="flex border-b border-slate-200 text-sm overflow-x-auto bg-slate-50/20" id="editor-tabs">
        {(["general", "items", "taxes", "json"] as const).map(tab => {
          const labelMap: Record<ActiveTab, string> = {
            general: "Resumen General",
            items: "Conceptos u Hojas",
            taxes: "Desglose Impuestos",
            json: "Código RAW JSON"
          };
          const isActive = activeTab === tab;
          return (
            <button
              key={tab}
              onClick={() => setActiveTab(tab)}
              className={`px-5 py-3.5 border-b-2 font-semibold font-sans whitespace-nowrap transition-all outline-none cursor-pointer ${
                isActive 
                  ? "border-blue-600 text-blue-600 bg-white" 
                  : "border-transparent text-slate-500 hover:text-slate-800 hover:bg-slate-50"
              }`}
            >
              {labelMap[tab]}
            </button>
          );
        })}
      </div>

      {/* Editor Main Canvas */}
      <div className="p-5 md:p-6 overflow-y-auto flex-1 max-h-[500px] md:max-h-[600px] bg-white">
        
        {/* TAB 1: GENERAL INFO */}
        {activeTab === "general" && (
          <div className="space-y-6" id="tab-general-content">
            
            {/* Header dates/numbers row */}
            <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-4 pb-6 border-b border-stone-100">
              <div>
                <label className="block text-xs font-sans font-semibold text-stone-500 uppercase tracking-wider mb-1.5 flex items-center gap-1.5">
                  <span className="w-1.5 h-1.5 rounded-full bg-slate-800" /> Nro. Factura
                </label>
                <input
                  type="text"
                  value={data.invoiceNumber || ""}
                  onChange={(e) => handleFieldChange(null, "invoiceNumber", e.target.value)}
                  className="w-full text-sm rounded-lg border border-stone-200 px-3 py-2 focus:ring-1 focus:ring-slate-900 focus:outline-none"
                  placeholder="Ej. F-2026-0001"
                />
              </div>

              <div>
                <label className="block text-xs font-sans font-semibold text-stone-500 uppercase tracking-wider mb-1.5 flex items-center gap-1.5">
                  <Calendar className="w-3.5 h-3.5 text-stone-400" /> Fecha Emisión
                </label>
                <input
                  type="text"
                  value={data.invoiceDate || ""}
                  onChange={(e) => handleFieldChange(null, "invoiceDate", e.target.value)}
                  className="w-full text-sm rounded-lg border border-stone-200 px-3 py-2 focus:ring-1 focus:ring-slate-900 focus:outline-none"
                  placeholder="YYYY-MM-DD"
                />
              </div>

              <div>
                <label className="block text-xs font-sans font-semibold text-stone-500 uppercase tracking-wider mb-1.5 flex items-center gap-1.5">
                  <Calendar className="w-3.5 h-3.5 text-stone-400" /> Vencimiento
                </label>
                <input
                  type="text"
                  value={data.dueDate || ""}
                  onChange={(e) => handleFieldChange(null, "dueDate", e.target.value)}
                  className="w-full text-sm rounded-lg border border-stone-200 px-3 py-2 focus:ring-1 focus:ring-slate-900 focus:outline-none"
                  placeholder="YYYY-MM-DD"
                />
              </div>

              <div>
                <label className="block text-xs font-sans font-semibold text-stone-500 uppercase tracking-wider mb-1.5 flex items-center gap-1.5">
                  Orden de Compra / PO
                </label>
                <input
                  type="text"
                  value={data.poNumber || ""}
                  onChange={(e) => handleFieldChange(null, "poNumber", e.target.value)}
                  className="w-full text-sm rounded-lg border border-stone-200 px-3 py-2 focus:ring-1 focus:ring-slate-900 focus:outline-none"
                  placeholder="Opcional"
                />
              </div>

              <div>
                <label className="block text-xs font-sans font-semibold text-stone-500 uppercase tracking-wider mb-1.5 flex items-center gap-1.5">
                  Método / Términos de Pago
                </label>
                <input
                  type="text"
                  value={data.paymentTerms || ""}
                  onChange={(e) => handleFieldChange(null, "paymentTerms", e.target.value)}
                  className="w-full text-sm rounded-lg border border-stone-200 px-3 py-2 focus:ring-1 focus:ring-slate-900 focus:outline-none"
                  placeholder="Ej. Transferencia 30 días"
                />
              </div>

              <div>
                <label className="block text-xs font-sans font-semibold text-stone-500 uppercase tracking-wider mb-1.5 flex items-center gap-1.5">
                  Divisa Principal
                </label>
                <input
                  type="text"
                  value={data.currency || ""}
                  onChange={(e) => handleFieldChange(null, "currency", e.target.value)}
                  className="w-full text-sm rounded-lg border border-stone-200 px-3 py-2 focus:ring-1 focus:ring-slate-900 focus:outline-none font-medium"
                  placeholder="EUR, USD, GBP"
                />
              </div>
            </div>

            {/* Vendor and Customer section side-by-side */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6 pt-2">
              {/* VENDOR CARD */}
              <div className="space-y-3.5 border border-stone-150 p-4 rounded-xl bg-stone-50/30">
                <span className="text-xs font-sans font-bold text-stone-600 flex items-center gap-1.5 uppercase tracking-wider">
                  <Building2 className="w-4 h-4 text-slate-800" /> Emisor (Proveedor)
                </span>
                
                <div>
                  <label className="block text-xs text-stone-500 mb-1">Nombre Comercial / Fiscal</label>
                  <input
                    type="text"
                    value={data.vendor?.name || ""}
                    onChange={(e) => handleFieldChange("vendor", "name", e.target.value)}
                    className="w-full text-sm bg-white rounded-lg border border-stone-200 px-3 py-1.5 focus:outline-none focus:ring-1 focus:ring-slate-900"
                  />
                </div>

                <div>
                  <label className="block text-xs text-stone-500 mb-1">C.I.F. / N.I.F. / Tax ID</label>
                  <input
                    type="text"
                    value={data.vendor?.taxId || ""}
                    onChange={(e) => handleFieldChange("vendor", "taxId", e.target.value)}
                    className="w-full text-sm bg-white rounded-lg border border-stone-200 px-3 py-1.5 focus:outline-none focus:ring-1 focus:ring-slate-900 font-mono"
                  />
                </div>

                <div>
                  <label className="block text-xs text-stone-500 mb-1">Dirección Fiscal</label>
                  <textarea
                    rows={2}
                    value={data.vendor?.address || ""}
                    onChange={(e) => handleFieldChange("vendor", "address", e.target.value)}
                    className="w-full text-sm bg-white rounded-lg border border-stone-200 px-3 py-1.5 focus:outline-none focus:ring-1 focus:ring-slate-900"
                  />
                </div>

                <div className="grid grid-cols-2 gap-2">
                  <div>
                    <label className="block text-xs text-stone-500 mb-1">Teléfono</label>
                    <input
                      type="text"
                      value={data.vendor?.phone || ""}
                      onChange={(e) => handleFieldChange("vendor", "phone", e.target.value)}
                      className="w-full text-sm bg-white rounded-lg border border-stone-200 px-3 py-1.5 focus:outline-none focus:ring-1 focus:ring-slate-900"
                    />
                  </div>
                  <div>
                    <label className="block text-xs text-stone-500 mb-1">Email</label>
                    <input
                      type="text"
                      value={data.vendor?.email || ""}
                      onChange={(e) => handleFieldChange("vendor", "email", e.target.value)}
                      className="w-full text-sm bg-white rounded-lg border border-stone-200 px-3 py-1.5 focus:outline-none focus:ring-1 focus:ring-slate-900"
                    />
                  </div>
                </div>
              </div>

              {/* CUSTOMER CARD */}
              <div className="space-y-3.5 border border-stone-150 p-4 rounded-xl bg-stone-50/30">
                <span className="text-xs font-sans font-bold text-stone-600 flex items-center gap-1.5 uppercase tracking-wider">
                  <User className="w-4 h-4 text-slate-800" /> Receptor (Cliente)
                </span>

                <div>
                  <label className="block text-xs text-stone-500 mb-1">Razón Social Cliente</label>
                  <input
                    type="text"
                    value={data.customer?.name || ""}
                    onChange={(e) => handleFieldChange("customer", "name", e.target.value)}
                    className="w-full text-sm bg-white rounded-lg border border-stone-200 px-3 py-1.5 focus:outline-none focus:ring-1 focus:ring-slate-900"
                  />
                </div>

                <div>
                  <label className="block text-xs text-stone-500 mb-1">N.I.F. / C.I.F. / Tax ID</label>
                  <input
                    type="text"
                    value={data.customer?.taxId || ""}
                    onChange={(e) => handleFieldChange("customer", "taxId", e.target.value)}
                    className="w-full text-sm bg-white rounded-lg border border-stone-200 px-3 py-1.5 focus:outline-none focus:ring-1 focus:ring-slate-900 font-mono"
                  />
                </div>

                <div>
                  <label className="block text-xs text-stone-500 mb-1">Dirección Cliente</label>
                  <textarea
                    rows={2}
                    value={data.customer?.address || ""}
                    onChange={(e) => handleFieldChange("customer", "address", e.target.value)}
                    className="w-full text-sm bg-white rounded-lg border border-stone-200 px-3 py-1.5 focus:outline-none focus:ring-1 focus:ring-slate-900"
                  />
                </div>

                <div>
                  <label className="block text-xs text-stone-500 mb-1">Email de Contacto</label>
                  <input
                    type="text"
                    value={data.customer?.email || ""}
                    onChange={(e) => handleFieldChange("customer", "email", e.target.value)}
                    className="w-full text-sm bg-white rounded-lg border border-stone-200 px-3 py-1.5 focus:outline-none focus:ring-1 focus:ring-slate-900"
                  />
                </div>
              </div>
            </div>

            {/* Payment instructions */}
            <div className="pt-2 border-t border-stone-100">
              <label className="block text-xs font-sans font-semibold text-stone-500 uppercase tracking-wider mb-1.5">
                Instrucciones de Pago / Datos Bancarios / Cuenta IBAN
              </label>
              <textarea
                rows={2}
                value={data.paymentInstructions || ""}
                onChange={(e) => handleFieldChange(null, "paymentInstructions", e.target.value)}
                className="w-full text-sm rounded-lg border border-stone-200 px-3 py-2 focus:ring-1 focus:ring-slate-900 focus:outline-none"
                placeholder="Introduzca los datos sobre transferencias, IBAN o instrucciones de facturación"
              />
            </div>

            {/* Total recap grid */}
            <div className="grid grid-cols-2 md:grid-cols-4 gap-4 pt-4 border-t border-slate-200">
              <div className="p-3 bg-slate-50 border border-slate-150 rounded-lg text-center">
                <span className="text-[10px] font-bold text-slate-400 uppercase tracking-tighter block">Subtotal</span>
                <span className="text-base font-bold text-slate-850 font-mono block mt-1">{formatCurrency(data.subtotal, data.currency)}</span>
              </div>
              <div className="p-3 bg-slate-50 border border-slate-150 rounded-lg text-center">
                <span className="text-[10px] font-bold text-slate-400 uppercase tracking-tighter block">Descuento</span>
                <input
                  type="number"
                  value={data.discount || 0}
                  onChange={(e) => handleFieldChange(null, "discount", parseFloat(e.target.value) || 0)}
                  className="w-20 text-center inline text-sm border-b border-slate-200 focus:outline-none focus:ring-1 focus:ring-blue-600 font-bold mt-1 text-red-600 block mx-auto"
                />
              </div>
              <div className="p-3 bg-slate-50 border border-slate-150 rounded-lg text-center">
                <span className="text-[10px] font-bold text-slate-400 uppercase tracking-tighter block">IVA Impuestos</span>
                <span className="text-base font-bold text-slate-850 font-mono block mt-1">{formatCurrency(data.taxTotal, data.currency)}</span>
              </div>
              <div className="p-3 bg-blue-50 border border-blue-200 rounded-lg text-center">
                <span className="text-[10px] font-bold text-blue-900 uppercase tracking-tighter block">TOTAL</span>
                <span className="text-lg font-bold text-blue-900 font-mono block mt-1">{formatCurrency(data.totalAmount, data.currency)}</span>
              </div>
            </div>

          </div>
        )}

        {/* TAB 2: LINE ITEMS TABLE WITH ADD / REMOVE / RECALCULATE */}
        {activeTab === "items" && (
          <div className="space-y-4" id="tab-items-content">
            <div className="flex items-center justify-between pb-3 border-b border-slate-100">
              <span className="text-xs text-slate-500">
                Puedes editar descripciones, cantidades y precios. Haz clic en "Recalcular Totales" para aplicar los cambios fiscales correspondientes.
              </span>
              <div className="flex gap-2">
                <button
                  type="button"
                  onClick={addLineItem}
                  className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-slate-700 bg-white border border-slate-200 rounded-md hover:bg-slate-50 transition shadow-sm cursor-pointer"
                >
                  <Plus className="w-3.5 h-3.5" /> Añadir Línea
                </button>
                <button
                  type="button"
                  onClick={recalculateTotals}
                  className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-white bg-blue-600 rounded-md hover:bg-blue-700 transition shadow-sm cursor-pointer"
                >
                  Recalcular Totales
                </button>
              </div>
            </div>

            <div className="overflow-x-auto -mx-6">
              <div className="inline-block min-w-full align-middle px-6">
                <table className="min-w-full divide-y divide-stone-200">
                  <thead className="bg-stone-50/70">
                    <tr>
                      <th scope="col" className="px-3 py-3 text-left text-xs font-semibold text-stone-500 uppercase tracking-wider w-8">#</th>
                      <th scope="col" className="px-3 py-3 text-left text-xs font-semibold text-stone-500 uppercase tracking-wider">Concepto / Descripción</th>
                      <th scope="col" className="px-3 py-3 text-right text-xs font-semibold text-stone-500 uppercase tracking-wider w-20">Cant</th>
                      <th scope="col" className="px-3 py-3 text-right text-xs font-semibold text-stone-500 uppercase tracking-wider w-28">Precio Unit. (Bruto)</th>
                      <th scope="col" className="px-3 py-3 text-right text-xs font-semibold text-stone-500 uppercase tracking-wider w-24">Tasa IVA %</th>
                      <th scope="col" className="px-3 py-3 text-right text-xs font-semibold text-stone-500 uppercase tracking-wider w-28">Total Línea</th>
                      <th scope="col" className="px-3 py-3 text-center text-xs font-semibold text-stone-500 uppercase tracking-wider w-12">Acción</th>
                    </tr>
                  </thead>
                  <tbody className="bg-white divide-y divide-stone-200">
                    {data.lineItems && data.lineItems.length > 0 ? (
                      data.lineItems.map((item, index) => (
                        <tr key={index} className="hover:bg-stone-50/20">
                          <td className="px-3 py-3 whitespace-nowrap text-sm text-stone-500 font-mono">
                            {index + 1}
                          </td>
                          <td className="px-3 py-3 text-sm">
                            <input
                              type="text"
                              value={item.description || ""}
                              onChange={(e) => handleLineItemChange(index, "description", e.target.value)}
                              className="w-full text-sm border-b border-transparent focus:border-stone-400 py-1 focus:outline-none"
                            />
                          </td>
                          <td className="px-3 py-3 text-right text-sm">
                            <input
                              type="number"
                              value={item.quantity}
                              step="any"
                              onChange={(e) => handleLineItemChange(index, "quantity", e.target.value)}
                              className="w-full text-right text-sm border-b border-transparent focus:border-stone-400 py-1 focus:outline-none"
                            />
                          </td>
                          <td className="px-3 py-3 text-right text-sm font-mono">
                            <input
                              type="number"
                              value={item.unitPrice}
                              step="any"
                              onChange={(e) => handleLineItemChange(index, "unitPrice", e.target.value)}
                              className="w-full text-right text-sm border-b border-transparent focus:border-stone-400 py-1 focus:outline-none font-mono"
                            />
                          </td>
                          <td className="px-3 py-3 text-right text-sm font-mono">
                            <input
                              type="number"
                              value={item.taxRate}
                              onChange={(e) => handleLineItemChange(index, "taxRate", e.target.value)}
                              className="w-14 text-right text-sm border-b border-transparent focus:border-stone-400 py-1 focus:outline-none font-mono"
                            />
                          </td>
                          <td className="px-3 py-3 text-right text-sm font-mono font-medium text-stone-900">
                            {formatCurrency(item.amount, data.currency)}
                          </td>
                          <td className="px-3 py-3 text-center whitespace-nowrap text-sm">
                            <button
                              type="button"
                              onClick={() => removeLineItem(index)}
                              className="text-stone-400 hover:text-red-600 transition p-1 cursor-pointer"
                              title="Eliminar concepto"
                            >
                              <Trash2 className="w-4 h-4" />
                            </button>
                          </td>
                        </tr>
                      ))
                    ) : (
                      <tr>
                        <td colSpan={7} className="px-4 py-8 text-center text-stone-400 text-sm">
                          No hay conceptos en la factura. Añade uno para comenzar.
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>
            
            <div className="flex justify-end pt-4">
              <div className="w-full max-w-sm space-y-1.5 p-4 bg-slate-50 rounded-xl border border-slate-200">
                <div className="flex justify-between text-xs text-slate-500">
                  <span>Monto Imponible (Suma líneas):</span>
                  <span className="font-mono">{formatCurrency(data.subtotal, data.currency)}</span>
                </div>
                <div className="flex justify-between text-xs text-slate-500">
                  <span>Suma Impuestos calculada:</span>
                  <span className="font-mono">{formatCurrency(data.taxTotal, data.currency)}</span>
                </div>
                {data.discount > 0 && (
                  <div className="flex justify-between text-xs text-red-650">
                    <span>Descuento descontado:</span>
                    <span className="font-mono">-{formatCurrency(data.discount, data.currency)}</span>
                  </div>
                )}
                <div className="flex justify-between items-center py-2 px-3 bg-blue-50 rounded-lg border border-blue-100 text-sm font-bold text-blue-900 mt-2">
                  <span>TOTAL CALCULADO:</span>
                  <span className="font-mono">{formatCurrency(data.totalAmount, data.currency)}</span>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* TAB 3: TAX BREAKDOWN */}
        {activeTab === "taxes" && (
          <div className="space-y-4" id="tab-taxes-content">
            <span className="text-xs text-stone-500 block mb-2">
              Desglose e identificación fiscal de los tipos impositivos detectados en el documento.
            </span>

            <div className="border border-stone-200 rounded-xl overflow-hidden">
              <table className="min-w-full divide-y divide-stone-200">
                <thead className="bg-stone-50">
                  <tr>
                    <th scope="col" className="px-4 py-3 text-left text-xs font-semibold text-stone-500 uppercase tracking-wider">Tipo Impuesto (%)</th>
                    <th scope="col" className="px-4 py-3 text-right text-xs font-semibold text-stone-500 uppercase tracking-wider">Base Imponible</th>
                    <th scope="col" className="px-4 py-3 text-right text-xs font-semibold text-stone-500 uppercase tracking-wider">Importe Cuota Impuesto</th>
                  </tr>
                </thead>
                <tbody className="bg-white divide-y divide-stone-200">
                  {data.taxes && data.taxes.length > 0 ? (
                    data.taxes.map((tax, i) => (
                      <tr key={i} className="font-mono text-sm">
                        <td className="px-4 py-3.5 text-stone-700 font-medium">
                          {tax.taxRate}%
                        </td>
                        <td className="px-4 py-3.5 text-right text-stone-600">
                          {formatCurrency(tax.taxableAmount, data.currency)}
                        </td>
                        <td className="px-4 py-3.5 text-right font-semibold text-stone-950">
                          {formatCurrency(tax.taxAmount, data.currency)}
                        </td>
                      </tr>
                    ))
                  ) : (
                    <tr>
                      <td colSpan={3} className="px-4 py-6 text-center text-stone-400 text-sm">
                        No se ha detectado ningún tipo impositivo diferenciado o exento.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* TAB 4: RAW JSON PREVIEW */}
        {activeTab === "json" && (
          <div className="space-y-4" id="tab-json-content">
            <div className="flex items-center justify-between">
              <span className="text-xs text-stone-500">
                Copie o exporte esta estructura JSON estandarizada directamente hacia su software ERP o contable.
              </span>
              <button
                onClick={() => {
                  navigator.clipboard.writeText(JSON.stringify(data, null, 2));
                  setShowSavedFeedback(true);
                  setTimeout(() => setShowSavedFeedback(false), 2000);
                }}
                className="px-3 py-1.5 text-xs bg-stone-100 hover:bg-stone-200 text-stone-850 rounded-lg transition"
                id="btn-copy-json"
              >
                Copiar JSON
              </button>
            </div>
            
            <div className="bg-stone-950 text-stone-200 p-4 rounded-xl overflow-x-auto text-[13px] font-mono leading-relaxed shadow-inner max-h-[380px]">
              <pre id="raw-json-code">
                {JSON.stringify(data, null, 2)}
              </pre>
            </div>
          </div>
        )}

      </div>
    </div>
  );
}
