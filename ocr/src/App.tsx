import React, { useState, useEffect } from "react";
import { motion, AnimatePresence } from "motion/react";
import { 
  FileText, Search, Landmark, History, Trash2, ChevronRight, 
  Sparkles, Layers, Info, CheckCircle2, AlertCircle, Play 
} from "lucide-react";
import { InvoiceData, OCRHistoryEntry } from "./types";
import { UploadZone } from "./components/UploadZone";
import { InvoiceEditor } from "./components/InvoiceEditor";
import { InvoiceAnalytics } from "./components/InvoiceAnalytics";
import { formatCurrency, formatDate } from "./utils";

export default function App() {
  const [history, setHistory] = useState<OCRHistoryEntry[]>([]);
  const [activeInvoice, setActiveInvoice] = useState<InvoiceData | null>(null);
  const [activeFileName, setActiveFileName] = useState<string>("");
  const [activeFileId, setActiveFileId] = useState<string>("");

  const [isProcessing, setIsProcessing] = useState(false);
  const [processingStep, setProcessingStep] = useState("");
  const [apiError, setApiError] = useState<string | null>(null);
  const [searchQuery, setSearchQuery] = useState("");
  const [showAnalytics, setShowAnalytics] = useState(false);

  // Load history from localStorage
  useEffect(() => {
    try {
      const stored = localStorage.getItem("ocr_invoice_history");
      if (stored) {
        setHistory(JSON.parse(stored));
      }
    } catch (e) {
      console.error("No se pudo cargar el historial de localStorage:", e);
    }
  }, []);

  // Save history helper
  const saveHistoryToStorage = (updatedHistory: OCRHistoryEntry[]) => {
    setHistory(updatedHistory);
    try {
      localStorage.setItem("ocr_invoice_history", JSON.stringify(updatedHistory));
    } catch (e) {
      console.error("No se pudo guardar el historial en localStorage:", e);
    }
  };

  // Automated step messenger for high fidelity UX
  useEffect(() => {
    if (!isProcessing) return;

    const steps = [
      "📂 Analizando la estructura inicial del archivo...",
      "🚀 Procesando archivo binario en base64...",
      "🧠 Analizando el documento con inteligencia artificial Gemini 3.5...",
      "✍️ Transcribiendo conceptos, importes y CIF fiscales...",
      "📊 Estructurando desglose impositivo e IVA general...",
      "✨ Consolidando todos los datos en formato contable..."
    ];

    let currentIdx = 0;
    setProcessingStep(steps[0]);

    const interval = setInterval(() => {
      currentIdx = (currentIdx + 1) % steps.length;
      setProcessingStep(steps[currentIdx]);
    }, 2200);

    return () => clearInterval(interval);
  }, [isProcessing]);

  // Handle file loaded base64 from UploadZone
  const handleFileLoaded = async (base64: string, mimeType: string, fileName: string) => {
    setIsProcessing(true);
    setApiError(null);
    setActiveInvoice(null);

    try {
      const response = await fetch("/api/ocr", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ mimeType, fileData: base64 })
      });

      const result = await response.json();

      if (!response.ok || !result.success) {
        throw new Error(result.error || "Ocurrió un error en el servidor OCR de Gemini.");
      }

      const invoiceData: InvoiceData = result.data;
      const newId = `invoice_${Date.now()}`;

      // Save to history with thumbnail/preview flag
      const newEntry: OCRHistoryEntry = {
        id: newId,
        fileName,
        fileType: mimeType,
        uploadedAt: new Date().toISOString(),
        invoiceData,
        fileData: base64 // cache for preview/inspect
      };

      const updatedHistory = [newEntry, ...history];
      saveHistoryToStorage(updatedHistory);

      // Focus on loaded invoice
      setActiveInvoice(invoiceData);
      setActiveFileName(fileName);
      setActiveFileId(newId);
      setShowAnalytics(false);

    } catch (err: any) {
      console.error("Error al procesar el OCR:", err);
      setApiError(err.message || "No se pudo extraer información contable del documento.");
    } finally {
      setIsProcessing(false);
    }
  };

  // Mock sample injector for user to inspect features easily without files
  const handleLoadSampleInvoice = () => {
    const sampleInvoice: InvoiceData = {
      invoiceNumber: "VLD-2026-4890",
      invoiceDate: "2026-05-15",
      dueDate: "2026-06-15",
      poNumber: "PO-MDR-9921",
      paymentTerms: "Transferencia Bancaria a 30 días",
      vendor: {
        name: "Voldis Distribución Madrid S.L.",
        taxId: "ESB86493012",
        address: "Av. de la Industria 45, Polígono Industrial Coslada, 28823 Madrid, España",
        phone: "+34 91 621 4455",
        email: "facturacion@voldis.es",
        website: "www.voldis.es"
      },
      customer: {
        name: "Restaurante El Laurel",
        taxId: "ESA28049103",
        address: "Calle Mayor 12, Planta Baja, 28013 Madrid, España",
        email: "compras@restauranteellaurel.com"
      },
      lineItems: [
        { description: "Cerveza Especial Premium Barril 50 Litros", quantity: 2, unitPrice: 120.00, taxRate: 21, amount: 240.00 },
        { description: "Refresco de Cola Caja 24 Botellas de Vidrio", quantity: 4, unitPrice: 18.50, taxRate: 21, amount: 74.00 },
        { description: "Agua Mineral Natural 1.5L Caja de 12 Unidades", quantity: 5, unitPrice: 6.20, taxRate: 10, amount: 31.00 }
      ],
      taxes: [
        { taxRate: 21, taxableAmount: 314.00, taxAmount: 65.94 },
        { taxRate: 10, taxableAmount: 31.00, taxAmount: 3.10 }
      ],
      subtotal: 345.00,
      discount: 15.00,
      taxTotal: 69.04,
      totalAmount: 399.04,
      currency: "EUR",
      paymentInstructions: "CaixaBank IBAN ES30 2100 0412 8901 2345 6789 - Concepto: Voldis-4890",
      summaryOfAccuracy: "Datos extraídos de plantilla estándar con un 100% de confianza. Se identificó desglose correspondiente a tipos reducidos (10%) y generales (21%) de IVA."
    };

    const newId = `invoice_sample_${Date.now()}`;
    const newEntry: OCRHistoryEntry = {
      id: newId,
      fileName: "factura_proveedor_voldis.pdf",
      fileType: "application/pdf",
      uploadedAt: new Date().toISOString(),
      invoiceData: sampleInvoice
    };

    const updatedHistory = [newEntry, ...history];
    saveHistoryToStorage(updatedHistory);

    setActiveInvoice(sampleInvoice);
    setActiveFileName("factura_proveedor_voldis.pdf");
    setActiveFileId(newId);
    setShowAnalytics(false);
  };

  // Save changes made in InvoiceEditor back to list and Storage
  const handleSaveInvoice = (updatedData: InvoiceData) => {
    setActiveInvoice(updatedData);
    const updatedHistory = history.map(item => {
      if (item.id === activeFileId) {
        return { ...item, invoiceData: updatedData };
      }
      return item;
    });
    saveHistoryToStorage(updatedHistory);
  };

  // Delete invoice from history
  const handleDeleteItem = (id: string, e: React.MouseEvent) => {
    e.stopPropagation();
    const updated = history.filter(item => item.id !== id);
    saveHistoryToStorage(updated);

    if (activeFileId === id) {
      setActiveInvoice(null);
      setActiveFileName("");
      setActiveFileId("");
    }
  };

  // Focus and inspect a past invoice from history list
  const handleSelectHistoryEntry = (entry: OCRHistoryEntry) => {
    setActiveInvoice(entry.invoiceData);
    setActiveFileName(entry.fileName);
    setActiveFileId(entry.id);
    setShowAnalytics(false);
    setApiError(null);
  };

  // Filter history entries by search query
  const filteredHistory = history.filter(item => {
    const q = searchQuery.toLowerCase();
    const vendorName = item.invoiceData.vendor?.name?.toLowerCase() || "";
    const invoiceNum = item.invoiceData.invoiceNumber?.toLowerCase() || "";
    const totalAmount = item.invoiceData.totalAmount?.toString() || "";
    const date = item.invoiceData.invoiceDate?.toLowerCase() || "";

    return vendorName.includes(q) || invoiceNum.includes(q) || totalAmount.includes(q) || date.includes(q);
  });

  return (
    <div className="min-h-screen bg-slate-50 text-slate-800 pb-12 flex flex-col font-sans" id="app-root-container">
      
      {/* Sleek Minimal Top Navigation Bar */}
      <header className="sticky top-0 bg-white border-b border-slate-200 z-20 h-16 shrink-0">
        <div className="max-w-7xl mx-auto px-4 md:px-6 h-full flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 bg-blue-600 text-white rounded-lg flex items-center justify-center font-bold text-xl shadow-sm">
              Σ
            </div>
            <div>
              <h1 className="text-lg font-bold text-slate-800 leading-none">Lector OCR Pro</h1>
              <p className="text-xs text-slate-500 mt-1">Procesamiento de Facturas de Alta Precisión</p>
            </div>
          </div>

          <div className="flex items-center gap-3">
            <div className="hidden lg:flex items-center gap-2 px-3 py-1.5 bg-green-50 text-green-700 rounded-full text-xs font-semibold border border-green-200">
              <span className="w-2 h-2 bg-green-500 rounded-full animate-pulse"></span>
              Motor OCR V4.2 Activo
            </div>
            
            <button
              onClick={() => {
                setActiveInvoice(null);
                setActiveFileId("");
                setShowAnalytics(!showAnalytics);
              }}
              className={`text-xs px-3.5 py-1.5 rounded-lg font-medium transition cursor-pointer ${
                showAnalytics && !activeInvoice
                  ? "bg-blue-50 text-blue-700 border border-blue-200"
                  : "bg-stone-50 hover:bg-stone-100 text-stone-700 border border-stone-200"
              }`}
              id="btn-toggle-analytics"
            >
              📊 Estadísticas
            </button>
            <button
              onClick={handleLoadSampleInvoice}
              className="px-4 py-2 bg-blue-600 text-white rounded-md text-sm font-semibold shadow-sm hover:bg-blue-700 transition duration-150 inline-flex items-center gap-1.5 cursor-pointer"
              id="btn-sample-invoice"
            >
              <Sparkles className="w-3.5 h-3.5 text-blue-100 animate-pulse" />
              <span>Instanciar Demo</span>
            </button>
          </div>
        </div>
      </header>

      {/* Main content body grid layout */}
      <main className="max-w-7xl mx-auto w-full px-4 md:px-6 mt-6 md:mt-8 flex-1 grid grid-cols-1 lg:grid-cols-12 gap-6 lg:items-start">
        
        {/* COLUMN 1: CONTROLS & HISTORY (SPAN 4) */}
        <div className="lg:col-span-4 space-y-6">
          
          {/* Main upload manager */}
          <div className="bg-white p-5 border border-slate-200 rounded-2xl shadow-sm space-y-5" id="control-panel">
            <div className="flex items-center justify-between">
              <span className="text-slate-400 text-[10px] font-bold uppercase tracking-wider block">Procesar Documento</span>
              <span className="w-2.5 h-2.5 rounded-full bg-blue-600 animate-pulse" title="Sistema listo" />
            </div>

            <UploadZone onFileLoaded={handleFileLoaded} isProcessing={isProcessing} />

            {/* In-app instructional note */}
            <div className="flex gap-2.5 py-3 px-3.5 bg-slate-50 rounded-xl text-xs text-slate-600 border border-slate-100">
              <Info className="w-4 h-4 text-blue-600 shrink-0 mt-0.5" />
              <p>
                Sube una factura y Gemini extraerá de forma automática los campos, desglose de IVA/IRPF y tablas de conceptos.
              </p>
            </div>
          </div>

          {/* Past Uploads History List */}
          <div className="bg-white p-5 border border-slate-200 rounded-2xl shadow-sm flex flex-col min-h-[280px]" id="history-panel">
            <span className="text-slate-400 text-[10px] font-bold uppercase tracking-wider block mb-3">Historial Reciente</span>
            
            {/* Search filter input */}
            <div className="relative mb-4">
              <Search className="absolute left-3 top-2.5 w-4 h-4 text-slate-400" />
              <input
                type="text"
                placeholder="Buscar por emisor, CIF, nro..."
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                className="w-full pl-9 pr-3 py-1.5 text-sm rounded-lg border border-slate-200 focus:outline-none focus:ring-1 focus:ring-blue-600 bg-slate-50/50 focus:bg-white transition-all"
                id="search-invoice-history"
              />
            </div>

            {/* List entries scroll container */}
            <div className="space-y-2.5 flex-1 max-h-[300px] overflow-y-auto pr-1" id="history-items-list">
              {filteredHistory.length > 0 ? (
                filteredHistory.map((item) => {
                  const isActive = activeFileId === item.id;
                  const vName = item.invoiceData.vendor?.name || "Proveedor no identificado";
                  const nFactura = item.invoiceData.invoiceNumber || "S/N";
                  const date = item.invoiceData.invoiceDate ? formatDate(item.invoiceData.invoiceDate) : "Sin fecha";
                  const formattedTotal = formatCurrency(item.invoiceData.totalAmount, item.invoiceData.currency);

                  return (
                    <div
                      key={item.id}
                      onClick={() => handleSelectHistoryEntry(item)}
                      className={`group p-3 rounded-xl border transition-all cursor-pointer flex items-center justify-between gap-3 ${
                        isActive
                          ? "bg-blue-50 border-blue-200 shadow-sm"
                          : "bg-white hover:bg-slate-50 border-slate-200 hover:border-slate-300"
                      }`}
                    >
                      <div className="flex items-center gap-2.5 truncate">
                        <div className={`w-8 h-8 rounded-lg flex items-center justify-center shrink-0 ${
                          isActive ? "bg-blue-105 text-blue-700" : "bg-slate-100 text-slate-500"
                        }`}>
                          <FileText className="w-4 h-4" />
                        </div>
                        <div className="truncate text-left leading-tight">
                          <p className={`text-xs font-bold font-sans truncate ${isActive ? "text-slate-800 font-semibold" : "text-slate-700"}`}>
                            {vName}
                          </p>
                          <p className={`text-[10px] font-mono mt-0.5 ${isActive ? "text-slate-500" : "text-slate-400"}`}>
                            {nFactura} • {date}
                          </p>
                        </div>
                      </div>

                      <div className="flex items-center gap-2 block shrink-0">
                        <span className={`text-xs font-bold font-mono ${isActive ? "text-blue-700" : "text-slate-800"}`}>
                          {formattedTotal}
                        </span>
                        
                        <button
                          onClick={(e) => handleDeleteItem(item.id, e)}
                          className={`p-1 rounded transition opacity-0 group-hover:opacity-100 ${
                            isActive ? "hover:bg-blue-150 text-slate-500 hover:text-red-600" : "hover:bg-slate-100 text-slate-450 hover:text-red-600"
                          }`}
                          title="Eliminar del historial"
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                        </button>
                      </div>
                    </div>
                  );
                })
              ) : (
                <div className="flex flex-col items-center justify-center py-10 px-4 text-center text-slate-400 flex-1">
                  <History className="w-8 h-8 text-slate-300 mb-2" />
                  <p className="text-xs font-medium">Historial vacío</p>
                  <p className="text-[10px] text-slate-400 mt-0.5">Sube facturas o instancia una demo.</p>
                </div>
              )}
            </div>
          </div>

        </div>

        {/* COLUMN 2: PRIMARY INTERACTIVE CANVAS (SPAN 8) */}
        <div className="lg:col-span-8">
          
          <AnimatePresence mode="wait">
            
            {/* 1. OCCUPIED PROCESSING STATE SCREEN */}
            {isProcessing && (
              <motion.div
                key="processing"
                initial={{ opacity: 0, y: 15 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -15 }}
                className="bg-white rounded-2xl border border-slate-200 p-8 md:p-12 text-center shadow-sm flex flex-col items-center justify-center min-h-[480px]"
                id="processing-stage"
              >
                <div className="relative mb-6">
                  {/* Outer breathing ring */}
                  <div className="w-16 h-16 rounded-2xl border-4 border-slate-200 border-t-blue-600 animate-spin flex items-center justify-center" />
                  {/* Floating particles background mockup */}
                  <Sparkles className="w-5 h-5 text-blue-500 absolute -top-1.5 -right-1.5 animate-bounce" />
                </div>

                <h3 className="font-sans font-semibold text-lg text-slate-850 mb-1 leading-snug">
                  Leyendo de forma inteligente tu factura
                </h3>
                
                <p className="text-slate-500 text-sm max-w-sm mb-6 font-medium">
                  La IA de Gemini está realizando la transcripción ultra-precisa y calculando los desgloses fiscales...
                </p>

                {/* Simulated/Sequential status loop */}
                <div className="bg-slate-50 border border-slate-200 py-3.5 px-6 rounded-xl w-full max-w-md font-mono text-xs text-slate-600 flex items-center gap-2 justify-center shadow-inner">
                  <span className="w-2 h-2 rounded-full bg-blue-500 animate-pulse" />
                  <span id="processing-step-text" className="truncate select-none">{processingStep}</span>
                </div>
              </motion.div>
            )}

            {/* 2. ERROR DISPLAY CONTAINER */}
            {!isProcessing && apiError && (
              <motion.div
                key="error"
                initial={{ opacity: 0 }}
                animate={{ opacity: 1 }}
                exit={{ opacity: 0 }}
                className="bg-white rounded-2xl border border-red-200 p-8 text-center shadow-sm flex flex-col items-center justify-center min-h-[400px]"
                id="error-stage"
              >
                <div className="w-12 h-12 bg-red-50 rounded-full flex items-center justify-center text-red-600 mb-4">
                  <AlertCircle className="w-6 h-6" />
                </div>
                <h3 className="font-sans font-bold text-lg text-slate-800 mb-2">
                  No se pudo leer la factura
                </h3>
                <p className="text-slate-600 text-sm max-w-md mb-6 leading-relaxed">
                  {apiError}
                </p>
                <div className="flex gap-2">
                  <button
                    onClick={() => setApiError(null)}
                    className="px-4 py-2 text-xs bg-slate-900 text-white font-medium rounded-lg hover:bg-slate-800 transition shadow-sm cursor-pointer"
                  >
                    Intentar de nuevo
                  </button>
                  <button
                    onClick={handleLoadSampleInvoice}
                    className="px-4 py-2 text-xs bg-stone-100 hover:bg-stone-200 text-stone-850 font-medium rounded-lg transition cursor-pointer"
                  >
                    Usar Factura de Ejemplo
                  </button>
                </div>
              </motion.div>
            )}

            {/* 3. ACTIVE INVOICE EDITOR CANVAS */}
            {!isProcessing && !apiError && activeInvoice && (
              <motion.div
                key="invoice-editor"
                initial={{ opacity: 0, scale: 0.99 }}
                animate={{ opacity: 1, scale: 1 }}
                exit={{ opacity: 0, scale: 0.99 }}
                className="h-full"
              >
                <InvoiceEditor
                  initialData={activeInvoice}
                  onSave={handleSaveInvoice}
                  fileName={activeFileName}
                />
              </motion.div>
            )}

            {/* 4. ANALYTICS CHART CANVAS */}
            {!isProcessing && !apiError && !activeInvoice && showAnalytics && (
              <motion.div
                key="analytics-charts"
                initial={{ opacity: 0 }}
                animate={{ opacity: 1 }}
                className="space-y-6"
              >
                <div className="flex justify-between items-center bg-white p-4 border border-slate-200 rounded-2xl">
                  <div>
                    <h3 className="font-sans font-semibold text-slate-850">Estadísticas de Gasto</h3>
                    <p className="text-xs text-slate-500 mt-0.5">Visión agregada consolidando todo tu historial.</p>
                  </div>
                  <button
                    onClick={() => setShowAnalytics(false)}
                    className="text-xs font-bold text-slate-600 hover:text-slate-905 transition cursor-pointer"
                  >
                    Volver
                  </button>
                </div>
                <InvoiceAnalytics history={history} />
              </motion.div>
            )}

            {/* 5. GREETING/INITIAL PLACEHOLDER CANVAS */}
            {!isProcessing && !apiError && !activeInvoice && !showAnalytics && (
              <motion.div
                key="empty-greeting"
                initial={{ opacity: 0 }}
                animate={{ opacity: 1 }}
                exit={{ opacity: 0 }}
                className="bg-white rounded-2xl border border-slate-200 p-8 md:p-12 text-center shadow-sm flex flex-col items-center justify-center min-h-[480px]"
                id="empty-placeholder-root"
              >
                <div className="w-14 h-14 bg-slate-50 border border-slate-200 rounded-2xl flex items-center justify-center mb-6 text-slate-600 transition-transform hover:rotate-6">
                  <Layers className="w-6 h-6 text-blue-600" />
                </div>

                <h3 className="font-sans font-bold text-[20px] text-slate-850 mb-2 leading-tight">
                  Sube tu primer documento contable
                </h3>
                
                <p className="text-slate-500 text-sm max-w-md mb-8 leading-relaxed">
                  Para comenzar, arrastra una factura en formato PDF o una foto de un recibo al bloque izquierdo de control. También puedes evaluar el lector instante haciendo clic abajo.
                </p>

                {/* Secondary call-to-actions */}
                <div className="flex flex-col sm:flex-row gap-3">
                  <button
                    onClick={handleLoadSampleInvoice}
                    className="px-5 py-2.5 bg-blue-600 text-white rounded-md text-sm font-semibold shadow-sm hover:bg-blue-700 transition flex items-center gap-2 justify-center cursor-pointer"
                    id="btn-trigger-sample-blank"
                  >
                    <Play className="w-3.5 h-3.5 fill-current text-blue-200" /> Cargar Factura de Ejemplo
                  </button>
                  <button
                    onClick={() => setShowAnalytics(true)}
                    className="px-5 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-750 font-medium text-xs rounded-xl border border-slate-200 transition justify-center cursor-pointer"
                  >
                    Explorar Indicadores
                  </button>
                </div>

                {/* Features bento list */}
                <div className="grid grid-cols-1 sm:grid-cols-3 gap-4 border-t border-slate-100 pt-8 mt-10 w-full text-left">
                  <div className="space-y-1">
                    <span className="w-6 h-6 rounded-lg bg-blue-50 text-blue-600 flex items-center justify-center text-xs font-mono font-bold">1</span>
                    <h4 className="text-xs font-bold text-slate-800">Cero Configuración</h4>
                    <p className="text-[11px] text-slate-500 leading-normal">Listo para usar de inmediato directo con el motor del modelo Flash.</p>
                  </div>
                  <div className="space-y-1">
                    <span className="w-6 h-6 rounded-lg bg-blue-50 text-blue-600 flex items-center justify-center text-xs font-mono font-bold">2</span>
                    <h4 className="text-xs font-bold text-slate-800">Tablas Detalladas</h4>
                    <p className="text-[11px] text-slate-500 leading-normal font-sans">Extrae cada concepto individual, cantidad e impuestos aplicados.</p>
                  </div>
                  <div className="space-y-1">
                    <span className="w-6 h-6 rounded-lg bg-blue-50 text-blue-600 flex items-center justify-center text-xs font-mono font-bold">3</span>
                    <h4 className="text-xs font-bold text-slate-800">Confiabilidad Fiscal</h4>
                    <p className="text-[11px] text-slate-500 leading-normal">Lee CIF/NIF, calcula subtotales, cuotas e IVA automáticamente.</p>
                  </div>
                </div>
              </motion.div>
            )}

          </AnimatePresence>

        </div>

      </main>

      {/* Styled Professional Footer Bar integrated directly */}
      <footer className="h-12 bg-white border-t border-slate-200 px-8 flex items-center justify-between text-[11px] text-slate-400 shrink-0 mt-8 w-full">
        <div>Soporte: support@ocrpro.io | v4.2.0-stable</div>
        <div className="flex gap-4">
          <span>Cumplimiento RGPD</span>
          <span>Cifrado de Extremo a Extremo</span>
        </div>
      </footer>

    </div>
  );
}
