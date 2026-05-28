import { useState, useRef, DragEvent, ChangeEvent } from "react";
import { UploadCloud, FileText, ImageIcon, AlertCircle } from "lucide-react";

interface UploadZoneProps {
  onFileLoaded: (base64: string, mimeType: string, fileName: string) => void;
  isProcessing: boolean;
}

export function UploadZone({ onFileLoaded, isProcessing }: UploadZoneProps) {
  const [isDragActive, setIsDragActive] = useState(false);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFile = (file: File) => {
    setErrorMessage(null);

    // Validate type
    const validTypes = ["application/pdf", "image/png", "image/jpeg", "image/jpg", "image/webp"];
    if (!validTypes.includes(file.type)) {
      setErrorMessage("Formato no soportado. Sube una imagen (PNG, JPG, WeBP) o un documento PDF.");
      return;
    }

    // Validate size (12MB max for high-res invoices)
    if (file.size > 12 * 1024 * 1024) {
      setErrorMessage("El archivo es demasiado grande. El límite recomendado es de 12 MB.");
      return;
    }

    const reader = new FileReader();
    reader.onload = () => {
      const result = reader.result as string;
      onFileLoaded(result, file.type, file.name);
    };
    reader.onerror = () => {
      setErrorMessage("Error al leer el archivo. Inténtalo de nuevo.");
    };
    reader.readAsDataURL(file);
  };

  const onDragOver = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    e.stopPropagation();
    if (!isProcessing) {
      setIsDragActive(true);
    }
  };

  const onDragLeave = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragActive(false);
  };

  const onDrop = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragActive(false);

    if (isProcessing) return;

    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      handleFile(e.dataTransfer.files[0]);
    }
  };

  const onFileInputChange = (e: ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      handleFile(e.target.files[0]);
    }
  };

  const triggerInputClick = () => {
    if (!isProcessing && fileInputRef.current) {
      fileInputRef.current.click();
    }
  };

  return (
    <div className="w-full">
      <div
        id="upload-dropzone"
        onDragOver={onDragOver}
        onDragLeave={onDragLeave}
        onDrop={onDrop}
        onClick={triggerInputClick}
        className={`relative w-full border-2 border-dashed rounded-2xl p-8 md:p-12 text-center transition-all cursor-pointer flex flex-col items-center justify-center min-h-[220px] ${
          isProcessing ? "opacity-60 cursor-not-allowed bg-slate-50 border-slate-200" : ""
        } ${
          isDragActive
            ? "border-blue-500 bg-blue-50/30 scale-[1.01]"
            : "border-slate-300 hover:border-blue-500 bg-slate-50/50 hover:bg-slate-50"
        }`}
      >
        <input
          id="invoice-file-input"
          ref={fileInputRef}
          type="file"
          accept=".pdf,image/png,image/jpeg,image/jpg,image/webp"
          className="hidden"
          onChange={onFileInputChange}
          disabled={isProcessing}
        />

        <div className="p-4 bg-white rounded-2xl shadow-sm border border-slate-100 flex items-center justify-center mb-4 transition-transform group-hover:scale-110">
          <UploadCloud className={`w-8 h-8 ${isDragActive ? "text-blue-600" : "text-slate-500"}`} />
        </div>

        <h3 className="font-sans font-medium text-slate-800 text-lg mb-1 leading-snug">
          Arrastra tu factura o haz clic para subirla
        </h3>
        <p className="text-slate-500 text-sm max-w-sm mb-4">
          Soporta documentos contables en formato <span className="font-medium text-blue-600">PDF</span>, imagenes <span className="font-medium text-amber-600">PNG, JPG, WebP</span> (máx 12MB).
        </p>

        <div className="flex items-center gap-3 text-xs text-slate-400 mt-2">
          <span className="flex items-center gap-1 bg-slate-100 px-2 py-1 rounded">
            <FileText className="w-3.5 h-3.5" /> PDF
          </span>
          <span className="flex items-center gap-1 bg-slate-100 px-2 py-1 rounded">
            <ImageIcon className="w-3.5 h-3.5" /> Imágenes
          </span>
        </div>
      </div>

      {errorMessage && (
        <div className="mt-4 flex items-start gap-2 bg-red-50 border border-red-200 text-red-800 rounded-xl p-3.5 text-sm leading-relaxed" id="upload-error">
          <AlertCircle className="w-4 h-4 shrink-0 mt-0.5 text-red-600" />
          <span className="font-sans font-medium">{errorMessage}</span>
        </div>
      )}
    </div>
  );
}
