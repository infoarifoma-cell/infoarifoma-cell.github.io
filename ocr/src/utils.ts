/**
 * Formatea un número como moneda localizada en euros o divisa dinámica.
 */
export function formatCurrency(amount: number, currency: string = "EUR"): string {
  const normCurrency = currency.trim().toUpperCase();
  const symbolMap: Record<string, string> = {
    EUR: "€",
    USD: "$",
    GBP: "£",
    MXN: "MXN$",
    COP: "COP$",
    ARS: "ARS$",
    CLP: "CLP$",
    PEN: "PEN S/",
    "€": "€",
    "$": "$",
    "£": "£"
  };
  
  const symbol = symbolMap[normCurrency] || normCurrency || "€";
  
  // Format to 2 decimal places with localized thousands separator
  const formatted = new Intl.NumberFormat("es-ES", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  }).format(amount);

  return `${formatted} ${symbol}`;
}

/**
 * Formatea una fecha en formato legible.
 */
export function formatDate(dateStr: string): string {
  if (!dateStr) return "-";
  try {
    const date = new Date(dateStr);
    if (isNaN(date.getTime())) return dateStr; // fallback if it is a random string
    return new Intl.DateTimeFormat("es-ES", {
      year: "numeric",
      month: "long",
      day: "numeric"
    }).format(date);
  } catch {
    return dateStr;
  }
}

/**
 * Descarga datos JSON como archivo.
 */
export function downloadJson(data: any, filename: string) {
  const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

/**
 * Convierte los conceptos de factura a CSV y descarga el archivo.
 */
export function exportInvoiceToCSV(invoice: any, filename: string) {
  const headers = ["Factura Nro", "Fecha", "Proveedor", "Concepto", "Cantidad", "Precio Unitario", "Impuesto %", "Importe Línea"];
  
  const rows = invoice.lineItems.map((item: any) => [
    invoice.invoiceNumber,
    invoice.invoiceDate,
    invoice.vendor.name,
    item.description.replace(/"/g, '""'),
    item.quantity,
    item.unitPrice,
    item.taxRate,
    item.amount
  ]);
  
  const csvContent = [
    headers.join(","),
    ...rows.map((row: any[]) => row.map(val => `"${val}"`).join(","))
  ].join("\n");
  
  const blob = new Blob([new Uint8Array([0xEF, 0xBB, 0xBF]), csvContent], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
