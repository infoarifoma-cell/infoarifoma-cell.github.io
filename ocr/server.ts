import express from "express";
import path from "path";
import { createServer as createViteServer } from "vite";
import { GoogleGenAI, Type } from "@google/genai";

const app = express();
const PORT = 3000;

// Increase raw payload limit for processing images and PDFs
app.use(express.json({ limit: "25mb" }));
app.use(express.urlencoded({ limit: "25mb", extended: true }));

// Definition of strict JSON response schema for invoices
const invoiceSchema = {
  type: Type.OBJECT,
  properties: {
    invoiceNumber: { type: Type.STRING, description: "Número de factura o serie, ej. F-2026-003, o vacío si no existe." },
    invoiceDate: { type: Type.STRING, description: "Fecha de emisión de la factura en formato YYYY-MM-DD o el formato original." },
    dueDate: { type: Type.STRING, description: "Fecha de vencimiento en formato YYYY-MM-DD o vacío si no se indica." },
    poNumber: { type: Type.STRING, description: "Número de pedido/orden de compra asociado si está visible." },
    paymentTerms: { type: Type.STRING, description: "Condiciones de pago, ej. Transferencia, Contado, Net 30, Tarjeta, etc." },
    vendor: {
      type: Type.OBJECT,
      properties: {
        name: { type: Type.STRING, description: "Nombre legal, comercial o razón social del emisor (proveedor)." },
        taxId: { type: Type.STRING, description: "Identificación fiscal: CIF, NIF, VAT ID, RFC, RUT, etc." },
        address: { type: Type.STRING, description: "Dirección física o fiscal completa del emisor." },
        phone: { type: Type.STRING, description: "Teléfono de contacto del emisor o vacío." },
        email: { type: Type.STRING, description: "Email del emisor o vacío." },
        website: { type: Type.STRING, description: "Sitio web del emisor." }
      },
      required: ["name"]
    },
    customer: {
      type: Type.OBJECT,
      properties: {
        name: { type: Type.STRING, description: "Nombre o razón social del cliente (receptor)." },
        taxId: { type: Type.STRING, description: "Identificación fiscal del cliente: CIF, NIF, VAT ID, RFC, etc." },
        address: { type: Type.STRING, description: "Dirección fiscal o del cliente." },
        email: { type: Type.STRING, description: "Email de contacto del cliente." }
      },
      required: ["name"]
    },
    lineItems: {
      type: Type.ARRAY,
      description: "Detalle de los conceptos facturados.",
      items: {
        type: Type.OBJECT,
        properties: {
          description: { type: Type.STRING, description: "Detalle o concepto del artículo o servicio." },
          quantity: { type: Type.NUMBER, description: "Cantidad o unidades de este elemento." },
          unitPrice: { type: Type.NUMBER, description: "Precio o coste unitario sin impuestos." },
          taxRate: { type: Type.NUMBER, description: "Porcentaje de IVA o tasa aplicada, ej. 21 para 21% de IVA." },
          amount: { type: Type.NUMBER, description: "Importe total neto de la línea (generalmente cantidad * precio)." }
        },
        required: ["description", "amount"]
      }
    },
    taxes: {
      type: Type.ARRAY,
      description: "Resumen de impuestos aplicados separados por tipo impositivo.",
      items: {
        type: Type.OBJECT,
        properties: {
          taxRate: { type: Type.NUMBER, description: "Porcentaje de impuesto como entero o decimal, ej. 21." },
          taxableAmount: { type: Type.NUMBER, description: "Base imponible correspondiente a este impuesto." },
          taxAmount: { type: Type.NUMBER, description: "Importe total del impuesto calulado para esta base." }
        },
        required: ["taxRate", "taxAmount"]
      }
    },
    subtotal: { type: Type.NUMBER, description: "Suma de importes netos de las líneas antes de impuestos." },
    discount: { type: Type.NUMBER, description: "Descuento total aplicado en la factura si lo hay." },
    taxTotal: { type: Type.NUMBER, description: "Suma total de los impuestos aplicados." },
    totalAmount: { type: Type.NUMBER, description: "Importe total a pagar (subtotal + taxTotal - descuento)." },
    currency: { type: Type.STRING, description: "Divisa de la factura, ej. EUR, USD, GBP, o símbolo €" },
    paymentInstructions: { type: Type.STRING, description: "Datos de transferencia, IBAN, cuenta bancaria, instrucciones de pago o notas." },
    summaryOfAccuracy: { type: Type.STRING, description: "Un breve comentario en español del OCR sobre la calidad de la lectura o advertencias de datos faltantes." }
  },
  required: [
    "invoiceNumber",
    "invoiceDate",
    "vendor",
    "customer",
    "lineItems",
    "subtotal",
    "taxTotal",
    "totalAmount",
    "currency"
  ]
};

// Health status API route
app.get("/api/health", (req, res) => {
  res.json({ status: "healthy", timestamp: new Date().toISOString() });
});

// Extraction OCR API Route
app.post("/api/ocr", async (req, res) => {
  try {
    const { mimeType, fileData } = req.body;

    if (!mimeType || !fileData) {
      return res.status(400).json({ 
        error: "Falta el tipo 'mimeType' o la información del archivo 'fileData' en formato Base64." 
      });
    }

    const apiKey = process.env.GEMINI_API_KEY;
    if (!apiKey) {
      return res.status(500).json({ 
        error: "No se ha configurado la clave API de Gemini. Configúrala en Settings > Secrets en AI Studio de forma segura." 
      });
    }

    // Lazy initialization of the official modern @google/genai SDK
    const ai = new GoogleGenAI({
      apiKey,
      httpOptions: {
        headers: {
          "User-Agent": "aistudio-build",
        },
      },
    });

    // Strip out base64 prefixes to send raw base64 contents to the API
    let base64Clean = fileData;
    if (base64Clean.includes(";base64,")) {
      base64Clean = base64Clean.split(";base64,").pop() || "";
    }

    const imagePart = {
      inlineData: {
        mimeType,
        data: base64Clean,
      },
    };

    const promptPart = {
      text: "Realiza una extracción precisa del texto y metadatos de esta factura o recibo. Rellena todos los campos indicados en el 'responseSchema'. Si algún campo no aparece explícitamente y no puede calcularse lógicamente, establécelo en una cadena vacía o 0. Presta especial atención al total general, la fecha de emisión, el número de factura y el desglose de impuestos.",
    };

    // Use gemini-3.5-flash for accurate and fast OCR extraction
    const response = await ai.models.generateContent({
      model: "gemini-3.5-flash",
      contents: { parts: [imagePart, promptPart] },
      config: {
        systemInstruction: "Eres un analista contable de facturas OCR de nivel experto. Transcribes con exactitud matemática extrema. Respetas símbolos de divisas, códigos de impuestos y nombres fiscales. Devuelves estrictamente JSON estructurado.",
        temperature: 0.1,
        responseMimeType: "application/json",
        responseSchema: invoiceSchema,
      },
    });

    const responseText = response.text;
    if (!responseText) {
      throw new Error("El servicio de análisis inteligente Gemini devolvió una respuesta vacía.");
    }

    const parsedData = JSON.parse(responseText);
    return res.json({ 
      success: true, 
      data: parsedData 
    });

  } catch (error: any) {
    console.error("Error en OCR:", error);
    return res.status(500).json({ 
      error: error.message || "Ocurrió un error inesperado al procesar la factura con Gemini OCR." 
    });
  }
});

async function start() {
  // Vite Integration
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    // Serve static files in production
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*", (req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`[OCR Server] Escuchando en el puerto ${PORT} (http://localhost:${PORT})`);
  });
}

start();
