// POST /api/ocr-hf
// Envía imagen a NuExtract3 o Qwen VL vía Hugging Face Inference API
// Autodetecta tipo de documento (factura o acta de ensayo)
// Requiere env var HF_TOKEN en Vercel

export const config = { api: { bodyParser: { sizeLimit: '10mb' } } };

const TEMPLATE_FACTURA = JSON.stringify({
  doc_type: "factura",
  proveedor: "verbatim-string",
  nfactura: "verbatim-string",
  fecha: "date",
  base_imponible: "number",
  iva_porcentaje: "number",
  total: "number",
  lineas: [
    {
      descripcion: "verbatim-string",
      cantidad: "number",
      precio_unitario: "number",
      importe: "number"
    }
  ]
});

const TEMPLATE_ENSAYO = JSON.stringify({
  doc_type: "ensayo",
  num_acta: "verbatim-string",
  num_albaran: "verbatim-string",
  fecha_toma: "date (YYYY-MM-DD)",
  fecha_acta: "date (YYYY-MM-DD)",
  fraccion: "verbatim-string (e.g. 0/4, 4/12, 12/20, 20/40, ZA25)",
  tipo_ensayo: "one of: granulometria, cont_finos, eq_arena, ind_lajas, caras_fractura",
  resultados: {
    gran_80: "number or null", gran_63: "number or null", gran_50: "number or null",
    gran_40: "number or null", gran_32: "number or null", gran_20: "number or null",
    gran_16: "number or null", gran_14: "number or null", gran_12_5: "number or null",
    gran_10: "number or null", gran_8: "number or null", gran_6_3: "number or null",
    gran_4: "number or null", gran_2: "number or null", gran_1: "number or null",
    gran_0_5: "number or null", gran_0_25: "number or null", gran_0_125: "number or null",
    gran_0_063: "number or null",
    eq_arena: "number or null",
    ind_lajas: "number or null",
    cont_finos: "number or null",
    caras_fractura: "number or null"
  }
});

const PROMPT_AUTO = `Analyze this document image and determine what type it is, then extract structured data.

STEP 1 — Detect document type:
- If it is an INVOICE / RECEIPT / FACTURA → use "factura" schema
- If it is a LABORATORY TEST REPORT / ACTA DE ENSAYO (aggregate testing, granulometry, sieve analysis) → use "ensayo" schema

STEP 2 — Extract data using the matching schema below. Return ONLY valid JSON.

=== If doc_type = "factura", use this schema: ===
${TEMPLATE_FACTURA}
- "proveedor" is the SELLER/SUPPLIER, NOT the buyer.
- "ARIDOS FONOLITICOS DE MASPALOMAS" / "ARIFOMA" is the BUYER, never the supplier.
- "lineas": array of line items with description, quantity, unit price, line total.

=== If doc_type = "ensayo", use this schema: ===
${TEMPLATE_ENSAYO}
- "num_acta": report number (e.g. "2026/258")
- "num_albaran": delivery note / sample number (e.g. "2026/101"), NOT the same as num_acta
- "fecha_toma": sample date, YYYY-MM-DD
- "fecha_acta": report date / fin de ensayos, YYYY-MM-DD
- "fraccion": aggregate size (e.g. "0/4", "4/12", "12/20", "20/40", "ZA25")
- "tipo_ensayo": "granulometria" if sieve table, "eq_arena" if sand equivalent, "ind_lajas" if flakiness, "cont_finos" if fines content, "caras_fractura" if crushed faces
- "resultados": only fill relevant fields. gran_X = % passing sieve X mm (gran_12_5 for 12.5mm, gran_6_3 for 6.3mm, gran_0_5 for 0.5mm, etc.)

If a field is not found, use null.`;

export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ ok: false, error: 'Método no permitido' });

  const token = process.env.HF_TOKEN;
  if (!token) return res.status(500).json({ ok: false, error: 'HF_TOKEN no configurado' });

  try {
    const { image } = req.body;
    if (!image) return res.status(400).json({ ok: false, error: 'Falta imagen' });
    const PROMPT = PROMPT_AUTO;

    const match = image.match(/^data:(image\/\w+);base64,(.+)$/);
    if (!match) return res.status(400).json({ ok: false, error: 'Formato de imagen inválido' });

    const mimeType = match[1];
    const base64Data = match[2];
    const dataUrl = `data:${mimeType};base64,${base64Data}`;

    // Provider list — intentar en orden
    const providers = [
      // NuExtract3 via featherless (soporta VLMs)
      {
        url: 'https://router.huggingface.co/featherless-ai/v1/chat/completions',
        model: 'numind/NuExtract3',
        extraBody: {
          chat_template_kwargs: {
            enable_thinking: false
          }
        }
      },
      // Qwen3-VL-4B como fallback
      {
        url: 'https://router.huggingface.co/featherless-ai/v1/chat/completions',
        model: 'Qwen/Qwen3-VL-4B-Instruct',
        extraBody: {}
      },
      // Qwen2.5-VL-7B via novita
      {
        url: 'https://router.huggingface.co/novita/v3/openai/chat/completions',
        model: 'Qwen/Qwen2.5-VL-7B-Instruct',
        extraBody: {}
      },
    ];

    let response = null;
    let lastErr = '';

    for (const prov of providers) {
      try {
        const body = {
          model: prov.model,
          messages: [
            {
              role: 'user',
              content: [
                { type: 'image_url', image_url: { url: dataUrl } },
                { type: 'text', text: PROMPT }
              ]
            }
          ],
          max_tokens: 2048,
          temperature: 0.2,
          ...prov.extraBody
        };

        response = await fetch(prov.url, {
          method: 'POST',
          headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json',
          },
          body: JSON.stringify(body),
        });

        if (response.ok) break;
        const errBody = await response.text().catch(() => '');
        lastErr = `${prov.model}@${prov.url} → ${response.status}: ${errBody.slice(0, 300)}`;
        console.error('HF provider failed:', lastErr);
        response = null;
      } catch (e) {
        lastErr = `${prov.model}@${prov.url} → ${e.message}`;
        console.error('HF provider error:', lastErr);
        response = null;
      }
    }

    if (!response) {
      return res.status(502).json({ ok: false, error: `All providers failed. Last: ${lastErr}` });
    }

    const data = await response.json();
    const content = data.choices?.[0]?.message?.content || '';

    // Extraer JSON de la respuesta
    const jsonMatch = content.match(/\{[\s\S]*\}/);
    if (!jsonMatch) {
      return res.status(200).json({ ok: true, parsed: null, raw: content });
    }

    const parsed = JSON.parse(jsonMatch[0]);
    return res.status(200).json({ ok: true, parsed, raw: content });

  } catch (e) {
    console.error('ocr-hf error:', e.message);
    return res.status(500).json({ ok: false, error: e.message });
  }
}
