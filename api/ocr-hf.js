// POST /api/ocr-hf
// Envía imagen de factura a Qwen2.5-VL vía Hugging Face Inference API
// Requiere env var HF_TOKEN en Vercel

export const config = { api: { bodyParser: { sizeLimit: '10mb' } } };

const PROMPT = `Analiza esta imagen de factura/albarán y extrae los siguientes datos en formato JSON:
{
  "proveedor": "nombre del proveedor/empresa emisora",
  "nfactura": "número de factura o albarán",
  "fecha": "fecha en formato YYYY-MM-DD",
  "base_imponible": numero sin IVA (solo número, sin símbolo €),
  "iva_porcentaje": porcentaje IVA (solo número),
  "total": importe total (solo número, sin símbolo €)
}
Si no encuentras algún campo, pon null. Responde SOLO con el JSON, sin explicaciones.`;

export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ ok: false, error: 'Método no permitido' });

  const token = process.env.HF_TOKEN;
  if (!token) return res.status(500).json({ ok: false, error: 'HF_TOKEN no configurado' });

  try {
    const { image } = req.body; // base64 data URL
    if (!image) return res.status(400).json({ ok: false, error: 'Falta imagen' });

    // Extraer base64 puro y mime type
    const match = image.match(/^data:(image\/\w+);base64,(.+)$/);
    if (!match) return res.status(400).json({ ok: false, error: 'Formato de imagen inválido' });

    const mimeType = match[1];
    const base64Data = match[2];

    // Intentar múltiples providers en orden
    const providers = [
      {
        url: 'https://router.huggingface.co/hf-inference/models/Qwen/Qwen2.5-VL-3B-Instruct/v1/chat/completions',
        model: 'Qwen/Qwen2.5-VL-3B-Instruct',
      },
      {
        url: 'https://router.huggingface.co/hf-inference/models/Qwen/Qwen2.5-VL-7B-Instruct/v1/chat/completions',
        model: 'Qwen/Qwen2.5-VL-7B-Instruct',
      },
    ];

    let response = null;
    let lastErr = '';
    for (const prov of providers) {
      try {
        response = await fetch(prov.url, {
          method: 'POST',
          headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({
            model: prov.model,
            messages: [
              {
                role: 'user',
                content: [
                  { type: 'text', text: PROMPT },
                  { type: 'image_url', image_url: { url: `data:${mimeType};base64,${base64Data}` } }
                ]
              }
            ],
            max_tokens: 1024,
            temperature: 0.1,
          }),
        });
        if (response.ok) break;
        const errBody = await response.text().catch(() => '');
        lastErr = `${prov.url} → ${response.status}: ${errBody.slice(0, 300)}`;
        console.error('HF provider failed:', lastErr);
        response = null;
      } catch (e) {
        lastErr = `${prov.url} → ${e.message}`;
        console.error('HF provider error:', lastErr);
        response = null;
      }
    }

    if (!response) {
      return res.status(502).json({ ok: false, error: `All HF providers failed. Last: ${lastErr}` });
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
