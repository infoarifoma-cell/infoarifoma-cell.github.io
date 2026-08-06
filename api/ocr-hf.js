// POST /api/ocr-hf
// Envía imagen de factura a NuExtract3 o Qwen VL vía Hugging Face Inference API
// Requiere env var HF_TOKEN en Vercel

export const config = { api: { bodyParser: { sizeLimit: '10mb' } } };

const TEMPLATE = JSON.stringify({
  proveedor: "verbatim-string",
  nfactura: "verbatim-string",
  fecha: "date",
  base_imponible: "number",
  iva_porcentaje: "number",
  total: "number"
});

const PROMPT = `Extract the following fields from this invoice/receipt image.
Return ONLY valid JSON matching this schema: ${TEMPLATE}
If a field is not found, use null.`;

export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ ok: false, error: 'Método no permitido' });

  const token = process.env.HF_TOKEN;
  if (!token) return res.status(500).json({ ok: false, error: 'HF_TOKEN no configurado' });

  try {
    const { image } = req.body;
    if (!image) return res.status(400).json({ ok: false, error: 'Falta imagen' });

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
            template: TEMPLATE,
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
          max_tokens: 1024,
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
