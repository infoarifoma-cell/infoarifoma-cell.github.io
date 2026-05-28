// POST /api/ocr
// Proxy para Gemini Flash — extrae datos de facturas con IA
// La API key queda en el servidor (env var)

export const config = { api: { bodyParser: { sizeLimit: '10mb' } } };

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ ok: false, error: 'Método no permitido' });
  }

  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) return res.status(500).json({ ok: false, error: 'GEMINI_API_KEY no configurada' });

  try {
    const { base64, mimeType } = req.body;
    if (!base64) return res.status(400).json({ ok: false, error: 'base64 requerido' });

    const resp = await fetch(
      'https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=' + apiKey,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          contents: [{
            parts: [
              {
                inlineData: {
                  mimeType: mimeType || 'image/png',
                  data: base64
                }
              },
              {
                text: `Analiza esta imagen de una factura o albarán. Extrae TODOS los datos que puedas ver.

Devuelve un JSON con esta estructura exacta (sin markdown, solo el JSON):
{
  "proveedor": "nombre del proveedor/emisor",
  "nFactura": "número de factura o albarán",
  "fecha": "YYYY-MM-DD",
  "textoCompleto": "transcripción completa de todo el texto visible en el documento"
}

Si no encuentras algún campo, déjalo como cadena vacía "".
La fecha debe estar en formato YYYY-MM-DD.
En textoCompleto pon TODO el texto que veas en la imagen, línea por línea.`
              }
            ]
          }]
        })
      }
    );

    const json = await resp.json();
    if (json.error) throw new Error(json.error.message || 'Gemini error');

    const text = json.candidates?.[0]?.content?.parts?.[0]?.text || '';

    // Intentar parsear JSON de la respuesta
    let parsed = null;
    try {
      const jsonMatch = text.match(/\{[\s\S]*\}/);
      if (jsonMatch) parsed = JSON.parse(jsonMatch[0]);
    } catch (e) { /* no pasa nada, devolvemos texto raw */ }

    return res.status(200).json({ ok: true, parsed, rawText: text });
  } catch (error) {
    console.error('Gemini OCR error:', error.message);
    return res.status(500).json({ ok: false, error: error.message });
  }
}
