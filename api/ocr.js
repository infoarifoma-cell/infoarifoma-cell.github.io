// POST /api/ocr
// Proxy para OCR.space — la API key queda en el servidor (env var)

export const config = { api: { bodyParser: { sizeLimit: '10mb' } } };

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ ok: false, error: 'Método no permitido' });
  }

  const apiKey = process.env.OCR_SPACE_KEY;
  if (!apiKey) return res.status(500).json({ ok: false, error: 'OCR_SPACE_KEY no configurada' });

  try {
    const { base64, filename, language, engine, isTable, scale, detectOrientation } = req.body;
    if (!base64) return res.status(400).json({ ok: false, error: 'base64 requerido' });

    const form = new URLSearchParams();
    form.append('apikey', apiKey);
    form.append('base64Image', 'data:image/png;base64,' + base64);
    form.append('language', language || 'spa');
    form.append('isOverlayRequired', 'false');
    form.append('scale', scale || 'true');
    form.append('isTable', isTable || 'true');
    form.append('detectOrientation', detectOrientation || 'true');
    form.append('OCREngine', String(engine || 1));
    if (filename) form.append('filename', filename);

    const resp = await fetch('https://api.ocr.space/parse/image', {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: form.toString()
    });

    const json = await resp.json();
    return res.status(200).json(json);
  } catch (error) {
    console.error('OCR proxy error:', error.message);
    return res.status(500).json({ ok: false, error: error.message });
  }
}
