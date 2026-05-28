// POST /api/ocr
// Proxy hacia la app OCR desplegada en Google Cloud Run

export const config = { api: { bodyParser: { sizeLimit: '25mb' } } };

const OCR_BACKEND = 'https://lector-ocr-de-facturas-272247425176.europe-west2.run.app/api/ocr';

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ ok: false, error: 'Método no permitido' });
  }

  try {
    const resp = await fetch(OCR_BACKEND, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(req.body)
    });

    const json = await resp.json();
    return res.status(resp.status).json(json);
  } catch (error) {
    console.error('OCR proxy error:', error.message);
    return res.status(500).json({ ok: false, error: error.message });
  }
}
