// POST /api/ocr
// Proxy para OCR.space — la API key queda en el servidor (env var)

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ ok: false, error: 'Método no permitido' });
  }

  const apiKey = process.env.OCR_SPACE_KEY;
  if (!apiKey) return res.status(500).json({ ok: false, error: 'OCR_SPACE_KEY no configurada' });

  try {
    // Reenviar el FormData tal cual a OCR.space
    const contentType = req.headers['content-type'];

    // Leer body raw
    const chunks = [];
    for await (const chunk of req) {
      chunks.push(chunk);
    }
    const body = Buffer.concat(chunks);

    const resp = await fetch('https://api.ocr.space/parse/image', {
      method: 'POST',
      headers: { 'Content-Type': contentType },
      body: body
    });

    const json = await resp.json();
    return res.status(200).json(json);
  } catch (error) {
    console.error('OCR proxy error:', error.message);
    return res.status(500).json({ ok: false, error: error.message });
  }
}
