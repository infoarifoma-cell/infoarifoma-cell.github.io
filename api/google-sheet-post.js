// POST /api/google-sheet-post
// Proxy para enviar datos a Google Apps Script (Producción y Gasoil)

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ ok: false, error: 'Método no permitido' });
  }

  const SHEETS_API = process.env.SHEETS_API;
  if (!SHEETS_API) {
    return res.status(500).json({ ok: false, error: 'SHEETS_API no configurada' });
  }

  try {
    const response = await fetch(SHEETS_API, {
      method: 'POST',
      redirect: 'follow',
      headers: { 'Content-Type': 'text/plain' },
      body: JSON.stringify(req.body)
    });

    if (!response.ok) {
      return res.status(response.status).json({ ok: false, error: 'HTTP ' + response.status });
    }

    const json = await response.json();
    return res.status(200).json(json);
  } catch (e) {
    console.error('google-sheet-post error:', e.message);
    return res.status(500).json({ ok: false, error: e.message });
  }
}
