// POST /api/gasoil-edit
// Proxy para editar fila de gasoil en Google Sheets via Apps Script

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
      body: JSON.stringify({ ...req.body, secret: 'ar1f0ma-2025-sh3ets' })
    });

    const text = await response.text();
    try {
      return res.status(200).json(JSON.parse(text));
    } catch {
      return res.status(200).json({ ok: false, error: 'Respuesta no JSON: ' + text.slice(0, 200) });
    }
  } catch (e) {
    return res.status(500).json({ ok: false, error: e.message });
  }
}
