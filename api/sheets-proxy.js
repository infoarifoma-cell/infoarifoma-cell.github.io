// POST /api/sheets-proxy
// Proxy para escribir en Google Sheets (evita CORS)

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ ok: false, error: 'Método no permitido' });
  }

  const SHEETS_API = 'https://script.google.com/macros/s/AKfycbwPIIgZCg03i4aJN8HIxKf20P5IPc-j3HOkoHmt2Jx0-vqiWrmq4Gz2WZmZvyopYJlv/exec';

  try {
    const response = await fetch(SHEETS_API, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain' },
      body: JSON.stringify(req.body),
      redirect: 'follow'
    });

    const text = await response.text();
    let json;
    try { json = JSON.parse(text); } catch { json = { ok: false, error: text }; }

    return res.status(200).json(json);
  } catch (error) {
    return res.status(500).json({ ok: false, error: error.message });
  }
}
