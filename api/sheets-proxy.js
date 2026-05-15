// POST /api/sheets-proxy
// Proxy para escribir en Google Sheets (evita CORS)

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ ok: false, error: 'Método no permitido' });
  }

  const SHEETS_API = 'https://script.google.com/macros/s/AKfycbwPIIgZCg03i4aJN8HIxKf20P5IPc-j3HOkoHmt2Jx0-vqiWrmq4Gz2WZmZvyopYJlv/exec';

  try {
    // Google Apps Script: POST → 302 redirect → GET a URL final
    // Seguir toda la cadena de redirects manualmente
    let url = SHEETS_API;
    let method = 'POST';
    let body = JSON.stringify(req.body);
    let finalResponse;

    for (let i = 0; i < 5; i++) {
      const opts = {
        method,
        redirect: 'manual',
        headers: method === 'POST' ? { 'Content-Type': 'text/plain' } : {}
      };
      if (method === 'POST') opts.body = body;

      const r = await fetch(url, opts);

      if (r.status >= 300 && r.status < 400) {
        url = r.headers.get('location');
        if (!url) throw new Error('Redirect sin location');
        method = 'GET';
        body = null;
        continue;
      }

      finalResponse = r;
      break;
    }

    if (!finalResponse) throw new Error('Demasiados redirects');

    const text = await finalResponse.text();
    let json;
    try { json = JSON.parse(text); } catch { json = { ok: false, error: text.substring(0, 300) }; }

    return res.status(200).json(json);
  } catch (error) {
    return res.status(500).json({ ok: false, error: error.message });
  }
}
