// POST /api/sheets-proxy
// Proxy para escribir en Google Sheets (evita CORS)

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ ok: false, error: 'Método no permitido' });
  }

  const SHEETS_API = 'https://script.google.com/macros/s/AKfycbwPIIgZCg03i4aJN8HIxKf20P5IPc-j3HOkoHmt2Jx0-vqiWrmq4Gz2WZmZvyopYJlv/exec';

  try {
    // Google Apps Script responde con 302 redirect tras POST
    // Hay que seguir el redirect manualmente
    const response = await fetch(SHEETS_API, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain' },
      body: JSON.stringify(req.body),
      redirect: 'manual'
    });

    let finalText;

    if (response.status >= 300 && response.status < 400) {
      // Seguir redirect con GET (como hace el navegador)
      const redirectUrl = response.headers.get('location');
      if (!redirectUrl) throw new Error('Redirect sin location');
      const r2 = await fetch(redirectUrl, { method: 'GET', redirect: 'follow' });
      finalText = await r2.text();
    } else {
      finalText = await response.text();
    }

    let json;
    try { json = JSON.parse(finalText); } catch { json = { ok: false, error: finalText.substring(0, 200) }; }

    return res.status(200).json(json);
  } catch (error) {
    return res.status(500).json({ ok: false, error: error.message });
  }
}
