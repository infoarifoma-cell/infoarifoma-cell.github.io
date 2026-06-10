// POST /api/google-sheet
// Proxy para leer Google Sheets con Service Account
// Body: { spreadsheetId, sheet, range }

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ ok: false, error: 'Método no permitido' });
  }

  const { spreadsheetId, sheet, range } = req.body;
  if (!spreadsheetId) return res.status(400).json({ ok: false, error: 'spreadsheetId requerido' });

  const clientEmail = process.env.GOOGLE_CLIENT_EMAIL;
  const privateKey = (process.env.GOOGLE_PRIVATE_KEY || '').replace(/\\n/g, '\n');

  if (!clientEmail || !privateKey) {
    return res.status(500).json({ ok: false, error: 'Credenciales Google no configuradas' });
  }

  try {
    // Generar JWT para Google OAuth2
    const now = Math.floor(Date.now() / 1000);
    const payload = {
      iss: clientEmail,
      scope: 'https://www.googleapis.com/auth/spreadsheets.readonly',
      aud: 'https://oauth2.googleapis.com/token',
      exp: now + 3600,
      iat: now
    };

    // Crear JWT manualmente (header.payload.signature)
    const encode = obj => btoa(JSON.stringify(obj)).replace(/\+/g,'-').replace(/\//g,'_').replace(/=/g,'');
    const header = encode({ alg: 'RS256', typ: 'JWT' });
    const body = encode(payload);
    const unsigned = `${header}.${body}`;

    // Importar clave privada y firmar
    const keyData = privateKey
      .replace('-----BEGIN PRIVATE KEY-----', '')
      .replace('-----END PRIVATE KEY-----', '')
      .replace(/\s/g, '');
    const binaryKey = Uint8Array.from(atob(keyData), c => c.charCodeAt(0));
    const cryptoKey = await crypto.subtle.importKey(
      'pkcs8', binaryKey.buffer,
      { name: 'RSASSA-PKCS1-v1_5', hash: 'SHA-256' },
      false, ['sign']
    );
    const signature = await crypto.subtle.sign(
      'RSASSA-PKCS1-v1_5', cryptoKey,
      new TextEncoder().encode(unsigned)
    );
    const sig = btoa(String.fromCharCode(...new Uint8Array(signature)))
      .replace(/\+/g,'-').replace(/\//g,'_').replace(/=/g,'');
    const jwt = `${unsigned}.${sig}`;

    // Obtener access token
    const tokenRes = await fetch('https://oauth2.googleapis.com/token', {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: `grant_type=urn%3Aietf%3Aparams%3Aoauth%3Agrant-type%3Ajwt-bearer&assertion=${jwt}`
    });
    if (!tokenRes.ok) throw new Error('Error obteniendo token: ' + await tokenRes.text());
    const { access_token } = await tokenRes.json();

    // Leer Sheet
    const sheetRange = range || (sheet ? `${sheet}!A:H` : 'A:H');
    const sheetUrl = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(sheetRange)}`;
    const sheetRes = await fetch(sheetUrl, {
      headers: { 'Authorization': `Bearer ${access_token}` }
    });
    if (!sheetRes.ok) throw new Error('Error leyendo sheet: ' + await sheetRes.text());
    const sheetJson = await sheetRes.json();

    return res.status(200).json({ ok: true, values: sheetJson.values || [] });
  } catch (e) {
    console.error('google-sheet error:', e.message);
    return res.status(500).json({ ok: false, error: e.message });
  }
}
