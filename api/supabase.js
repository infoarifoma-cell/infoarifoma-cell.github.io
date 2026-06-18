// POST /api/supabase
// Proxy seguro: Frontend envía acción + datos, backend ejecuta contra Supabase con service_role key
// La key NUNCA se expone al navegador — solo vive en env vars de Vercel

export default async function handler(req, res) {
  // Solo POST
  if (req.method !== 'POST') {
    return res.status(405).json({ ok: false, error: 'Método no permitido' });
  }

  const SUPABASE_URL = process.env.SUPABASE_URL;
  const SUPABASE_SERVICE_KEY = process.env.SUPABASE_SERVICE_KEY;

  if (!SUPABASE_URL || !SUPABASE_SERVICE_KEY) {
    return res.status(500).json({ ok: false, error: 'Config servidor incompleta' });
  }

  const { action, table, data, filters, options } = req.body;

  // ── Validar token de sesión del usuario ──
  const authHeader = req.headers.authorization;
  if (!authHeader || !authHeader.startsWith('Bearer ')) {
    return res.status(401).json({ ok: false, error: 'No autorizado' });
  }
  const userToken = authHeader.replace('Bearer ', '');

  // Verificar token contra Supabase Auth
  const userRes = await fetch(`${SUPABASE_URL}/auth/v1/user`, {
    headers: {
      'apikey': SUPABASE_SERVICE_KEY,
      'Authorization': `Bearer ${userToken}`
    }
  });
  if (!userRes.ok) {
    return res.status(401).json({ ok: false, error: 'Token inválido o expirado' });
  }

  const userData = await userRes.json();

  // ── Bloquear escritura para rol lectura ──
  if (['insert', 'update', 'delete'].includes(action)) {
    const email = userData.email;
    if (email) {
      const rolRes = await fetch(`${SUPABASE_URL}/rest/v1/tblUsuarios?select=rol&email=eq.${encodeURIComponent(email)}&limit=1`, {
        headers: { 'apikey': SUPABASE_SERVICE_KEY, 'Authorization': `Bearer ${SUPABASE_SERVICE_KEY}` }
      });
      if (rolRes.ok) {
        const rows = await rolRes.json();
        if (rows.length && rows[0].rol === 'lectura') {
          return res.status(403).json({ ok: false, error: 'Sin permisos de escritura' });
        }
      }
    }
  }

  // ── Whitelist de tablas permitidas ──
  const ALLOWED_TABLES = [
    'tblFichaje', 'tblpedidos', 'tblcamiones', 'tblobras',
    'PRODUCCION', 'GASOIL', 'GASOIL_STOCK', 'horometros',
    'tblGamasOT', 'tblAusencias', 'tblControlDocumental',
    'tblGamasNormas', 'tblGamasDependientes', 'tblGamasActivos',
    'tblGamasListadoPreventivo', 'tblUsuarios',
    'tblactivos', 'tblcaja', 'tblnotas', 'tblchoferes', 'tblTareas',
    'ensayos_semanas', 'ensayos_registros', 'ensayos_prestaciones', 'ensayos_limites', 'ensayos_anuales',
    'tblTopografia', 'tblRegularizaciones', 'tblPesoBascula'
  ];

  if (!table || !ALLOWED_TABLES.includes(table)) {
    return res.status(400).json({ ok: false, error: 'Tabla no permitida: ' + table });
  }

  // ── Whitelist de acciones ──
  const ALLOWED_ACTIONS = ['select', 'insert', 'update', 'delete'];
  if (!action || !ALLOWED_ACTIONS.includes(action)) {
    return res.status(400).json({ ok: false, error: 'Acción no permitida: ' + action });
  }

  const headers = {
    'apikey': SUPABASE_SERVICE_KEY,
    'Authorization': `Bearer ${SUPABASE_SERVICE_KEY}`,
    'Content-Type': 'application/json',
    'Prefer': action === 'insert' ? 'return=representation' : 'return=minimal',
    'Range-Unit': 'items',
    'Range': '0-9999'
  };

  try {
    let url = `${SUPABASE_URL}/rest/v1/${table}`;
    let method = 'GET';
    let body = null;

    // ── SELECT ──
    if (action === 'select') {
      method = 'GET';
      const params = [];
      params.push('select=' + encodeURIComponent(options?.select || '*'));
      if (options?.order) params.push('order=' + encodeURIComponent(options.order));
      if (options?.limit) params.push('limit=' + encodeURIComponent(options.limit));
      if (options?.offset) params.push('offset=' + encodeURIComponent(options.offset));
      // Filtros: [{column, op, value}]  op='in' usa sintaxis Supabase: column=in.(v1,v2,...)
      if (filters && Array.isArray(filters)) {
        for (const f of filters) {
          if (f.op === 'in' && Array.isArray(f.value)) {
            params.push(`${encodeURIComponent(f.column)}=in.(${f.value.join(',')})`);
          } else {
            params.push(`${encodeURIComponent(f.column)}=${encodeURIComponent(f.op)}.${encodeURIComponent(f.value)}`);
          }
        }
      }
      url += '?' + params.join('&');
    }

    // ── INSERT ──
    else if (action === 'insert') {
      method = 'POST';
      body = JSON.stringify(Array.isArray(data) ? data : [data]);
      if (options?.select) {
        url += '?select=' + encodeURIComponent(options.select);
      }
    }

    // ── UPDATE ──
    else if (action === 'update') {
      method = 'PATCH';
      body = JSON.stringify(data);
      // Requiere al menos un filtro para no actualizar toda la tabla
      if (!filters || !filters.length) {
        return res.status(400).json({ ok: false, error: 'Update requiere filtros' });
      }
      const params = [];
      for (const f of filters) {
        params.push(`${encodeURIComponent(f.column)}=${encodeURIComponent(f.op)}.${encodeURIComponent(f.value)}`);
      }
      url += '?' + params.join('&');
    }

    // ── DELETE ──
    else if (action === 'delete') {
      method = 'DELETE';
      if (!filters || !filters.length) {
        return res.status(400).json({ ok: false, error: 'Delete requiere filtros' });
      }
      const params = [];
      for (const f of filters) {
        params.push(`${encodeURIComponent(f.column)}=${encodeURIComponent(f.op)}.${encodeURIComponent(f.value)}`);
      }
      url += '?' + params.join('&');
    }

    const response = await fetch(url, { method, headers, body });
    const text = await response.text();

    if (!response.ok) {
      console.error('Supabase error:', response.status, text);
      return res.status(response.status).json({ ok: false, error: text });
    }

    // SELECT y INSERT devuelven JSON, UPDATE/DELETE devuelven vacío
    if (action === 'select' || action === 'insert') {
      const json = text ? JSON.parse(text) : [];
      return res.status(200).json({ ok: true, data: json });
    }
    return res.status(200).json({ ok: true });

  } catch (e) {
    console.error('Proxy error:', e);
    return res.status(500).json({ ok: false, error: e.message });
  }
}
