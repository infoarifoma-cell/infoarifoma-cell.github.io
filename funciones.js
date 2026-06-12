// ============================================================
// ARIFOMA · FUNCIONES — versión segura (proxy backend)
// ============================================================

// ── INPUT VALIDATION ───────────────────────────────────────────
function validateMatricula(val) {
  return val && /^[A-Z0-9]{1,20}$/.test(String(val).toUpperCase());
}
function validateEmail(val) {
  return val && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(val);
}
function validateNumber(val) {
  return !isNaN(parseFloat(val)) && isFinite(val);
}
function validateString(val, maxLen = 255) {
  return val && String(val).length > 0 && String(val).length <= maxLen;
}

// ── HTML SANITIZATION ───────────────────────────────────────────
function escapeHTML(text) {
  const div = document.createElement('div');
  div.textContent = text;
  return div.innerHTML;
}

// ── BUSINESS CENTRAL CONFIG ───────────────────────────────────
async function enviarLineaBCPesada(data) {
  try {
    if (typeof getBCToken !== 'function') {
      console.warn('BC: Función getBCToken no disponible');
      return;
    }
    const response = await fetch('/api/bc', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        action: 'linea-pesada',
        codigoCliente: data.codigoCliente,
        proyectoCod: data.proyectoCod,
        productoCod: data.productoCod,
        productoNombre: data.productoNombre,
        pesoNeto: data.pesoNeto,
        matriculacam: data.matriculacam,
        proyectoName: data.proyectoName
      })
    });
    if (!response.ok) {
      console.warn('BC: Error al enviar línea pesada');
      return;
    }
    console.log('BC línea enviada OK');
  } catch (e) {
    console.warn('BC error:', e.message);
  }
}

// ── SUPABASE AUTH (anon key solo para auth — es seguro, diseño oficial) ──
const SUPABASE_URL = 'https://bnsfgzjqmibsrklllqxb.supabase.co';
const SUPABASE_ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImJuc2ZnempxbWlic3JrbGxscXhiIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzQzNzYwNzksImV4cCI6MjA4OTk1MjA3OX0.8mTQHPdO954ICBd1Xam-kKmcA69CMyO2v3x1liFgWyk';
let _supabase = supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);
let _sessionToken = null;

function _initAuthClient(token) {
  _sessionToken = token;
}

async function cerrarSesion() {
  await _supabase.auth.signOut();
  _sessionToken = null;
  _supabase = supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);
  window._appInitialized = false;
}

// ── PROXY HELPER ────────────────────────────────────────────────
// Todas las operaciones de datos pasan por /api/supabase (backend seguro)
async function _refreshToken() {
  try {
    const { data: { session } } = await _supabase.auth.refreshSession();
    if (session && session.access_token) {
      _sessionToken = session.access_token;
      scheduleTokenRefresh(session.access_token);
      return true;
    }
  } catch (e) { console.warn('Token refresh failed:', e); }
  return false;
}

async function dbQuery({ action, table, data, filters, options }) {
  if (!_sessionToken) {
    return { ok: false, error: 'No hay sesión activa' };
  }
  const doFetch = () => fetch('/api/supabase', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + _sessionToken
    },
    body: JSON.stringify({ action, table, data, filters, options })
  });
  try {
    let res = await doFetch();
    // Si token expirado, refrescar y reintentar una vez
    if (res.status === 401) {
      const refreshed = await _refreshToken();
      if (refreshed) res = await doFetch();
    }
    const json = await res.json();
    return json;
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ── GOOGLE AUTH ──────────────────────────────────────────────
async function googleLogin() {
  try {
    document.getElementById('login-loading').style.display = 'block';
    const { data, error } = await _supabase.auth.signInWithOAuth({
      provider: 'google',
      options: {
        redirectTo: 'https://infoarifoma-cell.github.io/'
      }
    });
    if (error) {
      console.error('Google login error:', error.message);
      document.getElementById('login-error').textContent = 'Error: ' + error.message;
      document.getElementById('login-loading').style.display = 'none';
    }
  } catch(e) {
    console.error('Google login exception:', e.message);
    document.getElementById('login-error').textContent = 'Excepción: ' + e.message;
    document.getElementById('login-loading').style.display = 'none';
  }
}

let _checkingSession = false;
async function checkGoogleSession(existingSession) {
  if (_checkingSession || window._appInitialized) return;
  _checkingSession = true;
  try {
  const session = existingSession || (await _supabase.auth.getSession()).data.session;
  if (session && session.user) {
    const email = session.user.email;
    _sessionToken = session.access_token;

    // Verificar usuario en tblUsuarios via proxy
    const result = await dbQuery({
      action: 'select',
      table: 'tblUsuarios',
      filters: [{ column: 'email', op: 'eq', value: email }],
      options: { select: 'id,nombre,rol,activo' }
    });

    if (!result.ok || !result.data || result.data.length === 0) {
      document.getElementById('login-error').textContent = 'Credenciales inválidas o usuario no autorizado';
      await _supabase.auth.signOut();
      _sessionToken = null;
      return;
    }

    const usuarios = result.data[0];
    if (!usuarios.activo) {
      document.getElementById('login-error').textContent = 'Credenciales inválidas o usuario no autorizado';
      await _supabase.auth.signOut();
      _sessionToken = null;
      return;
    }

    loginUser = {
      id: usuarios.id,
      nombre: usuarios.nombre,
      email: email,
      rol: usuarios.rol,
      provider: 'google'
    };
    scheduleTokenRefresh(session.access_token);

    document.getElementById('login-loading').style.display = 'block';
    const r = await fetch('_shell.html?v=' + Date.now());
    if (!r.ok) throw new Error('HTTP ' + r.status + ' loading shell');
    const html = await r.text();
    const ph = document.getElementById('shell-placeholder');
    if (ph) {
      ph.insertAdjacentHTML('beforeend', html);
      document.getElementById('pinScreen').style.display = 'none';
      document.getElementById('shell').style.display = 'flex';
      setTimeout(() => initApp(), 50);
    }
  }
  } finally {
    _checkingSession = false;
  }
}

// Session timeout: 30 minutos de inactividad
let _sessionTimeout;
let _tokenRefreshTimeout;
const SESSION_TIMEOUT_MS = 30 * 60 * 1000;

function decodeJWT(token) {
  try {
    const parts = token.split('.');
    if (parts.length !== 3) return null;
    const decoded = JSON.parse(atob(parts[1]));
    return decoded;
  } catch (e) {
    return null;
  }
}

async function scheduleTokenRefresh(token) {
  clearTimeout(_tokenRefreshTimeout);
  const decoded = decodeJWT(token);
  if (!decoded || !decoded.exp) return;

  const expiresAt = decoded.exp * 1000;
  const now = Date.now();
  const timeUntilExpiry = expiresAt - now;
  const refreshTime = timeUntilExpiry - (5 * 60 * 1000);

  if (refreshTime > 0) {
    _tokenRefreshTimeout = setTimeout(async () => {
      console.log('Refrescando token Google...');
      const { data, error } = await _supabase.auth.refreshSession();
      const session = data?.session;
      if (session && session.access_token) {
        _initAuthClient(session.access_token);
        scheduleTokenRefresh(session.access_token);
      } else {
        console.warn('Error refrescando token Google:', error?.message);
        // Token irrecuperable — forzar re-login
        await cerrarSesion();
        document.getElementById('login-error').textContent = 'Sesión expirada. Inicie sesión nuevamente.';
        document.getElementById('pinScreen').style.display = 'flex';
        document.getElementById('shell').style.display = 'none';
      }
    }, refreshTime);
  }
}

function resetSessionTimeout() {
  clearTimeout(_sessionTimeout);
  _sessionTimeout = setTimeout(async () => {
    console.warn('Sesión expirada por inactividad');
    await cerrarSesion();
    document.getElementById('login-error').textContent = 'Sesión expirada. Inicie sesión nuevamente.';
    document.getElementById('pinScreen').style.display = 'flex';
    document.getElementById('shell').style.display = 'none';
  }, SESSION_TIMEOUT_MS);
}

['mousedown', 'keydown', 'scroll', 'touchstart', 'click'].forEach(event => {
  document.addEventListener(event, resetSessionTimeout, true);
});

window.appInitialized = false;

// Detectar retorno de OAuth: mostrar loading inmediatamente para evitar flash del login
(function() {
  const hash = window.location.hash || '';
  const search = window.location.search || '';
  if (hash.includes('access_token') || search.includes('code=')) {
    const loading = document.getElementById('login-loading');
    if (loading) loading.style.display = 'block';
  }
})();

// Usar onAuthStateChange para reaccionar a sesión sin polling
_supabase.auth.onAuthStateChange((event, session) => {
  if ((event === 'SIGNED_IN' || event === 'INITIAL_SESSION') && session && session.user) {
    if (!window.appInitialized) {
      window.appInitialized = true;
      checkGoogleSession(session).catch(e => console.error('checkGoogleSession error:', e));
      resetSessionTimeout();
    }
  }
});

// ── FICHAJES ─────────────────────────────────────────────────

async function getFichajes() {
  const result = await dbQuery({
    action: 'select',
    table: 'tblFichaje',
    options: { select: '*', order: 'fentrada.desc', limit: 300 }
  });
  return result;
}

async function doPostEntrada(data) {
  const result = await dbQuery({
    action: 'insert',
    table: 'tblFichaje',
    data: {
      empleado:    data.empleado,
      fecha:       data.fecha    || new Date().toISOString().slice(0,10),
      entrada:     data.entrada  || new Date().toLocaleTimeString('es-ES', { hour: '2-digit', minute: '2-digit' }),
      tipoTrabajo: data.tipoTrabajo || 'JORNADA',
      fentrada:    data.fentrada || new Date().toISOString()
    }
  });
  if (!result.ok) { console.error('doPostEntrada error:', result.error); }
  return result;
}

async function doEditSalida(data) {
  // Buscar última entrada abierta
  const search = await dbQuery({
    action: 'select',
    table: 'tblFichaje',
    filters: [
      { column: 'empleado', op: 'eq', value: data.empleado },
      { column: 'salida', op: 'is', value: 'null' }
    ],
    options: { select: 'id', order: 'fentrada.desc', limit: 1 }
  });

  const registro = search.ok && search.data && search.data.length ? search.data[0] : null;

  if (!registro || !registro.id) {
    // Insertar registro solo con salida
    const insResult = await dbQuery({
      action: 'insert',
      table: 'tblFichaje',
      data: {
        empleado: data.empleado,
        salida: data.salida,
        fsalida: data.fsalida || new Date().toISOString(),
        tiempodia: data.tiempodia
      }
    });
    return insResult;
  }

  const result = await dbQuery({
    action: 'update',
    table: 'tblFichaje',
    data: {
      salida:    data.salida,
      fsalida:   data.fsalida || new Date().toISOString(),
      tiempodia: data.tiempodia
    },
    filters: [{ column: 'id', op: 'eq', value: registro.id }]
  });
  return result;
}

async function doEditFichaje(data) {
  const updates = {};
  if (data.empleado  !== undefined) updates.empleado  = data.empleado;
  if (data.fecha     !== undefined) updates.fecha     = data.fecha;
  if (data.entrada   !== undefined) updates.entrada   = data.entrada;
  if (data.salida    !== undefined) updates.salida    = data.salida;
  if (data.tiempodia !== undefined) updates.tiempodia = data.tiempodia;
  if (data.fentrada  !== undefined) updates.fentrada  = data.fentrada;
  if (data.fsalida   !== undefined) updates.fsalida   = data.fsalida;
  const result = await dbQuery({
    action: 'update',
    table: 'tblFichaje',
    data: updates,
    filters: [{ column: 'id', op: 'eq', value: data.id }]
  });
  return result;
}

async function doDeleteFichaje(data) {
  return dbQuery({
    action: 'delete',
    table: 'tblFichaje',
    filters: [{ column: 'id', op: 'eq', value: data.id }]
  });
}

// ── PEDIDOS (Pesajes) ────────────────────────────────────────

async function getPedidos(diasAtras = 90) {
  const corte = new Date();
  corte.setDate(corte.getDate() - diasAtras);
  return dbQuery({
    action: 'select',
    table: 'tblpedidos',
    filters: [{ column: 'fechaHora', op: 'gte', value: corte.toISOString() }],
    options: { select: '*', order: 'fechaHora.desc' }
  });
}

async function doPostPesada(data) {
  if (!validateMatricula(data.matriculacam)) return { ok: false, error: 'Matrícula vehículo inválida' };
  if (!validateNumber(data.pesoBruto) || Number(data.pesoBruto) <= 0) return { ok: false, error: 'Peso bruto inválido' };
  if (!validateNumber(data.pesoNeto) || Number(data.pesoNeto) < 0) return { ok: false, error: 'Peso neto inválido' };
  if (!validateString(data.chofer, 100)) return { ok: false, error: 'Chofer requerido' };
  if (!validateString(data.nombreCliente, 150)) return { ok: false, error: 'Cliente requerido' };

  const result = await dbQuery({
    action: 'insert',
    table: 'tblpedidos',
    data: {
      matriculacam:  String(data.matriculacam).toUpperCase(),
      matricularem:  data.matricularem ? String(data.matricularem).toUpperCase() : null,
      tara:          Number(data.tara || 0),
      chofer:        String(data.chofer).substring(0, 100),
      nombreCliente: String(data.nombreCliente).substring(0, 150),
      codigoCliente: data.codigoCliente ? String(data.codigoCliente).substring(0, 50) : null,
      productoNombre:data.productoNombre ? String(data.productoNombre).substring(0, 150) : null,
      productoCod:   data.productoCod ? String(data.productoCod).substring(0, 50) : null,
      pesoBruto:     Number(data.pesoBruto),
      pesoNeto:      Number(data.pesoNeto),
      proyectoName:  data.proyectoName ? String(data.proyectoName).substring(0, 150) : null,
      proyectoCod:   data.proyectoCod ? String(data.proyectoCod).substring(0, 50) : null,
      numPedido:     data.numPedido ? String(data.numPedido).substring(0, 50) : null,
      numLinea:      data.numLinea ? String(data.numLinea).substring(0, 50) : null,
      fechaHora:     new Date().toISOString()
    },
    options: { select: 'id' }
  });
  if (!result.ok) return { ok: false, error: 'Error al grabar pesada. Contacte administrador.' };
  if (typeof enviarLineaBCPesada === 'function') {
    enviarLineaBCPesada(data).catch(e => console.warn('BC línea:', e.message));
  }
  return { ok: true, id: result.data && result.data[0] ? result.data[0].id : null };
}

async function doDeletePedido(data) {
  const id = Number(data.id);
  if (!id || isNaN(id)) return { ok: false, error: 'ID inválido' };
  return dbQuery({
    action: 'delete',
    table: 'tblpedidos',
    filters: [{ column: 'id', op: 'eq', value: id }]
  });
}

async function doEditarPedido(data) {
  const updates = {};
  if (data.matriculacam  !== undefined) updates.matriculacam  = data.matriculacam;
  if (data.matricularem  !== undefined) updates.matricularem  = data.matricularem;
  if (data.chofer        !== undefined) updates.chofer        = data.chofer;
  if (data.codigoCliente !== undefined) updates.codigoCliente = data.codigoCliente;
  if (data.nombreCliente !== undefined) updates.nombreCliente = data.nombreCliente;
  if (data.productoCod   !== undefined) updates.productoCod   = data.productoCod;
  if (data.productoNombre!== undefined) updates.productoNombre= data.productoNombre;
  if (data.pesoBruto     !== undefined) updates.pesoBruto     = Number(data.pesoBruto);
  if (data.pesoNeto      !== undefined) updates.pesoNeto      = Number(data.pesoNeto);
  if (data.proyectoCod   !== undefined) updates.proyectoCod   = data.proyectoCod;
  if (data.proyectoName  !== undefined) updates.proyectoName  = data.proyectoName;
  return dbQuery({
    action: 'update',
    table: 'tblpedidos',
    data: updates,
    filters: [{ column: 'id', op: 'eq', value: data.id }]
  });
}

// ── CAMIONES ─────────────────────────────────────────────────

async function getCamiones() {
  return dbQuery({
    action: 'select',
    table: 'tblcamiones',
    options: { select: '*', order: 'matriculacam.asc' }
  });
}

async function doNuevoCamion(data) {
  const { tipo, id, ...campos } = data;
  campos.fechaCreacion = new Date().toISOString();
  return dbQuery({ action: 'insert', table: 'tblcamiones', data: campos });
}

async function doEditarCamion(data) {
  const { id, tipo, ...updates } = data;
  return dbQuery({
    action: 'update',
    table: 'tblcamiones',
    data: updates,
    filters: [{ column: 'id', op: 'eq', value: id }]
  });
}

async function doEliminarCamion(data) {
  const id = Number(data.id);
  if (!id || isNaN(id)) return { ok: false, error: 'ID inválido' };
  return dbQuery({
    action: 'delete',
    table: 'tblcamiones',
    filters: [{ column: 'id', op: 'eq', value: id }]
  });
}

// ── OBRAS / PROYECTOS ───────────────────────────────────────

async function getObras() {
  return dbQuery({
    action: 'select',
    table: 'tblobras',
    options: { select: '*', order: 'codigo.asc' }
  });
}

async function doNuevaObra(data) {
  const { tipo, id, ...campos } = data;
  campos.fechaCreacion = new Date().toISOString();
  return dbQuery({ action: 'insert', table: 'tblobras', data: campos });
}

async function doEditarObra(data) {
  const { id, tipo, ...updates } = data;
  return dbQuery({
    action: 'update',
    table: 'tblobras',
    data: updates,
    filters: [{ column: 'id', op: 'eq', value: id }]
  });
}

async function doEliminarObra(data) {
  const id = Number(data.id);
  if (!id || isNaN(id)) return { ok: false, error: 'ID inválido' };
  return dbQuery({
    action: 'delete',
    table: 'tblobras',
    filters: [{ column: 'id', op: 'eq', value: id }]
  });
}

// ── CHOFERES ─────────────────────────────────────────────────

async function getChoferes() {
  return dbQuery({
    action: 'select',
    table: 'tblchoferes',
    options: { select: '*', order: 'nombre.asc' }
  });
}

async function doNuevoChofer(data) {
  const { tipo, id, ...campos } = data;
  campos.fechacreacion = new Date().toISOString();
  return dbQuery({ action: 'insert', table: 'tblchoferes', data: campos });
}

async function doEditarChofer(data) {
  const { id, tipo, ...updates } = data;
  return dbQuery({
    action: 'update',
    table: 'tblchoferes',
    data: updates,
    filters: [{ column: 'id', op: 'eq', value: id }]
  });
}

async function doEliminarChofer(data) {
  const id = Number(data.id);
  if (!id || isNaN(id)) return { ok: false, error: 'ID inválido' };
  return dbQuery({
    action: 'delete',
    table: 'tblchoferes',
    filters: [{ column: 'id', op: 'eq', value: id }]
  });
}

// ── PRODUCCIÓN ───────────────────────────────────────────────

async function getProduccion(mes, anyo) {
  const filters = [];
  if (anyo) {
    if (mes) {
      const start = `${anyo}-${String(mes).padStart(2, '0')}-01`;
      const end   = `${anyo}-${String(mes).padStart(2, '0')}-31`;
      filters.push({ column: 'fecha', op: 'gte', value: start });
      filters.push({ column: 'fecha', op: 'lte', value: end });
    } else {
      filters.push({ column: 'fecha', op: 'gte', value: `${anyo}-01-01` });
      filters.push({ column: 'fecha', op: 'lte', value: `${anyo}-12-31` });
    }
  }
  return dbQuery({
    action: 'select',
    table: 'PRODUCCION',
    filters,
    options: { select: '*', order: 'fecha.asc' }
  });
}

async function doEditProduccion(data) {
  const { id, ...campos } = data;
  return dbQuery({
    action: 'update',
    table: 'PRODUCCION',
    data: {
      tipoDia:      campos.tipoDia,
      tnDia:        Number(campos.tnDia),
      t04:          Number(campos.t04),
      t04h:         Number(campos.t04h),
      t412:         Number(campos.t412),
      t412h:        Number(campos.t412h),
      t1220:        Number(campos.t1220),
      t1220h:       Number(campos.t1220h),
      t2040:        Number(campos.t2040),
      t2040h:       Number(campos.t2040h),
      otroTipo:     campos.otroTipo,
      otroTn:       Number(campos.otroTn),
      otroTnh:      Number(campos.otroTnh),
      horasPlanta:  Number(campos.horasPlanta),
      observaciones:campos.observaciones
    },
    filters: [{ column: 'id', op: 'eq', value: id }]
  });
}

async function doAddProduccion(data) {
  return dbQuery({
    action: 'insert',
    table: 'PRODUCCION',
    data: {
      fecha:         data.fecha,
      tipoDia:       data.tipoDia,
      tnDia:         Number(data.tnDia  || 0),
      t04:           Number(data.t04    || 0),
      t04h:          Number(data.t04h   || 0),
      t412:          Number(data.t412   || 0),
      t412h:         Number(data.t412h  || 0),
      t1220:         Number(data.t1220  || 0),
      t1220h:        Number(data.t1220h || 0),
      t2040:         Number(data.t2040  || 0),
      t2040h:        Number(data.t2040h || 0),
      otroTipo:      data.otroTipo  || '',
      otroTn:        Number(data.otroTn  || 0),
      otroTnh:       Number(data.otroTnh || 0),
      horasPlanta:   Number(data.horasPlanta || 0),
      observaciones: data.observaciones || ''
    }
  });
}

// ── GASOIL ───────────────────────────────────────────────────

async function getGasoil() {
  const [gasoilRes, horoRes, stockRes] = await Promise.all([
    dbQuery({ action: 'select', table: 'GASOIL', options: { select: '*', order: 'fecha.desc' } }),
    dbQuery({ action: 'select', table: 'horometros', options: { select: '*' } }),
    dbQuery({ action: 'select', table: 'GASOIL_STOCK', options: { select: '*' } })
  ]);
  if (!gasoilRes.ok) return { ok: false, error: gasoilRes.error };

  const data = gasoilRes.data || [];
  let dep1 = 0, dep2 = 0;
  (stockRes.data || []).forEach(r => {
    if (r.deposito === 'DEP1') dep1 = Number(r.stock) || 0;
    if (r.deposito === 'DEP2') dep2 = Number(r.stock) || 0;
  });

  const consumos = (horoRes.data || []).map(r => ({
    activo:      r.activo,
    max:         r.horometro,
    actualizado: r.actualizado ? r.actualizado.slice(0, 10) : null
  }));

  return { ok: true, data, dep1, dep2, consumos };
}

async function doPostGasoil(data) {
  return dbQuery({
    action: 'insert',
    table: 'GASOIL',
    data: {
      fecha:      data.fecha,
      proveedor:  data.proveedor,
      origen:     data.origen,
      destino:    data.destino,
      tipo:       data.tipoMovimiento,
      litros:     Number(data.litros),
      horometro:  data.horometro ? Number(data.horometro) : null
    }
  });
}

async function doEditarGasoil(data) {
  return dbQuery({
    action: 'update',
    table: 'GASOIL',
    data: {
      fecha:      data.fecha,
      proveedor:  data.proveedor,
      origen:     data.origen,
      destino:    data.destino,
      tipo:       data.tipoMovimiento,
      litros:     Number(data.litros),
      horometro:  data.horometro ? Number(data.horometro) : null
    },
    filters: [{ column: 'id', op: 'eq', value: data.id }]
  });
}

// ── OT (Órdenes de Trabajo / Mantenimiento) ──────────────────

async function getHistorialOT() {
  return dbQuery({
    action: 'select',
    table: 'tblGamasOT',
    options: { select: '*', order: 'fecha.desc', limit: 200 }
  });
}

async function doPostOT(data) {
  return dbQuery({
    action: 'insert',
    table: 'tblGamasOT',
    data: {
      activo:   data.activo,
      fecha:    data.fecha,
      operario: data.operario,
      tiempo:   data.tiempo,
      texto:    data.texto,
      estado:   data.estado,
      gama:     data.gama,
      medicion: data.medicion,
      checks:   data.checks
    },
    options: { select: 'id' }
  });
}

async function doEditarOT(data) {
  const updates = {};
  if (data.fecha    !== undefined) updates.fecha    = data.fecha;
  if (data.operario !== undefined) updates.operario = data.operario;
  if (data.texto    !== undefined) updates.texto    = data.texto;
  if (data.medicion !== undefined) updates.medicion = data.medicion;
  if (data.estado   !== undefined) updates.estado   = data.estado;
  return dbQuery({
    action: 'update',
    table: 'tblGamasOT',
    data: updates,
    filters: [{ column: 'id', op: 'eq', value: data.id }]
  });
}

// ── AUSENCIAS ───────────────────────────────────────────────

async function getAusencias() {
  return dbQuery({ action: 'select', table: 'tblAusencias', options: { select: '*' } });
}

async function doPostAusencia(data) {
  return dbQuery({
    action: 'insert',
    table: 'tblAusencias',
    data: {
      tipo:          data.categoria,
      trabajador:    data.trabajador,
      start:         data.start,
      end:           data.end,
      dias:          data.dias,
      subtipo:       data.subtipo,
      horas:         data.horas,
      motivo:        data.motivo,
      fechaCreacion: new Date().toISOString()
    },
    options: { select: 'id' }
  });
}

async function doEditAusencia(data) {
  const updates = {};
  if (data.start   !== undefined) updates.start   = data.start;
  if (data.end     !== undefined) updates.end     = data.end;
  if (data.dias    !== undefined) updates.dias    = data.dias;
  if (data.subtipo !== undefined) updates.subtipo = data.subtipo;
  if (data.horas   !== undefined) updates.horas   = data.horas;
  if (data.motivo  !== undefined) updates.motivo  = data.motivo;
  return dbQuery({
    action: 'update',
    table: 'tblAusencias',
    data: updates,
    filters: [{ column: 'id', op: 'eq', value: data.id }]
  });
}

async function doDeleteAusencia(data) {
  return dbQuery({
    action: 'delete',
    table: 'tblAusencias',
    filters: [{ column: 'id', op: 'eq', value: data.id }]
  });
}

// ── DOCUMENTOS ───────────────────────────────────────────────

async function getDocumentos() {
  const result = await dbQuery({
    action: 'select',
    table: 'tblControlDocumental',
    options: { select: '*' }
  });
  if (!result.ok) return result;

  const hoy = new Date(); hoy.setHours(0, 0, 0, 0);
  const docs = (result.data || []).map(r => {
    const fv   = r.fechavencimiento ? new Date(r.fechavencimiento) : null;
    const dias = fv ? Math.round((fv - hoy) / 86400000) : null;
    return {
      fuente:        'nuevo',
      id:            r.id,
      numero:        r.numero,
      nombre:        r.nombre,
      estado:        r.estado,
      organo:        r.organo,
      fechaInicio:   r.fechainicio,
      fechaVig:      r.fechavencimiento,
      creado:        r.creado,
      idDocumento:   r.iddocumento,
      tipoDocumento: r.tipodocumento,
      diasRestantes: dias,
      tiempoAviso:   30
    };
  });
  return { ok: true, data: docs };
}

// ── CARGA INICIAL ────────────────────────────────────────────

async function getInit() {
  try {
    const [f, a, c, g] = await Promise.all([
      getFichajes(),
      getAusencias(),
      getCamiones(),
      getGasoil()
    ]);
    return {
      ok:          true,
      fichajes:    f.data || [],
      ausencias:   a.data || [],
      camiones:    c.data || [],
      gasoilData:  g.data || [],
      gasoilDep1:  g.dep1 || 0,
      gasoilDep2:  g.dep2 || 0
    };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

// ── GAMAS NORMAS ─────────────────────────────────────────────

async function getGamasNormas() {
  return dbQuery({ action: 'select', table: 'tblGamasNormas', options: { select: '*', order: 'id.asc' } });
}
async function doPostGamaNorma(d) {
  return dbQuery({ action: 'insert', table: 'tblGamasNormas', data: { nombre: d.nombre, modelo: d.modelo, intervalo: Number(d.intervalo), unidad: d.unidad || 'H', checks: d.checks || [] } });
}
async function doEditGamaNorma(d) {
  return dbQuery({ action: 'update', table: 'tblGamasNormas', data: { nombre: d.nombre, modelo: d.modelo, intervalo: Number(d.intervalo), unidad: d.unidad || 'H', checks: d.checks || [] }, filters: [{ column: 'id', op: 'eq', value: d.id }] });
}
async function doDeleteGamaNorma(id) {
  return dbQuery({ action: 'delete', table: 'tblGamasNormas', filters: [{ column: 'id', op: 'eq', value: id }] });
}

// ── GAMAS DEPENDIENTES ───────────────────────────────────────

async function getGamasDependientes() {
  return dbQuery({ action: 'select', table: 'tblGamasDependientes', options: { select: '*', order: 'id.asc' } });
}
async function doPostGamaDependiente(d) {
  return dbQuery({ action: 'insert', table: 'tblGamasDependientes', data: { normaId: d.normaId, nombre: d.nombre, checks: d.checks || [] } });
}
async function doEditGamaDependiente(d) {
  return dbQuery({ action: 'update', table: 'tblGamasDependientes', data: { normaId: d.normaId, nombre: d.nombre, checks: d.checks || [] }, filters: [{ column: 'id', op: 'eq', value: d.id }] });
}
async function doDeleteGamaDependiente(id) {
  return dbQuery({ action: 'delete', table: 'tblGamasDependientes', filters: [{ column: 'id', op: 'eq', value: id }] });
}

// ── GAMAS ACTIVOS ────────────────────────────────────────────

async function getGamasActivos() {
  return dbQuery({ action: 'select', table: 'tblGamasActivos', options: { select: '*', order: 'activo.asc' } });
}
async function doPostGamaActivo(d) {
  return dbQuery({ action: 'insert', table: 'tblGamasActivos', data: { activo: d.activo, modelo: d.modelo, codigogama: d.codigogama } });
}
async function doEditGamaActivo(d) {
  return dbQuery({ action: 'update', table: 'tblGamasActivos', data: { activo: d.activo, modelo: d.modelo, codigogama: d.codigogama }, filters: [{ column: 'id', op: 'eq', value: d.id }] });
}
async function doDeleteGamaActivo(id) {
  return dbQuery({ action: 'delete', table: 'tblGamasActivos', filters: [{ column: 'id', op: 'eq', value: id }] });
}

// ── GAMAS LISTADO PREVENTIVO ─────────────────────────────────

async function getGamasListado() {
  return dbQuery({ action: 'select', table: 'tblGamasListadoPreventivo', options: { select: '*', order: 'activo.asc' } });
}
async function doPostGamaListado(d) {
  const proximo = Number(d.proximo) || 0;
  const ultima  = Number(d.ultima)  || 0;
  return dbQuery({
    action: 'insert',
    table: 'tblGamasListadoPreventivo',
    data: {
      activo: d.activo, codigogama: d.codigogama, medidor: d.medidor || 'H',
      proximo, ultima, ultimafecha: d.ultimafecha || null, falta: proximo - ultima
    },
    options: { select: 'id' }
  });
}
async function doEditGamaListado(d) {
  const proximo = Number(d.proximo) || 0;
  const ultima  = Number(d.ultima)  || 0;
  return dbQuery({
    action: 'update',
    table: 'tblGamasListadoPreventivo',
    data: {
      activo: d.activo, codigogama: d.codigogama, medidor: d.medidor || 'H',
      proximo, ultima, ultimafecha: d.ultimafecha || null, falta: proximo - ultima
    },
    filters: [{ column: 'id', op: 'eq', value: d.id }]
  });
}
async function doDeleteGamaListado(id) {
  return dbQuery({ action: 'delete', table: 'tblGamasListadoPreventivo', filters: [{ column: 'id', op: 'eq', value: id }] });
}

// ── ROUTER: apiFetch → proxy ────────────────────────────────
// apiFetch y apiPost definidos en app.js
