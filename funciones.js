// ============================================================
// ARIFOMA · FUNCIONES SUPABASE — versión limpia
// ============================================================

// ── BUSINESS CENTRAL CONFIG ───────────────────────────────────
const BC_TENANT_F  = '5bd828f2-1899-48ba-a269-c37733f41806';
const BC_ENV_F     = 'Production';
const BC_COMPANY_F = 'ARIFOMA 25P.V06';

async function enviarLineaBCPesada(data) {
  if (typeof getBCToken !== 'function') return;
  let token;
  try { token = await getBCToken(); } catch(e) { console.warn('BC token:', e.message); return; }
  const base = `https://api.businesscentral.dynamics.com/v2.0/${BC_TENANT_F}/${BC_ENV_F}/api/v2.0/companies`;
  const headers = { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' };

  const cJson = await (await fetch(base, { headers })).json();
  const company = cJson.value.find(c => c.name === BC_COMPANY_F);
  if (!company) { console.warn('BC: Company no encontrada'); return; }
  const companyId = company.id;

  const filter = `customerNumber eq '${data.codigoCliente}' and externalDocumentNumber eq '${data.proyectoCod}'`;
  const ordersJson = await (await fetch(`${base}(${companyId})/salesOrders?$filter=${encodeURIComponent(filter)}&$select=id,number`, { headers })).json();
  let orderId;

  if (ordersJson.value && ordersJson.value.length > 0) {
    orderId = ordersJson.value[0].id;
  } else {
    const newOrder = await fetch(`${base}(${companyId})/salesOrders`, {
      method: 'POST', headers,
      body: JSON.stringify({ customerNumber: data.codigoCliente, externalDocumentNumber: data.proyectoCod })
    });
    if (!newOrder.ok) { console.warn('BC crear pedido:', await newOrder.text()); return; }
    orderId = (await newOrder.json()).id;
  }

  const lineBody = {
    lineType: 'Item',
    lineObjectNumber: data.productoCod,
    description: `${data.productoNombre} | ${data.proyectoName||data.proyectoCod} | ${(Number(data.pesoNeto)/1000).toFixed(3)} Tn | ${data.matriculacam}`,
    quantity: parseFloat((Number(data.pesoNeto) / 1000).toFixed(3))
  };
  const lineRes = await fetch(`${base}(${companyId})/salesOrders(${orderId})/salesOrderLines`, {
    method: 'POST', headers,
    body: JSON.stringify(lineBody)
  });
  if (!lineRes.ok) {
    const errText = await lineRes.text();
    console.warn('BC línea 400 detalle:', errText);
  } else console.log('BC línea creada OK');
}

const SUPABASE_URL = 'https://bnsfgzjqmibsrklllqxb.supabase.co';
const SUPABASE_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImJuc2ZnempxbWlic3JrbGxscXhiIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzQzNzYwNzksImV4cCI6MjA4OTk1MjA3OX0.8mTQHPdO954ICBd1Xam-kKmcA69CMyO2v3x1liFgWyk';
let _supabase = supabase.createClient(SUPABASE_URL, SUPABASE_KEY);
let _sessionToken = null;

// Recrear cliente Supabase con token de sesión en headers
function _initAuthClient(token) {
  _sessionToken = token;
  _supabase = supabase.createClient(SUPABASE_URL, SUPABASE_KEY, {
    global: { headers: { 'x-session-token': token } }
  });
}

// ── AUTENTICACIÓN SEGURA ────────────────────────────────────

async function obtenerUsuarios() {
  const { data, error } = await _supabase.rpc('obtener_usuarios');
  return error ? { ok: false, error: error.message } : { ok: true, data: data || [] };
}

async function verificarPin(nombre, pin) {
  const { data, error } = await _supabase.rpc('verificar_pin', {
    p_nombre: nombre,
    p_pin: pin
  });
  if (error) return { ok: false, error: error.message };
  if (data && data.ok) {
    _initAuthClient(data.token);
  }
  return data;
}

async function cerrarSesion() {
  if (_sessionToken) {
    await _supabase.rpc('cerrar_sesion', { p_token: _sessionToken });
  }
  _sessionToken = null;
  _supabase = supabase.createClient(SUPABASE_URL, SUPABASE_KEY);
}

// ── FICHAJES ─────────────────────────────────────────────────

async function getFichajes() {
  const { data, error } = await _supabase
    .from('tblFichaje')
    .select('*')
    .order('fentrada', { ascending: false })
    .limit(300);
  return error ? { ok: false, error: error.message } : { ok: true, data };
}

async function doPostEntrada(data) {
  const { error } = await _supabase.from('tblFichaje').insert([{
    empleado:    data.empleado,
    fecha:       data.fecha    || new Date().toLocaleDateString('es-ES'),
    entrada:     data.entrada  || new Date().toLocaleTimeString('es-ES', { hour: '2-digit', minute: '2-digit' }),
    tipoTrabajo: data.tipoTrabajo || 'JORNADA',
    fentrada:    data.fentrada || new Date().toISOString()
  }]);
  return error ? { ok: false, error: error.message } : { ok: true };
}

async function doEditSalida(data) {
  // Buscar la última entrada abierta (sin salida) del empleado
  const { data: registro, error: searchError } = await _supabase
    .from('tblFichaje')
    .select('id')
    .eq('empleado', data.empleado)
    .is('salida', null)
    .order('fentrada', { ascending: false })
    .limit(1)
    .single();

  if (searchError || !registro) {
    // Fallback: crear registro nuevo si no existe entrada abierta
    const { error: insErr } = await _supabase.from('tblFichaje').insert([{
      empleado:  data.empleado,
      salida:    data.salida,
      fsalida:   data.fsalida || new Date().toISOString(),
      tiempodia: data.tiempodia
    }]);
    return insErr ? { ok: false, error: insErr.message } : { ok: true };
  }

  const { error: updateErr } = await _supabase
    .from('tblFichaje')
    .update({
      salida:    data.salida,
      fsalida:   data.fsalida || new Date().toISOString(),
      tiempodia: data.tiempodia
    })
    .eq('id', registro.id);

  return updateErr ? { ok: false, error: updateErr.message } : { ok: true };
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
  const { error } = await _supabase.from('tblFichaje').update(updates).eq('id', data.id);
  return error ? { ok: false, error: error.message } : { ok: true };
}

async function doDeleteFichaje(data) {
  const { error } = await _supabase.from('tblFichaje').delete().eq('id', data.id);
  return error ? { ok: false, error: error.message } : { ok: true };
}

// ── PEDIDOS (Pesajes) ────────────────────────────────────────

async function getPedidos(diasAtras = 90) {
  const corte = new Date();
  corte.setDate(corte.getDate() - diasAtras);
  const { data, error } = await _supabase
    .from('tblpedidos')
    .select('*')
    .gte('fechaHora', corte.toISOString())
    .order('fechaHora', { ascending: false });
  return error ? { ok: false, error: error.message } : { ok: true, data };
}

async function doPostPesada(data) {
  const { data: inserted, error } = await _supabase.from('tblpedidos').insert([{
    matriculacam:  data.matriculacam,
    matricularem:  data.matricularem,
    tara:          data.tara,
    chofer:        data.chofer,
    nombreCliente: data.nombreCliente,
    codigoCliente: data.codigoCliente,
    productoNombre:data.productoNombre,
    productoCod:   data.productoCod,
    pesoBruto:     data.pesoBruto,
    pesoNeto:      data.pesoNeto,
    proyectoName:  data.proyectoName,
    proyectoCod:   data.proyectoCod,
    numPedido:     data.numPedido,
    numLinea:      data.numLinea,
    fechaHora:     new Date().toISOString()
  }]).select('id').single();
  if (error) return { ok: false, error: error.message };
  if (typeof enviarLineaBCPesada === 'function') {
    enviarLineaBCPesada(data).catch(e => console.warn('BC línea:', e.message));
  }
  return { ok: true, id: inserted?.id };
}

async function doDeletePedido(data) {
  const id=Number(data.id);
  if(!id||isNaN(id))return{ok:false,error:'ID inválido'};
  const { error } = await _supabase.from('tblpedidos').delete().eq('id', id);
  return error ? { ok: false, error: error.message } : { ok: true };
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
  const { error } = await _supabase.from('tblpedidos').update(updates).eq('id', data.id);
  return error ? { ok: false, error: error.message } : { ok: true };
}

// ── CAMIONES ─────────────────────────────────────────────────

async function getCamiones() {
  const { data, error } = await _supabase.from('tblcamiones').select('*').order('matriculacam');
  return error ? { ok: false, error: error.message } : { ok: true, data };
}

async function doNuevoCamion(data) {
  const { tipo, ...campos } = data; // quitar el campo "tipo" del payload
  const { error } = await _supabase.from('tblcamiones').insert([campos]);
  return error ? { ok: false, error: error.message } : { ok: true };
}

async function doEditarCamion(data) {
  const { id, tipo, ...updates } = data; // quitar "tipo" y "id" del update
  const { error } = await _supabase.from('tblcamiones').update(updates).eq('id', id);
  return error ? { ok: false, error: error.message } : { ok: true };
}

// ── PRODUCCIÓN ───────────────────────────────────────────────

async function getProduccion(mes, anyo) {
  let query = _supabase.from('PRODUCCION').select('*');
  if (anyo) {
    if (mes) {
      const start = `${anyo}-${String(mes).padStart(2, '0')}-01`;
      const end   = `${anyo}-${String(mes).padStart(2, '0')}-31`;
      query = query.gte('fecha', start).lte('fecha', end);
    } else {
      query = query.gte('fecha', `${anyo}-01-01`).lte('fecha', `${anyo}-12-31`);
    }
  }
  const { data, error } = await query.order('fecha', { ascending: true });
  return error ? { ok: false, error: error.message } : { ok: true, data };
}

async function doEditProduccion(data) {
  const { id, ...campos } = data;
  const { error } = await _supabase.from('PRODUCCION').update({
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
  }).eq('id', id);
  return error ? { ok: false, error: error.message } : { ok: true };
}

async function doAddProduccion(data) {
  const { error } = await _supabase.from('PRODUCCION').insert([{
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
  }]);
  return error ? { ok: false, error: error.message } : { ok: true };
}

// ── GASOIL ───────────────────────────────────────────────────

async function getGasoil() {
  const [gasoilRes, horoRes] = await Promise.all([
    _supabase.from('GASOIL').select('*').order('fecha', { ascending: false }),
    _supabase.from('horometros').select('*')
  ]);
  if (gasoilRes.error) return { ok: false, error: gasoilRes.error.message };

  const data = gasoilRes.data || [];
  // Stock
  let dep1 = 0, dep2 = 0;
  data.forEach(m => {
    const litros = Number(m.litros) || 0;
    if (m.destino === 'DEP1') dep1 += litros;
    if (m.origen  === 'DEP1') dep1 -= litros;
    if (m.destino === 'DEP2') dep2 += litros;
    if (m.origen  === 'DEP2') dep2 -= litros;
  });

  // Horómetros desde tabla horometros (sincronizada desde Sheet)
  const consumos = (horoRes.data || []).map(r => ({
    activo: r.activo,
    max:    r.horometro
  }));

  return { ok: true, data, dep1, dep2, consumos };
}

async function doPostGasoil(data) {
  const { error } = await _supabase.from('GASOIL').insert([{
    fecha:     data.fecha,
    proveedor: data.proveedor,
    origen:    data.origen,
    destino:   data.destino,
    tipo:      data.tipoMovimiento,
    litros:    Number(data.litros)
  }]);
  return error ? { ok: false, error: error.message } : { ok: true };
}

async function doEditarGasoil(data) {
  const { error } = await _supabase.from('GASOIL').update({
    fecha:     data.fecha,
    proveedor: data.proveedor,
    origen:    data.origen,
    destino:   data.destino,
    tipo:      data.tipoMovimiento,
    litros:    Number(data.litros)
  }).eq('id', data.id);
  return error ? { ok: false, error: error.message } : { ok: true };
}

// ── OT (Órdenes de Trabajo / Mantenimiento) ──────────────────

async function getHistorialOT() {
  const { data, error } = await _supabase
    .from('tblGamasOT')
    .select('*')
    .order('fecha', { ascending: false })
    .limit(200);
  return error ? { ok: false, error: error.message } : { ok: true, data };
}

async function doPostOT(data) {
  const { data: inserted, error } = await _supabase.from('tblGamasOT').insert([{
    activo:   data.activo,
    fecha:    data.fecha,
    operario: data.operario,
    tiempo:   data.tiempo,
    texto:    data.texto,
    estado:   data.estado,
    gama:     data.gama,
    medicion: data.medicion,
    checks:   data.checks   // columna JSONB en Supabase
  }]).select('id').single();
  return error ? { ok: false, error: error.message } : { ok: true, ot: inserted?.id };
}

async function doEditarOT(data) {
  const updates = {};
  if (data.fecha    !== undefined) updates.fecha    = data.fecha;
  if (data.operario !== undefined) updates.operario = data.operario;
  if (data.texto    !== undefined) updates.texto    = data.texto;
  if (data.medicion !== undefined) updates.medicion = data.medicion;
  if (data.estado   !== undefined) updates.estado   = data.estado;
  const { error } = await _supabase.from('tblGamasOT').update(updates).eq('id', data.id);
  return error ? { ok: false, error: error.message } : { ok: true };
}

// ── AUSENCIAS (vacaciones, bajas, días libres, extras) ───────

async function getAusencias() {
  const { data, error } = await _supabase.from('tblAusencias').select('*');
  return error ? { ok: false, error: error.message } : { ok: true, data };
}

async function doPostAusencia(data) {
  const { data: inserted, error } = await _supabase.from('tblAusencias').insert([{
    tipo:          data.categoria,
    trabajador:    data.trabajador,
    start:         data.start,
    end:           data.end,
    dias:          data.dias,
    subtipo:       data.subtipo,
    horas:         data.horas,
    motivo:        data.motivo,
    fechaCreacion: new Date().toISOString()
  }]).select('id').single();
  return error ? { ok: false, error: error.message } : { ok: true, id: inserted?.id };
}

async function doEditAusencia(data) {
  const updates = {};
  if (data.start   !== undefined) updates.start   = data.start;
  if (data.end     !== undefined) updates.end     = data.end;
  if (data.dias    !== undefined) updates.dias    = data.dias;
  if (data.subtipo !== undefined) updates.subtipo = data.subtipo;
  if (data.horas   !== undefined) updates.horas   = data.horas;
  if (data.motivo  !== undefined) updates.motivo  = data.motivo;
  const { error } = await _supabase.from('tblAusencias').update(updates).eq('id', data.id);
  return error ? { ok: false, error: error.message } : { ok: true };
}

async function doDeleteAusencia(data) {
  const { error } = await _supabase.from('tblAusencias').delete().eq('id', data.id);
  return error ? { ok: false, error: error.message } : { ok: true };
}

// ── DOCUMENTOS ───────────────────────────────────────────────

async function getDocumentos() {
  const { data, error } = await _supabase
    .from('tblcontroldocumental')
    .select('*');
  if (error) return { ok: false, error: error.message };

  const hoy = new Date(); hoy.setHours(0, 0, 0, 0);
  const docs = (data || []).map(r => {
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
  const { data, error } = await _supabase.from('tblGamasNormas').select('*').order('id', { ascending: true });
  return error ? { ok: false, error: error.message } : { ok: true, data: data || [] };
}
async function doPostGamaNorma(d) {
  const { error } = await _supabase.from('tblGamasNormas').insert([{ nombre: d.nombre, modelo: d.modelo, intervalo: Number(d.intervalo), unidad: d.unidad || 'H', checks: d.checks || [] }]);
  return error ? { ok: false, error: error.message } : { ok: true };
}
async function doEditGamaNorma(d) {
  const { error } = await _supabase.from('tblGamasNormas').update({ nombre: d.nombre, modelo: d.modelo, intervalo: Number(d.intervalo), unidad: d.unidad || 'H', checks: d.checks || [] }).eq('id', d.id);
  return error ? { ok: false, error: error.message } : { ok: true };
}
async function doDeleteGamaNorma(id) {
  const { error } = await _supabase.from('tblGamasNormas').delete().eq('id', id);
  return error ? { ok: false, error: error.message } : { ok: true };
}

// ── GAMAS DEPENDIENTES (Subgamas) ─────────────────────────────

async function getGamasDependientes() {
  const { data, error } = await _supabase.from('tblGamasDependientes').select('*').order('id', { ascending: true });
  return error ? { ok: false, error: error.message } : { ok: true, data: data || [] };
}
async function doPostGamaDependiente(d) {
  const { error } = await _supabase.from('tblGamasDependientes').insert([{ normaId: d.normaId, nombre: d.nombre, checks: d.checks || [] }]);
  return error ? { ok: false, error: error.message } : { ok: true };
}
async function doEditGamaDependiente(d) {
  const { error } = await _supabase.from('tblGamasDependientes').update({ normaId: d.normaId, nombre: d.nombre, checks: d.checks || [] }).eq('id', d.id);
  return error ? { ok: false, error: error.message } : { ok: true };
}
async function doDeleteGamaDependiente(id) {
  const { error } = await _supabase.from('tblGamasDependientes').delete().eq('id', id);
  return error ? { ok: false, error: error.message } : { ok: true };
}

// ── GAMAS ACTIVOS ─────────────────────────────────────────────

async function getGamasActivos() {
  const { data, error } = await _supabase.from('tblGamasActivos').select('*').order('activo', { ascending: true });
  return error ? { ok: false, error: error.message } : { ok: true, data: data || [] };
}
async function doPostGamaActivo(d) {
  const { error } = await _supabase.from('tblGamasActivos').insert([{ activo: d.activo, modelo: d.modelo, codigogama: d.codigogama }]);
  return error ? { ok: false, error: error.message } : { ok: true };
}
async function doEditGamaActivo(d) {
  const { error } = await _supabase.from('tblGamasActivos').update({ activo: d.activo, modelo: d.modelo, codigogama: d.codigogama }).eq('id', d.id);
  return error ? { ok: false, error: error.message } : { ok: true };
}
async function doDeleteGamaActivo(id) {
  const { error } = await _supabase.from('tblGamasActivos').delete().eq('id', id);
  return error ? { ok: false, error: error.message } : { ok: true };
}

// ── GAMAS LISTADO PREVENTIVO ──────────────────────────────────

async function getGamasListado() {
  const { data, error } = await _supabase.from('tblGamasListadoPreventivo').select('*').order('activo', { ascending: true });
  return error ? { ok: false, error: error.message } : { ok: true, data: data || [] };
}
async function doPostGamaListado(d) {
  const proximo = Number(d.proximo) || 0;
  const ultima  = Number(d.ultima)  || 0;
  const { data: ins, error } = await _supabase.from('tblGamasListadoPreventivo').insert([{
    activo: d.activo, codigogama: d.codigogama, medidor: d.medidor || 'H',
    proximo, ultima, ultimafecha: d.ultimafecha || null, falta: proximo - ultima
  }]).select('id').single();
  return error ? { ok: false, error: error.message } : { ok: true, id: ins?.id };
}
async function doEditGamaListado(d) {
  const proximo = Number(d.proximo) || 0;
  const ultima  = Number(d.ultima)  || 0;
  const { error } = await _supabase.from('tblGamasListadoPreventivo').update({
    activo: d.activo, codigogama: d.codigogama, medidor: d.medidor || 'H',
    proximo, ultima, ultimafecha: d.ultimafecha || null, falta: proximo - ultima
  }).eq('id', d.id);
  return error ? { ok: false, error: error.message } : { ok: true };
}
async function doDeleteGamaListado(id) {
  const { error } = await _supabase.from('tblGamasListadoPreventivo').delete().eq('id', id);
  return error ? { ok: false, error: error.message } : { ok: true };
}

// ── ROUTER: apiFetch → Supabase ──────────────────────────────

// apiFetch y apiPost definidos en index.html