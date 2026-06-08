
// ============================================================
// ARIFOMA · CAPA DE API — SUPABASE
// ============================================================
// _supabase (solo auth), dbQuery (proxy backend), _initAuthClient definidos en funciones.js

// ── GOOGLE SHEETS (para Producción y Gasoil) ────────────────
const SHEETS_API = 'https://script.google.com/macros/s/AKfycbwPIIgZCg03i4aJN8HIxKf20P5IPc-j3HOkoHmt2Jx0-vqiWrmq4Gz2WZmZvyopYJlv/exec';

function sheetsFetch(params) {
  return new Promise((resolve, reject) => {
    const cb = '_cb' + Date.now();
    const script = document.createElement('script');
    const timeout = setTimeout(() => {
      reject(new Error('Timeout'));
      delete window[cb];
      if (script.parentNode) script.parentNode.removeChild(script);
    }, 15000);
    window[cb] = function(data) {
      clearTimeout(timeout);
      resolve(data);
      delete window[cb];
      if (script.parentNode) script.parentNode.removeChild(script);
    };
    script.onerror = () => { clearTimeout(timeout); reject(new Error('Script error')); };
    script.src = SHEETS_API + params + '&callback=' + cb;
    document.head.appendChild(script);
  });
}

async function sheetsPost(payload) {
  try {
    const res = await fetch(SHEETS_API, {
      method: 'POST',
      mode: 'cors',
      redirect: 'follow',
      headers: { 'Content-Type': 'text/plain' },
      body: JSON.stringify(payload)
    });
    if (!res.ok) return { ok: false, error: 'HTTP ' + res.status };
    const json = await res.json();
    return json;
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ── FICHAJES ────────────────────────────────────────────────
async function getFichajes() {
  return dbQuery({ action: 'select', table: 'tblFichaje', options: { select: '*', order: 'fentrada.desc', limit: 1500 } });
}
async function doPostEntrada(data) {
  const result = await dbQuery({ action: 'insert', table: 'tblFichaje', data: {
    empleado: data.empleado,
    fecha: data.fecha || new Date().toISOString().slice(0, 10),
    entrada: data.entrada || new Date().toLocaleTimeString('es-ES', { hour: '2-digit', minute: '2-digit' }),
    tipoTrabajo: data.tipoTrabajo || 'JORNADA',
    fentrada: data.fentrada || new Date().toISOString()
  }});
  if (!result.ok) console.error('doPostEntrada error:', result.error);
  return result;
}
async function doPostFichajeManual(data) {
  const result = await dbQuery({ action: 'insert', table: 'tblFichaje', data: {
    empleado: data.empleado, fecha: data.fecha, entrada: data.entrada,
    salida: data.salida || null, tipoTrabajo: data.tipoTrabajo || 'JORNADA',
    fentrada: data.fentrada, fsalida: data.fsalida || null, tiempodia: data.tiempodia || null
  }, options: { select: 'id' }});
  if (!result.ok) { console.error('doPostFichajeManual error:', result.error); return result; }
  return { ok: true, id: result.data && result.data[0] ? result.data[0].id : null };
}
async function doEditSalida(data) {
  const search = await dbQuery({ action: 'select', table: 'tblFichaje',
    filters: [{ column: 'empleado', op: 'eq', value: data.empleado }, { column: 'salida', op: 'is', value: 'null' }],
    options: { select: 'id', order: 'fentrada.desc', limit: 1 }
  });
  const registro = search.ok && search.data && search.data.length ? search.data[0] : null;
  if (!registro) {
    return dbQuery({ action: 'insert', table: 'tblFichaje', data: {
      empleado: data.empleado, salida: data.salida,
      fsalida: data.fsalida || new Date().toISOString(), tiempodia: data.tiempodia
    }});
  }
  return dbQuery({ action: 'update', table: 'tblFichaje',
    data: { salida: data.salida, fsalida: data.fsalida || new Date().toISOString(), tiempodia: data.tiempodia },
    filters: [{ column: 'id', op: 'eq', value: registro.id }]
  });
}
async function doEditFichaje(data) {
  return dbQuery({ action: 'update', table: 'tblFichaje', data: {
    empleado: data.empleado, fecha: data.fecha, entrada: data.entrada,
    salida: data.salida, fentrada: data.fentrada, fsalida: data.fsalida, tiempodia: data.tiempodia
  }, filters: [{ column: 'id', op: 'eq', value: data.id }]});
}
async function doDeleteFichaje(data) {
  return dbQuery({ action: 'delete', table: 'tblFichaje', filters: [{ column: 'id', op: 'eq', value: data.id }] });
}

// ── PEDIDOS ─────────────────────────────────────────────────
async function getPedidos(diasAtras = 90) {
  const corte = new Date();
  corte.setDate(corte.getDate() - diasAtras);
  return dbQuery({ action: 'select', table: 'tblpedidos',
    filters: [{ column: 'fechaHora', op: 'gte', value: corte.toISOString() }],
    options: { select: '*', order: 'fechaHora.desc' }
  });
}
async function doPostPesada(data) {
  const result = await dbQuery({ action: 'insert', table: 'tblpedidos', data: {
    matriculacam: data.matriculacam, matricularem: data.matricularem,
    tara: data.tara, chofer: data.chofer,
    nombreCliente: data.nombreCliente, codigoCliente: data.codigoCliente,
    productoNombre: data.productoNombre, productoCod: data.productoCod,
    pesoBruto: data.pesoBruto, pesoNeto: data.pesoNeto,
    proyectoName: data.proyectoName, proyectoCod: data.proyectoCod,
    numPedido: data.numPedido, numLinea: data.numLinea,
    idproyecto: data.proyectoCod || null,
    numalbarancalle: (data.numPedido && data.numLinea) ? `${data.numPedido}/${data.numLinea}` : null,
    fechaHora: new Date().toISOString()
  }, options: { select: 'id' }});
  if (!result.ok) return result;
  const pedidoId = result.data && result.data[0] ? result.data[0].id : null;
  enviarLineaBCPesada(data).catch(e => console.warn('BC línea:', e.message));
  return { ok: true, id: pedidoId };
}

async function enviarLineaBCPesada(data) {
  const token = await getBCToken();
  const res = await fetch('/api/bc/linea-pesada', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      token,
      codigoCliente: data.codigoCliente,
      proyectoCod: data.proyectoCod,
      productoCod: data.productoCod,
      productoNombre: data.productoNombre,
      pesoNeto: data.pesoNeto,
      matriculacam: data.matriculacam,
      proyectoName: data.proyectoName
    })
  });
  if (!res.ok) throw new Error(await res.text());
  const json = await res.json();
  if (!json.ok) throw new Error(json.error);
  return json.numalbarancalle;
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
  if (data.observaciones !== undefined) updates.observaciones = data.observaciones;
  return dbQuery({ action: 'update', table: 'tblpedidos', data: updates, filters: [{ column: 'id', op: 'eq', value: data.id }] });
}

// ── CAMIONES ────────────────────────────────────────────────
async function getCamiones() {
  return dbQuery({ action: 'select', table: 'tblcamiones', options: { select: '*', order: 'matriculacam.asc' } });
}
async function doNuevoCamion(data) {
  const { tipo, id, ...payload } = data;
  return dbQuery({ action: 'insert', table: 'tblcamiones', data: payload });
}
async function doEditarCamion(data) {
  const { id, tipo, ...updates } = data;
  return dbQuery({ action: 'update', table: 'tblcamiones', data: updates, filters: [{ column: 'id', op: 'eq', value: id }] });
}
async function doEliminarCamion(data) {
  const id = Number(data.id);
  if (!id || isNaN(id)) return { ok: false, error: 'ID inválido' };
  return dbQuery({ action: 'delete', table: 'tblcamiones', filters: [{ column: 'id', op: 'eq', value: id }] });
}

// ── PRODUCCIÓN y GASOIL → Google Sheets (via sheetsFetch/sheetsPost) ──

// ── OT (Órdenes de Trabajo) ──────────────────────────────────
async function doPostOT(data) {
  // Calcular siguiente número OT
  const maxRes = await dbQuery({ action: 'select', table: 'tblGamasOT', options: { select: 'Ot', order: 'Ot.desc', limit: 1 } });
  const maxRow = maxRes.ok ? maxRes.data : [];
  const nextOt = (maxRow && maxRow.length && maxRow[0].Ot ? maxRow[0].Ot : 0) + 1;
  const row = {
    Activo: data.activo, Fecha: data.fecha, Operario: data.operario,
    Tiempo: data.tiempo, Texto: data.texto, Estado: data.estado,
    Gama: data.gama, Medicion: data.medicion, Ot: nextOt
  };
  if (data.checks && Array.isArray(data.checks)) {
    data.checks.forEach((v, i) => { row['n' + (i + 1)] = !!v; });
  }
  const result = await dbQuery({ action: 'insert', table: 'tblGamasOT', data: row, options: { select: 'id,Ot' } });
  if (!result.ok) return result;
  const inserted = result.data && result.data[0] ? result.data[0] : {};
  return { ok: true, ot: inserted.Ot || inserted.id };
}
async function getHistorialOT() {
  const result = await dbQuery({ action: 'select', table: 'tblGamasOT', options: { select: '*', order: 'Fecha.desc', limit: 200 } });
  if (!result.ok) return result;
  const mapped = (result.data || []).map(r => {
    const checks = [];
    for (let i = 1; i <= 60; i++) { if (r['n' + i] !== undefined) checks.push(!!r['n' + i]); }
    return {
      id: r.id, activo: r.Activo, fecha: r.Fecha, operario: r.Operario,
      tiempo: r.Tiempo, texto: r.Texto, estado: r.Estado,
      gama: r.Gama, medicion: r.Medicion, ot: r.Ot, checks
    };
  });
  return { ok: true, data: mapped };
}
async function doEditarOT(data) {
  const { id, ...rest } = data;
  const updates = {};
  if (rest.activo !== undefined) updates.Activo = rest.activo;
  if (rest.fecha !== undefined) updates.Fecha = rest.fecha;
  if (rest.operario !== undefined) updates.Operario = rest.operario;
  if (rest.tiempo !== undefined) updates.Tiempo = rest.tiempo;
  if (rest.texto !== undefined) updates.Texto = rest.texto;
  if (rest.estado !== undefined) updates.Estado = rest.estado;
  if (rest.gama !== undefined) updates.Gama = rest.gama;
  if (rest.medicion !== undefined) updates.Medicion = rest.medicion;
  if (rest.checks && Array.isArray(rest.checks)) {
    rest.checks.forEach((v, i) => { updates['n' + (i + 1)] = !!v; });
  }
  return dbQuery({ action: 'update', table: 'tblGamasOT', data: updates, filters: [{ column: 'id', op: 'eq', value: id }] });
}
async function doDeleteOT(data) {
  return dbQuery({ action: 'delete', table: 'tblGamasOT', filters: [{ column: 'id', op: 'eq', value: data.id }] });
}

// ── AUSENCIAS ────────────────────────────────────────────────
async function getAusencias() {
  return dbQuery({ action: 'select', table: 'tblAusencias', options: { select: '*' } });
}
async function doPostAusencia(data) {
  const result = await dbQuery({ action: 'insert', table: 'tblAusencias', data: {
    tipo: data.categoria, trabajador: data.trabajador,
    start: data.start, end: data.end, dias: data.dias,
    subtipo: data.subtipo, horas: data.horas, motivo: data.motivo,
    fechaCreacion: new Date().toISOString()
  }, options: { select: 'id' }});
  if (!result.ok) return result;
  return { ok: true, id: result.data && result.data[0] ? result.data[0].id : null };
}
async function doEditAusencia(data) {
  const updates = {};
  if (data.start   !== undefined) updates.start   = data.start;
  if (data.end     !== undefined) updates.end     = data.end;
  if (data.dias    !== undefined) updates.dias    = data.dias;
  if (data.subtipo !== undefined) updates.subtipo = data.subtipo;
  if (data.horas   !== undefined) updates.horas   = data.horas;
  if (data.motivo  !== undefined) updates.motivo  = data.motivo;
  return dbQuery({ action: 'update', table: 'tblAusencias', data: updates, filters: [{ column: 'id', op: 'eq', value: data.id }] });
}
async function doDeleteAusencia(data) {
  return dbQuery({ action: 'delete', table: 'tblAusencias', filters: [{ column: 'id', op: 'eq', value: data.id }] });
}

// ── DOCUMENTOS ───────────────────────────────────────────────
async function getDocumentos() {
  const result = await dbQuery({ action: 'select', table: 'tblControlDocumental', options: { select: '*' } });
  if (!result.ok) return result;
  const hoy = new Date(); hoy.setHours(0,0,0,0);
  const docs = (result.data||[]).map(r => {
    const fv = r.fechavencimiento ? new Date(r.fechavencimiento) : null;
    const dias = fv ? Math.round((fv - hoy) / 86400000) : null;
    return { fuente:'nuevo', id:r.id, numero:r.numero, nombre:r.nombre, estado:r.estado,
      organo:r.organo, fechaInicio:r.fechainicio, fechaVig:r.fechavencimiento,
      creado:r.creado, idDocumento:r.iddocumento, tipoDocumento:r.tipoDocumento,
      diasRestantes:dias, tiempoAviso:30 };
  });
  return { ok: true, data: docs };
}

// ── CARGA INICIAL ────────────────────────────────────────────
async function getInit() {
  try {
    const [f, a, c] = await Promise.all([
      getFichajes(), getAusencias(), getCamiones()
    ]);
    return {
      ok: true,
      fichajes:  f.data || [],
      ausencias: a.data || [],
      camiones:  c.data || [],
      gasoilData: [],
      gasoilDep1: 0,
      gasoilDep2: 0
    };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

// ── GAMAS NORMAS ─────────────────────────────────────────────
// Columnas reales: id, Numero, Gama, Modelo, Intervalo, n1..n60
async function getGamasNormas() {
  return dbQuery({ action: 'select', table: 'tblGamasNormas', options: { select: '*', order: 'id.asc' } });
}
async function doPostGamaNorma(d) {
  // Calcular siguiente id manualmente (la tabla no tiene serial)
  const maxId = normasData.length ? Math.max(...normasData.map(x=>Number(x.id)||0)) : 0;
  const row = { id: maxId+1, Numero: d.Numero||'', Gama: d.Gama||'', Modelo: d.Modelo||'', Intervalo: Number(d.Intervalo)||0 };
  for(let i=1;i<=60;i++) row['n'+i] = d['n'+i]||null;
  return dbQuery({ action: 'insert', table: 'tblGamasNormas', data: row });
}
async function doEditGamaNorma(d) {
  const row = { Numero: d.Numero||'', Gama: d.Gama||'', Modelo: d.Modelo||'', Intervalo: Number(d.Intervalo)||0 };
  for(let i=1;i<=60;i++) row['n'+i] = d['n'+i]||null;
  return dbQuery({ action: 'update', table: 'tblGamasNormas', data: row, filters: [{ column: 'id', op: 'eq', value: d.id }] });
}
async function doDeleteGamaNorma(id) {
  return dbQuery({ action: 'delete', table: 'tblGamasNormas', filters: [{ column: 'id', op: 'eq', value: id }] });
}

// ── GAMAS DEPENDIENTES (Subgamas) ─────────────────────────────
// Columnas reales: id, Gama_Principal, Gama_1..Gama_6
async function getGamasDependientes() {
  return dbQuery({ action: 'select', table: 'tblGamasDependientes', options: { select: '*', order: 'id.asc' } });
}
async function doPostGamaDependiente(d) {
  const maxId = subgamasData.length ? Math.max(...subgamasData.map(x=>Number(x.id)||0)) : 0;
  const row = { id: maxId+1, Gama_Principal: d.Gama_Principal||'' };
  for(let i=1;i<=6;i++) row['Gama_'+i] = d['Gama_'+i]||null;
  return dbQuery({ action: 'insert', table: 'tblGamasDependientes', data: row });
}
async function doEditGamaDependiente(d) {
  const row = { Gama_Principal: d.Gama_Principal||'' };
  for(let i=1;i<=6;i++) row['Gama_'+i] = d['Gama_'+i]||null;
  return dbQuery({ action: 'update', table: 'tblGamasDependientes', data: row, filters: [{ column: 'id', op: 'eq', value: d.id }] });
}
async function doDeleteGamaDependiente(id) {
  return dbQuery({ action: 'delete', table: 'tblGamasDependientes', filters: [{ column: 'id', op: 'eq', value: id }] });
}

// ── GAMAS ACTIVOS ─────────────────────────────────────────────
// Columnas reales: id, Activo, Gama_1..Gama_9, Check_1..Check_3
async function getGamasActivos() {
  return dbQuery({ action: 'select', table: 'tblGamasActivos', options: { select: '*', order: 'Activo.asc' } });
}
async function doPostGamaActivo(d) {
  const maxId = activoGamaData.length ? Math.max(...activoGamaData.map(x=>Number(x.id)||0)) : 0;
  const row = { id: maxId+1, Activo: d.Activo||'' };
  for(let i=1;i<=9;i++) row['Gama_'+i] = d['Gama_'+i]||null;
  for(let i=1;i<=3;i++) row['Check_'+i] = d['Check_'+i]||null;
  return dbQuery({ action: 'insert', table: 'tblGamasActivos', data: row });
}
async function doEditGamaActivo(d) {
  const row = { Activo: d.Activo||'' };
  for(let i=1;i<=9;i++) row['Gama_'+i] = d['Gama_'+i]||null;
  for(let i=1;i<=3;i++) row['Check_'+i] = d['Check_'+i]||null;
  return dbQuery({ action: 'update', table: 'tblGamasActivos', data: row, filters: [{ column: 'id', op: 'eq', value: d.id }] });
}
async function doDeleteGamaActivo(id) {
  return dbQuery({ action: 'delete', table: 'tblGamasActivos', filters: [{ column: 'id', op: 'eq', value: id }] });
}

// ── GAMAS LISTADO PREVENTIVO ──────────────────────────────────
// Columnas reales: id, Activo, Gama, Medidor, Proximo, U_Medicion_med, U_Medicion_fecha, Falta, Estado, Principal, Aviso
async function getGamasListado() {
  return dbQuery({ action: 'select', table: 'tblGamasListadoPreventivo', options: { select: '*', order: 'Activo.asc' } });
}
async function doPostGamaListado(d) {
  const proximo = Number(d.Proximo) || 0;
  const umed    = Number(d.U_Medicion_med) || 0;
  const maxId = listadoPrevData.length ? Math.max(...listadoPrevData.map(x=>Number(x.id)||0)) : 0;
  return dbQuery({ action: 'insert', table: 'tblGamasListadoPreventivo', data: {
    id: maxId+1, Activo: d.Activo||'', Gama: d.Gama||'', Medidor: d.Medidor||'H',
    Proximo: proximo, U_Medicion_med: umed, U_Medicion_fecha: d.U_Medicion_fecha||null,
    Falta: proximo - umed, Estado: d.Estado||null, Principal: d.Principal||false, Aviso: d.Aviso||null
  }, options: { select: 'id' }});
}
async function doEditGamaListado(d) {
  const proximo = Number(d.Proximo) || 0;
  const umed    = Number(d.U_Medicion_med) || 0;
  return dbQuery({ action: 'update', table: 'tblGamasListadoPreventivo', data: {
    Activo: d.Activo||'', Gama: d.Gama||'', Medidor: d.Medidor||'H',
    Proximo: proximo, U_Medicion_med: umed, U_Medicion_fecha: d.U_Medicion_fecha||null,
    Falta: proximo - umed, Estado: d.Estado||null, Principal: d.Principal||false, Aviso: d.Aviso||null
  }, filters: [{ column: 'id', op: 'eq', value: d.id }]});
}
async function doDeleteGamaListado(id) {
  return dbQuery({ action: 'delete', table: 'tblGamasListadoPreventivo', filters: [{ column: 'id', op: 'eq', value: id }] });
}

// ── GAMAS OT ─────────────────────────────────────────────────
// Columnas reales: id, Activo, Ot, Fecha, Operario, Tiempo, Texto, Estado, Gama, Medicion, n1..n60
async function getGamasOT(activo) {
  const filters = activo ? [{ column: 'Activo', op: 'eq', value: activo }] : [];
  return dbQuery({ action: 'select', table: 'tblGamasOT', filters, options: { select: '*', order: 'id.desc' } });
}
async function doPostGamasOT(d) {
  const row = { Activo: d.Activo||'', Ot: Number(d.Ot)||0, Fecha: d.Fecha||null, Operario: d.Operario||'', Tiempo: Number(d.Tiempo)||0, Texto: d.Texto||'', Estado: d.Estado||false, Gama: d.Gama||'', Medicion: Number(d.Medicion)||0 };
  for(let i=1;i<=60;i++) row['n'+i] = d['n'+i]||false;
  return dbQuery({ action: 'insert', table: 'tblGamasOT', data: row });
}
async function doEditGamasOT(d) {
  const row = { Activo: d.Activo||'', Ot: Number(d.Ot)||0, Fecha: d.Fecha||null, Operario: d.Operario||'', Tiempo: Number(d.Tiempo)||0, Texto: d.Texto||'', Estado: d.Estado||false, Gama: d.Gama||'', Medicion: Number(d.Medicion)||0 };
  for(let i=1;i<=60;i++) row['n'+i] = d['n'+i]||false;
  return dbQuery({ action: 'update', table: 'tblGamasOT', data: row, filters: [{ column: 'id', op: 'eq', value: d.id }] });
}
async function doDeleteGamasOT(id) {
  return dbQuery({ action: 'delete', table: 'tblGamasOT', filters: [{ column: 'id', op: 'eq', value: id }] });
}

// ── ROUTER: apiFetch (híbrido Supabase + Google Sheets) ─────
async function apiFetch(params) {
  const p = new URLSearchParams(params.replace(/^\?/,''));
  const accion = p.get('accion') || '';

  // ── Producción y Gasoil → Supabase ──
  if (accion === 'produccion') {
    const mes  = parseInt(p.get('mes'))  || 0;
    const anyo = parseInt(p.get('anyo')) || new Date().getFullYear();
    return getProduccion(mes, anyo);
  }
  if (accion === 'gasoil') return getGasoil();

  // ── Todo lo demás → Supabase ──
  const dias = parseInt(p.get('dias')) || 90;
  if (accion === 'init')        return getInit();
  if (accion === 'fichajes')    return getFichajes();
  if (accion === 'ausencias')   return getAusencias();
  if (accion === 'camiones')    return getCamiones();
  if (accion === 'obras')       return getObras();
  if (accion === 'choferes')    return getChoferes();
  if (accion === 'pedidos')     return getPedidos(dias);
  if (accion === 'historialOT')   return getHistorialOT();
  if (accion === 'documentos')    return getDocumentos();
  if (accion === 'gamasNormas')   return getGamasNormas();
  if (accion === 'gamasSubgamas') return getGamasDependientes();
  if (accion === 'gamasActivos')  return getGamasActivos();
  if (accion === 'gamasListado')  return getGamasListado();
  if (accion === 'gamasOT')       return getGamasOT(p.get('activo')||null);
  if (accion === 'peso')          return { ok: false, error: 'Lectura de báscula sólo por puerto serie' };
  return { ok: false, error: 'Acción desconocida: ' + accion };
}

// ── ROUTER: apiPost (híbrido Supabase + Google Sheets) ──────
async function apiPost(payload) {
  const t = payload.tipo || '';

  // ── Producción y Gasoil → Google Sheets ──
  if (t === 'gasoil' || t === 'editarGasoil' || t === 'editProduccion' || t === 'addProduccion') {
    return sheetsPost(payload);
  }

  // ── Todo lo demás → Supabase ──
  if (t === 'fichajeEntrada')  return doPostEntrada(payload);
  if (t === 'fichajeManual')   return doPostFichajeManual(payload);
  if (t === 'fichajesSalida')  return doEditSalida(payload);
  if (t === 'editFichaje')     return doEditFichaje(payload);
  if (t === 'delFichaje')      return doDeleteFichaje(payload);
  if (t === 'pesada')          return doPostPesada(payload);
  if (t === 'deletePedido')    return doDeletePedido(payload);
  if (t === 'editarPedido')    return doEditarPedido(payload);
  if (t === 'nuevoCamion')     return doNuevoCamion(payload);
  if (t === 'editarCamion')    return doEditarCamion(payload);
  if (t === 'eliminarCamion')  return doEliminarCamion(payload);
  if (t === 'nuevaObra')       return doNuevaObra(payload);
  if (t === 'editarObra')      return doEditarObra(payload);
  if (t === 'eliminarObra')    return doEliminarObra(payload);
  if (t === 'nuevoChofer')     return doNuevoChofer(payload);
  if (t === 'editarChofer')    return doEditarChofer(payload);
  if (t === 'eliminarChofer')  return doEliminarChofer(payload);
  if (t === 'ausencia')        return doPostAusencia(payload);
  if (t === 'editAusencia')    return doEditAusencia(payload);
  if (t === 'delAusencia')     return doDeleteAusencia(payload);
  if (t === 'editarOT')            return doEditarOT(payload);
  if (t === 'deleteOT')            return doDeleteOT(payload);
  if (t === 'postGamaNorma')       return doPostGamaNorma(payload);
  if (t === 'editGamaNorma')       return doEditGamaNorma(payload);
  if (t === 'delGamaNorma')        return doDeleteGamaNorma(payload.id);
  if (t === 'postGamaSubgama')     return doPostGamaDependiente(payload);
  if (t === 'editGamaSubgama')     return doEditGamaDependiente(payload);
  if (t === 'delGamaSubgama')      return doDeleteGamaDependiente(payload.id);
  if (t === 'postGamaActivo')      return doPostGamaActivo(payload);
  if (t === 'editGamaActivo')      return doEditGamaActivo(payload);
  if (t === 'delGamaActivo')       return doDeleteGamaActivo(payload.id);
  if (t === 'postGamaListado')     return doPostGamaListado(payload);
  if (t === 'editGamaListado')     return doEditGamaListado(payload);
  if (t === 'delGamaListado')      return doDeleteGamaListado(payload.id);
  if (t === 'postGamasOT')         return doPostGamasOT(payload);
  if (t === 'editGamasOT')         return doEditGamasOT(payload);
  if (t === 'delGamasOT')          return doDeleteGamasOT(payload.id);
  return doPostOT(payload);
}

let loginUser=null; // {id, nombre, rol} — SIN pin
const WORKERS=['Gabriel Reyes','David Espacios','Antonio Juan Martel','Rubén Díaz'];
const TOTAL_VAC=35;
// Horas esperadas por día de semana: 1=Lun,2=Mar,3=Mié,4=Jue,5=Vie,6=Sáb,0=Dom
const HORARIOS={
  'David Espacios':      {1:8,2:8,3:8,4:8,5:8,6:0,0:0},
  'Gabriel Reyes':       {1:10,2:8,3:8,4:8,5:8,6:0,0:0},
  'Antonio Juan Martel': {1:10,2:10,3:10,4:10,5:0,6:0,0:0},
  'Rubén Díaz':          {1:10,2:10,3:10,4:10,5:0,6:0,0:0},
};
const MESES=['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
const FESTIVOS=['2026-01-01','2026-01-02','2026-01-05','2026-01-06','2026-04-02','2026-04-03','2026-05-01','2026-08-15','2026-10-12','2026-11-01','2026-12-06','2026-12-08','2026-12-25'];
// Días laborables por mes según convenio de la construcción (0=ene...11=dic)
const DIAS_LAB_2026=[18,18,22,20,20,21,23,21,22,21,21,18];
const HORAS_DIA_STD=8; // Jornada estándar convenio
const PRODS=['ARIDO AF-T-0/4-I','ARIDO AG-T-4/12-I','ARIDO AG-T-12/20-I','ARIDO AG-T-20/40-I','ARIDO AG-T-40/70-I','REVUELTO 0/20','REVUELTO 0/10','PIEDRA PARA MURO (UD)','MATERIAL DE RELLENO 0/4'];
const PROD_CAT={'ARIDO AF-T-0/4-I':'0/4','ARIDO AG-T-4/12-I':'4/12','ARIDO AG-T-12/20-I':'12/20','ARIDO AG-T-20/40-I':'20/40'};
const PAGE_TITLES={inicio:'Inicio',bascula:'Pesada',pedidos:'Pedidos',facturacion:'Facturación',ventas:'Ventas','historico-ventas':'Histórico de Ventas',caja:'Caja',costes:'Análisis de Costes',produccion:'Producción Planta',informes:'Informes Planta',stock:'Stock Áridos',camiones:'Camiones',gasoil:'Gasoil',activos:'Activos / Maquinaria',fichaje:'Fichaje',resumen:'Resumen',vacaciones:'Vacaciones',calendario:'Calendario laboral',editar:'Editar fichajes',ot:'Nueva OT','historial-ot':'Historial OT',documentos:'Control Documental',tareas:'Tareas',preventivo:'Mantenimiento Preventivo',compras:'Escanear Factura',choferes:'Conductores',ensayos:'Control de Ensayos'};

// Login via Google OAuth — ver funciones.js: googleLogin() y checkGoogleSession()

// ── SHELL ─────────────────────────────────────────────────────
let menuOpen=false;
function toggleMenu(){menuOpen=!menuOpen;document.getElementById('sidebar').classList.toggle('open',menuOpen);document.getElementById('menu-btn').classList.toggle('open',menuOpen);}

// ── Bottom Nav (móvil) ───────────────────────────────────────
const BNAV_MAIN=['inicio','bascula','fichaje','ot'];
function bnavGo(id){closeBnavMore();goPage(id);}
function toggleBnavMore(){
  const panel=document.getElementById('bnav-more-panel');
  const overlay=document.getElementById('bnav-overlay');
  const open=panel.classList.toggle('open');
  overlay.classList.toggle('open',open);
}
function closeBnavMore(){
  const panel=document.getElementById('bnav-more-panel');
  const overlay=document.getElementById('bnav-overlay');
  if(panel){panel.classList.remove('open');}
  if(overlay){overlay.classList.remove('open');}
}
function updateBnav(id){
  document.querySelectorAll('.bnav-btn').forEach(b=>b.classList.remove('active'));
  document.querySelectorAll('.bnav-more-btn').forEach(b=>b.classList.remove('active'));
  const mainBtn=document.getElementById('bn-'+id);
  if(mainBtn){mainBtn.classList.add('active');}
  else{
    const moreBtn=document.getElementById('bm-'+id);
    if(moreBtn)moreBtn.classList.add('active');
    document.getElementById('bn-more').classList.add('active');
  }
}

function goPage(id){
  document.querySelectorAll('.page').forEach(p=>p.classList.remove('active'));
  document.querySelectorAll('.snav-btn').forEach(b=>b.classList.remove('active'));
  document.getElementById('pg-'+id).classList.add('active');
  const btn=document.getElementById('snav-'+id);if(btn)btn.classList.add('active');
  document.getElementById('topbar-title').textContent=PAGE_TITLES[id]||id;
  if(menuOpen)toggleMenu();
  updateBnav(id);
  if(id==='resumen'){renderResumenCards();renderMeses();renderHistorial();renderExtras();}
  if(id==='vacaciones'){renderVac();renderBajas();renderDiasLibres();renderExtrasManual();}
  if(id==='calendario')renderCal();
  if(id==='editar')renderEditar();
  if(id==='historial-ot')cargarHistorialOT();
  if(id==='documentos')cargarDocumentos();
  if(id==='preventivo')cargarMantenimientoPreventivo();
  if(id==='inicio')renderInicioDocs();
  if(id==='pedidos')cargarPedidos();
  if(id==='ventas'){if(ventasData.length===0)cargarVentas();else renderVentas();}
  if(id==='camiones')cargarCamiones();
  if(id==='choferes')cargarChoferes();
  if(id==='obras')cargarObras();
  if(id==='activos')initActivos();
  if(id==='compras')comprasInitProveedores();
  if(id==='gasoil')cargarGasoil();
  if(id==='produccion')initProduccion();
  if(id==='informes'){const fi=document.getElementById('inf-fecha');if(fi&&!fi.value)fi.value=new Date().toISOString().slice(0,10);}
  if(id==='stock')initStock();
  if(id==='facturacion')initFacturacion();
  if(id==='caja')initCaja();
  if(id==='costes')initCostes();
  if(id==='historico-ventas')initHistoricoVentas();
  if(id==='tareas')initTareasPanel();
  if(id==='ensayos')initEnsayos();
}

function goTareasSeccion(seccion){
  goPage('tareas');
  const sel=document.getElementById('tarea-filt-seccion');
  if(sel){sel.value=seccion;tareaSeccionChange();}
}

function pad(n){return String(n).padStart(2,'0')}
function fmtH(ts){const d=new Date(ts);return pad(d.getHours())+':'+pad(d.getMinutes())+':'+pad(d.getSeconds())}
function fmtDur(ms){const h=Math.floor(ms/3600000),m=Math.floor((ms%3600000)/60000);return h+'h '+m+'m'}
function fmtDurDec(ms){return(ms/3600000).toFixed(1)+'h'}
function dateStr(d){return d.getFullYear()+'-'+pad(d.getMonth()+1)+'-'+pad(d.getDate())}
function fmtDate(s){const p=s.split('-');return p[2]+'/'+p[1]+'/'+p[0]}
function fmtFecha(ts){const d=new Date(ts);return pad(d.getDate())+'/'+pad(d.getMonth()+1)+'/'+d.getFullYear();}
function fmtFechaHora(ts){const d=new Date(ts);return d.getFullYear()+'-'+pad(d.getMonth()+1)+'-'+pad(d.getDate())+'T'+pad(d.getHours())+':'+pad(d.getMinutes())+':00';}

function renderInicioDocs(){
  const el=document.getElementById('inicio-docs-alert');
  if(!docData.length){el.style.display='none';}else{
    const rojos=docData.filter(d=>docEstado(d.diasRestantes,d.tiempoAviso)==='rojo').length;
    const amarillos=docData.filter(d=>docEstado(d.diasRestantes,d.tiempoAviso)==='amarillo').length;
    if(!rojos&&!amarillos){el.style.display='none';}else{
      el.style.display='block';
      el.innerHTML=`<div style="background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:12px 14px;cursor:pointer" onclick="goPage('documentos')">
        <div style="font-size:.7rem;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.06em;margin-bottom:6px">⚠ Alertas documentales</div>
        ${rojos?`<div style="color:#ff4d4d;font-size:.82rem;font-weight:700">${rojos} documento${rojos>1?'s':''} vencido${rojos>1?'s':''}</div>`:''}
        ${amarillos?`<div style="color:#f5a623;font-size:.82rem;font-weight:700">${amarillos} documento${amarillos>1?'s':''} por vencer</div>`:''}
      </div>`;
    }
  }
  renderInicioMant();
}

async function cargarAusencias(){
  try{const json=await apiFetch('?accion=ausencias');procesarAusencias(json);}catch(e){console.log('Error cargando ausencias:',e);}
}

async function cargarInit(){
  try{
    const json=await apiFetch('?accion=init');
    if(!json.ok)throw new Error('Error init');
    // Fichajes
    if(json.fichajes&&json.fichajes.length) procesarFichajes({ok:true,data:json.fichajes});
    // Ausencias
    if(json.ausencias) procesarAusencias({ok:true,data:json.ausencias});
    // Camiones
    if(json.camiones){camionesData=json.camiones;renderCamGrid(camionesData);}
    // Gasoil
    if(json.gasoilData){
      gasoilData=json.gasoilData;
      gasoilConsumos=json.gasoilConsumos||[];
      gasoilStock={dep1:json.gasoilDep1||0,dep2:json.gasoilDep2||0};
    }
  }catch(e){
    console.warn('Init falló, cargando por separado...',e);
    cargarFichajes();cargarAusencias();initBasculaCamiones();
  }
}

function _marcarBotonesLectura(){
  document.querySelectorAll('button:not(.write-action)').forEach(b=>{
    const txt=(b.textContent||'').trim().toLowerCase();
    const oc=(b.getAttribute('onclick')||'').toLowerCase();
    if(txt.startsWith('+ ')||txt.startsWith('guardar')||oc.includes('eliminar')||oc.includes('delete')||oc.includes('save')||oc.includes('modal(null')||oc.includes('guardar'))
      b.classList.add('write-action');
  });
}

async function initApp(){
  if (window._appInitialized) {
    console.warn('initApp() already running, skipping duplicate call');
    return;
  }
  window._appInitialized = true;

  // Rol lectura: ocultar todos los botones de escritura
  if(loginUser&&loginUser.rol==='lectura'){
    document.body.classList.add('rol-lectura');
    _marcarBotonesLectura();
    // Observer para botones renderizados dinámicamente
    new MutationObserver(_marcarBotonesLectura).observe(document.body,{childList:true,subtree:true});
  }

  loadFst();

  // Cargar productos y vendors de BC en background
  cargarProductosBC();

  // Event delegation para botones de fichaje
  document.addEventListener('click', e => {
    if (e.target.classList.contains('wbtn')) {
      const nombre = e.target.dataset.worker;
      if (nombre) handleFichar(nombre);
    }
  });

  // Lanzar queries en paralelo para reducir tiempo de carga
  const hoy = new Date().toISOString().slice(0, 10);
  const [fichajeHoy, initData] = await Promise.all([
    dbQuery({ action: 'select', table: 'tblFichaje',
      filters: [{ column: 'fecha', op: 'eq', value: hoy }],
      options: { select: 'empleado,entrada,salida,fentrada' }
    }).catch(e => { console.error('Error fichaje hoy:', e); return { ok: false }; }),
    cargarInit().catch(e => { console.warn('cargarInit error:', e); })
  ]);

  // Procesar fichajes de hoy
  try {
    const data = fichajeHoy.data;
    if (fichajeHoy.ok && data && data.length > 0) {
      WORKERS.forEach(n => {
        fst.workers[n].working = false;
        fst.workers[n].entradaTs = null;
      });
      data.forEach(r => {
        const nombreDB = r.empleado.toUpperCase();
        const worker = WORKERS.find(w => w.toUpperCase() === nombreDB);
        if (worker && r.entrada && !r.salida) {
          fst.workers[worker].working = true;
          if (r.fentrada) {
            fst.workers[worker].entradaTs = new Date(r.fentrada).getTime();
          }
        }
      });
    } else {
      WORKERS.forEach(n => recalcWorker(n));
    }
  } catch(e) {
    console.error('Error actualizando estado HOY:', e);
    WORKERS.forEach(n => recalcWorker(n));
  }

  renderWgrid();renderStats();renderVac();renderCal();initOT();
  WORKERS.forEach(n => renderWcard(n));
  initBasculaUI();
  // Load OT history in background for reminders
  apiFetch('?accion=historialOT').then(j=>{if(j.ok){prevData=j.data;renderInicioMant();}}).catch(()=>{});
  goPage('inicio');
  cargarNotas();
  // Comprobar facturas vencidas en segundo plano
  checkFacturasVencidasBackground().catch(() => {});
  ['filt-w','filt-mes-w','filt-edit'].forEach(id=>{const s=document.getElementById(id);if(!s)return;WORKERS.forEach(n=>{const o=document.createElement('option');o.value=n;o.textContent=n;s.appendChild(o);});});
  ['vm-worker','em-worker','bm-worker','lm-worker','xm-worker'].forEach(id=>{const s=document.getElementById(id);if(!s)return;WORKERS.forEach(n=>{const o=document.createElement('option');o.value=n;o.textContent=n;s.appendChild(o);});});
  document.getElementById('ventas-mes-sel').value=new Date().getMonth();
  setInterval(actualizarReloj,1000);
  setInterval(()=>{WORKERS.filter(n=>fst.workers[n].working).forEach(n=>renderWcard(n));renderStats();},5000);
  actualizarReloj();
  const gfecha=document.getElementById('gasoil-fecha-inp');if(gfecha)gfecha.value=dateStr(new Date());
}

function actualizarReloj(){
  const now=new Date();
  const t=pad(now.getHours())+':'+pad(now.getMinutes())+':'+pad(now.getSeconds());
  const el=document.getElementById('clock');if(el)el.textContent=t;
  document.getElementById('topbar-clock').textContent=t;
  const dias=['Domingo','Lunes','Martes','Miércoles','Jueves','Viernes','Sábado'];
  const dl=document.getElementById('date-lbl');if(dl)dl.textContent=dias[now.getDay()]+', '+now.getDate()+' de '+MESES[now.getMonth()]+' de '+now.getFullYear();
}

// ── BÁSCULA ───────────────────────────────────────────────────
let CLIENTES=[
  {nombre:'ALONSO MAYOR AMBROSIO AMADO',codigo:'CLI-00027'},
  {nombre:'ANGEL SALVADOR SANTANA PEÑA',codigo:'CLI-00024'},
  {nombre:'API MOVILIDAD, S.A.',codigo:'CLI-00005'},
  {nombre:'CANARIAS BETON SL',codigo:'CLI-00002'},
  {nombre:'CLIENTES VARIOS',codigo:'CLI-00000'},
  {nombre:'CONTROLES Y ACCIONAMIENTOS CANARIOS SL',codigo:'CLI-00026'},
  {nombre:'CORRALO SL',codigo:'CLI-00009'},
  {nombre:'CRONISTAURO PLATER SL',codigo:'CLI-00032'},
  {nombre:'DOCARFRA SLU',codigo:'CLI-00006'},
  {nombre:'FELIPE Y NICOLAS SL',codigo:'CLI-00025'},
  {nombre:'FRANCISCO ARAÑA MELIAN',codigo:'CLI-00046'},
  {nombre:'HORMISOL CANARIAS SA',codigo:'CLI-00004'},
  {nombre:'ISMAEL GARCIA OJEDA',codigo:'CLI-00007'},
  {nombre:'JOSE ANTONIO MESA GONZALEZ',codigo:'CLI-00033'},
  {nombre:'JOSE ENCINOSO SANCHEZ',codigo:'CLI-00014'},
  {nombre:'JOSÉ MIGUEL SUÁREZ MARTÍN',codigo:'CLI-00031'},
  {nombre:'JUAN RAMON LEON ARENCIBIA',codigo:'CLI-00045'},
  {nombre:'MADRIAN OBRAS Y REFORMAS S.L.',codigo:'CLI-00017'},
  {nombre:'MANTENIMIENTOS LAS TIRAJANA S.L',codigo:'CLI-00019'},
  {nombre:'MAQUINARIA HERMANOS ASCANIO SL',codigo:'CLI-00029'},
  {nombre:'NOELIO LOPEZ MARTEL',codigo:'CLI-00037'},
  {nombre:'NORTE SUR OBRAS Y REFORMAS S.L',codigo:'CLI-00035'},
  {nombre:'PREFABRICADOS ARCHIPIELAGO SL',codigo:'CLI-00003'},
  {nombre:'PREFABRICADOS LEMES SL',codigo:'CLI-00041'},
  {nombre:'SANTANA Y MADRE SL',codigo:'CLI-00013'},
  {nombre:'SECULAR 2022 SLU',codigo:'CLI-00038'},
  {nombre:'SERVICIOS Y MANTENIMIENTOS LAS TIRAJANA',codigo:'CLI-00019'},
  {nombre:'SHUTTLE TRUCK SL',codigo:'CLI-00008'},
  {nombre:'SURHISA SUAREZ E HIJOS SL',codigo:'CLI-00044'},
  {nombre:'TRANSPORTES GUEPEVECA SL',codigo:'CLI-00028'},
  {nombre:'TRANSPORTES LUJAN S.L',codigo:'CLI-00020'},
  {nombre:'TRANSPORTES ROMANO PERERA SL',codigo:'CLI-00011'},
  {nombre:'TRANSPORTES Y GRUAS SANCHEZ CANARIAS SL',codigo:'CLI-00012'},
];
let CLI_PROY={
  'ALONSO MAYOR AMBROSIO AMADO':[{nombre:'OBRAS ALONSO MAYOR AMBROSIO AMADO',codigo:'PV-020'}],
  'ANGEL SALVADOR SANTANA PEÑA':[{nombre:'CLIENTES VARIOS',codigo:'PV-000'},{nombre:'OBRAS ANGEL SALVADOR',codigo:'PV-027'}],
  'API MOVILIDAD, S.A.':[{nombre:'APIMOVILIDAD CRTRA VALLESECO-VALSENDERO',codigo:'PV-007'}],
  'CANARIAS BETON SL':[{nombre:'CANARIAS BETON ARINAGA',codigo:'PV-001'},{nombre:'CANARIAS BETON JINAMAR',codigo:'PV-002'}],
  'CLIENTES VARIOS':[{nombre:'CLIENTES VARIOS',codigo:'PV-000'}],
  'CONTROLES Y ACCIONAMIENTOS CANARIOS SL':[{nombre:'CLIENTES VARIOS',codigo:'PV-000'},{nombre:'OBRAS CONTROLES Y ACCIONAMIENTOS CANARIO',codigo:'PV-022'}],
  'CORRALO SL':[{nombre:'CORRALO BAHIA FELIZ',codigo:'PV-010'}],
  'CRONISTAURO PLATER SL':[{nombre:'OBRAS CRONISTAURO',codigo:'PV-025'}],
  'DOCARFRA SLU':[{nombre:'DOCARFRA HOYA DE TUNTE',codigo:'PV-008'}],
  'FELIPE Y NICOLAS SL':[{nombre:'FELIPE Y NICOLAS',codigo:'PV-029'},{nombre:'OBRAS FELIPE Y NICOLAS',codigo:'PV-029'}],
  'FRANCISCO ARAÑA MELIAN':[{nombre:'CLIENTES VARIOS',codigo:'PV-000'}],
  'HORMISOL CANARIAS SA':[{nombre:'HORMISOL ARGUINEGUÍN',codigo:'PV-006'},{nombre:'HORMISOL ARINAGA',codigo:'PV-004'},{nombre:'HORMISOL LAS TORRES',codigo:'PV-005'}],
  'ISMAEL GARCIA OJEDA':[{nombre:'CLIENTES VARIOS',codigo:'PV-000'},{nombre:'OBRAS ISMAEL GARCIA OJEDA',codigo:'PV-011'}],
  'JOSE ANTONIO MESA GONZALEZ':[{nombre:'CLIENTES VARIOS',codigo:'PV-000'},{nombre:'JOSE ANTONIO MESA GONZALEZ',codigo:'PV-024'}],
  'JOSE ENCINOSO SANCHEZ':[{nombre:'JOSE ENCINOSO SANCHEZ OBRAS',codigo:'PV-014'}],
  'JOSÉ MIGUEL SUÁREZ MARTÍN':[{nombre:'CLIENTES VARIOS',codigo:'PV-000'},{nombre:'JOSÉ MIGUEL SUÁREZ MARTÍN',codigo:'PV-023'}],
  'JUAN RAMON LEON ARENCIBIA':[{nombre:'CLIENTES VARIOS',codigo:'PV-000'},{nombre:'VENTAS JUAN RAMON LEON ARENCIBIA',codigo:'PV-034'}],
  'MADRIAN OBRAS Y REFORMAS S.L.':[{nombre:'TRANSPORTE GUTIERREZ SARDINA',codigo:'PV-015'}],
  'MANTENIMIENTOS LAS TIRAJANA S.L':[{nombre:'OBRAS SERVICIOS LAS TIRAJANA S.L',codigo:'PV-018'}],
  'MAQUINARIA HERMANOS ASCANIO SL':[{nombre:'CLIENTES VARIOS',codigo:'PV-000'},{nombre:'OBRAS MAQUINARIA ASCANIO',codigo:'PV-026'}],
  'NOELIO LOPEZ MARTEL':[{nombre:'NOELIO LOPEZ MARTEL',codigo:'PV-030'}],
  'NORTE SUR OBRAS Y REFORMAS S.L':[{nombre:'OBRAS NORTE SUR PILAR',codigo:'PV-028'}],
  'PREFABRICADOS ARCHIPIELAGO SL':[{nombre:'PREARSA ARINAGA',codigo:'PV-003'},{nombre:'PREARSA JINAMAR',codigo:'PV-009'}],
  'PREFABRICADOS LEMES SL':[{nombre:'OBRAS PREFABRICADOS LEMES',codigo:'PV-032'}],
  'SANTANA Y MADRE SL':[{nombre:'OBRAS SANTANA Y MADRE',codigo:'PV-013'}],
  'SECULAR 2022 SLU':[{nombre:'OBRAS SECULAR 2022',codigo:'PV-031'}],
  'SERVICIOS Y MANTENIMIENTOS LAS TIRAJANA':[{nombre:'OBRAS SERVICIOS LAS TIRAJANA S.L',codigo:'PV-018'}],
  'SHUTTLE TRUCK SL':[{nombre:'SHUTTLE TRUCK OJOS DE GARZA',codigo:'PV-019'}],
  'SURHISA SUAREZ E HIJOS SL':[{nombre:'SURHISA SUAREZ E HIJOS SL',codigo:'PV-033'}],
  'TRANSPORTES GUEPEVECA SL':[{nombre:'OBRAS TRANSPORTES GUEPEVECA',codigo:'PV-021'}],
  'TRANSPORTES LUJAN S.L':[{nombre:'OBRAS TRANSPORTES LUJAN S.L',codigo:'PV-016'}],
  'TRANSPORTES ROMANO PERERA SL':[{nombre:'TRANSPORTE ROMANO MASPALOMAS',codigo:'PV-012'}],
  'TRANSPORTES Y GRUAS SANCHEZ CANARIAS SL':[{nombre:'OBRAS GRUAS SANCHEZ CANARIAS SL',codigo:'PV-017'}],
};
// Fallback local — se sobreescribe con datos de BC al cargar
let PROD_MAP={
  'PROD-000027':'ARIDO AF-T-0/4-I','PROD-000028':'ARIDO AG-T-4/12-I',
  'PROD-000029':'ARIDO AG-T-12/20-I','PROD-000030':'ARIDO AG-T-20/40-I',
  'PROD-000031':'ARIDO AG-T-40/70-I','PROD-000032':'ESCOLLERA',
  'PROD-000033':'REVUELTO 0/10','PROD-000034':'PIEDRA PARA MURO (UD)',
  'PROD-000035':'MATERIAL DE RELLENO 0/4',
  'PROD-000038':'ZAHORRA','PROD-000057':'REVUELTO 0/20','PROD-000058':'ZAHORRA 0/40',
  'PROD-000048':'MATERIAL DE RELLENO 0/4','PROD-000059':'TIERRA VEGETAL',
  'PROD-000062':'PIEDRA PARA MURO',
  'PROD-000003':'MATERIAL TODO UNO CANTERA',
};
async function cargarProductosBC(){
  try{
    const token=await getBCTokenSilent();
    const resp=await fetch('/api/bc/items',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({token})});
    const data=await resp.json();
    if(data.ok&&data.items.length){
      const map={};
      data.items.forEach(i=>{if(i.number&&(i.displayName||i.name))map[i.number]=i.displayName||i.name;});
      PROD_MAP=map;
      console.log('PROD_MAP cargado desde BC:',Object.keys(map).length,'productos');
    }
  }catch(e){console.warn('No se pudo cargar productos de BC, usando mapa local:',e.message);}
}
const PRECIOS={
  'ARIDO AF-T-0/4-I':16.10,'ARIDO AG-T-4/12-I':15.10,
  'ARIDO AG-T-12/20-I':15.10,'ARIDO AG-T-20/40-I':15.10,
  'ARIDO AG-T-40/70-I':16.00,'REVUELTO 0/20':15.60,
  'REVUELTO 0/10':15.60,'ZAHORRA 0/40':15.60,
  'MATERIAL DE RELLENO 0/4':2.00,'TIERRA VEGETAL':18.00,
  'PIEDRA PARA MURO (UD)':19.80,'MATERIAL TODO UNO CANTERA':15.10,
};
const IGIC_PCT=3;
const PRECIOS_ESP={
  'CANARIAS BETON SL':{'ARIDO AF-T-0/4-I':12.30,'_default':11.00},
  'PREFABRICADOS ARCHIPIELAGO SL':{'ARIDO AF-T-0/4-I':12.30,'_default':11.00},
};
function getPrecioTn(cliente,producto){
  const esp=PRECIOS_ESP[cliente];
  if(esp) return esp[producto]!==undefined?esp[producto]:(esp._default!==undefined?esp._default:(PRECIOS[producto]||0));
  return PRECIOS[producto]||0;
}

let camionesData=[];
let basSelCamion=null;
let basSelCliente=null;
let basNumPedido=null;
let basLineasSesion=[];
let basCurrentLinea=10000;

function initBasculaUI(){
  const now=new Date();
  const hoy=pad(now.getDate())+'/'+pad(now.getMonth()+1)+'/'+now.getFullYear();
  const el=document.getElementById('bas-fecha-hoy');if(el)el.textContent=hoy;
  const fp=document.getElementById('bas-fecha-pedido');if(fp)fp.value=dateStr(now);
  renderCamGrid(camionesData);
  renderCliDropdown('');
  // Cargar clientes y proyectos desde BC en background
  _cargarClientesYProyectosBC();
}

async function _ensureBCCompanyId(token){
  if(window._bcCompanyId)return window._bcCompanyId;
  const base=`https://api.businesscentral.dynamics.com/v2.0/${BC_TENANT}/${BC_ENV}/api/v2.0/companies`;
  const cJson=await(await fetch(base,{headers:{'Authorization':`Bearer ${token}`}})).json();
  const company=(cJson.value||[]).find(c=>c.name===BC_COMPANY);
  if(!company)return null;
  window._bcCompanyId=company.id;
  return company.id;
}

async function _cargarClientesYProyectosBC(){
  // 1. Cargar obras desde tblobras (Supabase)
  try{
    const obrasJson=await apiFetch('?accion=obras');
    if(obrasJson.ok&&obrasJson.data&&obrasJson.data.length>0){
      obrasGestData=obrasJson.data;
      _actualizarCliProyDesdeObras();
      console.log('Obras: '+obrasJson.data.length+' cargadas desde tblobras');
    }
  }catch(e){console.warn('Error cargando obras:',e.message);}

  // 2. Cargar clientes desde BC
  try{
    if(typeof getBCTokenSilent!=='function')return;
    const token=await getBCTokenSilent();
    const companyId=await _ensureBCCompanyId(token);
    if(!companyId)return;
    const base=`https://api.businesscentral.dynamics.com/v2.0/${BC_TENANT}/${BC_ENV}/api/v2.0/companies(${companyId})`;
    const headers={'Authorization':`Bearer ${token}`};

    const custUrl=`${base}/customers?$select=number,displayName&$orderby=displayName&$top=500`;
    const custRes=await fetch(custUrl,{headers});
    if(custRes.ok){
      const custJson=await custRes.json();
      const bcClientes=(custJson.value||[]).map(c=>({nombre:c.displayName,codigo:c.number})).filter(c=>c.nombre&&c.codigo);
      if(bcClientes.length>0){
        CLIENTES=bcClientes;
        renderCliDropdown('');
        console.log('BC: '+bcClientes.length+' clientes cargados');
      }
    }

  }catch(e){console.warn('BC clientes:',e.message);}
}
async function initBasculaCamiones(){
  try{
    const json=await apiFetch('?accion=camiones');
    if(json.ok)camionesData=json.data;
  }catch(e){console.warn('Error cargando camiones',e);}
  renderCamGrid(camionesData);
}
// Alias para compatibilidad
async function initBascula(){initBasculaUI();await initBasculaCamiones();}

let camFavoritos=new Set(JSON.parse(localStorage.getItem('camFavoritos')||'[]'));
function toggleFav(id,e){
  e.stopPropagation();
  id=String(id);
  if(camFavoritos.has(id))camFavoritos.delete(id);else camFavoritos.add(id);
  localStorage.setItem('camFavoritos',JSON.stringify([...camFavoritos]));
  basCamBuscar();
}
function camCard(c){
  const sel=basSelCamion&&basSelCamion.id===c.id;
  const fav=camFavoritos.has(String(c.id));
  return `<div class="mat-btn${sel?' selected':''}" onclick="basSeleccionarCamion('${c.id}')" style="position:relative">
    <span onclick="toggleFav('${c.id}',event)" style="position:absolute;top:4px;right:4px;font-size:.75rem;cursor:pointer;opacity:${fav?1:.3}" title="Favorito">${fav?'★':'☆'}</span>
    ${c.matriculacam}
    <span class="mat-sub">${c.chofer||''}</span>
    <span class="mat-sub" style="color:var(--muted);font-size:.58rem">${c.proveedor||''}</span>
  </div>`;
}
function renderCamGrid(lista){
  const grid=document.getElementById('bas-cam-grid');if(!grid)return;
  const q=document.getElementById('bas-cam-buscar').value.toUpperCase();
  if(q||!camFavoritos.size){
    grid.innerHTML=lista.map(camCard).join('');
    return;
  }
  const favs=lista.filter(c=>camFavoritos.has(String(c.id)));
  const resto=lista.filter(c=>!camFavoritos.has(String(c.id)));
  grid.innerHTML=
    (favs.length?`<div style="grid-column:1/-1;font-size:.62rem;font-weight:700;color:var(--accent);text-transform:uppercase;letter-spacing:.07em;padding:2px 0 4px">★ Favoritos</div>${favs.map(camCard).join('')}`:'')
    +(resto.length?`<div style="grid-column:1/-1;font-size:.62rem;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.07em;padding:6px 0 4px">Todos</div>${resto.map(camCard).join('')}`:'');
}

function basCamBuscar(){
  const q=document.getElementById('bas-cam-buscar').value.toUpperCase();
  const filtered=q?camionesData.filter(c=>c.matriculacam.includes(q)||String(c.chofer||'').toUpperCase().includes(q)):camionesData;
  renderCamGrid(filtered);
}

function basSeleccionarCamion(id){
  const c=camionesData.find(x=>x.id==id);if(!c)return;
  basSelCamion=c;
  renderCamGrid(camionesData.filter(x=>{
    const q=document.getElementById('bas-cam-buscar').value.toUpperCase();
    return !q||x.matriculacam.includes(q)||String(x.chofer||'').toUpperCase().includes(q);
  }));
  // Update header
  document.getElementById('hdr-cam').textContent=c.matriculacam;
  document.getElementById('hdr-rem').textContent=c.matricularem||'—';
  document.getElementById('hdr-tara').textContent=Number(c.tara||0).toLocaleString()+' kg';
  basActualizarPeso();
  // Check if can continue
  checkBasStep1();
}

// ── SERIAL ───────────────────────────────────────────────────
let serialPort = null;
let serialConnected = false;

async function conectarBascula() {
  const btn = document.getElementById('btn-conectar-serial');
  if (!('serial' in navigator)) {
    alert('Web Serial API no disponible.\nUsa Chrome o Edge en el PC de la báscula.');
    return;
  }
  if (serialConnected) {
    try { await serialPort.close(); } catch(e) {}
    serialPort = null; serialConnected = false;
    btn.innerHTML = '<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.8" style="width:15px;height:15px"><path d="M7 3v4M13 3v4M5 7h10a1 1 0 0 1 1 1v2a5 5 0 0 1-10 0V8a1 1 0 0 1 1-1zM10 14v3M7 17h6"/></svg> Conectar serie';
    btn.style.borderColor = 'var(--muted)'; btn.style.color = 'var(--muted)';
    return;
  }
  try {
    serialPort = await navigator.serial.requestPort();
    await serialPort.open({ baudRate: 9600, dataBits: 8, parity: 'none', stopBits: 1, flowControl: 'none' });
    serialConnected = true;
    btn.innerHTML = '<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.8" style="width:15px;height:15px"><path d="M4 10l4 4 8-8"/></svg> Serie conectada';
    btn.style.borderColor = 'var(--accent2)'; btn.style.color = 'var(--accent2)';
  } catch(e) {
    if (e.name !== 'NotFoundError') alert('Error al conectar: ' + e.message);
  }
}

async function leerDesdeSerial() {
  const decoder = new TextDecoder('latin1');

  // Paso 1: vaciar buffer acumulado (datos viejos) — leer y descartar durante 800ms
  const flush = serialPort.readable.getReader();
  try {
    const flushDeadline = Date.now() + 800;
    while (Date.now() < flushDeadline) {
      const timeLeft = flushDeadline - Date.now();
      if (timeLeft <= 0) break;
      await Promise.race([flush.read(), new Promise(r => setTimeout(r, timeLeft))]);
    }
  } catch(e) {}
  flush.releaseLock();

  // Paso 2: leer datos frescos
  const reader = serialPort.readable.getReader();
  let buffer = '';
  let rawBytes = [];
  const deadline = Date.now() + 5000;
  try {
    while (Date.now() < deadline) {
      let res;
      try {
        const timeLeft = deadline - Date.now();
        res = await Promise.race([
          reader.read(),
          new Promise((_, rej) => setTimeout(() => rej(new Error('timeout')), timeLeft))
        ]);
      } catch(e) { break; }
      if (res.done) break;
      rawBytes.push(...res.value);
      buffer += decoder.decode(res.value);
      if (rawBytes.length >= 200) break;
    }
  } finally {
    reader.releaseLock();
  }


  // Mostrar debug
  const hex = rawBytes.map(b => b.toString(16).padStart(2,'0')).join(' ');
  const dbg = document.getElementById('serial-debug');
  if (dbg) {
    dbg.style.display = 'block';
    dbg.textContent = 'HEX: ' + (hex||'(vacío)') + '\nTEXTO: ' + JSON.stringify(buffer);
  }

  if (rawBytes.length === 0) throw new Error('Sin datos. ¿Báscula encendida y cable conectado?');

  // Texto ASCII: coger el ÚLTIMO número del buffer (el más reciente)
  const nums = [...buffer.matchAll(/\d+\.?\d*/g)]
    .map(m => parseFloat(m[0])).filter(n => n >= 0);
  if (nums.length > 0) {
    const nonZero = nums.filter(n => n > 0);
    return nonZero.length > 0 ? Math.max(...nonZero) : 0;
  }

  // Intento 2: Modbus RTU — registros 16-bit big-endian
  for (let i = 3; i < rawBytes.length - 1; i++) {
    const val = (rawBytes[i] << 8) | rawBytes[i+1];
    if (val > 0 && val < 99999) return val;
  }

  throw new Error('Datos recibidos pero formato desconocido. Ver debug.');
}

async function leerBascula(){
  const btn=document.getElementById('btn-leer-bascula');
  const svgScale='<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.8" style="width:18px;height:18px"><ellipse cx="10" cy="6" rx="7" ry="2.5"/><path d="M3 6v4c0 1.4 3.1 2.5 7 2.5s7-1.1 7-2.5V6"/><rect x="8.5" y="12.5" width="3" height="4" rx=".5"/><rect x="5" y="16.5" width="10" height="1.5" rx=".7"/></svg>';
  btn.textContent='Leyendo...';btn.disabled=true;
  try{
    let peso;
    if (serialConnected) {
      peso = await leerDesdeSerial();
    } else {
      const json=await apiFetch('?accion=peso');
      if(!json.ok)throw new Error(json.error);
      peso = json.peso;
    }
    document.getElementById('bas-peso-input').value=peso;
    basActualizarPeso();
    btn.innerHTML='<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.8" style="width:18px;height:18px"><path d="M4 10l4 4 8-8" stroke="var(--accent2)"/></svg> '+peso.toLocaleString()+' kg';
    btn.style.borderColor='var(--accent2)';btn.style.color='var(--accent2)';
    setTimeout(()=>{
      btn.innerHTML=svgScale+' Leer báscula';
      btn.style.borderColor='var(--accent)';btn.style.color='var(--accent)';btn.disabled=false;
    },3000);
  }catch(e){
    btn.textContent='Error: '+e.message;
    setTimeout(()=>{
      btn.innerHTML=svgScale+' Leer báscula';
      btn.style.borderColor='var(--accent)';btn.style.color='var(--accent)';btn.disabled=false;
    },3000);
  }
}

function basActualizarPeso(){
  const bruto=parseFloat(document.getElementById('bas-peso-input').value)||0;
  const tara=basSelCamion?basSelCamion.tara||0:0;
  const neto=bruto-tara;
  document.getElementById('bas-bruto-display').innerHTML=bruto.toLocaleString()+' <span style="font-size:1.2rem;opacity:.8">Kg</span>';
  document.getElementById('hdr-bruto').textContent=bruto>0?bruto.toLocaleString()+' kg':'—';
  document.getElementById('hdr-neto').textContent=neto>0?neto.toLocaleString()+' kg':'—';
  checkBasStep1();
}

function checkBasStep1(){
  const bruto=parseFloat(document.getElementById('bas-peso-input').value)||0;
  document.getElementById('bas-btn-continuar').disabled=!(basSelCamion&&bruto>0);
}

let basCurrentStep=1;

function basGoStepTab(n){
  // Solo permitir ir a pasos ya completados o al actual
  if(n>basCurrentStep) return;
  basGoStep(n);
}

function cancelarPesada(){
  if(!confirm('¿Cancelar pesada en curso? Se perderán los datos no guardados.')) return;
  basCurrentStep=1;
  basSelCamion=null;
  basNumPedido=null;
  basSelCliente=null;
  document.getElementById('bas-peso-input').value='';
  basActualizarPeso();
  document.getElementById('hdr-bruto').textContent='—';
  document.getElementById('hdr-cam').textContent='—';
  document.getElementById('hdr-rem').textContent='—';
  document.getElementById('hdr-tara').textContent='—';
  document.getElementById('hdr-neto').textContent='—';
  basGoStep(1);
  goPage('inicio');
}

function basGoStep(n){
  if(n>basCurrentStep) basCurrentStep=n;
  [1,2,3].forEach(i=>{
    const el=document.getElementById('bas-step-'+i);
    if(el)el.style.display=i===n?'block':'none';
  });
  // Update step tabs
  [1,2,3].forEach(i=>{
    const tab=document.getElementById('stab'+i);
    const num=document.getElementById('snum'+i);
    const title=document.getElementById('stab'+i+'-title');
    if(!tab)return;
    if(i===n){
      tab.style.background='var(--accent)';tab.style.opacity='1';
      num.style.background='rgba(0,0,0,.3)';num.style.color='#fff';
      title.style.color='#fff';
    }else if(i<n||i<=basCurrentStep){
      tab.style.background='var(--surface2)';tab.style.opacity='1';
      num.style.background='var(--accent2)';num.style.color='#fff';
      title.style.color='var(--accent2)';
      tab.style.cursor='pointer';
    }else{
      tab.style.background='var(--surface2)';tab.style.opacity='.5';
      num.style.background='var(--border)';num.style.color='var(--muted)';
      title.style.color='var(--muted)';
      tab.style.cursor='default';
    }
  });
  const labels={1:'PESADA',2:'CLIENTE',3:'ALBARÁN'};
  document.getElementById('bas-step-label').textContent=labels[n];
  const hdrGuardar=document.getElementById('hdr-btn-guardar');
  if(hdrGuardar) hdrGuardar.style.display=n===3?'inline-block':'none';

  if(n===2){
    // Only reset client/pedido when navigating FORWARD (new weighing, not going back from step 3)
    const comingForward=n>basCurrentStep||(n===2&&basCurrentStep<=2);
    if(!basNumPedido){
      // No active pedido — always reset UI
      document.getElementById('bas-btn-s2').disabled=true;
      document.getElementById('bas-pedido-generado').style.display='none';
      basSelCliente=null;
      document.getElementById('bas-cli-text').textContent='Seleccionar cliente...';
      document.getElementById('bas-cli-text').style.color='var(--muted)';
    }
    cargarPedidosHoy();
  }
  if(n===3) renderAlbaranStep3();
}


async function cargarPedidosHoy(){
  const listEl=document.getElementById('bas-pedidos-hoy-list');
  if(!listEl)return;
  listEl.innerHTML='<div style="color:var(--muted);font-size:1rem;padding:8px 0">Cargando...</div>';

  // Cargar pedidos BC abiertos en paralelo con Supabase
  const hoy=dateStr(new Date());
  const [json, bcPedidos] = await Promise.all([
    apiFetch('?accion=pedidos&dias=2').catch(()=>({ok:false,data:[]})),
    _cargarPedidosAbiertosBC().catch(()=>[])
  ]);

  // Pedidos de Supabase de hoy
  const seen=new Set();
  const pedidosHoy=[];
  if(json.ok){
    json.data.forEach(r=>{
      if(!r.numPedido||seen.has(r.numPedido))return;
      const d=parseFechaHoraObj(r.fechaHora)||parseFechaHoraObj(r.fechaPedido);
      if(d&&dateStr(d)===hoy){
        seen.add(r.numPedido);
        const numLineas=json.data.filter(x=>x.numPedido===r.numPedido).length;
        pedidosHoy.push({numPedido:r.numPedido,nombreCliente:r.nombreCliente,numLineas,fuente:'local'});
      }
    });
  }

  // Pedidos BC abiertos (no duplicar si ya está en Supabase)
  bcPedidos.forEach(p=>{
    if(!seen.has(p.numPedido)){
      seen.add(p.numPedido);
      pedidosHoy.push(p);
    }
  });

  if(!pedidosHoy.length){
    listEl.innerHTML='<div style="color:var(--muted);font-size:1rem;padding:8px 0">Sin pedidos hoy — crea una nueva cabecera</div>';
    return;
  }
  listEl.innerHTML='';
  pedidosHoy.forEach(p=>{
    const div=document.createElement('div');
    div.style.cssText='background:var(--surface2);border:1.5px solid var(--border);border-radius:10px;padding:14px 16px;margin-bottom:8px;cursor:pointer;transition:border-color .2s';
    const esBC=p.fuente==='bc';
    div.innerHTML=`<div style="display:flex;justify-content:space-between;align-items:center">
      <div>
        <div style="font-family:'DM Mono',monospace;font-weight:700;font-size:1.1rem;color:var(--accent)">${p.numPedido}</div>
        <div style="font-size:.92rem;color:var(--muted);margin-top:3px">${p.nombreCliente}${p.numLineas?' · '+p.numLineas+' líneas':''}</div>
      </div>
      <div style="font-size:.78rem;font-weight:700;color:${esBC?'#1a5faa':'var(--accent2)'}>${esBC?'BC Abierto':'Abierto'}</div>
    </div>`;
    div.onmouseover=()=>div.style.borderColor='var(--accent)';
    div.onmouseout=()=>div.style.borderColor='var(--border)';
    div.onclick=()=>seleccionarPedidoExistente(p.numPedido, p.nombreCliente, json.data||[], p.fuente);
    listEl.appendChild(div);
  });
}

async function _cargarPedidosAbiertosBC(){
  try{
    if(typeof getBCToken!=='function')return[];
    const token=await getBCToken();
    const base=`https://api.businesscentral.dynamics.com/v2.0/${BC_TENANT}/${BC_ENV}/api/v2.0/companies`;
    const headers={'Authorization':`Bearer ${token}`};
    if(!window._bcCompanyId){
      const cJson=await(await fetch(base,{headers})).json();
      const company=(cJson.value||[]).find(c=>c.name===BC_COMPANY);
      if(!company)return[];
      window._bcCompanyId=company.id;
    }
    const hoy=new Date().toISOString().slice(0,10);
    const filter=`status eq 'Draft' and orderDate eq ${hoy}`;
    const url=`${base}(${window._bcCompanyId})/salesOrders?$filter=${encodeURIComponent(filter)}&$select=number,customerName,orderDate,status&$orderby=number desc`;
    const res=await fetch(url,{headers});
    if(!res.ok)return[];
    const json=await res.json();
    return(json.value||[]).map(o=>({
      numPedido:o.number,
      nombreCliente:o.customerName||'',
      numLineas:0,
      fuente:'bc'
    }));
  }catch(e){console.warn('BC pedidos abiertos:',e.message);return[];}
}

async function seleccionarPedidoExistente(numPedido, nombreCliente, allData, fuente){
  basNumPedido=numPedido;
  // Find client
  const cli=CLIENTES.find(c=>c.nombre.toUpperCase()===nombreCliente.toUpperCase());
  basSelCliente={nombre:nombreCliente, codigo:cli?cli.codigo:''};
  // Calculate next linea and load existing lines
  let lineasData=allData.filter(r=>r.numPedido===numPedido&&Number(r.numLinea||0)>0).sort((a,b)=>Number(a.numLinea)-Number(b.numLinea));

  // Si es pedido BC y no hay líneas en local, cargar de BC
  if(fuente==='bc'&&lineasData.length===0){
    try{
      if(typeof getBCToken==='function'){
        const token=await getBCToken();
        const base=`https://api.businesscentral.dynamics.com/v2.0/${BC_TENANT}/${BC_ENV}/api/v2.0/companies`;
        const headers={'Authorization':`Bearer ${token}`};
        if(!window._bcCompanyId){
          const cJson=await(await fetch(base,{headers})).json();
          const company=(cJson.value||[]).find(c=>c.name===BC_COMPANY);
          if(company)window._bcCompanyId=company.id;
        }
        if(window._bcCompanyId){
          const url=`${base}(${window._bcCompanyId})/salesOrderLines?$filter=documentNumber eq '${numPedido}'&$select=documentNumber,sequence,quantity,unitPrice,lineType&$orderby=sequence`;
          const res=await fetch(url,{headers});
          if(res.ok){
            const json=await res.json();
            lineasData=(json.value||[]).map((l,i)=>({
              numPedido:l.documentNumber,
              numLinea:(i+1)*10000,
              pesoNeto:l.quantity||0
            }));
          }
        }
      }
    }catch(e){console.warn('BC líneas:',e.message);}
  }

  const lineasNums=lineasData.map(r=>Number(r.numLinea));
  basCurrentLinea=lineasNums.length>0?Math.max(...lineasNums)+10000:10000;
  basLineasSesion=lineasData.map(r=>({
    numLinea:Number(r.numLinea),
    matriculacam:r.matriculacam||'',
    matricularem:r.matricularem||'',
    pesoNeto:r.pesoNeto,
    productoCod:r.productoCod||'',
    productoNombre:r.productoNombre||'',
    proyectoName:r.proyectoName||'',
  }));
  // Update obras dropdown for this client
  const obras=_getObrasCliente(nombreCliente,cli?cli.codigo:'');
  const sel=document.getElementById('bas-obra-sel');
  if(sel)sel.innerHTML='<option value="">Seleccionar obra...</option>'+obras.map(o=>`<option value="${o.codigo}">${o.codigo} — ${o.nombre}</option>`).join('');
  // Show selected
  document.getElementById('bas-cli-text').textContent=nombreCliente;
  document.getElementById('bas-cli-text').style.color='var(--text)';
  document.getElementById('bas-num-pedido').textContent=numPedido;
  document.getElementById('bas-pedido-estado').textContent='(línea '+basCurrentLinea+')';
  document.getElementById('bas-pedido-generado').style.display='block';
  document.getElementById('bas-btn-s2').disabled=false;
  // Highlight selected card
  document.querySelectorAll('#bas-pedidos-hoy-list > div').forEach(d=>d.style.borderColor='var(--border)');
  event.currentTarget.style.borderColor='var(--accent2)';
}

// CLIENTE — Favoritos
function _getFavClientes(){try{return JSON.parse(localStorage.getItem('cli_favs')||'[]');}catch(e){return[];}}
function _setFavClientes(arr){localStorage.setItem('cli_favs',JSON.stringify(arr));}
function _isCliFav(nombre){return _getFavClientes().includes(nombre);}
function toggleCliFav(nombre,e){
  if(e)e.stopPropagation();
  let favs=_getFavClientes();
  if(favs.includes(nombre))favs=favs.filter(n=>n!==nombre);
  else favs.unshift(nombre);
  _setFavClientes(favs);
  const q=document.getElementById('bas-cli-search');
  renderCliDropdown(q?q.value:'');
}

function toggleCliDropdown(){
  const dd=document.getElementById('bas-cli-dropdown');
  const isOpen=dd.style.display==='flex';
  if(isOpen){dd.style.display='none';return;}
  // Position below the display button
  const btn=document.getElementById('bas-cli-display');
  const r=btn.getBoundingClientRect();
  dd.style.top=(r.bottom+4)+'px';
  dd.style.display='flex';
  document.getElementById('bas-cli-search').value='';
  document.getElementById('bas-cli-search').focus();
  renderCliDropdown('');
}

function _renderCliRow(c,container){
  const div=document.createElement('div');
  div.style.cssText='padding:12px 14px;cursor:pointer;font-size:1.05rem;border-bottom:1px solid var(--border);display:flex;align-items:center;justify-content:space-between;gap:8px';
  const nameSpan=document.createElement('span');
  nameSpan.textContent=c.nombre;
  nameSpan.style.cssText='flex:1;font-weight:500';
  const star=document.createElement('span');
  star.textContent=_isCliFav(c.nombre)?'★':'☆';
  star.style.cssText='font-size:1.3rem;color:'+(_isCliFav(c.nombre)?'var(--accent)':'var(--muted)')+';cursor:pointer;padding:0 4px;flex-shrink:0';
  star.onclick=(e)=>toggleCliFav(c.nombre,e);
  div.appendChild(nameSpan);
  div.appendChild(star);
  div.onmouseover=()=>div.style.background='var(--surface2)';
  div.onmouseout=()=>div.style.background='transparent';
  div.onclick=()=>seleccionarCliente(c.nombre,c.codigo);
  container.appendChild(div);
}

function renderCliDropdown(q){
  const list=document.getElementById('bas-cli-list');if(!list)return;
  const filtered=q?CLIENTES.filter(c=>c.nombre.toUpperCase().includes(q.toUpperCase())):CLIENTES;
  list.innerHTML='';
  const favs=_getFavClientes();
  // Favoritos section
  const favClientes=filtered.filter(c=>favs.includes(c.nombre)).sort((a,b)=>favs.indexOf(a.nombre)-favs.indexOf(b.nombre));
  const noFavClientes=filtered.filter(c=>!favs.includes(c.nombre));
  if(favClientes.length&&!q){
    const hdr=document.createElement('div');
    hdr.style.cssText='padding:8px 14px;font-size:.75rem;font-weight:700;color:var(--accent);text-transform:uppercase;letter-spacing:.5px;background:var(--surface2)';
    hdr.textContent='★ Favoritos';
    list.appendChild(hdr);
    favClientes.forEach(c=>_renderCliRow(c,list));
    if(noFavClientes.length){
      const sep=document.createElement('div');
      sep.style.cssText='padding:8px 14px;font-size:.75rem;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.5px;background:var(--surface2);border-top:2px solid var(--border)';
      sep.textContent='Todos los clientes';
      list.appendChild(sep);
    }
  }
  (q?filtered:noFavClientes).forEach(c=>_renderCliRow(c,list));
}

function filtrarClientes(){
  const q=document.getElementById('bas-cli-search').value;
  renderCliDropdown(q);
}

function seleccionarCliente(nombre,codigo){
  basSelCliente={nombre,codigo};
  basNumPedido=null;
  basLineasSesion=[];
  basCurrentLinea=10000;
  // Update UI
  const cliText=document.getElementById('bas-cli-text');
  if(cliText){cliText.textContent=nombre;cliText.style.color='var(--text)';}
  const dd=document.getElementById('bas-cli-dropdown');
  if(dd)dd.style.display='none';
  const pedGen=document.getElementById('bas-pedido-generado');
  if(pedGen)pedGen.style.display='none';
  const btnS2=document.getElementById('bas-btn-s2');
  if(btnS2)btnS2.disabled=true;
  // Load obras
  const obras=_getObrasCliente(nombre,codigo);
  const obraSel=document.getElementById('bas-obra-sel');
  if(obraSel)obraSel.innerHTML='<option value="">Seleccionar obra...</option>'+obras.map(o=>`<option value="${o.codigo}">${o.codigo} — ${o.nombre}</option>`).join('');
  // Show nueva cab button
  const btn=document.getElementById('bas-btn-nueva-cab');
  if(btn)btn.style.display='block';
}

document.addEventListener('click',function(e){
  const dd=document.getElementById('bas-cli-dropdown');
  if(dd&&!e.target.closest('#bas-cli-display')&&!e.target.closest('#bas-cli-dropdown')){
    dd.style.display='none';
  }
});

async function crearCabeceraPedido(){
  if(!basSelCliente){alert('Selecciona un cliente primero.');return;}
  try{
    const json=await apiFetch('?accion=pedidos&dias=365');
    let lastNum=817;
    if(json.ok&&json.data.length){
      const nums=json.data
        .map(r=>parseInt(String(r.numPedido||'').split('-').pop()))
        .filter(n=>!isNaN(n)&&n>0);
      if(nums.length)lastNum=Math.max(...nums);
    }
    const year=new Date().getFullYear().toString().slice(2);
    basNumPedido='PEDV'+year+'-'+String(lastNum+1).padStart(6,'0');
  }catch(e){
    const year=new Date().getFullYear().toString().slice(2);
    basNumPedido='PEDV'+year+'-000001';
  }
  basLineasSesion=[];
  basCurrentLinea=10000;
  document.getElementById('bas-num-pedido').textContent=basNumPedido;
  document.getElementById('bas-pedido-estado').textContent='(nuevo pedido)';
  document.getElementById('bas-pedido-generado').style.display='block';
  document.getElementById('bas-btn-nueva-cab').style.display='none';
  document.getElementById('bas-btn-s2').disabled=false;
}

function renderAlbaranStep3(){
  if(!basSelCamion||!basNumPedido)return;
  const bruto=parseFloat(document.getElementById('bas-peso-input').value)||0;
  const tara=basSelCamion.tara||0;
  const neto=bruto-tara;
  document.getElementById('alb-header-num').textContent=basNumPedido+'/'+basCurrentLinea;
  document.getElementById('alb-mat2').textContent=basSelCamion.matriculacam;
  document.getElementById('alb-rem2').textContent=basSelCamion.matricularem||'—';
  document.getElementById('alb-cond').textContent=basSelCamion.chofer||'—';
  document.getElementById('alb-bruto2').textContent=bruto.toLocaleString();
  document.getElementById('alb-tara2').textContent=tara.toLocaleString();
  document.getElementById('alb-neto2').textContent=(neto>0?neto:0).toLocaleString();
  document.getElementById('alb-numpedido').textContent=basNumPedido;
  document.getElementById('alb-linea-step3').textContent=basCurrentLinea;
  document.getElementById('alb-cliente2').textContent=basSelCliente?basSelCliente.nombre:'—';
  document.getElementById('alb-codcli').textContent=basSelCliente?basSelCliente.codigo:'—';
  renderLineasAlbaran();
}

async function eliminarLineaSesion(idx){
  const l=basLineasSesion[idx];
  if(!l)return;
  if(!confirm('¿Eliminar línea '+l.numLinea+' · '+l.matriculacam+' · '+Number(l.pesoNeto).toLocaleString()+' kg?'))return;
  // Borrar de Supabase si tiene ID real
  const lid=l._id;
  if(lid&&lid!=='LOCAL'&&lid!=='null'&&!isNaN(Number(lid))){
    try{ await dbQuery({ action: 'delete', table: 'tblpedidos', filters: [{ column: 'id', op: 'eq', value: Number(lid) }] }); }catch(e){ console.warn('Error borrando línea:',e); }
  }
  basLineasSesion.splice(idx,1);
  renderLineasAlbaran();
}

function renderLineasAlbaran(){
  const el=document.getElementById('alb-lineas-list');
  if(!basLineasSesion.length){
    el.innerHTML='<div style="padding:14px;text-align:center;font-size:.78rem;color:var(--muted)">Sin líneas aún</div>';
    return;
  }
  el.innerHTML=basLineasSesion.map((l,i)=>`
    <div style="display:flex;gap:8px;padding:8px 10px;border-bottom:1px solid var(--border);font-size:.82rem;color:var(--text);align-items:center">
      <div style="flex:.6;font-family:monospace;color:var(--muted)">${l.numLinea}</div>
      <div style="flex:1;font-family:monospace;font-weight:700;color:var(--accent)">${l.matriculacam}</div>
      <div style="flex:.8;font-family:monospace">${l.matricularem||'—'}</div>
      <div style="flex:1;text-align:right;font-family:monospace;font-weight:700">${Number(l.pesoNeto).toLocaleString()}</div>
      <div style="flex:1.5;font-size:.75rem">${l.productoNombre||''}</div>
      <div style="flex:.5;text-align:right">
        <button onclick="eliminarLineaSesion(${i})" style="background:transparent;border:1.5px solid #e05;color:#e05;border-radius:6px;padding:4px 8px;cursor:pointer;font-size:.8rem;font-weight:700" title="Eliminar línea">✕</button>
      </div>
    </div>`).join('');
}

// Update obra name when selected
document.addEventListener('change',function(e){
  if(e.target.id==='bas-obra-sel'){
    const cod=e.target.value;
    const obras=(basSelCliente&&basSelCliente.nombre)?_getObrasCliente(basSelCliente.nombre,basSelCliente.codigo):[];
    const obra=obras.find(o=>o.codigo===cod);
    document.getElementById('alb-obra-nombre').textContent=obra?'Obra: '+obra.nombre:'';
  }
  if(e.target.id==='bas-producto-sel'){
    // no extra action needed
  }
});

async function guardarLinea(){
  const prodCod=document.getElementById('bas-producto-sel').value;
  const obraCod=document.getElementById('bas-obra-sel').value;
  if(!prodCod){alert('Selecciona un producto.');return;}
  if(!obraCod){if(!confirm('¿Seguro que quieres añadir un albarán sin obra?'))return;}
  if(!basNumPedido){alert('Crea primero la cabecera del pedido.');return;}

  const bruto=parseFloat(document.getElementById('bas-peso-input').value)||0;
  const tara=basSelCamion.tara||0;
  const neto=bruto-tara>0?bruto-tara:0;
  const prodNombre=PROD_MAP[prodCod]||prodCod;
  const obras=(basSelCliente&&basSelCliente.nombre)?_getObrasCliente(basSelCliente.nombre,basSelCliente.codigo):[];
  const obra=obras.find(o=>o.codigo===obraCod);
  const obraNombre=obra?obra.nombre:'';
  const fechaPedido=document.getElementById('bas-fecha-pedido').value;

  const payload={
    tipo:'pesada',
    matriculacam:basSelCamion.matriculacam,
    matricularem:basSelCamion.matricularem||'',
    tara,
    chofer:basSelCamion.chofer||'',
    cliente:basSelCliente?basSelCliente.codigo:'',
    codigoCliente:basSelCliente?basSelCliente.codigo:'',
    nombreCliente:basSelCliente?basSelCliente.nombre:'',
    fechaPedido,
    productoCod:prodCod,
    productoNombre:prodNombre,
    pesoBruto:bruto,
    pesoNeto:neto,
    proyectoCod:obraCod,
    proyectoName:obraNombre,
    numPedido:basNumPedido,
    numLinea:basCurrentLinea,
    observaciones:(document.getElementById('bas-observaciones').value||'').trim(),
  };

  const btn=document.querySelector('#bas-step-3 .btn-pri');
  if(btn){btn.disabled=true;btn.textContent='Guardando...';}
  try{
    const json=await apiPost(payload);
    if(!json.ok){
      alert('Error guardando pesada: '+(json.error||'Error desconocido')+'\n\nReintenta o recarga la página.');
      return;
    }
    const savedId=json.id!=null?json.id:Date.now();
    basLineasSesion.push({
      numLinea:basCurrentLinea,
      matriculacam:basSelCamion.matriculacam,
      matricularem:basSelCamion.matricularem||'',
      pesoNeto:neto,
      productoCod:prodCod,
      productoNombre:prodNombre,
      _id:savedId,
      _payload:payload,
    });
    basCurrentLinea+=10000;
    mostrarExitoLinea(savedId,payload);
    renderAlbaranStep3();
  }catch(e){
    alert('Error de conexión guardando pesada: '+e.message+'\n\nReintenta o recarga la página.');
    return;
  }finally{
    if(btn){btn.disabled=false;btn.textContent='Guardar';}
  }
}

// Store last saved line for albaran button (avoids inline JSON in onclick)
let _lastSavedId=null, _lastSavedPayload=null;

function mostrarExitoLinea(id,payload){
  _lastSavedId=id; _lastSavedPayload=payload;
  let banner=document.getElementById('linea-ok-banner');
  if(!banner){
    banner=document.createElement('div');
    banner.id='linea-ok-banner';
    banner.style.cssText='background:rgba(107,125,46,.12);border:1px solid rgba(107,125,46,.4);border-radius:8px;padding:10px 14px;margin-bottom:10px;display:flex;align-items:center;justify-content:space-between;gap:8px';
    const step3=document.getElementById('bas-step-3');
    if(step3)step3.insertBefore(banner,step3.firstChild);
  }
  const neto=Number(payload.pesoNeto).toLocaleString();
  banner.innerHTML=`<span style="font-size:.82rem;color:var(--accent2);font-weight:700">✓ Guardado — ${payload.productoNombre} · ${neto} kg</span>`+
    `<button onclick="mostrarAlbaranUltimaLinea()" style="padding:5px 12px;background:var(--accent);color:#fff;border:none;border-radius:6px;font-size:.72rem;font-weight:700;cursor:pointer;white-space:nowrap">🖨 Albarán</button>`;
}
function mostrarAlbaranUltimaLinea(){
  try{
  if(_lastSavedId===null||!_lastSavedPayload){alert('Sin datos guardados aún');return;}
  // Mostrar albarán inmediatamente sin esperar BC
  const p=_lastSavedPayload;
  const id=_lastSavedId;
  const now=new Date();
  const fecha=pad(now.getDate())+'/'+pad(now.getMonth()+1)+'/'+now.getFullYear()+' '+pad(now.getHours())+':'+pad(now.getMinutes());
  const idNum=(id&&id!=='LOCAL'&&!isNaN(Number(id)))?String(id).padStart(6,'0'):'PEND';
  document.getElementById('alb-num').textContent='PEDV'+now.getFullYear()+'-'+idNum+'/'+String(p.numLinea||0).padStart(5,'0');
  document.getElementById('alb-fecha').textContent=fecha;
  document.getElementById('alb-mat').textContent=p.matriculacam||'—';
  document.getElementById('alb-rem').textContent=p.matricularem||'—';
  document.getElementById('alb-tara-hdr').textContent=Number(p.tara).toLocaleString();
  document.getElementById('alb-chofer').textContent=(p.chofer||'—').toUpperCase();
  document.getElementById('alb-cli-cod').textContent=p.codigoCliente?'COD.CLIENTE: '+p.codigoCliente:'';
  document.getElementById('alb-cliente').textContent=p.nombreCliente||'—';
  document.getElementById('alb-cif').textContent='';
  document.getElementById('alb-dir1').textContent='';
  document.getElementById('alb-dir2').textContent='';
  document.getElementById('alb-obra-cod').textContent=p.codigoProyecto?'COD.OBRA: '+p.codigoProyecto:'';
  document.getElementById('alb-obra').textContent=p.proyectoName?'OBRA: '+p.proyectoName:'—';
  document.getElementById('alb-linea').textContent=p.numLinea||'—';
  document.getElementById('alb-cod-prod').textContent=p.productoCod||'—';
  document.getElementById('alb-producto').textContent=p.productoNombre||'—';
  document.getElementById('alb-bruto').textContent=Number(p.pesoBruto).toLocaleString();
  document.getElementById('alb-neto').textContent=Number(p.pesoNeto).toLocaleString();
  // Observaciones
  const obsWrap2=document.getElementById('alb-obs-wrap');
  const obsText2=document.getElementById('alb-obs-text');
  if(obsWrap2&&obsText2){
    const obs2=p.observaciones||'';
    if(obs2){obsText2.textContent=obs2;obsWrap2.style.display='block';}
    else{obsWrap2.style.display='none';}
  }
  const nombre=(p.productoNombre||'').toUpperCase();
  const esCE=/\b(0\/4|4\/12|12\/20)\b/.test(nombre);
  const ceImg=document.getElementById('alb-ce-img');
  if(ceImg)ceImg.style.display=esCE?'inline':'none';
  _renderAlbaranQR(p.productoNombre);
  const aw=document.getElementById('albaran-wrap');
  aw.style.display='flex';
  aw.classList.add('print-active');
  const notasBtn=document.getElementById('btn-notas-float');
  if(notasBtn)notasBtn.style.display='none';
  // BC en segundo plano
  if(p.codigoCliente){
    _cargarDatosFiscalesBC(p.codigoCliente).then(d=>{
      if(!d)return;
      document.getElementById('alb-cif').textContent=d.cif?'CIF: '+d.cif:'';
      document.getElementById('alb-dir1').textContent=d.dir1?'DIRECCIÓN: '+d.dir1:'';
      document.getElementById('alb-dir2').textContent=d.dir2||'';
    });
  }
  }catch(e){alert('Error albarán: '+e.message);console.error(e);}
}

// Cache datos fiscales clientes BC (codigo → {cif, dir1, dir2})
const _bcClienteCache={};

async function _cargarDatosFiscalesBC(codigoCliente){
  if(!codigoCliente||codigoCliente==='CLI-00000')return null;
  if(_bcClienteCache[codigoCliente])return _bcClienteCache[codigoCliente];
  try{
    if(typeof getBCToken!=='function')return null;
    const token=await getBCToken();
    const base=`https://api.businesscentral.dynamics.com/v2.0/${BC_TENANT}/${BC_ENV}/api/v2.0/companies`;
    const headers={'Authorization':`Bearer ${token}`};
    if(!window._bcCompanyId){
      const cJson=await(await fetch(base,{headers})).json();
      const company=(cJson.value||[]).find(c=>c.name===BC_COMPANY);
      if(!company)return null;
      window._bcCompanyId=company.id;
    }
    const url=`${base}(${window._bcCompanyId})/customers?$filter=number eq '${codigoCliente}'`;
    const res=await fetch(url,{headers});
    if(!res.ok)return null;
    const json=await res.json();
    const c=(json.value||[])[0];
    if(!c)return null;
    const datos={
      cif:c.taxRegistrationNumber||c.vatRegistrationNo||c.taxLiable||'',
      dir1:c.addressLine1||c.address||'',
      dir2:[c.postCode||c.postalCode,c.city].filter(Boolean).join(' ')
    };
    _bcClienteCache[codigoCliente]=datos;
    return datos;
  }catch(e){console.warn('BC cliente fiscal:',e.message);return null;}
}

async function mostrarAlbaran(id,p){
  const now=new Date();
  const fecha=pad(now.getDate())+'/'+pad(now.getMonth()+1)+'/'+now.getFullYear()+' '+pad(now.getHours())+':'+pad(now.getMinutes());
  const idNum=(id&&id!=='LOCAL'&&!isNaN(Number(id)))?String(id).padStart(6,'0'):'PEND';
  document.getElementById('alb-num').textContent='PEDV'+now.getFullYear()+'-'+idNum+'/'+String(p.numLinea||0).padStart(5,'0');
  document.getElementById('alb-fecha').textContent=fecha;
  document.getElementById('alb-mat').textContent=p.matriculacam||'—';
  document.getElementById('alb-rem').textContent=p.matricularem||'—';
  document.getElementById('alb-tara-hdr').textContent=Number(p.tara).toLocaleString();
  document.getElementById('alb-chofer').textContent=(p.chofer||'—').toUpperCase();
  // Cliente — datos básicos inmediatos
  document.getElementById('alb-cli-cod').textContent=p.codigoCliente?'COD.CLIENTE: '+p.codigoCliente:'';
  document.getElementById('alb-cliente').textContent=p.nombreCliente||'—';
  document.getElementById('alb-cif').textContent='';
  document.getElementById('alb-dir1').textContent='';
  document.getElementById('alb-dir2').textContent='';
  // Obra
  document.getElementById('alb-obra-cod').textContent=p.codigoProyecto?'COD.OBRA: '+p.codigoProyecto:'';
  document.getElementById('alb-obra').textContent=p.proyectoName?'OBRA: '+p.proyectoName:'—';
  // Línea
  document.getElementById('alb-linea').textContent=p.numLinea||'—';
  document.getElementById('alb-cod-prod').textContent=p.productoCod||'—';
  document.getElementById('alb-producto').textContent=p.productoNombre||'—';
  document.getElementById('alb-bruto').textContent=Number(p.pesoBruto).toLocaleString();
  document.getElementById('alb-neto').textContent=Number(p.pesoNeto).toLocaleString();
  // Observaciones
  const obsWrap=document.getElementById('alb-obs-wrap');
  const obsText=document.getElementById('alb-obs-text');
  if(obsWrap&&obsText){
    const obs=p.observaciones||'';
    if(obs){obsText.textContent=obs;obsWrap.style.display='block';}
    else{obsWrap.style.display='none';}
  }
  // CE logo
  const nombre=(p.productoNombre||'').toUpperCase();
  const esCE=/\b(0\/4|4\/12|12\/20)\b/.test(nombre);
  const ceImg=document.getElementById('alb-ce-img');
  if(ceImg)ceImg.style.display=esCE?'inline':'none';
  _renderAlbaranQR(p.productoNombre);
  // Mostrar albarán
  const aw=document.getElementById('albaran-wrap');
  if(aw) {
    aw.style.display='flex';
    aw.style.position='fixed';
    aw.classList.add('print-active');
    const notasBtn=document.getElementById('btn-notas-float');
    if(notasBtn)notasBtn.style.display='none';
    setTimeout(()=>window.print(),100);
  }
  // Cargar datos fiscales BC en segundo plano (deshabilitado por error MSAL)
  /*if(p.codigoCliente){
    _cargarDatosFiscalesBC(p.codigoCliente).then(d=>{
      if(!d)return;
      document.getElementById('alb-cif').textContent=d.cif?'CIF: '+d.cif:'';
      document.getElementById('alb-dir1').textContent=d.dir1?'DIRECCIÓN: '+d.dir1:'';
      document.getElementById('alb-dir2').textContent=d.dir2||'';
    });
  }*/
}
function cerrarAlbaran(){
  const aw=document.getElementById('albaran-wrap');
  aw.style.display='none';
  aw.classList.remove('print-active');
  const notasBtn=document.getElementById('btn-notas-float');
  if(notasBtn)notasBtn.style.display='flex';
  // Limpiar peso anterior para siguiente pesada
  document.getElementById('bas-peso-input').value='';
  basActualizarPeso();
}



// ── PEDIDOS ───────────────────────────────────────────────────
let pedidosData=[];
async function cargarPedidos(){
  const el=document.getElementById('pedidos-list');
  el.innerHTML='<div class="tbl"><div class="empty">Cargando...</div></div>';
  try{
    const json=await apiFetch('?accion=pedidos&dias=90');
    if(!json.ok)throw new Error(json.error);
    pedidosData=json.data;
    filtrarPedidos();
  }catch(e){el.innerHTML='<div class="tbl"><div class="empty">Error: '+e.message+'</div></div>';}
}
function filtrarPedidos(){
  const fm=(document.getElementById('filt-ped-mat').value||'').toUpperCase();
  const fc=(document.getElementById('filt-ped-cli').value||'').toUpperCase();
  const ff=document.getElementById('filt-ped-fecha').value;
  let data=pedidosData;
  if(fm)data=data.filter(r=>String(r.matriculacam).toUpperCase().includes(fm));
  if(fc)data=data.filter(r=>String(r.nombreCliente).toUpperCase().includes(fc));
  if(ff)data=data.filter(r=>{
    const fh=String(r.fechaHora||'');
    // Try ISO format: 2026-03-20T16:33:00.000Z
    if(fh.includes('T')){
      return fh.slice(0,10)===ff;
    }
    // Try dd/mm/yy format
    const parts=fh.split(' ')[0].split('/');
    if(parts.length<3)return false;
    const yr=parts[2].length===2?'20'+parts[2]:parts[2];
    const d=new Date(yr,parseInt(parts[1])-1,parseInt(parts[0]));
    return dateStr(d)===ff;
  });
  renderPedidos(data);
}
function formatFechaHoraPed(fh){
  if(!fh) return '—';
  // ISO: 2026-03-27T23:01:00.000Z
  if(typeof fh==='string' && fh.includes('T')){
    const d=new Date(fh);
    if(!isNaN(d)) return pad(d.getDate())+'/'+pad(d.getMonth()+1)+'/'+String(d.getFullYear()).slice(2)+' '+pad(d.getHours())+':'+pad(d.getMinutes());
  }
  // ya formateado dd/MM/yy HH:mm
  return String(fh).substring(0,14);
}
function renderPedidos(data){
  const el=document.getElementById('pedidos-list');
  if(!data.length){el.innerHTML='<div class="tbl"><div class="empty">Sin resultados</div></div>';return;}
  el.innerHTML='<div class="tbl"><div class="tr th">'+
    '<div class="tc" style="flex:.35">#</div>'+
    '<div class="tc" style="flex:.8">Mat.</div>'+
    '<div class="tc" style="flex:1.2">Cliente</div>'+
    '<div class="tc" style="flex:1">Producto</div>'+
    '<div class="tc" style="flex:.65;text-align:right">Neto</div>'+
    '<div class="tc" style="flex:.75">Fecha</div>'+
    '<div class="tc" style="flex:.7;text-align:right">Acciones</div>'+
    '</div>'+
    data.map(r=>'<div class="tr">'+
      '<div class="tc" style="flex:.35;font-family:monospace;color:var(--muted);font-size:.85rem">#'+r.id+'</div>'+
      '<div class="tc" style="flex:.8;font-family:monospace;font-weight:700;color:var(--accent);font-size:.9rem">'+r.matriculacam+'</div>'+
      '<div class="tc" style="flex:1.2;color:var(--muted);font-size:.88rem">'+r.nombreCliente+'</div>'+
      '<div class="tc" style="flex:1;font-size:.85rem">'+r.productoNombre+'</div>'+
      '<div class="tc" style="flex:.65;text-align:right;font-family:monospace;color:var(--accent2);font-weight:700;font-size:.9rem">'+Number(r.pesoNeto).toLocaleString()+'</div>'+
      '<div class="tc" style="flex:.75;font-size:.85rem;color:var(--muted)">'+formatFechaHoraPed(r.fechaHora)+'</div>'+
      '<div class="tc" style="flex:.7;text-align:right;gap:4px;display:flex;justify-content:flex-end">'+
        '<button class="btn-sm" title="Albarán" onclick=\'abrirAlbaranPedido('+r.id+')\' style="padding:4px 8px;font-size:.8rem;color:var(--accent);border-color:var(--accent)">🖨</button>'+
        '<button class="btn-sm" title="Editar" onclick=\'editarPedidoModal('+r.id+')\' style="padding:4px 8px;font-size:.8rem">✏</button>'+
        '<button class="btn-sm" title="Eliminar" onclick=\'eliminarPedido('+r.id+')\' style="padding:4px 8px;font-size:.8rem;color:#e05;border-color:#e05">🗑</button>'+
      '</div>'+
    '</div>').join('')+
  '</div>';
}
function abrirAlbaranPedido(id){
  const r=pedidosData.find(x=>x.id==id);
  if(!r) return;
  mostrarAlbaran(r.id, r);
}
async function eliminarPedido(id){
  if(!id||id==='null'||id==='LOCAL')return;
  const r=pedidosData.find(x=>x.id==id);
  if(!r) return;
  const confirmMsg=`¿Eliminar pedido #${id}?\n${r.matriculacam} · ${r.productoNombre} · ${Number(r.pesoNeto).toLocaleString()} kg\n${r.nombreCliente||'Sin cliente'}`;
  if(!confirm(confirmMsg)) return;
  const json=await apiPost({tipo:'deletePedido',id:Number(id)});
  if(!json.ok){alert('Error: '+json.error);return;}
  pedidosData=pedidosData.filter(x=>x.id!=id);
  filtrarPedidos();
}
function editarPedidoModal(id){
  const r=pedidosData.find(x=>x.id==id);
  if(!r) return;
  document.getElementById('eped-id').value=r.id;
  // Rellenar select clientes
  const cliSel=document.getElementById('eped-cliente');
  cliSel.innerHTML='<option value="">Seleccionar...</option>'+CLIENTES.map(c=>`<option value="${c.nombre}"${c.nombre===r.nombreCliente?' selected':''}>${c.nombre}</option>`).join('');
  // Rellenar select productos
  const prodSel=document.getElementById('eped-producto');
  prodSel.innerHTML='<option value="">Seleccionar...</option>'+Object.entries(PROD_MAP).map(([cod,nom])=>`<option value="${nom}"${nom===r.productoNombre?' selected':''}>${nom}</option>`).join('');
  // Rellenar select proyectos según cliente
  epedCargarProyectos(r.nombreCliente, r.proyectoName||r.proyectoCod||'');
  document.getElementById('eped-mat').value=r.matriculacam||'';
  document.getElementById('eped-rem').value=r.matricularem||'';
  document.getElementById('eped-chofer').value=r.chofer||'';
  document.getElementById('eped-bruto').value=r.pesoBruto||'';
  document.getElementById('eped-neto').value=r.pesoNeto||'';
  document.getElementById('eped-obs').value=r.observaciones||'';
  document.getElementById('eped-fecha').value=formatFechaHoraPed(r.fechaHora);
  document.getElementById('eped-msg').textContent='';
  document.getElementById('modal-eped').classList.add('open');
}
function epedClienteChange(){
  const cli=document.getElementById('eped-cliente').value;
  epedCargarProyectos(cli,'');
}
function epedCargarProyectos(nombreCliente, selectedVal){
  const proySel=document.getElementById('eped-proyecto');
  const cliObj=CLIENTES.find(c=>c.nombre===nombreCliente);
  const obras=_getObrasCliente(nombreCliente,cliObj?cliObj.codigo:'');
  if(!obras.length)obras.push({nombre:'CLIENTES VARIOS',codigo:'PV-000'});
  proySel.innerHTML=obras.map(o=>`<option value="${o.codigo}"${(o.nombre===selectedVal||o.codigo===selectedVal)?' selected':''}>${o.codigo} — ${o.nombre}</option>`).join('');
}
function cerrarModalEped(){document.getElementById('modal-eped').classList.remove('open');}
async function guardarPedidoEditar(){
  const btn=document.getElementById('eped-save-btn');
  const msg=document.getElementById('eped-msg');
  btn.disabled=true; btn.textContent='Guardando...';
  try{
    const cliNombre=document.getElementById('eped-cliente').value;
    const cli=CLIENTES.find(c=>c.nombre===cliNombre);
    const prodNombre=document.getElementById('eped-producto').value;
    const prodEntry=Object.entries(PROD_MAP).find(([k,v])=>v===prodNombre);
    const proySel=document.getElementById('eped-proyecto');
    const proyCod=proySel.value;
    const proyName=proySel.options[proySel.selectedIndex]?proySel.options[proySel.selectedIndex].text.split(' — ')[1]||'':'';
    const payload={
      tipo:'editarPedido',
      id:document.getElementById('eped-id').value,
      nombreCliente:cliNombre,
      codigoCliente:cli?cli.codigo:'',
      matriculacam:document.getElementById('eped-mat').value,
      matricularem:document.getElementById('eped-rem').value,
      chofer:document.getElementById('eped-chofer').value,
      productoNombre:prodNombre,
      productoCod:prodEntry?prodEntry[0]:'',
      proyectoCod:proyCod,
      proyectoName:proyName,
      pesoBruto:document.getElementById('eped-bruto').value,
      pesoNeto:document.getElementById('eped-neto').value,
      observaciones:document.getElementById('eped-obs').value,
    };
    const json=await apiPost(payload);
    if(json.ok){
      msg.style.color='var(--accent)'; msg.textContent='Guardado correctamente';
      // Actualizar dato local
      const r=pedidosData.find(x=>x.id==payload.id);
      if(r){Object.assign(r,{nombreCliente:payload.nombreCliente,matriculacam:payload.matriculacam,matricularem:payload.matricularem,chofer:payload.chofer,productoNombre:payload.productoNombre,pesoBruto:Number(payload.pesoBruto),pesoNeto:Number(payload.pesoNeto),observaciones:payload.observaciones});}
      filtrarPedidos();
      setTimeout(cerrarModalEped,900);
    } else {
      msg.style.color='var(--danger)'; msg.textContent='Error: '+(json.error||'desconocido');
    }
  }catch(e){
    msg.style.color='var(--danger)'; msg.textContent='Error de red';
  }finally{
    btn.disabled=false; btn.textContent='Guardar';
  }
}

// ── VENTAS ────────────────────────────────────────────────────
let ventasData=[];
const PROD_CAT2={'ARIDO AF-T-0/4-I':'0/4','ARIDO AG-T-4/12-I':'4/12','ARIDO AG-T-12/20-I':'12/20','ARIDO AG-T-20/40-I':'20/40'};

const FICHA_QR_URLS={'0/4':'https://www.arifoma.com/ficha%20tecnica/04v2.pdf','4/12':'https://www.arifoma.com/ficha%20tecnica/412v2.pdf','12/20':'https://www.arifoma.com/ficha%20tecnica/1220v2.pdf','20/40':'https://www.arifoma.com/ficha%20tecnica/2040v2.pdf'};
function _renderAlbaranQR(productoNombre){
  const qrWrap=document.getElementById('alb-qr-wrap');
  const qrCanvas=document.getElementById('alb-qr-canvas');
  if(!qrWrap||!qrCanvas)return;
  const nombre=(productoNombre||'').toUpperCase();
  let cat=null;
  if(nombre.includes('20/40'))cat='20/40';
  else if(nombre.includes('12/20'))cat='12/20';
  else if(nombre.includes('4/12'))cat='4/12';
  else if(nombre.includes('0/4'))cat='0/4';
  if(!cat||!FICHA_QR_URLS[cat]){qrWrap.style.display='none';return;}
  qrWrap.style.display='flex';
  qrCanvas.innerHTML='';
  function _doQR(){
    if(typeof QRCode!=='undefined'){
      new QRCode(qrCanvas,{text:FICHA_QR_URLS[cat],width:44,height:44,correctLevel:QRCode.CorrectLevel.M});
    } else {
      setTimeout(_doQR,200);
    }
  }
  _doQR();
}
function getCat(prod){
  const p=String(prod||'').toUpperCase();
  if(p.includes('12/20'))return '12/20';
  if(p.includes('20/40'))return '20/40';
  if(p.includes('4/12'))return '4/12';
  if(p.includes('0/4'))return '0/4';
  return 'Otros';
}
function kgToT(kg){return(Number(kg)/1000).toFixed(2);}

async function cargarVentas(){
  try{
    const json=await apiFetch('?accion=pedidos&dias=90');
    if(json.ok){ventasData=json.data;renderVentas();}
  }catch(e){console.warn('Error cargando ventas',e);}
}

function parseFechaHoraObj(fh){
  if(!fh)return null;
  try{
    const s=String(fh).trim();
    let d=null;
    // ISO format: 2026-03-20T16:33:00.000Z
    if(s.includes('T')){
      const [datePart,timePart]=s.split('T');
      const [y,m,day]=datePart.split('-');
      const [h,min]=timePart.split(':');
      d=new Date(Date.UTC(parseInt(y),parseInt(m)-1,parseInt(day),parseInt(h),parseInt(min)));
    }
    // yyyy-mm-dd with time
    else if(s.match(/^\d{4}-\d{2}-\d{2}/)){
      const match=s.match(/^(\d{4})-(\d{2})-(\d{2})\s*(?:(\d{2}):(\d{2}))?/);
      if(!match)return null;
      const [,y,m,day,h,min]=match;
      d=new Date(parseInt(y),parseInt(m)-1,parseInt(day),parseInt(h||0),parseInt(min||0));
    }
    // dd/mm/yyyy HH:mm:ss or dd/mm/yy HH:mm
    else {
      const parts=s.split(' ');
      const dp=parts[0].split('/');
      if(dp.length<3)return null;
      const year=parseInt(dp[2])<100?2000+parseInt(dp[2]):parseInt(dp[2]);
      const h=parts[1]?parseInt(parts[1].split(':')[0]):0;
      const min=parts[1]?parseInt(parts[1].split(':')[1]):0;
      d=new Date(year,parseInt(dp[1])-1,parseInt(dp[0]),h,min);
    }
    return d;
  }catch(e){return null;}
}

function calcVentasTotales(data){
  const cats=['0/4','4/12','12/20','20/40','Otros'];
  const t={};cats.forEach(c=>t[c]=0);
  data.forEach(r=>{const cat=getCat(r.productoNombre);t[cat]=(t[cat]||0)+Number(r.pesoNeto||0);});
  return t;
}

function renderTarjetaVentas(el,titulo,data){
  if(!el)return;
  const cats=['0/4','4/12','12/20','20/40','Otros'];
  const t=calcVentasTotales(data);
  const total=Object.values(t).reduce((a,b)=>a+b,0);
  el.innerHTML='<div class="ventas-card-title">'+titulo+'</div>'+
    cats.map(c=>'<div class="prod-row"><span class="prod-pill">'+c+'</span><span class="prod-val">'+kgToT(t[c])+' T</span></div>').join('')+
    '<div class="total-row-v"><span class="total-label-v">Total</span><span class="total-val-v">'+kgToT(total)+' T</span></div>';
}

function renderVentas(){
  const now=new Date();
  const curM=now.getMonth();const curY=now.getFullYear();
  const selM=parseInt(document.getElementById('ventas-mes-sel').value||curM);

  const dataMesAct=ventasData.filter(r=>{const d=parseFechaHoraObj(r.fechaHora);return d&&d.getMonth()===curM&&d.getFullYear()===curY;});
  const dataMesSel=ventasData.filter(r=>{const d=parseFechaHoraObj(r.fechaHora);return d&&d.getMonth()===selM&&d.getFullYear()===curY;});
  const dataAnyo=ventasData.filter(r=>{const d=parseFechaHoraObj(r.fechaHora);return d&&d.getFullYear()===curY;});
  const dataDia=ventasData.filter(r=>{const d=parseFechaHoraObj(r.fechaHora);return d&&dateStr(d)===dateStr(now);});

  renderTarjetaVentas(document.getElementById('v-mesactual'),'Mes actual ('+MESES[curM]+')',dataMesAct);
  renderTarjetaVentas(document.getElementById('v-messelec'),MESES[selM],dataMesSel);
  renderTarjetaVentas(document.getElementById('v-anyo'),'Año '+curY,dataAnyo);
  renderTarjetaVentas(document.getElementById('v-dia'),'Hoy '+now.getDate()+'/'+pad(now.getMonth()+1),dataDia);

  const ultimas=dataDia.slice(0,8);
  const ulEl=document.getElementById('v-ultimas');
  if(ulEl)ulEl.innerHTML=ultimas.length?
    ultimas.map(r=>'<div style="padding:4px 0;border-bottom:1px solid var(--border);display:flex;justify-content:space-between;font-size:.75rem">'+
      '<span style="font-family:monospace;color:var(--accent)">'+r.matriculacam+'</span>'+
      '<span style="flex:1;margin:0 8px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">'+r.productoNombre+'</span>'+
      '<span style="color:var(--accent2);font-family:monospace;font-weight:700">'+kgToT(r.pesoNeto)+'T</span>'+
    '</div>').join('')
    :'<div style="color:var(--muted);font-size:.78rem;padding:8px 0">Sin pesadas hoy</div>';

  renderTablaVentasDetalle(dataMesSel);
}

function copiarTablaVentas(){
  const tabla=document.getElementById('ventas-tabla');
  if(!tabla)return;
  let tsv='';
  const rows=tabla.querySelectorAll('tr');
  rows.forEach(row=>{
    const cols=row.querySelectorAll('th,td');
    const texto=[];
    cols.forEach(col=>{
      const txt=col.textContent.trim();
      texto.push(txt);
    });
    tsv+=texto.join('\t')+'\n';
  });
  navigator.clipboard.writeText(tsv).then(()=>{
    alert('Tabla copiada al portapapeles. Puedes pegarla en Excel.');
  }).catch(()=>{
    alert('Error al copiar. Intenta seleccionar la tabla manualmente.');
  });
}

function renderTablaVentasDetalle(data){
  const fechaEl=document.getElementById('ventas-detalle-fecha');
  const fechaFiltro=fechaEl?fechaEl.value:null;

  let filtered=data||[];
  if(fechaFiltro){
    filtered=filtered.filter(r=>{
      const d=parseFechaHoraObj(r.fechaHora);
      if(!d)return false;
      const dStr=d.getFullYear()+'-'+pad(d.getMonth()+1)+'-'+pad(d.getDate());
      return dStr===fechaFiltro;
    });
  }

  const tbody=document.getElementById('ventas-detalle-tbody');
  if(!filtered.length){
    tbody.innerHTML='<tr><td colspan="7" style="padding:20px;text-align:center;color:var(--muted)">'+
      (fechaFiltro?'Sin pesadas en esta fecha':'Sin pesadas este mes')+
    '</td></tr>';
    document.getElementById('ventas-resumen-materiales').innerHTML='<div style="padding:16px;text-align:center;color:var(--muted)">Sin materiales</div>';
    document.getElementById('total-toneladas').textContent='—';
    return;
  }

  let totalNeto=0;
  const materiales={};

  tbody.innerHTML=filtered.map(r=>{
    const d=parseFechaHoraObj(r.fechaHora);
    const fecha=d?pad(d.getDate())+'/'+pad(d.getMonth()+1)+'/'+d.getFullYear():'—';
    const fechaHora=d?pad(d.getDate())+'/'+pad(d.getMonth()+1)+'/'+d.getFullYear()+' '+pad(d.getHours())+':'+pad(d.getMinutes())+':'+pad(d.getSeconds()):'—';
    const bruto=Number(r.pesoBruto||0);
    const neto=Number(r.pesoNeto||0);
    const material=r.productoNombre||'Sin material';

    totalNeto+=neto;
    if(!materiales[material])materiales[material]=0;
    materiales[material]+=neto;

    return '<tr style="border-bottom:1px solid var(--border);cursor:pointer" onmouseover="this.style.background=\'var(--surface2)\'" onmouseout="this.style.background=\'transparent\'">'+
      '<td style="padding:8px 12px;color:var(--text);font-family:monospace">'+fecha+'</td>'+
      '<td style="padding:8px 12px;color:var(--text);font-family:monospace">'+fechaHora+'</td>'+
      '<td style="padding:8px 12px;color:var(--accent);font-family:monospace;font-weight:600">'+r.matriculacam+'</td>'+
      '<td style="padding:8px 12px;text-align:right;color:var(--text);font-family:monospace">'+bruto.toLocaleString('es-ES')+'</td>'+
      '<td style="padding:8px 12px;text-align:right;color:var(--accent2);font-family:monospace;font-weight:600">'+neto.toLocaleString('es-ES')+'</td>'+
      '<td style="padding:8px 12px;color:var(--text)">'+material+'</td>'+
      '<td style="padding:8px 12px;color:var(--text)">'+r.nombreCliente||'—'+'</td>'+
    '</tr>';
  }).join('');

  const resumenEl=document.getElementById('ventas-resumen-materiales');
  const materialesHtml=Object.entries(materiales).map(([mat,neto])=>{
    return '<div style="padding:12px 16px;border-bottom:1px solid var(--border);display:flex;justify-content:space-between;align-items:center">'+
      '<span style="color:var(--text);font-size:.9rem">'+mat+'</span>'+
      '<span style="color:var(--accent);font-family:\'DM Mono\',monospace;font-weight:700;font-size:.95rem">'+kgToT(neto)+' T</span>'+
    '</div>';
  }).join('');
  resumenEl.innerHTML=materialesHtml||'<div style="padding:16px;text-align:center;color:var(--muted)">Sin materiales</div>';

  document.getElementById('total-toneladas').textContent=kgToT(totalNeto);
}

// ── CAMIONES GESTIÓN ──────────────────────────────────────────
let camGestData=[];let camEditingId=null;
async function cargarCamiones(){
  const el=document.getElementById('camiones-list');
  el.innerHTML='<div class="tbl"><div class="empty">Cargando...</div></div>';
  // Cargar choferes en background si no están cargados
  if(!choferesData.length){apiFetch('?accion=choferes').then(j=>{if(j.ok)choferesData=j.data||[];}).catch(()=>{});}
  try{
    const json=await apiFetch('?accion=camiones');
    if(!json.ok)throw new Error(json.error);
    camGestData=json.data;camionesData=json.data;
    filtrarCamionesGestion();
  }catch(e){el.innerHTML='<div class="tbl"><div class="empty">Error: '+e.message+'</div></div>';}
}
function filtrarCamionesGestion(){
  const q=(document.getElementById('filt-cam').value||'').toUpperCase();
  let data=camGestData;
  if(q)data=data.filter(c=>String(c.matriculacam).toUpperCase().includes(q)||String(c.chofer||'').toUpperCase().includes(q)||String(c.proveedor||'').toUpperCase().includes(q));
  renderCamionesGestion(data);
}
function renderCamionesGestion(data){
  const el=document.getElementById('camiones-list');
  if(!data.length){el.innerHTML='<div class="tbl"><div class="empty">Sin resultados</div></div>';return;}
  el.innerHTML='<div class="tbl"><div class="tr th"><div class="tc" style="flex:1">Matrícula</div><div class="tc" style="flex:.7">Remolque</div><div class="tc" style="flex:1.2">Chofer</div><div class="tc" style="flex:1">Proveedor</div><div class="tc" style="flex:.7;text-align:right">Tara</div><div class="tc" style="flex:.4"></div></div>'+
  data.map(c=>`<div class="tr"><div class="tc" style="flex:1;font-family:monospace;font-weight:700;color:var(--accent)">${c.matriculacam}</div><div class="tc" style="flex:.7;font-family:monospace">${c.matricularem||'—'}</div><div class="tc" style="flex:1.2">${c.chofer||'—'}</div><div class="tc" style="flex:1;color:var(--muted)">${c.proveedor||'—'}</div><div class="tc" style="flex:.7;text-align:right;font-family:monospace">${Number(c.tara||0).toLocaleString()} kg</div><div class="tc" style="flex:.4;text-align:right"><button class="btn-sm" onclick="openCamModal(${c.id})">Editar</button></div></div>`).join('')+'</div>';
}
function openCamModal(id){
  camEditingId=id;
  const modal=document.getElementById('cam-modal');
  document.getElementById('cam-modal-title').textContent=id?'Editar camión':'Nuevo camión';
  const delBtn=document.getElementById('cm-del-btn');
  if(id){
    const c=camGestData.find(x=>x.id==id);if(!c)return;
    document.getElementById('cm-mat').value=c.matriculacam||'';
    document.getElementById('cm-rem').value=c.matricularem||'';
    document.getElementById('cm-tara').value=c.tara||0;
    document.getElementById('cm-chofer').value=c.chofer||'';
    document.getElementById('cm-prov').value=c.proveedor||'';
    document.getElementById('cm-tel').value=c.telefono||'';
    delBtn.style.display='block';
  } else {
    ['cm-mat','cm-rem','cm-tara','cm-chofer','cm-prov','cm-tel'].forEach(i=>{const el=document.getElementById(i);if(el)el.value='';});
    delBtn.style.display='none';
  }
  modal.classList.add('open');
}
function toggleCamChoferDropdown(show){
  const dd=document.getElementById('cm-chofer-dropdown');
  if(show){dd.style.display='block';filtrarCamChofer();}
  else setTimeout(()=>{dd.style.display='none';},150);
}
function filtrarCamChofer(){
  const q=(document.getElementById('cm-chofer').value||'').toUpperCase();
  const list=document.getElementById('cm-chofer-list');if(!list)return;
  const sorted=[...(choferesData||[])].sort((a,b)=>(a.nombre||'').localeCompare(b.nombre||''));
  const filtered=q?sorted.filter(c=>(c.nombre||'').toUpperCase().includes(q)):sorted;
  list.innerHTML='';
  filtered.forEach(c=>{
    const div=document.createElement('div');
    div.style.cssText='padding:9px 14px;cursor:pointer;font-size:.82rem;border-bottom:1px solid var(--border);display:flex;justify-content:space-between;align-items:center';
    const caeTag=c.cae?'<span style="font-size:.65rem;color:#0c6;font-weight:700;margin-left:8px">CAE ✓</span>':'';
    div.innerHTML='<span>'+c.nombre+caeTag+'</span><span style="font-size:.7rem;color:var(--muted)">'+(c.empresa||'')+'</span>';
    div.onmouseover=()=>div.style.background='var(--surface2)';
    div.onmouseout=()=>div.style.background='transparent';
    div.onclick=()=>{
      document.getElementById('cm-chofer').value=c.nombre;
      // Auto-rellenar empresa y teléfono si están vacíos
      const provEl=document.getElementById('cm-prov');
      const telEl=document.getElementById('cm-tel');
      if(!provEl.value&&c.empresa)provEl.value=c.empresa;
      if(!telEl.value&&c.telefono)telEl.value=c.telefono;
      document.getElementById('cm-chofer-dropdown').style.display='none';
    };
    list.appendChild(div);
  });
  if(!filtered.length){
    list.innerHTML='<div style="padding:10px 14px;font-size:.78rem;color:var(--muted);font-style:italic">Sin coincidencias — se usará el texto escrito</div>';
  }
}
function closeCamModal(){document.getElementById('cam-modal').classList.remove('open');camEditingId=null;document.getElementById('cm-chofer-dropdown').style.display='none';}
async function saveCamion(){
  const payload={
    tipo:camEditingId?'editarCamion':'nuevoCamion',
    id:camEditingId,
    matriculacam:document.getElementById('cm-mat').value.toUpperCase(),
    matricularem:document.getElementById('cm-rem').value,
    tara:parseFloat(document.getElementById('cm-tara').value)||0,
    chofer:document.getElementById('cm-chofer').value,
    proveedor:document.getElementById('cm-prov').value,
    telefono:document.getElementById('cm-tel').value,
  };
  try{
    const json=await apiPost(payload);
    if(json.ok){
      closeCamModal();
      // Actualizar lista local en lugar de recargar
      if(camEditingId){
        const idx=camGestData.findIndex(x=>x.id==camEditingId);
        if(idx>=0) camGestData[idx]={...camGestData[idx],...payload};
      } else {
        // Nuevo camión: agregar con ID temporal o esperar desde API
        payload.id=Math.max(...camGestData.map(c=>c.id||0),0)+1;
        camGestData.push(payload);
      }
      filtrarCamionesGestion();
    }
    else alert('Error: '+json.error);
  }catch(e){alert('Error de conexión');}
}
async function eliminarCamion(){
  if(!camEditingId)return;
  const c=camGestData.find(x=>x.id==camEditingId);
  const mat=(c?.matriculacam||'').trim();
  if(!mat){alert('No se pueden eliminar camiones vacíos. Completa los datos primero.');return;}
  if(!confirm('¿Eliminar este camión de la base de datos?'))return;
  const payload={tipo:'eliminarCamion',id:camEditingId};
  try{
    const json=await apiPost(payload);
    if(json.ok){
      closeCamModal();
      camGestData=camGestData.filter(x=>x.id!=camEditingId);
      filtrarCamionesGestion();
    }
    else alert('Error: '+json.error);
  }catch(e){alert('Error de conexión');}
}

// ── CHOFERES ────────────────────────────────────────────────
let choferesData=[];let choferEditingId=null;
const CHOFERES_ONEDRIVE_BASE='Escritorio/Arifoma/13. SEGURIDAD Y SALUD/13.02 SERVICIO DE PREVENCION/COORDINACION AE/0. CAE DOCUMENTACION';

async function cargarChoferes(){
  const el=document.getElementById('choferes-list');
  el.innerHTML='<div class="tbl"><div class="empty">Cargando...</div></div>';
  try{
    const json=await apiFetch('?accion=choferes');
    if(!json.ok)throw new Error(json.error);
    choferesData=json.data||[];
    filtrarChoferesGestion();
  }catch(e){el.innerHTML='<div class="tbl"><div class="empty">Error: '+e.message+'</div></div>';}
}

function filtrarChoferesGestion(){
  const q=(document.getElementById('filt-chofer').value||'').toUpperCase();
  let data=choferesData;
  if(q)data=data.filter(c=>String(c.nombre||'').toUpperCase().includes(q)||String(c.dni||'').toUpperCase().includes(q)||String(c.empresa||'').toUpperCase().includes(q));
  renderChoferesGestion(data);
}

function renderChoferesGestion(data){
  const el=document.getElementById('choferes-list');
  if(!data.length){el.innerHTML='<div class="tbl"><div class="empty">Sin resultados</div></div>';return;}
  const hoy=new Date().toISOString().slice(0,10);
  el.innerHTML='<div class="tbl"><div class="tr th"><div class="tc" style="flex:.4;text-align:center">CAE</div><div class="tc" style="flex:1.2">Nombre</div><div class="tc" style="flex:.8">DNI</div><div class="tc" style="flex:.8">Teléfono</div><div class="tc" style="flex:1">Empresa</div><div class="tc" style="flex:.7">Venc. CAE</div><div class="tc" style="flex:.8;text-align:right"></div></div>'+
  data.map(c=>{
    const caeOk=c.cae;
    const venc=c.cae_vencimiento||'';
    const vencido=venc&&venc<hoy;
    const caeBadge=caeOk?(vencido?'<span style="color:#e44;font-size:.85rem" title="CAE vencido">&#9888;</span>':'<span style="color:#0c6;font-size:1rem">&#10003;</span>'):'<span style="color:var(--muted)">—</span>';
    const vencTxt=venc?'<span style="'+(vencido?'color:#e44;font-weight:700':'color:var(--text)')+'">'+venc.split('-').reverse().join('/')+'</span>':'—';
    const carpeta=c.cae_carpeta?`<button class="btn-sec btn-read" onclick="abrirCarpetaCAE(${c.id})" title="Ver documentos CAE" style="padding:4px 8px;font-size:.8rem;color:var(--accent);border-color:var(--accent)">📁</button>`:'';
    return `<div class="tr"><div class="tc" style="flex:.4;text-align:center">${caeBadge}</div><div class="tc" style="flex:1.2;font-weight:600">${c.nombre}</div><div class="tc" style="flex:.8;font-family:monospace">${c.dni||'—'}</div><div class="tc" style="flex:.8;font-family:monospace">${c.telefono||'—'}</div><div class="tc" style="flex:1;color:var(--muted)">${c.empresa||'—'}</div><div class="tc" style="flex:.7">${vencTxt}</div><div class="tc" style="flex:.8;text-align:right;display:flex;gap:4px;justify-content:flex-end">${carpeta}<button class="btn-sec btn-read" onclick="imprimirChofer(${c.id})" title="Imprimir ficha" style="padding:4px 8px;font-size:.8rem">🖨</button><button class="btn-sm" onclick="openChoferModal(${c.id})">Editar</button></div></div>`;
  }).join('')+'</div>';
}

function openChoferModal(id){
  choferEditingId=id;
  const modal=document.getElementById('chofer-modal');
  document.getElementById('chofer-modal-title').textContent=id?'Editar conductor':'Nuevo conductor';
  const delBtn=document.getElementById('ch-del-btn');
  const caeCheck=document.getElementById('ch-cae');
  const caeExtra=document.getElementById('ch-cae-extra');
  if(id){
    const c=choferesData.find(x=>x.id==id);if(!c)return;
    document.getElementById('ch-nombre').value=c.nombre||'';
    document.getElementById('ch-dni').value=c.dni||'';
    document.getElementById('ch-telefono').value=c.telefono||'';
    document.getElementById('ch-empresa').value=c.empresa||'';
    caeCheck.checked=!!c.cae;
    caeExtra.style.display=c.cae?'block':'none';
    document.getElementById('ch-cae-carpeta').value=c.cae_carpeta||'';
    document.getElementById('ch-cae-fecha').value=c.cae_fecha||'';
    document.getElementById('ch-cae-file').value='';
    document.getElementById('ch-cae-cam').value='';
    document.getElementById('ch-cae-file-name').textContent='';
    calcVencCAE();
    delBtn.style.display='block';
  } else {
    ['ch-nombre','ch-dni','ch-telefono','ch-empresa','ch-cae-fecha','ch-cae-carpeta'].forEach(i=>{const el=document.getElementById(i);if(el)el.value='';});
    document.getElementById('ch-cae-venc-info').textContent='';
    caeCheck.checked=false;
    caeExtra.style.display='none';
    document.getElementById('ch-cae-file').value='';
    document.getElementById('ch-cae-cam').value='';
    document.getElementById('ch-cae-file-name').textContent='';
    delBtn.style.display='none';
  }
  caeCheck.onchange=()=>{caeExtra.style.display=caeCheck.checked?'block':'none';};
  document.getElementById('ch-cae-fecha').onchange=()=>{calcVencCAE();};
  document.getElementById('ch-cae-file').onchange=function(){
    document.getElementById('ch-cae-cam').value='';
    document.getElementById('ch-cae-file-name').textContent=this.files[0]?this.files[0].name:'';
  };
  document.getElementById('ch-cae-cam').onchange=function(){
    document.getElementById('ch-cae-file').value='';
    document.getElementById('ch-cae-file-name').textContent=this.files[0]?'📷 '+this.files[0].name:'';
  };
  modal.classList.add('open');
}

function calcVencCAE(){
  const fecha=document.getElementById('ch-cae-fecha').value;
  const info=document.getElementById('ch-cae-venc-info');
  if(!fecha){info.textContent='';return;}
  const d=new Date(fecha);d.setFullYear(d.getFullYear()+1);
  const venc=d.toISOString().slice(0,10);
  const hoy=new Date().toISOString().slice(0,10);
  const vencido=venc<hoy;
  info.innerHTML='Vencimiento: <strong style="color:'+(vencido?'#e44':'#0c6')+'">'+venc.split('-').reverse().join('/')+'</strong>'+(vencido?' — <span style="color:#e44">VENCIDO</span>':'');
}
function closeChoferModal(){document.getElementById('chofer-modal').classList.remove('open');choferEditingId=null;}

async function saveChofer(){
  const nombre=document.getElementById('ch-nombre').value.trim();
  if(!nombre){alert('Introduce un nombre.');return;}
  const payload={
    tipo:choferEditingId?'editarChofer':'nuevoChofer',
    id:choferEditingId,
    nombre,
    dni:document.getElementById('ch-dni').value.toUpperCase().trim(),
    telefono:document.getElementById('ch-telefono').value.trim(),
    empresa:document.getElementById('ch-empresa').value.trim(),
    cae:document.getElementById('ch-cae').checked,
    cae_carpeta:document.getElementById('ch-cae-carpeta').value.trim()||null,
    cae_fecha:document.getElementById('ch-cae-fecha').value||null,
    cae_vencimiento:(()=>{const f=document.getElementById('ch-cae-fecha').value;if(!f)return null;const d=new Date(f);d.setFullYear(d.getFullYear()+1);return d.toISOString().slice(0,10);})(),
  };

  // Subir PDF CAE a OneDrive si hay archivo
  const fileInput=document.getElementById('ch-cae-file');
  const camInput=document.getElementById('ch-cae-cam');
  const carpetaCAE=document.getElementById('ch-cae-carpeta').value.trim();
  const uploadFile=fileInput.files.length>0?fileInput.files[0]:(camInput.files.length>0?camInput.files[0]:null);
  if(uploadFile&&document.getElementById('ch-cae').checked&&carpetaCAE){
    try{
      const token=await comprasGetToken();
      const folderPath=CHOFERES_ONEDRIVE_BASE+'/'+carpetaCAE;
      // Crear carpeta si no existe
      const parentEncoded=CHOFERES_ONEDRIVE_BASE.split('/').map(s=>encodeURIComponent(s)).join('/');
      await fetch('https://graph.microsoft.com/v1.0/me/drive/root:/'+parentEncoded+':/children',{
        method:'POST',
        headers:{'Authorization':'Bearer '+token,'Content-Type':'application/json'},
        body:JSON.stringify({name:carpetaCAE,folder:{},'@microsoft.graph.conflictBehavior':'fail'})
      });
      const encodedPath=folderPath.split('/').map(s=>encodeURIComponent(s)).join('/');
      const fileName=uploadFile.name;
      const uploadUrl='https://graph.microsoft.com/v1.0/me/drive/root:/'+encodedPath+'/'+encodeURIComponent(fileName)+':/content';
      const resp=await fetch(uploadUrl,{
        method:'PUT',
        headers:{'Authorization':'Bearer '+token,'Content-Type':uploadFile.type||'application/pdf'},
        body:uploadFile
      });
      if(!resp.ok){const err=await resp.text();alert('Error subiendo PDF: '+err);}
      else{payload.cae_documento=folderPath+'/'+fileName;}
    }catch(e){alert('Error OneDrive: '+e.message);}
  }

  try{
    const json=await apiPost(payload);
    if(json.ok){
      closeChoferModal();
      if(choferEditingId){
        const idx=choferesData.findIndex(x=>x.id==choferEditingId);
        if(idx>=0) choferesData[idx]={...choferesData[idx],...payload};
      } else {
        payload.id=Math.max(...choferesData.map(c=>c.id||0),0)+1;
        choferesData.push(payload);
      }
      filtrarChoferesGestion();
    }
    else alert('Error: '+json.error);
  }catch(e){alert('Error de conexión');}
}

async function eliminarChofer(){
  if(!choferEditingId)return;
  if(!confirm('¿Eliminar este conductor de la base de datos?'))return;
  try{
    const json=await apiPost({tipo:'eliminarChofer',id:choferEditingId});
    if(json.ok){
      closeChoferModal();
      choferesData=choferesData.filter(x=>x.id!=choferEditingId);
      filtrarChoferesGestion();
    }
    else alert('Error: '+json.error);
  }catch(e){alert('Error de conexión');}
}

function imprimirChofer(id){
  const c=choferesData.find(x=>x.id==id);if(!c)return;
  const venc=c.cae_vencimiento?c.cae_vencimiento.split('-').reverse().join('/'):'—';
  const w=window.open('','_blank','width=600,height=500');
  w.document.write(`<!DOCTYPE html><html><head><meta charset="utf-8"><title>Ficha conductor — ${c.nombre}</title>
<style>body{font-family:Arial,sans-serif;padding:30px;color:#222}h2{margin:0 0 4px;font-size:1.3rem}
.sub{color:#888;font-size:.85rem;margin-bottom:20px}
table{width:100%;border-collapse:collapse;margin-top:10px}
td{padding:8px 12px;border:1px solid #ddd;font-size:.9rem}
td:first-child{font-weight:700;width:170px;background:#f8f8f8}
.cae-ok{color:#0a0;font-weight:700}.cae-no{color:#c00;font-weight:700}
.footer{margin-top:30px;font-size:.7rem;color:#aaa;text-align:center}
@media print{body{padding:15px}}
</style></head><body>
<h2>${c.nombre}</h2>
<div class="sub">Ficha de conductor — ARIFOMA</div>
<table>
<tr><td>DNI / NIE</td><td>${c.dni||'—'}</td></tr>
<tr><td>Teléfono</td><td>${c.telefono||'—'}</td></tr>
<tr><td>Empresa</td><td>${c.empresa||'—'}</td></tr>
<tr><td>CAE</td><td class="${c.cae?'cae-ok':'cae-no'}">${c.cae?'&#10003; Sí':'&#10007; No'}</td></tr>
<tr><td>Vencimiento CAE</td><td>${venc}</td></tr>
</table>
<div class="footer">Generado el ${new Date().toLocaleDateString('es-ES')} — ARIFOMA</div>
<script>window.print();<\/script></body></html>`);
  w.document.close();
}

function imprimirCAE(){
  const base='%2Fpersonal%2Fgreyes%5Farifoma%5Fcom%2FDocuments%2FEscritorio%2FArifoma%2F13%2E%20SEGURIDAD%20Y%20SALUD%2F13%2E02%20SERVICIO%20DE%20PREVENCION%2FCOORDINACION%20AE%2F0%2E%20CAE%20DOCUMENTACION';
  const file=base+'%2F1%2E%20CAE%20SOLICITUD%20Y%20ENTREGA%20DE%20DOCUMENTACION%2Epdf';
  window.open('https://grpsite-my.sharepoint.com/personal/greyes_arifoma_com/_layouts/15/onedrive.aspx?id='+file+'&parent='+base,'_blank');
}

async function abrirCarpetaCAE(id){
  const c=choferesData.find(x=>x.id==id);if(!c||!c.cae_carpeta)return;
  const folderPath=CHOFERES_ONEDRIVE_BASE+'/'+c.cae_carpeta;
  const encodedPath=folderPath.split('/').map(s=>encodeURIComponent(s)).join('/');
  try{
    const token=await comprasGetToken();
    const resp=await fetch('https://graph.microsoft.com/v1.0/me/drive/root:/'+encodedPath+':/children?$select=name,webUrl,size,lastModifiedDateTime,file&$orderby=name',{
      headers:{'Authorization':'Bearer '+token}
    });
    if(!resp.ok)throw new Error('No se pudo acceder a la carpeta');
    const data=await resp.json();
    const files=data.value||[];
    // Mostrar en ventana
    const w=window.open('','_blank','width=700,height=500');
    w.document.write(`<!DOCTYPE html><html><head><meta charset="utf-8"><title>CAE — ${c.nombre}</title>
<style>body{font-family:Arial,sans-serif;padding:20px;color:#222}h2{margin:0 0 4px;font-size:1.2rem}
.sub{color:#888;font-size:.82rem;margin-bottom:16px}
table{width:100%;border-collapse:collapse}th,td{padding:8px 10px;border:1px solid #ddd;font-size:.85rem;text-align:left}
th{background:#f5f5f5;font-weight:700}a{color:#4a6e1f;text-decoration:none;font-weight:600}a:hover{text-decoration:underline}
.empty{color:#999;padding:20px;text-align:center}
</style></head><body>
<h2>📁 Documentos CAE — ${c.nombre}</h2>
<div class="sub">${c.cae_carpeta} · ${files.length} archivo${files.length!==1?'s':''}</div>
${files.length?'<table><tr><th>Archivo</th><th>Tamaño</th><th>Modificado</th></tr>'+files.map(f=>{
  const size=f.size?(f.size/1024).toFixed(0)+' KB':'—';
  const mod=f.lastModifiedDateTime?new Date(f.lastModifiedDateTime).toLocaleDateString('es-ES',{day:'2-digit',month:'2-digit',year:'2-digit'}):'—';
  return '<tr><td><a href="'+f.webUrl+'" target="_blank">'+f.name+'</a></td><td>'+size+'</td><td>'+mod+'</td></tr>';
}).join('')+'</table>':'<div class="empty">Carpeta vacía</div>'}
</body></html>`);
    w.document.close();
  }catch(e){alert('Error accediendo a OneDrive: '+e.message);}
}

// ── OBRAS / PROYECTOS ────────────────────────────────────────
let obrasGestData=[];
let obraEditingId=null;

async function cargarObras(){
  const el=document.getElementById('obras-list');
  el.innerHTML='<div class="tbl"><div class="empty">Cargando...</div></div>';
  try{
    const json=await apiFetch('?accion=obras');
    if(!json.ok)throw new Error(json.error);
    obrasGestData=json.data||[];
    filtrarObrasGestion();
  }catch(e){el.innerHTML='<div class="tbl"><div class="empty">Error: '+e.message+'</div></div>';}
}

function filtrarObrasGestion(){
  const q=(document.getElementById('filt-obra').value||'').toUpperCase();
  let data=obrasGestData;
  if(q)data=data.filter(o=>String(o.codigo||'').toUpperCase().includes(q)||String(o.nombre||'').toUpperCase().includes(q)||String(o.nombreCliente||'').toUpperCase().includes(q));
  renderObrasGestion(data);
}

function renderObrasGestion(data){
  const el=document.getElementById('obras-list');
  if(!data.length){el.innerHTML='<div class="tbl"><div class="empty">Sin resultados</div></div>';return;}
  el.innerHTML='<div class="tbl"><div class="tr th"><div class="tc" style="flex:.6">Código</div><div class="tc" style="flex:1.3">Nombre obra</div><div class="tc" style="flex:1.2">Cliente</div><div class="tc" style="flex:.5">Estado</div><div class="tc" style="flex:.4"></div></div>'+
  data.map(o=>`<div class="tr"><div class="tc" style="flex:.6;font-family:monospace;font-weight:700;color:var(--accent)">${o.codigo}</div><div class="tc" style="flex:1.3">${o.nombre}</div><div class="tc" style="flex:1.2;color:var(--muted)">${o.nombreCliente||'—'}</div><div class="tc" style="flex:.5">${o.activo!==false?'<span style="color:#0c6">Activa</span>':'<span style="color:var(--muted)">Inactiva</span>'}</div><div class="tc" style="flex:.4;text-align:right"><button class="btn-sm" onclick="openObraModal(${o.id})">Editar</button></div></div>`).join('')+'</div>';
}

let _obraSelCliente=null; // {codigo, nombre}

function openObraModal(id){
  obraEditingId=id;
  _obraSelCliente=null;
  const modal=document.getElementById('obra-modal');
  document.getElementById('obra-modal-title').textContent=id?'Editar obra':'Nueva obra';
  const delBtn=document.getElementById('ob-del-btn');
  const cliText=document.getElementById('ob-cli-text');
  document.getElementById('ob-cli-dropdown').style.display='none';
  document.getElementById('ob-cli-search').value='';
  if(id){
    const o=obrasGestData.find(x=>x.id==id);if(!o)return;
    _obraSelCliente={codigo:o.codigoCliente,nombre:o.nombreCliente};
    cliText.textContent=o.nombreCliente||'Seleccionar cliente...';
    cliText.style.color=o.nombreCliente?'var(--text)':'var(--muted)';
    document.getElementById('ob-codigo').value=o.codigo||'';
    document.getElementById('ob-nombre').value=o.nombre||'';
    document.getElementById('ob-activo').checked=o.activo!==false;
    delBtn.style.display='block';
  } else {
    cliText.textContent='Seleccionar cliente...';
    cliText.style.color='var(--muted)';
    document.getElementById('ob-codigo').value='';
    document.getElementById('ob-nombre').value='';
    document.getElementById('ob-activo').checked=true;
    delBtn.style.display='none';
  }
  modal.classList.add('open');
}

function toggleObraCliDropdown(){
  const dd=document.getElementById('ob-cli-dropdown');
  dd.style.display=dd.style.display==='none'?'block':'none';
  if(dd.style.display==='block'){
    document.getElementById('ob-cli-search').value='';
    document.getElementById('ob-cli-search').focus();
    renderObraCliList('');
  }
}

function renderObraCliList(q){
  const list=document.getElementById('ob-cli-list');if(!list)return;
  const sorted=[...CLIENTES].sort((a,b)=>a.nombre.localeCompare(b.nombre));
  const filtered=q?sorted.filter(c=>c.nombre.toUpperCase().includes(q.toUpperCase())||c.codigo.toUpperCase().includes(q.toUpperCase())):sorted;
  list.innerHTML='';
  filtered.forEach(c=>{
    const div=document.createElement('div');
    div.style.cssText='padding:10px 14px;cursor:pointer;font-size:.82rem;border-bottom:1px solid var(--border)';
    div.innerHTML='<span style="font-family:monospace;color:var(--accent);margin-right:8px">'+c.codigo+'</span>'+c.nombre;
    div.onmouseover=()=>div.style.background='var(--surface2)';
    div.onmouseout=()=>div.style.background='transparent';
    div.onclick=()=>{
      _obraSelCliente={codigo:c.codigo,nombre:c.nombre};
      document.getElementById('ob-cli-text').textContent=c.nombre;
      document.getElementById('ob-cli-text').style.color='var(--text)';
      document.getElementById('ob-cli-dropdown').style.display='none';
    };
    list.appendChild(div);
  });
}

function filtrarObraClientes(){
  renderObraCliList(document.getElementById('ob-cli-search').value);
}

function closeObraModal(){document.getElementById('obra-modal').classList.remove('open');obraEditingId=null;_obraSelCliente=null;}

async function saveObra(){
  if(!_obraSelCliente){alert('Selecciona un cliente.');return;}
  const codigo=document.getElementById('ob-codigo').value.toUpperCase().trim();
  const nombre=document.getElementById('ob-nombre').value.trim();
  if(!codigo){alert('Introduce un código de obra.');return;}
  if(!nombre){alert('Introduce un nombre de obra.');return;}
  const payload={
    tipo:obraEditingId?'editarObra':'nuevaObra',
    id:obraEditingId,
    codigo,
    nombre,
    codigoCliente:_obraSelCliente.codigo,
    nombreCliente:_obraSelCliente.nombre,
    activo:document.getElementById('ob-activo').checked,
  };
  try{
    const json=await apiPost(payload);
    if(json.ok){
      closeObraModal();
      if(obraEditingId){
        const idx=obrasGestData.findIndex(x=>x.id==obraEditingId);
        if(idx>=0) obrasGestData[idx]={...obrasGestData[idx],...payload};
      } else {
        payload.id=Math.max(...obrasGestData.map(o=>o.id||0),0)+1;
        obrasGestData.push(payload);
      }
      filtrarObrasGestion();
      // Actualizar CLI_PROY en memoria
      _actualizarCliProyDesdeObras();
    }
    else alert('Error: '+json.error);
  }catch(e){alert('Error de conexión');}
}

async function eliminarObra(){
  if(!obraEditingId)return;
  if(!confirm('¿Eliminar esta obra de la base de datos?'))return;
  try{
    const json=await apiPost({tipo:'eliminarObra',id:obraEditingId});
    if(json.ok){
      closeObraModal();
      obrasGestData=obrasGestData.filter(x=>x.id!=obraEditingId);
      filtrarObrasGestion();
      _actualizarCliProyDesdeObras();
    }
    else alert('Error: '+json.error);
  }catch(e){alert('Error de conexión');}
}

function _actualizarCliProyDesdeObras(){
  const porNombre={};
  const porCodigo={};
  obrasGestData.filter(o=>o.activo!==false).sort((a,b)=>(a.codigo||'').localeCompare(b.codigo||'')).forEach(o=>{
    const cli=o.nombreCliente||'';
    const cod=o.codigoCliente||'';
    if(cli){
      if(!porNombre[cli])porNombre[cli]=[];
      porNombre[cli].push({nombre:o.nombre,codigo:o.codigo});
    }
    if(cod){
      if(!porCodigo[cod])porCodigo[cod]=[];
      porCodigo[cod].push({nombre:o.nombre,codigo:o.codigo});
    }
  });
  CLI_PROY=porNombre;
  window._CLI_PROY_COD=porCodigo;
}

// Buscar obras para un cliente por nombre o código
function _getObrasCliente(nombre,codigo){
  return CLI_PROY[nombre]||
    (codigo&&window._CLI_PROY_COD?window._CLI_PROY_COD[codigo]:null)||
    [];
}

// ── ACTIVOS ───────────────────────────────────────────────────
let activosData=[];
let activosFiltroTipo='TODOS';

async function cargarActivos(){
  const el=document.getElementById('activos-list');
  if(el)el.innerHTML='<div class="tbl"><div class="empty">Cargando...</div></div>';
  try{
    const result=await dbQuery({ action: 'select', table: 'tblactivos', options: { select: '*', order: 'Codigo.asc' } });
    if(!result.ok)throw new Error(result.error);
    activosData=result.data||[];
    _buildOTFromActivos(activosData);
    buildActivosFiltros();
    filtrarActivos();
  }catch(e){
    if(el)el.innerHTML='<div class="tbl"><div class="empty">Error: '+e.message+'</div></div>';
  }
}

function activosTab(tab){
  ['maquinas','consumos'].forEach(t=>{
    document.getElementById('atab-'+t).style.display=t===tab?'block':'none';
    const btn=document.getElementById('atab-btn-'+t);
    btn.style.background=t===tab?'var(--accent)':'transparent';
    btn.style.color=t===tab?'#fff':'var(--muted)';
  });
  if(tab==='consumos')renderActivosConsumos();
}

function initActivos(){
  cargarActivos();
}

function buildActivosFiltros(){
  const ft=document.getElementById('activos-filter-tabs');
  if(!ft)return;
  ft.innerHTML='';
  const tipos=['TODOS',...new Set(activosData.map(m=>m.tipoactivo).filter(Boolean))];
  tipos.forEach(t=>{
    const d=document.createElement('div');
    d.className='filter-tab'+(t==='TODOS'?' active':'');
    d.textContent=t==='TODOS'?'Todos':t.charAt(0)+t.slice(1).toLowerCase();
    d.dataset.tipo=t;
    d.onclick=()=>{
      activosFiltroTipo=t;
      document.querySelectorAll('#activos-filter-tabs .filter-tab').forEach(x=>x.classList.remove('active'));
      d.classList.add('active');
      filtrarActivos();
    };
    ft.appendChild(d);
  });
  activosFiltroTipo='TODOS';
}

function filtrarActivos(){
  const q=(document.getElementById('filt-activos').value||'').toLowerCase();
  let list=activosData.filter(m=>m.Codigo); // skip blank rows
  if(activosFiltroTipo!=='TODOS')list=list.filter(m=>m.tipoactivo===activosFiltroTipo);
  if(q)list=list.filter(m=>
    (m.Codigo||'').toLowerCase().includes(q)||
    (m.Activo||'').toLowerCase().includes(q)||
    (m.fabricante||'').toLowerCase().includes(q)||
    (m.proveedor||'').toLowerCase().includes(q)
  );
  renderActivos(list);
}

function renderActivos(data){
  const el=document.getElementById('activos-list');
  if(!data.length){el.innerHTML='<div class="tbl"><div class="empty">Sin resultados</div></div>';return;}
  el.innerHTML='<div class="tbl">'+
    '<div class="tr th">'+
      '<div class="tc" style="flex:.7">Código</div>'+
      '<div class="tc" style="flex:1.3">Tipo</div>'+
      '<div class="tc" style="flex:.8">Fabricante</div>'+
      '<div class="tc" style="flex:1">Proveedor</div>'+
      '<div class="tc" style="flex:.5;text-align:right">€/h</div>'+
      '<div class="tc" style="flex:.4"></div>'+
    '</div>'+
    data.map(m=>'<div class="tr">'+
      '<div class="tc" style="flex:.7;font-family:monospace;font-weight:700;color:var(--accent)">'+(m.Codigo||'—')+'</div>'+
      '<div class="tc" style="flex:1.3;font-size:.8rem;text-transform:uppercase;color:var(--text)">'+(m.tipoactivo||'—')+'</div>'+
      '<div class="tc" style="flex:.8;color:var(--muted)">'+(m.fabricante||'—')+'</div>'+
      '<div class="tc" style="flex:1;color:var(--muted);font-size:.8rem">'+(m.proveedor||'—')+'</div>'+
      '<div class="tc" style="flex:.5;font-family:monospace;text-align:right">'+(m.phora||'—')+'</div>'+
      '<div class="tc" style="flex:.7;text-align:right;display:flex;gap:4px;justify-content:flex-end"><button class="btn-sm" onclick="openActivosModal('+m.id+')">Editar</button><button class="btn-sm" style="background:var(--danger,#e53935);color:#fff;border-color:var(--danger,#e53935)" onclick="eliminarActivo('+m.id+',\''+( m.Codigo||'').replace(/'/g,"\\'")+'\')" >Eliminar</button></div>'+
    '</div>').join('')+
  '</div>';
}

function renderActivosConsumos(){
  const el=document.getElementById('activos-consumos-list');
  if(!el)return;
  const data=typeof gasoilConsumos!=='undefined'?gasoilConsumos:[];
  if(!data.length){
    el.innerHTML='<div class="tbl"><div class="empty">Sin datos — abre Gasoil y pulsa Actualizar primero</div></div>';return;
  }
  // enrich with phora from activosData
  const phoraMap={};
  activosData.forEach(m=>{if(m.Codigo&&m.phora)phoraMap[m.Codigo.toUpperCase()]=m.phora;});
  const rows=data.map(c=>{
    const ph=phoraMap[String(c.activo||'').toUpperCase()]||'—';
    return '<div class="tr">'+
      '<div class="tc" style="flex:1.1;font-weight:700;color:var(--text)">'+(c.activo||'—')+'</div>'+
      '<div class="tc" style="flex:.9;font-family:monospace;color:var(--accent);text-align:right">'+Number(c.litros||0).toLocaleString()+' L</div>'+
      '<div class="tc" style="flex:.65;font-family:monospace;color:var(--muted);text-align:right">'+Number(c.max||0).toLocaleString()+'</div>'+
      '<div class="tc" style="flex:.65;font-family:monospace;color:var(--muted);text-align:right">'+Number(c.min||0).toLocaleString()+'</div>'+
      '<div class="tc" style="flex:.55;font-family:monospace;color:var(--accent2);text-align:right">'+(c.lh||0)+' L/H</div>'+
      '<div class="tc" style="flex:.5;font-family:monospace;color:var(--muted);text-align:right">'+ph+' €/h</div>'+
    '</div>';
  }).join('');
  el.innerHTML='<div class="tbl">'+
    '<div class="tr th">'+
      '<div class="tc" style="flex:1.1">Activo</div>'+
      '<div class="tc" style="flex:.9;text-align:right">Total L</div>'+
      '<div class="tc" style="flex:.65;text-align:right">Máx</div>'+
      '<div class="tc" style="flex:.65;text-align:right">Mín</div>'+
      '<div class="tc" style="flex:.55;text-align:right">L/H</div>'+
      '<div class="tc" style="flex:.5;text-align:right">€/h</div>'+
    '</div>'+rows+'</div>';
}

function openActivosModal(dbid){
  const modal=document.getElementById('activos-modal');
  document.getElementById('activos-modal-title').textContent=dbid?'Editar activo':'Nuevo activo';
  document.getElementById('act-dbid').value=dbid||'';
  if(dbid){
    const m=activosData.find(x=>x.id==dbid);if(!m)return;
    document.getElementById('act-code').value=m.Codigo||'';
    document.getElementById('act-name').value=m.Activo||'';
    document.getElementById('act-nserie').value=m.N_Serie||'';
    document.getElementById('act-tipo').value=m.tipoactivo||'';
    document.getElementById('act-modelo').value=m.modelo||'';
    document.getElementById('act-fab').value=m.fabricante||'';
    document.getElementById('act-prov').value=m.proveedor||'';
    document.getElementById('act-phora').value=m.phora||'';
    document.getElementById('act-ptonelada').value=m.ptonelada||'';
  }else{
    ['act-code','act-name','act-nserie','act-tipo','act-modelo','act-fab','act-prov','act-phora','act-ptonelada'].forEach(i=>{document.getElementById(i).value='';});
  }
  modal.classList.add('open');
}
function closeActivosModal(){document.getElementById('activos-modal').classList.remove('open');}
async function saveActivo(){
  const dbid=document.getElementById('act-dbid').value;
  const payload={
    Codigo:document.getElementById('act-code').value.trim().toUpperCase()||null,
    Activo:document.getElementById('act-name').value.trim()||null,
    N_Serie:document.getElementById('act-nserie').value.trim()||null,
    tipoactivo:document.getElementById('act-tipo').value.trim().toUpperCase()||null,
    modelo:document.getElementById('act-modelo').value.trim()||null,
    fabricante:document.getElementById('act-fab').value.trim().toUpperCase()||null,
    proveedor:document.getElementById('act-prov').value.trim()||null,
    phora:document.getElementById('act-phora').value.trim()||null,
    ptonelada:document.getElementById('act-ptonelada').value.trim()||null,
  };
  if(!payload.Codigo){alert('El Código no puede estar vacío');return;}
  try{
    let error;
    let result;
    if(dbid){
      result=await dbQuery({ action: 'update', table: 'tblactivos', data: payload, filters: [{ column: 'id', op: 'eq', value: dbid }] });
    }else{
      result=await dbQuery({ action: 'insert', table: 'tblactivos', data: payload });
    }
    if(!result.ok)throw new Error(result.error);
    closeActivosModal();
    cargarActivos();
  }catch(e){alert('Error al guardar: '+e.message);}
}
async function eliminarActivo(dbid,codigo){
  if(!confirm('¿Eliminar el activo "'+codigo+'"? Esta acción no se puede deshacer.'))return;
  try{
    let result;
    if(dbid){
      result=await dbQuery({ action: 'delete', table: 'tblactivos', filters: [{ column: 'id', op: 'eq', value: dbid }] });
    }else{
      result=await dbQuery({ action: 'delete', table: 'tblactivos', filters: [{ column: 'Codigo', op: 'eq', value: codigo }] });
    }
    if(!result.ok)throw new Error(result.error);
    cargarActivos();
  }catch(e){alert('Error al eliminar: '+e.message);}
}

// ── FICHAJE ───────────────────────────────────────────────────
const fst={workers:{},registros:[],vacaciones:{},bajas:{},diasLibres:{},extrasManual:{},calYear:new Date().getFullYear(),calMonth:new Date().getMonth(),editingId:null};
let frid=0;
WORKERS.forEach(n=>{fst.workers[n]={working:false,entradaTs:null,totalMs:0};fst.vacaciones[n]=[];fst.bajas[n]=[];fst.diasLibres[n]=[];fst.extrasManual[n]=[];});

function saveFst(){
  try{localStorage.setItem('arifoma_fst',JSON.stringify({registros:fst.registros,vacaciones:fst.vacaciones,bajas:fst.bajas,diasLibres:fst.diasLibres,extrasManual:fst.extrasManual,frid}));}catch(e){}
}
function loadFst(){
  try{
    const raw=localStorage.getItem('arifoma_fst');if(!raw)return;
    const d=JSON.parse(raw);
    if(d.frid)frid=d.frid;
    if(d.registros)fst.registros=d.registros;
    if(d.vacaciones)WORKERS.forEach(n=>{if(d.vacaciones[n])fst.vacaciones[n]=d.vacaciones[n];});
    if(d.bajas)WORKERS.forEach(n=>{if(d.bajas[n])fst.bajas[n]=d.bajas[n];});
    if(d.diasLibres)WORKERS.forEach(n=>{if(d.diasLibres[n])fst.diasLibres[n]=d.diasLibres[n];});
    if(d.extrasManual)WORKERS.forEach(n=>{if(d.extrasManual[n])fst.extrasManual[n]=d.extrasManual[n];});
  }catch(e){}
}

function parseFechaHora(fecha,hora){
  try{
    if(!fecha||!hora)return null;
    const d=String(fecha).split('/');
    const t=String(hora).split(':');
    if(d.length<3||t.length<2)return null;
    const year=parseInt(d[2])<100?2000+parseInt(d[2]):parseInt(d[2]);
    const month=parseInt(d[1])-1;
    const day=parseInt(d[0]);
    const hh=parseInt(t[0]);
    const mm=parseInt(t[1]);
    if(isNaN(year)||isNaN(month)||isNaN(day)||isNaN(hh)||isNaN(mm))return null;
    return new Date(year,month,day,hh,mm,0).getTime();
  }catch(e){return null;}
}
function parseFechaHoraStr(str){
  try{
    if(!str)return null;
    // Timestamp numérico
    const n=Number(str);if(!isNaN(n)&&n>1000000000000)return n;
    const s=String(str).trim();
    // ISO con timezone: "2026-04-07T05:00:00Z" o "2026-05-11T09:29:00+00:00"
    if(s.includes('T')&&(s.endsWith('Z')||/[+-]\d{2}:\d{2}$/.test(s))){const ts=new Date(s).getTime();if(!isNaN(ts))return ts;}
    const sp=s.split(' ');if(sp.length<2)return null;
    const dp=sp[0],tp=sp[1];
    // Formato ISO: "2026-04-01" (guiones)
    if(dp.includes('-')){
      const d=dp.split('-');const t=tp.split(':');
      if(d.length>=3&&t.length>=2)
        return new Date(parseInt(d[0]),parseInt(d[1])-1,parseInt(d[2]),parseInt(t[0]),parseInt(t[1]),0).getTime();
    }
    // Formato "dd/MM/yy" o "dd/MM/yyyy"
    return parseFechaHora(dp,tp);
  }catch(e){return null;}
}

function procesarFichajes(json){
  const hEl=document.getElementById('res-hist');
  if(!json.ok){if(hEl)hEl.innerHTML='<div class="tbl"><div class="empty">Error API: '+(json.error||'desconocido')+'</div></div>';return;}
    fst.registros=[];
    let maxId=0;
    json.data.forEach(r=>{
      const id=Number(r.id)||0;
      if(id>maxId)maxId=id;
      const nombre=WORKERS.find(w=>w.toUpperCase()===String(r.empleado||'').toUpperCase())||String(r.empleado||'');
      // Fecha base desde col C (siempre D/M/Y fiable). Extraer horas de col D/E o fallback a fentrada/fsalida
      const parseFechaBase=fs=>{
        const s=String(fs||'').trim();if(!s)return null;
        const n=Number(s);if(!isNaN(n)&&n>1e12){const d=new Date(n);return new Date(d.getUTCFullYear(),d.getUTCMonth(),d.getUTCDate());}
        if(s.includes('-')){const p=s.split('T')[0].split('-');if(p.length>=3)return new Date(parseInt(p[0]),parseInt(p[1])-1,parseInt(p[2]));}
        if(s.includes('/')){const p=s.split(' ')[0].split('/');if(p.length>=3){const yy=parseInt(p[2])<100?2000+parseInt(p[2]):parseInt(p[2]);return new Date(yy,parseInt(p[1])-1,parseInt(p[0]));}}
        return null;
      };
      const extraerHora=v=>{
        if(v instanceof Date)return[v.getHours(),v.getMinutes()];
        const s=String(v||'').trim();if(!s)return null;
        const n=Number(s);if(!isNaN(n)&&n>1e12){const d=new Date(n);return[d.getHours(),d.getMinutes()];}
        const mT=s.match(/(\d{1,2}):(\d{2})/);if(mT)return[parseInt(mT[1]),parseInt(mT[2])];
        return null;
      };
      const fBase=parseFechaBase(r.fecha);
      const tieneHoraEntrada=String(r.entrada||'').trim()!=='';
      const tieneHoraSalida=String(r.salida||'').trim()!=='';
      let tsE=null,tsS=null;
      // Caso 1: cols D/E con horas → usar fBase + esas horas
      if(fBase&&tieneHoraEntrada){
        const hE=extraerHora(r.entrada),hS=tieneHoraSalida?extraerHora(r.salida):null;
        if(hE){const d=new Date(fBase);d.setHours(hE[0],hE[1],0,0);tsE=d.getTime();}
        if(hS){const d=new Date(fBase);d.setHours(hS[0],hS[1],0,0);if(tsE&&d.getTime()<tsE)d.setDate(d.getDate()+1);tsS=d.getTime();}
      }
      // Caso 2: sin horas en D/E pero con tiempodia → fallback a 8:00 + duración
      if((!tsE||!tsS)&&fBase&&r.tiempodia){
        const horas=Number(String(r.tiempodia).replace(',','.'));
        if(!isNaN(horas)&&horas>0){
          const d=new Date(fBase);d.setHours(8,0,0,0);
          tsE=d.getTime()+(id%60)*1000;
          tsS=tsE+Math.round(horas*3600000);
        }
      }
      // Caso 3: último recurso, parsear fentrada/fsalida directos
      if(!tsE)tsE=parseFechaHoraStr(r.fentrada);
      if(!tsS)tsS=parseFechaHoraStr(r.fsalida);
      // Fallback: fila manual sin fentrada/fsalida pero con fecha+tiempodia
      if((!tsE||!tsS)&&r.tiempodia){
        const horas=Number(String(r.tiempodia).replace(',','.'));
        if(!isNaN(horas)&&horas>0){
          let dBase=null;
          const fs=String(r.fecha||'').trim();
          const nf=Number(fs);
          if(!isNaN(nf)&&nf>1000000000000)dBase=new Date(nf);
          else if(fs.includes('-')){const p=fs.split('T')[0].split('-');if(p.length>=3)dBase=new Date(parseInt(p[0]),parseInt(p[1])-1,parseInt(p[2]),8,0,0);}
          else if(fs.includes('/')){const p=fs.split('/');if(p.length>=3){const yy=parseInt(p[2])<100?2000+parseInt(p[2]):parseInt(p[2]);dBase=new Date(yy,parseInt(p[1])-1,parseInt(p[0]),8,0,0);}}
          if(dBase){
            dBase.setHours(8,0,0,0);
            // desfase por id para que filas del mismo día no colisionen en ts
            tsE=dBase.getTime()+(id%60)*1000;
            tsS=tsE+Math.round(horas*3600000);
          }
        }
      }
      if(tsE)fst.registros.push({id:id*2,nombre,tipo:'entrada',ts:tsE});
      if(tsS&&tsE)fst.registros.push({id:id*2+1,nombre,tipo:'salida',ts:tsS,duracion:tsS-tsE});
    });
if(fst.registros.length===0&&json.data.length>0){
      const sample=json.data[0];
      if(hEl)hEl.innerHTML=`<div class="tbl"><div class="empty" style="color:var(--danger)">Sin datos parseados. fentrada="${sample.fentrada}" fecha="${sample.fecha}" entrada="${sample.entrada}"</div></div>`;
      return;
    }
    frid=maxId*2+100;
    WORKERS.forEach(n=>recalcWorker(n));
    saveFst();
    WORKERS.forEach(n=>renderWcard(n));
    renderStats();
    renderMeses();renderHistorial();
    const pg=document.querySelector('.page.active');
    if(pg&&pg.id==='pg-editar')renderEditar();
}
async function cargarFichajes(){
  const hEl=document.getElementById('res-hist');
  if(hEl)hEl.innerHTML='<div class="tbl"><div class="empty">Cargando fichajes...</div></div>';
  try{const json=await apiFetch('?accion=fichajes');procesarFichajes(json);}catch(e){console.warn('Error cargando fichajes:',e);}
}

function procesarAusencias(json){
  if(!json.ok||!json.data||!json.data.length)return;
  WORKERS.forEach(n=>{fst.vacaciones[n]=[];fst.bajas[n]=[];fst.diasLibres[n]=[];fst.extrasManual[n]=[];});
  function toDateStr(v){if(!v)return '';const s=String(v);if(s.includes('T'))return s.substring(0,10);return s;}
  let maxId=frid;
  json.data.forEach(r=>{
    const w=r.trabajador;
    if(!w||!WORKERS.includes(w))return;
    const id=Number(r.id)||0;
    if(id>=maxId)maxId=id+1;
    const st=toDateStr(r.start),en=toDateStr(r.end);
    if(r.tipo==='vacaciones'){fst.vacaciones[w].push({id,start:st,end:en,dias:Number(r.dias)||0});}
    else if(r.tipo==='baja'){fst.bajas[w].push({id,tipo:r.subtipo||'',start:st,end:en||null,dias:Number(r.dias)||0});}
    else if(r.tipo==='libre'){fst.diasLibres[w].push({id,tipo:r.subtipo||'',start:st,end:en,dias:Number(r.dias)||0});}
    else if(r.tipo==='extra'){fst.extrasManual[w].push({id,fecha:st,horas:Number(r.horas)||0,motivo:r.motivo||''});}
  });
  frid=maxId;
  saveFst();
  renderVac();renderBajas();renderDiasLibres();renderExtrasManual();renderCal();
}

async function verificarFichajePendiente(nombre) {
  const hoy = new Date().toISOString().slice(0, 10);
  const result = await dbQuery({ action: 'select', table: 'tblFichaje',
    filters: [{ column: 'empleado', op: 'eq', value: nombre.toUpperCase() }, { column: 'fecha', op: 'eq', value: hoy }],
    options: { select: 'entrada,salida' }
  });

  if (result.data && result.data.length > 0 && result.data[0].entrada && !result.data[0].salida) {
    return { bloqueado: true, motivo: 'Ya fichaste hoy. Debes desfichar primero.' };
  }
  return { bloqueado: false };
}

function handleFichar(nombre) {
  ficharWorker(nombre).catch(e => console.error('handleFichar error:', e));
}

async function ficharWorker(nombre){
  const w=fst.workers[nombre];const now=Date.now();
  if(!w.working){
    // Verificar si ya existe fichaje pendiente hoy
    const resultado = await verificarFichajePendiente(nombre);
    if (resultado && resultado.bloqueado) {
      alert(resultado.motivo);
      return;
    }
    w.working=true;w.entradaTs=now;
    fst.registros.push({id:frid++,nombre,tipo:'entrada',ts:now});
    enviarEntrada(nombre,now);
    saveFst();renderWcard(nombre);renderStats();
  }else{
    if(!w.entradaTs){console.warn('entradaTs nulo para',nombre);return;}
    const tsE=w.entradaTs; // capturar ANTES de que recalcWorker lo resetee
    const ms=now-tsE;
    if(ms<3000)return; // doble-click accidental
    w.totalMs+=ms;w.working=false;
    fst.registros.push({id:frid++,nombre,tipo:'salida',ts:now,duracion:ms});
    recalcWorker(nombre);
    enviarSalida(nombre,tsE,now,ms);
    saveFst();renderWcard(nombre);renderStats();
  }
}
async function enviarEntrada(nombre,tsE){
  const _d=new Date(tsE);const _fecha=_d.getFullYear()+'-'+pad(_d.getMonth()+1)+'-'+pad(_d.getDate());
  const payload={tipo:'fichajeEntrada',empleado:nombre.toUpperCase(),fecha:_fecha,entrada:fmtHM(tsE),fentrada:fmtFechaHora(tsE)};
  try{
    const result = await apiPost(payload);
    if (!result.ok) console.error('enviarEntrada error:', result.error);
  } catch(e) {
    console.error('enviarEntrada exception:', e.message);
  }
}
async function enviarSalida(nombre,tsE,tsS,ms){
  const payload={tipo:'fichajesSalida',empleado:nombre.toUpperCase(),fentrada:fmtFechaHora(tsE),salida:fmtHM(tsS),fsalida:fmtFechaHora(tsS),tiempodia:Math.round(ms/60000)/60};
  try{
    const result = await apiPost(payload);
    if (!result.ok) console.error('enviarSalida error:', result.error);
  } catch(e) {
    console.error('enviarSalida exception:', e.message);
  }
}


function recalcWorker(nombre){
  // totalMs = solo hoy (para el panel de fichaje)
  const now=new Date();const todayStart=new Date(now.getFullYear(),now.getMonth(),now.getDate()).getTime();
  let total=0,lastE=null;
  const todayRegs=fst.registros.filter(r=>r.nombre===nombre&&r.ts>=todayStart&&r.ts<todayStart+86400000);
  todayRegs.sort((a,b)=>a.ts-b.ts||a.id-b.id)
    .forEach(r=>{if(r.tipo==='entrada')lastE=r.ts;else if(lastE!==null){total+=safeDur(lastE,r.ts);lastE=null;}});
  fst.workers[nombre].totalMs=total;
  // Estado working: si hay entrada hoy sin salida
  if(lastE!==null){fst.workers[nombre].working=true;fst.workers[nombre].entradaTs=lastE;}
  else{fst.workers[nombre].working=false;fst.workers[nombre].entradaTs=null;}
}
function safeDur(tsE,tsS){
  let dur=tsS-tsE;
  if(dur<0)dur+=86400000; // turno nocturno: añadir 24h
  if(dur<=0||dur>86400000*1.5)return 0; // descartar si sigue inválido (>36h)
  return dur;
}
function getMsMonth(nombre,y,m){
  const monthStart=new Date(y,m,1).getTime();const monthEnd=new Date(y,m+1,1).getTime();
  let total=0,lastE=null;
  fst.registros.filter(r=>r.nombre===nombre&&r.ts>=monthStart&&r.ts<monthEnd)
    .sort((a,b)=>a.ts-b.ts||a.id-b.id)
    .forEach(r=>{
      if(r.tipo==='entrada')lastE=r.ts;
      else if(lastE!==null){total+=safeDur(lastE,r.ts);lastE=null;}
    });
  if(lastE&&fst.workers[nombre].working)total+=Date.now()-lastE;
  return total;
}
function renderWcard(nombre){
  const w=fst.workers[nombre];
  const liveMs=w.working?(Date.now()-w.entradaTs):0;
  const total=w.totalMs+liveMs;
  const c=document.getElementById('wc-'+nombre);
  if(!c) return;
  c.className='wcard'+(w.working?' on':'');
  c.querySelector('.wst').innerHTML=w.working?'<span class="ldot"></span>Desde '+fmtH(w.entradaTs):'Sin fichar';
  c.querySelector('.wtime').textContent=total>0?'Hoy: '+fmtDur(total):'';
  const btn = c.querySelector('.wbtn');
  if(btn) btn.textContent=w.working?'Registrar salida':'Registrar entrada';
}
function renderStats(){
  const act=WORKERS.filter(n=>fst.workers[n].working).length;
  const now=new Date();const todayStart=new Date(now.getFullYear(),now.getMonth(),now.getDate()).getTime();
  const ficHoy=fst.registros.filter(r=>r.ts>=todayStart&&r.ts<todayStart+86400000).length;
  let ms=0;WORKERS.forEach(n=>{const w=fst.workers[n];ms+=w.totalMs+(w.working?Date.now()-w.entradaTs:0);});
  document.getElementById('i-act').textContent=act;
  document.getElementById('i-fich').textContent=ficHoy;
  document.getElementById('i-hrs').textContent=fmtDur(ms);
}
function renderWgrid(){
  document.getElementById('wgrid').innerHTML=WORKERS.map(n=>`<div class="wcard" id="wc-${n}"><div class="wname">${n}</div><div class="wst">Sin fichar</div><div class="wtime"></div><button class="wbtn" data-worker="${n}">Registrar entrada</button></div>`).join('');
}
function renderResumenCards(){
  const resGrid=document.getElementById('res-grid');
  if(!resGrid)return;
  resGrid.innerHTML=WORKERS.map(n=>{
    const w=fst.workers[n];const liveMs=w.working?(Date.now()-w.entradaTs):0;const total=w.totalMs+liveMs;
    const rw=fst.registros.filter(r=>r.nombre===n);
    return `<div class="res-card"><div class="res-name">${w.working?'<span class="res-on"></span>':''}${n}</div><div class="res-row"><span class="res-lbl">Horas hoy</span><span class="res-val">${fmtDur(total)}</span></div><div class="res-row"><span class="res-lbl">Entradas</span><span class="res-val">${rw.filter(r=>r.tipo==='entrada').length}</span></div><div class="res-row"><span class="res-lbl">Salidas</span><span class="res-val">${rw.filter(r=>r.tipo==='salida').length}</span></div><div class="res-row"><span class="res-lbl">Estado</span><span class="res-val">${w.working?'Trabajando':'Fuera'}</span></div></div>`;
  }).join('');
}
// Horas laborables del mes según convenio: días_lab × 8h (uniforme para todos)
function getExpectedMonth(nombre,y,m){
  if(y===2026&&DIAS_LAB_2026[m]!=null)return DIAS_LAB_2026[m]*HORAS_DIA_STD*3600000;
  // Fallback: calcular desde festivos con 8h/día Mon-Fri
  let dias=0;const dim=new Date(y,m+1,0).getDate();
  for(let d=1;d<=dim;d++){
    const dow=new Date(y,m,d).getDay();
    if(dow===0||dow===6)continue;
    const ds=y+'-'+pad(m+1)+'-'+pad(d);
    if(FESTIVOS.includes(ds))continue;
    dias++;
  }
  return dias*HORAS_DIA_STD*3600000;
}
// Horas de una lista (vacaciones/bajas/libres): 8h/día (todos los días del rango). start y end inclusivos
function _msListaMes(lista,y,m,hpd){
  hpd=hpd||HORAS_DIA_STD;
  let ms=0;
  (lista||[]).forEach(item=>{
    if(!item.start)return;
    let cur=new Date(item.start+'T00:00:00');
    const fin=item.end?new Date(item.end+'T00:00:00'):new Date(item.start+'T00:00:00');
    while(cur<=fin){
      if(cur.getFullYear()===y&&cur.getMonth()===m)ms+=hpd*3600000;
      cur.setDate(cur.getDate()+1);
    }
  });
  return ms;
}
const _HPD={'Antonio Juan Martel':10,'Rubén Díaz':10};
function getVacacionesMs(n,y,m){return _msListaMes(fst.vacaciones[n],y,m,_HPD[n]);}
function getBajasMs(n,y,m){return _msListaMes(fst.bajas[n],y,m,_HPD[n]);}
function getLibresMs(n,y,m){return _msListaMes(fst.diasLibres[n],y,m,_HPD[n]);}
function getAusenciasMs(nombre,y,m){return getVacacionesMs(nombre,y,m)+getBajasMs(nombre,y,m)+getLibresMs(nombre,y,m);}
// Total del mes = trabajado + ausencias justificadas + ajustes manuales
function getMsMonthTotal(nombre,y,m){return getMsMonth(nombre,y,m)+getAusenciasMs(nombre,y,m)+getExtrasManualMs(nombre,y,m);}
// Ajustes manuales de extras en el mes (en ms)
function getExtrasManualMs(nombre,y,m){
  return (fst.extrasManual[nombre]||[]).filter(x=>{const d=new Date(x.fecha+'T00:00:00');return d.getFullYear()===y&&d.getMonth()===m;}).reduce((a,x)=>a+x.horas*3600000,0);
}
// Extras = (realizadas + libre + baja + manual) - horas laborables del mes
function getExtrasMonth(nombre,y,m){return getMsMonthTotal(nombre,y,m)-getExpectedMonth(nombre,y,m);}
function getExtrasTotal(nombre){const y=new Date().getFullYear();let t=0;for(let m=0;m<=11;m++)t+=getExtrasMonth(nombre,y,m);return t;}
function fmtExtra(ms){if(ms===0)return'0h';const sign=ms<0?'-':'+';const abs=Math.abs(ms);const h=Math.floor(abs/3600000);const mn=Math.floor((abs%3600000)/60000);return sign+(mn?h+'h'+pad(mn)+'m':h+'h');}
function renderMeses(){
  const fw=document.getElementById('filt-mes-w').value;const y=new Date().getFullYear();const cm=new Date().getMonth();
  let header=fw?`<div class="mtr mth"><div class="mtc mn">Mes</div><div class="mtc mh">Esperadas</div><div class="mtc mh">Fichadas</div><div class="mtc mh">Vacac.</div><div class="mtc mh">Bajas</div><div class="mtc mh">Libres</div><div class="mtc mh">Manual</div><div class="mtc mh">Total</div><div class="mtc mh">Extras</div></div>`:`<div class="mtr mth"><div class="mtc mn">Mes</div>${WORKERS.map(n=>`<div class="mtc mh" style="flex:.7;font-size:.65rem">${n.split(' ')[0]}</div>`).join('')}<div class="mtc mf">Total</div></div>`;
  let rows='';let gt=0;let gx=0;let gfi=0;let gva=0;let gba=0;let gli=0;let gma=0;let gesp=0;const wTotals=WORKERS.reduce((o,n)=>(o[n]=0,o),{});
  for(let m=0;m<=cm;m++){
    if(fw){
      const esp=getExpectedMonth(fw,y,m);const fich=getMsMonth(fw,y,m);const vac=getVacacionesMs(fw,y,m);const baj=getBajasMs(fw,y,m);const lib=getLibresMs(fw,y,m);const manual=getExtrasManualMs(fw,y,m);
      const ms=fich+vac+baj+lib+manual;const ex=getExtrasMonth(fw,y,m);
      gesp+=esp;gt+=ms;gx+=ex;gfi+=fich;gva+=vac;gba+=baj;gli+=lib;gma+=manual;const exCls=ex>0?'extra-pos':ex<0?'extra-neg':'';const maCls=manual>0?'extra-pos':manual<0?'extra-neg':'';
      rows+=`<div class="mtr"><div class="mtc mn">${MESES[m]}</div><div class="mtc mh">${fmtDurDec(esp)}</div><div class="mtc mh">${fich>0?fmtDurDec(fich):'—'}</div><div class="mtc mh">${vac>0?fmtDurDec(vac):'—'}</div><div class="mtc mh">${baj>0?fmtDurDec(baj):'—'}</div><div class="mtc mh">${lib>0?fmtDurDec(lib):'—'}</div><div class="mtc mh"><span class="${maCls}">${manual!==0?fmtExtra(manual):'—'}</span></div><div class="mtc mh">${ms>0?fmtDurDec(ms):'—'}</div><div class="mtc mh"><span class="${exCls}">${(ms>0||ex!==0)?fmtExtra(ex):'—'}</span></div></div>`;
    }else{
      let rt=0;const celdas=WORKERS.map(n=>{const ms=getMsMonthTotal(n,y,m);rt+=ms;wTotals[n]+=ms;return`<div class="mtc mh" style="flex:.7">${ms>0?fmtDurDec(ms):'—'}</div>`;}).join('');
      gt+=rt;rows+=`<div class="mtr"><div class="mtc mn">${MESES[m]}</div>${celdas}<div class="mtc mf">${rt>0?fmtDurDec(rt):'—'}</div></div>`;
    }
  }
  const totRow=fw?`<div class="mtr tot"><div class="mtc mn">Total</div><div class="mtc mh">${fmtDurDec(gesp)}</div><div class="mtc mh">${fmtDurDec(gfi)}</div><div class="mtc mh">${fmtDurDec(gva)}</div><div class="mtc mh">${fmtDurDec(gba)}</div><div class="mtc mh">${fmtDurDec(gli)}</div><div class="mtc mh"><span class="${gma>0?'extra-pos':gma<0?'extra-neg':''}">${gma!==0?fmtExtra(gma):'—'}</span></div><div class="mtc mh">${fmtDurDec(gt)}</div><div class="mtc mh"><span class="${gx>0?'extra-pos':gx<0?'extra-neg':''}">${fmtExtra(gx)}</span></div></div>`:`<div class="mtr tot"><div class="mtc mn">Total</div>${WORKERS.map(n=>`<div class="mtc mh" style="flex:.7">${wTotals[n]>0?fmtDurDec(wTotals[n]):'—'}</div>`).join('')}<div class="mtc mf">${fmtDurDec(gt)}</div></div>`;
  document.getElementById('mes-tabla').innerHTML=`<div class="month-table">${header}${rows}${totRow}</div>`;
}
function renderExtras(){
  const y=new Date().getFullYear();
  const rows=WORKERS.map(n=>{
    let acc=0;let total=0;
    for(let m=0;m<=11;m++){
      const ex=getExtrasMonth(n,y,m)+acc;
      if(ex>=0){total+=ex;acc=0;}else{acc=ex;}
    }
    const cls=total>0?'extra-pos':'';return`<div class="extras-row"><span style="font-weight:600">${n}</span><span class="${cls}">${fmtExtra(total)}</span></div>`;
  }).join('');
  document.getElementById('extras-sum').innerHTML=`<div class="extras-sum">${rows}</div>`;
}
function renderHistorial(){
  const f=document.getElementById('filt-w').value;const h=document.getElementById('res-hist');
  const byJid=new Map();
  fst.registros.forEach(r=>{const jid=Math.floor(r.id/2);if(!byJid.has(jid))byJid.set(jid,{});byJid.get(jid)[r.tipo]=r;});
  const jornadas=[...byJid.entries()]
    .filter(([,v])=>v.entrada&&(!f||v.entrada.nombre===f))
    .map(([jid,v])=>({jid,nombre:v.entrada.nombre,fecha:fmtFecha(v.entrada.ts),entrada:fmtH(v.entrada.ts),salida:v.salida?fmtH(v.salida.ts):'—',duracion:v.salida?fmtDur(safeDur(v.entrada.ts,v.salida.ts)):'—'}))
    .sort((a,b)=>b.jid-a.jid);
  if(!jornadas.length){h.innerHTML='<div class="tbl"><div class="empty">Sin registros aún</div></div>';return;}
  h.innerHTML=`<div class="tbl"><div class="tr th"><div class="tc" style="flex:1.2;font-weight:700">Trabajador</div><div class="tc" style="flex:.7">Fecha</div><div class="tc" style="flex:.8">Entrada</div><div class="tc" style="flex:.8">Salida</div><div class="tc" style="flex:.8">Duración</div></div>${jornadas.map(j=>`<div class="tr"><div class="tc" style="flex:1.2;font-size:.82rem">${j.nombre}</div><div class="tc" style="flex:.7;color:var(--muted);font-size:.75rem">${j.fecha}</div><div class="tc" style="flex:.8;color:var(--accent2);font-family:'DM Mono',monospace;font-size:.82rem">${j.entrada}</div><div class="tc" style="flex:.8;color:var(--danger);font-family:'DM Mono',monospace;font-size:.82rem">${j.salida}</div><div class="tc" style="flex:.8;color:var(--muted);font-family:'DM Mono',monospace;font-size:.82rem">${j.duracion}</div></div>`).join('')}</div>`;
}

// VAC
// s y e inclusivos. Cuenta TODOS los días del rango (incluye findes y festivos)
function calcWorkDays(s,e){let c=0,cur=new Date(s+'T00:00:00');const end=new Date(e+'T00:00:00');while(cur<=end){c++;cur.setDate(cur.getDate()+1);}return c;}
function usedDays(n){return fst.vacaciones[n].reduce((a,v)=>a+v.dias,0)}
let _editAus=null; // {tipo:'vac'|'baja'|'libre'|'extra', worker:string, id:number}
const _btnEdit='style="background:transparent;border:none;cursor:pointer;color:var(--accent);font-size:.75rem;margin-right:2px"';

function openVacModal(editWorker,editId){
  _editAus=null;
  document.getElementById('vm-start').value=dateStr(new Date());document.getElementById('vm-end').value=dateStr(new Date());document.getElementById('vm-warn').textContent='';
  if(editWorker&&editId!=null){
    const v=fst.vacaciones[editWorker].find(x=>x.id===editId);
    if(v){_editAus={tipo:'vac',worker:editWorker,id:editId};document.getElementById('vm-worker').value=editWorker;document.getElementById('vm-start').value=v.start;document.getElementById('vm-end').value=v.end;}
  }
  document.getElementById('vac-modal').classList.add('open');
}
function closeVacModal(){document.getElementById('vac-modal').classList.remove('open');_editAus=null;}
async function saveVac(){const w=document.getElementById('vm-worker').value,s=document.getElementById('vm-start').value,e=document.getElementById('vm-end').value;if(!w||!s||!e)return;const dias=calcWorkDays(s,e);
  if(_editAus&&_editAus.tipo==='vac'){
    const old=fst.vacaciones[_editAus.worker].find(x=>x.id===_editAus.id);
    const rest=TOTAL_VAC-usedDays(w)+(old?old.dias:0);
    if(dias>rest){document.getElementById('vm-warn').textContent='Solo quedan '+rest+' días para '+w+'.';return;}
    if(old){old.start=s;old.end=e;old.dias=dias;}
    saveFst();closeVacModal();renderVac();renderCal();
    try{apiPost({tipo:'editAusencia',id:_editAus.id,start:s,end:e,dias});}catch(e){}
    return;
  }
  const rest=TOTAL_VAC-usedDays(w);if(dias>rest){document.getElementById('vm-warn').textContent='Solo quedan '+rest+' días para '+w+'.';return;}
  let id=frid++;
  try{const r=await apiPost({tipo:'ausencia',categoria:'vacaciones',trabajador:w,start:s,end:e,dias});if(r.ok)id=r.id;}catch(e){}
  fst.vacaciones[w].push({id,start:s,end:e,dias});saveFst();closeVacModal();renderVac();renderCal();}
function delVac(n,id){fst.vacaciones[n]=fst.vacaciones[n].filter(v=>v.id!==id);saveFst();renderVac();renderCal();
  try{apiPost({tipo:'delAusencia',id});}catch(e){}}
function renderVac(){
  document.getElementById('vac-grid').innerHTML=WORKERS.map(n=>{
    const used=usedDays(n);const pct=Math.min(100,Math.round(used/TOTAL_VAC*100));
    const items=fst.vacaciones[n].map(v=>`<div class="vac-item"><span>${fmtDate(v.start)} – ${fmtDate(v.end)} (${v.dias}d)</span><span><button onclick="openVacModal('${n}',${v.id})" ${_btnEdit}>&#9998;</button><button onclick="delVac('${n}',${v.id})" style="background:transparent;border:none;cursor:pointer;color:var(--danger);font-size:.82rem">✕</button></span></div>`).join('');
    return `<div class="vac-card"><div style="font-size:.86rem;font-weight:600;color:var(--text);margin-bottom:6px">${n}</div><div style="font-size:1.4rem;font-weight:700;color:var(--text);font-family:'DM Mono',monospace">${TOTAL_VAC-used}<span style="font-size:.82rem;font-weight:400;color:var(--muted)"> días restantes</span></div><div style="font-size:.68rem;color:var(--muted);margin-bottom:8px">${used} de ${TOTAL_VAC} días usados</div><div class="vac-bar"><div class="vac-bar-fill" style="width:${pct}%"></div></div>${items?'<div>'+items+'</div>':'<div style="font-size:.68rem;color:var(--muted)">Sin vacaciones registradas</div>'}</div>`;
  }).join('');
}

// BAJAS
function openBajaModal(editWorker,editId){
  _editAus=null;
  document.getElementById('bm-start').value=dateStr(new Date());document.getElementById('bm-end').value='';
  if(editWorker&&editId!=null){
    const b=fst.bajas[editWorker].find(x=>x.id===editId);
    if(b){_editAus={tipo:'baja',worker:editWorker,id:editId};document.getElementById('bm-worker').value=editWorker;document.getElementById('bm-tipo').value=b.tipo||'';document.getElementById('bm-start').value=b.start;document.getElementById('bm-end').value=b.end||'';}
  }
  document.getElementById('baja-modal').classList.add('open');
}
function closeBajaModal(){document.getElementById('baja-modal').classList.remove('open');_editAus=null;}
async function saveBaja(){const w=document.getElementById('bm-worker').value,tipo=document.getElementById('bm-tipo').value,s=document.getElementById('bm-start').value,e=document.getElementById('bm-end').value;if(!w||!s)return;const dias=e?calcWorkDays(s,e):null;
  if(_editAus&&_editAus.tipo==='baja'){
    const old=fst.bajas[_editAus.worker].find(x=>x.id===_editAus.id);
    if(old){old.tipo=tipo;old.start=s;old.end=e||null;old.dias=dias;}
    saveFst();closeBajaModal();renderBajas();
    try{apiPost({tipo:'editAusencia',id:_editAus.id,start:s,end:e||'',dias:dias||0,subtipo:tipo});}catch(e){}
    return;
  }
  let id=frid++;
  try{const r=await apiPost({tipo:'ausencia',categoria:'baja',trabajador:w,start:s,end:e||'',dias:dias||0,subtipo:tipo});if(r.ok)id=r.id;}catch(e){}
  fst.bajas[w].push({id,tipo,start:s,end:e||null,dias});saveFst();closeBajaModal();renderBajas();}
function delBaja(n,id){fst.bajas[n]=fst.bajas[n].filter(b=>b.id!==id);saveFst();renderBajas();
  try{apiPost({tipo:'delAusencia',id});}catch(e){}}
function renderBajas(){
  document.getElementById('baja-grid').innerHTML=WORKERS.map(n=>{
    const items=fst.bajas[n].map(b=>`<div class="vac-item"><span style="color:var(--danger)">${b.tipo}</span><span style="margin-left:4px">${fmtDate(b.start)}${b.end?' – '+fmtDate(b.end):'  (en curso)'}${b.dias?' ('+b.dias+'d)':''}</span><span><button onclick="openBajaModal('${n}',${b.id})" ${_btnEdit}>&#9998;</button><button onclick="delBaja('${n}',${b.id})" style="background:transparent;border:none;cursor:pointer;color:var(--danger);font-size:.82rem">✕</button></span></div>`).join('');
    const total=fst.bajas[n].reduce((a,b)=>a+(b.dias||0),0);
    return `<div class="baja-card"><div style="font-size:.86rem;font-weight:600;color:var(--text);margin-bottom:6px">${n}</div><div style="font-size:1.1rem;font-weight:700;color:var(--danger);font-family:'DM Mono',monospace">${total} días<span style="font-size:.72rem;font-weight:400;color:var(--muted)"> este año</span></div>${items?'<div style="margin-top:8px">'+items+'</div>':'<div style="font-size:.68rem;color:var(--muted);margin-top:6px">Sin bajas registradas</div>'}</div>`;
  }).join('');
}

// DÍAS LIBRES
function openLibreModal(editWorker,editId){
  _editAus=null;
  document.getElementById('lm-start').value=dateStr(new Date());document.getElementById('lm-end').value=dateStr(new Date());
  if(editWorker&&editId!=null){
    const l=fst.diasLibres[editWorker].find(x=>x.id===editId);
    if(l){_editAus={tipo:'libre',worker:editWorker,id:editId};document.getElementById('lm-worker').value=editWorker;document.getElementById('lm-tipo').value=l.tipo||'';document.getElementById('lm-start').value=l.start;document.getElementById('lm-end').value=l.end||'';}
  }
  document.getElementById('libre-modal').classList.add('open');
}
function closeLibreModal(){document.getElementById('libre-modal').classList.remove('open');_editAus=null;}
async function saveLibre(){const w=document.getElementById('lm-worker').value,tipo=document.getElementById('lm-tipo').value,s=document.getElementById('lm-start').value,e=document.getElementById('lm-end').value;if(!w||!s||!e)return;const dias=calcWorkDays(s,e);
  if(_editAus&&_editAus.tipo==='libre'){
    const old=fst.diasLibres[_editAus.worker].find(x=>x.id===_editAus.id);
    if(old){old.tipo=tipo;old.start=s;old.end=e;old.dias=dias;}
    saveFst();closeLibreModal();renderDiasLibres();
    try{apiPost({tipo:'editAusencia',id:_editAus.id,start:s,end:e,dias,subtipo:tipo});}catch(e){}
    return;
  }
  let id=frid++;
  try{const r=await apiPost({tipo:'ausencia',categoria:'libre',trabajador:w,start:s,end:e,dias,subtipo:tipo});if(r.ok)id=r.id;}catch(e){}
  fst.diasLibres[w].push({id,tipo,start:s,end:e,dias});saveFst();closeLibreModal();renderDiasLibres();}
function delLibre(n,id){fst.diasLibres[n]=fst.diasLibres[n].filter(l=>l.id!==id);saveFst();renderDiasLibres();
  try{apiPost({tipo:'delAusencia',id});}catch(e){}}
function renderDiasLibres(){
  document.getElementById('libre-grid').innerHTML=WORKERS.map(n=>{
    const items=fst.diasLibres[n].map(l=>`<div class="vac-item"><span style="color:#f5c842">${l.tipo}</span><span style="margin-left:4px">${fmtDate(l.start)}${l.end!==l.start?' – '+fmtDate(l.end):''} (${l.dias}d)</span><span><button onclick="openLibreModal('${n}',${l.id})" ${_btnEdit}>&#9998;</button><button onclick="delLibre('${n}',${l.id})" style="background:transparent;border:none;cursor:pointer;color:var(--danger);font-size:.82rem">✕</button></span></div>`).join('');
    const total=fst.diasLibres[n].reduce((a,l)=>a+l.dias,0);
    return `<div class="baja-card" style="border-color:rgba(107,125,46,.3)"><div style="font-size:.86rem;font-weight:600;color:var(--text);margin-bottom:6px">${n}</div><div style="font-size:1.1rem;font-weight:700;color:#f5c842;font-family:'DM Mono',monospace">${total} días<span style="font-size:.72rem;font-weight:400;color:var(--muted)"> concedidos</span></div>${items?'<div style="margin-top:8px">'+items+'</div>':'<div style="font-size:.68rem;color:var(--muted);margin-top:6px">Sin días libres registrados</div>'}</div>`;
  }).join('');
}

// EXTRAS MANUALES
function openExtraModal(editWorker,editId){
  _editAus=null;
  document.getElementById('xm-fecha').value=dateStr(new Date());document.getElementById('xm-horas').value='';document.getElementById('xm-motivo').value='';
  if(editWorker&&editId!=null){
    const x=(fst.extrasManual[editWorker]||[]).find(v=>v.id===editId);
    if(x){_editAus={tipo:'extra',worker:editWorker,id:editId};document.getElementById('xm-worker').value=editWorker;document.getElementById('xm-fecha').value=x.fecha;document.getElementById('xm-horas').value=x.horas;document.getElementById('xm-motivo').value=x.motivo||'';}
  }
  document.getElementById('extra-modal').classList.add('open');
}
function closeExtraModal(){document.getElementById('extra-modal').classList.remove('open');_editAus=null;}
async function saveExtraManual(){const w=document.getElementById('xm-worker').value,f=document.getElementById('xm-fecha').value,h=parseFloat(document.getElementById('xm-horas').value),mo=document.getElementById('xm-motivo').value;if(!w||!f||isNaN(h))return;
  if(_editAus&&_editAus.tipo==='extra'){
    const old=(fst.extrasManual[_editAus.worker]||[]).find(x=>x.id===_editAus.id);
    if(old){old.fecha=f;old.horas=h;old.motivo=mo;}
    saveFst();closeExtraModal();renderExtrasManual();renderMeses();renderExtras();
    try{apiPost({tipo:'editAusencia',id:_editAus.id,start:f,end:f,horas:h,motivo:mo});}catch(e){}
    return;
  }
  let id=frid++;
  try{const r=await apiPost({tipo:'ausencia',categoria:'extra',trabajador:w,start:f,end:f,dias:0,horas:h,motivo:mo});if(r.ok)id=r.id;}catch(e){}
  fst.extrasManual[w].push({id,fecha:f,horas:h,motivo:mo});saveFst();closeExtraModal();renderExtrasManual();renderMeses();renderExtras();}
function delExtraManual(n,id){fst.extrasManual[n]=fst.extrasManual[n].filter(x=>x.id!==id);saveFst();renderExtrasManual();renderMeses();renderExtras();
  try{apiPost({tipo:'delAusencia',id});}catch(e){}}
function renderExtrasManual(){
  document.getElementById('extra-grid').innerHTML=WORKERS.map(n=>{
    const list=fst.extrasManual[n]||[];
    const items=list.slice().sort((a,b)=>b.fecha.localeCompare(a.fecha)).map(x=>`<div class="vac-item"><span style="color:${x.horas>=0?'var(--accent2)':'var(--danger)'}">${x.horas>=0?'+':''}${x.horas}h</span><span style="margin-left:4px">${fmtDate(x.fecha)}${x.motivo?' · '+x.motivo:''}</span><span><button onclick="openExtraModal('${n}',${x.id})" ${_btnEdit}>&#9998;</button><button onclick="delExtraManual('${n}',${x.id})" style="background:transparent;border:none;cursor:pointer;color:var(--danger);font-size:.82rem">✕</button></span></div>`).join('');
    const total=list.reduce((a,x)=>a+x.horas,0);
    const cls=total>0?'extra-pos':total<0?'extra-neg':'';
    return `<div class="baja-card"><div style="font-size:.86rem;font-weight:600;color:var(--text);margin-bottom:6px">${n}</div><div style="font-size:1.1rem;font-weight:700;font-family:'DM Mono',monospace" class="${cls}">${total>=0?'+':''}${total}h<span style="font-size:.72rem;font-weight:400;color:var(--muted)"> ajuste</span></div>${items?'<div style="margin-top:8px">'+items+'</div>':'<div style="font-size:.68rem;color:var(--muted);margin-top:6px">Sin ajustes</div>'}</div>`;
  }).join('');
}

// TABS CONTROL LABORAL
function switchCtrlTab(tab){
  const tabs=['vac','baja','libre','extra'];
  document.querySelectorAll('.ctrl-tab').forEach((t,i)=>t.classList.toggle('active',tabs[i]===tab));
  document.querySelectorAll('.ctrl-section').forEach(s=>s.classList.remove('active'));
  document.getElementById('ctrl-'+tab).classList.add('active');
}

// CAL
function renderCal(){
  const y=fst.calYear,m=fst.calMonth;
  document.getElementById('cal-title').textContent=MESES[m]+' '+y;
  const offset=(new Date(y,m,1).getDay()+6)%7;const dim=new Date(y,m+1,0).getDate();const today=dateStr(new Date());
  const vacDays=new Set();WORKERS.forEach(n=>fst.vacaciones[n].forEach(v=>{let cur=new Date(v.start+'T00:00:00');const end=new Date(v.end+'T00:00:00');while(cur<=end){vacDays.add(dateStr(cur));cur.setDate(cur.getDate()+1);}}));
  let html=['L','M','X','J','V','S','D'].map(d=>`<div class="cal-dh">${d}</div>`).join('');
  for(let i=0;i<offset;i++)html+=`<div class="cal-day"></div>`;
  for(let d=1;d<=dim;d++){const ds=y+'-'+pad(m+1)+'-'+pad(d);const dow=new Date(y,m,d).getDay();let cls='cal-day';if(ds===today)cls+=' today';else if(FESTIVOS.includes(ds))cls+=' festivo';else if(vacDays.has(ds))cls+=' vac-day';else if(dow===0||dow===6)cls+=' finde';html+=`<div class="${cls}">${d}</div>`;}
  document.getElementById('cal-grid').innerHTML=html;
  const eventos=[];FESTIVOS.filter(f=>f.startsWith(y+'-'+pad(m+1))).forEach(f=>eventos.push({date:f,name:'Festivo nacional',type:'fest'}));WORKERS.forEach(n=>fst.vacaciones[n].forEach(v=>{if(v.start.startsWith(y+'-'+pad(m+1))||v.end.startsWith(y+'-'+pad(m+1)))eventos.push({date:v.start,name:'Vacaciones – '+n,type:'vacd'});}));eventos.sort((a,b)=>a.date.localeCompare(b.date));
  const evEl=document.getElementById('cal-eventos');if(!eventos.length){evEl.innerHTML='<div class="empty">Sin eventos este mes</div>';return;}
  evEl.innerHTML=eventos.map(e=>`<div class="ev-row"><div style="color:var(--muted);min-width:76px;font-family:'DM Mono',monospace">${fmtDate(e.date)}</div><div style="flex:1">${e.name}</div><span class="ev-type ${e.type}">${e.type==='fest'?'Festivo':'Vacaciones'}</span></div>`).join('');
}
function calMove(dir){fst.calMonth+=dir;if(fst.calMonth>11){fst.calMonth=0;fst.calYear++;}if(fst.calMonth<0){fst.calMonth=11;fst.calYear--;}renderCal();}

// EDITAR FICHAJES
function fmtHM(ts){const d=new Date(ts);return pad(d.getHours())+':'+pad(d.getMinutes());}
function renderEditar(){
  const f=document.getElementById('filt-edit').value;const el=document.getElementById('edit-list');
  const byJid=new Map();
  fst.registros.forEach(r=>{const jid=Math.floor(r.id/2);if(!byJid.has(jid))byJid.set(jid,{});byJid.get(jid)[r.tipo]=r;});
  const jornadas=[...byJid.entries()]
    .filter(([,v])=>v.entrada&&(!f||v.entrada.nombre===f))
    .map(([jid,v])=>({jid,nombre:v.entrada.nombre,tsE:v.entrada.ts,tsS:v.salida?v.salida.ts:null,idE:v.entrada.id,idS:v.salida?v.salida.id:null}))
    .sort((a,b)=>b.tsE-a.tsE);
  if(!jornadas.length){el.innerHTML='<div class="tbl"><div class="empty">Sin fichajes</div></div>';return;}
  el.innerHTML=`<div class="tbl">
    <div class="tr th">
      <div class="tc" style="flex:1.3">Trabajador</div>
      <div class="tc" style="flex:.9">Fecha</div>
      <div class="tc" style="flex:.7">Entrada</div>
      <div class="tc" style="flex:.7">Salida</div>
      <div class="tc" style="flex:.4"></div>
    </div>
    ${jornadas.map(j=>`<div class="tr">
      <div class="tc" style="flex:1.3;font-size:.82rem">${j.nombre}</div>
      <div class="tc" style="flex:.9;color:var(--muted);font-size:.8rem">${fmtFecha(j.tsE)}</div>
      <div class="tc" style="flex:.7;color:var(--accent2);font-family:'DM Mono',monospace;font-size:.82rem">${fmtHM(j.tsE)}</div>
      <div class="tc" style="flex:.7;color:var(--danger);font-family:'DM Mono',monospace;font-size:.82rem">${j.tsS?fmtHM(j.tsS):'—'}</div>
      <div class="tc" style="flex:.4;text-align:right"><button class="btn-sm" onclick="openEditModal(${j.jid})">Editar</button></div>
    </div>`).join('')}
  </div>`;
}
function closeEditModal(){document.getElementById('edit-modal').classList.remove('open');fst.editingId=null;}
function deleteEditing(){
  const jid=fst.editingId;
  try{apiPost({tipo:'delFichaje',id:jid});}catch(e){}
  fst.registros=fst.registros.filter(x=>Math.floor(x.id/2)!==jid);
  WORKERS.forEach(n=>recalcWorker(n));saveFst();closeEditModal();renderEditar();renderMeses();renderHistorial();renderStats();WORKERS.forEach(n=>renderWcard(n));
}

// ── NUEVO FICHAJE MANUAL ──
function openNewFichajeModal(){
  fst.editingId='__new__';
  document.getElementById('em-title').textContent='Nuevo fichaje';
  document.getElementById('em-btn-delete').style.display='none';
  document.getElementById('edit-modal').classList.add('open');
  document.getElementById('em-worker').value=WORKERS[0];
  document.getElementById('em-fecha').value='';
  document.getElementById('em-hora-e').value='';
  document.getElementById('em-hora-s').value='';
}
function openEditModal(jid){
  const re=fst.registros.find(x=>Math.floor(x.id/2)===jid&&x.tipo==='entrada');
  const rs=fst.registros.find(x=>Math.floor(x.id/2)===jid&&x.tipo==='salida');
  if(!re)return;
  fst.editingId=jid;
  document.getElementById('em-title').textContent='Editar jornada';
  document.getElementById('em-btn-delete').style.display='';
  document.getElementById('edit-modal').classList.add('open');
  document.getElementById('em-worker').value=re.nombre;
  document.getElementById('em-fecha').value=dateStr(new Date(re.ts));
  document.getElementById('em-hora-e').value=fmtHM(re.ts);
  document.getElementById('em-hora-s').value=rs?fmtHM(rs.ts):'';
}
async function saveEdit(){
  const isNew=fst.editingId==='__new__';
  const nombre=document.getElementById('em-worker').value;
  const fechaVal=document.getElementById('em-fecha').value;
  const horaE=document.getElementById('em-hora-e').value;
  const horaS=document.getElementById('em-hora-s').value;
  if(!nombre||!fechaVal||!horaE){alert('Trabajador, fecha y hora de entrada son obligatorios.');return;}
  const [fy,fm,fd]=fechaVal.split('-').map(Number);
  const [heH,heM]=horaE.split(':').map(Number);
  const entradaStr=pad(heH)+':'+pad(heM);
  let salidaStr='';
  let tsE=new Date(fy,fm-1,fd,heH,heM,0).getTime();
  let tsS=null;
  let durMs=0;
  if(horaS){
    const[hsH,hsM]=horaS.split(':').map(Number);
    salidaStr=pad(hsH)+':'+pad(hsM);
    tsS=new Date(fy,fm-1,fd,hsH,hsM,0).getTime();
    if(tsS<tsE)tsS+=86400000;
    durMs=tsS-tsE;
  }
  const tiempodia=durMs>0?(durMs/3600000).toFixed(2):'';
  const fechaISO=fechaVal; // YYYY-MM-DD
  const fentradaISO=fechaISO+'T'+entradaStr+':00';
  const fsalidaISO=salidaStr?(fechaISO+'T'+salidaStr+':00'):null;

  if(isNew){
    // Insertar en Supabase
    try{
      const result=await apiPost({tipo:'fichajeManual',empleado:nombre.toUpperCase(),fecha:fechaISO,entrada:entradaStr,salida:salidaStr||null,fentrada:fentradaISO,fsalida:fsalidaISO,tiempodia:tiempodia||null});
      if(!result.ok){alert('Error al guardar: '+(result.error||'desconocido'));return;}
      // Añadir a registros locales
      const newId=result.id||Math.floor(Date.now()/1000);
      fst.registros.push({id:newId*2,nombre,tipo:'entrada',ts:tsE});
      if(tsS)fst.registros.push({id:newId*2+1,nombre,tipo:'salida',ts:tsS,duracion:durMs});
    }catch(e){alert('Error de conexión');return;}
  }else{
    // Editar existente (código original)
    const jid=fst.editingId;
    const re=fst.registros.find(x=>Math.floor(x.id/2)===jid&&x.tipo==='entrada');
    const rs=fst.registros.find(x=>Math.floor(x.id/2)===jid&&x.tipo==='salida');
    if(!re)return;
    re.nombre=nombre;re.ts=tsE;
    if(rs&&horaS){const[hsH,hsM]=horaS.split(':').map(Number);rs.nombre=nombre;rs.ts=new Date(fy,fm-1,fd,hsH,hsM,0).getTime();if(rs.ts<re.ts)rs.ts+=86400000;rs.duracion=rs.ts-re.ts;}
    else if(!rs&&horaS){fst.registros.push({id:jid*2+1,nombre,tipo:'salida',ts:tsS,duracion:durMs});}
    try{apiPost({tipo:'editFichaje',id:jid,empleado:nombre.toUpperCase(),fecha:fechaISO,entrada:entradaStr,salida:salidaStr||null,tiempodia:tiempodia||null,fentrada:fentradaISO,fsalida:fsalidaISO});}catch(e){}
  }
  WORKERS.forEach(n=>recalcWorker(n));saveFst();closeEditModal();renderEditar();renderMeses();renderHistorial();renderStats();WORKERS.forEach(n=>renderWcard(n));
}

// ── OT ────────────────────────────────────────────────────────
let selMachine=null,selGama=null,checkStates=[],currentFilter='TODOS';
let MACHINES=[
  {id:'M966G.01',name:'Pala Cargadora 966G',tipo:'PALA CARGADORA',modelo:'966G',fabricante:'CATERPILLAR'},
  {id:'M349.1',name:'Excavadora 349',tipo:'EXCAVADORA',modelo:'349',fabricante:'CATERPILLAR'},
  {id:'M725.1',name:'Dumper Articulado 725C',tipo:'DUMPER ARTICULADO',modelo:'725C',fabricante:'CATERPILLAR'},
  {id:'M40.9.1',name:'Manipulador Telescópico P40.9',tipo:'MANIPULADOR TELESCOPICO',modelo:'P40.9',fabricante:'MERLO'},
  {id:'M330.1',name:'Excavadora 330',tipo:'EXCAVADORA',modelo:'330',fabricante:'CASE POCLAIN'},
  {id:'M336.1',name:'Excavadora 336',tipo:'EXCAVADORA',modelo:'336',fabricante:'CATERPILLAR'},
  {id:'M962.1',name:'Pala Cargadora 962',tipo:'PALA CARGADORA',modelo:'962',fabricante:'CATERPILLAR'},
  {id:'MC32.1',name:'Grupo Electrógeno C32',tipo:'GRUPO ELECTROGENO',modelo:'C32',fabricante:'CATERPILLAR'},
  {id:'1274MMF',name:'Tractocamión Renault',tipo:'TRACTOCAMION',modelo:'-',fabricante:'RENAULT'},
  {id:'6352MJV',name:'Rígido Volvo',tipo:'RIGIDO',modelo:'-',fabricante:'VOLVO'},
  {id:'DE22',name:'Grupo Generador E22',tipo:'GRUPO GENERADOR',modelo:'E22',fabricante:'CATERPILLAR'},
  {id:'GC6825BV',name:'Cuba de Agua',tipo:'CUBA DE AGUA',modelo:'-',fabricante:'MERCEDES'},
  {id:'M1',name:'Machacadora RT12010',tipo:'MACHACADORA',modelo:'RT1210',fabricante:'MOPSA'},
  {id:'AP1',name:'Alimentador AP1',tipo:'ALIMENTADOR',modelo:'CINTAS',fabricante:'MOPSA'},
  {id:'AM3',name:'Alimentador Banda AM3',tipo:'ALIMENTADOR',modelo:'CINTAS',fabricante:'MOPSA'},
  {id:'AT1',name:'Alimentador Bandeja AT1',tipo:'ALIMENTADOR',modelo:'AT1',fabricante:'MOPSA'},
  {id:'AT2',name:'Alimentador Bandeja AT2',tipo:'ALIMENTADOR',modelo:'AT2',fabricante:'MOPSA'},
  {id:'T1',name:'Cinta T1',tipo:'CINTA TRANSPORTADORA',modelo:'CINTAS',fabricante:'MOPSA'},{id:'T2',name:'Cinta T2',tipo:'CINTA TRANSPORTADORA',modelo:'CINTAS',fabricante:'MOPSA'},{id:'T3',name:'Cinta T3',tipo:'CINTA TRANSPORTADORA',modelo:'CINTAS',fabricante:'MOPSA'},{id:'T4',name:'Cinta T4',tipo:'CINTA TRANSPORTADORA',modelo:'CINTAS',fabricante:'MOPSA'},{id:'T5',name:'Cinta T5',tipo:'CINTA TRANSPORTADORA',modelo:'CINTAS',fabricante:'MOPSA'},{id:'T6',name:'Cinta T6',tipo:'CINTA TRANSPORTADORA',modelo:'CINTAS',fabricante:'MOPSA'},{id:'T7',name:'Cinta T7',tipo:'CINTA TRANSPORTADORA',modelo:'CINTAS',fabricante:'MOPSA'},
  {id:'TE1',name:'Cinta TE1',tipo:'CINTA TRANSPORTADORA',modelo:'CINTAS',fabricante:'MOPSA'},{id:'TE2',name:'Cinta TE2',tipo:'CINTA TRANSPORTADORA',modelo:'CINTAS',fabricante:'MOPSA'},{id:'TE3',name:'Cinta TE3',tipo:'CINTA TRANSPORTADORA',modelo:'CINTAS',fabricante:'MOPSA'},{id:'T04',name:'Cinta T04',tipo:'CINTA TRANSPORTADORA',modelo:'CINTAS',fabricante:'MOPSA'},{id:'T412',name:'Cinta T412',tipo:'CINTA TRANSPORTADORA',modelo:'CINTAS',fabricante:'MOPSA'},{id:'T1220',name:'Cinta T1220',tipo:'CINTA TRANSPORTADORA',modelo:'CINTAS',fabricante:'MOPSA'},{id:'TS01',name:'Cinta TS01',tipo:'CINTA TRANSPORTADORA',modelo:'CINTAS',fabricante:'MOPSA'},{id:'TS02',name:'Cinta TS02',tipo:'CINTA TRANSPORTADORA',modelo:'CINTAS',fabricante:'MOPSA'},
  {id:'C1',name:'Criba Vibrante C1',tipo:'CRIBA VIBRANTE',modelo:'VKW7203',fabricante:'GRANIER'},{id:'C2',name:'Criba Vibrante C2',tipo:'CRIBA VIBRANTE',modelo:'VKW7203',fabricante:'GRANIER'},{id:'CE',name:'Criba Vibrante CE',tipo:'CRIBA VIBRANTE',modelo:'VKW7203',fabricante:'GRANIER'},
  {id:'M2',name:'Molino de Cono HP4',tipo:'MOLINO DE CONO',modelo:'HP4',fabricante:'NORDBERG'},{id:'M3',name:'Molino Arenero OM120',tipo:'MOLINO ARENERO',modelo:'OM120',fabricante:'ORE SIZER'},
  {id:'M769C.01',name:'Dumper Rígido 769C',tipo:'DUMPER RIGIDO',modelo:'769C',fabricante:'CATERPILLAR'},{id:'M365B.01',name:'Excavadora 365B',tipo:'EXCAVADORA',modelo:'365B',fabricante:'CATERPILLAR'},
  {id:'8590FBV',name:'Renault Kangoo',tipo:'VEHICULO',modelo:'-',fabricante:'RENAULT'},{id:'3833BNX',name:'Hyundai H1',tipo:'VEHICULO',modelo:'-',fabricante:'HYUNDAI'},{id:'8212FLC',name:'Tractocamión Martel',tipo:'TRACTOCAMION',modelo:'-',fabricante:'-'},{id:'VARIOSITC',name:'Tractocamión ITC',tipo:'TRACTOCAMION',modelo:'-',fabricante:'VARIOS'},{id:'P-01',name:'Planta Mesa Cañadas',tipo:'PLANTA',modelo:'-',fabricante:'ARIFOMA'},
];
const GAMAS=[
  {id:'M500H',modelo:'P40.9',nombre:'500 Horas',intervalo:500,checks:['LIMPIAR FILTRO RESPIRADERO TANQUE ACEITE HIDRAULICO','NIVEL ACEITE REDUCTOR RUEDAS','NIVEL ACEITE DIFERENCIALES','NIVEL ACEITE DEL CAMBIO','CONTROL CIERRE BULONES','ENGRASE JUNTAS CARDANICAS','VACIAR AGUA TANQUE GASOIL','SUSTITUIR FILTRO TRANSMISION','NIVEL ACEITE BRAZOS PORTARUEDA','ENGRASE TUBO GRUIA LATIGUILLOS','SUSTITUIR FILTRO RETORNO HIDRAULICO','SUSTITUIR FILTRO DEL AIRE','Cambio de aceite y filtro motor','Cambio filtro combustible']},
  {id:'M1000H',modelo:'P40.9',nombre:'1000 Horas',intervalo:1000,checks:['LIMPIAR FILTRO RESPIRADERO TANQUE ACEITE HIDRAULICO','NIVEL ACEITE REDUCTOR RUEDAS','NIVEL ACEITE DIFERENCIALES','NIVEL ACEITE DEL CAMBIO','CONTROL CIERRE BULONES','ENGRASE JUNTAS CARDANICAS','VACIAR AGUA TANQUE GASOIL','SUSTITUIR FILTRO TRANSMISION','NIVEL ACEITE BRAZOS PORTARUEDA','ENGRASE TUBO GRUIA LATIGUILLOS','SUSTITUIR FILTRO RETORNO HIDRAULICO','SUSTITUIR FILTRO DEL AIRE','VERIFICAR JUEGOS Y ENGRASAR ARTICULACIONES','ENGRASAR PATINES DEL BRAZO','SUSTITUIR FILTRO ANTIPOLVO']},
  {id:'M2000H',modelo:'P40.9',nombre:'2000 Horas',intervalo:2000,checks:['LIMPIAR FILTRO RESPIRADERO TANQUE ACEITE HIDRAULICO','SUSTITUIR ACEITE REDUCTOR RUEDAS','SUSTITUIR ACEITE DIFERENCIALES','SUSTITUIR ACEITE DEL CAMBIO','CONTROL CIERRE BULONES','ENGRASE JUNTAS CARDANICAS','VACIAR AGUA TANQUE GASOIL','SUSTITUIR FILTRO TRANSMISION','NIVEL ACEITE BRAZOS PORTARUEDA','ENGRASE TUBO GRUIA LATIGUILLOS','SUSTITUIR FILTRO RETORNO HIDRAULICO','SUSTITUIR FILTRO DEL AIRE','VERIFICAR JUEGOS Y ENGRASAR ARTICULACIONES','ENGRASAR PATINES DEL BRAZO','SUSTITUIR FILTRO ANTIPOLVO','SUSTITUIR ACEITE HIDRAULICO','SUSTITUIR ACEITE TRANSMISION HIDROSTATICA','SUSTITUIR ACEITE DE LOS FRENOS']},
  {id:'966G-10H',modelo:'966G',nombre:'Diario (10h)',intervalo:10,checks:['Probar Alarma de retroceso','Inspeccionar cuchillas del cucharón','Comprobar nivel de refrigerante','Comprobar nivel de aceite del motor','Comprobar nivel de aceite hidraúlico','Comprobar nivel de aceite de la transmisión','Limpiar Ventanas']},
  {id:'966G-50H',modelo:'966G',nombre:'Semanal (50h)',intervalo:50,checks:['Probar Alarma de retroceso','Inspeccionar cuchillas','Comprobar nivel de refrigerante','Comprobar nivel de aceite del motor','Comprobar nivel de aceite hidraúlico','Lubricar cojinetes de pivote','Limpiar filtro de aire de la cabina','Drenar agua del tanque de combustible','Comprobar inflado de neumáticos']},
  {id:'966G-250H',modelo:'966G',nombre:'250 Horas',intervalo:250,checks:['Comprobar aire acondicionado','Limpiar batería','Inspeccionar correas','Comprobar acumulador de freno','Probar sistema de frenos','Lubricar Estrías del eje motriz','Limpiar el respiradero del cárter','Cambiar aceite y filtro del motor']},
  {id:'966G-500H',modelo:'966G',nombre:'500 Horas',intervalo:500,checks:['Comprobar aire acondicionado','Limpiar batería','Inspeccionar correas','Cambiar aceite y filtro del motor','Reemplazar Filtro primario de combustible','Reemplazar Filtro secundario de combustible','Reemplazar Filtro de aceite hidraúlico','Reemplazar Filtro de aceite de la transmisión']},
  {id:'966G-1000H',modelo:'966G',nombre:'1000 Horas',intervalo:1000,checks:['Comprobar aire acondicionado','Limpiar batería','Cambiar aceite y filtro del motor','Reemplazar filtros de combustible','Reemplazar Filtro de aceite hidraúlico','Cambio de aceite de la transmisión']},
  {id:'966G-2000H',modelo:'966G',nombre:'2000 Horas',intervalo:2000,checks:['Comprobar aire acondicionado','Cambiar aceite y filtro del motor','Reemplazar filtros de combustible','Reemplazar Filtro de aceite hidraúlico','Cambiar aceite de la transmisión','Cambiar aceite del diferencial','Comprobar Juego de válvulas del motor','Cambiar Aceite del sistema hidraúlico']},
  {id:'769C-10H',modelo:'769C',nombre:'Diario (10h)',intervalo:10,checks:['Nivel de refrigerante.','Nivel de aceite motor.','Drenar depósitos de aire.','Drenar agua del depósito de combustible.','Comprobar frenos.','Comprobar luces.','Comprobar dirección auxiliar.','Comprobar alarma marcha atrás.','Control daños en neumáticos.']},
  {id:'769C-250H',modelo:'769C',nombre:'250 Horas',intervalo:250,checks:['Cambio aceite y filtro del motor (15W40 - 46L).','Nivel de aceite cojinetes ruedas delanteras.','Engrase estrías del eje impulsor.','Nivel desgaste cintas de freno.','Revisar correa alternador.','Cambio filtros hidráulicos y respiradero.']},
  {id:'769C-1000H',modelo:'769C',nombre:'1000 Horas',intervalo:1000,checks:['Cambio aceite y filtro del motor.','Cambio filtros hidráulicos.','Cambio filtros gasoil.','Cambio de aceite hidráulico transmisión (ISO46).','Cambio aceite hidráulico de dirección (ISO46).','Cambio aceite mandos finales y diferenciales.']},
  {id:'C32-250H',modelo:'C32',nombre:'250 Horas',intervalo:250,checks:['Comprobar nivel del electrólito de la batería','Inspeccionar correas','Limpiar radiador','Cambiar aceite y filtro del motor','Reemplazar filtro secundario de combustible','Inspeccionar mangueras y abrazaderas']},
  {id:'C32-2000H',modelo:'C32',nombre:'2000 Horas',intervalo:2000,checks:['Cambiar aceite y filtro del motor','Reemplazar filtros de combustible','Comprobar Juego de las válvulas del motor','Inspeccionar Soportes del motor','Prueba Aislamiento del devanado del generador','Cambiar Refrigerante (DEAC)']},
  {id:'MACH-80H',modelo:'RT1210',nombre:'80 Horas',intervalo:80,checks:['Engrase rodamientos (80 Gramos c/u) NLGI2.','Revisión de los asientos de las mandíbulas.','Revisión aprietes de los tacos de goma.','Revisión del reglaje entre pico y valle.','Engrase del cardán del alimentador.','Verificar tensión de las correas.','Revisión resortes laterales y principales.','Revisión del nivel de aceite (ISO200).','Revisión tope de seguridad.']},
  {id:'HP4-DIARIO',modelo:'HP4',nombre:'Diario',intervalo:1,checks:['Comprobar el nivel de aceite del depósito.','Comprobar la temperatura de entrada y retorno de aceite.','Comprobar la presión del contra-árbol.','Comprobar el reglaje del lado cerrado.','Comprobar la presión de sujeción (140 bares).','Comprobar ruidos anormales o señales de desgaste.']},
  {id:'HP4-200H',modelo:'HP4',nombre:'200 Horas',intervalo:200,checks:['Comprobar la precarga de los acumuladores (200 Bares).','Desbloquear el tazón y hacerlo girar en ambos sentidos.','Comprobar el aceite de lubricación.','Comprobar la holgura axial del contra-árbol (1,3 a 1,8mm).','Limpiar o cambiar el filtro de aceite.']},
  {id:'OM120-DIARIO',modelo:'OM120',nombre:'Diario (8h)',intervalo:1,checks:['ENGRASE LOS 4 PUNTOS DEL ROTOR CON 15 GRAMOS CADA 8 HORAS DE TRABAJO NLGI2.','INSPECCIÓN VISUAL DE LA MÁQUINA Y DEL ROTOR.']},
  {id:'OM120-80H',modelo:'OM120',nombre:'80 Horas',intervalo:80,checks:['ENGRASE DEL MOLINO 15GR C/U NLGI2.','ESTADO DE LAS CORREAS.','REVISIÓN DEL ROTOR (DESGASTE TIPS 20MM DESDE EL BORDE).']},
  {id:'CINTAS',modelo:'CINTAS',nombre:'80 Horas',intervalo:80,checks:['Engrase de rodamientos.','Verificar estado de la banda.','Verificar estado de los baberos.','Verificar estado de los guías de carga.','Revisar rodillos.']},
  {id:'GRANIERVKW7203-80H',modelo:'VKW7203',nombre:'80 Horas',intervalo:80,checks:['Engrase mecanismo de rotación 20Gr c/u (4 engrasadores).','Revisar tensión de las mallas.','Revisar tensión de las correas.','Inspeccionar estructura de la criba.','Revisar estado de las mallas.']},
];
const MODEL_TO_GAMAS={};GAMAS.forEach(g=>{if(!MODEL_TO_GAMAS[g.modelo])MODEL_TO_GAMAS[g.modelo]=[];MODEL_TO_GAMAS[g.modelo].push(g);});

function _buildOTFromActivos(data){
  if(!data||!data.length) return;
  MACHINES = data.filter(m=>m.Codigo).map(m=>({
    id: m.Codigo,
    name: m.Activo || m.Codigo,
    tipo: (m.tipoactivo||'OTROS').toUpperCase(),
    modelo: m.modelo||'-',
    fabricante: (m.fabricante||'-').toUpperCase()
  }));
}

async function initOT(){
  // Cargar activos desde Supabase si no están cargados aún
  if(!activosData.length){
    const result = await dbQuery({ action:'select', table:'tblactivos', options:{ select:'*', order:'Codigo.asc' } });
    if(result.ok && result.data && result.data.length){
      activosData = result.data;
    }
  }
  if(activosData.length) _buildOTFromActivos(activosData);

  const ft=document.getElementById('filterTabs');ft.innerHTML='';
  const labelsMap={'TODOS':'Todos','PALA CARGADORA':'Palas','EXCAVADORA':'Excavadoras','DUMPER ARTICULADO':'Dumpers','DUMPER RIGIDO':'Dumpers','MANIPULADOR TELESCOPICO':'Manipuladores','GRUPO ELECTROGENO':'Electrógenos','GRUPO GENERADOR':'Generadores','TRACTOCAMION':'Tractocamiones','MACHACADORA':'Machacadoras','MOLINO DE CONO':'Molinos','MOLINO ARENERO':'Molinos','CINTA TRANSPORTADORA':'Cintas','CRIBA VIBRANTE':'Cribas','CUBA DE AGUA':'Cubas','ALIMENTADOR':'Alimentadores','RIGIDO':'Rígidos','PLANTA':'Plantas','VEHICULO':'Vehículos'};
  const shown=new Set();
  ['TODOS',...new Set(MACHINES.map(m=>m.tipo))].forEach(t=>{const label=labelsMap[t]||t;if(shown.has(label))return;shown.add(label);const div=document.createElement('div');div.className='filter-tab'+(t==='TODOS'?' active':'');div.textContent=label;div.dataset.tipo=t;div.onclick=()=>{currentFilter=t;document.querySelectorAll('.filter-tab').forEach(x=>x.classList.remove('active'));div.classList.add('active');filterMachines();};ft.appendChild(div);});
  renderMachinesOT(MACHINES);
  const now=new Date();now.setMinutes(now.getMinutes()-now.getTimezoneOffset());
  document.getElementById('inputFecha').value=now.toISOString().slice(0,16);
  const selOp=document.getElementById('inputOperario');if(selOp){selOp.innerHTML='<option value="">— Seleccionar —</option>';WORKERS.forEach(n=>{const o=document.createElement('option');o.value=n;o.textContent=n;selOp.appendChild(o);});}
}
function renderMachinesOT(list){
  const grid=document.getElementById('machineGrid');
  if(!list.length){grid.innerHTML='<div class="no-results">No se encontraron máquinas</div>';return;}
  grid.innerHTML='';
  list.forEach(m=>{const d=document.createElement('div');d.className='machine-btn'+(selMachine&&selMachine.id===m.id?' selected':'');d.innerHTML=`<div class="machine-name">${m.name}</div><div class="machine-code">${m.id}</div><div class="machine-type">${m.fabricante}</div>`;d.onclick=()=>{selMachine=m;document.querySelectorAll('.machine-btn').forEach(x=>x.classList.remove('selected'));d.classList.add('selected');document.getElementById('btnS1').disabled=false;};grid.appendChild(d);});
}
function filterMachines(){
  const q=document.getElementById('searchInput').value.toLowerCase();let list=MACHINES;
  if(currentFilter!=='TODOS'){const labelsMap2={Palas:['PALA CARGADORA'],Excavadoras:['EXCAVADORA'],Dumpers:['DUMPER ARTICULADO','DUMPER RIGIDO'],Manipuladores:['MANIPULADOR TELESCOPICO'],Electrógenos:['GRUPO ELECTROGENO'],Generadores:['GRUPO GENERADOR'],Tractocamiones:['TRACTOCAMION'],Machacadoras:['MACHACADORA'],Molinos:['MOLINO DE CONO','MOLINO ARENERO'],Cintas:['CINTA TRANSPORTADORA'],Cribas:['CRIBA VIBRANTE'],Cubas:['CUBA DE AGUA'],Alimentadores:['ALIMENTADOR'],Rígidos:['RIGIDO'],Plantas:['PLANTA']};const label=document.querySelector('.filter-tab.active')?.textContent;const tipos=labelsMap2[label]||[currentFilter];list=list.filter(m=>tipos.includes(m.tipo));}
  if(q)list=list.filter(m=>m.name.toLowerCase().includes(q)||m.id.toLowerCase().includes(q)||m.fabricante.toLowerCase().includes(q));
  renderMachinesOT(list);
}
function showOT(id){['screen1','screen2','screen3','screen4','screen5','screenOK'].forEach(s=>{const el=document.getElementById(s);el.classList.add('hidden');el.style.display='';});const el=document.getElementById(id);el.classList.remove('hidden');if(id==='screenOK')el.style.display='block';}
function setSteps(active){for(let n=1;n<=5;n++){const el=document.getElementById('s'+n);el.style.opacity='1';el.classList.remove('active','done');if(n<active)el.classList.add('done');else if(n===active)el.classList.add('active');else el.style.opacity='0.4';}for(let n=1;n<=4;n++)document.getElementById('l'+n).classList.toggle('done',n<active);}
function goStep1(){showOT('screen1');setSteps(1);}
async function goStep2(){if(!selMachine)return;const grid=document.getElementById('gamaGrid');grid.innerHTML='<div class="no-results">Cargando gamas...</div>';showOT('screen2');setSteps(2);
  // Gamas: Supabase tiene prioridad, hardcoded como fallback
  let gamas=[];
  try{
    // Forzar recarga desde Supabase siempre
    const json=await apiFetch('?accion=gamasNormas').catch(()=>({ok:false}));
    if(json.ok)normasData=json.data||[];
    const mid=(selMachine.id||'').trim().toUpperCase();
    const mmod=(selMachine.modelo||'').trim().toUpperCase();
    const dynGamas=normasData.filter(n=>{
      const nm=(n.Modelo||'').trim().toUpperCase();
      // Coincide por ID de activo o por modelo
      return nm===mid||nm===mmod;
    });
    dynGamas.forEach(n=>{
      const checks=[];for(let i=1;i<=60;i++)if(n['n'+i])checks.push(n['n'+i]);
      gamas.push({id:n.Numero||'DB-'+n.id,modelo:n.Modelo,nombre:n.Gama||n.Numero||'Gama '+n.id,intervalo:n.Intervalo||0,checks,_src:'db',_dbId:n.id});
    });
  }catch(e){console.warn('Error cargando gamas dinámicas:',e);}
  // Hardcoded solo como fallback si no hay nada en Supabase para este modelo
  if(!gamas.length){
    gamas=(MODEL_TO_GAMAS[selMachine.modelo]||[]).map(g=>({...g,_src:'local'}));
  }
  if(!gamas.length){grid.innerHTML='<div class="no-results">No hay gamas definidas para '+selMachine.modelo+'.</div>';}else{grid.innerHTML='';gamas.forEach(g=>{const d=document.createElement('div');d.className='gama-btn'+(selGama&&selGama.id===g.id?' selected':'');d.innerHTML=`<div><div class="gama-name">${g.nombre}</div><div class="gama-meta">${g.checks.length} puntos</div></div><div class="gama-interval">${g.intervalo===1?'Diario':g.intervalo+'h'}</div>`;d.onclick=()=>{selGama=g;document.querySelectorAll('.gama-btn').forEach(x=>x.classList.remove('selected'));d.classList.add('selected');document.getElementById('btnS2').disabled=false;};grid.appendChild(d);});}}
function goStep3(){showOT('screen3');setSteps(3);}
function checkStep3(){document.getElementById('btnS3').disabled=false;}
function goStep4(){if(!document.getElementById('inputHoras').value){mostrarConfirm('¿Continuar sin horómetro?',goStep4Real);}else goStep4Real();}
function mostrarConfirm(msg,cb){const o=document.createElement('div');o.style.cssText='position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,.7);z-index:9999;display:flex;align-items:center;justify-content:center;padding:20px;';o.innerHTML=`<div style="background:var(--surface);border:1px solid var(--border);border-radius:14px;padding:22px;max-width:300px;width:100%;text-align:center;"><div style="font-size:.9rem;color:var(--text);margin-bottom:18px;line-height:1.5;">${msg}</div><div style="display:flex;gap:9px;"><button id="cNo" style="flex:1;padding:10px;background:var(--surface2);border:1.5px solid var(--border);border-radius:8px;color:var(--text);font-family:DM Sans,sans-serif;font-size:.88rem;cursor:pointer;">Cancelar</button><button id="cSi" style="flex:1;padding:10px;background:var(--accent);border:none;border-radius:8px;color:#fff;font-family:DM Sans,sans-serif;font-size:.88rem;font-weight:700;cursor:pointer;">Continuar</button></div></div>`;document.body.appendChild(o);document.getElementById('cSi').onclick=()=>{document.body.removeChild(o);cb();};document.getElementById('cNo').onclick=()=>{document.body.removeChild(o);};}
function goStep4Real(){checkStates=new Array(selGama.checks.length).fill(false);document.getElementById('checkTitle').textContent=selMachine.name;document.getElementById('checkBanner').textContent=`⏱ Gama: ${selGama.nombre}  ·  ${selGama.checks.length} puntos`;const list=document.getElementById('checkItems');list.innerHTML='';selGama.checks.forEach((c,i)=>{const item=document.createElement('div');item.className='check-item';item.innerHTML=`<div class="check-box"></div><div class="check-text">${c}</div>`;item.onclick=()=>{checkStates[i]=!checkStates[i];item.classList.toggle('checked',checkStates[i]);updateProg();};list.appendChild(item);});updateProg();showOT('screen4');setSteps(4);}
function toggleAllChecks(){const allDone=checkStates.every(Boolean);checkStates.fill(!allDone);document.querySelectorAll('.check-item').forEach((item,i)=>{item.classList.toggle('checked',checkStates[i]);});updateProg();}
function updateProg(){const done=checkStates.filter(Boolean).length,total=checkStates.length;document.getElementById('progBar').style.width=(total?done/total*100:0)+'%';document.getElementById('progLabel').textContent=`${done} / ${total} completados`;document.getElementById('btnS4').disabled=false;}
function goStep5(){const tipos={preventivo:'Preventivo programado',correctivo:'Correctivo',revision:'Revisión rápida'};const tipo=document.getElementById('inputTipo').value;const horas=document.getElementById('inputHoras').value;const obs=document.getElementById('inputObs').value;const operario=document.getElementById('inputOperario').value||'—';const estadoOT=document.getElementById('inputEstado').value==='cerrada'?'Cerrada':'Abierta';const fechaRaw=document.getElementById('inputFecha').value;const fecha=fechaRaw?new Date(fechaRaw).toLocaleString('es-ES',{day:'2-digit',month:'2-digit',year:'numeric',hour:'2-digit',minute:'2-digit'}):'—';document.getElementById('summaryMain').innerHTML=`<div class="sum-box"><div class="sum-row"><span class="sum-key">Máquina</span><span class="sum-val">${selMachine.name}</span></div><div class="sum-row"><span class="sum-key">Código</span><span class="sum-val">${selMachine.id}</span></div><div class="sum-row"><span class="sum-key">Fecha</span><span class="sum-val">${fecha}</span></div><div class="sum-row"><span class="sum-key">Operario</span><span class="sum-val">${operario}</span></div><div class="sum-row"><span class="sum-key">Estado</span><span class="sum-val">${estadoOT}</span></div><div class="sum-row"><span class="sum-key">Horómetro</span><span class="sum-val">${horas} h</span></div><div class="sum-row"><span class="sum-key">Gama</span><span class="sum-val">${selGama.nombre}</span></div><div class="sum-row"><span class="sum-key">Tipo</span><span class="sum-val">${tipos[tipo]}</span></div>${obs?`<div class="sum-row"><span class="sum-key">Obs.</span><span class="sum-val">${obs}</span></div>`:''}</div>`;document.getElementById('summaryChecks').innerHTML=`<div class="sum-box">${selGama.checks.map((c,i)=>`<div class="sum-row"><span class="sum-key" style="max-width:70%;font-size:.62rem">${c}</span><span class="sum-val" style="color:${checkStates[i]?'var(--accent2)':'var(--muted)'}">${checkStates[i]?'✓ OK':'— Pendiente'}</span></div>`).join('')}</div>`;showOT('screen5');setSteps(5);}
async function submitOT(){
  const fechaRaw=document.getElementById('inputFecha').value;const fecha=fechaRaw?new Date(fechaRaw).toISOString():'';const horas=document.getElementById('inputHoras').value;const obs=document.getElementById('inputObs').value;const todosOk=checkStates.every(Boolean);const checks60=Array(60).fill(false);checkStates.forEach((v,i)=>{checks60[i]=v;});
  const operario=document.getElementById('inputOperario').value||null;const estadoOT=document.getElementById('inputEstado').value||'abierta';
  const payload={activo:selMachine.id,fecha:fecha||null,operario:operario,tiempo:null,texto:obs||null,estado:estadoOT==='cerrada',gama:selGama.id,medicion:horas?Number(horas):null,checks:checks60};
  document.getElementById('stepsBar').style.display='none';showOT('screenOK');document.getElementById('okRef').textContent='Enviando...';
  document.getElementById('okSummary').innerHTML=`<div class="sum-row"><span class="sum-key">Máquina</span><span class="sum-val">${selMachine.name}</span></div><div class="sum-row"><span class="sum-key">Gama</span><span class="sum-val">${selGama.nombre}</span></div><div class="sum-row"><span class="sum-key">Horómetro</span><span class="sum-val">${horas} h</span></div>`;
  try{const json=await apiPost(payload);if(json.ok)document.getElementById('okRef').textContent='OT-'+String(json.ot).padStart(4,'0');else{console.error('OT save error:',json.error);document.getElementById('okRef').textContent='Error: '+(json.error||'desconocido');}}catch(e){console.error('OT network error:',e);document.getElementById('okRef').textContent='Sin conexión';}
}
function resetAll(){selMachine=null;selGama=null;checkStates=[];document.getElementById('inputHoras').value='';document.getElementById('inputObs').value='';document.getElementById('inputOperario').value='';document.getElementById('inputEstado').value='abierta';document.getElementById('btnS1').disabled=true;document.getElementById('btnS2').disabled=true;document.getElementById('stepsBar').style.display='flex';document.getElementById('searchInput').value='';currentFilter='TODOS';document.querySelectorAll('.filter-tab').forEach((t,i)=>t.classList.toggle('active',i===0));initOT();goStep1();}

// ── HISTORIAL OT ──────────────────────────────────────────────
let otHistData=[];let otEditingId=null;let otHistSort='fecha'; // 'fecha' or 'ot'
// ── CONTROL DOCUMENTAL ───────────────────────────────────────
let docData=[];
function docEstado(dias,aviso){
  if(dias===null||dias===''||isNaN(dias))return 'gris';
  const d=Number(dias);
  if(d<=0)return 'rojo';
  if(d<=Number(aviso||30))return 'amarillo';
  return 'verde';
}
function docColor(e){return{rojo:'#ff4d4d',amarillo:'#f5a623',verde:'#4caf50',gris:'#666'}[e]||'#666';}
function fmtFechaDoc(v){
  if(!v)return '—';
  if(v instanceof Date)return v.toLocaleDateString('es-ES');
  if(typeof v==='string'&&v.includes('T'))return new Date(v).toLocaleDateString('es-ES');
  return v;
}
async function cargarDocumentos(){
  const el=document.getElementById('doc-list');
  el.innerHTML='<div class="tbl"><div class="empty">Cargando...</div></div>';
  try{
    const json=await apiFetch('?accion=documentos');
    if(!json.ok)throw new Error(json.error);
    docData=json.data;
    let r=0,a=0,v=0;
    docData.forEach(d=>{const e=docEstado(d.diasRestantes,d.tiempoAviso);if(e==='rojo')r++;else if(e==='amarillo')a++;else if(e==='verde')v++;});
    document.getElementById('doc-cnt-rojo').textContent=r;
    document.getElementById('doc-cnt-amarillo').textContent=a;
    document.getElementById('doc-cnt-verde').textContent=v;
    renderDocumentos();
  }catch(e){el.innerHTML='<div class="tbl"><div class="empty">Error: '+e.message+'</div></div>';}
}
function renderDocumentos(){
  const el=document.getElementById('doc-list');
  const fuente=document.getElementById('filt-doc-fuente').value;
  const estado=document.getElementById('filt-doc-estado').value;
  const txt=document.getElementById('filt-doc-texto').value.toLowerCase();
  let datos=docData.filter(d=>{
    if(fuente&&d.fuente!==fuente)return false;
    if(estado&&docEstado(d.diasRestantes,d.tiempoAviso)!==estado)return false;
    if(txt){
      const hay=(d.identificacion||d.nombre||'')+(d.activoNombre||d.tipoDocumento||'')+(d.descripcion||d.organo||'');
      if(!hay.toLowerCase().includes(txt))return false;
    }
    return true;
  });
  if(!datos.length){el.innerHTML='<div class="tbl"><div class="empty">Sin resultados</div></div>';return;}
  el.innerHTML=datos.map(d=>{
    const e=docEstado(d.diasRestantes,d.tiempoAviso);
    const col=docColor(e);
    const dias=d.diasRestantes!==null&&d.diasRestantes!==''?Number(d.diasRestantes):null;
    const diasTxt=dias===null?'Sin fecha':dias<=0?'Vencido hace '+(Math.abs(dias))+' días':dias+' días restantes';
    if(d.fuente==='principal'){
      return `<div style="border-left:3px solid ${col};padding:10px 12px;margin-bottom:6px;background:var(--card);border-radius:0 6px 6px 0">
        <div style="display:flex;justify-content:space-between;align-items:flex-start;gap:8px">
          <div>
            <div style="font-size:.72rem;color:var(--muted)">${d.tema||''} · ${d.departamento||''}</div>
            <div style="font-weight:700;font-size:.82rem;margin:2px 0">${d.identificacion||''}</div>
            <div style="font-size:.78rem;color:var(--text)">${d.activoNombre||''} — ${d.descripcion||''}</div>
          </div>
          <div style="text-align:right;flex-shrink:0">
            <div style="font-size:.7rem;color:${col};font-weight:700">${diasTxt}</div>
            <div style="font-size:.68rem;color:var(--muted)">Vig: ${fmtFechaDoc(d.fechaVig)}</div>
          </div>
        </div>
      </div>`;
    } else {
      return `<div style="border-left:3px solid ${col};padding:10px 12px;margin-bottom:6px;background:var(--card);border-radius:0 6px 6px 0">
        <div style="display:flex;justify-content:space-between;align-items:flex-start;gap:8px">
          <div>
            <div style="font-size:.72rem;color:var(--muted)">${d.tipoDocumento||''} · ${d.organo||''}</div>
            <div style="font-weight:700;font-size:.82rem;margin:2px 0">${d.nombre||''}</div>
            <div style="font-size:.78rem;color:var(--text)">#${d.numero||''} — ${d.estado||''}</div>
          </div>
          <div style="text-align:right;flex-shrink:0">
            <div style="font-size:.7rem;color:${col};font-weight:700">${diasTxt}</div>
            <div style="font-size:.68rem;color:var(--muted)">Vig: ${fmtFechaDoc(d.fechaVig)}</div>
          </div>
        </div>
      </div>`;
    }
  }).join('');
}

async function cargarHistorialOT(){
  const el=document.getElementById('ot-hist-list');el.innerHTML='<div class="tbl"><div class="empty">Cargando...</div></div>';
  try{const json=await apiFetch('?accion=historialOT');if(!json.ok)throw new Error(json.error);otHistData=json.data;
    const maquinas=[...new Set(json.data.map(r=>r.activo))].sort();const selM=document.getElementById('filt-ot-maquina');selM.innerHTML='<option value="">Todas las máquinas</option>';maquinas.forEach(m=>{const o=document.createElement('option');o.value=m;o.textContent=m;selM.appendChild(o);});
    const operarios=[...new Set(json.data.map(r=>r.operario))].sort();const selO=document.getElementById('filt-ot-operario');selO.innerHTML='<option value="">Todos los operarios</option>';operarios.forEach(op=>{const o=document.createElement('option');o.value=op;o.textContent=op;selO.appendChild(o);});
    filtrarHistorialOT();
  }catch(e){el.innerHTML='<div class="tbl"><div class="empty">Error: '+e.message+'</div></div>';}
}
function filtrarHistorialOT(){
  const fm=document.getElementById('filt-ot-maquina').value;const fo=document.getElementById('filt-ot-operario').value;
  const q=(document.getElementById('filt-ot-buscar')?document.getElementById('filt-ot-buscar').value:'').toUpperCase();
  let data=otHistData;
  if(fm)data=data.filter(r=>r.activo===fm);
  if(fo)data=data.filter(r=>r.operario===fo);
  if(q)data=data.filter(r=>{const txt=[r.activo,r.gama,r.operario,r.texto,String(r.ot||r.id)].join(' ').toUpperCase();return txt.includes(q);});
  // Sort by fecha descending (newest first) — handle dd/mm/yyyy and yyyy-mm-dd
  function _otFechaToNum(f){
    if(!f)return 0;
    const p=f.split(/[\/\-]/);
    if(p.length===3){
      // dd/mm/yyyy
      if(p[0].length===4)return parseInt(p[0]+p[1]+p[2]);
      return parseInt(p[2]+p[1]+p[0]);
    }
    return 0;
  }
  if(otHistSort==='ot') data=[...data].sort((a,b)=>(b.ot||b.id)-(a.ot||a.id));
  else data=[...data].sort((a,b)=>_otFechaToNum(b.fecha)-_otFechaToNum(a.fecha));
  const el=document.getElementById('ot-hist-list');
  if(!data.length){el.innerHTML='<div class="tbl"><div class="empty">Sin resultados</div></div>';return;}
  const sOT=otHistSort==='ot'?' ▼':'',sFecha=otHistSort==='fecha'?' ▼':'';
  el.innerHTML='<div class="tbl"><div class="tr th"><div class="tc" style="flex:.5;cursor:pointer" onclick="otHistSort=\'ot\';filtrarHistorialOT()">OT'+sOT+'</div><div class="tc" style="flex:1.2">Máquina</div><div class="tc" style="flex:1">Gama</div><div class="tc" style="flex:.8;cursor:pointer" onclick="otHistSort=\'fecha\';filtrarHistorialOT()">Fecha'+sFecha+'</div><div class="tc" style="flex:.8">Operario</div><div class="tc" style="flex:.4;text-align:center">Estado</div><div class="tc" style="flex:.6"></div></div>'+
  data.map(r=>`<div class="tr"><div class="tc" style="flex:.5;font-family:monospace;color:var(--muted)">#${String(r.ot||r.id).padStart(4,'0')}</div><div class="tc" style="flex:1.2;font-size:.78rem">${r.activo}</div><div class="tc" style="flex:1;color:var(--muted);font-size:.72rem">${r.gama}</div><div class="tc" style="flex:.8;color:var(--muted);font-size:.72rem">${r.fecha}</div><div class="tc" style="flex:.8;font-size:.75rem">${r.operario||'—'}</div><div class="tc" style="flex:.4;text-align:center"><span class="badge ${r.estado===true||r.estado==='TRUE'?'badge-ok':'badge-pend'}" style="cursor:pointer" onclick="toggleEstadoOT(${r.id})" title="Click para cambiar estado">${r.estado===true||r.estado==='TRUE'?'OK':'Pend'}</span></div><div class="tc" style="flex:.6;text-align:right;display:flex;gap:4px;justify-content:flex-end"><button class="btn-sm" onclick="printOTHistorial(${r.id})" title="Imprimir" style="font-size:.65rem;padding:2px 5px">🖨</button><button class="btn-sm" onclick="openOTEditModal(${r.id})">Editar</button><button class="btn-sm" onclick="eliminarOT(${r.id})" title="Eliminar" style="font-size:.65rem;padding:2px 5px;color:#e05;border-color:#e05">🗑</button></div></div>`).join('')+'</div>';
}
function openOTEditModal(id){const r=otHistData.find(x=>x.id==id);if(!r)return;otEditingId=id;document.getElementById('ot-edit-modal').classList.add('open');document.getElementById('oem-fecha').value=r.fecha||'';document.getElementById('oem-operario').value=r.operario||'';document.getElementById('oem-medicion').value=r.medicion||'';document.getElementById('oem-texto').value=r.texto||'';}
function closeOTEditModal(){document.getElementById('ot-edit-modal').classList.remove('open');otEditingId=null;}
async function saveOTEdit(){
  const payload={tipo:'editarOT',id:otEditingId,fecha:document.getElementById('oem-fecha').value,operario:document.getElementById('oem-operario').value,medicion:document.getElementById('oem-medicion').value,texto:document.getElementById('oem-texto').value};
  try{const json=await apiPost(payload);if(json.ok){closeOTEditModal();cargarHistorialOT();}else alert('Error: '+json.error);}catch(e){alert('Error de conexión');}
}
async function toggleEstadoOT(id){
  const r=otHistData.find(x=>x.id==id);if(!r)return;
  const nuevoEstado=!(r.estado===true||r.estado==='TRUE');
  try{const json=await apiPost({tipo:'editarOT',id,estado:nuevoEstado});if(!json.ok){alert('Error: '+json.error);return;}
    r.estado=nuevoEstado;filtrarHistorialOT();
  }catch(e){alert('Error de conexión');}
}
async function eliminarOT(id){
  const r=otHistData.find(x=>x.id==id);if(!r)return;
  if(!confirm('¿Eliminar OT #'+String(r.ot||r.id).padStart(4,'0')+'?\n'+r.activo+' — '+r.gama))return;
  try{const json=await apiPost({tipo:'deleteOT',id:Number(id)});if(!json.ok){alert('Error: '+json.error);return;}
    otHistData=otHistData.filter(x=>x.id!=id);filtrarHistorialOT();
  }catch(e){alert('Error de conexión');}
}

// ── GASOIL ────────────────────────────────────────────────────
let gasoilData=[];
let gasoilConsumos=[];
let gasoilStock={dep1:0,dep2:0};

function gasoilTab(tab){
  ['registro','historial','consumos','horometros'].forEach(t=>{
    document.getElementById('gtab-'+t).style.display=t===tab?'block':'none';
    const btn=document.getElementById('gtab-btn-'+t);
    btn.style.background=t===tab?'var(--accent)':'transparent';
    btn.style.color=t===tab?'#fff':'var(--muted)';
  });
  if(tab==='horometros')renderGasoilHorometros();
}

function gasoilTipoChange(){
  const tipo=document.getElementById('gasoil-tipo-sel').value;
  const provRow=document.getElementById('gasoil-prov-row');
  const origenSel=document.getElementById('gasoil-origen-sel');
  if(tipo==='ENTRADA'){
    provRow.style.display='block';
    origenSel.value='PROVEEDOR';
  } else {
    provRow.style.display='none';
    if(origenSel.value==='PROVEEDOR')origenSel.value='DEPOSITO 1';
  }
}

async function cargarGasoil(){
  const el=document.getElementById('gasoil-hist-list');
  if(el)el.innerHTML='<div class="tbl"><div class="empty">Cargando...</div></div>';
  try{
    const json=await apiFetch('?accion=gasoil');
    if(json.ok){
      gasoilData=json.data||[];
      gasoilConsumos=json.consumos||[];
      gasoilStock={dep1:json.dep1||0,dep2:json.dep2||0};
      renderGasoilStock();
      renderGasoilHistorial();
      renderGasoilConsumos();
    } else throw new Error(json.error||'Error');
  }catch(e){
    if(el)el.innerHTML='<div class="tbl"><div class="empty">Error: '+e.message+'</div></div>';
    console.warn('Gasoil error:',e);
  }
}

function renderGasoilStock(){
  const dep1=gasoilStock.dep1, dep2=gasoilStock.dep2;
  const total=dep1+dep2;
  const d1El=document.getElementById('gasoil-dep1-display');
  const d2El=document.getElementById('gasoil-dep2-display');
  const totEl=document.getElementById('gasoil-total-display');
  if(d1El)d1El.textContent=dep1.toLocaleString()+' L';
  if(d2El)d2El.textContent=dep2.toLocaleString()+' L';
  if(totEl)totEl.textContent=total.toLocaleString()+' L';
}

function renderGasoilHistorial(){
  const filtTipo=(document.getElementById('gasoil-filt-tipo')||{}).value||'';
  const filtOrigen=(document.getElementById('gasoil-filt-origen')||{}).value||'';
  let filtered=[...gasoilData];
  if(filtTipo)filtered=filtered.filter(r=>String(r.tipo||'').toUpperCase().includes(filtTipo));
  if(filtOrigen)filtered=filtered.filter(r=>String(r.origen||'').toUpperCase().includes(filtOrigen.toUpperCase()));
  const el=document.getElementById('gasoil-hist-list');
  if(!el)return;
  if(!filtered.length){el.innerHTML='<div class="tbl"><div class="empty">Sin registros</div></div>';return;}
  const COLOR={'SALIDA':'rgba(255,95,95,.15)','ENTRADA':'rgba(107,125,46,.15)','ENTRADA-SALIDA':'rgba(107,125,46,.12)','REAJUSTE COMBUSTIBLE':'rgba(107,125,46,.15)'};
  const CTEXT={'SALIDA':'var(--danger)','ENTRADA':'var(--accent)','ENTRADA-SALIDA':'var(--accent2)','REAJUSTE COMBUSTIBLE':'#f5c842'};
  const rows=filtered.map(r=>{
    const t=String(r.tipo||'').toUpperCase();
    const bg=COLOR[t]||'rgba(122,132,160,.1)';
    const col=CTEXT[t]||'var(--muted)';
    const tShort=t==='REAJUSTE COMBUSTIBLE'?'REAJUSTE':t==='ENTRADA-SALIDA'?'E-S':t;
    const horo=r.horometro?Number(r.horometro).toLocaleString()+' h':'—';
    return '<div class="tr" style="cursor:pointer" onclick="openGasoilEditModal('+JSON.stringify(r)+')">'+
      '<div class="tc" style="flex:.7;color:var(--muted);font-size:.8rem">'+(r.fecha||'—')+'</div>'+
      '<div class="tc" style="flex:.8"><span class="badge" style="background:'+bg+';color:'+col+';white-space:nowrap;font-size:.72rem">'+tShort+'</span></div>'+
      '<div class="tc" style="flex:.6;color:var(--muted);font-size:.78rem">'+(r.origen||'—')+'</div>'+
      '<div class="tc" style="flex:.8;font-weight:700;color:var(--text);font-size:.85rem">'+(r.destino||'—')+'</div>'+
      '<div class="tc" style="flex:.45;font-family:monospace;font-weight:700;color:'+col+';text-align:right">'+Number(r.litros||0).toLocaleString()+'</div>'+
      '<div class="tc prev-hide-sm" style="flex:.55;font-family:monospace;font-size:.78rem;color:var(--accent2);text-align:right">'+horo+'</div>'+
      '<div class="tc" style="flex:.25;text-align:right"><span style="color:var(--muted);font-size:.85rem">✎</span></div>'+
    '</div>';
  }).join('');
  el.innerHTML='<div class="tbl"><div class="tr th"><div class="tc" style="flex:.7">Fecha</div><div class="tc" style="flex:.8">Tipo</div><div class="tc" style="flex:.6">Origen</div><div class="tc" style="flex:.8">Destino</div><div class="tc" style="flex:.45;text-align:right">L</div><div class="tc prev-hide-sm" style="flex:.55;text-align:right">Horómetro</div><div class="tc" style="flex:.25"></div></div>'+rows+'</div>';
}

function renderGasoilConsumos(){
  const el=document.getElementById('gasoil-consumos-list');
  if(!el)return;
  if(!gasoilConsumos.length){el.innerHTML='<div class="tbl"><div class="empty">Sin datos de consumo</div></div>';return;}
  const rows=gasoilConsumos.map(c=>'<div class="tr">'+
    '<div class="tc" style="flex:1;font-weight:700;color:var(--text)">'+(c.activo||'—')+'</div>'+
    '<div class="tc" style="flex:.9;font-family:monospace;color:var(--accent);text-align:right">'+Number(c.litros||0).toLocaleString()+'</div>'+
    '<div class="tc" style="flex:.8;font-family:monospace;color:var(--muted);text-align:right">'+Number(c.max||0).toLocaleString()+'</div>'+
    '<div class="tc" style="flex:.8;font-family:monospace;color:var(--muted);text-align:right">'+Number(c.min||0).toLocaleString()+'</div>'+
    '<div class="tc" style="flex:.6;font-family:monospace;color:var(--accent2);text-align:right">'+(c.lh||0)+'</div>'+
  '</div>').join('');
  el.innerHTML='<div class="tbl"><div class="tr th"><div class="tc" style="flex:1">Activo</div><div class="tc" style="flex:.9;text-align:right">Total L</div><div class="tc" style="flex:.8;text-align:right">Horómetro</div><div class="tc" style="flex:.8;text-align:right">Mín</div><div class="tc" style="flex:.6;text-align:right">L/H</div></div>'+rows+'</div>';
}

function renderGasoilHorometros(){
  const el=document.getElementById('gasoil-horometros-list');
  if(!el)return;
  if(!gasoilConsumos.length){el.innerHTML='<div class="tbl"><div class="empty">Sin datos</div></div>';return;}
  // Ordenar por horómetro descendente
  const sorted=[...gasoilConsumos].filter(c=>c.activo).sort((a,b)=>Number(b.max||0)-Number(a.max||0));
  const rows=sorted.map(c=>{
    const horo=Number(c.max||0);
    const color=horo>0?'var(--accent2)':'var(--muted)';
    return '<div class="tr">'+
      '<div class="tc" style="flex:1.2;font-weight:700;color:var(--text)">'+(c.activo||'—')+'</div>'+
      '<div class="tc" style="flex:1;font-family:monospace;font-size:.9rem;font-weight:700;color:'+color+';text-align:right">'+(horo>0?horo.toLocaleString()+' h':'—')+'</div>'+
      '<div class="tc" style="flex:.8;font-family:monospace;color:var(--muted);text-align:right">'+(Number(c.litros||0).toLocaleString())+' L</div>'+
    '</div>';
  }).join('');
  el.innerHTML='<div class="tbl"><div class="tr th"><div class="tc" style="flex:1.2">Activo</div><div class="tc" style="flex:1;text-align:right">Horómetro actual</div><div class="tc" style="flex:.8;text-align:right">Total L</div></div>'+rows+'</div>';
}

function openGasoilEditModal(r){
  document.getElementById('gedit-rownum').value=r.id||r.rowNum||'';
  document.getElementById('gedit-tipo').value=r.tipo||'SALIDA';
  document.getElementById('gedit-origen').value=r.origen||'DEPOSITO 1';
  document.getElementById('gedit-destino').value=r.destino||'';
  document.getElementById('gedit-prov').value=r.proveedor||'';
  document.getElementById('gedit-litros').value=r.litros||'';
  document.getElementById('gedit-horometro').value=r.horometro||'';
  // Convert d/mm/yyyy to yyyy-mm-dd for date input
  let fechaVal='';
  const fp=String(r.fecha||'').split('/');
  if(fp.length===3){const dd=fp[0].padStart(2,'0'),mm=fp[1].padStart(2,'0'),yy=fp[2];fechaVal=yy+'-'+mm+'-'+dd;}
  else if(String(r.fecha||'').match(/^\d{4}-\d{2}-\d{2}/)){fechaVal=String(r.fecha).slice(0,10);}
  document.getElementById('gedit-fecha').value=fechaVal;
  document.getElementById('gedit-msg').textContent='';
  document.getElementById('gasoil-edit-modal').style.display='flex';
}

function closeGasoilEditModal(){
  document.getElementById('gasoil-edit-modal').style.display='none';
}

async function saveGasoilEdit(){
  const recordId=document.getElementById('gedit-rownum').value;
  const tipo=document.getElementById('gedit-tipo').value;
  const origen=document.getElementById('gedit-origen').value;
  const destino=document.getElementById('gedit-destino').value;
  const proveedor=document.getElementById('gedit-prov').value;
  const litros=parseFloat(document.getElementById('gedit-litros').value)||0;
  const fecha=document.getElementById('gedit-fecha').value;
  const msg=document.getElementById('gedit-msg');
  if(!fecha){msg.style.color='var(--danger)';msg.textContent='Introduce la fecha.';return;}
  if(!litros){msg.style.color='var(--danger)';msg.textContent='Introduce los litros.';return;}
  const [y,m,d]=fecha.split('-');
  const fechaFmt=parseInt(d)+'/'+m+'/'+y;
  msg.style.color='var(--muted)';msg.textContent='Guardando...';
  try{
    const horometro=parseInt(document.getElementById('gedit-horometro').value)||null;
    await apiPost({tipo:'editarGasoil',id:recordId,fecha:fechaFmt,proveedor,origen,destino,tipoMovimiento:tipo,litros,horometro});
    // Update local data
    const idx=gasoilData.findIndex(r=>String(r.id)===String(recordId));
    if(idx>=0)gasoilData[idx]={...gasoilData[idx],fecha:fechaFmt,proveedor,origen,destino,tipo,litros};
    closeGasoilEditModal();
    renderGasoilHistorial();
    setTimeout(()=>cargarGasoil(),1500); // reload to get updated stock
  }catch(e){msg.style.color='var(--danger)';msg.textContent='Error de conexión';}
}

async function guardarGasoil(){
  const tipo=document.getElementById('gasoil-tipo-sel').value;
  const origen=document.getElementById('gasoil-origen-sel').value;
  const destino=document.getElementById('gasoil-destino-sel').value;
  const proveedor=tipo==='ENTRADA'?document.getElementById('gasoil-prov-inp').value:'';
  const litros=parseFloat(document.getElementById('gasoil-litros-inp').value)||0;
  const fecha=document.getElementById('gasoil-fecha-inp').value;
  const msg=document.getElementById('gasoil-save-msg');
  if(!destino){if(msg){msg.style.color='var(--danger)';msg.textContent='Selecciona el destino.';}return;}
  if(!litros||litros<=0){if(msg){msg.style.color='var(--danger)';msg.textContent='Introduce los litros.';}return;}
  if(!fecha){if(msg){msg.style.color='var(--danger)';msg.textContent='Introduce la fecha.';}return;}
  const [y,m,d]=fecha.split('-');
  const fechaFmt=parseInt(d)+'/'+m+'/'+y;
  const horometroNuevo=parseInt(document.getElementById('gasoil-horometro-inp')?.value)||null;
  const payload={tipo:'gasoil',fecha:fechaFmt,proveedor,origen,destino,tipoMovimiento:tipo,litros,horometro:horometroNuevo};
  try{
    if(msg){msg.style.color='var(--muted)';msg.textContent='Guardando...';}
    await apiPost(payload);
    if(msg){msg.style.color='var(--accent)';msg.textContent='✓ Guardado correctamente';}
    gasoilData.push({fecha:fechaFmt,proveedor,origen,destino,tipo,litros});
    renderGasoilHistorial();
    document.getElementById('gasoil-litros-inp').value='';
    document.getElementById('gasoil-destino-sel').value='';
    setTimeout(()=>{if(msg)msg.textContent='';cargarGasoil();},2000);
  }catch(e){if(msg){msg.style.color='var(--danger)';msg.textContent='Error de conexión';}}
}

// Alias para compatibilidad con filtros
function renderGasoil(){renderGasoilHistorial();}

// ── PRODUCCIÓN ───────────────────────────────────────────────
let prodData=[];
let prodInited=false;

function initProduccion(){
  if(!prodInited){
    const aSel=document.getElementById('prod-anyo');
    const cy=new Date().getFullYear();
    aSel.innerHTML='';for(let y=cy;y>=cy-3;y--)aSel.innerHTML+=`<option value="${y}">${y}</option>`;
    const mSel=document.getElementById('prod-mes');
    if(mSel)mSel.value=String(new Date().getMonth()+1);
    prodInited=true;
  }
  cargarProduccion();
}

async function cargarProduccion(){
  const anyo=parseInt(document.getElementById('prod-anyo').value);
  document.getElementById('prod-tabla').innerHTML='<div style="padding:20px;text-align:center;color:var(--muted)">Cargando...</div>';
  try{
    // Cargar producción + ventas del año en paralelo para que % Ventas funcione
    const proms=[apiFetch('?accion=produccion&mes=0&anyo='+anyo)];
    if(!ventasData||!ventasData.length) proms.push(apiFetch('?accion=pedidos&dias=365'));
    const results=await Promise.all(proms);
    const json=results[0];
    if(!json.ok)throw new Error(json.error||'Error');
    prodData=json.data;
    // Si se cargaron ventas, actualizar ventasData global
    if(results[1]&&results[1].ok) ventasData=results[1].data;
    renderProduccion(anyo);
  }catch(e){
    document.getElementById('prod-tabla').innerHTML='<div style="padding:20px;text-align:center;color:var(--danger)">Error: '+e.message+'</div>';
  }
}



function renderProduccionFiltrada(){
  const anyo=parseInt(document.getElementById('prod-anyo').value);
  renderProduccion(anyo);
}

function renderProduccion(anyo){
  const mesSel=parseInt((document.getElementById('prod-mes')||{}).value)||0;
  const filteredByMonth=mesSel>0?prodData.filter(r=>{
    if(!r.fecha)return false;
    const parts=r.fecha.split('-');
    return parseInt(parts[1])===mesSel;
  }):prodData;
  // Solo mostrar días que tengan datos reales (producción, mantenimiento, festivo registrado)
  const allData=filteredByMonth.filter(r=>{
    const tn=Number(r.tnDia)||0;
    const t04=Number(r.t04)||0;const t412=Number(r.t412)||0;
    const t1220=Number(r.t1220)||0;const t2040=Number(r.t2040)||0;
    const horas=Number(r.horasPlanta)||0;
    const tipo=r.tipoDia||'';
    // Tiene producción o tiene tipo explícito (M/FS/F) o tiene horas
    return tn>0||t04>0||t412>0||t1220>0||t2040>0||horas>0||tipo==='M'||tipo==='FS'||tipo==='F';
  });
  const data=allData.filter(r=>r.tipoDia==='P'||Number(r.tnDia)>0);
  // Resumen
  const totalTn=allData.reduce((a,r)=>a+(Number(r.tnDia)||0),0);
  const total04=allData.reduce((a,r)=>a+(Number(r.t04)||0),0);
  const total412=allData.reduce((a,r)=>a+(Number(r.t412)||0),0);
  const total1220=allData.reduce((a,r)=>a+(Number(r.t1220)||0),0);
  const total2040=allData.reduce((a,r)=>a+(Number(r.t2040)||0),0);
  const totalMat=total04+total412+total1220+total2040;
  const pct04=totalMat>0?(total04/totalMat*100).toFixed(1):'0';
  const pct412=totalMat>0?(total412/totalMat*100).toFixed(1):'0';
  const pct1220=totalMat>0?(total1220/totalMat*100).toFixed(1):'0';
  const pct2040=totalMat>0?(total2040/totalMat*100).toFixed(1):'0';
  const diasProd=data.length;
  const totalHoras=allData.reduce((a,r)=>a+(Number(r.horasPlanta)||0),0);
  const rendMedio=totalHoras>0?(totalTn/totalHoras).toFixed(1):'—';

  document.getElementById('prod-resumen').innerHTML=`
    <div style="flex:1;min-width:120px;background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:12px 14px">
      <div style="font-size:.62rem;color:var(--muted);text-transform:uppercase;font-weight:700">Total Tn</div>
      <div style="font-size:1.4rem;font-weight:700;color:var(--accent2);font-family:'DM Mono',monospace">${totalTn.toLocaleString()}</div>
    </div>
    <div style="flex:1;min-width:120px;background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:12px 14px">
      <div style="font-size:.62rem;color:var(--muted);text-transform:uppercase;font-weight:700">Días producción</div>
      <div style="font-size:1.4rem;font-weight:700;color:var(--text);font-family:'DM Mono',monospace">${diasProd}</div>
    </div>
    <div style="flex:1;min-width:120px;background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:12px 14px">
      <div style="font-size:.62rem;color:var(--muted);text-transform:uppercase;font-weight:700">Rend. medio</div>
      <div style="font-size:1.4rem;font-weight:700;color:var(--accent);font-family:'DM Mono',monospace">${rendMedio} Tn/h</div>
    </div>
    <div style="flex:1;min-width:120px;background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:12px 14px">
      <div style="font-size:.62rem;color:var(--muted);text-transform:uppercase;font-weight:700">0/4</div>
      <div style="font-size:1.1rem;font-weight:700;color:var(--text);font-family:'DM Mono',monospace">${total04.toLocaleString()} <span style="font-size:.7rem;color:var(--muted)">Tn</span></div>
      <div style="font-size:.72rem;font-weight:600;color:var(--accent);margin-top:2px">${pct04}%</div>
    </div>
    <div style="flex:1;min-width:120px;background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:12px 14px">
      <div style="font-size:.62rem;color:var(--muted);text-transform:uppercase;font-weight:700">4/12</div>
      <div style="font-size:1.1rem;font-weight:700;color:var(--text);font-family:'DM Mono',monospace">${total412.toLocaleString()} <span style="font-size:.7rem;color:var(--muted)">Tn</span></div>
      <div style="font-size:.72rem;font-weight:600;color:var(--accent);margin-top:2px">${pct412}%</div>
    </div>
    <div style="flex:1;min-width:120px;background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:12px 14px">
      <div style="font-size:.62rem;color:var(--muted);text-transform:uppercase;font-weight:700">12/20</div>
      <div style="font-size:1.1rem;font-weight:700;color:var(--text);font-family:'DM Mono',monospace">${total1220.toLocaleString()} <span style="font-size:.7rem;color:var(--muted)">Tn</span></div>
      <div style="font-size:.72rem;font-weight:600;color:var(--accent);margin-top:2px">${pct1220}%</div>
    </div>
    <div style="flex:1;min-width:120px;background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:12px 14px">
      <div style="font-size:.62rem;color:var(--muted);text-transform:uppercase;font-weight:700">20/40</div>
      <div style="font-size:1.1rem;font-weight:700;color:var(--text);font-family:'DM Mono',monospace">${total2040.toLocaleString()} <span style="font-size:.7rem;color:var(--muted)">Tn</span></div>
      <div style="font-size:.72rem;font-weight:600;color:var(--accent);margin-top:2px">${pct2040}%</div>
    </div>`;

  // Gráfico barras
  const maxTn=Math.max(...allData.map(r=>Number(r.tnDia)||0),1);
  document.getElementById('prod-chart').innerHTML=allData.map(r=>{
    const tn=Number(r.tnDia)||0;
    const h=tn>0?Math.max(4,Math.round(tn/maxTn*290)):2;
    const fecha=r.fecha?r.fecha.substring(8,10):'';
    const col=r.tipoDia==='M'?'var(--danger)':r.tipoDia==='P'&&tn>0?'var(--accent)':'var(--border)';
    const realIdx=prodData.indexOf(r);
    return `<div style="display:flex;flex-direction:column;align-items:center;flex:1;min-width:18px;max-width:36px" title="${fecha}: ${tn} Tn">
      <div style="font-size:.55rem;color:var(--muted);margin-bottom:2px">${tn>0?tn:''}</div>
      <div style="width:100%;height:${h}px;background:${col};border-radius:3px 3px 0 0;cursor:pointer" onclick="editarProdDia(${realIdx})"></div>
      <div style="font-size:.55rem;color:var(--muted);margin-top:2px">${fecha}</div>
    </div>`;
  }).join('');

  // Horómetros
  let hPrim=0,hHP4=0,hOre=0;
  allData.forEach(r=>{hPrim+=Number(r.primarioH)||0;hHP4+=Number(r.hp4H)||0;hOre+=Number(r.oreSizerH)||0;});
  document.getElementById('prod-horometros').innerHTML=`
    <div style="display:flex;justify-content:space-between;padding:6px 0;border-bottom:1px solid var(--border)"><span>Primario</span><span style="font-family:'DM Mono',monospace;font-weight:700">${hPrim.toFixed(0)} h</span></div>
    <div style="display:flex;justify-content:space-between;padding:6px 0;border-bottom:1px solid var(--border)"><span>HP4</span><span style="font-family:'DM Mono',monospace;font-weight:700">${hHP4.toFixed(0)} h</span></div>
    <div style="display:flex;justify-content:space-between;padding:6px 0"><span>Ore Sizer</span><span style="font-family:'DM Mono',monospace;font-weight:700">${hOre.toFixed(0)} h</span></div>
    <div style="border-top:1px solid var(--accent);margin-top:6px;padding-top:6px;display:flex;justify-content:space-between;font-weight:700"><span>Total horas planta</span><span style="font-family:'DM Mono',monospace;color:var(--accent2)">${totalHoras.toFixed(0)} h</span></div>`;

  // Consumos gasoil del año
  if(typeof gasoilData==='undefined'||!gasoilData||!gasoilData.length){
    cargarGasoil().then(()=>renderProdGasoil(anyo));
  } else {
    renderProdGasoil(anyo);
  }

  // Porcentajes fabricación y ventas
  renderPctFabricacion(anyo);
  renderPctVentas(anyo);

  // Tabla — solo días con producción real (tnDia>0)
  const tablaData=allData.filter(r=>Number(r.tnDia)>0);
  let tbl='';
  if(!tablaData.length){
    tbl='<div style="padding:20px;text-align:center;color:var(--muted);font-size:.8rem">Sin días de producción en el período</div>';
  }else{
    // Cabecera con totales por árido del período filtrado
    const s04=tablaData.reduce((a,r)=>a+(Number(r.t04)||0),0);
    const s412=tablaData.reduce((a,r)=>a+(Number(r.t412)||0),0);
    const s1220=tablaData.reduce((a,r)=>a+(Number(r.t1220)||0),0);
    const s2040=tablaData.reduce((a,r)=>a+(Number(r.t2040)||0),0);
    const sTn=tablaData.reduce((a,r)=>a+(Number(r.tnDia)||0),0);
    const sH=tablaData.reduce((a,r)=>a+(Number(r.horasPlanta)||0),0);
    const cats=[
      {l:'0/4',v:s04},{l:'4/12',v:s412},{l:'12/20',v:s1220},{l:'20/40',v:s2040}
    ].filter(c=>c.v>0);
    const catBadges=cats.map(c=>`<span style="padding:3px 8px;border-radius:4px;background:var(--surface2);font-size:.68rem;font-weight:700">${c.l}: <span style="color:var(--accent2)">${c.v.toLocaleString()} Tn</span></span>`).join('');
    tbl=`<div style="display:flex;gap:8px;flex-wrap:wrap;align-items:center;padding:10px 12px;background:var(--surface2);border-bottom:1px solid var(--border)">
      <span style="font-size:.7rem;color:var(--muted);font-weight:700">${tablaData.length} días</span>
      <span style="padding:3px 8px;border-radius:4px;background:var(--surface2);font-size:.68rem;font-weight:700">Total: <span style="color:var(--accent2)">${sTn.toLocaleString()} Tn</span></span>
      ${catBadges}
      <span style="padding:3px 8px;border-radius:4px;background:var(--surface2);font-size:.68rem;font-weight:700">Horas: <span style="color:var(--accent)">${sH.toFixed(0)} h</span></span>
    </div>
    <table style="width:100%;border-collapse:collapse;font-size:.72rem;white-space:nowrap">
    <thead><tr style="background:var(--surface2);color:var(--muted);font-size:.62rem;text-transform:uppercase">
      <th style="padding:8px 6px;text-align:left">Fecha</th><th>Día</th>
      <th style="text-align:right">Tn/día</th><th style="text-align:right">0/4</th><th style="text-align:right">4/12</th>
      <th style="text-align:right">12/20</th><th style="text-align:right">20/40</th>
      <th style="text-align:right">Horas</th><th style="text-align:right">Rend.</th>
      <th>Obs.</th><th></th>
    </tr></thead><tbody>`;
    tablaData.forEach((r)=>{
      const realIdx=prodData.indexOf(r);
      const tn=Number(r.tnDia)||0;
      const f=r.fecha?r.fecha.substring(5).replace('-','/'):'-';
      tbl+=`<tr style="border-top:1px solid var(--border)">
        <td style="padding:6px">${f}</td><td style="text-align:center;color:var(--muted);font-size:.65rem">${r.diaSem||''}</td>
        <td style="text-align:right;font-family:'DM Mono',monospace;font-weight:700;color:var(--accent2)">${tn}</td>
        <td style="text-align:right;font-family:'DM Mono',monospace">${r.t04||'—'}</td>
        <td style="text-align:right;font-family:'DM Mono',monospace">${r.t412||'—'}</td>
        <td style="text-align:right;font-family:'DM Mono',monospace">${r.t1220||'—'}</td>
        <td style="text-align:right;font-family:'DM Mono',monospace">${r.t2040||'—'}</td>
        <td style="text-align:right;font-family:'DM Mono',monospace">${r.horasPlanta||'—'}</td>
        <td style="text-align:right;font-family:'DM Mono',monospace;color:var(--accent)">${r.rendimiento?Number(r.rendimiento).toFixed(1):'—'}</td>
        <td style="max-width:200px;overflow:hidden;text-overflow:ellipsis;font-size:.65rem;color:var(--muted)">${r.observaciones||''}</td>
        <td><button onclick="editarProdDia(${realIdx})" style="background:transparent;border:none;cursor:pointer;color:var(--accent);font-size:.72rem">&#9998;</button></td>
      </tr>`;
    });
    tbl+=`</tbody></table>`;
  }
  document.getElementById('prod-tabla').innerHTML=tbl;
}

function renderProdGasoil(anyo){
  const el=document.getElementById('prod-gasoil');
  // Usar gasoilData si está cargado
  if(typeof gasoilData==='undefined'||!gasoilData||!gasoilData.length){
    el.innerHTML='<div style="font-size:.75rem;color:var(--muted)">Carga Gasoil primero para ver consumos</div>';
    return;
  }
  // Filtrar por año
  const consumos={};let total=0;
  gasoilData.forEach(r=>{
    const d=parseFechaFact(r.fecha);
    if(!d||d.getFullYear()!==anyo)return;
    const t=String(r.tipo||'').toUpperCase();
    if(t==='ENTRADA'||t==='REAJUSTE COMBUSTIBLE')return;
    const dest=r.destino||r.maquina||'Otros';
    if(dest.toUpperCase().includes('DEPOSITO'))return;
    if(!consumos[dest])consumos[dest]=0;
    consumos[dest]+=Number(r.litros)||0;
    total+=Number(r.litros)||0;
  });
  if(!total){el.innerHTML='<div style="font-size:.75rem;color:var(--muted)">Sin consumos este año</div>';return;}

  // L/H de C32 y PRAMAC desde GASOIL_CONSUMOS
  const lhMap={};
  if(typeof gasoilConsumos!=='undefined'&&gasoilConsumos.length){
    gasoilConsumos.forEach(c=>{ lhMap[String(c.activo||'').toUpperCase()]=Number(c.lh)||0; });
  }

  el.innerHTML=Object.entries(consumos).sort((a,b)=>b[1]-a[1]).map(([dest,l])=>{
    const lh=lhMap[dest.toUpperCase()];
    const lhStr=lh?`<span style="color:var(--muted);font-size:.7rem;margin-left:6px">${lh} L/H</span>`:'';
    return `<div style="display:flex;justify-content:space-between;align-items:center;padding:5px 0;border-bottom:1px solid rgba(0,0,0,.06);font-size:.75rem">
      <span>${dest}${lhStr}</span><span style="font-family:'DM Mono',monospace;font-weight:700">${l.toFixed(0)} L</span></div>`;
  }).join('')+`<div style="display:flex;justify-content:space-between;padding:6px 0;font-weight:700;border-top:1px solid var(--accent);margin-top:4px;font-size:.78rem">
    <span>Total</span><span style="font-family:'DM Mono',monospace;color:var(--accent2)">${total.toFixed(0)} L</span></div>`;
}

function renderPctFabricacion(anyo){
  const el=document.getElementById('prod-pct-fab');
  if(!el)return;
  const meses=['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO','JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE'];
  const byMonth=Array.from({length:12},()=>({t04:0,t412:0,t1220:0,t2040:0}));
  prodData.forEach(r=>{
    if(!r.fecha)return;
    const m=parseInt(r.fecha.split('-')[1])-1;
    if(m<0||m>11)return;
    byMonth[m].t04+=Number(r.t04)||0;
    byMonth[m].t412+=Number(r.t412)||0;
    byMonth[m].t1220+=Number(r.t1220)||0;
    byMonth[m].t2040+=Number(r.t2040)||0;
  });
  const totY={t04:0,t412:0,t1220:0,t2040:0};
  byMonth.forEach(m=>{totY.t04+=m.t04;totY.t412+=m.t412;totY.t1220+=m.t1220;totY.t2040+=m.t2040;});
  function pctRow(d){
    const sum=d.t04+d.t412+d.t1220+d.t2040;
    if(!sum)return {p04:'—',p412:'—',p1220:'—',p2040:'—'};
    return {p04:Math.round(d.t04/sum*100)+'%',p412:Math.round(d.t412/sum*100)+'%',p1220:Math.round(d.t1220/sum*100)+'%',p2040:Math.round(d.t2040/sum*100)+'%'};
  }
  let html=`<table style="width:100%;border-collapse:collapse;font-size:.72rem">
    <thead><tr style="background:var(--surface2);color:var(--muted);font-size:.62rem;text-transform:uppercase">
      <th style="padding:6px;text-align:left"></th><th style="text-align:right">0/4</th><th style="text-align:right">4/12</th><th style="text-align:right">12/20</th><th style="text-align:right">20/40</th>
    </tr></thead><tbody>`;
  byMonth.forEach((m,i)=>{
    const sum=m.t04+m.t412+m.t1220+m.t2040;
    if(!sum)return;
    const p=pctRow(m);
    html+=`<tr style="border-top:1px solid var(--border)">
      <td style="padding:5px 6px;font-weight:600">${meses[i]}</td>
      <td style="text-align:right;font-family:'DM Mono',monospace">${p.p04}</td>
      <td style="text-align:right;font-family:'DM Mono',monospace">${p.p412}</td>
      <td style="text-align:right;font-family:'DM Mono',monospace">${p.p1220}</td>
      <td style="text-align:right;font-family:'DM Mono',monospace">${p.p2040}</td>
    </tr>`;
  });
  const pt=pctRow(totY);
  html+=`<tr style="border-top:2px solid var(--accent);font-weight:700">
    <td style="padding:5px 6px;color:var(--accent)">TOTAL</td>
    <td style="text-align:right;font-family:'DM Mono',monospace;color:var(--accent)">${pt.p04}</td>
    <td style="text-align:right;font-family:'DM Mono',monospace;color:var(--accent)">${pt.p412}</td>
    <td style="text-align:right;font-family:'DM Mono',monospace;color:var(--accent)">${pt.p1220}</td>
    <td style="text-align:right;font-family:'DM Mono',monospace;color:var(--accent)">${pt.p2040}</td>
  </tr></tbody></table>`;
  el.innerHTML=html;
}

function renderPctVentas(anyo){
  const el=document.getElementById('prod-pct-ventas');
  if(!el)return;
  if(!ventasData||!ventasData.length){el.innerHTML='<div style="font-size:.75rem;color:var(--muted)">Carga Ventas primero para ver porcentajes</div>';return;}
  const meses=['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO','JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE'];
  const byMonth=Array.from({length:12},()=>({v04:0,v412:0,v1220:0,v2040:0,otros:0}));
  ventasData.forEach(r=>{
    const d=parseFechaHoraObj(r.fechaHora);
    if(!d||d.getFullYear()!==anyo)return;
    const m=d.getMonth();
    const cat=getCat(r.productoNombre);
    const kg=Number(r.pesoNeto)||0;
    if(cat==='0/4')byMonth[m].v04+=kg;
    else if(cat==='4/12')byMonth[m].v412+=kg;
    else if(cat==='12/20')byMonth[m].v1220+=kg;
    else if(cat==='20/40')byMonth[m].v2040+=kg;
    else byMonth[m].otros+=kg;
  });
  const totY={v04:0,v412:0,v1220:0,v2040:0,otros:0};
  byMonth.forEach(m=>{totY.v04+=m.v04;totY.v412+=m.v412;totY.v1220+=m.v1220;totY.v2040+=m.v2040;totY.otros+=m.otros;});
  function pctRow(d){
    const sum=d.v04+d.v412+d.v1220+d.v2040+d.otros;
    if(!sum)return {p04:'—',p412:'—',p1220:'—',p2040:'—',pOtros:'—'};
    return {p04:Math.round(d.v04/sum*100)+'%',p412:Math.round(d.v412/sum*100)+'%',p1220:Math.round(d.v1220/sum*100)+'%',p2040:Math.round(d.v2040/sum*100)+'%',pOtros:Math.round(d.otros/sum*100)+'%'};
  }
  let html=`<table style="width:100%;border-collapse:collapse;font-size:.72rem">
    <thead><tr style="background:var(--surface2);color:var(--muted);font-size:.62rem;text-transform:uppercase">
      <th style="padding:6px;text-align:left"></th><th style="text-align:right">0/4</th><th style="text-align:right">4/12</th><th style="text-align:right">12/20</th><th style="text-align:right">20/40</th><th style="text-align:right">Otros</th>
    </tr></thead><tbody>`;
  byMonth.forEach((m,i)=>{
    const sum=m.v04+m.v412+m.v1220+m.v2040+m.otros;
    if(!sum)return;
    const p=pctRow(m);
    html+=`<tr style="border-top:1px solid var(--border)">
      <td style="padding:5px 6px;font-weight:600">${meses[i]}</td>
      <td style="text-align:right;font-family:'DM Mono',monospace">${p.p04}</td>
      <td style="text-align:right;font-family:'DM Mono',monospace">${p.p412}</td>
      <td style="text-align:right;font-family:'DM Mono',monospace">${p.p1220}</td>
      <td style="text-align:right;font-family:'DM Mono',monospace">${p.p2040}</td>
      <td style="text-align:right;font-family:'DM Mono',monospace">${p.pOtros}</td>
    </tr>`;
  });
  const pt=pctRow(totY);
  html+=`<tr style="border-top:2px solid var(--accent);font-weight:700">
    <td style="padding:5px 6px;color:var(--accent)">TOTAL</td>
    <td style="text-align:right;font-family:'DM Mono',monospace;color:var(--accent)">${pt.p04}</td>
    <td style="text-align:right;font-family:'DM Mono',monospace;color:var(--accent)">${pt.p412}</td>
    <td style="text-align:right;font-family:'DM Mono',monospace;color:var(--accent)">${pt.p1220}</td>
    <td style="text-align:right;font-family:'DM Mono',monospace;color:var(--accent)">${pt.p2040}</td>
    <td style="text-align:right;font-family:'DM Mono',monospace;color:var(--accent)">${pt.pOtros}</td>
  </tr></tbody></table>`;
  el.innerHTML=html;
}

function editarProdDia(idx){
  const r=prodData[idx];if(!r)return;
  document.getElementById('pe-fila').value=r.fila||'';
  document.getElementById('pe-anyo').value=document.getElementById('prod-anyo').value;
  document.getElementById('pe-fecha').value=r.fecha;
  document.getElementById('pe-tipo').value=r.tipoDia||'P';
  document.getElementById('pe-t04').value=r.t04||'';
  document.getElementById('pe-t412').value=r.t412||'';
  document.getElementById('pe-t1220').value=r.t1220||'';
  document.getElementById('pe-t2040').value=r.t2040||'';
  document.getElementById('pe-tndia').value=r.tnDia||'';
  document.getElementById('pe-horas').value=r.horasPlanta||'';
  document.getElementById('pe-obs').value=r.observaciones||'';
  document.getElementById('pe-msg').textContent='';
  document.getElementById('prod-edit-title').textContent='Editar '+r.fecha;
  document.getElementById('modal-prod-edit').classList.add('open');
}

function abrirAddProduccion(){
  document.getElementById('pe-fila').value='';
  document.getElementById('pe-anyo').value=document.getElementById('prod-anyo').value;
  document.getElementById('pe-fecha').value=dateStr(new Date());
  document.getElementById('pe-tipo').value='P';
  document.getElementById('pe-t04').value='';document.getElementById('pe-t412').value='';
  document.getElementById('pe-t1220').value='';document.getElementById('pe-t2040').value='';
  document.getElementById('pe-tndia').value='';document.getElementById('pe-horas').value='';
  document.getElementById('pe-obs').value='';document.getElementById('pe-msg').textContent='';
  document.getElementById('prod-edit-title').textContent='Añadir día de producción';
  document.getElementById('modal-prod-edit').classList.add('open');
}

function cerrarProdEdit(){document.getElementById('modal-prod-edit').classList.remove('open');}

async function guardarProdEdit(){
  const fila=document.getElementById('pe-fila').value;
  const anyo=document.getElementById('pe-anyo').value;
  const payload={
    anyo,
    tipoDia:document.getElementById('pe-tipo').value,
    t04:document.getElementById('pe-t04').value||0,
    t412:document.getElementById('pe-t412').value||0,
    t1220:document.getElementById('pe-t1220').value||0,
    t2040:document.getElementById('pe-t2040').value||0,
    tnDia:document.getElementById('pe-tndia').value||0,
    horasPlanta:document.getElementById('pe-horas').value||0,
    observaciones:document.getElementById('pe-obs').value,
  };
  const msg=document.getElementById('pe-msg');
  if(fila){
    payload.tipo='editProduccion';payload.fila=fila;
  }else{
    payload.tipo='addProduccion';payload.fecha=document.getElementById('pe-fecha').value;
  }
  try{
    const json=await apiPost(payload);
    if(json.ok){msg.style.color='var(--accent)';msg.textContent='Guardado';setTimeout(()=>{cerrarProdEdit();cargarProduccion();},800);}
    else{msg.style.color='var(--danger)';msg.textContent='Error: '+(json.error||'');}
  }catch(e){msg.style.color='var(--danger)';msg.textContent='Error de red';}
}

// ── FACTURACIÓN ──────────────────────────────────────────────
let factData=[];
let factInited=false;

function initFacturacion(){
  if(!factInited){
    // Rellenar select clientes
    const sel=document.getElementById('fact-cliente');
    sel.innerHTML='<option value="">Todos los clientes</option>'+CLIENTES.map(c=>`<option value="${c.nombre}">${c.nombre}</option>`).join('');
    // Inicializar fechas: desde primer día mes actual, hasta hoy
    const hoy=new Date();
    const primerDia=new Date(hoy.getFullYear(),hoy.getMonth(),1);
    document.getElementById('fact-fecha-desde').value=primerDia.getFullYear()+'-'+pad(primerDia.getMonth()+1)+'-'+pad(primerDia.getDate());
    document.getElementById('fact-fecha-hasta').value=hoy.getFullYear()+'-'+pad(hoy.getMonth()+1)+'-'+pad(hoy.getDate());
    initInformeMensual();
    factInited=true;
  }
  if(factData.length===0) cargarFacturacion();
  else { renderFacturacion(); renderInformeMensual(); }
}

async function cargarFacturacion(){
  document.getElementById('fact-desglose').innerHTML='<div style="color:var(--muted);padding:20px;text-align:center">Cargando pedidos...</div>';
  try{
    const json=await apiFetch('?accion=pedidos&dias=365');
    if(!json.ok)throw new Error(json.error);
    factData=json.data;
    renderFacturacion();
    renderInformeMensual();
  }catch(e){
    document.getElementById('fact-desglose').innerHTML='<div style="color:var(--danger);padding:20px;text-align:center">Error: '+e.message+'</div>';
  }
}

function parseFechaFact(fh){
  if(!fh)return null;
  if(fh instanceof Date)return fh;
  const s=String(fh);
  if(s.includes('T')){return new Date(s);}
  if(s.includes('/')){const p=s.split(' ')[0].split('/');if(p.length>=3){const yy=parseInt(p[2])<100?2000+parseInt(p[2]):parseInt(p[2]);return new Date(yy,parseInt(p[1])-1,parseInt(p[0]));}}
  if(/^\d{4}-\d{2}-\d{2}$/.test(s)){const p=s.split('-');return new Date(parseInt(p[0]),parseInt(p[1])-1,parseInt(p[2]));}
  return null;
}

function renderFacturacion(){
  const filtCli=document.getElementById('fact-cliente').value;
  const fechaDesdeStr=document.getElementById('fact-fecha-desde').value;
  const fechaHastaStr=document.getElementById('fact-fecha-hasta').value;

  let fechaDesde=null,fechaHasta=null;
  if(fechaDesdeStr){
    const [y,m,d]=fechaDesdeStr.split('-');
    fechaDesde=new Date(parseInt(y),parseInt(m)-1,parseInt(d),0,0,0);
  }
  if(fechaHastaStr){
    const [y,m,d]=fechaHastaStr.split('-');
    fechaHasta=new Date(parseInt(y),parseInt(m)-1,parseInt(d),23,59,59);
  }

  // Filtrar pedidos por intervalo fecha
  const pedidos=factData.filter(r=>{
    const d=parseFechaFact(r.fechaHora)||parseFechaFact(r.fechaPedido);
    if(!d)return false;
    if(fechaDesde&&d<fechaDesde)return false;
    if(fechaHasta&&d>fechaHasta)return false;
    if(filtCli&&(r.nombreCliente||'').trim()!==filtCli.trim())return false;
    return true;
  });

  // Agrupar por cliente → proyecto → producto
  const clientes={};
  pedidos.forEach(r=>{
    const cli=(r.nombreCliente||'Sin cliente').trim();
    const proy=r.proyectoName||r.proyectoCod||'Sin proyecto';
    const prod=r.productoNombre||r.productoCod||'Sin producto';
    const neto=Number(r.pesoNeto)||0;
    if(!clientes[cli])clientes[cli]={proyectos:{},totalKg:0,totalEur:0};
    if(!clientes[cli].proyectos[proy])clientes[cli].proyectos[proy]={productos:{},totalKg:0,totalEur:0,proyectoCod:r.proyectoCod||''};
    if(!clientes[cli].proyectos[proy].productos[prod])clientes[cli].proyectos[proy].productos[prod]={viajes:0,kg:0,cod:r.productoCod||''};
    clientes[cli].proyectos[proy].productos[prod].viajes++;
    clientes[cli].proyectos[proy].productos[prod].kg+=neto;
  });

  // Calcular precios
  let grandTotalBase=0,grandTotalIgic=0,grandTotalKg=0,grandViajes=0;
  Object.entries(clientes).forEach(([cli,cData])=>{
    Object.entries(cData.proyectos).forEach(([proy,pData])=>{
      Object.entries(pData.productos).forEach(([prod,info])=>{
        const precioTn=getPrecioTn(cli,prod);
        const tn=info.kg/1000;
        const importe=tn*precioTn;
        info.precioTn=precioTn;info.importe=importe;info.igic=importe*IGIC_PCT/100;
        pData.totalKg+=info.kg;pData.totalEur+=importe;
      });
      cData.totalKg+=pData.totalKg;cData.totalEur+=pData.totalEur;
    });
    const cIgic=cData.totalEur*IGIC_PCT/100;
    grandTotalBase+=cData.totalEur;grandTotalIgic+=cIgic;grandTotalKg+=cData.totalKg;
    grandViajes+=pedidos.filter(r=>r.nombreCliente===cli).length;
  });

  // Resumen cards
  const numClientes=Object.keys(clientes).length;
  document.getElementById('fact-resumen').innerHTML=`
    <div style="flex:1;min-width:140px;background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:14px 16px">
      <div style="font-size:.68rem;color:var(--muted);text-transform:uppercase;font-weight:700;margin-bottom:4px">Clientes</div>
      <div style="font-size:1.4rem;font-weight:700;color:var(--text);font-family:'DM Mono',monospace">${numClientes}</div>
    </div>
    <div style="flex:1;min-width:140px;background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:14px 16px">
      <div style="font-size:.68rem;color:var(--muted);text-transform:uppercase;font-weight:700;margin-bottom:4px">Viajes</div>
      <div style="font-size:1.4rem;font-weight:700;color:var(--text);font-family:'DM Mono',monospace">${grandViajes}</div>
    </div>
    <div style="flex:1;min-width:140px;background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:14px 16px">
      <div style="font-size:.68rem;color:var(--muted);text-transform:uppercase;font-weight:700;margin-bottom:4px">Toneladas</div>
      <div style="font-size:1.4rem;font-weight:700;color:var(--accent);font-family:'DM Mono',monospace">${(grandTotalKg/1000).toFixed(2)}</div>
    </div>
    <div style="flex:1;min-width:140px;background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:14px 16px">
      <div style="font-size:.68rem;color:var(--muted);text-transform:uppercase;font-weight:700;margin-bottom:4px">Base imponible</div>
      <div style="font-size:1.4rem;font-weight:700;color:var(--accent2);font-family:'DM Mono',monospace">${grandTotalBase.toFixed(2)} €</div>
    </div>
    <div style="flex:1;min-width:140px;background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:14px 16px">
      <div style="font-size:.68rem;color:var(--muted);text-transform:uppercase;font-weight:700;margin-bottom:4px">IGIC (${IGIC_PCT}%)</div>
      <div style="font-size:1.1rem;font-weight:700;color:var(--text);font-family:'DM Mono',monospace">${grandTotalIgic.toFixed(2)} €</div>
    </div>
    <div style="flex:1;min-width:140px;background:rgba(107,125,46,.08);border:1px solid rgba(107,125,46,.3);border-radius:var(--radius);padding:14px 16px">
      <div style="font-size:.68rem;color:var(--accent2);text-transform:uppercase;font-weight:700;margin-bottom:4px">Total con IGIC</div>
      <div style="font-size:1.4rem;font-weight:700;color:var(--accent2);font-family:'DM Mono',monospace">${(grandTotalBase+grandTotalIgic).toFixed(2)} €</div>
    </div>`;

  // Desglose por cliente
  if(!numClientes){
    document.getElementById('fact-desglose').innerHTML='<div style="color:var(--muted);padding:20px;text-align:center;font-size:.82rem">Sin pedidos en este periodo</div>';
    return;
  }

  window._bcClientesData = [];
  let html='';
  Object.entries(clientes).sort((a,b)=>b[1].totalEur-a[1].totalEur).forEach(([cli,cData],bcIdx)=>{
    const cIgic=cData.totalEur*IGIC_PCT/100;
    window._bcClientesData[bcIdx] = { cli, cData };
    html+=`<div style="background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);margin-bottom:16px;overflow:hidden">`;
    html+=`<div style="padding:14px 16px;border-bottom:1px solid var(--border);display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px">
      <div>
        <div style="font-family:'Syne',sans-serif;font-size:.92rem;font-weight:700;color:var(--accent)">${cli}</div>
        <div style="font-size:.68rem;color:var(--muted);margin-top:2px">${(cData.totalKg/1000).toFixed(2)} Tn · ${pedidos.filter(r=>(r.nombreCliente||'').trim()===cli).length} viajes</div>
      </div>
      <div style="display:flex;align-items:center;gap:10px">
        <span id="fact-estado-${bcIdx}" style="font-size:.7rem;font-weight:700;padding:4px 10px;border-radius:6px;background:rgba(150,150,150,.15);color:var(--muted)">⏳</span>
        <div style="text-align:right">
          <div style="font-family:'DM Mono',monospace;font-size:1.1rem;font-weight:700;color:var(--accent2)">${cData.totalEur.toFixed(2)} €</div>
          <div style="font-size:.65rem;color:var(--muted)">+ ${cIgic.toFixed(2)} € IGIC = ${(cData.totalEur+cIgic).toFixed(2)} €</div>
        </div>
        <button onclick="abrirModalAlbaranes('${cli.replace(/'/g,"\\'")}')" style="background:#0078d4;color:#fff;border:none;border-radius:8px;padding:8px 14px;font-size:.72rem;font-weight:700;cursor:pointer;white-space:nowrap">📋 Facturar albaranes</button>
        <button onclick="exportarExcelCliente(${bcIdx})" style="background:#217346;color:#fff;border:none;border-radius:8px;padding:8px 14px;font-size:.72rem;font-weight:700;cursor:pointer;white-space:nowrap">📥 Excel</button>
      </div>
    </div>`;

    Object.entries(cData.proyectos).forEach(([proy,pData],pIdx)=>{
      const bcPKey = `${bcIdx}_${pIdx}`;
      window._bcClientesData[bcPKey] = { cli, cData: { proyectos: { [proy]: pData } } };
      html+=`<div style="padding:10px 16px;border-bottom:1px solid var(--border);background:var(--surface2)">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px">
          <div style="font-size:.78rem;font-weight:700;color:#f5c842">📍 ${proy}</div>
          <button onclick="enviarBCCliente('${bcPKey}',this)" style="background:#0078d4;color:#fff;border:none;border-radius:6px;padding:5px 10px;font-size:.68rem;font-weight:700;cursor:pointer">Enviar a BC</button>
        </div>
        <div style="display:flex;padding:4px 0;font-size:.65rem;font-weight:700;color:var(--muted);text-transform:uppercase;gap:8px">
          <div style="flex:2">Producto</div><div style="flex:.7;text-align:right">Viajes</div><div style="flex:1;text-align:right">Kg Neto</div><div style="flex:.7;text-align:right">Tn</div><div style="flex:.8;text-align:right">€/Tn</div><div style="flex:1;text-align:right">Importe</div>
        </div>`;
      Object.entries(pData.productos).forEach(([prod,info])=>{
        html+=`<div style="display:flex;padding:5px 0;font-size:.75rem;color:var(--text);gap:8px;border-top:1px solid rgba(0,0,0,.06)">
          <div style="flex:2">${prod}</div>
          <div style="flex:.7;text-align:right;font-family:'DM Mono',monospace">${info.viajes}</div>
          <div style="flex:1;text-align:right;font-family:'DM Mono',monospace">${info.kg.toLocaleString()}</div>
          <div style="flex:.7;text-align:right;font-family:'DM Mono',monospace">${(info.kg/1000).toFixed(2)}</div>
          <div style="flex:.8;text-align:right;font-family:'DM Mono',monospace;color:var(--muted)">${info.precioTn.toFixed(2)}</div>
          <div style="flex:1;text-align:right;font-family:'DM Mono',monospace;font-weight:700;color:var(--accent2)">${info.importe.toFixed(2)} €</div>
        </div>`;
      });
      html+=`<div style="display:flex;padding:6px 0;font-size:.72rem;font-weight:700;color:var(--text);gap:8px;border-top:1px solid var(--border);margin-top:4px">
        <div style="flex:2">Subtotal ${proy}</div><div style="flex:.7"></div><div style="flex:1;text-align:right;font-family:'DM Mono',monospace">${pData.totalKg.toLocaleString()}</div><div style="flex:.7;text-align:right;font-family:'DM Mono',monospace">${(pData.totalKg/1000).toFixed(2)}</div><div style="flex:.8"></div><div style="flex:1;text-align:right;font-family:'DM Mono',monospace;color:var(--accent2)">${pData.totalEur.toFixed(2)} €</div>
      </div>`;
      html+=`</div>`;
    });

    // Resumen total por producto (sumado de todos los proyectos)
    const prodResumen={};
    Object.values(cData.proyectos).forEach(pData=>{
      Object.entries(pData.productos).forEach(([prod,info])=>{
        if(!prodResumen[prod])prodResumen[prod]={viajes:0,kg:0,importe:0};
        prodResumen[prod].viajes+=info.viajes;
        prodResumen[prod].kg+=info.kg;
        prodResumen[prod].importe+=info.importe;
      });
    });
    if(Object.keys(cData.proyectos).length>1){
      html+=`<div style="padding:10px 16px;background:rgba(107,125,46,.06);border-top:2px solid rgba(107,125,46,.3)">
        <div style="font-size:.72rem;font-weight:700;color:var(--accent2);text-transform:uppercase;margin-bottom:8px;letter-spacing:.04em">▸ Total por producto</div>
        <div style="display:flex;padding:4px 0;font-size:.65rem;font-weight:700;color:var(--muted);text-transform:uppercase;gap:8px">
          <div style="flex:2">Producto</div><div style="flex:.7;text-align:right">Viajes</div><div style="flex:1;text-align:right">Kg Neto</div><div style="flex:.7;text-align:right">Tn</div><div style="flex:.8"></div><div style="flex:1;text-align:right">Importe</div>
        </div>`;
      Object.entries(prodResumen).sort((a,b)=>b[1].kg-a[1].kg).forEach(([prod,info])=>{
        html+=`<div style="display:flex;padding:5px 0;font-size:.75rem;color:var(--text);gap:8px;border-top:1px solid rgba(107,125,46,.15)">
          <div style="flex:2;font-weight:600">${prod}</div>
          <div style="flex:.7;text-align:right;font-family:'DM Mono',monospace">${info.viajes}</div>
          <div style="flex:1;text-align:right;font-family:'DM Mono',monospace">${info.kg.toLocaleString()}</div>
          <div style="flex:.7;text-align:right;font-family:'DM Mono',monospace;font-weight:700;color:var(--accent)">${(info.kg/1000).toFixed(2)}</div>
          <div style="flex:.8"></div>
          <div style="flex:1;text-align:right;font-family:'DM Mono',monospace;font-weight:700;color:var(--accent2)">${info.importe.toFixed(2)} €</div>
        </div>`;
      });
      html+=`</div>`;
    }
    html+=`</div>`;
  });
  document.getElementById('fact-desglose').innerHTML=html;

  // Comprobar estado facturación en BC (background)
  _comprobarEstadoFacturacionBC(clientes);
}

async function _comprobarEstadoFacturacionBC(clientes) {
  try {
    if (typeof getBCToken !== 'function') return;
    const token = await getBCToken();
    const headers = { 'Authorization': `Bearer ${token}` };
    const base = `https://api.businesscentral.dynamics.com/v2.0/${BC_TENANT}/${BC_ENV}/api/v2.0/companies`;
    const cRes = await fetch(base, { headers });
    const cJson = await cRes.json();
    const company = (cJson.value || []).find(c => c.name.trim() === BC_COMPANY.trim());
    if (!company) return;

    // Determinar mes/año del filtro
    const now = new Date();
    const mesEl = document.getElementById('fact-mes');
    const anyoEl = document.getElementById('fact-anyo');
    const fechaDesdeStr = document.getElementById('fact-fecha-desde').value;
    const fechaHastaStr = document.getElementById('fact-fecha-hasta').value;
    let mes, anyo;
    if (fechaDesdeStr) {
      const [y, m] = fechaDesdeStr.split('-');
      mes = parseInt(m); anyo = parseInt(y);
    } else {
      mes = (mesEl?.value ? parseInt(mesEl.value) : now.getMonth()) + 1;
      anyo = anyoEl?.value ? parseInt(anyoEl.value) : now.getFullYear();
    }
    const mesPad = String(mes).padStart(2, '0');

    // Obtener todos los customerNumber únicos de los clientes mostrados
    const cliCustMap = {};
    Object.keys(clientes).forEach(cli => {
      const pedidoCli = factData?.find(r => (r.nombreCliente || '').trim() === cli.trim());
      cliCustMap[cli] = pedidoCli?.codigoCliente || '';
    });

    // Calcular último día del mes
    const ultimoDia = new Date(anyo, mes, 0).getDate();
    const desde = `${anyo}-${mesPad}-01`;
    const hasta = `${anyo}-${mesPad}-${String(ultimoDia).padStart(2,'0')}`;

    // Buscar borradores (salesInvoices)
    let allFacturas = [];
    try {
      const draftUrl = `${base}(${company.id})/salesInvoices?$filter=invoiceDate ge ${desde} and invoiceDate le ${hasta}&$select=id,number,customerNumber,customerName,totalAmountIncludingTax,externalDocumentNumber,status&$top=500`;
      const draftRes = await fetch(draftUrl, { headers });
      if (draftRes.ok) {
        const draftData = await draftRes.json();
        (draftData.value || []).forEach(f => { f._tipo = 'Borrador'; allFacturas.push(f); });
      }
    } catch (e) {}

    // Buscar facturas registradas/contabilizadas (postedSalesInvoices)
    try {
      const postUrl = `${base}(${company.id})/postedSalesInvoices?$filter=invoiceDate ge ${desde} and invoiceDate le ${hasta}&$select=id,number,customerNumber,customerName,totalAmountIncludingTax,externalDocumentNumber&$top=500`;
      const postRes = await fetch(postUrl, { headers });
      if (postRes.ok) {
        const postData = await postRes.json();
        const ids = new Set(allFacturas.map(f => f.id));
        (postData.value || []).forEach(f => { if (!ids.has(f.id)) { f._tipo = 'Registrada'; allFacturas.push(f); } });
      }
    } catch (e) {}

    // Mapear por customerNumber
    const facturadoMap = {};
    allFacturas.forEach(f => {
      const custNo = (f.customerNumber || '').trim();
      if (custNo) {
        if (!facturadoMap[custNo]) facturadoMap[custNo] = [];
        facturadoMap[custNo].push(f);
      }
    });

    // Actualizar badges
    Object.keys(clientes).sort((a, b) => clientes[b].totalEur - clientes[a].totalEur).forEach((cli, bcIdx) => {
      const el = document.getElementById(`fact-estado-${bcIdx}`);
      if (!el) return;
      const custNo = cliCustMap[cli];
      const facts = custNo ? facturadoMap[custNo] : null;
      if (facts && facts.length > 0) {
        const nums = facts.map(f => f.number).filter(Boolean).join(', ');
        const totalBC = facts.reduce((s, f) => s + (f.totalAmountIncludingTax || 0), 0);
        el.textContent = `✅ Facturada ${nums} · ${totalBC.toLocaleString('es-ES', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} €`;
        el.style.background = 'rgba(46,125,50,.12)';
        el.style.color = '#2e7d32';
      } else {
        el.textContent = '⬚ Sin facturar';
        el.style.background = 'rgba(211,47,47,.1)';
        el.style.color = '#d32f2f';
      }
    });
  } catch (e) {
    console.warn('Error comprobando estado facturación BC:', e.message);
  }
}

// ── INFORME MENSUAL POR CLIENTE ──────────────────────────────
function initInformeMensual() {
  const hoy = new Date();
  const selMes = document.getElementById('fact-informe-mes');
  const selAnyo = document.getElementById('fact-informe-anyo');
  selMes.value = hoy.getMonth();
  // Años disponibles
  const anyoActual = hoy.getFullYear();
  selAnyo.innerHTML = '';
  for (let y = anyoActual; y >= anyoActual - 3; y--) {
    const o = document.createElement('option');
    o.value = y; o.textContent = y;
    selAnyo.appendChild(o);
  }
  selAnyo.value = anyoActual;
}

function renderInformeMensual() {
  const mes = parseInt(document.getElementById('fact-informe-mes').value);
  const anyo = parseInt(document.getElementById('fact-informe-anyo').value);
  const body = document.getElementById('fact-informe-body');
  const MESES_NOMBRE = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];

  if (!factData.length) {
    body.innerHTML = '<div style="padding:20px;text-align:center;color:var(--muted);font-size:.82rem">Carga datos para ver el informe</div>';
    return;
  }

  // Filtrar pedidos del mes/año seleccionado
  const pedidosMes = factData.filter(r => {
    const d = parseFechaFact(r.fechaHora) || parseFechaFact(r.fechaPedido);
    if (!d) return false;
    return d.getMonth() === mes && d.getFullYear() === anyo;
  });

  if (!pedidosMes.length) {
    body.innerHTML = `<div style="padding:20px;text-align:center;color:var(--muted);font-size:.82rem">Sin pedidos en ${MESES_NOMBRE[mes]} ${anyo}</div>`;
    return;
  }

  // Agrupar por cliente
  const clientes = {};
  pedidosMes.forEach(r => {
    const cli = (r.nombreCliente || 'Sin cliente').trim();
    if (!clientes[cli]) clientes[cli] = { viajes: 0, kg: 0, importe: 0 };
    clientes[cli].viajes++;
    const neto = Number(r.pesoNeto) || 0;
    clientes[cli].kg += neto;
    const precioTn = getPrecioTn(cli, r.productoNombre || r.productoCod || '');
    clientes[cli].importe += (neto / 1000) * precioTn;
  });

  // Ordenar por importe desc
  const sorted = Object.entries(clientes).sort((a, b) => b[1].importe - a[1].importe);
  const totalKg = sorted.reduce((s, [, c]) => s + c.kg, 0);
  const totalImporte = sorted.reduce((s, [, c]) => s + c.importe, 0);
  const totalIgic = totalImporte * IGIC_PCT / 100;
  const totalViajes = sorted.reduce((s, [, c]) => s + c.viajes, 0);

  let html = `<div style="padding:12px 16px;background:rgba(107,125,46,.06);border-bottom:1px solid var(--border);display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px">
    <div style="font-size:.85rem;font-weight:700;color:var(--accent)">${MESES_NOMBRE[mes]} ${anyo}</div>
    <div style="font-size:.78rem;color:var(--text);font-weight:600">${sorted.length} clientes · ${totalViajes} viajes · ${(totalKg/1000).toFixed(2)} Tn · <span style="color:var(--accent2)">${totalImporte.toFixed(2)} € + ${totalIgic.toFixed(2)} € IGIC</span></div>
  </div>`;

  html += `<div style="overflow-x:auto"><table style="width:100%;border-collapse:collapse;font-size:.78rem">
    <thead><tr style="background:var(--surface2)">
      <th style="padding:10px 14px;text-align:left;font-weight:700;font-size:.7rem;text-transform:uppercase;color:var(--muted);letter-spacing:.04em">Cliente</th>
      <th style="padding:10px 14px;text-align:right;font-weight:700;font-size:.7rem;text-transform:uppercase;color:var(--muted)">Viajes</th>
      <th style="padding:10px 14px;text-align:right;font-weight:700;font-size:.7rem;text-transform:uppercase;color:var(--muted)">Toneladas</th>
      <th style="padding:10px 14px;text-align:right;font-weight:700;font-size:.7rem;text-transform:uppercase;color:var(--muted)">Base imp.</th>
      <th style="padding:10px 14px;text-align:right;font-weight:700;font-size:.7rem;text-transform:uppercase;color:var(--muted)">IGIC</th>
      <th style="padding:10px 14px;text-align:right;font-weight:700;font-size:.7rem;text-transform:uppercase;color:var(--muted)">Total</th>
      <th style="padding:10px 14px;text-align:right;font-weight:700;font-size:.7rem;text-transform:uppercase;color:var(--muted)">% Factur.</th>
    </tr></thead><tbody>`;

  sorted.forEach(([cli, c]) => {
    const tn = c.kg / 1000;
    const igic = c.importe * IGIC_PCT / 100;
    const pct = totalImporte > 0 ? (c.importe / totalImporte * 100) : 0;
    const barW = Math.max(pct, 1);
    html += `<tr style="border-bottom:1px solid var(--border)">
      <td style="padding:9px 14px;font-weight:600">${escapeHTML(cli)}</td>
      <td style="padding:9px 14px;text-align:right;font-family:'DM Mono',monospace">${c.viajes}</td>
      <td style="padding:9px 14px;text-align:right;font-family:'DM Mono',monospace;font-weight:600;color:var(--accent)">${tn.toFixed(2)}</td>
      <td style="padding:9px 14px;text-align:right;font-family:'DM Mono',monospace">${c.importe.toFixed(2)} €</td>
      <td style="padding:9px 14px;text-align:right;font-family:'DM Mono',monospace;color:var(--muted)">${igic.toFixed(2)} €</td>
      <td style="padding:9px 14px;text-align:right;font-family:'DM Mono',monospace;font-weight:700;color:var(--accent2)">${(c.importe + igic).toFixed(2)} €</td>
      <td style="padding:9px 14px;text-align:right">
        <div style="display:flex;align-items:center;justify-content:flex-end;gap:6px">
          <div style="width:60px;height:6px;background:var(--border);border-radius:3px;overflow:hidden"><div style="width:${barW}%;height:100%;background:var(--accent2);border-radius:3px"></div></div>
          <span style="font-family:'DM Mono',monospace;font-size:.72rem;min-width:38px;text-align:right">${pct.toFixed(1)}%</span>
        </div>
      </td>
    </tr>`;
  });

  // Fila totales
  html += `<tr style="background:rgba(107,125,46,.08);font-weight:700;border-top:2px solid var(--accent)">
    <td style="padding:10px 14px;font-size:.82rem">TOTAL</td>
    <td style="padding:10px 14px;text-align:right;font-family:'DM Mono',monospace">${totalViajes}</td>
    <td style="padding:10px 14px;text-align:right;font-family:'DM Mono',monospace;color:var(--accent)">${(totalKg/1000).toFixed(2)}</td>
    <td style="padding:10px 14px;text-align:right;font-family:'DM Mono',monospace">${totalImporte.toFixed(2)} €</td>
    <td style="padding:10px 14px;text-align:right;font-family:'DM Mono',monospace">${totalIgic.toFixed(2)} €</td>
    <td style="padding:10px 14px;text-align:right;font-family:'DM Mono',monospace;color:var(--accent2)">${(totalImporte + totalIgic).toFixed(2)} €</td>
    <td style="padding:10px 14px;text-align:right;font-family:'DM Mono',monospace">100%</td>
  </tr>`;

  html += '</tbody></table></div>';
  body.innerHTML = html;
}

function imprimirInformeMensual() {
  const MESES_NOMBRE = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
  const mes = parseInt(document.getElementById('fact-informe-mes').value);
  const anyo = parseInt(document.getElementById('fact-informe-anyo').value);
  const contenido = document.getElementById('fact-informe-body').innerHTML;
  if (!factData.length) return;

  const win = window.open('', '_blank');
  win.document.write(`<!DOCTYPE html><html><head><title>Informe ${MESES_NOMBRE[mes]} ${anyo} - ARIFOMA</title>
<style>
  *{margin:0;padding:0;box-sizing:border-box}
  body{font-family:Arial,Helvetica,sans-serif;font-size:9pt;color:#111;padding:12mm 15mm}
  .header{display:flex;justify-content:space-between;align-items:flex-end;margin-bottom:16px;padding-bottom:10px;border-bottom:2px solid #6b7d2e}
  .logo{font-size:16pt;font-weight:900;color:#6b7d2e;letter-spacing:.05em}
  .subtitle{font-size:8pt;color:#666}
  .title{font-size:13pt;font-weight:700;color:#333;text-align:right}
  .date{font-size:8pt;color:#888;text-align:right;margin-top:2px}
  table{width:100%;border-collapse:collapse;margin-top:8px}
  th{background:#f5f5f0;font-size:7.5pt;font-weight:700;text-transform:uppercase;color:#555;padding:7px 10px;border:1px solid #ccc;letter-spacing:.03em}
  td{padding:6px 10px;border:1px solid #ddd;font-size:8.5pt}
  .r{text-align:right;font-family:'Courier New',monospace}
  .b{font-weight:700}
  .accent{color:#6b7d2e}
  .total-row{background:#f0f2e6;font-weight:700;border-top:2px solid #6b7d2e}
  .total-row td{font-size:9pt}
  .bar-cell{width:90px}
  .bar-bg{width:60px;height:5px;background:#e0e0e0;border-radius:3px;display:inline-block;vertical-align:middle}
  .bar-fg{height:5px;background:#6b7d2e;border-radius:3px;display:block}
  .footer{margin-top:20px;font-size:7pt;color:#999;text-align:center;border-top:1px solid #ddd;padding-top:8px}
  @media print{body{padding:8mm 10mm}.no-print{display:none!important}}
</style></head><body>
<div class="header">
  <div><div class="logo">ARIFOMA</div><div class="subtitle">Cantera Mesa de las Cañadas</div></div>
  <div><div class="title">Informe Mensual · ${MESES_NOMBRE[mes]} ${anyo}</div><div class="date">Generado el ${new Date().toLocaleDateString('es-ES')}</div></div>
</div>
<div class="no-print" style="margin-bottom:12px"><button onclick="window.print()" style="background:#6b7d2e;color:#fff;border:none;border-radius:6px;padding:8px 20px;font-size:10pt;font-weight:700;cursor:pointer">🖨 Imprimir</button> <button onclick="window.close()" style="background:#eee;color:#333;border:1px solid #ccc;border-radius:6px;padding:8px 20px;font-size:10pt;cursor:pointer;margin-left:8px">Cerrar</button></div>`);

  // Rebuild table for print (clean, no CSS vars)
  const pedidosMes = factData.filter(r => {
    const d = parseFechaFact(r.fechaHora) || parseFechaFact(r.fechaPedido);
    if (!d) return false;
    return d.getMonth() === mes && d.getFullYear() === anyo;
  });

  const clientes = {};
  pedidosMes.forEach(r => {
    const cli = (r.nombreCliente || 'Sin cliente').trim();
    if (!clientes[cli]) clientes[cli] = { viajes: 0, kg: 0, importe: 0 };
    clientes[cli].viajes++;
    const neto = Number(r.pesoNeto) || 0;
    clientes[cli].kg += neto;
    clientes[cli].importe += (neto / 1000) * getPrecioTn(cli, r.productoNombre || r.productoCod || '');
  });

  const sorted = Object.entries(clientes).sort((a, b) => b[1].importe - a[1].importe);
  const totalKg = sorted.reduce((s, [, c]) => s + c.kg, 0);
  const totalImporte = sorted.reduce((s, [, c]) => s + c.importe, 0);
  const totalIgic = totalImporte * IGIC_PCT / 100;
  const totalViajes = sorted.reduce((s, [, c]) => s + c.viajes, 0);

  let tbl = `<table><thead><tr>
    <th style="text-align:left">Cliente</th><th>Viajes</th><th>Toneladas</th><th>Base imp.</th><th>IGIC (${IGIC_PCT}%)</th><th>Total</th><th>%</th>
  </tr></thead><tbody>`;

  sorted.forEach(([cli, c]) => {
    const tn = c.kg / 1000;
    const igic = c.importe * IGIC_PCT / 100;
    const pct = totalImporte > 0 ? (c.importe / totalImporte * 100) : 0;
    tbl += `<tr>
      <td class="b">${cli}</td>
      <td class="r">${c.viajes}</td>
      <td class="r b accent">${tn.toFixed(2)}</td>
      <td class="r">${c.importe.toFixed(2)} €</td>
      <td class="r" style="color:#888">${igic.toFixed(2)} €</td>
      <td class="r b accent">${(c.importe + igic).toFixed(2)} €</td>
      <td class="r bar-cell"><div class="bar-bg"><div class="bar-fg" style="width:${Math.max(pct, 1)}%"></div></div> ${pct.toFixed(1)}%</td>
    </tr>`;
  });

  tbl += `<tr class="total-row">
    <td>TOTAL</td>
    <td class="r">${totalViajes}</td>
    <td class="r accent">${(totalKg/1000).toFixed(2)}</td>
    <td class="r">${totalImporte.toFixed(2)} €</td>
    <td class="r">${totalIgic.toFixed(2)} €</td>
    <td class="r accent">${(totalImporte + totalIgic).toFixed(2)} €</td>
    <td class="r">100%</td>
  </tr></tbody></table>`;

  win.document.write(tbl);
  win.document.write(`<div class="footer">ARIFOMA · Cantera Mesa de las Cañadas · ${MESES_NOMBRE[mes]} ${anyo} · ${sorted.length} clientes · ${totalViajes} viajes · ${(totalKg/1000).toFixed(2)} Tn</div>`);
  win.document.write('</body></html>');
  win.document.close();
  setTimeout(() => win.print(), 300);
}

// ── EXPORTAR EXCEL ───────────────────────────────────────────
async function exportarExcelCliente(bcIdx) {
  const { cli } = window._bcClientesData[bcIdx];
  const fechaDesdeStr = document.getElementById('fact-fecha-desde').value;
  const fechaHastaStr = document.getElementById('fact-fecha-hasta').value;

  let fechaDesde = null, fechaHasta = null;
  if (fechaDesdeStr) {
    const [y, m, d] = fechaDesdeStr.split('-');
    fechaDesde = new Date(parseInt(y), parseInt(m) - 1, parseInt(d), 0, 0, 0);
  }
  if (fechaHastaStr) {
    const [y, m, d] = fechaHastaStr.split('-');
    fechaHasta = new Date(parseInt(y), parseInt(m) - 1, parseInt(d), 23, 59, 59);
  }

  const [anyo, mes] = fechaDesdeStr ? fechaDesdeStr.split('-').slice(0, 2).map(Number) : [new Date().getFullYear(), new Date().getMonth() + 1];
  const nomMes = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'][mes - 1];

  const viajes = factData.filter(r => {
    const d = parseFechaFact(r.fechaHora) || parseFechaFact(r.fechaPedido);
    if (!d) return false;
    if (fechaDesde && d < fechaDesde) return false;
    if (fechaHasta && d > fechaHasta) return false;
    return (r.nombreCliente || '').trim() === cli.trim();
  }).sort((a, b) => new Date(a.fechaHora) - new Date(b.fechaHora));

  const totalKg = viajes.reduce((s, r) => s + (Number(r.pesoNeto) || 0), 0);
  const totalEur = viajes.reduce((s, r) => {
    const tn = Number(r.pesoNeto) / 1000;
    return s + tn * getPrecioTn(cli, r.productoNombre || r.productoCod || '');
  }, 0);
  const igic = totalEur * IGIC_PCT / 100;

  const wb = new ExcelJS.Workbook();
  wb.creator = 'ARIFOMA';
  const ws = wb.addWorksheet(nomMes);

  // Anchos
  ws.columns = [
    {width:22},{width:13},{width:8},{width:15},{width:18},
    {width:24},{width:24},{width:13},{width:10},{width:9},{width:14}
  ];

  const VERDE  = '3D5A10';
  const VERDE2 = '6B7D2E';
  const GRIS   = 'F0F0EC';
  const GRIS2  = 'E0E0D8';
  const BLANCO = 'FFFFFF';

  const centerBold = (size, color='FFFFFF') => ({
    font: { bold:true, size, color:{argb:'FF'+color}, name:'Calibri' },
    alignment: { horizontal:'center', vertical:'middle' },
    fill: { type:'pattern', pattern:'solid', fgColor:{argb:'FF'+VERDE} }
  });

  // Fila 1 — ARIFOMA
  ws.mergeCells('A1:K1');
  const r1 = ws.getRow(1); r1.height = 36;
  Object.assign(ws.getCell('A1'), { value: 'ARIFOMA', ...centerBold(20) });

  // Fila 2 — título
  ws.mergeCells('A2:K2');
  const r2 = ws.getRow(2); r2.height = 22;
  Object.assign(ws.getCell('A2'), {
    value: `Desglose de viajes · ${nomMes} ${anyo}`,
    font: { bold:true, size:12, color:{argb:'FF'+BLANCO}, name:'Calibri' },
    alignment: { horizontal:'center', vertical:'middle' },
    fill: { type:'pattern', pattern:'solid', fgColor:{argb:'FF'+VERDE2} }
  });

  // Fila 3 — cliente
  ws.mergeCells('A3:K3');
  const r3 = ws.getRow(3); r3.height = 20;
  Object.assign(ws.getCell('A3'), {
    value: `Cliente: ${cli}`,
    font: { bold:true, size:11, color:{argb:'FF333333'}, name:'Calibri' },
    alignment: { horizontal:'center', vertical:'middle' },
    fill: { type:'pattern', pattern:'solid', fgColor:{argb:'FF'+GRIS} }
  });

  // Fila 4 — resumen
  ws.mergeCells('A4:K4');
  const r4 = ws.getRow(4); r4.height = 18;
  Object.assign(ws.getCell('A4'), {
    value: `Viajes: ${viajes.length}   ·   Tn: ${(totalKg/1000).toFixed(2)}   ·   Base: ${totalEur.toFixed(2)} €   ·   IGIC ${IGIC_PCT}%: ${igic.toFixed(2)} €   ·   Total: ${(totalEur+igic).toFixed(2)} €`,
    font: { size:10, color:{argb:'FF555555'}, name:'Calibri' },
    alignment: { horizontal:'center', vertical:'middle' },
    fill: { type:'pattern', pattern:'solid', fgColor:{argb:'FF'+GRIS2} }
  });

  // Fila 5 vacía
  ws.addRow([]);

  // Fila 6 — cabecera columnas
  const hdrRow = ws.addRow(['Fecha/Hora','Nº Pedido','Línea','Matrícula','Chofer','Proyecto','Producto','Kg Neto','Tn','€/Tn','Importe']);
  hdrRow.height = 18;
  hdrRow.eachCell(cell => {
    cell.font = { bold:true, size:10, color:{argb:'FF'+BLANCO}, name:'Calibri' };
    cell.fill = { type:'pattern', pattern:'solid', fgColor:{argb:'FF'+VERDE} };
    cell.alignment = { horizontal:'center', vertical:'middle' };
    cell.border = { bottom:{ style:'thin', color:{argb:'FFAAAAAA'} } };
  });

  // Filas datos
  viajes.forEach((r, i) => {
    const fecha = r.fechaHora ? new Date(r.fechaHora).toLocaleString('es-ES') : '';
    const tn = Number(r.pesoNeto) / 1000;
    const precioTn = getPrecioTn(cli, r.productoNombre || r.productoCod || '');
    const importe = tn * precioTn;
    const dataRow = ws.addRow([
      fecha, r.numPedido||'', r.numLinea||'', r.matriculacam||'', r.chofer||'',
      r.proyectoName||r.proyectoCod||'', r.productoNombre||r.productoCod||'',
      Number(r.pesoNeto)||0, parseFloat(tn.toFixed(3)),
      parseFloat(precioTn.toFixed(2)), parseFloat(importe.toFixed(2))
    ]);
    dataRow.height = 16;
    const bg = i % 2 === 0 ? 'FFFFFFFF' : 'FFF7F7F2';
    dataRow.eachCell(cell => {
      cell.fill = { type:'pattern', pattern:'solid', fgColor:{argb:bg} };
      cell.font = { size:9, name:'Calibri' };
      cell.alignment = { vertical:'middle' };
    });
    // Números alineados derecha
    [8,9,10,11].forEach(col => {
      dataRow.getCell(col).alignment = { horizontal:'right', vertical:'middle' };
    });
  });

  // Fila vacía
  ws.addRow([]);

  // Fila TOTAL
  const totRow = ws.addRow(['TOTAL','','','','','','', totalKg, parseFloat((totalKg/1000).toFixed(3)),'', parseFloat(totalEur.toFixed(2))]);
  totRow.height = 18;
  totRow.eachCell(cell => {
    cell.font = { bold:true, size:10, color:{argb:'FF'+BLANCO}, name:'Calibri' };
    cell.fill = { type:'pattern', pattern:'solid', fgColor:{argb:'FF'+VERDE2} };
    cell.alignment = { vertical:'middle' };
  });
  totRow.getCell(1).alignment = { horizontal:'center', vertical:'middle' };
  [8,9,11].forEach(col => totRow.getCell(col).alignment = { horizontal:'right', vertical:'middle' });

  // Descargar
  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], { type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `${cli} - ${nomMes} ${anyo}.xlsx`;
  a.click();
  URL.revokeObjectURL(url);
}

// ── CAJA — Facturas de compra desde BC → Google Sheet ────────
let cajaFacturas = [];
let cajaSelected = new Set();
let cajaRegistradas = new Set();
let cajaInited = false;

function initCaja() {
  if (!cajaInited) {
    const hoy = new Date();
    const primerDia = new Date(hoy.getFullYear(), hoy.getMonth(), 1);
    document.getElementById('caja-fecha-desde').value = primerDia.getFullYear() + '-' + pad(primerDia.getMonth() + 1) + '-' + pad(primerDia.getDate());
    document.getElementById('caja-fecha-hasta').value = hoy.getFullYear() + '-' + pad(hoy.getMonth() + 1) + '-' + pad(hoy.getDate());
    cajaInited = true;
  }
  if (cajaFacturas.length === 0) {
    // No auto-load, user clicks button
  } else {
    renderCajaList();
  }
}

async function cargarFacturasCompra() {
  const el = document.getElementById('caja-list');
  el.innerHTML = '<div style="color:var(--muted);text-align:center;padding:30px">Conectando con BC...</div>';
  cajaSelected.clear();
  updateCajaSelResumen();

  try {
    const token = await getBCToken();
    const headers = { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' };
    const base = `https://api.businesscentral.dynamics.com/v2.0/${BC_TENANT}/${BC_ENV}/api/v2.0/companies`;

    // Get company
    const cRes = await fetch(base, { headers });
    if (!cRes.ok) throw new Error('No se pudo obtener company');
    const cJson = await cRes.json();
    const company = cJson.value.find(c => c.name.trim() === BC_COMPANY.trim());
    if (!company) throw new Error('Company no encontrada');
    const companyId = company.id;

    // Get purchase invoices
    const invRes = await fetch(
      `${base}(${companyId})/purchaseInvoices?$select=id,number,postingDate,invoiceDate,vendorNumber,vendorName,totalAmountExcludingTax,totalAmountIncludingTax,vendorInvoiceNumber,status&$orderby=postingDate desc&$top=500`,
      { headers }
    );
    if (!invRes.ok) throw new Error('Error obteniendo facturas: ' + invRes.statusText);
    const invJson = await invRes.json();

    // Cargar facturas ya registradas en caja (Supabase + Google Sheet)
    const CAJA_SHEET_ID = '1fxHwVEgcIrRdyPh-TJ-k84QFBHXX-P3mNRCiWYaeDTQ';
    const CAJA_SHEET_TAB = 'LISTADO FACTS.CAJAS';
    const [cajaRes, sheetRes] = await Promise.all([
      dbQuery({ action: 'select', table: 'tblcaja', options: { select: 'facturabc' } }),
      fetch(`https://docs.google.com/spreadsheets/d/${CAJA_SHEET_ID}/gviz/tq?tqx=out:json&sheet=${encodeURIComponent(CAJA_SHEET_TAB)}&tq=${encodeURIComponent('select C')}`)
        .then(r => r.text()).catch(() => '')
    ]);
    const registradasSet = new Set((cajaRes.data || []).map(r => r.facturabc.trim()));
    // Añadir facturas del Google Sheet (columna C)
    if (sheetRes) {
      try {
        const jsonStr = sheetRes.replace(/^[^(]*\(/, '').replace(/\);?\s*$/, '');
        const gviz = JSON.parse(jsonStr);
        (gviz.table.rows || []).forEach(row => {
          const v = row.c && row.c[0] && row.c[0].v;
          if (v) registradasSet.add(String(v).trim());
        });
      } catch(_) {}
    }

    cajaRegistradas = registradasSet;
    cajaFacturas = (invJson.value || []).map(f => ({
      id: f.id,
      number: f.number || '',
      postingDate: f.postingDate || '',
      invoiceDate: f.invoiceDate || '',
      vendorNumber: f.vendorNumber || '',
      vendorName: f.vendorName || '',
      amount: f.totalAmountExcludingTax || 0,
      amountInc: f.totalAmountIncludingTax || 0,
      vendorInvoice: f.vendorInvoiceNumber || '',
      status: f.status || ''
    }));

    el.innerHTML = '';
    renderCajaList();
  } catch (e) {
    el.innerHTML = `<div style="color:var(--danger);text-align:center;padding:30px">Error: ${e.message}</div>`;
  }
}

function renderCajaList() {
  const el = document.getElementById('caja-list');
  if (!cajaFacturas.length) {
    el.innerHTML = '<div style="color:var(--muted);text-align:center;padding:30px;font-size:.82rem">Sin facturas cargadas</div>';
    return;
  }

  const fechaDesdeStr = document.getElementById('caja-fecha-desde').value;
  const fechaHastaStr = document.getElementById('caja-fecha-hasta').value;
  const buscar = (document.getElementById('caja-buscar').value || '').toLowerCase();

  let filtered = cajaFacturas.filter(f => {
    if (fechaDesdeStr && f.postingDate < fechaDesdeStr) return false;
    if (fechaHastaStr && f.postingDate > fechaHastaStr) return false;
    if (buscar && !f.vendorName.toLowerCase().includes(buscar) && !f.number.toLowerCase().includes(buscar) && !f.vendorInvoice.toLowerCase().includes(buscar)) return false;
    return true;
  });

  let html = `<div style="margin-bottom:8px;display:flex;justify-content:space-between;align-items:center">
    <div style="font-size:.72rem;color:var(--muted)">${filtered.length} factura${filtered.length !== 1 ? 's' : ''} · ${filtered.filter(f => cajaRegistradas.has(f.number.trim())).length} ya registrada${filtered.filter(f => cajaRegistradas.has(f.number.trim())).length !== 1 ? 's' : ''}</div>
    <button onclick="toggleAllCaja()" style="background:none;border:1px solid var(--border);border-radius:6px;padding:4px 10px;font-size:.68rem;color:var(--accent);cursor:pointer;font-weight:600">${cajaSelected.size > 0 && filtered.every(f => cajaSelected.has(f.id)) ? 'Deseleccionar todos' : 'Seleccionar todos'}</button>
  </div>`;

  // Header
  html += `<div style="display:flex;padding:6px 12px;font-size:.65rem;font-weight:700;color:var(--muted);text-transform:uppercase;gap:8px;border-bottom:1px solid var(--border)">
    <div style="width:28px"></div>
    <div style="flex:.8">Factura BC</div>
    <div style="flex:.6">F. Registro</div>
    <div style="flex:.6">F. Emisión</div>
    <div style="flex:2">Proveedor</div>
    <div style="flex:.8;text-align:right">Importe</div>
    <div style="flex:1">Nº Fact. Proveedor</div>
  </div>`;

  filtered.forEach(f => {
    const yaRegistrada = cajaRegistradas.has(f.number.trim());
    const checked = !yaRegistrada && cajaSelected.has(f.id);
    const dateReg = f.postingDate ? new Date(f.postingDate + 'T00:00:00').toLocaleDateString('es-ES') : '';
    const dateEmi = f.invoiceDate ? new Date(f.invoiceDate + 'T00:00:00').toLocaleDateString('es-ES') : '';
    if (yaRegistrada) {
      html += `<div style="display:flex;align-items:center;padding:10px 12px;gap:8px;border-bottom:1px solid rgba(0,0,0,.05);opacity:.45;pointer-events:none;background:rgba(0,0,0,.03)">
        <div style="width:22px;height:22px;border:2px solid var(--border);border-radius:6px;flex-shrink:0;display:flex;align-items:center;justify-content:center;background:var(--border)">
          <svg width="12" height="12" viewBox="0 0 12 12"><path d="M2 6l3 3 5-5" stroke="#fff" stroke-width="2" fill="none"/></svg>
        </div>
        <div style="flex:.8;font-family:'DM Mono',monospace;font-size:.75rem">${f.number}</div>
        <div style="flex:.6;font-size:.75rem">${dateReg}</div>
        <div style="flex:.6;font-size:.75rem">${dateEmi}</div>
        <div style="flex:2;font-size:.78rem;font-weight:600;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${f.vendorName}</div>
        <div style="flex:.8;text-align:right;font-family:'DM Mono',monospace;font-size:.78rem;font-weight:700">${f.amountInc.toLocaleString('es-ES',{minimumFractionDigits:2,maximumFractionDigits:2})} €</div>
        <div style="flex:1;font-size:.72rem;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${f.vendorInvoice} <span style="font-size:.6rem;font-style:italic">✓ registrada</span></div>
      </div>`;
    } else {
      html += `<div onclick="toggleCajaItem('${f.id}')" style="display:flex;align-items:center;padding:10px 12px;gap:8px;border-bottom:1px solid rgba(0,0,0,.05);cursor:pointer;background:${checked ? 'rgba(107,125,46,.06)' : 'transparent'};transition:background .15s">
        <div style="width:22px;height:22px;border:2px solid ${checked ? 'var(--accent2)' : 'var(--border)'};border-radius:6px;flex-shrink:0;display:flex;align-items:center;justify-content:center;background:${checked ? 'var(--accent2)' : 'transparent'};transition:all .15s">
          ${checked ? '<svg width="12" height="12" viewBox="0 0 12 12"><path d="M2 6l3 3 5-5" stroke="#fff" stroke-width="2" fill="none"/></svg>' : ''}
        </div>
        <div style="flex:.8;font-family:'DM Mono',monospace;font-size:.75rem;color:var(--text)">${f.number}</div>
        <div style="flex:.6;font-size:.75rem;color:var(--muted)">${dateReg}</div>
        <div style="flex:.6;font-size:.75rem;color:var(--muted)">${dateEmi}</div>
        <div style="flex:2;font-size:.78rem;color:var(--text);font-weight:600;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${f.vendorName}</div>
        <div style="flex:.8;text-align:right;font-family:'DM Mono',monospace;font-size:.78rem;font-weight:700;color:var(--accent2)">${f.amountInc.toLocaleString('es-ES',{minimumFractionDigits:2,maximumFractionDigits:2})} €</div>
        <div style="flex:1;font-size:.72rem;color:var(--muted);overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${f.vendorInvoice}</div>
      </div>`;
    }
  });

  el.innerHTML = html;
  updateCajaSelResumen();
}

function toggleCajaItem(id) {
  if (cajaSelected.has(id)) cajaSelected.delete(id);
  else cajaSelected.add(id);
  renderCajaList();
}

function toggleAllCaja() {
  const fechaDesdeStr = document.getElementById('caja-fecha-desde').value;
  const fechaHastaStr = document.getElementById('caja-fecha-hasta').value;
  const buscar = (document.getElementById('caja-buscar').value || '').toLowerCase();
  const filtered = cajaFacturas.filter(f => {
    if (fechaDesdeStr && f.postingDate < fechaDesdeStr) return false;
    if (fechaHastaStr && f.postingDate > fechaHastaStr) return false;
    if (buscar && !f.vendorName.toLowerCase().includes(buscar) && !f.number.toLowerCase().includes(buscar) && !f.vendorInvoice.toLowerCase().includes(buscar)) return false;
    return true;
  });
  const seleccionables = filtered.filter(f => !cajaRegistradas.has(f.number.trim()));
  const allSelected = seleccionables.every(f => cajaSelected.has(f.id));
  if (allSelected) seleccionables.forEach(f => cajaSelected.delete(f.id));
  else seleccionables.forEach(f => cajaSelected.add(f.id));
  renderCajaList();
}

function updateCajaSelResumen() {
  const wrap = document.getElementById('caja-sel-resumen');
  const n = cajaSelected.size;
  if (n === 0) { wrap.style.display = 'none'; return; }
  wrap.style.display = 'flex';
  let total = 0;
  cajaSelected.forEach(id => { const f = cajaFacturas.find(x => x.id === id); if (f) total += f.amountInc; });
  document.getElementById('caja-sel-text').textContent = `${n} factura${n !== 1 ? 's' : ''} seleccionada${n !== 1 ? 's' : ''} · ${total.toLocaleString('es-ES',{minimumFractionDigits:2,maximumFractionDigits:2})} €`;
}

async function enviarCajaSheet() {
  if (!cajaSelected.size) return;
  const btn = document.getElementById('caja-btn-enviar');
  btn.disabled = true;
  btn.textContent = 'Enviando...';

  try {
    const now = new Date();
    const anyo2 = String(now.getFullYear()).slice(-2);
    const mes2 = pad(now.getMonth() + 1);
    const codCaja = anyo2 + mes2;
    const meses = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
    const numCaja = 'Caja' + meses[now.getMonth()] + anyo2;

    const filas = [];
    cajaSelected.forEach(id => {
      const f = cajaFacturas.find(x => x.id === id);
      if (!f) return;
      filas.push({
        codCaja,
        numCaja,
        facturaBC: f.number,
        fechaFactura: f.invoiceDate ? new Date(f.invoiceDate + 'T00:00:00').toLocaleDateString('es-ES') : '',
        fechaRegistro: f.postingDate ? new Date(f.postingDate + 'T00:00:00').toLocaleDateString('es-ES') : '',
        proveedor: f.vendorName,
        importe: f.amountInc,
        factProveedor: f.vendorInvoice
      });
    });

    const rows = filas.map(f => ({
      codcaja: f.codCaja,
      numcaja: f.numCaja,
      facturabc: f.facturaBC,
      fechafactura: f.fechaFactura,
      fecharegistro: f.fechaRegistro,
      proveedor: f.proveedor,
      importe: f.importe,
      factproveedor: f.factProveedor
    }));
    const cajaInsert = await dbQuery({ action: 'insert', table: 'tblcaja', data: rows });
    if (!cajaInsert.ok) throw new Error(cajaInsert.error);

    btn.textContent = '✓ Enviado';
    btn.style.background = '#2e7d32';
    alert(`${filas.length} factura${filas.length > 1 ? 's' : ''} registrada${filas.length > 1 ? 's' : ''} en caja (${numCaja})`);

    // Limpiar selección
    cajaSelected.clear();
    setTimeout(() => {
      btn.disabled = false;
      btn.textContent = 'Registrar en Sheet';
      btn.style.background = '';
      renderCajaList();
    }, 1500);
  } catch (e) {
    btn.disabled = false;
    btn.textContent = 'Registrar en Sheet';
    btn.style.background = '';
    alert('Error: ' + e.message);
  }
}

// ── BUSINESS CENTRAL ─────────────────────────────────────────
const BC_TENANT   = '5bd828f2-1899-48ba-a269-c37733f41806';
const BC_CLIENT   = 'e2a57ff0-8ea7-433d-a2af-7335d3f01847';
const BC_ENV      = 'Production';
const BC_COMPANY  = 'ARIFOMA 25P.V06';
// Resetear cache companyId y clientes BC al cargar
window._bcCompanyId = null;
window._bcClientesData = [];
const BC_SCOPE    = `https://api.businesscentral.dynamics.com/.default`;

let msalApp = null;
async function getMsalApp() {
  if (msalApp) return msalApp;
  msalApp = new msal.PublicClientApplication({
    auth: {
      clientId: BC_CLIENT,
      authority: `https://login.microsoftonline.com/${BC_TENANT}`,
      redirectUri: window.location.origin + '/'
    },
    cache: { cacheLocation: 'sessionStorage' }
  });
  await msalApp.initialize();
  return msalApp;
}

// Limpia flags de interacción MSAL colgados en storage
function _clearMsalInteractionState() {
  const scanAndClear = (storage) => {
    const keys = [];
    for (let i = 0; i < storage.length; i++) keys.push(storage.key(i));
    keys.filter(k => k && k.includes('interaction.status')).forEach(k => storage.removeItem(k));
  };
  scanAndClear(sessionStorage);
  scanAndClear(localStorage);
}

let _bcTokenPromise = null;
async function getBCToken() {
  if(_bcTokenPromise) return _bcTokenPromise;
  _bcTokenPromise = _getBCTokenInner().finally(()=>{ _bcTokenPromise=null; });
  return _bcTokenPromise;
}
async function _getBCTokenInner(allowPopup=true) {
  const app = await getMsalApp();
  const req = { scopes: [BC_SCOPE] };
  try {
    const accounts = app.getAllAccounts();
    if (accounts.length > 0) {
      return (await app.acquireTokenSilent({ ...req, account: accounts[0] })).accessToken;
    }
  } catch(e) {}
  if (!allowPopup) throw new Error('No hay sesión BC activa');
  _clearMsalInteractionState();
  return (await app.acquireTokenPopup(req)).accessToken;
}

// Para llamadas en background (arranque): no abrir popup, solo silent
async function getBCTokenSilent() {
  const app = await getMsalApp();
  const accounts = app.getAllAccounts();
  if (!accounts.length) throw new Error('No hay sesión BC activa');
  const req = { scopes: [BC_SCOPE] };
  return (await app.acquireTokenSilent({ ...req, account: accounts[0] })).accessToken;
}

async function enviarBCCliente(bcIdx, btn) {
  const { cli, cData } = window._bcClientesData[bcIdx];
  btn.disabled = true;
  btn.textContent = 'Conectando...';
  try {
    const token = await getBCToken();

    // Buscar codigoCliente en factData
    const pedidoCli = factData?.find(r => (r.nombreCliente||'').trim() === cli.trim());
    const customerNo = pedidoCli?.codigoCliente || '';

    // Obtener fecha del filtro activo (usa hoy si no hay filtro)
    const now = new Date();
    const mesEl = document.getElementById('fact-mes');
    const anyoEl = document.getElementById('fact-anyo');
    const mes = (mesEl?.value ? parseInt(mesEl.value) : now.getMonth()) + 1;
    const anyo = anyoEl?.value ? parseInt(anyoEl.value) : now.getFullYear();
    const invoiceDate = `${anyo}-${String(mes).padStart(2,'0')}-01`;
    const extDoc = `APP-${cli.substring(0,10).replace(/\s/g,'-')}-${anyo}${String(mes).padStart(2,'0')}`;

    // Comprobar si ya existe factura para este cliente/mes en BC
    btn.textContent = 'Comprobando...';
    const chkHeaders = { 'Authorization': `Bearer ${token}` };
    const chkBase = `https://api.businesscentral.dynamics.com/v2.0/${BC_TENANT}/${BC_ENV}/api/v2.0/companies`;
    const chkCRes = await fetch(chkBase, { headers: chkHeaders });
    const chkCJson = await chkCRes.json();
    const chkCompany = chkCJson.value.find(c => c.name.trim() === BC_COMPANY.trim());
    if (chkCompany) {
      const chkUrl = `${chkBase}(${chkCompany.id})/salesInvoices?$filter=externalDocumentNumber eq '${extDoc}'&$select=id,number,externalDocumentNumber&$top=1`;
      const chkRes = await fetch(chkUrl, { headers: chkHeaders });
      if (chkRes.ok) {
        const chkData = await chkRes.json();
        if (chkData.value && chkData.value.length > 0) {
          const existing = chkData.value[0];
          if (!confirm(`⚠️ Ya existe una factura para "${cli}" en ${String(mes).padStart(2,'0')}/${anyo}:\n\nNº: ${existing.number || existing.id}\nRef: ${existing.externalDocumentNumber}\n\n¿Crear otra igualmente?`)) {
            btn.disabled = false;
            btn.textContent = 'Enviar a BC';
            return;
          }
        }
      }
    }

    btn.textContent = 'Creando factura...';

    // Crear factura vía backend
    const invRes = await fetch('/api/bc/facturas', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        token,
        customerNumber: customerNo,
        invoiceDate,
        externalDocumentNumber: extDoc
      })
    });

    if (!invRes.ok) throw new Error(await invRes.text());
    const invData = await invRes.json();
    if (!invData.ok) throw new Error(invData.error);

    const inv = invData.invoice;
    const invId = inv.id;
    let lineCount = 0;

    // Crear líneas por proyecto → producto (directamente en frontend por ahora)
    const headers = { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' };
    const base = `https://api.businesscentral.dynamics.com/v2.0/${BC_TENANT}/${BC_ENV}/api/v2.0/companies`;
    const cRes = await fetch(base, { headers });
    const cJson = await cRes.json();
    const company = cJson.value.find(c => c.name.trim() === BC_COMPANY.trim());
    const companyId = company.id;

    for (const [proy, pData] of Object.entries(cData.proyectos)) {
      for (const [prod, info] of Object.entries(pData.productos)) {
        const lineRes = await fetch(`${base}(${companyId})/salesInvoices(${invId})/salesInvoiceLines`, {
          method: 'POST', headers,
          body: JSON.stringify({
            lineType: 'Item',
            lineObjectNumber: info.cod,
            description: `${prod} - ${proy}`,
            quantity: parseFloat((info.kg / 1000).toFixed(3)),
            unitPrice: info.precioTn
          })
        });
        console.log('Enviando línea:', prod, '→ cod:', info.cod, '→ PROD-'+String(info.cod).padStart(6,'0'));
        if (lineRes.ok) {
          const lineData = await lineRes.json();
          const lineId = lineData.id;
          const etag = lineData['@odata.etag'] || '*';
          console.log('Línea creada OK, campos:', Object.keys(lineData).join(', '));
          // Asignar proyecto si hay proyectoCod
          const proyCod = pData.proyectoCod || '';
          if (proyCod) {
            try {
              const patchUrl = `${base}(${companyId})/salesInvoices(${invId})/salesInvoiceLines(${lineId})`;
              const patchRes = await fetch(patchUrl, {
                method: 'PATCH',
                headers: { ...headers, 'If-Match': etag },
                body: JSON.stringify({ jobNo: proyCod, jobTaskNo: 'INGRESOS' })
              });
              if (!patchRes.ok) console.warn('PATCH proyecto error (jobNo='+proyCod+'):', await patchRes.text());
              else console.log('Proyecto asignado OK:', proyCod, 'INGRESOS');
            } catch(e) { console.warn('Error asignando proyecto:', e.message); }
          }
          lineCount++;
        }
        else console.error('Línea fallida:', prod, await lineRes.text());
      }
    }

    btn.textContent = '✓ Enviado';
    btn.style.background = '#2e7d32';
    alert(`Factura creada en BC para ${cli}\n${lineCount} líneas añadidas.\nNº: ${inv.number||invId}`);
  } catch(e) {
    btn.disabled = false;
    btn.textContent = 'Enviar a BC';
    alert('Error al enviar a BC:\n' + e.message);
  }
}

// ── FACTURAR ALBARANES (seleccionar albaranes locales → crear salesInvoice en BC) ──
window._albModalData = { customerNo: '', customerName: '', albaranes: [], selected: new Set() };

function abrirModalAlbaranes(cli) {
  const modal = document.getElementById('modal-albaranes');
  modal.style.display = 'flex';
  document.getElementById('modal-alb-cliente').textContent = cli;
  document.getElementById('modal-alb-btn-facturar').disabled = true;
  document.getElementById('modal-alb-btn-facturar').textContent = 'Crear factura en BC';
  document.getElementById('modal-alb-btn-facturar').style.background = '';
  document.getElementById('modal-alb-sel').textContent = '0 albaranes seleccionados';

  const pedidoCli = factData?.find(r => (r.nombreCliente || '').trim() === cli.trim());
  const customerNo = pedidoCli?.codigoCliente || '';

  // Filtrar pedidos de este cliente según fechas activas
  const fechaDesdeStr = document.getElementById('fact-fecha-desde').value;
  const fechaHastaStr = document.getElementById('fact-fecha-hasta').value;
  let fechaDesde = null, fechaHasta = null;
  if (fechaDesdeStr) { const [y, m, d] = fechaDesdeStr.split('-'); fechaDesde = new Date(parseInt(y), parseInt(m) - 1, parseInt(d), 0, 0, 0); }
  if (fechaHastaStr) { const [y, m, d] = fechaHastaStr.split('-'); fechaHasta = new Date(parseInt(y), parseInt(m) - 1, parseInt(d), 23, 59, 59); }

  const pedidosCli = factData.filter(r => {
    const d = parseFechaFact(r.fechaHora) || parseFechaFact(r.fechaPedido);
    if (!d) return false;
    if (fechaDesde && d < fechaDesde) return false;
    if (fechaHasta && d > fechaHasta) return false;
    return (r.nombreCliente || '').trim() === cli.trim();
  });

  // Agrupar por numPedido (cada numPedido = 1 albarán)
  const groups = {};
  pedidosCli.forEach(r => {
    const key = r.numPedido || r.numalbarancalle?.split('/')[0] || r.id || 'sin-num';
    if (!groups[key]) groups[key] = { numPedido: key, lineas: [], totalKg: 0, totalEur: 0, fecha: null, proyecto: '' };
    groups[key].lineas.push(r);
    const neto = Number(r.pesoNeto) || 0;
    groups[key].totalKg += neto;
    const precioTn = getPrecioTn(cli, r.productoNombre || r.productoCod || '');
    groups[key].totalEur += (neto / 1000) * precioTn;
    if (!groups[key].fecha) groups[key].fecha = parseFechaFact(r.fechaHora) || parseFechaFact(r.fechaPedido);
    if (!groups[key].proyecto) groups[key].proyecto = r.proyectoName || r.proyectoCod || '';
  });

  const albaranes = Object.values(groups).sort((a, b) => (b.fecha || 0) - (a.fecha || 0));

  window._albModalData = { customerNo, customerName: cli, albaranes, selected: new Set() };
  renderModalAlbaranes();
}

function renderModalAlbaranes() {
  const { albaranes, selected } = window._albModalData;
  const body = document.getElementById('modal-alb-body');

  if (!albaranes.length) {
    body.innerHTML = '<div style="color:var(--muted);text-align:center;padding:30px">No hay albaranes para este cliente en el periodo seleccionado</div>';
    return;
  }

  let html = `<div style="margin-bottom:10px;display:flex;justify-content:space-between;align-items:center">
    <div style="font-size:.72rem;color:var(--muted)">${albaranes.length} albarán${albaranes.length > 1 ? 'es' : ''}</div>
    <button onclick="toggleAllAlbaranes()" style="background:none;border:1px solid var(--border);border-radius:6px;padding:4px 10px;font-size:.68rem;color:var(--accent);cursor:pointer;font-weight:600">${selected.size === albaranes.length ? 'Deseleccionar todos' : 'Seleccionar todos'}</button>
  </div>`;

  albaranes.forEach((alb, i) => {
    const checked = selected.has(i);
    const dateStr = alb.fecha ? alb.fecha.toLocaleDateString('es-ES') : '';
    const prodResumen = {};
    alb.lineas.forEach(r => {
      const p = r.productoNombre || r.productoCod || '?';
      if (!prodResumen[p]) prodResumen[p] = 0;
      prodResumen[p] += Number(r.pesoNeto) || 0;
    });
    const lineSummary = Object.entries(prodResumen).map(([p, kg]) => `${p} ${(kg / 1000).toFixed(2)}Tn`).join(', ');
    const albId = alb.lineas[0]?.id;
    const albAnyo = alb.fecha ? alb.fecha.getFullYear() : new Date().getFullYear();
    const albNum = albId ? `PEDV${albAnyo}-${String(albId).padStart(6,'0')}` : alb.numPedido;

    html += `<div onclick="toggleAlbaran(${i})" style="display:flex;gap:10px;padding:12px;margin-bottom:8px;background:var(--surface);border:1.5px solid ${checked ? 'var(--accent2)' : 'var(--border)'};border-radius:var(--radius);cursor:pointer;transition:all .15s">
      <div style="width:22px;height:22px;border:2px solid ${checked ? 'var(--accent2)' : 'var(--border)'};border-radius:6px;flex-shrink:0;display:flex;align-items:center;justify-content:center;background:${checked ? 'var(--accent2)' : 'transparent'};transition:all .15s;margin-top:2px">
        ${checked ? '<svg width="12" height="12" viewBox="0 0 12 12"><path d="M2 6l3 3 5-5" stroke="#fff" stroke-width="2" fill="none"/></svg>' : ''}
      </div>
      <div style="flex:1;min-width:0">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:4px">
          <div style="font-weight:700;font-size:.82rem;color:var(--text)">Albarán ${albNum}</div>
          <div style="font-family:'DM Mono',monospace;font-size:.82rem;font-weight:700;color:var(--accent2)">${alb.totalEur.toFixed(2)} €</div>
        </div>
        <div style="font-size:.68rem;color:var(--muted);margin-bottom:2px">${dateStr}${alb.proyecto ? ' · ' + alb.proyecto : ''} · ${(alb.totalKg / 1000).toFixed(2)} Tn</div>
        <div style="font-size:.68rem;color:var(--muted);overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${alb.lineas.length} línea${alb.lineas.length !== 1 ? 's' : ''}: ${lineSummary}</div>
      </div>
    </div>`;
  });

  // Total seleccionado
  let totalSel = 0, totalKgSel = 0;
  selected.forEach(i => { totalSel += albaranes[i].totalEur; totalKgSel += albaranes[i].totalKg; });
  if (selected.size > 0) {
    html += `<div style="margin-top:8px;padding:10px 12px;background:rgba(107,125,46,.08);border:1px solid rgba(107,125,46,.3);border-radius:var(--radius);display:flex;justify-content:space-between;align-items:center">
      <div style="font-size:.75rem;font-weight:600;color:var(--accent2)">Total seleccionado</div>
      <div style="font-family:'DM Mono',monospace;font-size:.88rem;font-weight:700;color:var(--accent2)">${(totalKgSel / 1000).toFixed(2)} Tn · ${totalSel.toFixed(2)} €</div>
    </div>`;
  }

  body.innerHTML = html;
  updateAlbSelCount();
}

function toggleAlbaran(idx) {
  const { selected } = window._albModalData;
  if (selected.has(idx)) selected.delete(idx);
  else selected.add(idx);
  renderModalAlbaranes();
}

function toggleAllAlbaranes() {
  const { albaranes, selected } = window._albModalData;
  if (selected.size === albaranes.length) selected.clear();
  else albaranes.forEach((_, i) => selected.add(i));
  renderModalAlbaranes();
}

function updateAlbSelCount() {
  const n = window._albModalData.selected.size;
  document.getElementById('modal-alb-sel').textContent = `${n} albarán${n !== 1 ? 'es' : ''} seleccionado${n !== 1 ? 's' : ''}`;
  document.getElementById('modal-alb-btn-facturar').disabled = n === 0;
}

function cerrarModalAlbaranes() {
  document.getElementById('modal-albaranes').style.display = 'none';
}

async function facturarAlbaranesSeleccionados() {
  const { customerNo, customerName, albaranes, selected } = window._albModalData;
  if (!selected.size) return;

  const btn = document.getElementById('modal-alb-btn-facturar');
  btn.disabled = true;
  btn.textContent = 'Creando factura...';

  try {
    const token = await getBCToken();
    const now = new Date();
    const invoiceDate = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;
    const selAlbs = [...selected].map(i => albaranes[i]);
    const numAlbaranes = selAlbs.map(a => {
      const linea = a.lineas[0];
      const id = linea?.id;
      const anyo = a.fecha ? a.fecha.getFullYear() : now.getFullYear();
      return id ? `PEDV${anyo}-${String(id).padStart(6,'0')}` : (a.numPedido || '');
    }).filter(Boolean);
    const extDocRaw = numAlbaranes.join(', ');
    const extDoc = extDocRaw.length <= 35 ? extDocRaw : extDocRaw.substring(0, 32) + '...';

    // Comprobar si ya existe factura para este cliente/mes en BC
    btn.textContent = 'Comprobando...';
    const chkHeaders = { 'Authorization': `Bearer ${token}` };
    const chkBase = `https://api.businesscentral.dynamics.com/v2.0/${BC_TENANT}/${BC_ENV}/api/v2.0/companies`;
    const chkCRes = await fetch(chkBase, { headers: chkHeaders });
    const chkCJson = await chkCRes.json();
    const chkCompany = chkCJson.value.find(c => c.name.trim() === BC_COMPANY.trim());
    if (chkCompany) {
      const mesStr = String(now.getMonth() + 1).padStart(2, '0');
      const chkUrl = `${chkBase}(${chkCompany.id})/salesInvoices?$filter=externalDocumentNumber eq '${extDoc}'&$select=id,number,externalDocumentNumber&$top=1`;
      const chkRes = await fetch(chkUrl, { headers: chkHeaders });
      if (chkRes.ok) {
        const chkData = await chkRes.json();
        if (chkData.value && chkData.value.length > 0) {
          const existing = chkData.value[0];
          if (!confirm(`⚠️ Ya existe una factura para "${customerName}" en ${mesStr}/${now.getFullYear()}:\n\nNº: ${existing.number || existing.id}\nRef: ${existing.externalDocumentNumber}\n\n¿Crear otra igualmente?`)) {
            btn.disabled = false;
            btn.textContent = 'Crear factura en BC';
            return;
          }
        }
      }
    }

    // Crear factura via backend
    const invRes = await fetch('/api/bc/facturas', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ token, customerNumber: customerNo, invoiceDate, externalDocumentNumber: extDoc })
    });
    if (!invRes.ok) throw new Error(await invRes.text());
    const invData = await invRes.json();
    if (!invData.ok) throw new Error(invData.error);

    const inv = invData.invoice;
    const invId = inv.id;

    // Crear líneas en la factura agrupando por proyecto → producto
    const headers2 = { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' };
    const base = `https://api.businesscentral.dynamics.com/v2.0/${BC_TENANT}/${BC_ENV}/api/v2.0/companies`;
    const cRes = await fetch(base, { headers: headers2 });
    const cJson = await cRes.json();
    const company = cJson.value.find(c => c.name.trim() === BC_COMPANY.trim());
    const companyId = company.id;

    // Agrupar líneas seleccionadas por producto
    const prodLines = {};
    [...selected].forEach(i => {
      const alb = albaranes[i];
      alb.lineas.forEach(r => {
        const prod = r.productoNombre || r.productoCod || 'Sin producto';
        const cod = r.productoCod || '';
        const proy = r.proyectoName || r.proyectoCod || '';
        const key = `${cod}||${proy}`;
        const proyCod = r.proyectoCod || '';
        if (!prodLines[key]) prodLines[key] = { cod, prod, proy, proyCod, kg: 0, viajes: 0, albNums: new Set() };
        prodLines[key].kg += Number(r.pesoNeto) || 0;
        prodLines[key].viajes++;
        prodLines[key].albNums.add(alb.numPedido);
      });
    });

    let lineCount = 0;
    for (const [, info] of Object.entries(prodLines)) {
      const tn = info.kg / 1000;
      const precioTn = getPrecioTn(customerName, info.prod);
      const albRef = [...info.albNums].join(',');
      const lineRes = await fetch(`${base}(${companyId})/salesInvoices(${invId})/salesInvoiceLines`, {
        method: 'POST',
        headers: headers2,
        body: JSON.stringify({
          lineType: 'Item',
          lineObjectNumber: info.cod,
          description: `${info.prod} - ${info.proy} [Alb.${albRef}]`,
          quantity: parseFloat(tn.toFixed(3)),
          unitPrice: precioTn
        })
      });
      if (lineRes.ok) {
        const lineData = await lineRes.json();
        const lineId = lineData.id;
        const etag = lineData['@odata.etag'] || '*';
        // Asignar proyecto si hay proyectoCod
        if (info.proyCod) {
          try {
            const patchUrl = `${base}(${companyId})/salesInvoices(${invId})/salesInvoiceLines(${lineId})`;
            const patchRes = await fetch(patchUrl, {
              method: 'PATCH',
              headers: { ...headers2, 'If-Match': etag },
              body: JSON.stringify({ jobNo: info.proyCod, jobTaskNo: 'INGRESOS' })
            });
            if (!patchRes.ok) console.warn('PATCH proyecto error:', await patchRes.text());
            else console.log('Proyecto asignado OK:', info.proyCod, 'INGRESOS');
          } catch(e) { console.warn('Error asignando proyecto:', e.message); }
        }
        lineCount++;
      }
      else console.error('Línea fallida:', info.prod, await lineRes.text());
    }

    btn.textContent = '✓ Factura creada';
    btn.style.background = '#2e7d32';
    alert(`Factura creada en BC para ${customerName}\nDesde ${selected.size} albarán${selected.size > 1 ? 'es' : ''}\n${lineCount} líneas añadidas.\nNº: ${inv.number || invId}`);

    setTimeout(() => cerrarModalAlbaranes(), 1500);
  } catch (e) {
    btn.disabled = false;
    btn.textContent = 'Crear factura en BC';
    btn.style.background = '';
    alert('Error al facturar:\n' + e.message);
  }
}

// ── OT PRINT ─────────────────────────────────────────────────
function printOT(filled){
  if(!selMachine||!selGama)return;
  const horas=document.getElementById('inputHoras').value||'';
  const obs=document.getElementById('inputObs')?document.getElementById('inputObs').value:'';
  const fechaRaw=document.getElementById('inputFecha').value;
  const fecha=fechaRaw?new Date(fechaRaw).toLocaleString('es-ES',{day:'2-digit',month:'2-digit',year:'numeric',hour:'2-digit',minute:'2-digit'}):'';
  const tipos={preventivo:'Preventivo programado',correctivo:'Correctivo',revision:'Revisión rápida'};
  const tipo=tipos[document.getElementById('inputTipo').value]||'Preventivo';
  document.getElementById('otp-ref').textContent='OT — '+new Date().toLocaleDateString('es-ES');
  document.getElementById('otp-maquina').textContent=selMachine.name+' ('+selMachine.id+')';
  document.getElementById('otp-fecha').textContent=fecha;
  document.getElementById('otp-activo').textContent=selMachine.id;
  document.getElementById('otp-gama').textContent=selGama.nombre+' ('+selGama.intervalo+'h)';
  document.getElementById('otp-horas').textContent=horas?horas+' h':'';
  document.getElementById('otp-tipo').textContent=tipo;
  document.getElementById('otp-operario').textContent='';
  document.getElementById('otp-fabricante').textContent=selMachine.fabricante;
  document.getElementById('otp-obs').textContent=filled?obs:'';
  // Build checks table
  const tbl=document.getElementById('otp-checks-table');
  let rows='';
  selGama.checks.forEach((c,i)=>{
    const done=filled&&checkStates[i];
    rows+=`<tr style="background:${i%2===0?'#f9f9f9':'#fff'}">
      <td style="border:1px solid #ddd;padding:3px 6px;width:60%;font-size:7.5pt">${c}</td>
      <td style="border:1px solid #ddd;padding:3px 6px;width:20%;text-align:center;font-size:7.5pt;color:${done?'#2e7d32':'#999'}">${filled?(done?'✓ OK':'—'):'□'}</td>
      <td style="border:1px solid #ddd;padding:3px 6px;width:20%;font-size:7pt;color:#999">Obs.</td>
    </tr>`;
  });
  tbl.innerHTML=`<thead><tr style="background:#444;color:#fff"><th style="padding:3px 6px;border:1px solid #444;text-align:left;width:60%">Punto de verificación</th><th style="padding:3px 6px;border:1px solid #444;width:20%">Estado</th><th style="padding:3px 6px;border:1px solid #444;width:20%">Observación</th></tr></thead><tbody>${rows}</tbody>`;
  const wrap=document.getElementById('ot-print-wrap');
  wrap.style.display='flex';
  wrap.style.position='fixed';
  wrap.classList.add('print-active');
  setTimeout(()=>window.print(),100);
}
function cerrarOTPrint(){
  const wrap=document.getElementById('ot-print-wrap');
  wrap.style.display='none';
  wrap.classList.remove('print-active');
}

// ── MANTENIMIENTO PREVENTIVO ──────────────────────────────────
let prevData=[];
let prevGasoilHoroMap={};   // {machine.id → horómetro} desde gasoil sheet col C
let prevGasoilFechaMap={}; // {machine.id → fecha último registro gasoil}
// Mapa destino gasoil (como aparece en el sheet) → machine.id
const GASOIL_DEST_MAP={
  '966G':'M966G.01','349':'M349.1','725C':'M725.1','MERLO':'M40.9.1',
  '336':'M336.1','769':'M769C.01','365B':'M365B.01','C32':'MC32.1',
  'DE22':'DE22','KANGOO':'8590FBV','3833BNX':'3833BNX','PRAMAC':'PRAMAC'
};
let customGamas=JSON.parse(localStorage.getItem('customGamas')||'[]');

function getEffectiveGamas(){
  // Merge hardcoded GAMAS with custom ones (custom overrides by id)
  const base=[...GAMAS];
  customGamas.forEach(cg=>{
    const idx=base.findIndex(g=>g.id===cg.id);
    if(idx>=0)base[idx]=cg;else base.push(cg);
  });
  return base;
}

async function cargarMantenimientoPreventivo(){
  const el=document.getElementById('prev-list');
  el.innerHTML='<div class="tbl"><div class="empty">Cargando...</div></div>';
  try{
    // Cargar listado preventivo (fuente principal del listado)
    const lRes=await apiFetch('?accion=gamasListado');
    if(lRes.ok)listadoPrevData=(lRes.data||[]).filter(r=>r.id!=null);
    // Fetch OT history + gasoil in parallel
    const [jsonOT, jsonGasoil]=await Promise.all([
      apiFetch('?accion=historialOT'),
      apiFetch('?accion=gasoil').catch(()=>({ok:false}))
    ]);
    if(!jsonOT.ok)throw new Error(jsonOT.error);
    prevData=jsonOT.data;
    // Build horómetro + fecha maps desde gasoil consumos (col C = MAX = horómetro actual)
    prevGasoilHoroMap={};
    prevGasoilFechaMap={};
    if(jsonGasoil.ok&&jsonGasoil.consumos){
      jsonGasoil.consumos.forEach(c=>{
        if(!c.activo)return;
        const horo=Number(c.max||0);
        if(horo<=0)return;
        prevGasoilHoroMap[c.activo]=horo;
        if(c.actualizado)prevGasoilFechaMap[c.activo]=c.actualizado;
        const machineId=GASOIL_DEST_MAP[c.activo]||null;
        if(machineId){prevGasoilHoroMap[machineId]=horo;if(c.actualizado)prevGasoilFechaMap[machineId]=c.actualizado;}
      });
    }
    // Última fecha por máquina desde historial gasoil
    if(jsonGasoil.ok&&jsonGasoil.data){
      jsonGasoil.data.forEach(r=>{
        const dest=r.destino||r.maquina;
        if(!dest)return;
        const machineId=GASOIL_DEST_MAP[dest]||null;
        if(!machineId)return;
        if(!prevGasoilFechaMap[machineId]||r.fecha>prevGasoilFechaMap[machineId])
          prevGasoilFechaMap[machineId]=r.fecha;
      });
    }
    // Populate machine filter desde listadoPrevData
    const maquinas=[...new Set(listadoPrevData.map(r=>r.Activo).filter(Boolean))].sort();
    const sel=document.getElementById('filt-prev-maquina');
    sel.innerHTML='<option value="">Todas las máquinas</option>';
    maquinas.forEach(m=>{const o=document.createElement('option');o.value=m;o.textContent=m;sel.appendChild(o);});
    renderMantenimientoPreventivo();
    notificarMantenimiento();
  }catch(e){el.innerHTML='<div class="tbl"><div class="empty">Error: '+e.message+'</div></div>';}
}

function calcMantenimiento(){
  const result=[];
  const HORAS_DIA=8;
  // Horómetro actual por activo desde OTs
  const machineHoroOT={};
  prevData.forEach(r=>{
    const m=Number(r.medicion)||0;
    if(!machineHoroOT[r.activo]||m>machineHoroOT[r.activo])machineHoroOT[r.activo]=m;
  });
  // Usar listadoPrevData (tblGamasListadoPreventivo) como fuente
  listadoPrevData.forEach(r=>{
    if(!r.Activo||!r.Gama)return;
    const gasoilHoro=prevGasoilHoroMap[r.Activo]||null;
    const gasoilFecha=prevGasoilFechaMap&&prevGasoilFechaMap[r.Activo]||null;
    const currentHoro=gasoilHoro||machineHoroOT[r.Activo]||Number(r.U_Medicion_med)||0;
    const proximo=Number(r.Proximo)||null;
    const falta=proximo&&currentHoro?proximo-currentHoro:null;
    const diasRestantes=falta!==null?Math.round(falta/HORAS_DIA):null;
    const estado=falta===null?'sin_datos':falta<=0?'pdte':falta<=(Number(r.Proximo)-Number(r.U_Medicion_med))*0.1?'prox':'ok';
    result.push({
      activo:r.Activo,
      maquina:r.Activo,
      gama:r.Gama,
      gamaNombre:r.Gama,
      intervalo:proximo&&r.U_Medicion_med?proximo-Number(r.U_Medicion_med):0,
      proximo,
      ultima:currentHoro||null,
      ultimaFecha:gasoilFecha||r.U_Medicion_fecha||'—',
      falta,
      diasRestantes,
      estado,
      isCustom:false,
      fromGasoil:gasoilHoro!=null,
      listadoId:r.id,
    });
  });
  return result;
}

function renderMantenimientoPreventivo(){
  const el=document.getElementById('prev-list');
  const fm=document.getElementById('filt-prev-maquina').value;
  const fe=document.getElementById('filt-prev-estado').value;
  const fcEl=document.getElementById('filt-prev-custom');const fc=fcEl?fcEl.checked:false;
  let data=calcMantenimiento();
  if(fm)data=data.filter(r=>r.activo===fm);
  if(fe)data=data.filter(r=>r.estado===fe);
  if(fc)data=data.filter(r=>r.isCustom);
  // Update counters
  const allData=calcMantenimiento();
  document.getElementById('prev-cnt-pdte').textContent=allData.filter(r=>r.estado==='pdte').length;
  document.getElementById('prev-cnt-prox').textContent=allData.filter(r=>r.estado==='prox').length;
  document.getElementById('prev-cnt-ok').textContent=allData.filter(r=>r.estado==='ok').length;
  // Alert banner
  const banner=document.getElementById('prev-alert-banner');
  if(banner){
    const allD=calcMantenimiento();
    const nPdte=allD.filter(r=>r.estado==='pdte').length;
    const nProx=allD.filter(r=>r.estado==='prox').length;
    if(nPdte||nProx){
      banner.style.display='block';
      const parts=[];
      if(nPdte)parts.push(`⚠ ${nPdte} gama${nPdte>1?'s':''} PENDIENTE${nPdte>1?'S':''} de mantenimiento`);
      if(nProx)parts.push(`⏰ ${nProx} gama${nProx>1?'s':''} PRÓXIMA${nProx>1?'S':''} — pulsa para filtrar`);
      banner.innerHTML=parts.join('&nbsp;&nbsp;|&nbsp;&nbsp;');
      banner.style.borderColor=nPdte?'#ff4d4d55':'#f5a62355';
      banner.style.background=nPdte?'#ff4d4d18':'#f5a62318';
      banner.style.color=nPdte?'#ff4d4d':'#f5a623';
    }else{banner.style.display='none';}
  }
  if(!data.length){el.innerHTML='<div class="tbl"><div class="empty">Sin datos</div></div>';return;}
  const estadoColor={pdte:'#ff4d4d',prox:'#f5a623',ok:'#4caf50',sin_datos:'#888'};
  const estadoLabel={pdte:'PDTE',prox:'PRÓXIMO',ok:'CORRECTO',sin_datos:'SIN DATOS'};
  el.innerHTML='<div class="tbl">'+
    '<div class="tr th">'+
      '<div class="tc" style="flex:1">Activo</div>'+
      '<div class="tc" style="flex:1.5">Gama</div>'+
      '<div class="tc" style="flex:.65;text-align:right">Horóm.</div>'+
      '<div class="tc" style="flex:.85;text-align:right">Fecha Horóm.</div>'+
      '<div class="tc" style="flex:.7;text-align:right">Próximo</div>'+
      '<div class="tc" style="flex:.6;text-align:right">Falta h</div>'+
      '<div class="tc" style="flex:.9;text-align:center">Estado</div>'+
      '<div class="tc" style="flex:.5;text-align:center"></div>'+
    '</div>'+
    data.map(r=>{
      const faltaStyle=r.falta!==null&&r.falta<0?'background:rgba(255,77,77,.15);':'';
      const faltaColor=r.falta!==null&&r.falta<0?'color:#ff4d4d;font-weight:700':'';
      const horomLabel=r.ultima!==null?(r.fromGasoil?`<span title="Gasoil">${r.ultima}</span>`:`${r.ultima}`):'—';
      const fechaLabel=r.ultimaFecha&&r.ultimaFecha!=='—'?r.ultimaFecha:'—';
      return `<div class="tr" style="${faltaStyle}">
        <div class="tc" style="flex:1;font-family:monospace;font-weight:700;color:var(--accent);font-size:.75rem">${r.activo}</div>
        <div class="tc" style="flex:1.5;font-size:.75rem">${r.gamaNombre}${r.isCustom?' <span style="font-size:.6rem;color:var(--accent2)">[C]</span>':''}</div>
        <div class="tc" style="flex:.65;text-align:right;font-family:monospace;font-size:.75rem${r.fromGasoil?';color:var(--accent2)':''}">${horomLabel}</div>
        <div class="tc" style="flex:.85;text-align:right;font-size:.7rem;color:var(--muted)">${fechaLabel}</div>
        <div class="tc" style="flex:.7;text-align:right;font-family:monospace;font-size:.75rem">${r.proximo!==null?r.proximo:'—'}</div>
        <div class="tc" style="flex:.6;text-align:right;font-family:monospace;font-size:.75rem;${faltaColor}">${r.falta!==null?r.falta:'—'}</div>
        <div class="tc" style="flex:.9;text-align:center"><span style="font-size:.65rem;font-weight:700;padding:2px 6px;border-radius:4px;background:${estadoColor[r.estado]}22;color:${estadoColor[r.estado]}">${estadoLabel[r.estado]}</span></div>
        <div class="tc" style="flex:.5;text-align:center"><button class="btn-sm" onclick="abrirModalListadoByActivoGama('${r.activo}','${r.gama}')" style="font-size:.62rem;padding:2px 6px">✏</button></div>
      </div>`;
    }).join('')+
  '</div>';
}

// ── NOTIFICACIÓN MANTENIMIENTO ────────────────────────────────
function notificarMantenimiento(){
  const data=calcMantenimiento();
  const pdtes=data.filter(r=>r.estado==='pdte');
  const proxs=data.filter(r=>r.estado==='prox');
  if(!pdtes.length&&!proxs.length)return;
  const lines=[];
  if(pdtes.length)lines.push(`⚠ ${pdtes.length} gama${pdtes.length>1?'s':''} pendiente${pdtes.length>1?'s':''}`);
  if(proxs.length)lines.push(`⏰ ${proxs.length} gama${proxs.length>1?'s':''} próxima${proxs.length>1?'s':''}`);
  const msg=lines.join(' · ');
  // Browser Notification API
  if('Notification' in window){
    if(Notification.permission==='granted'){
      new Notification('Mantenimiento Preventivo',{body:msg,icon:'',tag:'mant-preventivo'});
    }else if(Notification.permission!=='denied'){
      Notification.requestPermission().then(p=>{
        if(p==='granted')new Notification('Mantenimiento Preventivo',{body:msg,icon:'',tag:'mant-preventivo'});
      });
    }
  }
}

// ── GAMA CRUD ─────────────────────────────────────────────────
function abrirModalGama(gamaId){
  document.getElementById('mgama-msg').textContent='';
  const gamas=getEffectiveGamas();
  if(gamaId){
    const g=gamas.find(x=>x.id===gamaId);
    if(g){
      document.getElementById('mgama-title').textContent='Editar Gama';
      document.getElementById('mgama-id').value=g.id;
      document.getElementById('mgama-modelo').value=g.modelo;
      document.getElementById('mgama-nombre').value=g.nombre;
      document.getElementById('mgama-intervalo').value=g.intervalo;
      document.getElementById('mgama-checks').value=g.checks.join('\n');
    }
  }else{
    document.getElementById('mgama-title').textContent='Nueva Gama';
    document.getElementById('mgama-id').value='';
    document.getElementById('mgama-modelo').value='';
    document.getElementById('mgama-nombre').value='';
    document.getElementById('mgama-intervalo').value='';
    document.getElementById('mgama-checks').value='';
  }
  document.getElementById('modal-gama').classList.add('open');
}
function cerrarModalGama(){document.getElementById('modal-gama').classList.remove('open');}
function guardarGama(){
  const modelo=document.getElementById('mgama-modelo').value.trim();
  const nombre=document.getElementById('mgama-nombre').value.trim();
  const intervalo=parseInt(document.getElementById('mgama-intervalo').value)||0;
  const checksRaw=document.getElementById('mgama-checks').value;
  const checks=checksRaw.split('\n').map(s=>s.trim()).filter(Boolean);
  if(!modelo||!nombre||!intervalo||!checks.length){
    document.getElementById('mgama-msg').textContent='Completa todos los campos.';return;
  }
  const existingId=document.getElementById('mgama-id').value;
  const id=existingId||modelo.replace(/\s+/g,'')+'-'+intervalo+'H-'+Date.now();
  const gama={id,modelo,nombre,intervalo,checks};
  const idx=customGamas.findIndex(g=>g.id===id);
  if(idx>=0)customGamas[idx]=gama;else customGamas.push(gama);
  localStorage.setItem('customGamas',JSON.stringify(customGamas));
  cerrarModalGama();
  renderMantenimientoPreventivo();
}

// ── RECORDATORIOS MANTENIMIENTO EN INICIO ────────────────────
function renderInicioMant(){
  const el=document.getElementById('inicio-mant-alert');
  if(!el)return;
  if(!prevData.length)return;
  const data=calcMantenimiento();
  const pdtes=data.filter(r=>r.estado==='pdte').length;
  const proxs=data.filter(r=>r.estado==='prox').length;
  if(!pdtes&&!proxs){el.style.display='none';return;}
  el.style.display='block';
  el.innerHTML=`<div style="background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:12px 14px;cursor:pointer;margin-top:8px" onclick="goPage('preventivo')">
    <div style="font-size:.7rem;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.06em;margin-bottom:6px">⚙ Alertas mantenimiento</div>
    ${pdtes?`<div style="color:#ff4d4d;font-size:.82rem;font-weight:700">${pdtes} gama${pdtes>1?'s':''} pendiente${pdtes>1?'s':''}</div>`:''}
    ${proxs?`<div style="color:#f5a623;font-size:.82rem;font-weight:700">${proxs} gama${proxs>1?'s':''} próxima${proxs>1?'s':''}</div>`:''}
  </div>`;
}

// ── PREVENTIVO TABS ───────────────────────────────────────────
function prevTab(name){
  ['listado','normas','subgamas','activos','editar'].forEach(t=>{
    document.getElementById('psec-'+t).classList.toggle('active',t===name);
    document.getElementById('ptab-'+t).classList.toggle('active',t===name);
  });
  if(name==='normas')cargarNormas();
  if(name==='subgamas')cargarSubgamas();
  if(name==='activos')cargarActivosGama();
  if(name==='editar')cargarListadoPreventivo();
}

// ── NORMAS CRUD ───────────────────────────────────────────────
let normasData=[];
async function cargarNormas(){
  const el=document.getElementById('normas-list');
  el.innerHTML='<div class="tbl"><div class="empty">Cargando...</div></div>';
  const json=await apiFetch('?accion=gamasNormas').catch(e=>({ok:false,error:e.message}));
  if(!json.ok){el.innerHTML='<div class="tbl"><div class="empty">Error: '+json.error+'</div></div>';return;}
  normasData=json.data;
  // Populate norma select in subgama modal
  const sel=document.getElementById('msubg-norma');
  if(sel){sel.innerHTML=normasData.map(n=>`<option value="${n.id}">${n.Gama||n.Numero||n.id} — ${n.Modelo||''} (${n.Intervalo||0}H)</option>`).join('');}
  if(!normasData.length){el.innerHTML='<div class="tbl"><div class="empty">Sin normas</div></div>';return;}
  renderNormasFiltradas();
}
function abrirModalNorma(id){
  document.getElementById('mnorma-msg').textContent='';
  const n=id?normasData.find(x=>x.id===id):null;
  document.getElementById('mnorma-title').textContent=n?'Editar Norma':'Nueva Norma/Gama';
  document.getElementById('mnorma-id').value=n?n.id:'';
  document.getElementById('mnorma-numero').value=n?n.Numero||'':'';
  document.getElementById('mnorma-nombre').value=n?n.Gama||'':'';
  document.getElementById('mnorma-intervalo').value=n?n.Intervalo||'':'';
  // Poblar select de modelos desde activosData (campo modelo) o MACHINES
  const selModelo=document.getElementById('mnorma-modelo');
  const modelosSet=new Set();
  // Desde activosData (tblactivos)
  activosData.forEach(a=>{if(a.modelo)modelosSet.add(a.modelo);});
  // Desde MACHINES como fallback
  if(!modelosSet.size) MACHINES.forEach(m=>{if(m.modelo&&m.modelo!=='-')modelosSet.add(m.modelo);});
  const modelosList=[...modelosSet].sort();
  selModelo.innerHTML='<option value="">— Seleccionar modelo —</option>'+
    modelosList.map(m=>`<option value="${m}">${m}</option>`).join('');
  selModelo.value=n?n.Modelo||'':'';
  // checks n1..n60
  const checks=[]; for(let i=1;i<=60;i++) if(n&&n['n'+i]) checks.push(n['n'+i]);
  document.getElementById('mnorma-checks').value=checks.join('\n');
  document.getElementById('modal-norma').classList.add('open');
}
function cerrarModalNorma(){document.getElementById('modal-norma').classList.remove('open');}
function normaAutoNombre(){
  const modelo=document.getElementById('mnorma-modelo')?.value||'';
  const intervalo=document.getElementById('mnorma-intervalo')?.value||'';
  const nombreEl=document.getElementById('mnorma-nombre');
  if(!nombreEl)return;
  // Solo auto-suggest si el campo está vacío o tiene el patrón anterior auto-generado
  if(modelo&&intervalo){
    const sugerido='MANTENIMIENTO '+intervalo+'H '+modelo;
    // Sobreescribir si está vacío o si parece auto-generado (empieza por MANTENIMIENTO)
    if(!nombreEl.value||nombreEl.value.startsWith('MANTENIMIENTO ')){
      nombreEl.value=sugerido;
    }
  }
}
async function guardarNorma(){
  const Numero=document.getElementById('mnorma-numero').value.trim();
  const Gama=document.getElementById('mnorma-nombre').value.trim();
  const Modelo=document.getElementById('mnorma-modelo').value.trim();
  const Intervalo=parseInt(document.getElementById('mnorma-intervalo').value)||0;
  const checksArr=document.getElementById('mnorma-checks').value.split('\n').map(s=>s.trim()).filter(Boolean);
  if(!Gama||!Modelo||!Intervalo){document.getElementById('mnorma-msg').textContent='Gama, modelo e intervalo requeridos.';return;}
  const id=document.getElementById('mnorma-id').value;
  const payload={tipo:id?'editGamaNorma':'postGamaNorma',id:id||undefined,Numero,Gama,Modelo,Intervalo};
  for(let i=1;i<=60;i++) payload['n'+i]=checksArr[i-1]||null;
  const btn=document.querySelector('#modal-norma .btn-save');
  btn.disabled=true;btn.textContent='Guardando...';
  const json=await apiPost(payload).catch(e=>({ok:false,error:e.message}));
  btn.disabled=false;btn.textContent='Guardar';
  if(!json.ok){document.getElementById('mnorma-msg').textContent='Error: '+json.error;return;}
  cerrarModalNorma();
  await cargarNormas();
}
async function eliminarNorma(id){
  if(!confirm('¿Eliminar esta norma?'))return;
  await apiPost({tipo:'delGamaNorma',id});cargarNormas();
}

// ── SUBGAMAS CRUD ─────────────────────────────────────────────
let subgamasData=[];
async function cargarSubgamas(){
  const el=document.getElementById('subgamas-list');
  el.innerHTML='<div class="tbl"><div class="empty">Cargando...</div></div>';
  // Ensure normas loaded for dropdown
  if(!normasData.length)await cargarNormasQuiet();
  const json=await apiFetch('?accion=gamasSubgamas').catch(e=>({ok:false,error:e.message}));
  if(!json.ok){el.innerHTML='<div class="tbl"><div class="empty">Error: '+json.error+'</div></div>';return;}
  subgamasData=json.data;
  if(!subgamasData.length){el.innerHTML='<div class="tbl"><div class="empty">Sin subgamas</div></div>';return;}
  renderSubgamasFiltradas();
}
async function cargarNormasQuiet(){
  const json=await apiFetch('?accion=gamasNormas').catch(()=>({ok:false}));
  if(json.ok)normasData=json.data;
}
function renderNormasFiltradas(){
  const q=(document.getElementById('filt-normas-q')?.value||'').toLowerCase();
  const ord=document.getElementById('ord-normas')?.value||'id';
  const el=document.getElementById('normas-list');
  if(!normasData.length)return;
  let filtered=q?normasData.filter(n=>(n.Gama||'').toLowerCase().includes(q)||(n.Modelo||'').toLowerCase().includes(q)||(n.Numero||'').toLowerCase().includes(q)):[...normasData];
  if(ord==='id') filtered.sort((a,b)=>(Number(a.id)||0)-(Number(b.id)||0));
  else if(ord==='id_desc') filtered.sort((a,b)=>(Number(b.id)||0)-(Number(a.id)||0));
  else if(ord==='Gama') filtered.sort((a,b)=>(a.Gama||'').localeCompare(b.Gama||''));
  else if(ord==='Modelo') filtered.sort((a,b)=>(a.Modelo||'').localeCompare(b.Modelo||''));
  else if(ord==='Intervalo_asc') filtered.sort((a,b)=>(Number(a.Intervalo)||0)-(Number(b.Intervalo)||0));
  else if(ord==='Intervalo_desc') filtered.sort((a,b)=>(Number(b.Intervalo)||0)-(Number(a.Intervalo)||0));
  if(!filtered.length){el.innerHTML='<div class="tbl"><div class="empty">Sin coincidencias</div></div>';return;}
  el.innerHTML='<div class="tbl">'+
    '<div class="tr th"><div class="tc" style="flex:.4">#</div><div class="tc" style="flex:.5">Código</div><div class="tc" style="flex:1">Gama</div><div class="tc" style="flex:.8">Modelo</div><div class="tc" style="flex:.5;text-align:center">Intervalo</div><div class="tc" style="flex:.3;text-align:center">Checks</div><div class="tc" style="flex:.6"></div></div>'+
    filtered.map(n=>{
      let nc=0; for(let i=1;i<=60;i++) if(n['n'+i]) nc++;
      return `<div class="tr">
      <div class="tc" style="flex:.4;font-family:monospace;font-size:.7rem;color:var(--muted)">${n.id}</div>
      <div class="tc" style="flex:.5;font-size:.72rem;color:var(--muted)">${n.Numero||'—'}</div>
      <div class="tc" style="flex:1;font-size:.78rem;font-weight:600">${n.Gama||'—'}</div>
      <div class="tc" style="flex:.8;font-size:.75rem;color:var(--accent)">${n.Modelo||'—'}</div>
      <div class="tc" style="flex:.5;text-align:center;font-family:monospace;font-size:.75rem">${n.Intervalo||'—'}H</div>
      <div class="tc" style="flex:.3;text-align:center;font-size:.7rem;color:var(--muted)">${nc}</div>
      <div class="tc" style="flex:.6;text-align:right;display:flex;gap:4px;justify-content:flex-end">
        <button class="btn-sm" onclick="abrirModalNorma(${n.id})" style="font-size:.65rem;padding:2px 5px">✏</button>
        <button class="btn-sm" onclick="eliminarNorma(${n.id})" style="font-size:.65rem;padding:2px 5px;color:#ff4d4d;border-color:#ff4d4d">🗑</button>
      </div></div>`;}).join('')+
  '</div>';
}
function renderSubgamasFiltradas(){
  const q=(document.getElementById('filt-subgamas-q')?.value||'').toLowerCase();
  const ord=document.getElementById('ord-subgamas')?.value||'id';
  const el=document.getElementById('subgamas-list');
  if(!subgamasData.length)return;
  let filtered=q?subgamasData.filter(s=>(s.Gama_Principal||'').toLowerCase().includes(q)):[...subgamasData];
  if(ord==='Gama_Principal') filtered.sort((a,b)=>(a.Gama_Principal||'').localeCompare(b.Gama_Principal||''));
  if(!filtered.length){el.innerHTML='<div class="tbl"><div class="empty">Sin coincidencias</div></div>';return;}
  el.innerHTML='<div style="display:flex;flex-direction:column;gap:10px;padding:4px 0">'+
    filtered.map(s=>{
      const subs=[s.Gama_1,s.Gama_2,s.Gama_3,s.Gama_4,s.Gama_5,s.Gama_6].filter(Boolean);
      const subsHtml=subs.map((g,i)=>`
        <div style="display:flex;align-items:center;gap:6px;margin-left:28px;padding:5px 10px;background:var(--surface);border-radius:6px;border-left:2px solid var(--border)">
          <span style="font-size:.68rem;color:var(--muted);min-width:14px;font-family:monospace">${i+1}</span>
          <span style="font-size:.75rem;color:var(--fg)">${g}</span>
        </div>`).join('');
      return `<div style="background:var(--surface2);border-radius:8px;padding:10px 12px;border:1px solid var(--border)">
        <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:${subs.length?'8px':'0'}">
          <div style="display:flex;align-items:center;gap:8px">
            <span style="font-size:.65rem;color:var(--muted);font-family:monospace;min-width:20px">${s.id}</span>
            <span style="font-size:.8rem;font-weight:700;color:var(--accent)">▶ ${s.Gama_Principal||'—'}</span>
            ${subs.length?`<span style="font-size:.65rem;color:var(--muted);background:var(--surface);border-radius:10px;padding:1px 7px">${subs.length} subgama${subs.length>1?'s':''}</span>`:''}
          </div>
          <div style="display:flex;gap:4px">
            <button class="btn-sm" onclick="abrirModalSubgama(${s.id})" style="font-size:.65rem;padding:2px 5px">✏</button>
            <button class="btn-sm" onclick="eliminarSubgama(${s.id})" style="font-size:.65rem;padding:2px 5px;color:#ff4d4d;border-color:#ff4d4d">🗑</button>
          </div>
        </div>
        ${subsHtml}
      </div>`;
    }).join('')+
  '</div>';
}
function renderActivoGamaFiltrados(){
  const q=(document.getElementById('filt-activogama-q')?.value||'').toLowerCase();
  const ord=document.getElementById('ord-activogama')?.value||'id';
  const el=document.getElementById('activogama-list');
  if(!activoGamaData.length)return;
  const gamasCols=['Gama_1','Gama_2','Gama_3','Gama_4','Gama_5','Gama_6','Gama_7','Gama_8','Gama_9'];
  let filtered=q?activoGamaData.filter(a=>(a.Activo||'').toLowerCase().includes(q)):[...activoGamaData];
  if(ord==='Activo') filtered.sort((a,b)=>(a.Activo||'').localeCompare(b.Activo||''));
  if(!filtered.length){el.innerHTML='<div class="tbl"><div class="empty">Sin coincidencias</div></div>';return;}
  el.innerHTML='<div class="tbl">'+
    '<div class="tr th"><div class="tc" style="flex:.4">#</div><div class="tc" style="flex:1">Activo</div><div class="tc" style="flex:2">Gamas asignadas</div><div class="tc prev-hide-sm tc-checks" style="flex:.8">Checks</div><div class="tc" style="flex:.6"></div></div>'+
    filtered.map(a=>{
      const gamas=gamasCols.map(c=>a[c]).filter(Boolean).join(', ')||'—';
      const checks=[a.Check_1,a.Check_2,a.Check_3].filter(Boolean).join(', ')||'—';
      return `<div class="tr">
      <div class="tc" style="flex:.4;font-family:monospace;font-size:.7rem;color:var(--muted)">${a.id}</div>
      <div class="tc" style="flex:1;font-family:monospace;font-weight:700;color:var(--accent);font-size:.75rem">${a.Activo||'—'}</div>
      <div class="tc" style="flex:2;font-size:.72rem;color:var(--muted)">${gamas}</div>
      <div class="tc prev-hide-sm" style="flex:.8;font-size:.72rem;color:var(--muted)">${checks}</div>
      <div class="tc" style="flex:.6;text-align:right;display:flex;gap:4px;justify-content:flex-end">
        <button class="btn-sm" onclick="abrirModalActivoGama(${a.id})" style="font-size:.65rem;padding:2px 5px">✏</button>
        <button class="btn-sm" onclick="eliminarActivoGama(${a.id})" style="font-size:.65rem;padding:2px 5px;color:#ff4d4d;border-color:#ff4d4d">🗑</button>
      </div></div>`;}).join('')+
  '</div>';
}
function renderListadoFiltrado(){
  const q=(document.getElementById('filt-listado-q')?.value||'').toLowerCase();
  const ord=document.getElementById('ord-listado')?.value||'id';
  const el=document.getElementById('listado-prev-list');
  if(!listadoPrevData.length)return;
  let filtered=listadoPrevData.filter(r=>r.id!=null); // excluir filas sin ID
  if(q) filtered=filtered.filter(r=>(r.Activo||'').toLowerCase().includes(q)||(r.Gama||'').toLowerCase().includes(q));
  if(ord==='Activo') filtered.sort((a,b)=>(a.Activo||'').localeCompare(b.Activo||''));
  else if(ord==='Gama') filtered.sort((a,b)=>(a.Gama||'').localeCompare(b.Gama||''));
  else if(ord==='falta_asc') filtered.sort((a,b)=>{const fa=Number(a.Falta)||(Number(a.Proximo)-Number(a.U_Medicion_med));const fb=Number(b.Falta)||(Number(b.Proximo)-Number(b.U_Medicion_med));return fa-fb;});
  else if(ord==='falta_desc') filtered.sort((a,b)=>{const fa=Number(a.Falta)||(Number(a.Proximo)-Number(a.U_Medicion_med));const fb=Number(b.Falta)||(Number(b.Proximo)-Number(b.U_Medicion_med));return fb-fa;});
  else if(ord==='fecha_desc') filtered.sort((a,b)=>(b.U_Medicion_fecha||'').localeCompare(a.U_Medicion_fecha||''));
  else if(ord==='fecha_asc') filtered.sort((a,b)=>(a.U_Medicion_fecha||'').localeCompare(b.U_Medicion_fecha||''));
  if(!filtered.length){el.innerHTML='<div class="tbl"><div class="empty">Sin coincidencias</div></div>';return;}
  el.innerHTML='<div class="tbl">'+
    '<div class="tr th"><div class="tc" style="flex:1">Activo</div><div class="tc" style="flex:1.5">Gama</div><div class="tc prev-hide-sm" style="flex:.4;text-align:center">Med.</div><div class="tc" style="flex:.7;text-align:right">Horóm.</div><div class="tc" style="flex:.8;text-align:right">Fecha Horóm.</div><div class="tc" style="flex:.7;text-align:right">Próximo</div><div class="tc" style="flex:.6"></div></div>'+
    filtered.map(r=>{
      return `<div class="tr">
        <div class="tc" style="flex:1;font-family:monospace;font-weight:700;color:var(--accent);font-size:.75rem">${r.Activo||'—'}</div>
        <div class="tc" style="flex:1.5;font-size:.75rem">${r.Gama||'—'}</div>
        <div class="tc prev-hide-sm" style="flex:.4;text-align:center;font-size:.72rem;color:var(--muted)">${r.Medidor||'H'}</div>
        <div class="tc" style="flex:.7;text-align:right;font-family:monospace;font-size:.75rem;color:var(--muted)">${r.U_Medicion_med??'—'}</div>
        <div class="tc" style="flex:.8;text-align:right;font-size:.72rem;color:var(--muted)">${r.U_Medicion_fecha||'—'}</div>
        <div class="tc" style="flex:.7;text-align:right;font-family:monospace;font-size:.75rem">${r.Proximo??'—'}</div>
        <div class="tc" style="flex:.6;text-align:right;display:flex;gap:3px;justify-content:flex-end">
          <button class="btn-sm" onclick="abrirModalListado(${r.id})" style="font-size:.62rem;padding:2px 5px">✏</button>
          <button class="btn-sm" onclick="eliminarListado(${r.id})" style="font-size:.62rem;padding:2px 5px;color:#ff4d4d;border-color:#ff4d4d">🗑</button>
        </div></div>`;
    }).join('')+
  '</div>';
}
function abrirModalSubgama(id){
  document.getElementById('msubg-msg').textContent='';
  const s=id?subgamasData.find(x=>x.id===id):null;
  document.getElementById('msubg-title').textContent=s?'Editar Subgama':'Nueva Subgama';
  document.getElementById('msubg-id').value=s?s.id:'';
  const gamaOpts=normasData.map(n=>`<option value="${n.Numero||n.Gama||''}">${n.Numero||n.Gama||''}</option>`).join('');
  const principal=document.getElementById('msubg-principal');
  if(principal){principal.innerHTML='<option value="">— Seleccionar gama principal —</option>'+gamaOpts;principal.value=s?s.Gama_Principal||''  :'';}
  for(let i=1;i<=6;i++){const el2=document.getElementById('msubg-gama'+i);if(el2){el2.innerHTML='<option value="">—</option>'+gamaOpts;el2.value=s?s['Gama_'+i]||'':'';}};
  document.getElementById('modal-subgama').classList.add('open');
}
function cerrarModalSubgama(){document.getElementById('modal-subgama').classList.remove('open');}
async function guardarSubgama(){
  const Gama_Principal=document.getElementById('msubg-principal').value.trim();
  if(!Gama_Principal){document.getElementById('msubg-msg').textContent='Gama principal requerida.';return;}
  const payload={tipo:'postGamaSubgama',Gama_Principal};
  const id=document.getElementById('msubg-id').value;
  if(id){payload.tipo='editGamaSubgama';payload.id=id;}
  for(let i=1;i<=6;i++){const el2=document.getElementById('msubg-gama'+i);payload['Gama_'+i]=el2?el2.value.trim()||null:null;}
  const btn=document.querySelector('#modal-subgama .btn-save');
  btn.disabled=true;btn.textContent='Guardando...';
  const json=await apiPost(payload).catch(e=>({ok:false,error:e.message}));
  btn.disabled=false;btn.textContent='Guardar';
  if(!json.ok){document.getElementById('msubg-msg').textContent='Error: '+json.error;return;}
  cerrarModalSubgama();
  await cargarSubgamas();
}
async function eliminarSubgama(id){
  if(!confirm('¿Eliminar esta subgama?'))return;
  await apiPost({tipo:'delGamaSubgama',id});cargarSubgamas();
}

// ── ACTIVO-GAMA CRUD ──────────────────────────────────────────
let activoGamaData=[];
async function cargarActivosGama(){
  const el=document.getElementById('activogama-list');
  el.innerHTML='<div class="tbl"><div class="empty">Cargando...</div></div>';
  const json=await apiFetch('?accion=gamasActivos').catch(e=>({ok:false,error:e.message}));
  if(!json.ok){el.innerHTML='<div class="tbl"><div class="empty">Error: '+json.error+'</div></div>';return;}
  activoGamaData=json.data;
  if(!activoGamaData.length){el.innerHTML='<div class="tbl"><div class="empty">Sin registros</div></div>';return;}
  renderActivoGamaFiltrados();
  // Populate activo select in modal
  const sel=document.getElementById('mag-activo');
  if(sel){sel.innerHTML='<option value="">Seleccionar...</option>'+MACHINES.map(m=>`<option value="${m.id}">${m.id} — ${m.name}</option>`).join('');}
}
async function abrirModalActivoGama(id){
  document.getElementById('mag-msg').textContent='';
  // Cargar normasData y activosData si están vacíos
  const preloads=[];
  if(!normasData.length) preloads.push(apiFetch('?accion=gamasNormas').catch(()=>({ok:false})).then(r=>{if(r.ok)normasData=r.data||[];}));
  if(!activosData.length) preloads.push(dbQuery({action:'select',table:'tblactivos',options:{select:'*',order:'Codigo.asc'}}).then(r=>{if(r.ok&&r.data)activosData=r.data;}));
  if(preloads.length) await Promise.all(preloads);
  const a=id?activoGamaData.find(x=>x.id===id):null;
  document.getElementById('mag-title').textContent=a?'Editar Activo-Gama':'Nuevo Activo-Gama';
  document.getElementById('mag-id').value=a?a.id:'';
  document.getElementById('mag-activo').value=a?a.Activo||'':'';

  // Poblar selects filtrando por modelo del activo
  magActivoChange();
  // Si editando, restaurar valores guardados
  if(a){
    const gamaIds=['mag-codigogama','mag-gama2','mag-gama3','mag-gama4','mag-gama5','mag-gama6','mag-gama7','mag-gama8','mag-gama9'];
    gamaIds.forEach((sid,i)=>{
      const sel=document.getElementById(sid); if(!sel)return;
      const campo=i===0?'Gama_1':'Gama_'+(i+1);
      // Si el valor no está en las opciones, añadirlo
      const val=a[campo]||'';
      if(val && ![...sel.options].some(o=>o.value===val)){
        sel.insertAdjacentHTML('beforeend',`<option value="${val}">${val}</option>`);
      }
      sel.value=val;
    });
  }

  document.getElementById('modal-activogama').classList.add('open');
}
function cerrarModalActivoGama(){document.getElementById('modal-activogama').classList.remove('open');}
function magActivoChange(){
  const activoCodigo = document.getElementById('mag-activo').value;
  const aviso = document.getElementById('mag-gamas-aviso');
  const avisoTxt = document.getElementById('mag-gamas-aviso-txt');
  if(!activoCodigo){ if(aviso) aviso.style.display='none'; return; }
  // Buscar el modelo del activo seleccionado
  const activoRow = activosData.find(a=>a.Codigo===activoCodigo);
  const modelo = activoRow?.modelo || null;
  // Filtrar normas por modelo (si hay modelo)
  const normasFiltradas = modelo
    ? normasData.filter(n=>(n.Modelo||'').toLowerCase()===modelo.toLowerCase())
    : normasData;
  // Construir opciones
  const gamaOpts = '<option value="">—</option>' + normasFiltradas.map(n=>{
    const label=(n.Numero||n.id)+(n.Gama?' — '+n.Gama:'');
    return `<option value="${n.Numero||n.id}">${label}</option>`;
  }).join('');
  const gamaIds=['mag-codigogama','mag-gama2','mag-gama3','mag-gama4','mag-gama5','mag-gama6','mag-gama7','mag-gama8','mag-gama9'];
  gamaIds.forEach(sid=>{ const s=document.getElementById(sid); if(s){ s.innerHTML=gamaOpts; s.value=''; } });
  // Aviso si no hay gamas para este modelo
  if(aviso){
    if(normasFiltradas.length===0){
      aviso.style.display='flex';
      avisoTxt.textContent = modelo
        ? `No hay gamas definidas para el modelo "${modelo}"`
        : 'No hay gamas — crea una primero';
    } else {
      aviso.style.display='none';
    }
  }
}
async function guardarActivoGama(){
  const Activo=document.getElementById('mag-activo').value;
  if(!Activo){document.getElementById('mag-msg').textContent='Activo requerido.';return;}
  const id=document.getElementById('mag-id').value;
  const payload={tipo:id?'editGamaActivo':'postGamaActivo',id:id||undefined,Activo};
  // Gama_1 viene del campo codigogama por compatibilidad
  payload['Gama_1']=document.getElementById('mag-codigogama').value.trim()||null;
  for(let i=2;i<=9;i++){const el2=document.getElementById('mag-gama'+i);payload['Gama_'+i]=el2?el2.value.trim()||null:null;}
  for(let i=1;i<=3;i++){const el2=document.getElementById('mag-check'+i);payload['Check_'+i]=el2?el2.value.trim()||null:null;}
  const btn=document.querySelector('#modal-activogama .btn-save');
  btn.disabled=true;btn.textContent='Guardando...';
  const json=await apiPost(payload).catch(e=>({ok:false,error:e.message}));
  btn.disabled=false;btn.textContent='Guardar';
  if(!json.ok){document.getElementById('mag-msg').textContent='Error: '+json.error;return;}
  cerrarModalActivoGama();cargarActivosGama();
}
async function eliminarActivoGama(id){
  if(!confirm('¿Eliminar?'))return;
  await apiPost({tipo:'delGamaActivo',id});cargarActivosGama();
}

// ── LISTADO PREVENTIVO CRUD ───────────────────────────────────
let listadoPrevData=[];
async function cargarListadoPreventivo(){
  const el=document.getElementById('listado-prev-list');
  el.innerHTML='<div class="tbl"><div class="empty">Cargando...</div></div>';
  const json=await apiFetch('?accion=gamasListado').catch(e=>({ok:false,error:e.message}));
  if(!json.ok){el.innerHTML='<div class="tbl"><div class="empty">Error: '+json.error+'</div></div>';return;}
  listadoPrevData=json.data;
  if(!listadoPrevData.length){el.innerHTML='<div class="tbl"><div class="empty">Sin registros — pulsa + Nuevo para añadir</div></div>';return;}
  renderListadoFiltrado();
}
function mlistActivoChange(){
  const activo=document.getElementById('mlist-activo').value;
  const sel=document.getElementById('mlist-codigogama');
  sel.innerHTML='';
  if(!activo){sel.innerHTML='<option value="">— elige activo primero —</option>';return;}
  // Buscar en activoGamaData las gamas de este activo
  const row=activoGamaData.find(x=>x.Activo===activo);
  const gamas=[];
  if(row){for(let i=1;i<=9;i++){if(row['Gama_'+i])gamas.push(row['Gama_'+i]);}}
  // Si no hay gamas en tblGamasActivos, buscar en normasData por modelo del activo
  if(!gamas.length && normasData.length){
    const activoRow=activosData.find(a=>a.Codigo===activo||a.Activo===activo);
    const modelo=activoRow?activoRow.modelo||activoRow.Activo||activo:activo;
    normasData.forEach(n=>{
      const nModelo=n.Modelo||'';
      if(nModelo&&(nModelo===modelo||nModelo===activo)){
        const codigo=n.Numero||n.Gama||'';
        if(codigo&&!gamas.includes(codigo))gamas.push(codigo);
      }
    });
  }
  if(!gamas.length){sel.innerHTML='<option value="">Sin gamas asignadas</option>';return;}
  sel.innerHTML=gamas.map(g=>`<option value="${g}">${g}</option>`).join('');
}
async function abrirModalListado(id){
  document.getElementById('mlist-msg').textContent='';
  const r=id?listadoPrevData.find(x=>x.id===id):null;
  document.getElementById('mlist-title').textContent=r?'Editar entrada':'Nueva entrada';
  document.getElementById('mlist-id').value=r?r.id:'';
  // Cargar activoGamaData si está vacío
  if(!activoGamaData.length){
    const j=await apiFetch('?accion=gamasActivos').catch(()=>({ok:false}));
    if(j.ok)activoGamaData=j.data||[];
  }
  // Cargar activosData (tblactivos) si está vacío
  if(!activosData.length){
    const j=await dbQuery({action:'select',table:'tblactivos',options:{select:'*',order:'Codigo.asc',limit:500}});
    if(j.ok && j.data && j.data.length) activosData=j.data;
  }
  // Cargar normasData siempre (para buscar gamas por modelo)
  const jNormas=await apiFetch('?accion=gamasNormas').catch(()=>({ok:false}));
  if(jNormas.ok)normasData=jNormas.data||[];
  // Construir mapa Codigo→Nombre desde tblactivos
  const activoNombreMap={};
  activosData.forEach(a=>{if(a.Codigo)activoNombreMap[a.Codigo]=a.Activo||a.Codigo;});
  // Combinar IDs de ambas tablas sin duplicados
  const idsGamaActivos=activoGamaData.map(a=>a.Activo).filter(Boolean);
  const idsActivos=activosData.map(a=>a.Codigo||'').filter(Boolean);
  const todosActivos=[...new Set([...idsGamaActivos,...idsActivos])].sort();
  // Poblar selector activos
  const selAct=document.getElementById('mlist-activo');
  selAct.innerHTML='<option value="">Seleccionar activo...</option>'+
    todosActivos.map(a=>`<option value="${a}">${a}${activoNombreMap[a]&&activoNombreMap[a]!==a?' — '+activoNombreMap[a]:''}</option>`).join('');
  selAct.value=r?r.Activo||'':'';
  // Poblar gamas dependiendo del activo seleccionado
  mlistActivoChange();
  // Si editando, seleccionar la gama actual
  if(r&&r.Gama){
    const selGama=document.getElementById('mlist-codigogama');
    // Si la gama no está en las opciones (caso raro), añadirla
    if(![...selGama.options].some(o=>o.value===r.Gama)){
      selGama.insertAdjacentHTML('beforeend',`<option value="${r.Gama}">${r.Gama}</option>`);
    }
    selGama.value=r.Gama;
  }
  document.getElementById('mlist-medidor').value=r?r.Medidor||'H':'H';
  document.getElementById('mlist-proximo').value=r?r.Proximo??'':'';
  document.getElementById('mlist-ultima').value=r?r.U_Medicion_med??'':'';
  document.getElementById('mlist-ultimafecha').value=r?r.U_Medicion_fecha||'':'';
  document.getElementById('modal-listado').classList.add('open');
}
function cerrarModalListado(){document.getElementById('modal-listado').classList.remove('open');}
async function abrirModalListadoByActivoGama(activo, gama){
  // Ensure listadoPrevData loaded
  if(!listadoPrevData.length){
    const j=await apiFetch('?accion=gamasListado').catch(()=>({ok:false}));
    if(j.ok)listadoPrevData=j.data||[];
  }
  const r=listadoPrevData.find(x=>x.Activo===activo&&x.Gama===gama);
  if(r) abrirModalListado(r.id);
  else abrirModalListado(); // new entry pre-filled below if not found
}
async function guardarListado(){
  const Activo=document.getElementById('mlist-activo').value.trim();
  const Gama=document.getElementById('mlist-codigogama').value.trim();
  const Proximo=document.getElementById('mlist-proximo').value;
  const U_Medicion_med=document.getElementById('mlist-ultima').value;
  if(!Activo||!Gama){document.getElementById('mlist-msg').textContent='Activo y gama requeridos.';return;}
  const id=document.getElementById('mlist-id').value;
  const btn=document.querySelector('#modal-listado .btn-save');
  btn.disabled=true;btn.textContent='Guardando...';
  const payload={
    tipo:id?'editGamaListado':'postGamaListado',
    id:id||undefined,Activo,Gama,
    Medidor:document.getElementById('mlist-medidor').value,
    Proximo:Number(Proximo)||0,U_Medicion_med:Number(U_Medicion_med)||0,
    U_Medicion_fecha:document.getElementById('mlist-ultimafecha').value||null,
  };
  const json=await apiPost(payload).catch(e=>({ok:false,error:e.message}));
  btn.disabled=false;btn.textContent='Guardar';
  if(!json.ok){document.getElementById('mlist-msg').textContent='Error: '+json.error;return;}
  cerrarModalListado();cargarListadoPreventivo();
}
async function eliminarListado(id){
  const idNum=Number(id);
  if(!idNum){alert('ID inválido');return;}
  if(!confirm('¿Eliminar esta entrada del listado?'))return;
  const json=await apiPost({tipo:'delGamaListado',id:idNum}).catch(e=>({ok:false,error:e.message}));
  if(!json.ok){alert('Error: '+json.error);return;}
  cargarListadoPreventivo();
}

// ── OT HISTORIAL PRINT ────────────────────────────────────────
function printOTHistorial(id){
  const r=otHistData.find(x=>x.id==id);
  if(!r)return;
  const machine=MACHINES.find(m=>m.id===r.activo)||{name:r.activo,fabricante:'—'};
  const gama=getEffectiveGamas().find(g=>g.id===r.gama)||{nombre:r.gama,intervalo:'—',checks:[]};
  const checks=Array.isArray(r.checks)?r.checks:[];
  document.getElementById('otp-ref').textContent='OT-'+String(r.ot||r.id).padStart(4,'0');
  document.getElementById('otp-maquina').textContent=machine.name+' ('+r.activo+')';
  document.getElementById('otp-fecha').textContent=r.fecha||'';
  document.getElementById('otp-activo').textContent=r.activo;
  document.getElementById('otp-gama').textContent=gama.nombre+(gama.intervalo!=='—'?' ('+gama.intervalo+'h)':'');
  document.getElementById('otp-horas').textContent=r.medicion?r.medicion+' h':'';
  document.getElementById('otp-tipo').textContent='Preventivo programado';
  document.getElementById('otp-operario').textContent=r.operario||'';
  document.getElementById('otp-fabricante').textContent=machine.fabricante||'';
  document.getElementById('otp-obs').textContent=r.texto||'';
  const tbl=document.getElementById('otp-checks-table');
  let rows='';
  gama.checks.forEach((c,i)=>{
    const done=checks[i]===true||checks[i]==='TRUE';
    rows+=`<tr style="background:${i%2===0?'#f9f9f9':'#fff'}">
      <td style="border:1px solid #ddd;padding:3px 6px;width:60%;font-size:7.5pt">${c}</td>
      <td style="border:1px solid #ddd;padding:3px 6px;width:20%;text-align:center;font-size:7.5pt;color:${done?'#2e7d32':'#999'}">${done?'✓ OK':'—'}</td>
      <td style="border:1px solid #ddd;padding:3px 6px;width:20%;font-size:7pt;color:#999"></td>
    </tr>`;
  });
  if(!rows)rows='<tr><td colspan="3" style="padding:6px;font-size:7.5pt;color:#999">Sin puntos de verificación registrados</td></tr>';
  tbl.innerHTML=`<thead><tr style="background:#444;color:#fff"><th style="padding:3px 6px;border:1px solid #444;text-align:left;width:60%">Punto de verificación</th><th style="padding:3px 6px;border:1px solid #444;width:20%">Estado</th><th style="padding:3px 6px;border:1px solid #444;width:20%">Observación</th></tr></thead><tbody>${rows}</tbody>`;
  const wrap=document.getElementById('ot-print-wrap');
  wrap.style.display='flex';
  wrap.style.position='fixed';
  setTimeout(()=>window.print(),100);
}

// ── COSTES — Análisis cuentas 600/700 desde BC ──────────────────────────────

// Mapa de códigos CA → categoría (concepto) y nombre
const COSTES_CA_MAP = {
  C000:{cat:'0.SIN DEFINIR',name:'SIN DEFINIR',orden:'0.GASTO'},
  C100:{cat:'1.PLANTA',name:'COSTES DE PLANTA',orden:'0.GASTO'},
  C101:{cat:'1.PLANTA',name:'INVERSION INICIAL AMORTIZABLE',orden:'0.GASTO'},
  C102:{cat:'1.PLANTA',name:'ELECTRICIDAD',orden:'0.GASTO'},
  C103:{cat:'1.PLANTA',name:'PIEZAS DESGASTE',orden:'0.GASTO'},
  C104:{cat:'1.PLANTA',name:'MANTENIMIENTO',orden:'0.GASTO'},
  C105:{cat:'1.PLANTA',name:'SUBCONTRATACIONES',orden:'0.GASTO'},
  C106:{cat:'1.PLANTA',name:'VIGILANCIA Y SEGURIDAD',orden:'0.GASTO'},
  C107:{cat:'AMORTIZACION',name:'AMORTIZACIONES',orden:'1.AMORTIZA'},
  C108:{cat:'1.PLANTA',name:'CALIDAD Y MEDIO AMBIENTE',orden:'0.GASTO'},
  C200:{cat:'2.EXTRACCION PIEDRA',name:'COSTES EXTRACCION DE PIEDRA',orden:'0.GASTO'},
  C201:{cat:'2.EXTRACCION PIEDRA',name:'VOLADURA',orden:'0.GASTO'},
  C202:{cat:'2.EXTRACCION PIEDRA',name:'CANON',orden:'0.GASTO'},
  C203:{cat:'2.EXTRACCION PIEDRA',name:'EXCAVADORA-MARTILLO HIDRAULICO',orden:'0.GASTO'},
  C204:{cat:'2.EXTRACCION PIEDRA',name:'EXISTENCIAS ARIDOS',orden:'0.GASTO'},
  C300:{cat:'3.PERSONAL',name:'COSTES DE PERSONAL',orden:'0.GASTO'},
  C301:{cat:'3.PERSONAL',name:'SALARIOS',orden:'0.GASTO'},
  C302:{cat:'3.PERSONAL',name:'SEGURIDAD SOCIAL',orden:'0.GASTO'},
  C303:{cat:'3.PERSONAL',name:'PREV. RIESGOS LABORALES',orden:'0.GASTO'},
  C304:{cat:'3.PERSONAL',name:'FORMACIONES',orden:'0.GASTO'},
  C400:{cat:'4.ADMINISTRACION',name:'COSTES ADMINISTRACION',orden:'0.GASTO'},
  C401:{cat:'4.ADMINISTRACION',name:'COSTES FINANCIEROS',orden:'0.GASTO'},
  C402:{cat:'4.ADMINISTRACION',name:'MATERIAL OFICINA',orden:'0.GASTO'},
  C403:{cat:'4.ADMINISTRACION',name:'DIETAS, VIAJES, COMB...',orden:'0.GASTO'},
  C404:{cat:'4.ADMINISTRACION',name:'GASTOS DE GESTION',orden:'0.GASTO'},
  C405:{cat:'4.ADMINISTRACION',name:'COSTES OPERATIVOS',orden:'0.GASTO'},
  C406:{cat:'4.ADMINISTRACION',name:'OFIMATICA',orden:'0.GASTO'},
  C407:{cat:'4.ADMINISTRACION',name:'PRIMAS DE SEGUROS',orden:'0.GASTO'},
  C408:{cat:'4.ADMINISTRACION',name:'SEGURIDAD E HIGIENE',orden:'0.GASTO'},
  C409:{cat:'4.ADMINISTRACION',name:'TASAS Y TRIBUTOS',orden:'0.GASTO'},
  C500:{cat:'5.TALLER',name:'COSTES TALLER',orden:'0.GASTO'},
  C501:{cat:'5.TALLER',name:'HERRAMIENTAS TALLER',orden:'0.GASTO'},
  C502:{cat:'5.TALLER',name:'CONSUMIBLES TALLER',orden:'0.GASTO'},
  C503:{cat:'5.TALLER',name:'REPUESTOS',orden:'0.GASTO'},
  C600:{cat:'6.MAQUINARIA',name:'MAQUINARIA MOVIL',orden:'0.GASTO'},
  C601:{cat:'6.MAQUINARIA',name:'PALA CARGADORA',orden:'0.GASTO'},
  C602:{cat:'6.MAQUINARIA',name:'EXCAVADORA',orden:'0.GASTO'},
  C603:{cat:'6.MAQUINARIA',name:'DUMPER',orden:'0.GASTO'},
  C604:{cat:'6.MAQUINARIA',name:'CAMION',orden:'0.GASTO'},
  C605:{cat:'6.MAQUINARIA',name:'CUBA DE AGUA',orden:'0.GASTO'},
  C606:{cat:'6.MAQUINARIA',name:'CAMION GRUA',orden:'0.GASTO'},
  C607:{cat:'6.MAQUINARIA',name:'OTRA MAQUINARIA',orden:'0.GASTO'},
  C608:{cat:'7.COMBUSTIBLE',name:'COMBUSTIBLE',orden:'0.GASTO'},
  C700:{cat:'8.SERVICIOS',name:'PRESTACION SERVICIO',orden:'0.GASTO'},
  C999:{cat:'8.SERVICIOS',name:'TOTAL COSTES',orden:'0.GASTO'},
  I000:{cat:'INGRESOS',name:'INGRESOS',orden:'2.INGRESO'},
  I100:{cat:'INGRESOS',name:'FACTURACION',orden:'2.INGRESO'},
  I999:{cat:'INGRESOS',name:'TOTAL INGRESOS',orden:'2.INGRESO'}
};

// Orden de categorías para la tabla
const COSTES_CAT_ORDER = [
  '1.PLANTA','2.EXTRACCION PIEDRA','3.PERSONAL','4.ADMINISTRACION',
  '5.TALLER','6.MAQUINARIA','7.COMBUSTIBLE','8.SERVICIOS','AMORTIZACION','INGRESOS',
  '0.SIN DEFINIR','#N/D'
];

const MESES_NOMBRE = ['','Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];

let costesRawData = [];  // entries from BC
let costesProduccion = []; // producción from Supabase
let costesAnyoCargado = null;
let costesExcludedAccounts = new Set(JSON.parse(localStorage.getItem('costesExcludedAccounts')||'[]'));

function initCostes(){
  const sel = document.getElementById('costes-anyo');
  if(!sel.options.length){
    const y = new Date().getFullYear();
    for(let i=y;i>=y-3;i--){
      const o = document.createElement('option');
      o.value=i; o.textContent=i;
      sel.appendChild(o);
    }
  }
  // Default mes hasta = current month
  const curMonth = new Date().getMonth()+1;
  document.getElementById('costes-mes-hasta').value = curMonth;
  document.getElementById('costes-mes-desde').value = 1;
}

function switchCostesTab(tab){
  document.getElementById('costes-panel-analisis').style.display = tab==='analisis' ? '' : 'none';
  document.getElementById('costes-panel-config').style.display = tab==='config' ? '' : 'none';
  document.getElementById('costes-tab-analisis').classList.toggle('active', tab==='analisis');
  document.getElementById('costes-tab-config').classList.toggle('active', tab==='config');
  if(tab==='config') renderConfigCostes();
}

function renderConfigCostes(){
  const wrap = document.getElementById('costes-config-accounts');
  if(!costesRawData.length){
    wrap.innerHTML='<div style="color:var(--muted);font-style:italic;font-size:.78rem">Carga datos de BC primero para ver las cuentas disponibles.</div>';
    return;
  }
  // Collect unique accounts with their total debit-credit and name
  const accMap = {};
  for(const e of costesRawData){
    const acc = e.account || '?';
    if(!accMap[acc]) accMap[acc] = { desc: e.accountName || e.description||'', total: 0, count: 0 };
    else if(!accMap[acc].desc && e.accountName) accMap[acc].desc = e.accountName;
    accMap[acc].total += (e.debit||0) - (e.credit||0);
    accMap[acc].count++;
  }
  const accounts = Object.keys(accMap).sort();
  const fmtES = v => {
    if(!v) return '';
    const neg = v<0;
    const [ent,dec] = Math.abs(v).toFixed(2).split('.');
    return (neg?'-':'')+ent.replace(/\B(?=(\d{3})+(?!\d))/g,'.')+','+dec+' €';
  };
  let html = '<table class="costes-cfg-tbl"><thead><tr>'
    + '<th class="c">Incluir</th>'
    + '<th>Cuenta</th>'
    + '<th>Nombre</th>'
    + '<th class="r costes-cfg-hide-sm">Movim.</th>'
    + '<th class="r">Importe</th>'
    + '</tr></thead><tbody>';
  for(const acc of accounts){
    const info = accMap[acc];
    const excluded = costesExcludedAccounts.has(acc);
    html += `<tr class="${excluded?'excluded':''}">
      <td class="c"><input type="checkbox" ${excluded?'':'checked'} onchange="costesToggleAccount('${acc}',this.checked)" style="cursor:pointer;width:16px;height:16px"></td>
      <td class="costes-cfg-acc">${acc}</td>
      <td class="costes-cfg-name">${info.desc}</td>
      <td class="r costes-cfg-hide-sm" style="color:var(--muted)">${info.count}</td>
      <td class="r" style="font-variant-numeric:tabular-nums;white-space:nowrap">${fmtES(info.total)}</td>
    </tr>`;
  }
  html += '</tbody></table>';
  wrap.innerHTML = html;
}

function costesToggleAccount(acc, included){
  if(included) costesExcludedAccounts.delete(acc);
  else costesExcludedAccounts.add(acc);
  localStorage.setItem('costesExcludedAccounts', JSON.stringify([...costesExcludedAccounts]));
  renderCostes();
  renderConfigCostes();
}

function costesConfigSelAll(include){
  if(include){
    costesExcludedAccounts.clear();
  } else {
    for(const e of costesRawData) costesExcludedAccounts.add(e.account||'?');
  }
  localStorage.setItem('costesExcludedAccounts', JSON.stringify([...costesExcludedAccounts]));
  renderCostes();
  renderConfigCostes();
}

async function cargarCostes(){
  const btn = document.getElementById('costes-btn-cargar');
  const info = document.getElementById('costes-info');
  const anyo = document.getElementById('costes-anyo').value;
  btn.disabled=true; btn.textContent='Cargando...';
  info.textContent='Conectando con Business Central...';

  try {
    const token = await getBCToken();

    // Cargar GL entries y producción en paralelo
    const [bcRes, prodRes] = await Promise.all([
      fetch('/api/bc/costes', {
        method:'POST',
        headers:{'Content-Type':'application/json'},
        body:JSON.stringify({token, anyo})
      }).then(r=>r.json()),
      getProduccion(null, anyo)
    ]);

    if(!bcRes.ok) throw new Error(bcRes.error||'Error cargando datos BC');

    costesRawData = bcRes.entries || [];
    costesProduccion = prodRes.data || [];
    costesAnyoCargado = anyo;

    info.textContent = `${costesRawData.length} movimientos cargados del ${anyo} · Producción: ${costesProduccion.length} días`;
    renderCostes();
    renderRendimiento();
  } catch(e){
    info.textContent = 'Error: '+e.message;
    console.error('Costes error:', e);
  } finally {
    btn.disabled=false; btn.textContent='↻ Cargar de BC';
  }
}

function renderCostes(){
  const wrap = document.getElementById('costes-table-wrap');
  if(!costesRawData.length){ wrap.innerHTML='<div style="color:var(--muted);text-align:center;padding:40px;font-size:.82rem">Sin datos. Pulsa "Cargar de BC".</div>'; return; }

  const vista = document.getElementById('costes-vista').value;
  const mesDesde = parseInt(document.getElementById('costes-mes-desde').value);
  const mesHasta = parseInt(document.getElementById('costes-mes-hasta').value);
  const isEurTn = vista === 'eurtn' || vista === 'eurtn-acum';
  const isAcum = vista === 'acumulado' || vista === 'eurtn-acum';

  // Calcular producción por mes
  const prodMes = {};
  for(const p of costesProduccion){
    const m = parseInt(p.fecha.split('-')[1]);
    if(m>=mesDesde && m<=mesHasta){
      prodMes[m] = (prodMes[m]||0) + (parseFloat(p.tnDia)||0);
    }
  }

  // Agrupar movimientos por CA code + mes
  // Estructura: { caCode: { mes: importe } }
  const data = {};
  const mesesActivos = [];
  for(let m=mesDesde;m<=mesHasta;m++) mesesActivos.push(m);

  for(const e of costesRawData){
    if(costesExcludedAccounts.has(e.account||'?')) continue;
    const m = parseInt(e.date.split('-')[1]);
    if(m<mesDesde || m>mesHasta) continue;
    const ca = e.ca || '#N/D';
    if(!data[ca]) data[ca]={};
    const importe = (e.debit||0) - (e.credit||0);
    data[ca][m] = (data[ca][m]||0) + importe;
  }

  // Agrupar por categoría
  const catData = {}; // { cat: { subcats: { caCode: {mes:val} }, totals: {mes:val} } }
  for(const [ca, meses] of Object.entries(data)){
    const info = COSTES_CA_MAP[ca] || {cat:'#N/D', name:ca, orden:'0.GASTO'};
    const cat = info.cat;
    if(!catData[cat]) catData[cat] = {subcats:{}, totals:{}};
    catData[cat].subcats[ca] = {name:info.name, meses};
    for(const [m, v] of Object.entries(meses)){
      catData[cat].totals[m] = (catData[cat].totals[m]||0) + v;
    }
  }

  // Build HTML table
  const fmtES = (v, dec=2) => {
    if(v===0||v===undefined||v===null) return '';
    const neg = v < 0;
    const [ent, dec2] = Math.abs(v).toFixed(dec).split('.');
    const miles = ent.replace(/\B(?=(\d{3})+(?!\d))/g, '.');
    return (neg ? '-' : '') + miles + ',' + dec2;
  };
  const fmt = v => fmtES(v, 2);
  const fmtTn = v => (!v || v===0) ? '' : fmtES(v, 2);

  // Producción acumulada
  const prodAcum = {};
  if(isAcum){
    let acum = 0;
    for(const m of mesesActivos){ acum += (prodMes[m]||0); prodAcum[m]=acum; }
  }

  // Helper: apply €/tn division
  const applyTn = (v, m) => {
    if(!isEurTn) return v;
    const tn = isAcum ? (prodAcum[m]||0) : (prodMes[m]||0);
    return tn ? v/tn : 0;
  };
  const applyTnTotal = (v) => {
    if(!isEurTn) return v;
    const totalTn = Object.values(prodMes).reduce((a,b)=>a+b,0);
    return totalTn ? v/totalTn : 0;
  };
  const colUnit = isEurTn ? '€/Tn' : '€';

  let html = '<table class="costes-tbl"><thead><tr><th class="costes-cat-col">Concepto</th>';
  for(const m of mesesActivos){
    const mName = MESES_NOMBRE[m].substring(0,3);
    html += `<th class="costes-val-col">${mName} ${colUnit}</th>`;
  }
  html += `<th class="costes-val-col costes-total-col">Total ${colUnit}</th>`;
  html += '</tr>';
  // Producción row
  html += '<tr class="costes-prod-row"><td>Producción (Tn)</td>';
  let totalProd = 0;
  for(const m of mesesActivos){
    const prod = isAcum ? (prodAcum[m]||0) : (prodMes[m]||0);
    totalProd += (prodMes[m]||0);
    html += `<td style="text-align:center;font-weight:600">${fmtTn(prod)}</td>`;
  }
  html += `<td style="text-align:center;font-weight:600">${fmtTn(totalProd)}</td>`;
  html += '</tr></thead><tbody>';

  // Grand totals for total column
  const grandTotals = {};

  for(const cat of COSTES_CAT_ORDER){
    const cd = catData[cat];
    if(!cd) continue;

    // Acumulado: sumas progresivas (siempre en € antes de /tn)
    let catTotals = {};
    if(isAcum){
      let acum = 0;
      for(const m of mesesActivos){ acum += (cd.totals[m]||0); catTotals[m]=acum; }
    } else {
      catTotals = cd.totals;
    }

    // Category header row
    let catTotal = 0;
    for(const m of mesesActivos) catTotal += (cd.totals[m]||0);
    const isIngreso = cat==='INGRESOS';
    const catClass = isIngreso ? 'costes-cat-row costes-ingreso' : 'costes-cat-row';

    const catId = cat.replace(/[^A-Za-z0-9]/g,'_');
    html += `<tr class="${catClass}" data-cat="${catId}" onclick="toggleCostesCat('${catId}')"><td class="costes-cat-name"><span class="costes-cat-toggle open" id="tog-${catId}">▶</span>${cat}</td>`;
    for(const m of mesesActivos){
      const v = applyTn(catTotals[m]||0, m);
      html += `<td class="costes-val">${fmt(v)}</td>`;
    }
    html += `<td class="costes-val costes-total-col">${fmt(applyTnTotal(catTotal))}</td>`;
    html += '</tr>';

    // Subcategory rows
    const subcats = Object.entries(cd.subcats).sort((a,b)=>a[0].localeCompare(b[0]));
    for(const [ca, sub] of subcats){
      let subTotals = {};
      if(isAcum){
        let acum=0;
        for(const m of mesesActivos){ acum += (sub.meses[m]||0); subTotals[m]=acum; }
      } else {
        subTotals = sub.meses;
      }

      let subTotal = 0;
      for(const m of mesesActivos) subTotal += (sub.meses[m]||0);

      html += `<tr class="costes-sub-row" data-parent="${catId}"><td class="costes-sub-name">${sub.name}</td>`;
      for(const m of mesesActivos){
        const v = applyTn(subTotals[m]||0, m);
        html += `<td class="costes-val">${fmt(v)}</td>`;
      }
      html += `<td class="costes-val costes-total-col">${fmt(applyTnTotal(subTotal))}</td>`;
      html += '</tr>';
    }

    // Accumulate grand totals
    for(const m of mesesActivos){
      grandTotals[m] = (grandTotals[m]||0) + (cd.totals[m]||0);
    }
  }

  // Grand total row
  let grandTotal = 0;
  for(const m of mesesActivos) grandTotal += (grandTotals[m]||0);

  let gtAccum = {};
  if(isAcum){
    let acum=0;
    for(const m of mesesActivos){ acum += (grandTotals[m]||0); gtAccum[m]=acum; }
  } else {
    gtAccum = grandTotals;
  }

  html += '<tr class="costes-grand-row"><td>TOTAL GENERAL</td>';
  for(const m of mesesActivos){
    const v = applyTn(gtAccum[m]||0, m);
    html += `<td class="costes-val">${fmt(v)}</td>`;
  }
  html += `<td class="costes-val costes-total-col">${fmt(applyTnTotal(grandTotal))}</td>`;
  html += '</tr></tbody></table>';

  wrap.innerHTML = html;

  // Show toggle button
  const tw = document.getElementById('costes-toggle-wrap');
  if(tw) tw.style.display='';

  // Renderizar gráficos
  renderCostesCharts(catData, mesesActivos, prodMes, prodAcum, totalProd);
}

// ── COSTES COLLAPSE/EXPAND ───────────────────────────────────────────────────
function toggleCostesCat(catId){
  const rows = document.querySelectorAll(`.costes-sub-row[data-parent="${catId}"]`);
  const tog = document.getElementById('tog-'+catId);
  const collapsed = !rows[0]?.classList.contains('collapsed');
  rows.forEach(r=>r.classList.toggle('collapsed',collapsed));
  if(tog) tog.classList.toggle('open',!collapsed);
  updateCostesToggleBtn();
}
function toggleAllCostesRows(){
  const allSub = document.querySelectorAll('.costes-sub-row');
  const anyVisible = [...allSub].some(r=>!r.classList.contains('collapsed'));
  allSub.forEach(r=>r.classList.toggle('collapsed',anyVisible));
  document.querySelectorAll('.costes-cat-toggle').forEach(t=>t.classList.toggle('open',!anyVisible));
  updateCostesToggleBtn();
}
function updateCostesToggleBtn(){
  const btn = document.getElementById('costes-toggle-all');
  if(!btn) return;
  const allSub = document.querySelectorAll('.costes-sub-row');
  const anyVisible = [...allSub].some(r=>!r.classList.contains('collapsed'));
  btn.textContent = anyVisible ? '▼ Contraer todo' : '▶ Expandir todo';
}

// ── COSTES CHARTS ────────────────────────────────────────────────────────────

const COSTES_COLORS = [
  '#6b7d2e','#d4a017','#2e7d6b','#7d2e6b','#2e4a7d',
  '#a0522d','#4682b4','#8b4513','#556b2f','#8b008b'
];
let costesCharts = {};

function destroyCostesCharts(){
  Object.values(costesCharts).forEach(c=>{ try{c.destroy();}catch(e){} });
  costesCharts = {};
}

function renderCostesCharts(catData, mesesActivos, prodMes, prodAcum, totalProd){
  if(typeof Chart==='undefined') return;
  destroyCostesCharts();
  document.getElementById('costes-charts').style.display='block';

  const labels = mesesActivos.map(m=>MESES_NOMBRE[m].substring(0,3));
  const gastoCats = COSTES_CAT_ORDER.filter(c=>c!=='INGRESOS'&&c!=='AMORTIZACION');

  // Defaults Chart.js
  Chart.defaults.color = '#707070';
  Chart.defaults.borderColor = 'rgba(0,0,0,0.08)';

  // ── 1. Barras apiladas: coste por categoría/mes ──
  const stackedDatasets = gastoCats.map((cat,i)=>{
    const cd = catData[cat];
    return {
      label: cat,
      data: mesesActivos.map(m=> cd ? (cd.totals[m]||0) : 0),
      backgroundColor: COSTES_COLORS[i % COSTES_COLORS.length],
      borderWidth: 0
    };
  }).filter(ds=>ds.data.some(v=>v!==0));

  costesCharts.barras = new Chart(document.getElementById('chart-barras-cat'),{
    type:'bar',
    data:{ labels, datasets: stackedDatasets },
    options:{
      responsive:true, maintainAspectRatio:false,
      plugins:{
        legend:{ display:true, position:'bottom', labels:{boxWidth:10,font:{size:9}} },
        tooltip:{ callbacks:{ label:ctx=>ctx.dataset.label+': '+ctx.parsed.y.toLocaleString('es-ES',{minimumFractionDigits:2,maximumFractionDigits:2})+'€' } }
      },
      scales:{
        x:{ stacked:true },
        y:{ stacked:true, ticks:{ callback:v=>{const n=Math.abs(v);const s=Math.round(n).toString().replace(/\B(?=(\d{3})+(?!\d))/g,'.');return (v<0?'-':'')+s+'€';} } }
      }
    }
  });

  // ── 2. Línea: €/tn mensual por categoría ──
  const lineDatasets = gastoCats.map((cat,i)=>{
    const cd = catData[cat];
    if(!cd) return null;
    const vals = mesesActivos.map(m=>{
      const v = cd.totals[m]||0;
      const p = prodMes[m]||0;
      return p ? +(v/p).toFixed(2) : 0;
    });
    if(vals.every(v=>v===0)) return null;
    return {
      label: cat,
      data: vals,
      borderColor: COSTES_COLORS[i % COSTES_COLORS.length],
      backgroundColor: 'transparent',
      borderWidth: 2,
      tension: 0.3,
      pointRadius: 3
    };
  }).filter(Boolean);

  // Add total €/tn line
  const totalTnLine = mesesActivos.map(m=>{
    const p = prodMes[m]||0;
    if(!p) return 0;
    let total = 0;
    for(const cat of gastoCats){ const cd=catData[cat]; if(cd) total+=(cd.totals[m]||0); }
    return +(total/p).toFixed(2);
  });
  lineDatasets.push({
    label:'TOTAL',
    data:totalTnLine,
    borderColor:'#1a1a1a',
    backgroundColor:'transparent',
    borderWidth:3,
    borderDash:[6,3],
    tension:0.3,
    pointRadius:4
  });

  costesCharts.linea = new Chart(document.getElementById('chart-linea-tn'),{
    type:'line',
    data:{ labels, datasets: lineDatasets },
    options:{
      responsive:true, maintainAspectRatio:false,
      plugins:{
        legend:{ display:true, position:'bottom', labels:{boxWidth:10,font:{size:9}} },
        tooltip:{ callbacks:{ label:ctx=>ctx.dataset.label+': '+ctx.parsed.y.toLocaleString('es-ES',{minimumFractionDigits:2,maximumFractionDigits:2})+'€/tn' } }
      },
      scales:{
        y:{ ticks:{ callback:v=>v.toLocaleString('es-ES',{minimumFractionDigits:1,maximumFractionDigits:1})+'€/tn' } }
      }
    }
  });

  // ── 3. Donut: distribución de costes total ──
  const donutLabels = [];
  const donutValues = [];
  const donutColors = [];
  gastoCats.forEach((cat,i)=>{
    const cd = catData[cat];
    if(!cd) return;
    let total = 0;
    mesesActivos.forEach(m=> total += (cd.totals[m]||0));
    if(total > 0){
      donutLabels.push(cat);
      donutValues.push(+total.toFixed(2));
      donutColors.push(COSTES_COLORS[i % COSTES_COLORS.length]);
    }
  });

  costesCharts.donut = new Chart(document.getElementById('chart-donut'),{
    type:'doughnut',
    data:{
      labels: donutLabels,
      datasets:[{ data:donutValues, backgroundColor:donutColors, borderWidth:0 }]
    },
    options:{
      responsive:true, maintainAspectRatio:false,
      plugins:{
        legend:{ display:true, position:'right', labels:{boxWidth:10,font:{size:9}} },
        tooltip:{ callbacks:{ label:ctx=>ctx.label+': '+ctx.parsed.toLocaleString('es-ES',{minimumFractionDigits:2})+'€' } }
      }
    }
  });

  // ── 4. Barras: Ingresos vs Gastos por mes ──
  const gastosMes = mesesActivos.map(m=>{
    let total=0;
    gastoCats.forEach(cat=>{ const cd=catData[cat]; if(cd) total+=(cd.totals[m]||0); });
    return +total.toFixed(2);
  });
  const ingresosMes = mesesActivos.map(m=>{
    const cd = catData['INGRESOS'];
    return cd ? +Math.abs(cd.totals[m]||0).toFixed(2) : 0;
  });

  costesCharts.inggas = new Chart(document.getElementById('chart-ing-gas'),{
    type:'bar',
    data:{
      labels,
      datasets:[
        { label:'Gastos', data:gastosMes, backgroundColor:'#d4a017', borderWidth:0 },
        { label:'Ingresos', data:ingresosMes, backgroundColor:'#4caf50', borderWidth:0 }
      ]
    },
    options:{
      responsive:true, maintainAspectRatio:false,
      plugins:{
        legend:{ display:true, position:'bottom', labels:{boxWidth:10,font:{size:9}} },
        tooltip:{ callbacks:{ label:ctx=>ctx.dataset.label+': '+ctx.parsed.y.toLocaleString('es-ES',{minimumFractionDigits:2,maximumFractionDigits:2})+'€' } }
      },
      scales:{
        y:{ ticks:{ callback:v=>{const n=Math.abs(v);const s=Math.round(n).toString().replace(/\B(?=(\d{3})+(?!\d))/g,'.');return (v<0?'-':'')+s+'€';} } }
      }
    }
  });

  // ── 5. Línea: Acumulado ingresos vs gastos ──
  let acumGas=0, acumIng=0;
  const acumGasArr=[], acumIngArr=[], acumBenArr=[];
  mesesActivos.forEach((m,i)=>{
    acumGas += gastosMes[i];
    acumIng += ingresosMes[i];
    acumGasArr.push(+acumGas.toFixed(2));
    acumIngArr.push(+acumIng.toFixed(2));
    acumBenArr.push(+(acumIng-acumGas).toFixed(2));
  });

  costesCharts.acumulado = new Chart(document.getElementById('chart-acumulado'),{
    type:'line',
    data:{
      labels,
      datasets:[
        { label:'Gastos acum.', data:acumGasArr, borderColor:'#d4a017', backgroundColor:'rgba(212,160,23,0.1)', fill:true, borderWidth:2, tension:0.3 },
        { label:'Ingresos acum.', data:acumIngArr, borderColor:'#4caf50', backgroundColor:'rgba(76,175,80,0.1)', fill:true, borderWidth:2, tension:0.3 },
        { label:'Resultado', data:acumBenArr, borderColor:'#1a1a1a', backgroundColor:'transparent', borderWidth:2, borderDash:[6,3], tension:0.3, pointRadius:4 }
      ]
    },
    options:{
      responsive:true, maintainAspectRatio:false,
      plugins:{
        legend:{ display:true, position:'bottom', labels:{boxWidth:10,font:{size:9}} },
        tooltip:{ callbacks:{ label:ctx=>ctx.dataset.label+': '+ctx.parsed.y.toLocaleString('es-ES',{minimumFractionDigits:2,maximumFractionDigits:2})+'€' } }
      },
      scales:{
        y:{ ticks:{ callback:v=>{const n=Math.abs(v);const s=Math.round(n).toString().replace(/\B(?=(\d{3})+(?!\d))/g,'.');return (v<0?'-':'')+s+'€';} } }
      }
    }
  });
}

// ── RENDIMIENTO ──────────────────────────────────────────────────────────────

let rendCharts = {};

function renderRendimiento() {
  const wrap = document.getElementById('rendimiento-wrap');
  if (!wrap) return;
  if (!costesProduccion.length) { wrap.style.display = 'none'; return; }

  const mesDesde = parseInt(document.getElementById('costes-mes-desde').value);
  const mesHasta = parseInt(document.getElementById('costes-mes-hasta').value);
  const meses = [];
  for (let m = mesDesde; m <= mesHasta; m++) meses.push(m);

  // Agregar producción por mes
  const prodMes = {}; // { mes: { tn, horas, t04, t412, t1220, t2040 } }
  for (const p of costesProduccion) {
    const m = parseInt(p.fecha.split('-')[1]);
    if (m < mesDesde || m > mesHasta) continue;
    if (!prodMes[m]) prodMes[m] = { tn: 0, horas: 0, t04: 0, t412: 0, t1220: 0, t2040: 0 };
    prodMes[m].tn    += parseFloat(p.tnDia)     || 0;
    prodMes[m].horas += parseFloat(p.horasPlanta)|| 0;
    prodMes[m].t04   += parseFloat(p.t04)        || 0;
    prodMes[m].t412  += parseFloat(p.t412)        || 0;
    prodMes[m].t1220 += parseFloat(p.t1220)       || 0;
    prodMes[m].t2040 += parseFloat(p.t2040)       || 0;
  }

  const fmtN = (v, d=1) => v ? Number(v).toLocaleString('es-ES', { minimumFractionDigits: d, maximumFractionDigits: d }) : '—';

  // ── KPIs ──
  const totalTn    = meses.reduce((s, m) => s + (prodMes[m]?.tn    || 0), 0);
  const totalHoras = meses.reduce((s, m) => s + (prodMes[m]?.horas || 0), 0);
  const eficMedia  = totalHoras ? totalTn / totalHoras : 0;
  const mesesConDatos = meses.filter(m => prodMes[m]?.tn > 0);
  const mejorMes   = mesesConDatos.reduce((best, m) => (prodMes[m].tn > (prodMes[best]?.tn || 0) ? m : best), mesesConDatos[0]);

  const kpiBox = (label, value, unit = '') => `
    <div style="background:var(--surface2);border-radius:8px;padding:12px 14px;border:1px solid var(--border)">
      <div style="color:var(--muted);font-size:.68rem;text-transform:uppercase;letter-spacing:.05em;margin-bottom:4px">${label}</div>
      <div style="font-size:1.15rem;font-weight:800;color:var(--text)">${value}<span style="font-size:.72rem;font-weight:400;color:var(--muted);margin-left:3px">${unit}</span></div>
    </div>`;

  document.getElementById('rend-kpis').innerHTML =
    kpiBox('Total producido', fmtN(totalTn, 0), 'tn') +
    kpiBox('Horas planta', fmtN(totalHoras, 1), 'h') +
    kpiBox('Eficiencia media', fmtN(eficMedia, 2), 'tn/h') +
    (mejorMes ? kpiBox('Mejor mes', MESES_NOMBRE[mejorMes] + ' · ' + fmtN(prodMes[mejorMes].tn, 0), 'tn') : '');

  // ── Tabla mensual ──
  const MESES_NOMBRE_CORTO = { 1:'Ene',2:'Feb',3:'Mar',4:'Abr',5:'May',6:'Jun',7:'Jul',8:'Ago',9:'Sep',10:'Oct',11:'Nov',12:'Dic' };
  const thStyle = 'padding:5px 8px;background:var(--surface2);font-weight:700;font-size:.72rem;text-align:right;border:1px solid var(--border)';
  const thL = thStyle.replace('text-align:right', 'text-align:left');
  const td = (v, bold = false) => `<td style="padding:4px 8px;border:1px solid var(--border);text-align:right;font-size:.75rem${bold ? ';font-weight:700' : ''}">${v}</td>`;

  let tbl = `<thead><tr>
    <th style="${thL}">Mes</th>
    <th style="${thStyle}">Tn Total</th>
    <th style="${thStyle}">0/4</th>
    <th style="${thStyle}">4/12</th>
    <th style="${thStyle}">12/20</th>
    <th style="${thStyle}">20/40</th>
    <th style="${thStyle}">Horas</th>
    <th style="${thStyle}">Tn/h</th>
  </tr></thead><tbody>`;

  let sumTn=0, sumHoras=0, sumT04=0, sumT412=0, sumT1220=0, sumT2040=0;
  for (const m of meses) {
    const d = prodMes[m];
    if (!d) { tbl += `<tr><td style="padding:4px 8px;border:1px solid var(--border);font-size:.75rem">${MESES_NOMBRE[m]}</td>${td('—')}${td('—')}${td('—')}${td('—')}${td('—')}${td('—')}${td('—')}</tr>`; continue; }
    const efic = d.horas ? (d.tn / d.horas) : 0;
    sumTn += d.tn; sumHoras += d.horas; sumT04 += d.t04; sumT412 += d.t412; sumT1220 += d.t1220; sumT2040 += d.t2040;
    tbl += `<tr>
      <td style="padding:4px 8px;border:1px solid var(--border);font-size:.75rem;font-weight:600">${MESES_NOMBRE[m]}</td>
      ${td(fmtN(d.tn, 0), true)}${td(fmtN(d.t04, 0))}${td(fmtN(d.t412, 0))}${td(fmtN(d.t1220, 0))}${td(fmtN(d.t2040, 0))}${td(fmtN(d.horas, 1))}${td(fmtN(efic, 2))}
    </tr>`;
  }
  // Total row
  const eficTotal = sumHoras ? sumTn / sumHoras : 0;
  tbl += `<tr style="background:rgba(107,125,46,.08);font-weight:900;font-size:.76rem">
    <td style="padding:5px 8px;border:1px solid var(--border);border-top:3px double var(--accent)">TOTAL</td>
    ${td(fmtN(sumTn, 0), true)}${td(fmtN(sumT04, 0))}${td(fmtN(sumT412, 0))}${td(fmtN(sumT1220, 0))}${td(fmtN(sumT2040, 0))}${td(fmtN(sumHoras, 1))}${td(fmtN(eficTotal, 2))}
  </tr>`;
  tbl += '</tbody>';
  document.getElementById('rend-tabla').innerHTML = tbl;

  // ── Gráficos ──
  Object.values(rendCharts).forEach(c => { try { c.destroy(); } catch(e) {} });
  rendCharts = {};
  if (typeof Chart === 'undefined') { wrap.style.display = ''; return; }

  const labels = meses.map(m => MESES_NOMBRE[m].substring(0, 3));

  // Barras apiladas por producto
  rendCharts.tn = new Chart(document.getElementById('chart-rend-tn'), {
    type: 'bar',
    data: {
      labels,
      datasets: [
        { label: '0/4',   data: meses.map(m => prodMes[m]?.t04   || 0), backgroundColor: '#6b7d2e', borderWidth: 0 },
        { label: '4/12',  data: meses.map(m => prodMes[m]?.t412  || 0), backgroundColor: '#d4a017', borderWidth: 0 },
        { label: '12/20', data: meses.map(m => prodMes[m]?.t1220 || 0), backgroundColor: '#2e7d6b', borderWidth: 0 },
        { label: '20/40', data: meses.map(m => prodMes[m]?.t2040 || 0), backgroundColor: '#4682b4', borderWidth: 0 },
      ]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { display: true, position: 'bottom', labels: { boxWidth: 10, font: { size: 9 } } },
        tooltip: { callbacks: { label: ctx => ctx.dataset.label + ': ' + ctx.parsed.y.toLocaleString('es-ES') + ' tn' } }
      },
      scales: { x: { stacked: true }, y: { stacked: true, ticks: { callback: v => v.toLocaleString('es-ES') + ' tn' } } }
    }
  });

  // Línea tn/hora
  rendCharts.efic = new Chart(document.getElementById('chart-rend-efic'), {
    type: 'line',
    data: {
      labels,
      datasets: [{
        label: 'Tn/hora',
        data: meses.map(m => { const d = prodMes[m]; return d?.horas ? +(d.tn / d.horas).toFixed(2) : 0; }),
        borderColor: '#6b7d2e', backgroundColor: 'rgba(107,125,46,0.1)', fill: true,
        borderWidth: 2, tension: 0.3, pointRadius: 4
      }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: ctx => ctx.parsed.y.toFixed(2) + ' tn/h' } }
      },
      scales: { y: { beginAtZero: true, ticks: { callback: v => v.toFixed(1) + ' tn/h' } } }
    }
  });

  wrap.style.display = '';
}

// ── COSTES IMPRIMIR ──────────────────────────────────────────────────────────

function imprimirCostes(){
  if(!costesRawData.length){ alert('Carga datos primero'); return; }

  const wrap = document.getElementById('costes-print-wrap');
  const anyo = document.getElementById('costes-anyo').value;
  const mesDesde = parseInt(document.getElementById('costes-mes-desde').value);
  const mesHasta = parseInt(document.getElementById('costes-mes-hasta').value);
  const vista = document.getElementById('costes-vista').value;

  // Header
  document.getElementById('costes-print-header').innerHTML = `
    <div style="display:flex;justify-content:space-between;align-items:flex-end;margin-bottom:8mm;border-bottom:2px solid #333;padding-bottom:4mm">
      <div>
        <div style="font-size:14pt;font-weight:900;letter-spacing:.04em">ARIFOMA</div>
        <div style="font-size:8pt;color:#666">Áridos Fonolíticos de Maspalomas SL</div>
      </div>
      <div style="text-align:right">
        <div style="font-size:11pt;font-weight:700">${vista==='acumulado'?'ANÁLISIS ACUMULADO':'ANÁLISIS MENSUAL DE COSTES'}</div>
        <div style="font-size:9pt;color:#666">${MESES_NOMBRE[mesDesde]} — ${MESES_NOMBRE[mesHasta]} ${anyo}</div>
      </div>
    </div>`;

  // Copy table from costes-table-wrap
  const tableHtml = document.getElementById('costes-table-wrap').innerHTML;
  document.getElementById('costes-print-body').innerHTML = `
    <div class="costes-print-table" style="font-family:Arial,sans-serif;font-size:8pt;color:#111">
      ${tableHtml}
    </div>`;

  // Override table styles for print
  const printTable = document.querySelector('#costes-print-body .costes-tbl');
  if(printTable){
    printTable.style.fontSize = '7.5pt';
    printTable.style.color = '#111';
    printTable.querySelectorAll('th,td').forEach(c=>{
      c.style.border = '1px solid #bbb';
      c.style.padding = '2px 5px';
      c.style.color = '#111';
    });
    printTable.querySelectorAll('.costes-cat-row td').forEach(c=>{
      c.style.background = '#e8e8e0';
      c.style.fontWeight = '700';
    });
    printTable.querySelectorAll('.costes-grand-row td').forEach(c=>{
      c.style.background = '#d4d4c8';
      c.style.fontWeight = '900';
    });
    printTable.querySelectorAll('.costes-ingreso td').forEach(c=>{
      c.style.color = '#2e7d32';
    });
    printTable.querySelectorAll('.costes-tn').forEach(c=>{
      c.style.color = '#888';
      c.style.fontSize = '6.5pt';
    });
  }

  // Charts as images
  const chartsArea = document.getElementById('costes-print-charts-area');
  chartsArea.innerHTML = '<div style="font-size:10pt;font-weight:700;margin-bottom:4mm">GRÁFICOS</div>';
  const chartIds = ['chart-barras-cat','chart-linea-tn','chart-donut','chart-ing-gas','chart-acumulado'];
  const chartNames = ['Costes por categoría / mes','€/tn mensual','Distribución de costes','Ingresos vs Gastos / mes','Acumulado Ingresos vs Gastos'];
  const grid = document.createElement('div');
  grid.style.cssText = 'display:grid;grid-template-columns:1fr 1fr;gap:8mm';
  chartIds.forEach((id,i)=>{
    const canvas = document.getElementById(id);
    if(!canvas) return;
    const div = document.createElement('div');
    div.innerHTML = `<div style="font-size:7.5pt;font-weight:700;margin-bottom:2mm">${chartNames[i]}</div>`;
    const img = document.createElement('img');
    img.src = canvas.toDataURL('image/png');
    img.style.cssText = 'width:100%;max-height:200px;object-fit:contain';
    div.appendChild(img);
    grid.appendChild(div);
  });
  chartsArea.appendChild(grid);

  wrap.style.display = 'block';
}

function cerrarImprimirCostes(){
  document.getElementById('costes-print-wrap').style.display = 'none';
}

// ── HISTÓRICO DE VENTAS — Facturas registradas + estado pago desde BC ────────
let hvData = [];
let hvLoaded = false;
let hvSelectedClientes = new Set(); // vacío = todos

function initHistoricoVentas() {
  if (!hvLoaded) {
    // Defaults: último año
    const hoy = new Date();
    const desde = new Date(hoy.getFullYear(), 0, 1);
    document.getElementById('hv-fecha-desde').value = desde.toISOString().slice(0, 10);
    document.getElementById('hv-fecha-hasta').value = hoy.toISOString().slice(0, 10);
  }
  if (hvData.length === 0) cargarHistoricoVentas();
  else renderHistoricoVentas();
}

async function cargarHistoricoVentas() {
  const tbody = document.getElementById('hv-tbody');
  tbody.innerHTML = '<tr><td colspan="9" style="padding:30px;text-align:center;color:var(--muted)">Cargando facturas de BC...</td></tr>';
  document.getElementById('hv-alertas').style.display = 'none';
  document.getElementById('hv-kpis').innerHTML = '';

  try {
    const token = await getBCToken();
    const fechaDesde = document.getElementById('hv-fecha-desde').value || undefined;
    const fechaHasta = document.getElementById('hv-fecha-hasta').value || undefined;

    const res = await fetch('/api/bc/historico-ventas', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ token, fechaDesde, fechaHasta })
    });

    const json = await res.json();
    if (!json.ok) throw new Error(json.error || 'Error cargando facturas');

    hvData = json.data || [];
    hvLoaded = true;
    hvSelectedClientes = new Set();
    poblarHvClientes();
    renderHistoricoVentas();
    notificarFacturasVencidas();
  } catch (e) {
    console.error('Histórico ventas error:', e);
    tbody.innerHTML = `<tr><td colspan="9" style="padding:30px;text-align:center;color:#c62828">${escapeHTML(e.message)}</td></tr>`;
  }
}

function renderHistoricoVentas() {
  const buscar = (document.getElementById('hv-buscar').value || '').toLowerCase().trim();
  const estado = document.getElementById('hv-estado').value;
  const desde = document.getElementById('hv-fecha-desde').value;
  const hasta = document.getElementById('hv-fecha-hasta').value;

  let filtered = hvData.filter(f => {
    if (estado) {
      if (estado === 'vencida') { if (f.estado !== 'vencida') return false; }
      else if (estado === 'vencida30') { if (f.estado !== 'vencida' || f.diasVencido < 30) return false; }
      else if (estado === 'vencida60') { if (f.estado !== 'vencida' || f.diasVencido < 60) return false; }
      else if (estado === 'vencida90') { if (f.estado !== 'vencida' || f.diasVencido < 90) return false; }
      else { if (f.estado !== estado) return false; }
    }
    if (desde && f.fecha < desde) return false;
    if (hasta && f.fecha > hasta) return false;
    if (buscar) {
      const hay = (f.numero + ' ' + f.clienteCod).toLowerCase().includes(buscar);
      if (!hay) return false;
    }
    // Filtro multi-cliente
    if (hvSelectedClientes.size > 0 && !hvSelectedClientes.has(f.clienteNombre)) return false;
    return true;
  });

  renderHvClientesChips();

  // KPIs
  renderHvKpis(filtered);

  // Alertas vencidas
  renderHvAlertas(filtered);

  // Aging
  renderHvAging(filtered);

  // Tabla
  const tbody = document.getElementById('hv-tbody');
  if (filtered.length === 0) {
    tbody.innerHTML = '<tr><td colspan="9" style="padding:20px;text-align:center;color:var(--muted)">Sin facturas para los filtros seleccionados</td></tr>';
    return;
  }

  // Ordenar por fecha de registro (más reciente primero)
  filtered.sort((a, b) => b.fecha.localeCompare(a.fecha));

  window._hvFiltered = filtered;
  tbody.innerHTML = filtered.map((f, idx) => {
    const badge = hvBadge(f.estado);
    const diasTxt = f.estado === 'vencida' && f.diasVencido > 0
      ? `<span style="color:#c62828;font-weight:700">${f.diasVencido}d</span>`
      : (f.estado === 'pendiente' && f.vencimiento ? diasHastaVenc(f.vencimiento) : '—');
    const rowBg = f.estado === 'vencida' ? 'background:rgba(198,40,40,.06)' : '';
    const reminderBtn = f.estado !== 'pagada'
      ? `<td style="padding:8px 6px;border-bottom:1px solid var(--border);text-align:center"><button data-hv-idx="${idx}" onclick="hvEnviarRecordatorio(this.dataset.hvIdx)" style="background:none;border:none;cursor:pointer;font-size:1.1rem;padding:2px 6px;border-radius:6px;transition:background .15s" title="Enviar recordatorio de cobro">🔔</button></td>`
      : `<td style="padding:8px 6px;border-bottom:1px solid var(--border);text-align:center"></td>`;
    return `<tr style="${rowBg}">
      <td style="padding:8px 12px;border-bottom:1px solid var(--border);font-weight:600">${escapeHTML(f.numero)}</td>
      <td style="padding:8px 12px;border-bottom:1px solid var(--border)">${fmtDateISO(f.fecha)}</td>
      <td style="padding:8px 12px;border-bottom:1px solid var(--border)">${f.vencimiento ? fmtDateISO(f.vencimiento) : '—'}</td>
      <td style="padding:8px 12px;border-bottom:1px solid var(--border)">${escapeHTML(f.clienteNombre)}</td>
      <td style="padding:8px 12px;border-bottom:1px solid var(--border);text-align:right;font-variant-numeric:tabular-nums">${fmtEur(f.importe)}</td>
      <td style="padding:8px 12px;border-bottom:1px solid var(--border);text-align:right;font-variant-numeric:tabular-nums;${f.pendiente > 0 ? 'color:#c62828;font-weight:600' : ''}">${fmtEur(f.pendiente)}</td>
      <td style="padding:8px 12px;border-bottom:1px solid var(--border);text-align:center">${badge}</td>
      <td style="padding:8px 12px;border-bottom:1px solid var(--border);text-align:center">${diasTxt}</td>
      ${reminderBtn}
    </tr>`;
  }).join('');
}

function renderHvKpis(data) {
  const total = data.length;
  const totalImporte = data.reduce((s, f) => s + f.importe, 0);
  const pendientes = data.filter(f => f.estado !== 'pagada');
  const totalPendiente = pendientes.reduce((s, f) => s + f.pendiente, 0);
  const vencidas = data.filter(f => f.estado === 'vencida');
  const totalVencido = vencidas.reduce((s, f) => s + f.pendiente, 0);

  document.getElementById('hv-kpis').innerHTML = `
    <div style="background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:14px;text-align:center">
      <div style="font-size:.68rem;font-weight:700;text-transform:uppercase;letter-spacing:.05em;color:var(--muted);margin-bottom:4px">Facturas</div>
      <div style="font-family:'DM Mono',monospace;font-size:1.3rem;font-weight:700;color:var(--text)">${total}</div>
    </div>
    <div style="background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:14px;text-align:center">
      <div style="font-size:.68rem;font-weight:700;text-transform:uppercase;letter-spacing:.05em;color:var(--muted);margin-bottom:4px">Facturado</div>
      <div style="font-family:'DM Mono',monospace;font-size:1.3rem;font-weight:700;color:var(--accent)">${fmtEur(totalImporte)}</div>
    </div>
    <div style="background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:14px;text-align:center">
      <div style="font-size:.68rem;font-weight:700;text-transform:uppercase;letter-spacing:.05em;color:var(--muted);margin-bottom:4px">Pendiente cobro</div>
      <div style="font-family:'DM Mono',monospace;font-size:1.3rem;font-weight:700;color:${totalPendiente > 0 ? '#e65100' : 'var(--text)'}">${fmtEur(totalPendiente)}</div>
    </div>
    <div style="background:${totalVencido > 0 ? 'rgba(198,40,40,.08)' : 'var(--surface)'};border:1px solid ${totalVencido > 0 ? 'rgba(198,40,40,.3)' : 'var(--border)'};border-radius:var(--radius);padding:14px;text-align:center">
      <div style="font-size:.68rem;font-weight:700;text-transform:uppercase;letter-spacing:.05em;color:${totalVencido > 0 ? '#c62828' : 'var(--muted)'};margin-bottom:4px">Vencido</div>
      <div style="font-family:'DM Mono',monospace;font-size:1.3rem;font-weight:700;color:${totalVencido > 0 ? '#c62828' : 'var(--text)'}">${fmtEur(totalVencido)}</div>
    </div>`;
}

function renderHvAlertas(data) {
  const vencidas = data.filter(f => f.estado === 'vencida').sort((a, b) => b.diasVencido - a.diasVencido);
  const el = document.getElementById('hv-alertas');

  if (vencidas.length === 0) {
    el.style.display = 'none';
    return;
  }

  // Agrupar por cliente
  const porCliente = {};
  for (const f of vencidas) {
    if (!porCliente[f.clienteNombre]) porCliente[f.clienteNombre] = { facturas: [], total: 0, maxDias: 0 };
    porCliente[f.clienteNombre].facturas.push(f);
    porCliente[f.clienteNombre].total += f.pendiente;
    porCliente[f.clienteNombre].maxDias = Math.max(porCliente[f.clienteNombre].maxDias, f.diasVencido);
  }

  const clientes = Object.entries(porCliente).sort((a, b) => b[1].maxDias - a[1].maxDias);

  el.style.display = 'block';
  el.innerHTML = `
    <div style="background:rgba(198,40,40,.08);border:1px solid rgba(198,40,40,.3);border-radius:var(--radius);padding:14px 16px">
      <div style="font-size:.82rem;font-weight:700;color:#c62828;margin-bottom:8px">⚠ ${vencidas.length} factura${vencidas.length > 1 ? 's' : ''} vencida${vencidas.length > 1 ? 's' : ''} — Recordatorios de cobro</div>
      ${clientes.map(([cli, info]) => `
        <div style="margin-bottom:6px;padding:8px 12px;background:rgba(255,255,255,.5);border-radius:6px;font-size:.78rem">
          <div style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:4px">
            <span style="font-weight:700;color:#b71c1c">${escapeHTML(cli)}</span>
            <span style="font-weight:600">${fmtEur(info.total)} · ${info.facturas.length} fact. · máx ${info.maxDias} días</span>
          </div>
          <div style="font-size:.72rem;color:var(--muted);margin-top:2px">
            ${info.facturas.map(f => `${escapeHTML(f.numero)} (${f.diasVencido}d · ${fmtEur(f.pendiente)})`).join(' · ')}
          </div>
        </div>
      `).join('')}
    </div>`;
}

function renderHvAging(data) {
  const pendientes = data.filter(f => f.estado !== 'pagada' && f.pendiente > 0);
  const agingEl = document.getElementById('hv-aging');
  const gridEl = document.getElementById('hv-aging-grid');

  if (pendientes.length === 0) {
    agingEl.style.display = 'none';
    return;
  }

  const hoy = new Date();
  const buckets = [
    { label: 'Al día', min: -Infinity, max: 0, color: '#2e7d32', bg: 'rgba(46,125,50,.08)', total: 0, count: 0 },
    { label: '1-30 días', min: 1, max: 30, color: '#e65100', bg: 'rgba(230,81,0,.08)', total: 0, count: 0 },
    { label: '31-60 días', min: 31, max: 60, color: '#c62828', bg: 'rgba(198,40,40,.08)', total: 0, count: 0 },
    { label: '61-90 días', min: 61, max: 90, color: '#b71c1c', bg: 'rgba(183,28,28,.12)', total: 0, count: 0 },
    { label: '>90 días', min: 91, max: Infinity, color: '#880e4f', bg: 'rgba(136,14,79,.12)', total: 0, count: 0 }
  ];

  for (const f of pendientes) {
    let dias = 0;
    if (f.vencimiento) {
      dias = Math.floor((hoy.getTime() - new Date(f.vencimiento).getTime()) / 86400000);
    }
    for (const b of buckets) {
      if (dias >= b.min && dias <= b.max) {
        b.total += f.pendiente;
        b.count++;
        break;
      }
    }
  }

  agingEl.style.display = 'block';
  gridEl.innerHTML = buckets.map(b => `
    <div style="background:${b.bg};border:1px solid ${b.color}33;border-radius:var(--radius);padding:12px;text-align:center">
      <div style="font-size:.68rem;font-weight:700;text-transform:uppercase;letter-spacing:.04em;color:${b.color};margin-bottom:4px">${b.label}</div>
      <div style="font-family:'DM Mono',monospace;font-size:1.1rem;font-weight:700;color:${b.color}">${fmtEur(b.total)}</div>
      <div style="font-size:.68rem;color:var(--muted);margin-top:2px">${b.count} factura${b.count !== 1 ? 's' : ''}</div>
    </div>
  `).join('');
}

function hvBadge(estado) {
  const map = {
    pagada: { bg: 'rgba(46,125,50,.12)', color: '#2e7d32', text: 'Pagada' },
    pendiente: { bg: 'rgba(230,81,0,.1)', color: '#e65100', text: 'Pendiente' },
    vencida: { bg: 'rgba(198,40,40,.12)', color: '#c62828', text: 'Vencida' },
    desconocido: { bg: 'var(--surface2)', color: 'var(--muted)', text: '—' }
  };
  const s = map[estado] || map.desconocido;
  return `<span style="display:inline-block;padding:3px 10px;border-radius:20px;font-size:.68rem;font-weight:700;background:${s.bg};color:${s.color}">${s.text}</span>`;
}

function fmtDateISO(d) {
  if (!d) return '—';
  const p = d.slice(0, 10).split('-');
  return p.length === 3 ? p[2] + '/' + p[1] + '/' + p[0] : d;
}

function fmtEur(n) {
  return new Intl.NumberFormat('es-ES', { style: 'currency', currency: 'EUR' }).format(n || 0);
}

function diasHastaVenc(venc) {
  const dias = Math.ceil((new Date(venc).getTime() - Date.now()) / 86400000);
  if (dias < 0) return `<span style="color:#c62828;font-weight:700">${Math.abs(dias)}d venc.</span>`;
  if (dias <= 7) return `<span style="color:#e65100;font-weight:600">${dias}d</span>`;
  return `<span style="color:var(--muted)">${dias}d</span>`;
}

// ── HISTÓRICO VENTAS — MULTI-SELECT CLIENTES ──────────────────────────────────
let _hvClientesList = []; // lista ordenada de nombres únicos

function poblarHvClientes() {
  const nombres = [...new Set(hvData.map(f => f.clienteNombre).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'es'));
  _hvClientesList = nombres;
  renderHvClientesList();
  updateHvClientesLabel();
}

let _hvDdOpen = false;

function toggleHvClientes(evt) {
  if (evt) evt.stopPropagation();
  const dd = document.getElementById('hv-clientes-dd');
  _hvDdOpen = !_hvDdOpen;
  dd.style.display = _hvDdOpen ? 'block' : 'none';
  if (_hvDdOpen) {
    document.getElementById('hv-clientes-search').value = '';
    renderHvClientesList();
    setTimeout(() => document.getElementById('hv-clientes-search').focus(), 50);
    document.addEventListener('mousedown', _hvClickOutside);
  } else {
    document.removeEventListener('mousedown', _hvClickOutside);
  }
}

function _hvClickOutside(e) {
  const dd = document.getElementById('hv-clientes-dd');
  const btn = document.getElementById('hv-clientes-btn');
  if (dd && !dd.contains(e.target) && btn && !btn.contains(e.target)) {
    dd.style.display = 'none';
    _hvDdOpen = false;
    document.removeEventListener('mousedown', _hvClickOutside);
  }
}

function filtrarHvClientes() {
  renderHvClientesList();
}

function renderHvClientesList() {
  const search = (document.getElementById('hv-clientes-search').value || '').toLowerCase().trim();
  const list = document.getElementById('hv-clientes-list');
  const filtered = search ? _hvClientesList.filter(n => n.toLowerCase().includes(search)) : _hvClientesList;

  // Build DOM instead of innerHTML to avoid quote escaping issues
  list.innerHTML = '';
  if (filtered.length === 0) {
    list.innerHTML = '<div style="padding:12px;text-align:center;color:var(--muted);font-size:.75rem">Sin resultados</div>';
    return;
  }

  for (const name of filtered) {
    const checked = hvSelectedClientes.size === 0 || hvSelectedClientes.has(name);
    const count = hvData.filter(f => f.clienteNombre === name).length;

    const label = document.createElement('label');
    label.style.cssText = 'display:flex;align-items:center;gap:8px;padding:6px 10px;cursor:pointer;font-size:.75rem;transition:background .1s;border-bottom:1px solid var(--border)';
    label.onmouseover = function() { this.style.background = 'var(--surface2)'; };
    label.onmouseout = function() { this.style.background = ''; };

    const cb = document.createElement('input');
    cb.type = 'checkbox';
    cb.checked = checked;
    cb.style.cssText = 'accent-color:var(--accent2);flex-shrink:0';
    cb.addEventListener('change', function() { hvToggleCliente(name, this.checked); });

    const span = document.createElement('span');
    span.style.cssText = 'flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap';
    span.textContent = name;

    const cnt = document.createElement('span');
    cnt.style.cssText = 'font-size:.65rem;color:var(--muted);flex-shrink:0';
    cnt.textContent = count;

    label.appendChild(cb);
    label.appendChild(span);
    label.appendChild(cnt);
    list.appendChild(label);
  }
}

function hvToggleCliente(name, checked) {
  if (hvSelectedClientes.size === 0 && checked) {
    // Transitioning from "all" → need to add all except unchecked
    // Actually if size=0 means all are selected, and user unchecks one:
    // We need to handle this differently
  }

  if (hvSelectedClientes.size === 0) {
    // "Todos" mode — user unchecked one, so select all EXCEPT that one
    if (!checked) {
      for (const n of _hvClientesList) hvSelectedClientes.add(n);
      hvSelectedClientes.delete(name);
    }
    // If checked while in "all" mode, nothing to do
  } else {
    if (checked) hvSelectedClientes.add(name);
    else hvSelectedClientes.delete(name);
    // Si todos seleccionados, volver a modo "todos" (set vacío)
    if (hvSelectedClientes.size === _hvClientesList.length) hvSelectedClientes.clear();
  }

  updateHvClientesLabel();
  renderHistoricoVentas();
}

function hvClientesSelectAll() {
  hvSelectedClientes.clear();
  renderHvClientesList();
  updateHvClientesLabel();
  renderHistoricoVentas();
}

function hvClientesSelectNone() {
  hvSelectedClientes.clear();
  hvSelectedClientes.add('__none__'); // sentinel para que no matchee nada
  renderHvClientesList();
  updateHvClientesLabel();
  renderHistoricoVentas();
}

function updateHvClientesLabel() {
  const label = document.getElementById('hv-clientes-label');
  if (hvSelectedClientes.size === 0) {
    label.textContent = 'Todos los clientes';
    label.style.color = '';
  } else if (hvSelectedClientes.has('__none__')) {
    label.textContent = 'Ningún cliente';
    label.style.color = 'var(--muted)';
  } else {
    const n = hvSelectedClientes.size;
    label.textContent = n === 1 ? [...hvSelectedClientes][0] : `${n} clientes`;
    label.style.color = 'var(--accent2)';
  }
}

function renderHvClientesChips() {
  const el = document.getElementById('hv-clientes-chips');
  if (!el) return;
  if (hvSelectedClientes.size === 0 || hvSelectedClientes.has('__none__')) {
    el.style.display = 'none';
    return;
  }
  el.style.display = 'flex';
  el.innerHTML = '';
  for (const name of [...hvSelectedClientes].sort((a, b) => a.localeCompare(b, 'es'))) {
    const chip = document.createElement('span');
    chip.style.cssText = 'display:inline-flex;align-items:center;gap:4px;padding:3px 10px 3px 8px;background:rgba(107,125,46,.1);border:1px solid rgba(107,125,46,.25);border-radius:20px;font-size:.68rem;font-weight:600;color:var(--accent2)';
    chip.textContent = name;
    const x = document.createElement('span');
    x.style.cssText = 'cursor:pointer;font-size:.8rem;line-height:1;opacity:.6;margin-left:2px';
    x.textContent = '\u00d7';
    x.addEventListener('click', () => hvRemoveCliente(name));
    chip.appendChild(x);
    el.appendChild(chip);
  }
}

function hvRemoveCliente(name) {
  hvSelectedClientes.delete(name);
  if (hvSelectedClientes.size === 0) hvSelectedClientes.clear(); // back to "all"
  renderHvClientesList();
  updateHvClientesLabel();
  renderHistoricoVentas();
}

// ── RECORDATORIO COBRO POR EMAIL ──────────────────────────────────────────────
function hvEnviarRecordatorio(idx) {
  const f = window._hvFiltered[idx];
  if (!f) return;

  const email = f.clienteEmail || '';
  const asunto = `Recordatorio de pago - Factura ${f.numero}`;
  const vencTxt = f.vencimiento ? fmtDateISO(f.vencimiento) : 'no especificada';
  const diasTxt = f.diasVencido > 0 ? ` (${f.diasVencido} días de retraso)` : '';

  const cuerpo = `Estimado/a cliente,

Le escribimos desde ARIFOMA para recordarle que la siguiente factura se encuentra pendiente de pago:

  - Nº Factura: ${f.numero}
  - Fecha emisión: ${fmtDateISO(f.fecha)}
  - Fecha vencimiento: ${vencTxt}${diasTxt}
  - Importe pendiente: ${fmtEur(f.pendiente)}

Le agradeceríamos que nos indicase cuándo podemos esperar recibir el pago, o si existe alguna incidencia con esta factura.

Quedamos a su disposición para cualquier consulta.

Un saludo,`;

  // Abrir en Outlook web — URL directa que no redirige a app de escritorio
  const outlookUrl = 'https://outlook.office365.com/owa/?path=/mail/action/compose&to=' + encodeURIComponent(email)
    + '&subject=' + encodeURIComponent(asunto)
    + '&body=' + encodeURIComponent(cuerpo);

  window.open(outlookUrl, '_blank');
}

// ── NOTIFICACIONES FACTURAS VENCIDAS ──────────────────────────────────────────
function notificarFacturasVencidas() {
  const vencidas = hvData.filter(f => f.estado === 'vencida');
  if (!vencidas.length) return;

  const totalPendiente = vencidas.reduce((s, f) => s + f.pendiente, 0);
  const maxDias = Math.max(...vencidas.map(f => f.diasVencido));

  // Agrupar por cliente
  const clientes = {};
  for (const f of vencidas) {
    if (!clientes[f.clienteNombre]) clientes[f.clienteNombre] = { count: 0, total: 0 };
    clientes[f.clienteNombre].count++;
    clientes[f.clienteNombre].total += f.pendiente;
  }
  const topClientes = Object.entries(clientes).sort((a, b) => b[1].total - a[1].total).slice(0, 3);

  const lines = [`${vencidas.length} factura${vencidas.length > 1 ? 's' : ''} vencida${vencidas.length > 1 ? 's' : ''} · ${fmtEur(totalPendiente)}`];
  if (maxDias > 30) lines.push(`Máximo retraso: ${maxDias} días`);
  for (const [cli, info] of topClientes) {
    lines.push(`${cli}: ${info.count} fact. · ${fmtEur(info.total)}`);
  }
  const msg = lines.join('\n');

  // Browser Notification
  if ('Notification' in window) {
    if (Notification.permission === 'granted') {
      new Notification('Facturas vencidas', { body: msg, icon: '', tag: 'facturas-vencidas' });
    } else if (Notification.permission !== 'denied') {
      Notification.requestPermission().then(p => {
        if (p === 'granted') new Notification('Facturas vencidas', { body: msg, icon: '', tag: 'facturas-vencidas' });
      });
    }
  }

  // In-app alert en inicio si está visible
  const inicioAlert = document.getElementById('inicio-facturas-alert');
  if (inicioAlert) {
    inicioAlert.style.display = 'block';
    inicioAlert.innerHTML = `
      <div style="background:rgba(198,40,40,.08);border:1px solid rgba(198,40,40,.3);border-radius:var(--radius);padding:14px 16px;cursor:pointer" onclick="goPage('historico-ventas');document.getElementById('hv-estado').value='vencida';renderHistoricoVentas()">
        <div style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px">
          <div>
            <div style="font-size:.82rem;font-weight:700;color:#c62828">⚠ ${vencidas.length} factura${vencidas.length > 1 ? 's' : ''} vencida${vencidas.length > 1 ? 's' : ''}</div>
            <div style="font-size:.72rem;color:var(--muted);margin-top:2px">${fmtEur(totalPendiente)} pendiente · máx ${maxDias} días</div>
          </div>
          <div style="font-size:.72rem;color:var(--accent2);font-weight:600">Ver detalle →</div>
        </div>
      </div>`;
  }
}

// Comprobar facturas vencidas en segundo plano al cargar la app
async function checkFacturasVencidasBackground() {
  try {
    const token = await getBCTokenSilent();
    // Solo últimos 6 meses
    const desde = new Date();
    desde.setMonth(desde.getMonth() - 6);
    const res = await fetch('/api/bc/historico-ventas', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ token, fechaDesde: desde.toISOString().slice(0, 10) })
    });
    const json = await res.json();
    if (json.ok) {
      hvData = json.data || [];
      hvLoaded = true;
      notificarFacturasVencidas();
    }
  } catch (e) {
    console.warn('Check facturas vencidas background:', e.message);
  }
}

// ── NOTAS DEL SISTEMA ────────────────────────────────────────────────────────
let notasCache = [];

async function cargarNotas() {
  const result = await dbQuery({ action: 'select', table: 'tblnotas', options: { select: '*', order: 'created_at.desc' } });
  if (result.ok) notasCache = result.data || [];
  actualizarBadgeNotas();
}

function actualizarBadgeNotas() {
  const badge = document.getElementById('notas-badge');
  if (!badge) return;
  const n = notasCache.length;
  if (n > 0) {
    badge.textContent = n > 99 ? '99+' : n;
    badge.style.display = 'inline-flex';
  } else {
    badge.style.display = 'none';
  }
}

function abrirModalNotas() {
  const modal = document.getElementById('modal-notas');
  modal.style.display = 'flex';
  // Poblar dropdown de páginas
  const selFiltro = document.getElementById('notas-filtro-pagina');
  const selForm = document.getElementById('nota-nueva-pagina');
  const currentVal = selFiltro.value;
  selFiltro.innerHTML = '<option value="">Todas</option>';
  selForm.innerHTML = '';
  const entries = Object.entries(PAGE_TITLES).sort((a,b)=>a[1].localeCompare(b[1]));
  entries.forEach(([k,v]) => {
    selFiltro.innerHTML += `<option value="${k}">${v}</option>`;
    selForm.innerHTML += `<option value="${k}">${v}</option>`;
  });
  selFiltro.value = currentVal;
  // Pre-seleccionar página actual en el form
  const currentPage = document.querySelector('.page.active')?.id?.replace('pg-','') || '';
  if (currentPage && selForm.querySelector(`option[value="${currentPage}"]`)) {
    selForm.value = currentPage;
  }
  cargarNotas().then(() => renderNotas());
}

function cerrarModalNotas() {
  document.getElementById('modal-notas').style.display = 'none';
  cerrarFormNota();
}

function renderNotas() {
  const lista = document.getElementById('notas-lista');
  const filtPag = document.getElementById('notas-filtro-pagina').value;
  const filtTipo = document.getElementById('notas-filtro-tipo').value;
  let filtered = notasCache;
  if (filtPag) filtered = filtered.filter(n => n.pagina === filtPag);
  if (filtTipo) filtered = filtered.filter(n => n.tipo === filtTipo);

  if (!filtered.length) {
    lista.innerHTML = '<div style="color:var(--muted);text-align:center;padding:30px;font-size:.82rem">Sin notas. Pulsa "+ Nueva nota" para añadir.</div>';
    return;
  }

  const tipoCfg = {
    error:  { color:'#c0392b', bg:'rgba(192,57,43,.1)',  label:'Error'  },
    mejora: { color:'#d4a017', bg:'rgba(212,160,23,.1)', label:'Mejora' },
    nota:   { color:'#2e4a7d', bg:'rgba(46,74,125,.1)',  label:'Nota'   }
  };

  // Agrupar por página
  const grupos = {};
  filtered.forEach(n => {
    const pg = n.pagina || 'general';
    if (!grupos[pg]) grupos[pg] = [];
    grupos[pg].push(n);
  });

  let html = '';
  Object.entries(grupos).forEach(([pg, items]) => {
    const pgTitle = PAGE_TITLES[pg] || pg;
    html += `<div style="margin-bottom:16px"><div style="font-size:.7rem;font-weight:700;text-transform:uppercase;letter-spacing:.08em;color:var(--muted);margin-bottom:6px;padding-bottom:4px;border-bottom:1px solid var(--border)">${pgTitle}</div>`;
    items.forEach(n => {
      const cfg = tipoCfg[n.tipo] || tipoCfg.nota;
      const fecha = n.created_at ? new Date(n.created_at).toLocaleDateString('es-ES',{day:'2-digit',month:'2-digit',year:'2-digit',hour:'2-digit',minute:'2-digit'}) : '';
      html += `<div style="background:var(--surface);border:1px solid var(--border);border-left:3px solid ${cfg.color};border-radius:6px;padding:10px 12px;margin-bottom:8px">
        <div style="display:flex;align-items:center;gap:8px;margin-bottom:6px">
          <span style="background:${cfg.bg};color:${cfg.color};border-radius:4px;padding:2px 8px;font-size:.65rem;font-weight:700;text-transform:uppercase">${cfg.label}</span>
          <span style="font-size:.68rem;color:var(--muted)">${fecha}</span>
          <button onclick="eliminarNota('${n.id}')" style="margin-left:auto;background:none;border:none;color:var(--muted);cursor:pointer;font-size:1rem;line-height:1;padding:0 4px" title="Eliminar">&times;</button>
        </div>
        <div style="font-size:.82rem;color:var(--text);white-space:pre-wrap;line-height:1.5">${escapeHTML(n.texto||'')}</div>
      </div>`;
    });
    html += '</div>';
  });
  lista.innerHTML = html;
}

function abrirFormNota() {
  document.getElementById('notas-form').style.display = 'block';
  document.getElementById('nota-nueva-texto').focus();
}

function cerrarFormNota() {
  document.getElementById('notas-form').style.display = 'none';
  document.getElementById('nota-nueva-texto').value = '';
}

async function guardarNota() {
  const pagina = document.getElementById('nota-nueva-pagina').value;
  const tipo = document.getElementById('nota-nueva-tipo').value;
  const texto = document.getElementById('nota-nueva-texto').value.trim();
  if (!texto) { alert('Escribe el texto de la nota'); return; }

  const notaRes = await dbQuery({ action: 'insert', table: 'tblnotas', data: { pagina, tipo, texto } });
  if (!notaRes.ok) { alert('Error guardando nota: ' + notaRes.error); return; }

  cerrarFormNota();
  await cargarNotas();
  renderNotas();
}

async function eliminarNota(id) {
  if (!confirm('¿Eliminar esta nota?')) return;
  const delRes = await dbQuery({ action: 'delete', table: 'tblnotas', filters: [{ column: 'id', op: 'eq', value: id }] });
  if (!delRes.ok) { alert('Error: ' + delRes.error); return; }
  await cargarNotas();
  renderNotas();
}

// ── COMPRAS — ESCANEAR FACTURA Y SUBIR A ONEDRIVE ───────────
const COMPRAS_CLIENT_ID='20d8ca37-34e7-4ad4-b379-97c5b22f15ad';
const COMPRAS_TENANT_ID='5bd828f2-1899-48ba-a269-c37733f41806';
const COMPRAS_REDIRECT=location.origin+location.pathname;
const COMPRAS_SCOPES=['Files.ReadWrite.All'];
const COMPRAS_ONEDRIVE_BASE='06. ADMINISTRACION/06.01 PROVEEDORES';
const COMPRAS_SHARE_URL='https://grpsite-my.sharepoint.com/:f:/g/personal/greyes_arifoma_com/IgD8XOuwUpjWQ4E17TuO5-PoAWbx8HnqElIXhD2fQerh_QM';
const COMPRAS_MESES=['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO','JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE'];
const COMPRAS_PROVEEDORES=[
  '(ANEFA) ASOCIACION NACIONAL DE EMPRESARIOS FABRICANTES DE ARIDO','AENOR','AGONEY LUJAN PEREZ','AGUAS DE GUAYADEQUE SL','ALIANZA ALEMAN BLAKER',
  'APPLUS ITEUVE TECHNOLOGY, S.L.U','ARUESCABLE, S.L','ASESORIA RUPERTO PEREZ S.L.U','ASUAREZ MOTOR SERVICES SL','ATLANTIC CANARIAS',
  'AUTOS BASSO SA','AUTOS TRAG ALAMO SL','AVANTA PREVENCIÓN INTEGRAL S.L','AYTO San Bartolome','Aab','Aguiar Marrero','Ascanio',
  'B2Brouter GLOBAL S.L','BALCAN','Blumaq','CAMPSA ESTACIONES DE SERVICIO','CANARIAS DE TRANSPORTES ENTRE ISLAS 2020 SLU',
  'CANARIAS EXPLOSIVOS, SA','CENTAURO','CENTRAL UNIFORMES SL','CONTROLES Y ACCIONAMIENTOS CANARIOS SL','Canarias Beton','Crédito y Caución',
  'Diasan','Disa','E.S. JUAN GRANDE','EQUIPOS Y SERVICIOS HIDRAULICOS CANARIOS SL','ESOCAN SL','EXTINTORES CONTRAINCENDIOS CANARIOS',
  'Electrimega','Elevaciones Archipielago','Elmasa','Energy Power','Eteicomps','Euro Scrymo','FINANZAUTO S.A.U-CENTRAL',
  'FRANCISCO JAVIER VERA MARTÍN','GRAN CHOLLOS CANARIAS SLU','HARRUCASUNICO SL','HOTEL LA ERMITA','HROS. DE JOSE SUAREZ LOPEZ',
  'Hernandez Consultores','Hierros 7 islas','Hispano Japonesa','IGLESIAS FARRE ROS, SAU','INSTALADORA SUAREZ SL','ITC 2023 SL',
  'ITT CANARIAS SL','ITV DE MAQUINARIA','Insucan','Jacob Perez','Jose caldera','Juan Melian SL','Kalon',
  'LABORATORIO DE CERTIFICACIONES VEGA BAJA','LAS ROSAS','LEROY MERLIN ESPAÑA SLU','Labcer','Labetec',
  'MANUEL OLIVERA RODRIGUEZ SL','MAQUINARIA Y SEGUROS OJEDA GRANADO SL','MAQUINARIAS OPEIN SLU','MARK JOHNSTON','MASPALOMAS REPUESTOS',
  'MATERIALES Y SAN','MEDIAMARKT','MERCURI PUBLIC','MICROSOFT IBERICA','MICROSOFT IRELAND OPERATIONS LTD','MIKEL URIARTE ATEKA',
  'Marecan','Masanes','Microrriego','Mopsa','Msm Rodamientos','NESTOR CUBAS DIAZ','NEUMADRAO','NEUMATICOS ATLANTICO','NORAUTO SAU',
  'NOVELEC VECAPE SL','Neumaticos TEIDESUR','OBRAMAT','OPERACIONES TURÍSTICAS CANARIASVIAJA, S.A','PRENSA DIGITAL CANARIA SL',
  'PRESTA SERVICIOS AMBIENTALES SL','PROPIEDADES MEJORADAS SL','PaTuMovil','QUIMICAS LASSO SLU','REPUESTOS Y REPRESENTACIONES',
  'ROCA GESTION HOSPITALARIA SL','RUIJIA SCP','Rafael Peinado','Recacor','Roeirasa','Ronandez',
  'SEGURIDAD INDUSTRIAL, MEDIO AMBIENTE Y CALIDAD, S. L','SPAR JUAN GRANDE','SUMINISTROS SANTANA DOMINGUEZ SA','Salazar','Santana Jerez',
  'Secular 2022','Securitas Direct','Sernamol','Sertego','Señal Canary','Sika','Sopranes','Suim',
  'TALLER DE MECANIZADO SALVADOR ORTEGA MORENO','TRANSPORTES JUAN ELEUTERIO MARTERL SLU','Tamaran','Telefonica','Transportes Sanchez',
  'Transportistas','Vallate','WURTH','ISMAEL GARCIA OJEDA'
];

let _msalInstance=null;
let _comprasFile=null;
let _comprasFileName='';
let _comprasFileBuffer=null;
let _comprasFileType='';
let _comprasVendorsBC=[];
let _comprasItemsBC=[];

// Artículo por defecto según proveedor (nombre BC exacto → número artículo BC)
const COMPRAS_PROVEEDOR_ARTICULO_DEFAULT={
  'ELEVACIONES ARCHIPIELAGO SAU':'PROD-000047',
  'MASPALOMAS REPUESTOS, S.L.':'PROD-000041',
  'INSUCAN S.L':'PROD-000041',
  'EYSER ISLAS CANARIAS SL':'PROD-000041',
};

async function comprasInitArticulos(){
  const dl=document.getElementById('compras-articulos-list');
  if(!dl||_comprasItemsBC.length)return;
  try{
    const token=await getBCToken();
    const resp=await fetch('/api/bc/items',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({token})});
    const data=await resp.json();
    if(data.ok&&data.items.length){
      _comprasItemsBC=data.items;
      dl.innerHTML='';
      _comprasItemsBC.forEach(it=>{const o=document.createElement('option');o.value=it.number+' — '+it.displayName;dl.appendChild(o);});
    }
  }catch(e){console.warn('No se pudo cargar artículos BC:',e.message);}
}

function comprasAutoArticulo(proveedorNombre){
  const def=COMPRAS_PROVEEDOR_ARTICULO_DEFAULT[proveedorNombre];
  if(!def)return;
  const match=_comprasItemsBC.find(it=>it.number===def||it.displayName===def);
  if(match){
    const inp=document.getElementById('compras-articulo');
    if(inp&&!inp.value)inp.value=match.number+' — '+match.displayName;
  }
}

let _msalInstancePromise=null;
async function getMsalInstance(){
  if(_msalInstancePromise)return _msalInstancePromise;
  _msalInstancePromise=(async()=>{
    if(_msalInstance)return _msalInstance;
    const cfg={auth:{clientId:COMPRAS_CLIENT_ID,authority:'https://login.microsoftonline.com/'+COMPRAS_TENANT_ID,redirectUri:COMPRAS_REDIRECT},cache:{cacheLocation:'sessionStorage'}};
    _msalInstance=new msal.PublicClientApplication(cfg);
    await _msalInstance.initialize();
    return _msalInstance;
  })();
  return _msalInstancePromise;
}

let _comprasTokenPromise=null;
async function comprasGetToken(){
  if(_comprasTokenPromise)return _comprasTokenPromise;
  _comprasTokenPromise=_comprasGetTokenInner().finally(()=>{_comprasTokenPromise=null;});
  return _comprasTokenPromise;
}
async function _comprasGetTokenInner(){
  const m=await getMsalInstance();
  const accounts=m.getAllAccounts();
  if(accounts.length){
    try{const r=await m.acquireTokenSilent({scopes:COMPRAS_SCOPES,account:accounts[0]});return r.accessToken;}catch(e){}
  }
  _clearMsalInteractionState();
  const r = await m.loginPopup({scopes:COMPRAS_SCOPES});
  return r.accessToken;
}

async function comprasInitProveedores(){
  const dl=document.getElementById('compras-proveedores-list');
  if(!dl)return;
  // Si ya cargamos de BC, no recargar
  if(_comprasVendorsBC.length&&dl.children.length)return;
  // Intentar cargar desde BC
  try{
    const token=await getBCToken();
    const resp=await fetch('/api/bc/vendors',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({token})});
    const data=await resp.json();
    if(data.ok&&data.vendors.length){
      _comprasVendorsBC=data.vendors;
      dl.innerHTML='';
      _comprasVendorsBC.forEach(v=>{const o=document.createElement('option');o.value=v.name;dl.appendChild(o);});
      return;
    }
  }catch(e){console.warn('No se pudo cargar vendors de BC, usando lista local:',e.message);}
  // Fallback: lista hardcoded
  if(!dl.children.length){
    COMPRAS_PROVEEDORES.forEach(p=>{const o=document.createElement('option');o.value=p;dl.appendChild(o);});
  }
}

function comprasFileSelected(input){
  const file=input.files[0];if(!file)return;
  _comprasFile=file;
  _comprasFileName=file.name;
  const preview=document.getElementById('compras-preview');
  const isPdf=file.type==='application/pdf'||file.name.toLowerCase().endsWith('.pdf');
  document.getElementById('compras-pdf-viewer').style.display='none';
  if(isPdf){
    preview.style.display='none';
    comprasShowPdfViewer(file);
  }else{
    preview.src=URL.createObjectURL(file);
    preview.style.display='block';
  }
  comprasInitProveedores();
  if(isPdf){
    comprasRunOCRPdf(file);
  }else{
    comprasRunOCR(file);
  }
}


// ── Lazy loaders: Tesseract y PDF.js ───────────────────────
function _loadScript(src) {
  return new Promise((resolve, reject) => {
    const s = document.createElement('script');
    s.src = src; s.onload = resolve; s.onerror = reject;
    document.head.appendChild(s);
  });
}
async function _ensureTesseract() {
  if (typeof Tesseract !== 'undefined') return;
  await _loadScript('https://cdn.jsdelivr.net/npm/tesseract.js@5/dist/tesseract.min.js');
}
async function _ensurePdfjs() {
  if (typeof pdfjsLib !== 'undefined') return;
  await _loadScript('https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.min.js');
}

async function comprasTesseractOCR(file){
  await _ensureTesseract();
  const prog=document.getElementById('compras-ocr-progress');
  const {data:{text}}=await Tesseract.recognize(file,'spa',{
    logger:m=>{
      if(m.status==='recognizing text') prog.textContent='Reconociendo... '+Math.round(m.progress*100)+'%';
    }
  });
  return text;
}

async function comprasRunOCR(file){
  const s2=document.getElementById('compras-step2');
  const s3=document.getElementById('compras-step3');
  const prog=document.getElementById('compras-ocr-progress');
  s2.style.display='block';

  try{
    prog.textContent='Iniciando OCR...';
    const text=await comprasTesseractOCR(file);
    document.getElementById('compras-ocr-text').value=text;
    comprasParseOCR(text);
    s2.style.display='none';
    s3.style.display='block';
  }catch(e){
    s2.style.display='none';
    comprasShowError('Error OCR: '+e.message);
  }
}


function comprasNormalize(s){return s.toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^A-Z0-9]/g,'');}

function comprasFuzzyScore(haystack,needle){
  // Coincidencia exacta normalizada
  if(haystack.includes(needle))return 1;
  // Dividir needle en palabras y contar cuántas aparecen
  const words=needle.split(/[^A-Z0-9]+/).filter(w=>w.length>2);
  if(!words.length)return 0;
  let found=0;
  for(const w of words){if(haystack.includes(w))found++;}
  return found/words.length;
}

function comprasParseOCR(text){
  const normText=comprasNormalize(text);
  const lines=text.split(/\n/).map(l=>l.trim()).filter(Boolean);

  // ── Detectar proveedor (fuzzy) — preferir lista BC, fallback hardcoded ──
  const provList=_comprasVendorsBC.length?_comprasVendorsBC.map(v=>v.name):COMPRAS_PROVEEDORES;
  let bestMatch='',bestScore=0;
  for(const p of provList){
    const normP=comprasNormalize(p);
    const score=comprasFuzzyScore(normText,normP);
    // Priorizar matches más largos a igualdad de score
    if(score>bestScore||(score===bestScore&&p.length>bestMatch.length)){
      bestScore=score;bestMatch=p;
    }
  }
  // Solo aceptar si al menos 60% de las palabras coinciden
  document.getElementById('compras-proveedor').value=bestScore>=0.6?bestMatch:'';

  // ── Detectar nº factura (múltiples patrones) ──
  let nfac='';
  const facPatterns=[
    // "Factura nº 12345" o "Factura: ABC-12345"
    /(?:factura|fra|fact|invoice|albaran|albar[aá]n|ticket|recibo)\s*(?:simplificada|completa|proforma)?\s*[.:;\s\-#nº°n]+\s*([A-Z]{0,4}[\-\/]?\d[\w\/-]{1,20})/i,
    // "Nº 12345" o "Nº factura: 12345"
    /(?:nº|n°|num|numero|número)[.:;\s\-]*(?:de\s+)?(?:factura|fra|fac|albaran|doc)?\s*[.:;\s\-]*([A-Z]{0,4}[\-\/]?\d[\w\/-]{1,20})/i,
    // "Ref: ABC-123"
    /(?:doc|documento|ref|referencia)[.:;\s\-#]*\s*([A-Z]{0,4}[\-\/]?\d[\w\/-]{1,20})/i,
    // Patrón tipo "ABC-12345" suelto
    /\b([A-Z]{1,4}[\-\/]\d{3,10})\b/,
    // Número largo suelto (mínimo 5 dígitos para evitar falsos positivos)
    /\b(\d{5,10})\b/
  ];
  for(const pat of facPatterns){
    const m=text.match(pat);
    if(m){nfac=m[1].trim();break;}
  }
  document.getElementById('compras-nfactura').value=nfac;

  // ── Detectar fecha (múltiples formatos) ──
  let fechaFound='';
  const fechaPatterns=[
    /(?:fecha|date|fch)[.:;\s\-]*(\d{1,2})[\/\-.\s](\d{1,2})[\/\-.\s](\d{2,4})/i,
    /(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})/,
    /(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2})\b/,
    /(\d{1,2})\s+de\s+(enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|octubre|noviembre|diciembre)\s+(?:de\s+)?(\d{4})/i
  ];
  const meses={enero:'01',febrero:'02',marzo:'03',abril:'04',mayo:'05',junio:'06',julio:'07',agosto:'08',septiembre:'09',octubre:'10',noviembre:'11',diciembre:'12'};
  for(const pat of fechaPatterns){
    const m=text.match(pat);
    if(m){
      if(meses[m[2]&&m[2].toLowerCase()]){
        fechaFound=m[3]+'-'+meses[m[2].toLowerCase()]+'-'+m[1].padStart(2,'0');
      }else{
        let d=m[1].padStart(2,'0'),mo=m[2].padStart(2,'0'),y=m[3];
        if(y.length===2)y='20'+y;
        let yn=parseInt(y);
        if(yn<2000||yn>2099)yn=new Date().getFullYear();
        y=String(yn);
        let mn=parseInt(mo),dn=parseInt(d);
        if(mn>12&&dn<=12){const tmp=mo;mo=d;d=tmp;}
        fechaFound=y+'-'+mo+'-'+d;
      }
      break;
    }
  }
  document.getElementById('compras-fecha').value=fechaFound||new Date().toISOString().slice(0,10);
}

async function comprasSubir(){
  const prov=document.getElementById('compras-proveedor').value.trim();
  const nfac=document.getElementById('compras-nfactura').value.trim();
  const fecha=document.getElementById('compras-fecha').value;
  if(!prov){alert('Selecciona un proveedor');return;}
  if(!fecha){alert('Indica la fecha');return;}
  if(!_comprasFile){alert('No hay archivo');return;}

  const btn=document.getElementById('compras-btn-subir');
  btn.disabled=true;btn.textContent='Subiendo...';
  document.getElementById('compras-error').style.display='none';

  try{
    const token=await comprasGetToken();
    const d=new Date(fecha);
    const year=d.getFullYear();
    const mes=COMPRAS_MESES[d.getMonth()];

    // Leer archivo una sola vez para evitar error de File handle expirado
    _comprasFileBuffer=await _comprasFile.arrayBuffer();
    _comprasFileType=_comprasFile.type||'application/octet-stream';

    // Si es imagen, convertir a PDF
    let uploadFile;
    let uploadType=_comprasFileType;
    let ext='.pdf';
    if(_comprasFileType.startsWith('image/')){
      btn.textContent='Convirtiendo a PDF...';
      uploadFile=await comprasImgToPdf(new Blob([_comprasFileBuffer],{type:_comprasFileType}));
      uploadType='application/pdf';
    }else{
      uploadFile=new Blob([_comprasFileBuffer],{type:_comprasFileType});
      ext=_comprasFileName.includes('.')?_comprasFileName.substring(_comprasFileName.lastIndexOf('.')):'.pdf';
    }

    const safeNfac=nfac?nfac.replace(/[\/\\:*?"<>|]/g,'-'):'';
    const fileName=safeNfac?(safeNfac+' '+fecha+ext):(fecha+'_factura'+ext);
    const folderPath=COMPRAS_ONEDRIVE_BASE+'/'+prov+'/'+year+'/'+mes;

    // Resolver carpeta Arifoma via share link
    btn.textContent='Conectando con OneDrive...';
    const shareToken='u!'+btoa(COMPRAS_SHARE_URL).replace(/=+$/,'').replace(/\//g,'_').replace(/\+/g,'-');
    const shareRes=await fetch('https://graph.microsoft.com/v1.0/shares/'+shareToken+'/driveItem?$select=id,parentReference',{
      headers:{'Authorization':'Bearer '+token}
    });
    if(!shareRes.ok){
      const errText=await shareRes.text();
      console.error('Share resolve error:',shareRes.status,errText);
      throw new Error('No se pudo acceder a carpeta Arifoma ('+shareRes.status+')');
    }
    const shareItem=await shareRes.json();
    const driveId=shareItem.parentReference.driveId;
    let parentId=shareItem.id;

    // Navegar subcarpetas (06. ADMINISTRACION / 06.01 PROVEEDORES)
    btn.textContent='Creando carpetas...';
    for(const seg of COMPRAS_ONEDRIVE_BASE.split('/')){
      const listRes=await fetch('https://graph.microsoft.com/v1.0/drives/'+driveId+'/items/'+parentId+'/children?$select=id,name,folder&$top=200',{
        headers:{'Authorization':'Bearer '+token}
      });
      if(!listRes.ok) throw new Error('No se pudo listar carpeta: '+seg);
      const listJson=await listRes.json();
      const segNorm=seg.trim().toLowerCase();
      const found=(listJson.value||[]).find(i=>i.folder&&i.name.trim().toLowerCase()===segNorm);
      if(!found){
        console.error('Buscando "'+seg+'" en carpetas:',(listJson.value||[]).map(i=>i.name));
        throw new Error('Carpeta "'+seg+'" no encontrada en OneDrive');
      }
      parentId=found.id;
    }

    // Función helper para buscar o crear carpeta
    async function _findOrCreate(pid,name){
      const listRes=await fetch('https://graph.microsoft.com/v1.0/drives/'+driveId+'/items/'+pid+'/children?$select=id,name,folder&$top=200',{
        headers:{'Authorization':'Bearer '+token}
      });
      if(!listRes.ok) throw new Error('No se pudo listar carpeta para buscar "'+name+'"');
      const listJson=await listRes.json();
      const nameNorm=name.trim().toLowerCase().normalize('NFC').replace(/\s+/g,' ');
      const found=(listJson.value||[]).find(i=>i.folder&&i.name.trim().toLowerCase().normalize('NFC').replace(/\s+/g,' ')===nameNorm);
      if(found) return {id:found.id,items:listJson.value};
      const createRes=await fetch('https://graph.microsoft.com/v1.0/drives/'+driveId+'/items/'+pid+'/children',{
        method:'POST',
        headers:{'Authorization':'Bearer '+token,'Content-Type':'application/json'},
        body:JSON.stringify({name,folder:{}})
      });
      if(!createRes.ok) throw new Error('Error creando carpeta "'+name+'": '+createRes.status+' '+await createRes.text());
      return {id:(await createRes.json()).id,items:listJson.value};
    }

    // Proveedor
    const provResult=await _findOrCreate(parentId,prov);
    parentId=provResult.id;

    // Comprobar si existe subcarpeta FACTURAS dentro del proveedor
    const provChildRes=await fetch('https://graph.microsoft.com/v1.0/drives/'+driveId+'/items/'+parentId+'/children?$select=id,name,folder&$top=200',{
      headers:{'Authorization':'Bearer '+token}
    });
    if(provChildRes.ok){
      const provChildren=await provChildRes.json();
      const facFolder=(provChildren.value||[]).find(i=>i.folder&&i.name.trim().toUpperCase()==='FACTURAS');
      if(facFolder) parentId=facFolder.id;
    }

    // Año y mes
    const yearResult=await _findOrCreate(parentId,String(year));
    parentId=yearResult.id;
    const mesResult=await _findOrCreate(parentId,mes);
    parentId=mesResult.id;

    const folderId=parentId;

    const uploadUrl='https://graph.microsoft.com/v1.0/drives/'+driveId+'/items/'+folderId+':/'+encodeURIComponent(fileName)+':/content';

    btn.textContent='Subiendo...';
    const resp=await fetch(uploadUrl,{
      method:'PUT',
      headers:{'Authorization':'Bearer '+token,'Content-Type':uploadType},
      body:uploadFile
    });

    if(!resp.ok){const err=await resp.text();throw new Error(err);}

    document.getElementById('compras-step3').style.display='none';
    document.getElementById('compras-step4').style.display='block';
    document.getElementById('compras-ruta-destino').textContent='Arifoma/'+folderPath+'/'+fileName;
    await comprasInitArticulos();
    comprasAutoArticulo(prov);
  }catch(e){
    comprasShowError('Error al subir: '+e.message);
  }finally{
    btn.disabled=false;btn.textContent='Subir a OneDrive';
  }
}

async function comprasCrearPedidoCompra(){
  const prov=document.getElementById('compras-proveedor').value.trim();
  const nfac=document.getElementById('compras-nfactura').value.trim();
  const fecha=document.getElementById('compras-fecha').value;
  const btn=document.getElementById('compras-btn-pedido');
  const resDiv=document.getElementById('compras-pedido-resultado');

  if(!prov){alert('No hay proveedor');return;}

  btn.disabled=true;btn.textContent='Creando pedido...';
  resDiv.style.display='none';

  try{
    const token=await getBCToken();
    const orderDate=fecha||null;
    const articuloVal=(document.getElementById('compras-articulo')?.value||'').trim();
    const itemNumber=articuloVal?articuloVal.split(' — ')[0].trim():null;
    const quantity=document.getElementById('compras-cantidad')?.value||null;
    const unitPrice=document.getElementById('compras-precio')?.value||null;

    console.log('compras payload:', {vendorName:prov, vendorInvoiceNumber:nfac, itemNumber, quantity, unitPrice});
    const resp=await fetch('/api/bc/pedido-compra',{
      method:'POST',
      headers:{'Content-Type':'application/json'},
      body:JSON.stringify({
        token,
        vendorName:prov,
        orderDate,
        vendorInvoiceNumber:nfac||null,
        itemNumber:itemNumber||null,
        quantity:quantity?Number(quantity):null,
        unitPrice:unitPrice?Number(unitPrice):null
      })
    });

    const data=await resp.json();
    if(!data.ok) throw new Error(data.error);

    const o=data.order;

    // Adjuntar factura escaneada al pedido de compra
    if(_comprasFile&&o.id){
      btn.textContent='Adjuntando factura...';
      try{
        const bcBase=`https://api.businesscentral.dynamics.com/v2.0/${BC_TENANT}/${BC_ENV}/api/v2.0/companies`;
        const bcHeaders={'Authorization':'Bearer '+token,'Content-Type':'application/json'};

        const cRes=await fetch(bcBase,{headers:bcHeaders});
        const cJson=await cRes.json();
        const company=(cJson.value||[]).find(c=>c.name.trim()===BC_COMPANY.trim());
        if(company){
          const companyId=company.id;
          const safeNfac=nfac?nfac.replace(/[\/\\:*?"<>|]/g,'-'):'factura';
          const fileName=safeNfac+'.pdf';

          // Reusar buffer ya leído
          let uploadBlob;
          if(_comprasFileType.startsWith('image/')){
            uploadBlob=await comprasImgToPdf(new Blob([_comprasFileBuffer],{type:_comprasFileType}));
          }else{
            uploadBlob=new Blob([_comprasFileBuffer],{type:'application/pdf'});
          }

          // Crear attachment metadata
          const attUrl=`${bcBase}(${companyId})/purchaseOrders(${o.id})/attachments`;
          const attRes=await fetch(attUrl,{
            method:'POST',
            headers:bcHeaders,
            body:JSON.stringify({fileName})
          });
          if(attRes.ok){
            const att=await attRes.json();
            const contentUrl=`${attUrl}(${att.id})/attachmentContent`;
            const patchRes=await fetch(contentUrl,{
              method:'PATCH',
              headers:{'Authorization':'Bearer '+token,'Content-Type':'application/pdf','If-Match':att['@odata.etag']||'*'},
              body:uploadBlob
            });
            if(!patchRes.ok) console.warn('Attachment content error:',patchRes.status,await patchRes.text());
          }else{
            console.warn('Attachment create error:',attRes.status,await attRes.text());
          }
        }
      }catch(attErr){
        console.warn('No se pudo adjuntar factura:',attErr.message);
      }
    }

    resDiv.style.display='block';
    resDiv.style.background='var(--success-bg,#d4edda)';
    resDiv.style.color='var(--success-text,#155724)';
    resDiv.innerHTML=`✓ Pedido de compra <b>${o.number||''}</b> creado en BC`+(o.vendorName?` para <b>${o.vendorName}</b>`:'');
    btn.style.display='none';
  }catch(e){
    resDiv.style.display='block';
    resDiv.style.background='var(--danger-bg,#f8d7da)';
    resDiv.style.color='var(--danger-text,#721c24)';
    resDiv.textContent='Error: '+e.message;
  }finally{
    btn.disabled=false;btn.textContent='Crear Pedido de Compra en BC';
  }
}

async function comprasRunOCRPdf(file){
  const s2=document.getElementById('compras-step2');
  const s3=document.getElementById('compras-step3');
  const prog=document.getElementById('compras-ocr-progress');
  s2.style.display='block';

  try{
    // Renderizar primera página del PDF a imagen y enviar a Gemini
    prog.textContent='Leyendo PDF...';
    await _ensurePdfjs();
    const arrayBuf=await file.arrayBuffer();
    pdfjsLib.GlobalWorkerOptions.workerSrc='https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js';
    const pdf=await pdfjsLib.getDocument({data:arrayBuf}).promise;
    const page=await pdf.getPage(1);
    const viewport=page.getViewport({scale:2});
    const canvas=document.createElement('canvas');
    canvas.width=viewport.width;canvas.height=viewport.height;
    const ctx=canvas.getContext('2d');
    await page.render({canvasContext:ctx,viewport}).promise;
    const blob=await new Promise(r=>canvas.toBlob(r,'image/png'));
    const text=await comprasTesseractOCR(blob);
    document.getElementById('compras-ocr-text').value=text;
    comprasParseOCR(text);
    s2.style.display='none';
    s3.style.display='block';
  }catch(e){
    s2.style.display='none';
    comprasShowError('Error OCR PDF: '+(e.message||e));
  }
}

function comprasImgToPdf(file){
  return new Promise((resolve,reject)=>{
    const img=new Image();
    img.onload=()=>{
      const w=img.width,h=img.height;
      // PDF en puntos (72dpi), ajustar a A4 si es muy grande
      const a4w=595.28,a4h=841.89;
      let pw,ph;
      const ratio=w/h;
      if(ratio>a4w/a4h){pw=a4w;ph=a4w/ratio;}
      else{ph=a4h;pw=a4h*ratio;}

      // Dibujar imagen en canvas a resolución original
      const canvas=document.createElement('canvas');
      canvas.width=w;canvas.height=h;
      const ctx=canvas.getContext('2d');
      ctx.drawImage(img,0,0);
      const jpegData=canvas.toDataURL('image/jpeg',0.92);

      // Construir PDF manualmente (mínimo válido)
      const imgBytes=atob(jpegData.split(',')[1]);
      const imgLen=imgBytes.length;
      const imgStream=new Uint8Array(imgLen);
      for(let i=0;i<imgLen;i++)imgStream[i]=imgBytes.charCodeAt(i);

      let pdf='%PDF-1.4\n';
      const offsets=[];

      offsets.push(pdf.length);
      pdf+='1 0 obj<</Type/Catalog/Pages 2 0 R>>endobj\n';

      offsets.push(pdf.length);
      pdf+='2 0 obj<</Type/Pages/Kids[3 0 R]/Count 1>>endobj\n';

      offsets.push(pdf.length);
      pdf+='3 0 obj<</Type/Page/Parent 2 0 R/MediaBox[0 0 '+pw.toFixed(2)+' '+ph.toFixed(2)+']/Contents 4 0 R/Resources<</XObject<</Img 5 0 R>>>>>>endobj\n';

      const content='q '+pw.toFixed(2)+' 0 0 '+ph.toFixed(2)+' 0 0 cm /Img Do Q';
      offsets.push(pdf.length);
      pdf+='4 0 obj<</Length '+content.length+'>>stream\n'+content+'\nendstream\nendobj\n';

      offsets.push(pdf.length);
      const imgHeader='5 0 obj<</Type/XObject/Subtype/Image/Width '+w+'/Height '+h+'/ColorSpace/DeviceRGB/BitsPerComponent 8/Filter/DCTDecode/Length '+imgLen+'>>stream\n';

      // Combinar texto + imagen binaria + resto
      const encoder=new TextEncoder();
      const pdfBefore=encoder.encode(pdf+imgHeader);
      const streamEnd=encoder.encode('\nendstream\nendobj\n');

      const xrefStart=pdfBefore.length+imgLen+streamEnd.length;
      let xref='xref\n0 6\n0000000000 65535 f \n';
      for(let i=0;i<5;i++)xref+=String(offsets[i]).padStart(10,'0')+' 00000 n \n';
      xref+='trailer<</Size 6/Root 1 0 R>>\nstartxref\n'+xrefStart+'\n%%EOF';
      const xrefBytes=encoder.encode(xref);

      const final=new Uint8Array(pdfBefore.length+imgLen+streamEnd.length+xrefBytes.length);
      final.set(pdfBefore,0);
      final.set(imgStream,pdfBefore.length);
      final.set(streamEnd,pdfBefore.length+imgLen);
      final.set(xrefBytes,pdfBefore.length+imgLen+streamEnd.length);

      resolve(new Blob([final],{type:'application/pdf'}));
    };
    img.onerror=()=>reject(new Error('No se pudo cargar la imagen'));
    img.src=URL.createObjectURL(file);
  });
}

// ── Visor PDF ──
let _comprasPdf=null,_comprasPdfPage=1,_comprasPdfZoom=1;

async function comprasShowPdfViewer(file){
  await _ensurePdfjs();
  const viewer=document.getElementById('compras-pdf-viewer');
  viewer.style.display='block';
  const arrayBuf=await file.arrayBuffer();
  pdfjsLib.GlobalWorkerOptions.workerSrc='https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js';
  _comprasPdf=await pdfjsLib.getDocument({data:arrayBuf}).promise;
  _comprasPdfPage=1;_comprasPdfZoom=1;
  document.getElementById('compras-pdf-total').textContent=_comprasPdf.numPages;
  comprasRenderPdfPage();
}

async function comprasRenderPdfPage(){
  if(!_comprasPdf)return;
  const page=await _comprasPdf.getPage(_comprasPdfPage);
  const viewport=page.getViewport({scale:_comprasPdfZoom*1.5});
  const canvas=document.getElementById('compras-pdf-canvas');
  canvas.width=viewport.width;canvas.height=viewport.height;
  await page.render({canvasContext:canvas.getContext('2d'),viewport}).promise;
  document.getElementById('compras-pdf-page').textContent=_comprasPdfPage;
  document.getElementById('compras-pdf-zoom-label').textContent=Math.round(_comprasPdfZoom*100)+'%';
}

function comprasPdfZoom(delta){
  _comprasPdfZoom=Math.max(0.25,Math.min(4,_comprasPdfZoom+delta));
  comprasRenderPdfPage();
}

function comprasPdfNav(dir){
  if(!_comprasPdf)return;
  const next=_comprasPdfPage+dir;
  if(next<1||next>_comprasPdf.numPages)return;
  _comprasPdfPage=next;
  comprasRenderPdfPage();
}

function comprasShowError(msg){
  const el=document.getElementById('compras-error');
  el.textContent=msg;el.style.display='block';
}

function comprasReset(){
  _comprasFile=null;_comprasFileName='';_comprasPdf=null;
  document.getElementById('compras-preview').style.display='none';
  document.getElementById('compras-pdf-viewer').style.display='none';
  document.getElementById('compras-step2').style.display='none';
  document.getElementById('compras-step3').style.display='none';
  document.getElementById('compras-step4').style.display='none';
  document.getElementById('compras-error').style.display='none';
  document.getElementById('compras-file-input').value='';
  document.getElementById('compras-proveedor').value='';
  document.getElementById('compras-nfactura').value='';
  document.getElementById('compras-fecha').value='';
  document.getElementById('compras-ocr-text').value='';
  const artInp=document.getElementById('compras-articulo');if(artInp)artInp.value='';
  const cantInp=document.getElementById('compras-cantidad');if(cantInp)cantInp.value='';
  const precInp=document.getElementById('compras-precio');if(precInp)precInp.value='';
  const btnPed=document.getElementById('compras-btn-pedido');
  if(btnPed){btnPed.style.display='';btnPed.disabled=false;btnPed.textContent='Crear Pedido de Compra en BC';}
  const resDiv=document.getElementById('compras-pedido-resultado');
  if(resDiv){resDiv.style.display='none';resDiv.innerHTML='';}
}

// ============================================================
// INFORMES DIARIOS PLANTA
// ============================================================

const INF_MAQUINARIA = [
  'CAT 330','CAT 336','PALA 966G','DUMPER 769','LAGARTO HM400','CAT 365'
];

let _infData = null; // cache último informe cargado

async function cargarInformeDiario() {
  const fecha = document.getElementById('inf-fecha').value;
  if (!fecha) { alert('Selecciona una fecha'); return; }

  document.getElementById('inf-empty').style.display = 'none';
  document.getElementById('inf-content').style.display = 'none';
  document.getElementById('inf-loading').style.display = 'block';

  try {
    // fechaHora en tblpedidos es ISO: 2026-06-02T07:10:00Z → filtrar por rango ISO
    const fechaDesde = fecha + 'T00:00:00';
    const fechaHasta = fecha + 'T23:59:59';

    const [fichajesRes, pedidosRes, produccionRes, gasoilRes, stockRes] = await Promise.all([
      dbQuery({ action:'select', table:'tblFichaje', filters:[{column:'fecha',op:'eq',value:fecha}], options:{select:'empleado,entrada,salida,tiempodia,fentrada,fsalida'} }),
      dbQuery({ action:'select', table:'tblpedidos', filters:[{column:'fechaHora',op:'gte',value:fechaDesde},{column:'fechaHora',op:'lte',value:fechaHasta}], options:{select:'fechaHora,matriculacam,pesoBruto,pesoNeto,productoNombre,nombreCliente',order:'fechaHora'} }),
      dbQuery({ action:'select', table:'PRODUCCION', filters:[{column:'fecha',op:'eq',value:fecha}], options:{select:'fecha,tipoDia,t04,t412,t1220,t2040,tnDia,horasPlanta'} }),
      dbQuery({ action:'select', table:'GASOIL', filters:[{column:'fecha',op:'eq',value:fecha}], options:{select:'fecha,origen,destino,litros,tipo'} }),
      apiFetch('?accion=gasoil'),
    ]);

    const gasoilDelDia = gasoilRes.data || [];
    _infData = {
      fecha,
      fichajes: fichajesRes.data || [],
      pedidos: pedidosRes.data || [],
      produccion: (produccionRes.data || [])[0] || null,
      gasoil: gasoilDelDia,
      stock: { dep1: stockRes.dep1||0, dep2: stockRes.dep2||0 },
    };

    _renderInforme(_infData);
    document.getElementById('inf-loading').style.display = 'none';
    document.getElementById('inf-content').style.display = 'block';
  } catch(e) {
    document.getElementById('inf-loading').style.display = 'none';
    document.getElementById('inf-empty').style.display = 'block';
    document.getElementById('inf-empty').textContent = 'Error: ' + e.message;
    console.error('Error informe:', e);
  }
}

// Calcula horas reales desde entrada/salida (strings HH:MM locales); fallback a tiempodia
function calcHorasFichaje(f) {
  if (f.entrada && f.salida) {
    const [eh, em] = f.entrada.split(':').map(Number);
    const [sh, sm] = f.salida.split(':').map(Number);
    const mins = (sh * 60 + sm) - (eh * 60 + em);
    if (mins > 0) return (mins / 60).toFixed(2);
  }
  if (f.tiempodia != null) return parseFloat(String(f.tiempodia).replace(',','.')).toFixed(2);
  return '—';
}

function _renderInforme(d) {
  const fmt = v => Number(v||0).toLocaleString('es-ES',{minimumFractionDigits:2,maximumFractionDigits:2});
  const calcHoras = calcHorasFichaje;
  const fmtH = v => v!=null ? parseFloat(String(v).replace(',','.')).toFixed(1) : '—';

  // Maquinaria (fija) — inputs con value por defecto
  const maqBody = document.getElementById('inf-maquinaria-body');
  maqBody.innerHTML = INF_MAQUINARIA.map(m =>
    `<tr style="border-bottom:1px solid var(--border)">
      <td style="padding:5px 8px;font-weight:600">${m}</td>
      <td style="padding:5px 8px;text-align:center"><input type="number" min="0" max="24" value="7" style="width:50px;text-align:center;border:1px solid var(--border);border-radius:4px;padding:2px;background:var(--surface2)"></td>
      <td style="padding:5px 8px;text-align:center"><input type="number" min="0" max="24" value="17" style="width:50px;text-align:center;border:1px solid var(--border);border-radius:4px;padding:2px;background:var(--surface2)"></td>
      <td style="padding:5px 8px;text-align:center;font-weight:600" class="inf-maq-horas">10</td>
      <td style="padding:5px 8px"><input type="text" style="width:100%;border:1px solid var(--border);border-radius:4px;padding:2px 4px;background:var(--surface2)" placeholder="Trabajo..."></td>
    </tr>`
  ).join('');
  maqBody.querySelectorAll('tr').forEach(row => {
    const [desdeTd, hastaTd, horasTd] = [row.children[1], row.children[2], row.children[3]];
    const calc = () => {
      const desde = parseFloat(desdeTd.querySelector('input').value);
      const hasta = parseFloat(hastaTd.querySelector('input').value);
      horasTd.textContent = (!isNaN(desde)&&!isNaN(hasta)&&hasta>desde) ? (hasta-desde) : '—';
    };
    desdeTd.querySelector('input').addEventListener('input', calc);
    hastaTd.querySelector('input').addEventListener('input', calc);
  });

  // Personal
  const perBody = document.getElementById('inf-personal-body');
  if (!d.fichajes.length) {
    perBody.innerHTML = '<tr><td colspan="4" style="padding:8px;color:var(--muted)">Sin fichajes</td></tr>';
  } else {
    perBody.innerHTML = d.fichajes.map(f =>
      `<tr style="border-bottom:1px solid var(--border)">
        <td style="padding:5px 8px;font-weight:600">${f.empleado||'—'}</td>
        <td style="padding:5px 8px;text-align:center;color:var(--accent2)">${f.entrada||'—'}</td>
        <td style="padding:5px 8px;text-align:center;color:var(--danger)">${f.salida||'—'}</td>
        <td style="padding:5px 8px;text-align:center">${calcHoras(f)}</td>
      </tr>`
    ).join('');
  }

  // Producción
  const prodDiv = document.getElementById('inf-produccion-body');
  if (!d.produccion) {
    prodDiv.innerHTML = '<div style="color:var(--muted)">Sin datos de producción</div>';
  } else {
    const p = d.produccion;
    prodDiv.innerHTML = `
      <div style="display:flex;gap:16px;flex-wrap:wrap">
        <div><span style="color:var(--muted)">Tipo día:</span> <b>${p.tipoDia||'—'}</b></div>
        <div><span style="color:var(--muted)">Total Tn:</span> <b>${fmt(p.tnDia)}</b></div>
        <div><span style="color:var(--muted)">H.Planta:</span> <b>${p.horasPlanta||'—'}</b></div>
      </div>
      <div style="display:flex;gap:16px;flex-wrap:wrap;margin-top:8px">
        <div><span style="color:var(--muted)">0/4:</span> <b>${fmt(p.t04)} Tn</b></div>
        <div><span style="color:var(--muted)">4/12:</span> <b>${fmt(p.t412)} Tn</b></div>
        <div><span style="color:var(--muted)">12/20:</span> <b>${fmt(p.t1220)} Tn</b></div>
        <div><span style="color:var(--muted)">20/40:</span> <b>${fmt(p.t2040)} Tn</b></div>
      </div>`;
  }

  // Ventas
  const ventBody = document.getElementById('inf-ventas-body');
  let totalNeto = 0, totalImporte = 0;
  if (!d.pedidos.length) {
    ventBody.innerHTML = '<tr><td colspan="8" style="padding:8px;color:var(--muted)">Sin pesadas</td></tr>';
  } else {
    ventBody.innerHTML = d.pedidos.map(p => {
      const d2 = new Date(p.fechaHora);
      const hora = !isNaN(d2) ? pad(d2.getHours())+':'+pad(d2.getMinutes()) : (p.fechaHora||'').split(' ')[1]||'';
      const neto = Number(p.pesoNeto||0)/1000;
      const precio = getPrecioTn(p.nombreCliente, p.productoNombre);
      const total = neto * precio;
      totalNeto += neto; totalImporte += total;
      return `<tr style="border-bottom:1px solid var(--border)">
        <td style="padding:4px 8px;color:var(--muted)">${hora}</td>
        <td style="padding:4px 8px">${p.matriculacam||'—'}</td>
        <td style="padding:4px 8px;text-align:right">${Number(p.pesoBruto||0).toLocaleString()}</td>
        <td style="padding:4px 8px;text-align:right">${Number(p.pesoNeto||0).toLocaleString()}</td>
        <td style="padding:4px 8px">${p.productoNombre||'—'}</td>
        <td style="padding:4px 8px">${p.nombreCliente||'—'}</td>
        <td style="padding:4px 8px;text-align:right">${fmt(precio)}</td>
        <td style="padding:4px 8px;text-align:right;font-weight:600">${fmt(total)} €</td>
      </tr>`;
    }).join('');
  }
  document.getElementById('inf-ventas-total-neto').textContent = fmt(totalNeto) + ' Tn';
  document.getElementById('inf-ventas-total').textContent = fmt(totalImporte) + ' €';

  // Gasoil + Stock depósitos
  const gasoilDiv = document.getElementById('inf-gasoil-body');
  const totalLitros = d.gasoil.reduce((s,r)=>s+Number(r.litros||0),0);
  const stockHtml = `<div style="display:flex;gap:16px;margin-top:10px;flex-wrap:wrap">
    <div style="background:var(--surface2);border-radius:6px;padding:8px 14px;font-size:.8rem">
      <div style="color:var(--muted);font-size:.7rem;text-transform:uppercase">DEPÓSITO 1</div>
      <div style="font-weight:700;font-size:1rem">${Number(d.stock.dep1).toLocaleString()} L</div>
    </div>
    <div style="background:var(--surface2);border-radius:6px;padding:8px 14px;font-size:.8rem">
      <div style="color:var(--muted);font-size:.7rem;text-transform:uppercase">DEPÓSITO 2</div>
      <div style="font-weight:700;font-size:1rem">${Number(d.stock.dep2).toLocaleString()} L</div>
    </div>
  </div>`;
  if (!d.gasoil.length) {
    gasoilDiv.innerHTML = '<div style="color:var(--muted)">Sin movimientos de gasoil</div>' + stockHtml;
  } else {
    gasoilDiv.innerHTML = `<table style="width:100%;border-collapse:collapse">
      <thead><tr style="border-bottom:2px solid var(--border)">
        <th style="text-align:left;padding:4px 8px">Origen</th>
        <th style="text-align:left;padding:4px 8px">Destino</th>
        <th style="text-align:left;padding:4px 8px">Tipo</th>
        <th style="text-align:right;padding:4px 8px">Litros</th>
      </tr></thead>
      <tbody>${d.gasoil.map(g=>`<tr style="border-bottom:1px solid var(--border)">
        <td style="padding:4px 8px">${g.origen||'—'}</td>
        <td style="padding:4px 8px">${g.destino||'—'}</td>
        <td style="padding:4px 8px">${g.tipo||'—'}</td>
        <td style="padding:4px 8px;text-align:right">${Number(g.litros||0).toLocaleString()} L</td>
      </tr>`).join('')}</tbody>
      <tfoot><tr style="border-top:2px solid var(--border);font-weight:700">
        <td colspan="3" style="padding:4px 8px">TOTAL CONSUMIDO DÍA</td>
        <td style="padding:4px 8px;text-align:right">${totalLitros.toLocaleString()} L</td>
      </tr></tfoot>
    </table>${stockHtml}`;
  }
}

async function infEnviarEmail() {
  if (!_infData) { alert('Carga primero el informe'); return; }
  if (typeof ExcelJS === 'undefined') { alert('Librería Excel no cargada'); return; }
  const btn = document.querySelector('button[onclick="infEnviarEmail()"]');
  if (btn) { btn.disabled = true; btn.textContent = 'Enviando...'; }
  try {
    const buf = await _infGenerarBuffer(_infData);
    const base64 = btoa(String.fromCharCode(...new Uint8Array(buf)));
    const fechaFmt = _infData.fecha.split('-').reverse().join('/');
    const fileName = `Informe_Planta_${_infData.fecha}.xlsx`;
    const token = await comprasGetToken();
    const mail = {
      message: {
        subject: `ARIFOMA DATOS DIARIOS ${fechaFmt}`,
        body: { contentType: 'Text', content: `Buenas tardes,\n\nAdjunto datos diarios del día ${fechaFmt}.\n\nSaludos,` },
        toRecipients: [
          { emailAddress: { address: 'jpereira@lopesan.com' } },
          { emailAddress: { address: 'mleon@lopesan.com' } },
          { emailAddress: { address: 'asarmiento@lopesan.com' } }
        ],
        attachments: [{
          '@odata.type': '#microsoft.graph.fileAttachment',
          name: fileName,
          contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          contentBytes: base64
        }]
      }
    };
    const res = await fetch('https://graph.microsoft.com/v1.0/me/sendMail', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
      body: JSON.stringify(mail)
    });
    if (res.ok || res.status === 202) {
      alert('Email enviado correctamente a jpereira@lopesan.com');
    } else {
      const err = await res.text();
      alert('Error al enviar: ' + err);
    }
  } catch(e) {
    alert('Error: ' + e.message);
  } finally {
    if (btn) { btn.disabled = false; btn.textContent = 'Enviar por Email'; }
  }
}

async function _infGenerarBuffer(d) {
  const wb = new ExcelJS.Workbook();
  const fechaFmt = d.fecha.split('-').reverse().join('/');
  const ws = wb.addWorksheet('Informe ' + d.fecha);
  ws.columns = [{width:22},{width:20},{width:8},{width:8},{width:28},{width:30},{width:12},{width:12},{width:12},{width:12},{width:18},{width:28},{width:10},{width:14}];
  const hdrFill = {type:'pattern',pattern:'solid',fgColor:{argb:'FF374151'}};
  const subFill  = {type:'pattern',pattern:'solid',fgColor:{argb:'FFD1D5DB'}};
  const hdrFont  = {bold:true,color:{argb:'FFFFFFFF'},size:10};
  const subFont  = {bold:true,size:10};
  const boldFont = {bold:true,size:10};
  const border   = {top:{style:'thin'},left:{style:'thin'},bottom:{style:'thin'},right:{style:'thin'}};
  const addHdr = (label, cols=14) => {
    const r = ws.addRow([label]);
    r.getCell(1).fill=hdrFill; r.getCell(1).font=hdrFont;
    ws.mergeCells(r.number,1,r.number,cols);
    r.height=16;
  };
  const addSubHdr = (vals) => {
    const r = ws.addRow(vals);
    r.eachCell(c=>{c.fill=subFill;c.font=subFont;c.border=border;c.alignment={horizontal:'center'};});
    r.getCell(1).alignment={horizontal:'left'};
  };
  const addRow = (vals) => { const r = ws.addRow(vals); r.eachCell(c=>{c.border=border;}); return r; };
  const addBlank = () => ws.addRow([]);
  const titulo = ws.addRow([`INFORME DIARIO PLANTA — ${fechaFmt}`]);
  titulo.getCell(1).font={bold:true,size:13}; titulo.height=20;
  ws.mergeCells(titulo.number,1,titulo.number,14);
  addBlank();
  addHdr('MAQUINARIA', 5);
  addSubHdr(['MAQUINARIA','DESDE','HASTA','HORAS','TRABAJO']);
  const maqRows = document.querySelectorAll('#inf-maquinaria-body tr');
  INF_MAQUINARIA.forEach((m, i) => {
    const row = maqRows[i];
    const inputs = row ? row.querySelectorAll('input') : [];
    const desde = inputs[0]?.value||''; const hasta = inputs[1]?.value||''; const trabajo = inputs[2]?.value||'';
    const horas = (desde&&hasta&&Number(hasta)>Number(desde)) ? Number(hasta)-Number(desde) : '';
    addRow([m, desde?Number(desde):'', hasta?Number(hasta):'', horas, trabajo]);
  });
  addBlank();
  addHdr('PERSONAL', 4);
  addSubHdr(['NOMBRE','ENTRADA','SALIDA','HORAS']);
  d.fichajes.forEach(f => { const horas = calcHorasFichaje(f); addRow([f.empleado||'', f.entrada||'', f.salida||'', horas!=='—'?Number(horas):'']); });
  addBlank();
  addHdr('PRODUCCIÓN', 8);
  addSubHdr(['FECHA','DÍA','0/4 Tn','4/12 Tn','12/20 Tn','20/40 Tn','TOTAL Tn','H.PLANTA']);
  if (d.produccion) {
    const p = d.produccion;
    addRow([fechaFmt, p.tipoDia||'', Number(p.t04||0), Number(p.t412||0), Number(p.t1220||0), Number(p.t2040||0), Number(p.tnDia||0), p.horasPlanta||'']);
    const totProd = ws.addRow(['','TOTALES',Number(p.t04||0),Number(p.t412||0),Number(p.t1220||0),Number(p.t2040||0),Number(p.tnDia||0),'']);
    totProd.font=boldFont; totProd.eachCell(c=>{c.border=border;});
  }
  addBlank();
  addHdr('VENTAS', 8);
  addSubHdr(['FECHA-HORA','MATRÍCULA','BRUTO','NETO','MATERIAL','CLIENTE','PRECIO','TOTAL']);
  let totalNeto=0, totalImp=0;
  d.pedidos.forEach(p => {
    const d2 = new Date(p.fechaHora);
    const hora = !isNaN(d2) ? fechaFmt+' '+pad(d2.getHours())+':'+pad(d2.getMinutes()) : p.fechaHora||'';
    const neto = Number(p.pesoNeto||0)/1000;
    const precio = getPrecioTn(p.nombreCliente, p.productoNombre);
    const total = neto * precio;
    totalNeto += neto; totalImp += total;
    addRow([hora, p.matriculacam||'', Number(p.pesoBruto||0), Number(p.pesoNeto||0), p.productoNombre||'', p.nombreCliente||'', precio, total]);
  });
  const totVent = ws.addRow(['','','',totalNeto,'','','TOTAL DIARIO (sin IGIC)',totalImp]);
  totVent.font=boldFont; totVent.eachCell(c=>{c.border=border;});
  addBlank();
  addHdr('GASOIL', 5);
  addSubHdr(['FECHA','ORIGEN','DESTINO','TIPO','TOTAL (L)']);
  let totalL=0;
  d.gasoil.forEach(g => { totalL += Number(g.litros||0); addRow([fechaFmt, g.origen||'', g.destino||'', g.tipo||'', Number(g.litros||0)]); });
  const totGas = ws.addRow(['','','','TOTAL CONSUMIDO DÍA', totalL]);
  totGas.font=boldFont; totGas.eachCell(c=>{c.border=border;});
  addBlank();
  addHdr('STOCK DEPÓSITOS', 3);
  addSubHdr(['DEPÓSITO','STOCK ACTUAL']);
  addRow(['DEPOSITO 1', Number(d.stock.dep1)]);
  addRow(['DEPOSITO 2', Number(d.stock.dep2)]);
  return wb.xlsx.writeBuffer();
}

async function infExportarExcel() {
  if (!_infData) { alert('Carga primero el informe'); return; }
  if (typeof ExcelJS === 'undefined') { alert('Librería Excel no cargada'); return; }
  const buf = await _infGenerarBuffer(_infData);
  const blob = new Blob([buf], {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href=url; a.download=`Informe_Planta_${_infData.fecha}.xlsx`; a.click();
  URL.revokeObjectURL(url);

  // Subir a OneDrive
  try {
    const [yyyy, mm] = d.fecha.split('-');
    const fileName = `Informe_Planta_${d.fecha}.xlsx`;
    const INF_BASE = '06. ADMINISTRACION/06.11 DOCUMENTOS/Informes Diarios';
    const token = await comprasGetToken();
    const shareToken = 'u!' + btoa(COMPRAS_SHARE_URL).replace(/=+$/, '').replace(/\//g, '_').replace(/\+/g, '-');
    const shareRes = await fetch('https://graph.microsoft.com/v1.0/shares/' + shareToken + '/driveItem?$select=id,parentReference', {
      headers: { 'Authorization': 'Bearer ' + token }
    });
    if (!shareRes.ok) throw new Error('No se pudo acceder a OneDrive (' + shareRes.status + ')');
    const shareItem = await shareRes.json();
    const driveId = shareItem.parentReference.driveId;
    let parentId = shareItem.id;

    // Navegar/crear carpetas: 06. ADMINISTRACION / 06.11 DOCUMENTOS / Informes Diarios / año
    async function _infFindOrCreate(pid, name) {
      const listRes = await fetch('https://graph.microsoft.com/v1.0/drives/' + driveId + '/items/' + pid + '/children?$select=id,name,folder&$top=200', {
        headers: { 'Authorization': 'Bearer ' + token }
      });
      if (!listRes.ok) throw new Error('No se pudo listar carpeta "' + name + '"');
      const listJson = await listRes.json();
      const norm = s => s.trim().toLowerCase().normalize('NFC').replace(/\s+/g, ' ');
      const found = (listJson.value || []).find(i => i.folder && norm(i.name) === norm(name));
      if (found) return found.id;
      const createRes = await fetch('https://graph.microsoft.com/v1.0/drives/' + driveId + '/items/' + pid + '/children', {
        method: 'POST',
        headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
        body: JSON.stringify({ name, folder: {} })
      });
      if (!createRes.ok) throw new Error('Error creando carpeta "' + name + '": ' + await createRes.text());
      return (await createRes.json()).id;
    }

    for (const seg of INF_BASE.split('/')) {
      parentId = await _infFindOrCreate(parentId, seg);
    }
    parentId = await _infFindOrCreate(parentId, yyyy);
    parentId = await _infFindOrCreate(parentId, COMPRAS_MESES[parseInt(mm, 10) - 1]);

    const uploadUrl = 'https://graph.microsoft.com/v1.0/drives/' + driveId + '/items/' + parentId + ':/' + encodeURIComponent(fileName) + ':/content';
    const upRes = await fetch(uploadUrl, {
      method: 'PUT',
      headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' },
      body: new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
    });
    if (upRes.ok) {
      console.log('Informe subido a OneDrive: Arifoma/' + INF_BASE + '/' + yyyy + '/' + fileName);
    } else {
      console.warn('Error subiendo informe a OneDrive:', upRes.status, await upRes.text());
    }
  } catch(e) {
    console.warn('No se pudo subir a OneDrive:', e.message);
  }
}

// ============================================================
// STOCK ÁRIDOS — Gráficos montaña desde Google Sheet (hoja STOCK)
// ============================================================
// Hoja STOCK: cada fila = un día. Col C=fecha, N=0/4, O=4/12, P=12/20, Q=20/40
// Fila 5 = existencias iniciales, fila 6+ = datos diarios
// Se agrupa por mes: primer valor = inicio mes, último = final mes

const STOCK_SHEET_ID = '1fxHwVEgcIrRdyPh-TJ-k84QFBHXX-P3mNRCiWYaeDTQ';
const STOCK_SHEET_TAB = 'STOCK';
const STOCK_MESES = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
const STOCK_COLORS = {
  '04':   {line:'#6b7d2e', bg:'rgba(107,125,46,.25)'},
  '412':  {line:'#2e6b7d', bg:'rgba(46,107,125,.25)'},
  '1220': {line:'#7d2e6b', bg:'rgba(125,46,107,.25)'},
  '2040': {line:'#c0792b', bg:'rgba(192,121,43,.25)'}
};
const STOCK_LABELS = {'04':'0/4','412':'4/12','1220':'12/20','2040':'20/40'};

let _stockData = null;
let _stockRawDaily = null;
let _stockCharts = {mini04:null, mini412:null, mini1220:null, mini2040:null, grande:null};
let _stockInited = false;

function initStock() {
  if (!_stockInited) {
    const sel = document.getElementById('stock-anyo');
    const yr = new Date().getFullYear();
    for (let y = yr; y >= yr - 3; y--) {
      const o = document.createElement('option');
      o.value = y; o.textContent = y;
      sel.appendChild(o);
    }
    _stockInited = true;
  }
  cargarStock();
}

async function cargarStock() {
  const anyo = parseInt(document.getElementById('stock-anyo').value);
  try {
    // Leer fecha (C) y stock (N,O,P,Q)
    const query = `select C,N,O,P,Q`;
    const url = `https://docs.google.com/spreadsheets/d/${STOCK_SHEET_ID}/gviz/tq?tqx=out:json&sheet=${encodeURIComponent(STOCK_SHEET_TAB)}&tq=${encodeURIComponent(query)}`;
    const res = await fetch(url);
    const txt = await res.text();
    const jsonStr = txt.replace(/^[^(]*\(/, '').replace(/\);?\s*$/, '');
    const gviz = JSON.parse(jsonStr);
    const rows = gviz.table.rows || [];

    const keys = ['04', '412', '1220', '2040'];
    // Parse todas las filas con fecha y valores
    const daily = []; // {date, '04':v, '412':v, ...}
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      if (!r.c) continue;
      // Fecha en col 0 (C)
      const dateCell = r.c[0];
      let fecha = null;
      if (dateCell && dateCell.v) {
        // gviz dates: "Date(2026,0,1)" or string
        const dv = dateCell.v;
        if (typeof dv === 'string' && dv.startsWith('Date(')) {
          const parts = dv.replace('Date(','').replace(')','').split(',').map(Number);
          fecha = new Date(parts[0], parts[1], parts[2]);
        } else if (dv instanceof Date) {
          fecha = dv;
        } else if (typeof dv === 'string') {
          fecha = new Date(dv);
        }
      }
      // Stock values (cols 1-4 = N,O,P,Q)
      const vals = {};
      let hasAny = false;
      for (let k = 0; k < 4; k++) {
        const cell = r.c[k + 1];
        const v = cell && cell.v != null ? Number(cell.v) : null;
        vals[keys[k]] = v;
        if (v != null && !isNaN(v)) hasAny = true;
      }
      if (hasAny) daily.push({ fecha, ...vals });
    }

    _stockRawDaily = daily;

    // Filtrar por año y agrupar por mes
    const filtered = daily.filter(d => d.fecha && d.fecha.getFullYear() === anyo);
    _stockData = {};
    for (const key of keys) {
      const meses = [];
      for (let m = 0; m < 12; m++) {
        const delMes = filtered.filter(d => d.fecha.getMonth() === m && d[key] != null);
        if (delMes.length === 0) {
          meses.push({ mes: STOCK_MESES[m], inicio: null, final: null, actual: null });
          continue;
        }
        const inicio = delMes[0][key];
        const final_ = delMes[delMes.length - 1][key];
        // "actual" = último valor del mes actual, null para meses pasados
        const hoy = new Date();
        const esActual = (hoy.getFullYear() === anyo && hoy.getMonth() === m);
        const actual = esActual ? final_ : null;
        meses.push({ mes: STOCK_MESES[m], inicio, final: final_, actual });
      }
      _stockData[key] = meses;
    }

    renderStockOverview();
  } catch (e) {
    console.error('Error cargando stock:', e);
  }
}

function renderStockOverview() {
  if (!_stockData) return;
  document.getElementById('stock-overview').style.display = '';
  document.getElementById('stock-detalle').style.display = 'none';

  const keys = ['04', '412', '1220', '2040'];

  // --- KPI cards ---
  for (const key of keys) {
    const meses = _stockData[key] || [];
    // Find current month's actual value, or last month with final value
    let lastVal = null, prevVal = null;
    for (let i = meses.length - 1; i >= 0; i--) {
      if (lastVal == null && (meses[i].actual != null || meses[i].final != null)) {
        lastVal = meses[i].actual ?? meses[i].final;
        // Previous = inicio of same month
        prevVal = meses[i].inicio;
        break;
      }
    }
    const elActual = document.getElementById('stock-actual-' + key);
    elActual.textContent = lastVal != null ? lastVal.toLocaleString('es-ES') + ' T' : '—';

    const elDelta = document.getElementById('stock-delta-' + key);
    if (lastVal != null && prevVal != null && prevVal !== 0) {
      const pct = ((lastVal - prevVal) / prevVal * 100).toFixed(1);
      const up = pct >= 0;
      elDelta.innerHTML = `<span style="color:${up ? '#27ae60' : '#e74c3c'}">${up ? '▲' : '▼'} ${Math.abs(pct)}%</span> <span style="color:var(--muted)">vs inicio mes</span>`;
    } else {
      elDelta.textContent = '';
    }
  }

  // --- Evolution line chart (all products over time) ---
  if (_stockCharts.evolucion) _stockCharts.evolucion.destroy();
  const ctxEvo = document.getElementById('stock-chart-evolucion').getContext('2d');

  // Use monthly final values for all keys
  const allLabels = (_stockData['04'] || []).filter(d => d.inicio != null || d.final != null).map(d => d.mes);
  const evoDatasets = keys.map(key => {
    const datos = (_stockData[key] || []).filter(d => d.inicio != null || d.final != null);
    return {
      label: STOCK_LABELS[key],
      data: datos.map(d => d.final ?? d.inicio),
      borderColor: STOCK_COLORS[key].line,
      backgroundColor: STOCK_COLORS[key].bg,
      tension: 0.35,
      pointRadius: 3,
      pointHoverRadius: 6,
      borderWidth: 2.5,
      fill: false
    };
  });

  _stockCharts.evolucion = new Chart(ctxEvo, {
    type: 'line',
    data: { labels: allLabels, datasets: evoDatasets },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { position: 'bottom', labels: { usePointStyle: true, padding: 14, font: { size: 11 } } },
        tooltip: { callbacks: { label: c => c.dataset.label + ': ' + (c.parsed.y?.toLocaleString('es-ES') ?? '—') + ' T' } }
      },
      scales: {
        x: { grid: { display: false } },
        y: { beginAtZero: false, ticks: { callback: v => v.toLocaleString('es-ES') } }
      },
      animation: { duration: 1200, easing: 'easeOutQuart' },
      interaction: { intersect: false, mode: 'index' }
    }
  });

  // --- Grouped bar chart (latest values per product) ---
  if (_stockCharts.barras) _stockCharts.barras.destroy();
  const ctxBar = document.getElementById('stock-chart-barras').getContext('2d');

  const barLabels = keys.map(k => STOCK_LABELS[k]);
  const inicioVals = keys.map(k => {
    const datos = (_stockData[k] || []).filter(d => d.inicio != null || d.final != null);
    return datos.length ? datos[datos.length - 1].inicio : null;
  });
  const finalVals = keys.map(k => {
    const datos = (_stockData[k] || []).filter(d => d.inicio != null || d.final != null);
    return datos.length ? datos[datos.length - 1].final : null;
  });

  _stockCharts.barras = new Chart(ctxBar, {
    type: 'bar',
    data: {
      labels: barLabels,
      datasets: [
        { label: 'Inicio mes', data: inicioVals, backgroundColor: 'rgba(150,150,150,.35)', borderColor: 'rgba(150,150,150,.7)', borderWidth: 1, borderRadius: 4 },
        { label: 'Último', data: finalVals, backgroundColor: keys.map(k => STOCK_COLORS[k].bg.replace('.25)', '.6)')), borderColor: keys.map(k => STOCK_COLORS[k].line), borderWidth: 1.5, borderRadius: 4 }
      ]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { position: 'bottom', labels: { usePointStyle: true, padding: 14, font: { size: 11 } } },
        tooltip: { callbacks: { label: c => c.dataset.label + ': ' + (c.parsed.y?.toLocaleString('es-ES') ?? '—') + ' T' } }
      },
      scales: {
        x: { grid: { display: false } },
        y: { beginAtZero: false, ticks: { callback: v => v.toLocaleString('es-ES') } }
      },
      animation: { duration: 1200, easing: 'easeOutQuart' },
      interaction: { intersect: false, mode: 'index' }
    }
  });
}

function abrirStockDetalle(key) {
  document.getElementById('stock-overview').style.display = 'none';
  document.getElementById('stock-detalle').style.display = 'block';
  document.getElementById('stock-detalle-titulo').textContent = 'Stock ' + STOCK_LABELS[key];

  const datosAll = (_stockData && _stockData[key]) || [];
  const datos = datosAll.filter(d => d.inicio != null || d.final != null);
  const ctx = document.getElementById('stock-chart-grande').getContext('2d');

  if (_stockCharts.grande) _stockCharts.grande.destroy();

  const labels = datos.map(d => d.mes);
  const gradient = ctx.createLinearGradient(0, 0, 0, 300);
  gradient.addColorStop(0, STOCK_COLORS[key].bg);
  gradient.addColorStop(1, 'rgba(255,255,255,0)');

  const datasets = [];

  // Inicio mes
  datasets.push({
    label: 'Inicio mes',
    data: datos.map(d => d.inicio),
    backgroundColor: 'rgba(150,150,150,.3)',
    borderColor: 'rgba(150,150,150,.6)',
    borderWidth: 1,
    borderRadius: 3
  });

  // Final mes
  datasets.push({
    label: 'Final mes',
    data: datos.map(d => d.final),
    backgroundColor: STOCK_COLORS[key].bg,
    borderColor: STOCK_COLORS[key].line,
    borderWidth: 1.5,
    borderRadius: 4
  });

  // Actual
  datasets.push({
    label: 'Actual',
    data: datos.map(d => d.actual),
    backgroundColor: 'rgba(231,76,60,.6)',
    borderColor: '#e74c3c',
    borderWidth: 1.5,
    borderRadius: 4
  });

  _stockCharts.grande = new Chart(ctx, {
    type: 'bar',
    data: { labels, datasets },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: 'top', labels: { usePointStyle: true, padding: 16, font: { size: 12 } } },
        tooltip: { callbacks: { label: c => c.dataset.label + ': ' + (c.parsed.y?.toLocaleString('es-ES') ?? '—') + ' Tn' } }
      },
      scales: {
        x: { grid: { display: false } },
        y: { beginAtZero: false, ticks: { callback: v => v.toLocaleString('es-ES') } }
      },
      animation: { duration: 1500, easing: 'easeOutQuart' },
      interaction: { intersect: false, mode: 'index' }
    }
  });

  // Tabla
  const tbody = document.querySelector('#stock-tabla-detalle tbody');
  tbody.innerHTML = '';
  for (const d of datos) {
    const tr = document.createElement('tr');
    tr.style.borderBottom = '1px solid var(--border)';
    tr.innerHTML = `<td style="padding:6px 8px">${d.mes}</td>
      <td style="text-align:right;padding:6px 8px">${d.inicio != null ? d.inicio.toLocaleString('es-ES') : '—'}</td>
      <td style="text-align:right;padding:6px 8px">${d.final != null ? d.final.toLocaleString('es-ES') : '—'}</td>
      <td style="text-align:right;padding:6px 8px;font-weight:600;color:var(--accent)">${d.actual != null ? d.actual.toLocaleString('es-ES') : '—'}</td>`;
    tbody.appendChild(tr);
  }
}

function cerrarStockDetalle() {
  document.getElementById('stock-detalle').style.display = 'none';
  document.getElementById('stock-overview').style.display = '';
}

// ── TAREAS ────────────────────────────────────────────────────────────────────
let tareasData = [];
let _tareasPanelInit = false;
let _tareaEditId = null;
let _tareasVista = 'kanban'; // 'lista' | 'kanban'
let _draggingTareaId = null;

function setTareasVista(v){
  _tareasVista = v;
  renderTareas();
}

function initTareasPanel(){
  // Poblar select personas
  const sel = document.getElementById('tarea-filt-persona');
  if(sel.options.length <= 2){
    for(const w of WORKERS){
      const o = document.createElement('option');
      o.value = w; o.textContent = w;
      sel.appendChild(o);
    }
  }
  _tareasPanelInit = true;
  cargarTareas();
}

async function cargarTareas(){
  document.getElementById('tareas-list').innerHTML = '<div style="color:var(--muted);text-align:center;padding:30px;font-size:.82rem">Cargando...</div>';
  const res = await dbQuery({ action:'select', table:'tblTareas', options:{ select:'*', order:'created_at.desc', limit:500 } });
  if(!res.ok){ document.getElementById('tareas-list').innerHTML = '<div style="color:#e53935;text-align:center;padding:30px;font-size:.82rem">Error: '+escapeHTML(res.error)+'</div>'; return; }
  tareasData = res.data || [];
  renderTareas();
}

function _tareasAsignados(t){
  // asignado puede ser "Ana,Luis" o "Ana" — devuelve array
  if(!t.asignado) return [];
  return t.asignado.split(',').map(s=>s.trim()).filter(Boolean);
}
const TAREAS_SECCIONES = ['Planta','Maquinaria','Administración'];
const TAREAS_SECCIONES_ACTIVO = ['Planta','Maquinaria']; // secciones con selector de activo

function _filtrarTareas(){
  const filtPersona  = document.getElementById('tarea-filt-persona')?.value||'';
  const filtSeccion  = document.getElementById('tarea-filt-seccion')?.value||'';
  const filtActivo   = document.getElementById('tarea-filt-activo')?.value||'';
  return tareasData.filter(t => {
    if(filtSeccion && t.seccion !== filtSeccion) return false;
    if(filtActivo && t.activo !== filtActivo) return false;
    if(filtPersona === '__mias__' && loginUser){
      if(!_tareasAsignados(t).includes(loginUser.nombre)) return false;
    } else if(filtPersona && filtPersona !== '__mias__'){
      if(!_tareasAsignados(t).includes(filtPersona)) return false;
    }
    return true;
  });
}

function goTareasSeccion(seccion){
  goPage('tareas');
  setTimeout(()=>{
    const sel=document.getElementById('tarea-filt-seccion');
    if(sel){sel.value=seccion;tareaSeccionChange();}
  },100);
}
function tareaSeccionChange(){
  const sec = document.getElementById('tarea-filt-seccion')?.value||'';
  const selActivo = document.getElementById('tarea-filt-activo');
  const card = document.getElementById('tarea-activo-card');
  if(TAREAS_SECCIONES_ACTIVO.includes(sec)){
    // Poblar select activos
    if(!activosData.length){
      dbQuery({action:'select',table:'tblactivos',options:{select:'*',order:'Codigo.asc',limit:500}}).then(j=>{
        if(j.ok && j.data) activosData=j.data;
        _poblarSelectActivoTareas();
      });
    } else {
      _poblarSelectActivoTareas();
    }
    selActivo.style.display='';
  } else {
    selActivo.style.display='none';
    selActivo.value='';
    if(card) card.style.display='none';
  }
  renderTareas();
}

function _poblarSelectActivoTareas(){
  const sel = document.getElementById('tarea-filt-activo');
  if(!sel) return;
  sel.innerHTML='<option value="">Todos los activos</option>'+
    activosData.filter(a=>a.Codigo).map(a=>`<option value="${a.Codigo}">${a.Codigo}${a.Activo&&a.Activo!==a.Codigo?' — '+a.Activo:''}</option>`).join('');
}

function tareaActivoChange(){
  const codigo = document.getElementById('tarea-filt-activo')?.value||'';
  const card = document.getElementById('tarea-activo-card');
  if(codigo && card){
    const a = activosData.find(x=>x.Codigo===codigo);
    if(a){
      // Horómetro: buscar en prevGasoilHoroMap o machineHoroOT
      const horo = (typeof prevGasoilHoroMap!=='undefined'&&prevGasoilHoroMap[codigo]) || '';
      const items = [
        ['Código', a.Codigo],
        ['Nombre', a.Activo||'—'],
        ['Modelo', a.modelo||'—'],
        ['Fabricante', a.fabricante||'—'],
        horo ? ['Horómetro', horo+' h'] : null,
        a.matricula ? ['Matrícula', a.matricula] : null,
      ].filter(Boolean);
      card.innerHTML = items.map(([k,v])=>`<div style="font-size:.74rem"><span style="color:var(--muted);font-size:.68rem;display:block">${k}</span><strong>${v}</strong></div>`).join('');
      card.style.display='flex';
    }
  } else if(card){
    card.style.display='none';
  }
  renderTareas();
}

function tmSeccionChange(){
  const sec = document.getElementById('tm-seccion')?.value||'';
  const row = document.getElementById('tm-activo-row');
  if(!row) return;
  if(TAREAS_SECCIONES_ACTIVO.includes(sec)){
    row.style.display='';
    const sel = document.getElementById('tm-activo');
    if(sel.options.length<=1){
      sel.innerHTML='<option value="">— Sin activo específico —</option>'+
        activosData.filter(a=>a.Codigo).map(a=>`<option value="${a.Codigo}">${a.Codigo}${a.Activo&&a.Activo!==a.Codigo?' — '+a.Activo:''}</option>`).join('');
    }
  } else {
    row.style.display='none';
  }
}

function _renderTareaCard(t, hoy){
  const done = t.estado === 'hecha';
  const vencida = !done && t.fecha_limite && t.fecha_limite < hoy;
  const priLabel = {alta:'Alta',media:'Media',baja:'Baja'}[t.prioridad] || t.prioridad;
  const fechaStr = t.fecha_limite ? t.fecha_limite.split('-').reverse().join('/') : '';
  return `<div class="tarea-card${done?' done':''}" onclick="abrirModalTarea(${JSON.stringify(t).replace(/"/g,'&quot;')})">
    <div class="tarea-header">
      <div class="tarea-check${done?' done':''}" onclick="event.stopPropagation();toggleEstadoTarea(${t.id},'${t.estado}')" title="Cambiar estado">
        ${done ? '<svg width="12" height="12" viewBox="0 0 12 12"><path d="M2 6l3 3 5-5" stroke="currentColor" stroke-width="2" fill="none"/></svg>' : ''}
      </div>
      <div class="tarea-titulo">${escapeHTML(t.titulo)}</div>
    </div>
    <div class="tarea-meta">
      <span class="tarea-badge ${t.prioridad||'media'}">${priLabel}</span>
      ${_tareasAsignados(t).map(a=>`<span class="tarea-badge persona">${escapeHTML(a)}</span>`).join('')||'<span class="tarea-badge persona">—</span>'}
      ${fechaStr ? `<span class="tarea-badge ${vencida?'vencida':'vence'}">📅 ${fechaStr}</span>` : ''}
    </div>
    ${t.descripcion ? `<div class="tarea-desc">${escapeHTML(t.descripcion)}</div>` : ''}
  </div>`;
}

function renderTareas(){
  if(_tareasVista === 'kanban'){ renderTareasKanban(); return; }

  const lista = _filtrarTareas();
  const hoy = new Date().toISOString().slice(0,10);
  const wrap = document.getElementById('tareas-list');

  if(!lista.length){
    wrap.innerHTML = '<div style="color:var(--muted);text-align:center;padding:40px;font-size:.82rem">Sin tareas para los filtros seleccionados</div>';
    return;
  }

  const priOrd = {alta:0,media:1,baja:2};
  const activas = lista.filter(t => t.estado !== 'hecha').sort((a,b)=>(priOrd[a.prioridad]||1)-(priOrd[b.prioridad]||1)||(a.fecha_limite||'9').localeCompare(b.fecha_limite||'9'));
  const hechas  = lista.filter(t => t.estado === 'hecha');

  let html = activas.map(t=>_renderTareaCard(t,hoy)).join('');
  if(hechas.length){
    html += `<div class="tarea-group-title">Completadas (${hechas.length})</div>`;
    html += hechas.map(t=>_renderTareaCard(t,hoy)).join('');
  }
  wrap.innerHTML = html;
}

function renderTareasKanban(){
  const lista = _filtrarTareas();
  const hoy = new Date().toISOString().slice(0,10);
  const priOrd = {alta:0,media:1,baja:2};
  const colInfo = [
    {key:'pendiente', label:'Pendiente', icon:'○'},
    {key:'en_curso',  label:'En curso',  icon:'◑'},
    {key:'hecha',     label:'Hecha',     icon:'●'}
  ];
  const filtSeccion = document.getElementById('tarea-filt-seccion')?.value||'';

  const renderKanbanCard = t => {
    const vencida = t.estado!=='hecha' && t.fecha_limite && t.fecha_limite < hoy;
    const priLabel = {alta:'Alta',media:'Media',baja:'Baja'}[t.prioridad] || t.prioridad;
    const fechaStr = t.fecha_limite ? t.fecha_limite.split('-').reverse().join('/') : '';
    return `<div class="kanban-card${t.estado==='hecha'?' done':''}"
      draggable="true"
      data-id="${t.id}"
      ondragstart="kanbanDragStart(event,${t.id})"
      ondragend="kanbanDragEnd(event)"
      onclick="abrirModalTarea(${JSON.stringify(t).replace(/"/g,'&quot;')})">
      ${t.activo ? `<div style="font-size:.63rem;font-family:monospace;color:var(--accent);font-weight:700;margin-bottom:3px">${escapeHTML(t.activo)}</div>` : ''}
      <div class="kanban-card-title">${escapeHTML(t.titulo)}</div>
      <div class="kanban-card-meta">
        <span class="tarea-badge ${t.prioridad||'media'}">${priLabel}</span>
        ${_tareasAsignados(t).map(a=>`<span class="tarea-badge persona">${escapeHTML(a)}</span>`).join('')||'<span class="tarea-badge persona">—</span>'}
        ${fechaStr ? `<span class="tarea-badge ${vencida?'vencida':'vence'}">📅 ${fechaStr}</span>` : ''}
      </div>
      ${t.descripcion ? `<div class="kanban-card-desc">${escapeHTML(t.descripcion)}</div>` : ''}
    </div>`;
  };

  const buildBoard = (tareas, seccion) => {
    const cols = {pendiente:[], en_curso:[], hecha:[]};
    tareas.forEach(t => { if(cols[t.estado]) cols[t.estado].push(t); });
    Object.values(cols).forEach(arr => arr.sort((a,b)=>(priOrd[a.prioridad]||1)-(priOrd[b.prioridad]||1)||(a.fecha_limite||'9').localeCompare(b.fecha_limite||'9')));
    const board = document.createElement('div');
    board.className = 'kanban-board';
    colInfo.forEach(({key, label, icon}) => {
      const cards = cols[key];
      const col = document.createElement('div');
      col.className = `kanban-col ${key}`;
      col.dataset.estado = key;
      col.dataset.seccion = seccion;
      col.innerHTML = `<div class="kanban-col-header">
        <span class="kanban-col-title">${icon} ${label}</span>
        <span class="kanban-col-count">${cards.length}</span>
      </div>` + (cards.length ? cards.map(renderKanbanCard).join('') : `<div style="color:var(--muted);font-size:.74rem;text-align:center;padding:16px 8px;font-style:italic">Sin tareas</div>`);
      col.addEventListener('dragover', e=>{ e.preventDefault(); col.classList.add('drag-over'); });
      col.addEventListener('dragleave', ()=> col.classList.remove('drag-over'));
      col.addEventListener('drop', e=>{ e.preventDefault(); col.classList.remove('drag-over'); kanbanDrop(key); });
      board.appendChild(col);
    });
    return board;
  };

  const wrap = document.getElementById('tareas-list');
  wrap.innerHTML = '';

  // Secciones a mostrar
  const secciones = filtSeccion ? [filtSeccion] : TAREAS_SECCIONES;
  const seccionColors = {'Planta':'#4caf50','Maquinaria':'#2196f3','Administración':'#f5a623'};

  secciones.forEach(sec => {
    const tareasSeccion = lista.filter(t => (t.seccion||'Planta') === sec);
    const secDiv = document.createElement('div');
    secDiv.style.cssText = 'margin-bottom:24px;width:100%';
    const header = document.createElement('div');
    header.style.cssText = `display:flex;align-items:center;gap:8px;margin-bottom:8px;padding:6px 10px;border-radius:6px;background:${seccionColors[sec]||'var(--accent)'}18;border-left:3px solid ${seccionColors[sec]||'var(--accent)'}`;
    header.innerHTML = `<span style="font-weight:700;font-size:.82rem;color:${seccionColors[sec]||'var(--accent)'}">${sec}</span><span style="font-size:.72rem;color:var(--muted)">${tareasSeccion.length} tareas</span>`;
    secDiv.appendChild(header);
    secDiv.appendChild(buildBoard(tareasSeccion, sec));
    wrap.appendChild(secDiv);
  });
}

function kanbanDragStart(event, id){
  _draggingTareaId = id;
  event.dataTransfer.effectAllowed = 'move';
  setTimeout(()=>{ const el=document.querySelector(`.kanban-card[data-id="${id}"]`); if(el) el.classList.add('dragging'); }, 0);
}
function kanbanDragEnd(event){
  document.querySelectorAll('.kanban-card.dragging').forEach(el=>el.classList.remove('dragging'));
}
async function kanbanDrop(nuevoEstado){
  if(!_draggingTareaId) return;
  const id = _draggingTareaId;
  _draggingTareaId = null;
  const t = tareasData.find(x=>x.id===id);
  if(!t || t.estado === nuevoEstado) return;
  t.estado = nuevoEstado; // optimistic update
  renderTareasKanban();
  await dbQuery({ action:'update', table:'tblTareas', data:{estado:nuevoEstado}, filters:[{column:'id',op:'eq',value:id}] });
}

async function toggleEstadoTarea(id, estadoActual){
  const sig = estadoActual === 'hecha' ? 'pendiente' : estadoActual === 'pendiente' ? 'en_curso' : 'hecha';
  const res = await dbQuery({ action:'update', table:'tblTareas', data:{estado:sig}, filters:[{column:'id',op:'eq',value:id}] });
  if(res.ok){ const t = tareasData.find(x=>x.id===id); if(t) t.estado=sig; renderTareas(); }
}

function abrirModalTarea(tarea){
  _tareaEditId = tarea ? tarea.id : null;
  // Poblar select personas en modal
  const sel = document.getElementById('tm-persona');
  if(sel.options.length === 0){
    for(const w of WORKERS){
      const o = document.createElement('option');
      o.value = w; o.textContent = w;
      sel.appendChild(o);
    }
  }
  // Seleccionar asignados (multi)
  const asignados = tarea ? _tareasAsignados(tarea) : (loginUser ? [loginUser.nombre] : []);
  for(const opt of sel.options) opt.selected = asignados.includes(opt.value);
  // Rellenar campos
  document.getElementById('tm-seccion').value   = tarea ? (tarea.seccion||'Planta') : (document.getElementById('tarea-filt-seccion')?.value||'Planta');
  tmSeccionChange();
  document.getElementById('tm-activo').value    = tarea ? (tarea.activo||'') : (document.getElementById('tarea-filt-activo')?.value||'');
  document.getElementById('tm-titulo').value    = tarea ? tarea.titulo : '';
  document.getElementById('tm-desc').value      = tarea ? (tarea.descripcion||'') : '';
  document.getElementById('tm-prioridad').value = tarea ? (tarea.prioridad||'media') : 'media';
  document.getElementById('tm-fecha').value     = tarea ? (tarea.fecha_limite||'') : '';
  document.getElementById('tm-estado').value    = tarea ? tarea.estado : 'pendiente';
  document.getElementById('tm-recurrente').checked = false; // solo para nueva
  document.getElementById('tareas-modal-title').textContent = tarea ? 'Editar tarea' : 'Nueva tarea';
  const btnBorrar = document.getElementById('tm-btn-borrar');
  btnBorrar.style.display = (tarea && loginUser && loginUser.rol==='admin') ? '' : 'none';
  document.getElementById('tareas-modal-wrap').style.display = 'flex';
}

function cerrarModalTarea(){
  document.getElementById('tareas-modal-wrap').style.display = 'none';
}

async function guardarTarea(){
  const titulo = document.getElementById('tm-titulo').value.trim();
  if(!titulo){ alert('El título es obligatorio'); return; }
  const selPer = document.getElementById('tm-persona');
  const asignados = [...selPer.selectedOptions].map(o=>o.value).filter(Boolean).join(',');
  if(!asignados){ alert('Selecciona al menos una persona'); return; }
  const data = {
    seccion:     document.getElementById('tm-seccion').value,
    activo:      document.getElementById('tm-activo')?.value||null,
    titulo,
    descripcion: document.getElementById('tm-desc').value.trim() || null,
    asignado:    asignados,
    prioridad:   document.getElementById('tm-prioridad').value,
    fecha_limite: document.getElementById('tm-fecha').value || null,
    estado:      document.getElementById('tm-estado').value,
  };
  const recurrente = !_tareaEditId && document.getElementById('tm-recurrente').checked;
  let res;
  if(_tareaEditId){
    res = await dbQuery({ action:'update', table:'tblTareas', data, filters:[{column:'id',op:'eq',value:_tareaEditId}] });
    if(!res.ok){ alert('Error al guardar: '+res.error); return; }
  } else {
    data.creado_por = loginUser ? loginUser.nombre : '';
    if(recurrente){
      // Crear una entrada por cada mes del año actual
      const anyo = new Date().getFullYear();
      const base = data.fecha_limite ? data.fecha_limite.slice(0,7) : null;
      const baseMonth = base ? parseInt(base.split('-')[1]) : 1;
      const rows = [];
      for(let m=1; m<=12; m++){
        const mm = String(m).padStart(2,'0');
        const lastDay = new Date(anyo, m, 0).getDate();
        rows.push({...data, fecha_limite:`${anyo}-${mm}-${String(lastDay).padStart(2,'0')}`, titulo:`${data.titulo} (${mm}/${anyo})`});
      }
      // Insert one by one (proxy doesn't support batch easily)
      for(const row of rows){
        const r = await dbQuery({ action:'insert', table:'tblTareas', data:row });
        if(!r.ok){ alert('Error al guardar: '+r.error); return; }
      }
    } else {
      res = await dbQuery({ action:'insert', table:'tblTareas', data });
      if(!res.ok){ alert('Error al guardar: '+res.error); return; }
    }
  }
  cerrarModalTarea();
  // Mostrar todas tras guardar para que el resultado sea visible
  const filtEl = document.getElementById('tarea-filt-persona');
  if(filtEl && !_tareaEditId) filtEl.value = '';
  cargarTareas();
}

async function borrarTareaModal(){
  if(!_tareaEditId) return;
  if(!confirm('¿Eliminar esta tarea?')) return;
  const res = await dbQuery({ action:'delete', table:'tblTareas', filters:[{column:'id',op:'eq',value:_tareaEditId}] });
  if(!res.ok){ alert('Error al eliminar: '+res.error); return; }
  cerrarModalTarea();
  cargarTareas();
}

// ============================================================
// CONTROL DE ENSAYOS
// ============================================================

let _ensayosTab = 'control';
let _ensayosAnio = new Date().getFullYear();
let _ensayosSemanas = [];
let _ensayosRegistros = [];
let _ensayosPrestaciones = [];
let _ensayosLimites = [];
let _ensayosFraccion = '0/4';

// API
async function getEnsayosSemanas(anio) {
  const res = await dbQuery({ action:'select', table:'ensayos_semanas', options:{ select:'*', order:'fecha_lunes.desc', limit:500 } });
  if (!res.ok) return res;
  const desde = anio + '-01-01';
  const hasta = anio + '-12-31';
  res.data = res.data.filter(function(r){ const f = (r.fecha_lunes||'').slice(0,10); return f >= desde && f <= hasta; });
  return res;
}
async function insertEnsayoSemana(data) {
  return dbQuery({ action:'insert', table:'ensayos_semanas', data });
}
async function updateEnsayoSemana(id, data) {
  return dbQuery({ action:'update', table:'ensayos_semanas', data, filters:[{column:'id',op:'eq',value:id}] });
}
async function getEnsayosRegistros(anio) {
  const semanaIds = _ensayosSemanas.map(function(s){ return s.id; });
  if (!semanaIds.length) return { ok:true, data:[] };
  return dbQuery({ action:'select', table:'ensayos_registros', options:{ select:'*', order:'semana_id.asc', limit:1000 }, filters:[{ column:'semana_id', op:'in', value:semanaIds }] });
}
async function insertEnsayoRegistro(data) {
  return dbQuery({ action:'insert', table:'ensayos_registros', data });
}
async function updateEnsayoRegistro(id, data) {
  return dbQuery({ action:'update', table:'ensayos_registros', data, filters:[{column:'id',op:'eq',value:id}] });
}
async function deleteEnsayoRegistro(id) {
  return dbQuery({ action:'delete', table:'ensayos_registros', filters:[{column:'id',op:'eq',value:id}] });
}
async function getEnsayosPrestaciones() {
  return dbQuery({ action:'select', table:'ensayos_prestaciones', options:{ select:'*', order:'fraccion.asc' } });
}
async function updateEnsayoPrestacion(id, data) {
  return dbQuery({ action:'update', table:'ensayos_prestaciones', data, filters:[{column:'id',op:'eq',value:id}] });
}
async function getEnsayosLimites() {
  return dbQuery({ action:'select', table:'ensayos_limites', options:{ select:'*' } });
}

// INIT
async function initEnsayos() {
  _ensayosAnio = new Date().getFullYear();
  const anioSel = document.getElementById('ensayos-anio');
  if (anioSel) anioSel.value = _ensayosAnio;
  await _ensayosCargarTodo();
  _ensayosRenderTab('control');
}

async function _ensayosCargarTodo() {
  // Semanas primero — registros dependen de los semana_id
  const rSem = await getEnsayosSemanas(_ensayosAnio);
  _ensayosSemanas = rSem.ok ? rSem.data : [];
  const [rReg, rPrest, rLim] = await Promise.all([
    getEnsayosRegistros(_ensayosAnio),
    getEnsayosPrestaciones(),
    getEnsayosLimites()
  ]);
  _ensayosRegistros = rReg.ok ? rReg.data : [];
  _ensayosPrestaciones = rPrest.ok ? rPrest.data : [];
  _ensayosLimites = rLim.ok ? rLim.data : [];
}

function _ensayosRenderTab(tab) {
  _ensayosTab = tab;
  ['control','registros','anuales','prestaciones'].forEach(function(t) {
    const btn = document.getElementById('ensayos-tab-' + t);
    if (!btn) return;
    btn.classList.toggle('active', t === tab);
    const on = t === tab;
    btn.style.background = on ? 'var(--accent)' : 'var(--surface)';
    btn.style.color = on ? '#fff' : 'var(--text)';
    btn.style.fontWeight = on ? '600' : '400';
  });
  const body = document.getElementById('ensayos-body');
  if (!body) return;
  if (tab === 'control') body.innerHTML = _ensayosRenderControl();
  else if (tab === 'registros') body.innerHTML = _ensayosRenderRegistros();
  else if (tab === 'anuales') body.innerHTML = _ensayosRenderAnuales();
  else if (tab === 'prestaciones') body.innerHTML = _ensayosRenderPrestaciones();
}

// TAB CONTROL
function _ensayosRenderControl() {
  const semanas = _ensayosSemanas;
  const TH = 'padding:5px 8px;border:1px solid #d0d7c8;white-space:nowrap;font-size:.72rem;';
  const TH_GRP = 'padding:5px 8px;border:1px solid #d0d7c8;text-align:center;font-size:.72rem;font-weight:700;';
  const BG_HDR = 'background:#2c3a2c;color:#fff;';
  const BG_GRAN = 'background:#3a4a3a;color:#fff;';
  const BG_FINO = 'background:#3a4a30;color:#fff;';
  const BG_EQ   = 'background:#2a3a4a;color:#fff;';
  const BG_LAJ  = 'background:#4a3a2a;color:#fff;';
  const BG_CAR  = 'background:#4a2a3a;color:#fff;';

  let html = '<div style="font-size:.8rem;color:var(--muted);margin-bottom:8px">' + semanas.length + ' semanas</div>';
  html += '<div style="overflow-x:auto"><table style="border-collapse:collapse;font-size:.78rem;min-width:1100px;background:var(--surface)">';

  // Fila 1: grupos con colspan
  html += '<thead>';
  html += '<tr>';
  html += '<th colspan="2" style="' + TH + BG_HDR + '"></th>';
  html += '<th style="' + TH + BG_HDR + '">N\u00ba</th>';
  html += '<th style="' + TH + BG_HDR + '">FECHA</th>';
  html += '<th style="' + TH + BG_HDR + '">MAT.</th>';
  html += '<th style="' + TH + BG_HDR + '">TN TOTAL</th>';
  html += '<th colspan="4" style="' + TH_GRP + BG_HDR + '">PRODUCCI\u00d3N (TN)</th>';
  html += '<th colspan="4" style="' + TH_GRP + BG_GRAN + '">GRANULOMETR\u00cdA<br><span style="font-weight:400;font-size:.68rem">UNE-EN 933-1 \u00b7 SEMANAL</span></th>';
  html += '<th colspan="4" style="' + TH_GRP + BG_FINO + '">CONT. DE FINOS<br><span style="font-weight:400;font-size:.68rem">UNE-EN 933-1 \u00b7 SEMANAL</span></th>';
  html += '<th colspan="2" style="' + TH_GRP + BG_EQ + '">EQ. ARENA<br><span style="font-weight:400;font-size:.68rem">UNE-EN 933-8 \u00b7 SEMANAL</span></th>';
  html += '<th colspan="4" style="' + TH_GRP + BG_LAJ + '">\u00cdNDICE LAJAS<br><span style="font-weight:400;font-size:.68rem">UNE-EN 933-3 \u00b7 MENSUAL</span></th>';
  html += '<th colspan="3" style="' + TH_GRP + BG_CAR + '">% CAPAS FRAG.<br><span style="font-weight:400;font-size:.68rem">UNE-EN 933-5 \u00b7 MENSUAL</span></th>';
  html += '</tr>';

  // Fila 2: subcolumnas
  html += '<tr>';
  html += '<th style="' + TH + BG_HDR + '">EST.</th>';
  html += '<th style="' + TH + BG_HDR + '">REC.</th>';
  html += '<th style="' + TH + BG_HDR + '"></th>';
  html += '<th style="' + TH + BG_HDR + '"></th>';
  html += '<th style="' + TH + BG_HDR + '"></th>';
  html += '<th style="' + TH + BG_HDR + '"></th>';
  ['0/4','4/12','12/20','20/40'].forEach(function(f){ html += '<th style="' + TH + BG_HDR + '">' + f + '</th>'; });
  ['0/4','4/12','12/20','20/40'].forEach(function(f){ html += '<th style="' + TH + BG_GRAN + '">' + f + '</th>'; });
  ['0/4','4/12','12/20','20/40'].forEach(function(f){ html += '<th style="' + TH + BG_FINO + '">' + f + '</th>'; });
  ['0/4','ZA25'].forEach(function(f){ html += '<th style="' + TH + BG_EQ + '">' + f + '</th>'; });
  ['4/12','12/20','20/40','ZA25'].forEach(function(f){ html += '<th style="' + TH + BG_LAJ + '">' + f + '</th>'; });
  ['4/12','12/20','20/40'].forEach(function(f){ html += '<th style="' + TH + BG_CAR + '">' + f + '</th>'; });
  html += '</tr>';
  html += '</thead><tbody>';

  semanas.forEach(function(sem, i) {
    const num = semanas.length - i;
    const fecha = sem.fecha_lunes ? sem.fecha_lunes.slice(0,10) : '\u2014';
    const tnTotal = (sem.tn_04||0) + (sem.tn_412||0) + (sem.tn_1220||0) + (sem.tn_2040||0);
    const regs = _ensayosRegistros.filter(function(r){ return r.semana_id === sem.id; });

    function estadoReg(tipo, frac) {
      const r = regs.find(function(r){ return r.tipo_ensayo === tipo && r.fraccion === frac; });
      if (!r) return '<span style="color:#ccc;font-size:.8rem">+</span>';
      if (r.estado === 'conforme') return '<span style="color:#2e7d32;font-size:1rem;font-weight:700">\u2713\u2713</span>';
      if (r.estado === 'no_conforme') return '<span style="color:#c62828;font-size:1rem;font-weight:700">\u2717</span>';
      return '<span style="color:#e65100;font-size:1rem;font-weight:700">\u2713</span>';
    }

    let estadoGlobal = 'NP', estadoColor = 'var(--muted)';
    if (regs.length > 0) {
      if (regs.every(function(r){ return r.estado === 'conforme'; })) { estadoGlobal = 'C'; estadoColor = '#4caf50'; }
      else if (regs.some(function(r){ return r.estado === 'no_conforme'; })) { estadoGlobal = 'NC'; estadoColor = '#f44336'; }
      else { estadoGlobal = 'P'; estadoColor = '#ff9800'; }
    }
    const mat = sem.tipo_material || 'NP';
    const matColor = mat === 'AC' ? 'var(--accent)' : 'var(--muted)';

    const TD = 'padding:5px 8px;border:1px solid #e8ede4;';
    const TD_C = TD + 'text-align:center;';
    const TD_R = TD + 'text-align:right;';
    const BG_G = 'background:#eef4ea;';
    const BG_F = 'background:#f0f4e8;';
    const BG_E = 'background:#e8eef4;';
    const BG_L = 'background:#f4ede8;';
    const BG_Ca= 'background:#f4e8ee;';
    const rowBg = i % 2 === 0 ? '' : 'background:#f9faf8;';

    html += '<tr style="' + rowBg + 'cursor:pointer" onclick="ensayosAbrirSemana(\'' + sem.id + '\')">';
    html += '<td style="' + TD_C + 'color:' + estadoColor + ';font-weight:700;font-size:.7rem">' + estadoGlobal + '</td>';
    html += '<td style="' + TD_C + 'color:var(--muted)">\u2014</td>';
    html += '<td style="' + TD_C + 'font-weight:600">' + num + '</td>';
    html += '<td style="' + TD + 'white-space:nowrap">' + fecha + '</td>';
    html += '<td style="' + TD_C + 'color:' + matColor + ';font-weight:600">' + mat + '</td>';
    html += '<td style="' + TD_R + '">' + (tnTotal ? Number(tnTotal).toLocaleString('es') : '\u2014') + '</td>';
    html += '<td style="' + TD_R + '">' + (sem.tn_04 ? Number(sem.tn_04).toLocaleString('es') : '\u2014') + '</td>';
    html += '<td style="' + TD_R + '">' + (sem.tn_412 ? Number(sem.tn_412).toLocaleString('es') : '\u2014') + '</td>';
    html += '<td style="' + TD_R + '">' + (sem.tn_1220 ? Number(sem.tn_1220).toLocaleString('es') : '\u2014') + '</td>';
    html += '<td style="' + TD_R + '">' + (sem.tn_2040 ? Number(sem.tn_2040).toLocaleString('es') : '\u2014') + '</td>';
    html += '<td style="' + TD_C + BG_G + '">' + estadoReg('granulometria','0/4') + '</td>';
    html += '<td style="' + TD_C + BG_G + '">' + estadoReg('granulometria','4/12') + '</td>';
    html += '<td style="' + TD_C + BG_G + '">' + estadoReg('granulometria','12/20') + '</td>';
    html += '<td style="' + TD_C + BG_G + '">' + estadoReg('granulometria','20/40') + '</td>';
    html += '<td style="' + TD_C + BG_F + '">' + estadoReg('cont_finos','0/4') + '</td>';
    html += '<td style="' + TD_C + BG_F + '">' + estadoReg('cont_finos','4/12') + '</td>';
    html += '<td style="' + TD_C + BG_F + '">' + estadoReg('cont_finos','12/20') + '</td>';
    html += '<td style="' + TD_C + BG_F + '">' + estadoReg('cont_finos','20/40') + '</td>';
    html += '<td style="' + TD_C + BG_E + '">' + estadoReg('eq_arena','0/4') + '</td>';
    html += '<td style="' + TD_C + BG_E + '">' + estadoReg('eq_arena','ZA25') + '</td>';
    html += '<td style="' + TD_C + BG_L + '">' + estadoReg('ind_lajas','4/12') + '</td>';
    html += '<td style="' + TD_C + BG_L + '">' + estadoReg('ind_lajas','12/20') + '</td>';
    html += '<td style="' + TD_C + BG_L + '">' + estadoReg('ind_lajas','20/40') + '</td>';
    html += '<td style="' + TD_C + BG_L + '">' + estadoReg('ind_lajas','ZA25') + '</td>';
    html += '<td style="' + TD_C + BG_Ca + '">' + estadoReg('caras_fractura','4/12') + '</td>';
    html += '<td style="' + TD_C + BG_Ca + '">' + estadoReg('caras_fractura','12/20') + '</td>';
    html += '<td style="' + TD_C + BG_Ca + '">' + estadoReg('caras_fractura','20/40') + '</td>';
    html += '</tr>';
  });

  if (!semanas.length) html += '<tr><td colspan="27" style="padding:24px;text-align:center;color:var(--muted)">Sin semanas para ' + _ensayosAnio + '</td></tr>';
  html += '</tbody></table></div>';
  html += '<div style="display:flex;gap:16px;margin-top:10px;font-size:.72rem;color:var(--muted);flex-wrap:wrap">';
  html += '<span><span style="color:#2e7d32;font-weight:700">\u2713\u2713</span> Conforme</span>';
  html += '<span><span style="color:#c62828;font-weight:700">\u2717</span> No conforme</span>';
  html += '<span><span style="color:#e65100;font-weight:700">\u2713</span> Recogido (sin resultado)</span>';
  html += '<span style="color:#ccc">+ Sin ensayo</span>';
  html += '</div>';
  return html;
}

// TAB REGISTROS
function _ensayosRenderRegistros() {
  const fracciones = ['0/4','4/12','12/20','20/40'];
  let html = '<div style="display:flex;gap:8px;margin-bottom:16px">';
  fracciones.forEach(function(f) {
    const active = _ensayosFraccion === f;
    html += '<button onclick="ensayosFraccionTab(\'' + f + '\')" style="padding:5px 14px;border-radius:20px;border:1px solid var(--border);background:' + (active?'var(--accent)':'var(--surface)') + ';color:' + (active?'#fff':'var(--text)') + ';cursor:pointer;font-size:.82rem">' + f + '</button>';
  });
  html += '</div>';

  const regs = _ensayosRegistros.filter(function(r){ return r.fraccion === _ensayosFraccion; })
    .sort(function(a,b){ return (b.fecha_toma||'').localeCompare(a.fecha_toma||''); });

  html += '<div style="overflow-x:auto"><table style="border-collapse:collapse;font-size:.78rem;width:100%">';
  html += '<thead><tr style="background:#1e2a1e;color:#fff">';
  ['ESTADO','FECHA TOMA','FECHA ENSAYO','N\u00ba MUESTRA','N\u00ba ALBAR\u00c1N','DECL. PREST.','GRANULOMETR\u00cdA \u2014 % QUE PASA (UNE-EN 933-1)','EQ. ARENA','CONT. FINOS','COMENTARIO',''].forEach(function(h){
    html += '<th style="padding:7px 10px;white-space:nowrap">' + h + '</th>';
  });
  html += '</tr></thead><tbody>';

  regs.forEach(function(r) {
    const res = r.resultados || {};
    const estadoLabel = r.estado === 'conforme' ? 'Conforme' : r.estado === 'no_conforme' ? 'No conforme' : 'Pendiente';
    const estadoColor = r.estado === 'conforme' ? '#4caf50' : r.estado === 'no_conforme' ? '#f44336' : '#ff9800';
    const granStr = ['8','6.3','4','2','1','0.5','0.25','0.125','0.063'].map(function(t){ return res['gran_'+t] != null ? res['gran_'+t] : '\u2014'; }).join(' | ');
    html += '<tr style="border-bottom:1px solid var(--border)">';
    html += '<td style="padding:6px 10px"><span style="background:' + estadoColor + ';color:#fff;padding:2px 8px;border-radius:10px;font-size:.72rem">' + estadoLabel + '</span></td>';
    html += '<td style="padding:6px 10px;white-space:nowrap">' + (r.fecha_toma||'\u2014') + '</td>';
    html += '<td style="padding:6px 10px;white-space:nowrap">' + (r.fecha_acta||'\u2014') + '</td>';
    html += '<td style="padding:6px 10px">' + (r.num_muestra||'\u2014') + '</td>';
    html += '<td style="padding:6px 10px">' + (r.num_albaran||'\u2014') + '</td>';
    html += '<td style="padding:6px 10px">' + (r.num_acta||'\u2014') + '</td>';
    html += '<td style="padding:6px 10px;font-size:.72rem">' + granStr + '</td>';
    html += '<td style="padding:6px 10px">' + (res.eq_arena!=null?res.eq_arena:'\u2014') + '</td>';
    html += '<td style="padding:6px 10px">' + (res.cont_finos!=null?res.cont_finos:'\u2014') + '</td>';
    html += '<td style="padding:6px 10px;color:var(--muted)">' + (r.comentario||'') + '</td>';
    html += '<td style="padding:6px 10px;text-align:center"><button onclick="ensayosEliminarRegistro(\'' + r.id + '\')" style="background:none;border:none;color:#c62828;cursor:pointer;font-size:1rem;padding:2px 6px" title="Eliminar">\u2715</button></td>';
    html += '</tr>';
  });

  if (!regs.length) html += '<tr><td colspan="10" style="padding:20px;text-align:center;color:var(--muted)">Sin registros para ' + _ensayosFraccion + '</td></tr>';
  html += '</tbody></table></div>';
  return html;
}

// TAB ANUALES
function _ensayosRenderAnuales() {
  const BIANUALES = ['Contenido de Azufre','Sulfatos Solubles en \u00c1cido','CO. Ligeros','CO. H\u00famicos','CD. \u00c1cido F\u00falvico','CO. Mortero','Cloruros Solubles en Agua'];
  const ANUALES = [
    {norma:'UNE-EN 1097-6 \u00b7 ANUAL', label:'Absorci\u00f3n de Agua', fracciones:['0/4','4/12','12/20','20/40']},
    {norma:'UNE-EN 1097-6 \u00b7 ANUAL', label:'Densidad de Part\u00edculas', fracciones:['0/4','4/12','12/20','20/40']},
    {norma:'UNE-EN 1097-2 \u00b7 ANUAL', label:'Los \u00c1ngeles', fracciones:['4/12','12/20','20/40']},
    {norma:'UNE-EN 1097B \u00b7 ANUAL', label:'Resistencia al Pulimento Acel.', fracciones:['4/12']},
    {norma:'UNE-EN 1367-2 \u00b7 ANUAL', label:'Sulfatos de Magnesio', fracciones:['12/20']},
    {norma:'UNE 146512 \u00b7 ANUAL', label:'\u00c1lcali-S\u00edlice', fracciones:['0/4']},
  ];

  let html = '<div style="font-size:.8rem;font-weight:700;color:var(--muted);margin-bottom:8px;letter-spacing:.05em">BIANUALES \u2014 UNE-EN 1744-1</div>';
  html += '<div style="overflow-x:auto;margin-bottom:24px"><table style="border-collapse:collapse;font-size:.78rem"><thead><tr style="background:#1e2a1e;color:#fff"><th style="padding:7px 10px">MAT.</th>';
  BIANUALES.forEach(function(b){ html += '<th style="padding:7px 10px;white-space:nowrap">' + b + '<br><span style="font-size:.7rem;opacity:.6">0/4<br>FECHA</span></th>'; });
  html += '</tr></thead><tbody><tr style="border-bottom:1px solid var(--border)"><td style="padding:6px 10px;font-weight:600">AC</td>';
  BIANUALES.forEach(function(){ html += '<td style="padding:6px 10px;text-align:center;cursor:pointer;color:var(--muted)">+</td>'; });
  html += '</tr></tbody></table></div>';

  html += '<div style="font-size:.8rem;font-weight:700;color:var(--muted);margin-bottom:8px;letter-spacing:.05em">ANUALES</div>';
  html += '<div style="display:flex;flex-wrap:wrap;gap:12px">';
  ANUALES.forEach(function(a) {
    html += '<div style="background:#1e2a1e;border-radius:8px;padding:12px;min-width:200px">';
    html += '<div style="font-size:.68rem;color:var(--accent);margin-bottom:4px">' + a.norma + '</div>';
    html += '<div style="font-weight:700;color:#fff;margin-bottom:10px">' + a.label + '</div>';
    html += '<table style="border-collapse:collapse;font-size:.75rem;width:100%"><thead><tr><th style="padding:4px 8px;color:#aaa">MAT.</th>';
    a.fracciones.forEach(function(f){ html += '<th style="padding:4px 8px;color:#aaa">' + f + '<br><span style="font-size:.65rem;opacity:.6">FECHA</span></th>'; });
    html += '</tr></thead><tbody><tr><td style="padding:4px 8px;font-weight:600;color:#fff">AC</td>';
    a.fracciones.forEach(function(){ html += '<td style="padding:4px 8px;text-align:center;cursor:pointer;color:var(--muted)">+</td>'; });
    html += '</tr></tbody></table></div>';
  });
  html += '</div>';
  return html;
}

// TAB PRESTACIONES
function _ensayosRenderPrestaciones() {
  const PARAMS = [
    {key:'referencia',label:'REFERENCIA DoP'},
    {key:'fecha_emision',label:'FECHA EMISI\u00d3N'},
    {key:'gran_categoria',label:'GRAN. (CAT.)',norma:'EN 933-1'},
    {key:'cont_finos',label:'CONT. FINOS',norma:'EN 933-1'},
    {key:'eq_arena',label:'EQ. ARENA',norma:'EN 933-8'},
    {key:'ind_lajas',label:'\u00cdND. LAJAS',norma:'EN 933-3'},
    {key:'caras_fractura',label:'CARAS FRAC.',norma:'EN 933-5'},
    {key:'los_angeles',label:'LOS \u00c1NGELES',norma:'EN 1097-2'},
    {key:'azul_metileno',label:'AZUL MET.',norma:'EN 933-9'},
    {key:'densidad_m',label:'DENS. m [t/m\u00b3]',norma:'EN 1097-6'},
    {key:'densidad_max',label:'DENS. MAX [t/m\u00b3]',norma:'EN 1097-6'},
    {key:'densidad_min',label:'DENS. MIN [t/m\u00b3]',norma:'EN 1097-6'},
    {key:'cpa',label:'CPA',norma:'EN 1097-8'},
    {key:'sulfato_mg',label:'SULF. Mg',norma:'EN 1367-2'},
    {key:'absorcion',label:'ABSORCI\u00d3N [%]',norma:'EN 1097-6'},
    {key:'cloruros',label:'CLORUROS [%]',norma:'EN 1744-1'},
    {key:'azufre',label:'AZUFRE [%]',norma:'EN 1744-1'},
    {key:'sulfatos_acido',label:'SULF. \u00c1CIDO [%]',norma:'EN 1744-1'},
    {key:'cont_ligeros',label:'CONT. LIG.',norma:'EN 1744-1'},
    {key:'cont_humicos',label:'CONT. H\u00daM.',norma:'EN 1744-1'},
  ];
  const fracciones = ['0/4','4/12','12/20','20/40'];
  const byFrac = {};
  fracciones.forEach(function(f){ byFrac[f] = _ensayosPrestaciones.find(function(p){ return p.fraccion === f; }) || {}; });

  let html = '<p style="font-size:.78rem;color:var(--muted);margin-bottom:12px">Haz clic en una celda para editar. Los cambios se guardan autom\u00e1ticamente.</p>';
  html += '<div style="overflow-x:auto"><table style="border-collapse:collapse;font-size:.78rem">';
  html += '<thead><tr style="background:#1e2a1e;color:#fff"><th style="padding:7px 12px">PAR\u00c1METRO</th><th style="padding:7px 12px">NORMA</th>';
  fracciones.forEach(function(f){ html += '<th style="padding:7px 30px">' + f + '</th>'; });
  html += '</tr></thead><tbody>';

  PARAMS.forEach(function(p) {
    html += '<tr style="border-bottom:1px solid var(--border)"><td style="padding:6px 12px;font-weight:600;font-size:.75rem">' + p.label + '</td><td style="padding:6px 12px;color:var(--muted);font-size:.72rem">' + (p.norma||'') + '</td>';
    fracciones.forEach(function(f) {
      const prest = byFrac[f];
      const val = prest[p.key] != null ? prest[p.key] : '';
      const id = prest.id || '';
      html += '<td style="padding:6px 12px;text-align:center;cursor:pointer" onclick="ensayosEditPrestacion(this,\'' + id + '\',\'' + p.key + '\')" data-val="' + val + '">' + (val||'\u2014') + '</td>';
    });
    html += '</tr>';
  });
  html += '</tbody></table></div>';
  return html;
}

// ACCIONES
function ensayosFraccionTab(f) {
  _ensayosFraccion = f;
  document.getElementById('ensayos-body').innerHTML = _ensayosRenderRegistros();
}

function ensayosAbrirSemana(id) {
  const sem = _ensayosSemanas.find(function(s){ return s.id === id; });
  if (!sem) return;
  _ensayosAbrirModalSemana(sem);
}

function ensayosAbrirRegistro(id) {
  const reg = _ensayosRegistros.find(function(r){ return r.id === id; });
  if (!reg) return;
  alert('Registro: ' + (reg.num_muestra||'') + ' \u2014 ' + (reg.estado||''));
}

function ensayosNuevaSemana() {
  _ensayosAbrirModalSemana(null);
}

function _ensayosAbrirModalSemana(sem) {
  const modal = document.getElementById('ensayos-modal-semana');
  if (!modal) return;
  document.getElementById('ems-id').value = sem ? sem.id : '';
  document.getElementById('ems-fecha').value = sem ? sem.fecha_lunes : _ensayosLunesActual();
  document.getElementById('ems-mat').value = sem ? (sem.tipo_material||'NP') : 'NP';
  document.getElementById('ems-tn04').value = sem ? (sem.tn_04||'') : '';
  document.getElementById('ems-tn412').value = sem ? (sem.tn_412||'') : '';
  document.getElementById('ems-tn1220').value = sem ? (sem.tn_1220||'') : '';
  document.getElementById('ems-tn2040').value = sem ? (sem.tn_2040||'') : '';
  document.getElementById('ems-comentario').value = sem ? (sem.comentario||'') : '';
  modal.style.display = 'flex';
}

function ensayosCerrarModalSemana() {
  const modal = document.getElementById('ensayos-modal-semana');
  if (modal) modal.style.display = 'none';
}

async function ensayosGuardarSemana() {
  const id = document.getElementById('ems-id').value;
  const data = {
    fecha_lunes: document.getElementById('ems-fecha').value,
    tipo_material: document.getElementById('ems-mat').value,
    tn_04: parseFloat(document.getElementById('ems-tn04').value)||null,
    tn_412: parseFloat(document.getElementById('ems-tn412').value)||null,
    tn_1220: parseFloat(document.getElementById('ems-tn1220').value)||null,
    tn_2040: parseFloat(document.getElementById('ems-tn2040').value)||null,
    comentario: document.getElementById('ems-comentario').value||null,
  };
  let res;
  if (id) res = await updateEnsayoSemana(id, data);
  else res = await insertEnsayoSemana(data);
  if (!res.ok) { alert('Error: ' + res.error); return; }
  ensayosCerrarModalSemana();
  const rSem = await getEnsayosSemanas(_ensayosAnio);
  _ensayosSemanas = rSem.ok ? rSem.data : [];
  _ensayosRenderTab('control');
}

async function ensayosCambiarAnio(v) {
  _ensayosAnio = parseInt(v);
  await _ensayosCargarTodo();
  _ensayosRenderTab(_ensayosTab);
}

function ensayosEditPrestacion(td, id, key) {
  if (!id) return;
  const cur = td.dataset.val || '';
  const input = document.createElement('input');
  input.value = cur;
  input.style.cssText = 'width:80px;border:1px solid var(--accent);border-radius:4px;padding:2px 4px;font-size:.78rem;background:var(--surface);color:var(--text)';
  td.innerHTML = '';
  td.appendChild(input);
  input.focus();
  input.select();
  const save = async function() {
    const val = input.value.trim();
    td.dataset.val = val;
    td.textContent = val || '\u2014';
    const upd = {};
    upd[key] = val || null;
    await updateEnsayoPrestacion(id, upd);
    const prest = _ensayosPrestaciones.find(function(p){ return p.id === id; });
    if (prest) prest[key] = val || null;
  };
  input.addEventListener('blur', save);
  input.addEventListener('keydown', function(e){ if(e.key==='Enter') input.blur(); if(e.key==='Escape'){td.textContent=cur||'\u2014';} });
}

function _ensayosLunesActual() {
  const d = new Date();
  const day = d.getDay();
  const diff = (day === 0 ? -6 : 1 - day);
  d.setDate(d.getDate() + diff);
  return d.toISOString().slice(0,10);
}

async function ensayosSubirPDFs(files) {
  if (!files || !files.length) return;
  const toast = document.getElementById('ensayos-toast');
  const resultados = [];

  await _ensurePdfjs();

  for (const file of Array.from(files)) {
    if (toast) { toast.textContent = 'Procesando ' + file.name + '...'; toast.style.display = 'block'; }

    try {
      // Extraer texto del PDF en el cliente con PDF.js
      const arrayBuffer = await file.arrayBuffer();
      const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
      let texto = '';
      for (let i = 1; i <= pdf.numPages; i++) {
        const page = await pdf.getPage(i);
        const content = await page.getTextContent();
        texto += content.items.map(function(it){ return it.str; }).join(' ') + '\n';
      }

      // Parsear y abrir modal de confirmación
      const d = _ensayosParseActa(texto);
      if (toast) toast.style.display = 'none';
      ensayosCerrarDrop();
      const pdfUrl = URL.createObjectURL(file);
      await ensayosAbrirConfirm(file.name, d, pdfUrl);
      URL.revokeObjectURL(pdfUrl);

    } catch(e) {
      alert('Error procesando ' + file.name + ': ' + e.message);
    }
  }

  document.getElementById('ensayos-pdf-input').value = '';
}

function _ensayosParseActa(text) {
  const r = {};

  // Nº ACTA — buscar el número que sigue a la cabecera "Nº ACTA" o "ACTA"
  // En tablas PDF.js extrae texto en orden: cabecera fila1 fila2...
  // Formato tabla: "Nº ACTA ALBARAN Nº Nº SERIE Nº DE OBRA MUESTRA FECHA DE ACTA 2026/258  159 6716 2026/101 09/01/2026"
  const mActa = text.match(/N[ºo°]\s*ACTA[\s\S]{0,80}?(20\d{2}\/\d+)/i);
  if (mActa) r.num_acta = mActa[1];

  // Nº ALBARÁN / MUESTRA — puede venir como ".2026/101" (con punto delante)
  const mAlb = text.match(/MUESTRA[\s\S]{0,60}?\.?(20\d{2}\/\d+)/i);
  if (mAlb && mAlb[1] !== r.num_acta) r.num_albaran = mAlb[1];
  if (!r.num_albaran || r.num_albaran === r.num_acta) {
    const todos = [...text.matchAll(/\.?(20\d{2}\/\d+)/g)].map(m=>m[1]);
    const distinto = todos.find(function(n){ return n !== r.num_acta; });
    if (distinto) r.num_albaran = distinto;
  }

  // Fecha de toma
  const mToma = text.match(/Fecha de toma[:\s]+(\d{2}\/\d{2}\/\d{4})/i);
  if (mToma) r.fecha_toma = _ensayosIsoFecha(mToma[1]);

  // Fecha acta / fin ensayos
  // El PDF de ESOCAN tiene la fecha ANTES de "FECHA DE ACTA" en el flujo de texto
  const mFin = text.match(/Fin de ensayos[:\s]+(\d{2}\/\d{2}\/\d{4})/i)
    || text.match(/(\d{2}\/\d{2}\/\d{4})\s*(?=[\s\S]{0,30}FECHA DE ACTA)/i)
    || text.match(/FECHA DE ACTA[\s\S]{0,30}?(\d{2}\/\d{2}\/\d{4})/i);
  if (mFin) r.fecha_acta = _ensayosIsoFecha(mFin[1]);

  // Fracción
  const mFrac = text.match(/[ÁA]rido\s+([\d\/]+)/i) || text.match(/Tipo de material[:\s]+[ÁA]rido\s+([\d\/]+)/i);
  if (mFrac) r.fraccion = mFrac[1].trim();

  // Tipo ensayo
  if (/granulometr/i.test(text)) r.tipo_ensayo = 'granulometria';
  else if (/equivalente.*arena/i.test(text)) r.tipo_ensayo = 'eq_arena';
  else if (/contenido.*fino|finos.*tamiz/i.test(text)) r.tipo_ensayo = 'cont_finos';
  else if (/[íi]ndice.*laja/i.test(text)) r.tipo_ensayo = 'ind_lajas';
  else if (/caras.*fractura/i.test(text)) r.tipo_ensayo = 'caras_fractura';

  // Granulometría — tabla tamiz/pasa
  // PDF.js extrae la tabla como: "Tamiz (mm)   Pasa (%) 14   100 12,5   99 10   91 8   79 6,3   58 4   8 2   1 1   0"
  if (r.tipo_ensayo === 'granulometria') {
    const res = {};
    // Extraer todos los pares (tamiz, pasa) del bloque de tabla
    // Buscar desde "Tamiz" o "Pasa" hasta el final de la zona de datos
    var bloqueM = text.match(/Tamiz[\s\S]{0,20}?Pasa[\s\S]+?((?:\d[\d,\.]*\s+\d{1,3}\s*){3,})/i);
    var bloque = bloqueM ? bloqueM[1] : text;
    // Extraer pares: número_tamiz  número_pasa separados por espacios
    var pares = [...bloque.matchAll(/\b(\d+[,.]?\d*)\s{1,10}(\d{1,3})\b/g)];
    // Mapa tamiz normalizado -> clave gran_
    var tamMap = {'8':'8','6.3':'6.3','6,3':'6.3','4':'4','2':'2','1':'1',
                  '0.5':'0.5','0,5':'0.5','0.25':'0.25','0,25':'0.25',
                  '0.125':'0.125','0,125':'0.125','0.063':'0.063','0,063':'0.063',
                  }; // 14, 12.5, 10 se ignoran — no están en nuestros campos
    pares.forEach(function(p) {
      var tamiz = p[1].trim();
      var pasa = parseInt(p[2]);
      var key = tamMap[tamiz];
      if (key && pasa >= 0 && pasa <= 100) res['gran_' + key] = pasa;
    });
    r.resultados = res;
  }

  // Eq. Arena
  if (r.tipo_ensayo === 'eq_arena') {
    const m = text.match(/(\d{1,3}(?:[,.]\d+)?)\s*%/);
    if (m) r.resultados = { eq_arena: parseFloat(m[1].replace(',','.')) };
  }

  // Contenido en finos — "Contenido en finos que pasan por el tamiz 0.063   11,29"
  if (r.tipo_ensayo === 'cont_finos') {
    const m = text.match(/Contenido en finos[\s\S]{0,60}?([\d]+[,.]\d+)\s*(?:\n|$| {2})/i)
      || text.match(/tamiz 0[.,]063[\s\S]{0,20}?([\d]+[,.]\d+)/i);
    if (m) r.resultados = { cont_finos: parseFloat(m[1].replace(',','.')) };
  }

  r.estado = 'recogido';
  return r;
}

function _ensayosIsoFecha(ddmmyyyy) {
  const p = ddmmyyyy.split('/');
  return p.length === 3 ? p[2]+'-'+p[1]+'-'+p[0] : ddmmyyyy;
}

function ensayosAbrirDrop() {
  const m = document.getElementById('ensayos-drop-modal');
  if (m) m.style.display = 'flex';
}
function ensayosCerrarDrop() {
  const m = document.getElementById('ensayos-drop-modal');
  if (m) m.style.display = 'none';
}
function ensayosDropPDFs(files) {
  ensayosSubirPDFs(files);
}

// ── Modal confirmación acta ───────────────────────────────────
let _ensayosConfirmResolve = null;

async function ensayosAbrirConfirm(filename, d, pdfUrl) {
  return new Promise(function(resolve) {
    _ensayosConfirmResolve = resolve;

    // Visor PDF — renderizar con PDF.js en canvas (evita CSP frame-src)
    const pdfCol = document.getElementById('ecf-pdf-col');
    const pdfWrap = document.getElementById('ecf-pdf-canvas-wrap');
    if (pdfCol && pdfWrap) {
      if (pdfUrl) {
        pdfCol.style.display = 'flex';
        pdfWrap.innerHTML = '<div style="color:#aaa;font-size:.8rem;padding:12px">Cargando PDF...</div>';
        pdfjsLib.getDocument(pdfUrl).promise.then(function(pdf) {
          pdfWrap.innerHTML = '';
          var pageNums = [];
          for (var i = 1; i <= pdf.numPages; i++) pageNums.push(i);
          pageNums.reduce(function(p, pageNum) {
            return p.then(function() {
              return pdf.getPage(pageNum).then(function(page) {
                var vp = page.getViewport({ scale: 1.2 });
                var canvas = document.createElement('canvas');
                canvas.width = vp.width;
                canvas.height = vp.height;
                canvas.style.cssText = 'display:block;width:100%;margin-bottom:4px;border-radius:4px';
                pdfWrap.appendChild(canvas);
                return page.render({ canvasContext: canvas.getContext('2d'), viewport: vp }).promise;
              });
            });
          }, Promise.resolve());
        }).catch(function() {
          pdfWrap.innerHTML = '<div style="color:#f88;font-size:.8rem;padding:12px">No se pudo cargar el PDF</div>';
        });
      } else {
        pdfCol.style.display = 'none';
        pdfWrap.innerHTML = '';
      }
    }

    // Rellenar campos
    document.getElementById('ensayos-confirm-filename').textContent = filename;
    document.getElementById('ecf-num-acta').value = d.num_acta || '';
    document.getElementById('ecf-num-albaran').value = d.num_albaran || '';
    document.getElementById('ecf-fecha-toma').value = d.fecha_toma || '';
    document.getElementById('ecf-fecha-acta').value = d.fecha_acta || '';
    const fracSel = document.getElementById('ecf-fraccion');
    if (d.fraccion) fracSel.value = d.fraccion;
    const tipoSel = document.getElementById('ecf-tipo');
    if (d.tipo_ensayo) tipoSel.value = d.tipo_ensayo;

    // Semanas
    const semSel = document.getElementById('ecf-semana');
    semSel.innerHTML = '<option value="">— Sin vincular —</option>';
    _ensayosSemanas.forEach(function(s) {
      const opt = document.createElement('option');
      opt.value = s.id;
      opt.textContent = s.fecha_lunes + ' (' + (s.tipo_material||'NP') + ')';
      // Auto-seleccionar semana por fecha_toma
      if (d.fecha_toma) {
        const lunes = new Date(s.fecha_lunes);
        const toma = new Date(d.fecha_toma);
        const domingo = new Date(lunes); domingo.setDate(domingo.getDate() + 6);
        if (toma >= lunes && toma <= domingo) opt.selected = true;
      }
      semSel.appendChild(opt);
    });

    // Resultados según tipo
    const wrap = document.getElementById('ecf-resultados-wrap');
    wrap.innerHTML = '';
    if (d.tipo_ensayo === 'granulometria' && d.resultados) {
      wrap.innerHTML = '<div style="font-size:.8rem;color:var(--muted);margin-bottom:6px">Granulometría — % que pasa</div>'
        + '<div style="display:grid;grid-template-columns:repeat(5,1fr);gap:6px">'
        + ['8','6.3','4','2','1','0.5','0.25','0.125','0.063'].map(function(t) {
            const v = d.resultados['gran_'+t];
            return '<label style="font-size:.75rem;text-align:center">'
              + '<span style="display:block;color:var(--muted);margin-bottom:2px">' + t + '</span>'
              + '<input type="number" id="ecf-gran-' + t.replace('.','_') + '" value="' + (v!=null?v:'') + '" min="0" max="100"'
              + ' style="width:100%;padding:5px 4px;border:1px solid var(--border);border-radius:6px;background:var(--surface);color:var(--text);font-size:.82rem;text-align:center">'
              + '</label>';
          }).join('')
        + '</div>';
    } else if (d.tipo_ensayo === 'cont_finos') {
      const v = d.resultados ? d.resultados.cont_finos : '';
      wrap.innerHTML = '<label style="font-size:.8rem;color:var(--muted)">Contenido en finos (%)'
        + '<input type="number" id="ecf-cont-finos" value="' + (v!=null&&v!==''?v:'') + '" step="0.01" min="0" max="100"'
        + ' style="width:100%;margin-top:3px;padding:7px 10px;border:1px solid var(--border);border-radius:7px;background:var(--surface);color:var(--text);font-size:.85rem">'
        + '</label>';
    } else if (d.tipo_ensayo === 'eq_arena') {
      const v = d.resultados ? d.resultados.eq_arena : '';
      wrap.innerHTML = '<label style="font-size:.8rem;color:var(--muted)">Equivalente de arena (%)'
        + '<input type="number" id="ecf-eq-arena" value="' + (v!=null&&v!==''?v:'') + '" step="0.1" min="0" max="100"'
        + ' style="width:100%;margin-top:3px;padding:7px 10px;border:1px solid var(--border);border-radius:7px;background:var(--surface);color:var(--text);font-size:.85rem">'
        + '</label>';
    }

    document.getElementById('ensayos-confirm-modal').style.display = 'flex';
  });
}

function ensayosCerrarConfirm() {
  document.getElementById('ensayos-confirm-modal').style.display = 'none';
  if (_ensayosConfirmResolve) { _ensayosConfirmResolve(null); _ensayosConfirmResolve = null; }
}

async function ensayosConfirmarGuardar() {
  const semanaId = document.getElementById('ecf-semana').value;
  if (!semanaId) { alert('Selecciona una semana'); return; }

  const tipo = document.getElementById('ecf-tipo').value;
  const fraccion = document.getElementById('ecf-fraccion').value;

  // Recoger resultados
  let resultados = {};
  if (tipo === 'granulometria') {
    ['8','6.3','4','2','1','0.5','0.25','0.125','0.063'].forEach(function(t) {
      const el = document.getElementById('ecf-gran-' + t.replace('.','_'));
      if (el && el.value !== '') resultados['gran_'+t] = parseInt(el.value);
    });
  } else if (tipo === 'cont_finos') {
    const el = document.getElementById('ecf-cont-finos');
    if (el && el.value !== '') resultados.cont_finos = parseFloat(el.value);
  } else if (tipo === 'eq_arena') {
    const el = document.getElementById('ecf-eq-arena');
    if (el && el.value !== '') resultados.eq_arena = parseFloat(el.value);
  }

  const payload = {
    semana_id: semanaId,
    tipo_ensayo: tipo,
    fraccion: fraccion,
    fecha_toma: document.getElementById('ecf-fecha-toma').value || null,
    fecha_acta: document.getElementById('ecf-fecha-acta').value || null,
    num_acta: document.getElementById('ecf-num-acta').value || null,
    num_albaran: document.getElementById('ecf-num-albaran').value || null,
    resultados: resultados,
    estado: 'recogido'
  };

  const existing = _ensayosRegistros.find(function(r) {
    return r.semana_id === semanaId && r.tipo_ensayo === tipo && r.fraccion === fraccion;
  });

  let res;
  if (existing) res = await updateEnsayoRegistro(existing.id, payload);
  else res = await insertEnsayoRegistro(payload);

  if (!res.ok) { alert('Error al guardar: ' + res.error); return; }

  document.getElementById('ensayos-confirm-modal').style.display = 'none';
  if (_ensayosConfirmResolve) { _ensayosConfirmResolve(true); _ensayosConfirmResolve = null; }

  // Recargar y toast
  const rReg = await getEnsayosRegistros(_ensayosAnio);
  _ensayosRegistros = rReg.ok ? rReg.data : [];
  _ensayosRenderTab(_ensayosTab);

  const toast = document.getElementById('ensayos-toast');
  if (toast) {
    toast.style.background = '#2e7d32';
    toast.textContent = '\u2713 Acta guardada correctamente';
    toast.style.display = 'block';
    setTimeout(function(){ toast.style.display='none'; toast.style.background='#333'; }, 3000);
  }
}

// ── Eliminar registro ─────────────────────────────────────────
async function ensayosEliminarRegistro(id) {
  if (!confirm('¿Eliminar este registro?')) return;
  const res = await deleteEnsayoRegistro(id);
  if (!res.ok) { alert('Error: ' + res.error); return; }
  const rReg = await getEnsayosRegistros(_ensayosAnio);
  _ensayosRegistros = rReg.ok ? rReg.data : [];
  _ensayosRenderTab(_ensayosTab);
}
