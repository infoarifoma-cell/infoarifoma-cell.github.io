// POST /api/bc/costes
// Proxy: obtiene movimientos contables (G/L Entries) de cuentas 600-799
// con dimensiones CA y PROYECTO expandidas, filtrado por año

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ ok: false, error: 'Método no permitido' });
  }

  const { token, anyo } = req.body;
  if (!token) return res.status(401).json({ ok: false, error: 'Token requerido' });
  if (!anyo) return res.status(400).json({ ok: false, error: 'Año requerido' });

  const BC_TENANT = process.env.BC_TENANT;
  const BC_ENV = process.env.BC_ENV;
  const BC_COMPANY = process.env.BC_COMPANY;
  const base = `https://api.businesscentral.dynamics.com/v2.0/${BC_TENANT}/${BC_ENV}/api/v2.0/companies`;
  const headers = { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' };

  try {
    // Obtener company ID
    const cRes = await fetch(base, { headers });
    if (!cRes.ok) throw new Error('No se pudo obtener company: ' + cRes.statusText);
    const cJson = await cRes.json();
    const company = cJson.value.find(c => c.name === BC_COMPANY);
    if (!company) throw new Error('Company no encontrada: ' + BC_COMPANY);
    const cid = company.id;

    // Paginar G/L Entries del año con dimensiones expandidas
    // Filtro: cuentas 6* y 7* (gastos e ingresos) del año seleccionado
    const fechaInicio = `${anyo}-01-01`;
    const fechaFin = `${anyo}-12-31`;
    const filter = `postingDate ge ${fechaInicio} and postingDate le ${fechaFin} and ((accountNumber ge '6000000000' and accountNumber lt '8000000000') or (accountNumber ge '600' and accountNumber lt '800'))`;

    let allEntries = [];
    let url = `${base}(${cid})/generalLedgerEntries?$filter=${encodeURIComponent(filter)}&$expand=dimensionSetLines&$orderby=postingDate asc&$top=1000`;

    // Paginación OData
    while (url) {
      const glRes = await fetch(url, { headers });
      if (!glRes.ok) throw new Error('Error G/L Entries: ' + glRes.statusText);
      const glJson = await glRes.json();

      // Procesar cada entrada: extraer solo lo necesario
      for (const entry of (glJson.value || [])) {
        const dims = {};
        for (const dim of (entry.dimensionSetLines || [])) {
          dims[dim.code] = { valueCode: dim.valueCode, displayName: dim.valueDisplayName };
        }

        allEntries.push({
          date: entry.postingDate,
          account: entry.accountNumber,
          description: entry.description,
          debit: entry.debitAmount,
          credit: entry.creditAmount,
          ca: dims.CA?.valueCode || null,
          caName: dims.CA?.displayName || null,
          proyecto: dims.PROYECTO?.valueCode || null,
          proyectoName: dims.PROYECTO?.displayName || null,
          docNumber: entry.documentNumber
        });
      }

      url = glJson['@odata.nextLink'] || null;
    }

    return res.status(200).json({ ok: true, count: allEntries.length, entries: allEntries });
  } catch (error) {
    console.error('BC costes error:', error.message);
    return res.status(500).json({ ok: false, error: error.message });
  }
}
