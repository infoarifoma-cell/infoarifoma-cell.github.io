// POST /api/bc/costes
// Proxy: obtiene movimientos contables (G/L Entries) de cuentas 600-799
// con dimensiones CA y PROYECTO expandidas, filtrado por año
// Consulta mes a mes para evitar límite de 1000 registros de BC

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

    let allEntries = [];

    // Consultar mes a mes para no superar el límite de 1000 de BC
    for (let mes = 1; mes <= 12; mes++) {
      const mm = String(mes).padStart(2, '0');
      const lastDay = new Date(Number(anyo), mes, 0).getDate();
      const fechaInicio = `${anyo}-${mm}-01`;
      const fechaFin = `${anyo}-${mm}-${String(lastDay).padStart(2, '0')}`;
      const filter = `postingDate ge ${fechaInicio} and postingDate le ${fechaFin}`;

      let url = `${base}(${cid})/generalLedgerEntries?$filter=${encodeURIComponent(filter)}&$expand=dimensionSetLines&$top=1000`;

      while (url) {
        const glRes = await fetch(url, { headers });
        if (!glRes.ok) {
          const errText = await glRes.text().catch(() => glRes.statusText);
          throw new Error(`Error G/L Entries mes ${mes}: ` + errText);
        }
        const glJson = await glRes.json();

        // Procesar cada entrada: solo cuentas 6* y 7*
        for (const entry of (glJson.value || [])) {
          const acc = entry.accountNumber || '';
          if (!acc.startsWith('6') && !acc.startsWith('7')) continue;

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
    }

    return res.status(200).json({ ok: true, count: allEntries.length, entries: allEntries });
  } catch (error) {
    console.error('BC costes error:', error.message);
    return res.status(500).json({ ok: false, error: error.message });
  }
}
