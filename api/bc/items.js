// POST /api/bc/items
// Devuelve lista de artículos desde BC

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ ok: false, error: 'Método no permitido' });
  }

  const { token } = req.body;

  if (!token || typeof token !== 'string') {
    return res.status(401).json({ ok: false, error: 'Token requerido' });
  }

  const BC_TENANT = process.env.BC_TENANT;
  const BC_ENV = process.env.BC_ENV;
  const BC_COMPANY = process.env.BC_COMPANY;
  const base = `https://api.businesscentral.dynamics.com/v2.0/${BC_TENANT}/${BC_ENV}/api/v2.0/companies`;
  const headers = { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' };

  try {
    const cRes = await fetch(base, { headers });
    if (!cRes.ok) throw new Error('No se pudo obtener company: ' + cRes.statusText);

    const cJson = await cRes.json();
    const company = cJson.value.find(c => c.name === BC_COMPANY);
    if (!company) throw new Error('Company no encontrada: ' + BC_COMPANY);

    const companyId = company.id;

    let items = [];
    let url = `${base}(${companyId})/items?$select=number,displayName,unitPrice&$top=500`;

    while (url) {
      const iRes = await fetch(url, { headers });
      if (!iRes.ok) throw new Error('Error obteniendo artículos: ' + iRes.statusText);
      const iJson = await iRes.json();
      items = items.concat((iJson.value || []).map(i => ({ number: i.number, displayName: i.displayName, unitPrice: i.unitPrice })));
      url = iJson['@odata.nextLink'] || null;
    }

    return res.status(200).json({ ok: true, items });
  } catch (error) {
    console.error('BC items error:', error.message);
    return res.status(500).json({ ok: false, error: error.message });
  }
}
