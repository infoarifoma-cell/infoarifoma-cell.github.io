// GET /api/bc/vendors
// Devuelve lista de proveedores desde BC

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

    // Obtener todos los vendors (paginando si es necesario)
    let vendors = [];
    let url = `${base}(${companyId})/vendors?$select=number,displayName&$top=500`;

    while (url) {
      const vRes = await fetch(url, { headers });
      if (!vRes.ok) throw new Error('Error obteniendo vendors: ' + vRes.statusText);
      const vJson = await vRes.json();
      vendors = vendors.concat((vJson.value || []).map(v => ({ number: v.number, name: v.displayName })));
      url = vJson['@odata.nextLink'] || null;
    }

    return res.status(200).json({ ok: true, vendors });
  } catch (error) {
    console.error('BC vendors error:', error.message);
    return res.status(500).json({ ok: false, error: error.message });
  }
}
