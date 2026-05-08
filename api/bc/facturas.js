// POST /api/bc/facturas
// Crear factura en BC desde frontend

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ ok: false, error: 'Método no permitido' });
  }

  const { token, customerNumber, invoiceDate, externalDocumentNumber } = req.body;

  if (!token) {
    return res.status(401).json({ ok: false, error: 'Token requerido' });
  }

  const BC_TENANT = process.env.BC_TENANT;
  const BC_ENV = process.env.BC_ENV;
  const BC_COMPANY = process.env.BC_COMPANY;
  const base = `https://api.businesscentral.dynamics.com/v2.0/${BC_TENANT}/${BC_ENV}/api/v2.0/companies`;
  const headers = { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' };

  try {
    // Obtener company ID
    const cRes = await fetch(base, { headers });
    if (!cRes.ok) throw new Error('No se pudo obtener company');

    const cJson = await cRes.json();
    const company = cJson.value.find(c => c.name === BC_COMPANY);
    if (!company) throw new Error('Company no encontrada: ' + BC_COMPANY);

    const companyId = company.id;

    // Crear factura
    const invRes = await fetch(`${base}(${companyId})/salesInvoices`, {
      method: 'POST',
      headers,
      body: JSON.stringify({
        customerNumber,
        invoiceDate,
        externalDocumentNumber
      })
    });

    if (!invRes.ok) throw new Error('No se pudo crear factura: ' + await invRes.text());

    const invoice = await invRes.json();
    return res.status(200).json({ ok: true, invoice });
  } catch (error) {
    console.error('BC factura error:', error.message);
    return res.status(500).json({ ok: false, error: error.message });
  }
}
