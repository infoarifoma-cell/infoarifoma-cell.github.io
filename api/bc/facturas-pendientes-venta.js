// POST /api/bc/facturas-pendientes-venta
// Obtiene salesInvoices con status=open desde BC API v2.0
// Campos: number, invoiceDate, customerName, dueDate, totalAmountIncludingTax, remainingAmount

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ ok: false, error: 'Método no permitido' });
  }

  const { token } = req.body;
  if (!token) return res.status(401).json({ ok: false, error: 'Token requerido' });

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

    // Obtener facturas de venta pendientes (status=open)
    const filter = encodeURIComponent("status eq 'open'");
    const select = '$select=number,invoiceDate,customerName,dueDate,totalAmountIncludingTax,remainingAmount';
    const url = `${base}(${cid})/salesInvoices?$filter=${filter}&${select}&$top=500`;

    const invRes = await fetch(url, { headers });
    if (!invRes.ok) {
      const errText = await invRes.text().catch(() => invRes.statusText);
      throw new Error('Error salesInvoices: ' + errText);
    }
    const invJson = await invRes.json();

    const data = (invJson.value || []).map(inv => ({
      number: inv.number,
      invoiceDate: inv.invoiceDate,
      customerName: inv.customerName,
      dueDate: inv.dueDate,
      totalAmountIncludingTax: inv.totalAmountIncludingTax,
      remainingAmount: inv.remainingAmount
    }));

    return res.status(200).json({ ok: true, data });
  } catch (error) {
    console.error('BC facturas-pendientes-venta error:', error.message);
    return res.status(500).json({ ok: false, error: error.message });
  }
}
