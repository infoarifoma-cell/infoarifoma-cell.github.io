// POST /api/bc/facturas-pendientes-compra
// Obtiene purchaseInvoices con status=open desde BC API v2.0
// Campos: number, invoiceDate, vendorInvoiceNumber, vendorName, totalAmountExcludingTax, totalAmountIncludingTax, remainingAmount, dueDate

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

    // Obtener facturas de compra pendientes (status=open)
    const filter = encodeURIComponent("status eq 'open'");
    const select = '$select=number,invoiceDate,vendorInvoiceNumber,vendorName,totalAmountExcludingTax,totalAmountIncludingTax,dueDate';
    const url = `${base}(${cid})/purchaseInvoices?$filter=${filter}&${select}&$top=500`;

    const invRes = await fetch(url, { headers });
    if (!invRes.ok) {
      const errText = await invRes.text().catch(() => invRes.statusText);
      throw new Error('Error purchaseInvoices: ' + errText);
    }
    const invJson = await invRes.json();

    const data = (invJson.value || []).map(inv => ({
      number: inv.number,
      invoiceDate: inv.invoiceDate,
      vendorInvoiceNumber: inv.vendorInvoiceNumber,
      vendorName: inv.vendorName,
      totalAmountExcludingTax: inv.totalAmountExcludingTax,
      totalAmountIncludingTax: inv.totalAmountIncludingTax,
      dueDate: inv.dueDate
    }));

    return res.status(200).json({ ok: true, data });
  } catch (error) {
    console.error('BC facturas-pendientes-compra error:', error.message);
    return res.status(500).json({ ok: false, error: error.message });
  }
}
