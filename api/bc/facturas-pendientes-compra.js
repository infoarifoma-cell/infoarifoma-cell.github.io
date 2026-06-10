// POST /api/bc/facturas-pendientes-compra
// Obtiene facturas de compra desde purchaseInvoices API v2.0 (status=open)

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
    const cRes = await fetch(base, { headers });
    if (!cRes.ok) throw new Error('No se pudo obtener company: ' + cRes.statusText);
    const cJson = await cRes.json();
    const company = cJson.value.find(c => c.name === BC_COMPANY);
    if (!company) throw new Error('Company no encontrada: ' + BC_COMPANY);
    const cid = company.id;

    // Paginar para traer todas (no solo 500)
    let all = [];
    let url = `${base}(${cid})/purchaseInvoices?$filter=status eq 'open'&$select=number,invoiceDate,vendorInvoiceNumber,vendorName,totalAmountExcludingTax,totalAmountIncludingTax,dueDate&$orderby=invoiceDate desc&$top=500`;

    while (url) {
      const r = await fetch(url, { headers });
      if (!r.ok) {
        const errText = await r.text().catch(() => r.statusText);
        throw new Error('Error purchaseInvoices: ' + errText);
      }
      const j = await r.json();
      all = all.concat(j.value || []);
      url = j['@odata.nextLink'] || null;
    }

    const data = all.map(inv => ({
      number: inv.number,
      invoiceDate: inv.invoiceDate,
      vendorInvoiceNumber: inv.vendorInvoiceNumber,
      vendorName: inv.vendorName,
      dueDate: inv.dueDate,
      totalAmountExcludingTax: inv.totalAmountExcludingTax,
      totalAmountIncludingTax: inv.totalAmountIncludingTax,
      remainingAmount: null
    }));

    return res.status(200).json({ ok: true, data });
  } catch (error) {
    console.error('BC facturas-pendientes-compra error:', error.message);
    return res.status(500).json({ ok: false, error: error.message });
  }
}
