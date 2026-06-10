// POST /api/bc/facturas-pendientes
// type='venta'  → salesInvoices status=open
// type='compra' → purchaseInvoices status=open

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ ok: false, error: 'Método no permitido' });
  }

  const { token, type } = req.body;
  if (!token) return res.status(401).json({ ok: false, error: 'Token requerido' });
  if (type !== 'venta' && type !== 'compra') return res.status(400).json({ ok: false, error: 'type debe ser venta o compra' });

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

    if (type === 'venta') {
      const filter = encodeURIComponent("status eq 'open'");
      const select = '$select=number,invoiceDate,customerName,dueDate,totalAmountIncludingTax,remainingAmount';
      const url = `${base}(${cid})/salesInvoices?$filter=${filter}&${select}&$top=500`;
      const invRes = await fetch(url, { headers });
      if (!invRes.ok) throw new Error('Error salesInvoices: ' + await invRes.text().catch(() => invRes.statusText));
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
    }

    // compra
    let all = [];
    let url = `${base}(${cid})/purchaseInvoices?$filter=status eq 'open'&$select=number,invoiceDate,vendorInvoiceNumber,vendorName,totalAmountExcludingTax,totalAmountIncludingTax,dueDate&$orderby=invoiceDate desc&$top=500`;
    while (url) {
      const r = await fetch(url, { headers });
      if (!r.ok) throw new Error('Error purchaseInvoices: ' + await r.text().catch(() => r.statusText));
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
    console.error('BC facturas-pendientes error:', error.message);
    return res.status(500).json({ ok: false, error: error.message });
  }
}
