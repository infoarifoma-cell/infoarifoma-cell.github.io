// POST /api/bc/facturas-pendientes-compra
// Obtiene facturas de compra pendientes desde vendorLedgerEntries (API v2.0)
// Filtra documentType=Invoice y open=true → tiene remainingAmount real

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

    // vendorLedgerEntries: documentType=Invoice, open=true
    const filter = encodeURIComponent("documentType eq 'Invoice' and open eq true");
    const select = '$select=documentNumber,postingDate,externalDocumentNumber,vendorNumber,vendorName,dueDate,amount,remainingAmount';
    const url = `${base}(${cid})/vendorLedgerEntries?$filter=${filter}&${select}&$top=500`;

    const invRes = await fetch(url, { headers });
    if (!invRes.ok) {
      const errText = await invRes.text().catch(() => invRes.statusText);
      throw new Error('Error vendorLedgerEntries: ' + errText);
    }
    const invJson = await invRes.json();

    const data = (invJson.value || []).map(e => ({
      number: e.documentNumber,
      invoiceDate: e.postingDate,
      vendorInvoiceNumber: e.externalDocumentNumber,
      vendorNumber: e.vendorNumber,
      vendorName: e.vendorName,
      dueDate: e.dueDate,
      totalAmountExcludingTax: null,
      totalAmountIncludingTax: Math.abs(e.amount || 0),
      remainingAmount: Math.abs(e.remainingAmount || 0)
    }));

    return res.status(200).json({ ok: true, data });
  } catch (error) {
    console.error('BC facturas-pendientes-compra error:', error.message);
    return res.status(500).json({ ok: false, error: error.message });
  }
}
