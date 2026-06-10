// POST /api/bc/facturas-pendientes-compra
// Obtiene facturas de compra desde OData BC (página 138 Histórico_facturas_compra_Excel)

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ ok: false, error: 'Método no permitido' });
  }

  const { token } = req.body;
  if (!token) return res.status(401).json({ ok: false, error: 'Token requerido' });

  const BC_TENANT = process.env.BC_TENANT;
  const BC_ENV = process.env.BC_ENV;
  const BC_COMPANY = process.env.BC_COMPANY;
  const headers = { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' };
  const companyEncoded = encodeURIComponent(BC_COMPANY);

  try {
    const url = `https://api.businesscentral.dynamics.com/v2.0/${BC_TENANT}/${BC_ENV}/ODataV4/Company('${companyEncoded}')/Hist_rico_facturas_compra_Excel?$top=500`;

    const invRes = await fetch(url, { headers });
    if (!invRes.ok) {
      const errText = await invRes.text().catch(() => invRes.statusText);
      throw new Error('Error Histórico_facturas_compra_Excel: ' + errText);
    }
    const invJson = await invRes.json();

    const data = (invJson.value || []).map(e => ({
      number: e.No ?? '',
      invoiceDate: e.Posting_Date ?? e.Document_Date ?? '',
      vendorInvoiceNumber: e.Vendor_Invoice_No ?? '',
      vendorNumber: e.Buy_from_Vendor_No ?? '',
      vendorName: e.Buy_from_Vendor_Name ?? '',
      dueDate: e.Due_Date ?? '',
      paymentTerms: e.Payment_Terms_Code ?? '',
      paymentMethod: e.Payment_Method_Code ?? '',
      cancelled: e.Cancelled ?? false
    }));

    return res.status(200).json({ ok: true, data });
  } catch (error) {
    console.error('BC facturas-pendientes-compra error:', error.message);
    return res.status(500).json({ ok: false, error: error.message });
  }
}
