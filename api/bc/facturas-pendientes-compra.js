// POST /api/bc/facturas-pendientes-compra
// Obtiene facturas de compra pendientes desde OData v4 — página Histórico facturas compra
// usa vendorLedgerEntries OData con filtro open=true

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
    const filter = encodeURIComponent("Document_Type eq 'Invoice' and Open eq true");
    const url = `https://api.businesscentral.dynamics.com/v2.0/${BC_TENANT}/${BC_ENV}/ODataV4/Company('${companyEncoded}')/Vendor_Ledger_Entry?$filter=${filter}&$top=500`;

    const invRes = await fetch(url, { headers });
    if (!invRes.ok) {
      const errText = await invRes.text().catch(() => invRes.statusText);
      throw new Error('Error Vendor_Ledger_Entry: ' + errText);
    }
    const invJson = await invRes.json();

    const data = (invJson.value || []).map(e => ({
      number: e.Document_No ?? e.Document_No_ ?? '',
      invoiceDate: e.Posting_Date ?? '',
      vendorInvoiceNumber: e.External_Document_No ?? e.External_Document_No_ ?? '',
      vendorNumber: e.Vendor_No ?? e.Vendor_No_ ?? '',
      vendorName: e.Vendor_Name ?? '',
      dueDate: e.Due_Date ?? '',
      totalAmountExcludingTax: null,
      totalAmountIncludingTax: Math.abs(parseFloat(e.Original_Amount) || 0),
      remainingAmount: Math.abs(parseFloat(e.Remaining_Amount) || 0)
    }));

    return res.status(200).json({ ok: true, data });
  } catch (error) {
    console.error('BC facturas-pendientes-compra error:', error.message);
    return res.status(500).json({ ok: false, error: error.message });
  }
}
