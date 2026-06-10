// POST /api/bc/facturas-pendientes-compra
// Obtiene facturas de compra con importe pendiente desde OData BC (página 138 Histórico_facturas_compra_Excel)
// Campos: No_, Posting_Date, Vendor_Invoice_No_, Buy_from_Vendor_No_, Buy_from_Vendor_Name,
//         Amount, Amount_Including_VAT, Remaining_Amount, Due_Date

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ ok: false, error: 'Método no permitido' });
  }

  const { token } = req.body;
  if (!token) return res.status(401).json({ ok: false, error: 'Token requerido' });

  const BC_TENANT = process.env.BC_TENANT;
  const BC_ENV = process.env.BC_ENV;
  const BC_COMPANY = process.env.BC_COMPANY;

  const companyEncoded = encodeURIComponent(BC_COMPANY);
  const filter = encodeURIComponent("Remaining_Amount gt 0");
  const url = `https://api.businesscentral.dynamics.com/v2.0/${BC_TENANT}/${BC_ENV}/ODataV4/Company('${companyEncoded}')/Histórico_facturas_compra_Excel?$filter=${filter}&$top=500`;
  const headers = { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' };

  try {
    const invRes = await fetch(url, { headers });
    if (!invRes.ok) {
      const errText = await invRes.text().catch(() => invRes.statusText);
      throw new Error('Error Factura_compra_Excel: ' + errText);
    }
    const invJson = await invRes.json();

    const data = (invJson.value || []).map(inv => ({
      number: inv.No_ ?? inv.No ?? inv['No_'] ?? '',
      postingDate: inv.Posting_Date ?? inv.Document_Date ?? '',
      invoiceDate: inv.Document_Date ?? inv.Posting_Date ?? '',
      vendorInvoiceNumber: inv.Vendor_Invoice_No_ ?? inv.External_Document_No_ ?? '',
      vendorNumber: inv.Buy_from_Vendor_No_ ?? '',
      vendorName: inv.Buy_from_Vendor_Name ?? '',
      totalAmountExcludingTax: inv.Amount ?? 0,
      totalAmountIncludingTax: inv.Amount_Including_VAT ?? 0,
      remainingAmount: inv.Remaining_Amount ?? inv.Amount_Including_VAT ?? 0,
      dueDate: inv.Due_Date ?? ''
    }));

    return res.status(200).json({ ok: true, data });
  } catch (error) {
    console.error('BC facturas-pendientes-compra error:', error.message);
    return res.status(500).json({ ok: false, error: error.message });
  }
}
