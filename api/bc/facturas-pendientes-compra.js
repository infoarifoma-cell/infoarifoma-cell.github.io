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
  const url = `https://api.businesscentral.dynamics.com/v2.0/${BC_TENANT}/${BC_ENV}/ODataV4/Company('${companyEncoded}')/Histórico_facturas_compra_Excel?$top=1`;
  const headers = { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' };

  try {
    const invRes = await fetch(url, { headers });
    if (!invRes.ok) {
      const errText = await invRes.text().catch(() => invRes.statusText);
      throw new Error('Error Factura_compra_Excel: ' + errText);
    }
    const invJson = await invRes.json();

    // DEBUG: devolver primer registro completo para ver nombres de campos
    return res.status(200).json({ ok: true, debug: true, fields: Object.keys((invJson.value||[{}])[0]), sample: (invJson.value||[])[0] });
  } catch (error) {
    console.error('BC facturas-pendientes-compra error:', error.message);
    return res.status(500).json({ ok: false, error: error.message });
  }
}
