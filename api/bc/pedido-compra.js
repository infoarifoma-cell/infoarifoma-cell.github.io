// POST /api/bc/pedido-compra
// Crear pedido de compra en BC desde factura escaneada

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ ok: false, error: 'Método no permitido' });
  }

  const { token, vendorName, orderDate, vendorInvoiceNumber } = req.body;

  if (!token || typeof token !== 'string') {
    return res.status(401).json({ ok: false, error: 'Token requerido' });
  }
  if (!vendorName) {
    return res.status(400).json({ ok: false, error: 'vendorName requerido' });
  }
  if (orderDate && !/^\d{4}-\d{2}-\d{2}$/.test(orderDate)) {
    return res.status(400).json({ ok: false, error: 'Formato de fecha inválido (YYYY-MM-DD)' });
  }

  const BC_TENANT = process.env.BC_TENANT;
  const BC_ENV = process.env.BC_ENV;
  const BC_COMPANY = process.env.BC_COMPANY;
  const base = `https://api.businesscentral.dynamics.com/v2.0/${BC_TENANT}/${BC_ENV}/api/v2.0/companies`;
  const headers = { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' };

  const odataSafe = (val) => String(val || '').replace(/'/g, "''");

  try {
    // Obtener company ID
    const cRes = await fetch(base, { headers });
    if (!cRes.ok) throw new Error('No se pudo obtener company: ' + cRes.statusText);

    const cJson = await cRes.json();
    const company = cJson.value.find(c => c.name === BC_COMPANY);
    if (!company) throw new Error('Company no encontrada: ' + BC_COMPANY);

    const companyId = company.id;

    // Buscar proveedor por nombre
    const vendorFilter = `displayName eq '${odataSafe(vendorName)}'`;
    const vendorRes = await fetch(
      `${base}(${companyId})/vendors?$filter=${encodeURIComponent(vendorFilter)}&$select=id,number,displayName&$top=1`,
      { headers }
    );
    if (!vendorRes.ok) throw new Error('Error buscando proveedor: ' + vendorRes.statusText);

    const vendorJson = await vendorRes.json();
    let vendorNumber = '';
    if (vendorJson.value && vendorJson.value.length > 0) {
      vendorNumber = vendorJson.value[0].number;
    } else {
      return res.status(404).json({
        ok: false,
        error: `Proveedor "${vendorName}" no encontrado en BC. Créalo primero o verifica el nombre.`
      });
    }

    // Crear purchase order
    const orderBody = { vendorNumber };
    if (orderDate) orderBody.orderDate = orderDate;
    if (vendorInvoiceNumber) orderBody.vendorInvoiceNumber = vendorInvoiceNumber;

    const orderRes = await fetch(`${base}(${companyId})/purchaseOrders`, {
      method: 'POST',
      headers,
      body: JSON.stringify(orderBody)
    });

    if (!orderRes.ok) throw new Error('No se pudo crear pedido de compra: ' + await orderRes.text());

    const order = await orderRes.json();
    return res.status(200).json({ ok: true, order });
  } catch (error) {
    console.error('BC pedido-compra error:', error.message);
    return res.status(500).json({ ok: false, error: error.message });
  }
}
