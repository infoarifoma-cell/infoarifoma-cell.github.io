// POST /api/bc/pedido-compra
// Crear pedido de compra en BC desde factura escaneada

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ ok: false, error: 'Método no permitido' });
  }

  const { token, vendorName, orderDate, vendorInvoiceNumber, itemNumber, quantity, unitPrice } = req.body;

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
  const BC_COMPANY_ODATA = process.env.BC_COMPANY_ODATA || BC_COMPANY;
  const base = `https://api.businesscentral.dynamics.com/v2.0/${BC_TENANT}/${BC_ENV}/api/v2.0/companies`;
  const odataBase = `https://api.businesscentral.dynamics.com/v2.0/${BC_TENANT}/${BC_ENV}/ODataV4/Company('${encodeURIComponent(BC_COMPANY_ODATA)}')`;
  const headers = { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' };

  const odataSafe = (val) => String(val || '').replace(/'/g, "''");

  try {
    // Obtener company ID (para API v2.0)
    const cRes = await fetch(base, { headers });
    if (!cRes.ok) throw new Error('No se pudo obtener company: ' + cRes.statusText);
    const cJson = await cRes.json();
    const company = cJson.value.find(c => c.name === BC_COMPANY);
    if (!company) throw new Error('Company no encontrada: ' + BC_COMPANY);
    const companyId = company.id;

    // Buscar proveedor
    let vendorNumber = '';
    const exactFilter = `displayName eq '${odataSafe(vendorName)}'`;
    const exactRes = await fetch(
      `${base}(${companyId})/vendors?$filter=${encodeURIComponent(exactFilter)}&$select=id,number,displayName&$top=1`,
      { headers }
    );
    if (!exactRes.ok) throw new Error('Error buscando proveedor: ' + exactRes.statusText);
    const exactJson = await exactRes.json();

    if (exactJson.value && exactJson.value.length > 0) {
      vendorNumber = exactJson.value[0].number;
    } else {
      const partialFilter = `contains(displayName,'${odataSafe(vendorName)}')`;
      const partialRes = await fetch(
        `${base}(${companyId})/vendors?$filter=${encodeURIComponent(partialFilter)}&$select=id,number,displayName&$top=5`,
        { headers }
      );
      if (partialRes.ok) {
        const partialJson = await partialRes.json();
        if (partialJson.value && partialJson.value.length === 1) {
          vendorNumber = partialJson.value[0].number;
        } else if (partialJson.value && partialJson.value.length > 1) {
          const nombres = partialJson.value.map(v => v.displayName).join(', ');
          return res.status(404).json({ ok: false, error: `Varios proveedores coinciden con "${vendorName}": ${nombres}.` });
        }
      }
      if (!vendorNumber) {
        return res.status(404).json({ ok: false, error: `Proveedor "${vendorName}" no encontrado en BC.` });
      }
    }

    // Crear pedido via API v2.0
    const orderBody = { vendorNumber };
    if (orderDate) orderBody.orderDate = orderDate;

    const orderRes = await fetch(`${base}(${companyId})/purchaseOrders`, {
      method: 'POST',
      headers,
      body: JSON.stringify(orderBody)
    });
    if (!orderRes.ok) throw new Error('No se pudo crear pedido de compra: ' + await orderRes.text());
    const order = await orderRes.json();

    // PATCH Vendor_Invoice_No via ODataV4 (API v2.0 no expone este campo)
    if (vendorInvoiceNumber) {
      const patchOdata = await fetch(`${odataBase}/PurchaseOrder(Document_Type='Order',No='${odataSafe(order.number)}')`, {
        method: 'PATCH',
        headers: { ...headers, 'If-Match': '*' },
        body: JSON.stringify({ Vendor_Invoice_No: vendorInvoiceNumber })
      });
      if (!patchOdata.ok) console.warn('PATCH Vendor_Invoice_No falló:', await patchOdata.text());
      else console.log('PATCH Vendor_Invoice_No OK');
    }

    // Añadir línea via API v2.0
    if (itemNumber && quantity) {
      const lineBody = {
        documentId: order.id,
        lineType: 'Item',
        lineObjectNumber: itemNumber,
        quantity: Number(quantity)
      };
      if (unitPrice) lineBody.directUnitCost = Number(unitPrice);

      const lineRes = await fetch(`${base}(${companyId})/purchaseOrders(${order.id})/purchaseOrderLines`, {
        method: 'POST',
        headers,
        body: JSON.stringify(lineBody)
      });
      if (!lineRes.ok) console.warn('No se pudo crear línea:', await lineRes.text());
    }

    return res.status(200).json({ ok: true, order });
  } catch (error) {
    console.error('BC pedido-compra error:', error.message);
    return res.status(500).json({ ok: false, error: error.message });
  }
}
