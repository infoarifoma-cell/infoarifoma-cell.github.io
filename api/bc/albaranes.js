// POST /api/bc/albaranes
// Obtener albaranes (salesOrders) de un cliente en BC, con sus líneas

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ ok: false, error: 'Método no permitido' });
  }

  const { token, customerNumber } = req.body;

  if (!token || typeof token !== 'string') {
    return res.status(401).json({ ok: false, error: 'Token requerido' });
  }

  if (!customerNumber) {
    return res.status(400).json({ ok: false, error: 'customerNumber requerido' });
  }

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

    const companyId = company.id;

    // Obtener salesOrders del cliente
    const filter = `customerNumber eq '${customerNumber}'`;
    const ordersRes = await fetch(
      `${base}(${companyId})/salesOrders?$filter=${encodeURIComponent(filter)}&$select=id,number,externalDocumentNumber,orderDate,totalAmountExcludingTax,totalAmountIncludingTax,status`,
      { headers }
    );
    if (!ordersRes.ok) throw new Error('Error obteniendo pedidos: ' + ordersRes.statusText);

    const ordersJson = await ordersRes.json();
    const orders = ordersJson.value || [];

    // Para cada order, obtener sus líneas
    const result = [];
    for (const order of orders) {
      const linesRes = await fetch(
        `${base}(${companyId})/salesOrders(${order.id})/salesOrderLines?$select=id,lineObjectNumber,description,quantity,unitPrice,lineAmount,lineType`,
        { headers }
      );
      let lines = [];
      if (linesRes.ok) {
        const linesJson = await linesRes.json();
        lines = (linesJson.value || []).filter(l => l.lineType === 'Item');
      }

      result.push({
        id: order.id,
        number: order.number,
        externalDocumentNumber: order.externalDocumentNumber || '',
        orderDate: order.orderDate || '',
        totalAmount: order.totalAmountExcludingTax || 0,
        totalAmountInc: order.totalAmountIncludingTax || 0,
        status: order.status || '',
        lines
      });
    }

    return res.status(200).json({ ok: true, orders: result });
  } catch (error) {
    console.error('BC albaranes error:', error.message);
    return res.status(500).json({ ok: false, error: error.message });
  }
}
