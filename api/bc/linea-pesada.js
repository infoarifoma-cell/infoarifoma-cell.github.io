// POST /api/bc/linea-pesada
// Proxy seguro: Frontend envía token MSAL, backend envía línea a BC

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ ok: false, error: 'Método no permitido' });
  }

  const { token, codigoCliente, proyectoCod, productoCod, productoNombre, pesoNeto, matriculacam, proyectoName } = req.body;

  // Validar token (básico)
  if (!token || typeof token !== 'string') {
    return res.status(401).json({ ok: false, error: 'Token requerido' });
  }

  // Sanitizar inputs contra OData injection
  const odataSafe = (val) => String(val || '').replace(/'/g, "''");
  if (!codigoCliente || !proyectoCod) {
    return res.status(400).json({ ok: false, error: 'codigoCliente y proyectoCod requeridos' });
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

    // Buscar o crear pedido
    const filter = `customerNumber eq '${odataSafe(codigoCliente)}' and externalDocumentNumber eq '${odataSafe(proyectoCod)}'`;
    const ordersRes = await fetch(`${base}(${companyId})/salesOrders?$filter=${encodeURIComponent(filter)}&$select=id,number`, { headers });
    const ordersJson = await ordersRes.json();

    let orderId;
    if (ordersJson.value && ordersJson.value.length > 0) {
      orderId = ordersJson.value[0].id;
    } else {
      const newOrderRes = await fetch(`${base}(${companyId})/salesOrders`, {
        method: 'POST',
        headers,
        body: JSON.stringify({
          customerNumber: codigoCliente,
          externalDocumentNumber: proyectoCod
        })
      });
      if (!newOrderRes.ok) throw new Error('No se pudo crear pedido: ' + await newOrderRes.text());
      const newOrder = await newOrderRes.json();
      orderId = newOrder.id;
    }

    // Agregar línea al pedido
    const lineRes = await fetch(`${base}(${companyId})/salesOrders(${orderId})/salesOrderLines`, {
      method: 'POST',
      headers,
      body: JSON.stringify({
        lineType: 'Item',
        lineObjectNumber: productoCod,
        description: `${productoNombre} | ${proyectoName || proyectoCod} | ${(Number(pesoNeto) / 1000).toFixed(3)} Tn | ${matriculacam}`,
        quantity: parseFloat((Number(pesoNeto) / 1000).toFixed(3)),
        unitPrice: 0
      })
    });

    if (!lineRes.ok) throw new Error('No se pudo crear línea: ' + await lineRes.text());

    const lineJson = await lineRes.json();
    const docNum = lineJson.documentNumber || '';
    const lineSeq = lineJson.sequence || lineJson.lineSequenceNumber || lineJson.lineNumber || '';
    const numalbarancalle = docNum && lineSeq ? `${docNum}/${lineSeq}` : docNum || null;

    return res.status(200).json({ ok: true, numalbarancalle });
  } catch (error) {
    console.error('BC error:', error.message);
    return res.status(500).json({ ok: false, error: error.message });
  }
}
