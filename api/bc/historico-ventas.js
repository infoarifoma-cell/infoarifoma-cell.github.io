// POST /api/bc/historico-ventas
// Proxy: obtiene facturas de venta registradas (Posted Sales Invoices) de BC
// con información de pagos pendientes via Customer Ledger Entries

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ ok: false, error: 'Método no permitido' });
  }

  const { token, fechaDesde, fechaHasta } = req.body;
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

    // Construir filtro de fechas
    let filter = '';
    if (fechaDesde && fechaHasta) {
      filter = `invoiceDate ge ${fechaDesde} and invoiceDate le ${fechaHasta}`;
    } else if (fechaDesde) {
      filter = `invoiceDate ge ${fechaDesde}`;
    } else if (fechaHasta) {
      filter = `invoiceDate le ${fechaHasta}`;
    }

    // Obtener facturas de venta registradas (Posted Sales Invoices)
    let allInvoices = [];
    let url = `${base}(${cid})/salesInvoices?$filter=${encodeURIComponent(filter + (filter ? ' and ' : '') + "status eq 'Paid' or status eq 'Open' or status eq 'Draft'")}&$top=500&$orderby=invoiceDate desc`;

    // BC salesInvoices API devuelve solo facturas registradas (posted)
    // Alternativa: usar endpoint de posted sales invoices si disponible
    // Primero intentar con salesInvoices que incluye totalAmountIncludingTax
    url = `${base}(${cid})/salesInvoices?${filter ? '$filter=' + encodeURIComponent(filter) + '&' : ''}$top=500&$orderby=invoiceDate desc`;

    while (url) {
      const invRes = await fetch(url, { headers });
      if (!invRes.ok) {
        const errText = await invRes.text().catch(() => invRes.statusText);
        throw new Error('Error salesInvoices: ' + errText);
      }
      const invJson = await invRes.json();
      if (invJson.value) allInvoices = allInvoices.concat(invJson.value);
      url = invJson['@odata.nextLink'] || null;
    }

    // Obtener Customer Ledger Entries para info de pagos
    // Filtramos por Document Type = Invoice
    let ledgerFilter = "documentType eq 'Invoice'";
    if (fechaDesde) ledgerFilter += ` and postingDate ge ${fechaDesde}`;
    if (fechaHasta) ledgerFilter += ` and postingDate le ${fechaHasta}`;

    let allLedger = [];
    let ledgerUrl = `${base}(${cid})/customerLedgerEntries?$filter=${encodeURIComponent(ledgerFilter)}&$top=500`;

    while (ledgerUrl) {
      const lRes = await fetch(ledgerUrl, { headers });
      if (!lRes.ok) {
        // Si falla customer ledger entries, continuar sin info de pagos
        console.warn('Customer Ledger Entries no disponible, continuando sin datos de pago');
        break;
      }
      const lJson = await lRes.json();
      if (lJson.value) allLedger = allLedger.concat(lJson.value);
      ledgerUrl = lJson['@odata.nextLink'] || null;
    }

    // Mapear ledger entries por documentNumber para cruzar
    const ledgerMap = {};
    for (const le of allLedger) {
      ledgerMap[le.documentNumber] = {
        amount: le.amount || 0,
        remainingAmount: le.remainingAmount || 0,
        dueDate: le.dueDate || null,
        open: le.open || false,
        closedAtDate: le.closedAtDate || null
      };
    }

    // Combinar info de facturas con info de pagos
    const result = allInvoices.map(inv => {
      const ledger = ledgerMap[inv.number] || null;
      const importe = inv.totalAmountIncludingTax || inv.totalAmountExcludingTax || 0;
      const pendiente = ledger ? Math.abs(ledger.remainingAmount) : null;
      const vencimiento = ledger ? ledger.dueDate : (inv.dueDate || null);
      const abierta = ledger ? ledger.open : null;

      let estado = 'desconocido';
      if (ledger) {
        if (!ledger.open || ledger.remainingAmount === 0) {
          estado = 'pagada';
        } else if (Math.abs(ledger.remainingAmount) < Math.abs(ledger.amount)) {
          estado = 'parcial';
        } else {
          // Comprobar si vencida
          if (vencimiento && new Date(vencimiento) < new Date()) {
            estado = 'vencida';
          } else {
            estado = 'pendiente';
          }
        }
      } else {
        // Sin ledger entry: asumir estado de la factura
        if (inv.status === 'Paid') estado = 'pagada';
        else if (inv.status === 'Open' || inv.status === 'Draft') estado = 'pendiente';
      }

      return {
        numero: inv.number || '',
        fecha: inv.invoiceDate || '',
        vencimiento: vencimiento || '',
        clienteNombre: inv.customerName || inv.sellToCustomerName || '',
        clienteCod: inv.customerNumber || inv.sellToCustomerNumber || '',
        importe,
        pendiente: pendiente !== null ? pendiente : importe,
        estado,
        diasVencido: (estado === 'vencida' && vencimiento)
          ? Math.floor((Date.now() - new Date(vencimiento).getTime()) / 86400000)
          : 0
      };
    });

    return res.status(200).json({ ok: true, data: result });
  } catch (error) {
    console.error('BC histórico ventas error:', error.message);
    return res.status(500).json({ ok: false, error: error.message });
  }
}
