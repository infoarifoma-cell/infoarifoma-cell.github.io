// POST /api/bc/historico-ventas
// Proxy: obtiene facturas de venta de BC con estado de pago
// Usa salesInvoices (v2.0) que incluye status, remainingAmount, dueDate

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
    const filters = [];
    if (fechaDesde) filters.push(`invoiceDate ge ${fechaDesde}`);
    if (fechaHasta) filters.push(`invoiceDate le ${fechaHasta}`);
    const filterStr = filters.length ? '$filter=' + encodeURIComponent(filters.join(' and ')) + '&' : '';

    // 1. Obtener salesInvoices (posted = status Open/Paid, draft = Draft)
    let allInvoices = [];
    let url = `${base}(${cid})/salesInvoices?${filterStr}$top=500&$orderby=invoiceDate desc`;

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

    // Excluir borradores — solo registradas (Open/Paid)
    allInvoices = allInvoices.filter(inv => inv.status === 'Open' || inv.status === 'Paid');

    // 2. Intentar Customer Ledger Entries como fuente complementaria
    let ledgerMap = {};
    try {
      const ledgerFilters = ["documentType eq 'Invoice'"];
      if (fechaDesde) ledgerFilters.push(`postingDate ge ${fechaDesde}`);
      if (fechaHasta) ledgerFilters.push(`postingDate le ${fechaHasta}`);

      let allLedger = [];
      let ledgerUrl = `${base}(${cid})/customerLedgerEntries?$filter=${encodeURIComponent(ledgerFilters.join(' and '))}&$top=500`;

      while (ledgerUrl) {
        const lRes = await fetch(ledgerUrl, { headers });
        if (!lRes.ok) break; // No disponible, seguir sin datos extra
        const lJson = await lRes.json();
        if (lJson.value) allLedger = allLedger.concat(lJson.value);
        ledgerUrl = lJson['@odata.nextLink'] || null;
      }

      for (const le of allLedger) {
        ledgerMap[le.documentNumber] = {
          amount: le.amount || 0,
          remainingAmount: le.remainingAmount || 0,
          dueDate: le.dueDate || null,
          open: le.open !== undefined ? le.open : true,
        };
      }
    } catch (e) {
      console.warn('Customer Ledger Entries fallback:', e.message);
    }

    const now = new Date();

    const result = allInvoices.map(inv => {
      const ledger = ledgerMap[inv.number] || null;

      // Importe: usar totalAmountIncludingTax de la factura
      const importe = inv.totalAmountIncludingTax || inv.totalAmountExcludingTax || 0;

      // Pendiente: preferir remainingAmount del ledger, luego de la factura
      let pendiente;
      if (ledger && ledger.remainingAmount !== undefined) {
        pendiente = Math.abs(ledger.remainingAmount);
      } else if (inv.remainingAmount !== undefined) {
        pendiente = Math.abs(inv.remainingAmount);
      } else {
        // Sin dato de remaining: deducir del status
        pendiente = inv.status === 'Paid' ? 0 : Math.abs(importe);
      }

      // Vencimiento: preferir ledger, luego factura
      const vencimiento = (ledger && ledger.dueDate) || inv.dueDate || null;

      // Estado
      let estado;
      if (pendiente === 0 || inv.status === 'Paid') {
        estado = 'pagada';
      } else if (pendiente > 0 && pendiente < Math.abs(importe)) {
        estado = 'parcial';
      } else if (vencimiento && new Date(vencimiento) < now) {
        estado = 'vencida';
      } else {
        estado = 'pendiente';
      }

      const diasVencido = (estado === 'vencida' && vencimiento)
        ? Math.floor((now.getTime() - new Date(vencimiento).getTime()) / 86400000)
        : 0;

      return {
        numero: inv.number || '',
        fecha: inv.invoiceDate || '',
        vencimiento: vencimiento || '',
        clienteNombre: inv.customerName || inv.sellToCustomerName || '',
        clienteCod: inv.customerNumber || inv.sellToCustomerNumber || '',
        importe,
        pendiente,
        estado,
        diasVencido
      };
    });

    return res.status(200).json({ ok: true, data: result });
  } catch (error) {
    console.error('BC histórico ventas error:', error.message);
    return res.status(500).json({ ok: false, error: error.message });
  }
}
