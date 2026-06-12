// POST /api/bc
// Consolidated BC proxy — routes by req.body.action
// Actions: items, vendors, facturas, facturas-pendientes, historico-ventas, linea-pesada, pedido-compra, costes

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ ok: false, error: 'Método no permitido' });
  }

  const { action, token } = req.body;
  if (!token || typeof token !== 'string') {
    return res.status(401).json({ ok: false, error: 'Token requerido' });
  }

  const BC_TENANT = process.env.BC_TENANT;
  const BC_ENV = process.env.BC_ENV;
  const BC_COMPANY = process.env.BC_COMPANY;
  const base = `https://api.businesscentral.dynamics.com/v2.0/${BC_TENANT}/${BC_ENV}/api/v2.0/companies`;
  const headers = { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' };

  const odataSafe = (val) => String(val || '').replace(/'/g, "''");

  async function getCompanyId() {
    const cRes = await fetch(base, { headers });
    if (!cRes.ok) throw new Error('No se pudo obtener company: ' + cRes.statusText);
    const cJson = await cRes.json();
    const company = cJson.value.find(c => c.name === BC_COMPANY);
    if (!company) throw new Error('Company no encontrada: ' + BC_COMPANY);
    return company.id;
  }

  try {
    switch (action) {

      // ── ITEMS ─────────────────────────────────────────────
      case 'items': {
        const cid = await getCompanyId();
        let items = [];
        let url = `${base}(${cid})/items?$select=number,displayName,unitPrice&$top=500`;
        while (url) {
          const r = await fetch(url, { headers });
          if (!r.ok) throw new Error('Error obteniendo artículos: ' + r.statusText);
          const j = await r.json();
          items = items.concat((j.value || []).map(i => ({ number: i.number, displayName: i.displayName, unitPrice: i.unitPrice })));
          url = j['@odata.nextLink'] || null;
        }
        return res.status(200).json({ ok: true, items });
      }

      // ── VENDORS ───────────────────────────────────────────
      case 'vendors': {
        const cid = await getCompanyId();
        let vendors = [];
        let url = `${base}(${cid})/vendors?$select=number,displayName&$top=500`;
        while (url) {
          const r = await fetch(url, { headers });
          if (!r.ok) throw new Error('Error obteniendo vendors: ' + r.statusText);
          const j = await r.json();
          vendors = vendors.concat((j.value || []).map(v => ({ number: v.number, name: v.displayName })));
          url = j['@odata.nextLink'] || null;
        }
        return res.status(200).json({ ok: true, vendors });
      }

      // ── FACTURAS (crear factura venta) ────────────────────
      case 'facturas': {
        const { customerNumber, invoiceDate, externalDocumentNumber } = req.body;
        if (!customerNumber || typeof customerNumber !== 'string') {
          return res.status(400).json({ ok: false, error: 'customerNumber requerido' });
        }
        if (invoiceDate && !/^\d{4}-\d{2}-\d{2}$/.test(invoiceDate)) {
          return res.status(400).json({ ok: false, error: 'Formato de fecha inválido' });
        }
        const cid = await getCompanyId();
        const invRes = await fetch(`${base}(${cid})/salesInvoices`, {
          method: 'POST', headers,
          body: JSON.stringify({ customerNumber, invoiceDate, externalDocumentNumber })
        });
        if (!invRes.ok) throw new Error('No se pudo crear factura: ' + await invRes.text());
        const invoice = await invRes.json();
        return res.status(200).json({ ok: true, invoice });
      }

      // ── FACTURAS PENDIENTES ───────────────────────────────
      case 'facturas-pendientes': {
        const { type } = req.body;
        if (type !== 'venta' && type !== 'compra') {
          return res.status(400).json({ ok: false, error: 'type debe ser venta o compra' });
        }
        const cid = await getCompanyId();

        if (type === 'venta') {
          const filter = encodeURIComponent("status eq 'open'");
          const select = '$select=number,invoiceDate,customerName,dueDate,totalAmountIncludingTax,remainingAmount';
          const url = `${base}(${cid})/salesInvoices?$filter=${filter}&${select}&$top=500`;
          const invRes = await fetch(url, { headers });
          if (!invRes.ok) throw new Error('Error salesInvoices: ' + await invRes.text().catch(() => invRes.statusText));
          const invJson = await invRes.json();
          const data = (invJson.value || []).map(inv => ({
            number: inv.number, invoiceDate: inv.invoiceDate, customerName: inv.customerName,
            dueDate: inv.dueDate, totalAmountIncludingTax: inv.totalAmountIncludingTax, remainingAmount: inv.remainingAmount
          }));
          return res.status(200).json({ ok: true, data });
        }

        // compra
        let all = [];
        let url = `${base}(${cid})/purchaseInvoices?$filter=status eq 'open'&$select=number,invoiceDate,vendorInvoiceNumber,vendorName,totalAmountExcludingTax,totalAmountIncludingTax,dueDate&$orderby=invoiceDate desc&$top=500`;
        while (url) {
          const r = await fetch(url, { headers });
          if (!r.ok) throw new Error('Error purchaseInvoices: ' + await r.text().catch(() => r.statusText));
          const j = await r.json();
          all = all.concat(j.value || []);
          url = j['@odata.nextLink'] || null;
        }
        const data = all.map(inv => ({
          number: inv.number, invoiceDate: inv.invoiceDate, vendorInvoiceNumber: inv.vendorInvoiceNumber,
          vendorName: inv.vendorName, dueDate: inv.dueDate, totalAmountExcludingTax: inv.totalAmountExcludingTax,
          totalAmountIncludingTax: inv.totalAmountIncludingTax, remainingAmount: null
        }));
        return res.status(200).json({ ok: true, data });
      }

      // ── HISTORICO VENTAS ──────────────────────────────────
      case 'historico-ventas': {
        const { fechaDesde, fechaHasta } = req.body;
        const fechaValida = (f) => !f || /^\d{4}-\d{2}-\d{2}$/.test(f);
        if (!fechaValida(fechaDesde) || !fechaValida(fechaHasta)) {
          return res.status(400).json({ ok: false, error: 'Formato de fecha inválido' });
        }
        const cid = await getCompanyId();

        const filters = [];
        if (fechaDesde) filters.push(`invoiceDate ge ${fechaDesde}`);
        if (fechaHasta) filters.push(`invoiceDate le ${fechaHasta}`);
        const filterStr = filters.length ? '$filter=' + encodeURIComponent(filters.join(' and ')) + '&' : '';

        let allInvoices = [];
        let url = `${base}(${cid})/salesInvoices?${filterStr}$top=500&$orderby=invoiceDate desc`;
        while (url) {
          const r = await fetch(url, { headers });
          if (!r.ok) throw new Error('Error salesInvoices: ' + await r.text().catch(() => r.statusText));
          const j = await r.json();
          if (j.value) allInvoices = allInvoices.concat(j.value);
          url = j['@odata.nextLink'] || null;
        }
        allInvoices = allInvoices.filter(inv => inv.status === 'Open' || inv.status === 'Paid');

        let customerEmails = {};
        try {
          let custUrl = `${base}(${cid})/customers?$select=number,displayName,email&$top=500`;
          while (custUrl) {
            const r = await fetch(custUrl, { headers });
            if (!r.ok) break;
            const j = await r.json();
            for (const c of (j.value || [])) { if (c.number) customerEmails[c.number] = c.email || ''; }
            custUrl = j['@odata.nextLink'] || null;
          }
        } catch (e) { /* no crítico */ }

        let ledgerMap = {};
        try {
          const lf = ["documentType eq 'Invoice'"];
          if (fechaDesde) lf.push(`postingDate ge ${fechaDesde}`);
          if (fechaHasta) lf.push(`postingDate le ${fechaHasta}`);
          let allLedger = [];
          let lUrl = `${base}(${cid})/customerLedgerEntries?$filter=${encodeURIComponent(lf.join(' and '))}&$top=500`;
          while (lUrl) {
            const r = await fetch(lUrl, { headers });
            if (!r.ok) break;
            const j = await r.json();
            if (j.value) allLedger = allLedger.concat(j.value);
            lUrl = j['@odata.nextLink'] || null;
          }
          for (const le of allLedger) {
            ledgerMap[le.documentNumber] = {
              amount: le.amount || 0, remainingAmount: le.remainingAmount || 0,
              dueDate: le.dueDate || null, open: le.open !== undefined ? le.open : true
            };
          }
        } catch (e) { /* fallback */ }

        const now = new Date();
        const result = allInvoices.map(inv => {
          const ledger = ledgerMap[inv.number] || null;
          const importe = inv.totalAmountIncludingTax || inv.totalAmountExcludingTax || 0;
          let pendiente;
          if (ledger && ledger.remainingAmount !== undefined) pendiente = Math.abs(ledger.remainingAmount);
          else if (inv.remainingAmount !== undefined) pendiente = Math.abs(inv.remainingAmount);
          else pendiente = inv.status === 'Paid' ? 0 : Math.abs(importe);
          const vencimiento = (ledger && ledger.dueDate) || inv.dueDate || null;
          let estado;
          if (pendiente === 0 || inv.status === 'Paid') estado = 'pagada';
          else if (vencimiento && new Date(vencimiento) < now) estado = 'vencida';
          else estado = 'pendiente';
          const diasVencido = (estado === 'vencida' && vencimiento) ? Math.floor((now.getTime() - new Date(vencimiento).getTime()) / 86400000) : 0;
          const custNum = inv.customerNumber || inv.sellToCustomerNumber || '';
          return {
            numero: inv.number || '', fecha: inv.invoiceDate || '', vencimiento: vencimiento || '',
            clienteNombre: inv.customerName || inv.sellToCustomerName || '', clienteCod: custNum,
            clienteEmail: customerEmails[custNum] || '', importe, pendiente, estado, diasVencido
          };
        });
        return res.status(200).json({ ok: true, data: result });
      }

      // ── LINEA PESADA ──────────────────────────────────────
      case 'linea-pesada': {
        const { codigoCliente, proyectoCod, productoCod, productoNombre, pesoNeto, matriculacam, proyectoName } = req.body;
        if (!codigoCliente || !proyectoCod) {
          return res.status(400).json({ ok: false, error: 'codigoCliente y proyectoCod requeridos' });
        }
        const cid = await getCompanyId();

        const filter = `customerNumber eq '${odataSafe(codigoCliente)}' and externalDocumentNumber eq '${odataSafe(proyectoCod)}'`;
        const ordersRes = await fetch(`${base}(${cid})/salesOrders?$filter=${encodeURIComponent(filter)}&$select=id,number`, { headers });
        const ordersJson = await ordersRes.json();

        let orderId;
        if (ordersJson.value && ordersJson.value.length > 0) {
          orderId = ordersJson.value[0].id;
        } else {
          const newOrderRes = await fetch(`${base}(${cid})/salesOrders`, {
            method: 'POST', headers,
            body: JSON.stringify({ customerNumber: codigoCliente, externalDocumentNumber: proyectoCod })
          });
          if (!newOrderRes.ok) throw new Error('No se pudo crear pedido: ' + await newOrderRes.text());
          const newOrder = await newOrderRes.json();
          orderId = newOrder.id;
        }

        const lineRes = await fetch(`${base}(${cid})/salesOrders(${orderId})/salesOrderLines`, {
          method: 'POST', headers,
          body: JSON.stringify({
            lineType: 'Item', lineObjectNumber: productoCod,
            description: `${productoNombre} | ${proyectoName || proyectoCod} | ${(Number(pesoNeto) / 1000).toFixed(3)} Tn | ${matriculacam}`,
            quantity: parseFloat((Number(pesoNeto) / 1000).toFixed(3)), unitPrice: 0
          })
        });
        if (!lineRes.ok) throw new Error('No se pudo crear línea: ' + await lineRes.text());
        const lineJson = await lineRes.json();
        const docNum = lineJson.documentNumber || '';
        const lineSeq = lineJson.sequence || lineJson.lineSequenceNumber || lineJson.lineNumber || '';
        const numalbarancalle = docNum && lineSeq ? `${docNum}/${lineSeq}` : docNum || null;
        return res.status(200).json({ ok: true, numalbarancalle });
      }

      // ── PEDIDO COMPRA ─────────────────────────────────────
      case 'pedido-compra': {
        const { vendorName, orderDate, vendorInvoiceNumber, itemNumber, quantity, unitPrice } = req.body;
        if (!vendorName) return res.status(400).json({ ok: false, error: 'vendorName requerido' });
        if (orderDate && !/^\d{4}-\d{2}-\d{2}$/.test(orderDate)) {
          return res.status(400).json({ ok: false, error: 'Formato de fecha inválido (YYYY-MM-DD)' });
        }
        const BC_COMPANY_ODATA = process.env.BC_COMPANY_ODATA || BC_COMPANY;
        const odataBase = `https://api.businesscentral.dynamics.com/v2.0/${BC_TENANT}/${BC_ENV}/ODataV4/Company('${encodeURIComponent(BC_COMPANY_ODATA)}')`;
        const cid = await getCompanyId();

        // Buscar proveedor
        let vendorNumber = '';
        const exactFilter = `displayName eq '${odataSafe(vendorName)}'`;
        const exactRes = await fetch(`${base}(${cid})/vendors?$filter=${encodeURIComponent(exactFilter)}&$select=id,number,displayName&$top=1`, { headers });
        if (!exactRes.ok) throw new Error('Error buscando proveedor: ' + exactRes.statusText);
        const exactJson = await exactRes.json();

        if (exactJson.value && exactJson.value.length > 0) {
          vendorNumber = exactJson.value[0].number;
        } else {
          const partialFilter = `contains(displayName,'${odataSafe(vendorName)}')`;
          const partialRes = await fetch(`${base}(${cid})/vendors?$filter=${encodeURIComponent(partialFilter)}&$select=id,number,displayName&$top=5`, { headers });
          if (partialRes.ok) {
            const partialJson = await partialRes.json();
            if (partialJson.value && partialJson.value.length === 1) {
              vendorNumber = partialJson.value[0].number;
            } else if (partialJson.value && partialJson.value.length > 1) {
              return res.status(404).json({ ok: false, error: `Varios proveedores coinciden con "${vendorName}": ${partialJson.value.map(v => v.displayName).join(', ')}.` });
            }
          }
          if (!vendorNumber) return res.status(404).json({ ok: false, error: `Proveedor "${vendorName}" no encontrado en BC.` });
        }

        const orderBody = { vendorNumber };
        if (orderDate) orderBody.orderDate = orderDate;
        const orderRes = await fetch(`${base}(${cid})/purchaseOrders`, { method: 'POST', headers, body: JSON.stringify(orderBody) });
        if (!orderRes.ok) throw new Error('No se pudo crear pedido de compra: ' + await orderRes.text());
        const order = await orderRes.json();

        if (vendorInvoiceNumber) {
          const patchOdata = await fetch(`${odataBase}/PurchaseOrder(Document_Type='Order',No='${odataSafe(order.number)}')`, {
            method: 'PATCH', headers: { ...headers, 'If-Match': '*' },
            body: JSON.stringify({ Vendor_Invoice_No: vendorInvoiceNumber })
          });
          if (!patchOdata.ok) console.warn('PATCH Vendor_Invoice_No falló:', await patchOdata.text());
        }

        if (itemNumber && quantity) {
          const lineBody = { documentId: order.id, lineType: 'Item', lineObjectNumber: itemNumber, quantity: Number(quantity) };
          if (unitPrice) lineBody.directUnitCost = Number(unitPrice);
          const lineRes = await fetch(`${base}(${cid})/purchaseOrders(${order.id})/purchaseOrderLines`, { method: 'POST', headers, body: JSON.stringify(lineBody) });
          if (!lineRes.ok) console.warn('No se pudo crear línea:', await lineRes.text());
        }

        return res.status(200).json({ ok: true, order });
      }

      // ── COSTES ────────────────────────────────────────────
      case 'costes': {
        const { anyo } = req.body;
        if (!anyo || !/^\d{4}$/.test(String(anyo))) return res.status(400).json({ ok: false, error: 'Año inválido' });
        const cid = await getCompanyId();

        const coaMap = {};
        try {
          const coaUrl = `${base}(${cid})/accounts?$filter=startswith(number,'6') or startswith(number,'7')&$select=number,displayName&$top=500`;
          const coaRes = await fetch(coaUrl, { headers });
          if (coaRes.ok) {
            const coaJson = await coaRes.json();
            for (const acc of (coaJson.value || [])) coaMap[acc.number] = acc.displayName || '';
          }
        } catch(e) { /* no crítico */ }

        let allEntries = [];
        for (let mes = 1; mes <= 12; mes++) {
          const mm = String(mes).padStart(2, '0');
          const lastDay = new Date(Number(anyo), mes, 0).getDate();
          const fechaInicio = `${anyo}-${mm}-01`;
          const fechaFin = `${anyo}-${mm}-${String(lastDay).padStart(2, '0')}`;
          const filter = `postingDate ge ${fechaInicio} and postingDate le ${fechaFin}`;
          let url = `${base}(${cid})/generalLedgerEntries?$filter=${encodeURIComponent(filter)}&$expand=dimensionSetLines&$top=1000`;
          while (url) {
            const r = await fetch(url, { headers });
            if (!r.ok) throw new Error(`Error G/L Entries mes ${mes}: ` + await r.text().catch(() => r.statusText));
            const j = await r.json();
            for (const entry of (j.value || [])) {
              const acc = entry.accountNumber || '';
              if (!acc.startsWith('6') && !acc.startsWith('7')) continue;
              const dims = {};
              for (const dim of (entry.dimensionSetLines || [])) dims[dim.code] = { valueCode: dim.valueCode, displayName: dim.valueDisplayName };
              allEntries.push({
                date: entry.postingDate, account: entry.accountNumber, accountName: coaMap[entry.accountNumber] || '',
                description: entry.description, debit: entry.debitAmount, credit: entry.creditAmount,
                ca: dims.CA?.valueCode || null, caName: dims.CA?.displayName || null,
                proyecto: dims.PROYECTO?.valueCode || null, proyectoName: dims.PROYECTO?.displayName || null,
                docNumber: entry.documentNumber
              });
            }
            url = j['@odata.nextLink'] || null;
          }
        }
        return res.status(200).json({ ok: true, count: allEntries.length, entries: allEntries });
      }

      default:
        return res.status(400).json({ ok: false, error: 'Acción desconocida: ' + action });
    }
  } catch (error) {
    console.error('BC proxy error:', error.message);
    return res.status(500).json({ ok: false, error: error.message });
  }
}
