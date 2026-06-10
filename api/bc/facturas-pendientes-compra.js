// POST /api/bc/facturas-pendientes-compra
// PENDIENTE: BC no expone vendorLedgerEntries en API v2.0 ni OData estándar
// Devuelve array vacío hasta encontrar el endpoint correcto

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ ok: false, error: 'Método no permitido' });
  }
  return res.status(200).json({ ok: true, data: [] });
}
