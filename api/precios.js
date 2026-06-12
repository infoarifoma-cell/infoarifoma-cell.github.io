// GET /api/precios
// Devuelve precios desde variables de entorno (no expone datos de negocio en frontend)

export default function handler(req, res) {
  if (req.method !== 'GET') {
    return res.status(405).json({ ok: false, error: 'Método no permitido' });
  }

  try {
    const precios = JSON.parse(process.env.PRECIOS_JSON || '{}');
    const preciosEsp = JSON.parse(process.env.PRECIOS_ESP_JSON || '{}');
    const igicPct = Number(process.env.IGIC_PCT || '3');

    return res.status(200).json({ ok: true, precios, preciosEsp, igicPct });
  } catch (e) {
    console.error('precios error:', e.message);
    return res.status(500).json({ ok: false, error: e.message });
  }
}
