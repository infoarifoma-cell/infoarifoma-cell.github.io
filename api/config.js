// GET /api/config
// Sirve configuración pública (IDs de app registration, etc.)
// Los valores vienen de env vars de Vercel, no hardcodeados en frontend

export default function handler(req, res) {
  if (req.method !== 'GET') {
    return res.status(405).json({ ok: false, error: 'Método no permitido' });
  }

  return res.status(200).json({
    bc: {
      tenant: process.env.BC_TENANT,
      client: process.env.BC_CLIENT,
      env: process.env.BC_ENV,
      company: process.env.BC_COMPANY,
      scope: process.env.BC_SCOPE || 'https://api.businesscentral.dynamics.com/.default',
    },
    compras: {
      clientId: process.env.COMPRAS_CLIENT_ID,
      tenantId: process.env.COMPRAS_TENANT_ID,
      onedriveBase: process.env.COMPRAS_ONEDRIVE_BASE || '06. ADMINISTRACION/06.01 PROVEEDORES',
      shareUrl: process.env.COMPRAS_SHARE_URL,
    }
  });
}
