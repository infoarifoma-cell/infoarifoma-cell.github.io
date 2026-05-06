// ============================================================
// ARIFOMA · CARGADOR DE VARIABLES DE ENTORNO
// ============================================================
// Carga .env en desarrollo (local)
// En producción (GitHub Pages), las vars vienen de Supabase o environment variables

(async function loadEnv() {
  try {
    // Solo intenta cargar .env en desarrollo (localhost)
    if (window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1') {
      const response = await fetch('.env');
      if (response.ok) {
        const envText = await response.text();
        const lines = envText.split('\n');
        lines.forEach(line => {
          line = line.trim();
          if (line && !line.startsWith('#')) {
            const [key, ...valueParts] = line.split('=');
            const value = valueParts.join('=');
            // Guardar en window global para acceso desde otros scripts
            window[`__${key}__`] = value;
          }
        });
        console.log('✓ Variables de entorno cargadas desde .env');
      }
    }
  } catch (e) {
    // En producción, esto falla (esperado) — las vars vienen de otro lado
    console.debug('Env vars no disponibles en .env (esperado en producción)');
  }
})();

// Exportar función para obtener vars
function getEnvVar(key) {
  return window[`__${key}__`] || null;
}
