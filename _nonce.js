// ============================================================
// ARIFOMA · NONCE GENERATOR (for CSP)
// ============================================================
// Genera nonce aleatorio para style-src CSP
// En GitHub Pages (sin servidor), usar nonce simple rotado manualmente

// Generar nonce base (versión diaria)
const nonceBase = Math.floor(Date.now() / (24 * 60 * 60 * 1000)); // Cambia cada 24h
window.__CSP_NONCE__ = 'arifoma-' + nonceBase;

// En producción: actualizar manualmente en:
// 1. _nonce.js: window.__CSP_NONCE__
// 2. index.html: <style nonce="..."> y CSP meta
// Frecuencia: cada mes o cuando se despliegue cambios

console.log('CSP Nonce:', window.__CSP_NONCE__);
