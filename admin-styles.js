// ============================================================
// ARIFOMA · PANEL DE ADMINISTRACIÓN DE ESTILOS
// ============================================================

const STYLE_KEY = 'arifoma_styles';

const DEFAULTS = {
  '--accent': '#6b7d2e',
  '--accent2': '#5a6b25',
  '--bg': '#f0f0ec',
  '--surface': '#ffffff',
  '--surface2': '#f5f5f0',
  '--topbar-bg': '#1a1a1a',
  '--text': '#1a1a1a',
  '--muted': '#707070',
  '--danger': '#c0392b',
  '--border': '#c8c8b8',
  '--radius': '12px',
  '--font-base': '19px'
};

const COLOR_VARS = [
  { name: '--accent', label: 'Accent (primario)', color: '#6b7d2e' },
  { name: '--accent2', label: 'Accent 2 (secundario)', color: '#5a6b25' },
  { name: '--bg', label: 'Fondo (background)', color: '#f0f0ec' },
  { name: '--surface', label: 'Surface', color: '#ffffff' },
  { name: '--surface2', label: 'Surface 2', color: '#f5f5f0' },
  { name: '--topbar-bg', label: 'Topbar fondo', color: '#1a1a1a' },
  { name: '--text', label: 'Texto (text)', color: '#1a1a1a' },
  { name: '--muted', label: 'Muted (gris)', color: '#707070' },
  { name: '--danger', label: 'Danger (rojo)', color: '#c0392b' },
  { name: '--border', label: 'Borde (border)', color: '#c8c8b8' }
];

// Cargar estilos guardados y aplicar al página
function initStylePanel() {
  const saved = localStorage.getItem(STYLE_KEY);
  if (saved) {
    try {
      const styles = JSON.parse(saved);
      Object.entries(styles).forEach(([name, value]) => {
        document.documentElement.style.setProperty(name, value);
      });
    } catch (e) {
      console.warn('Error cargando estilos:', e.message);
    }
  }
}

// Aplicar un CSS var al documento y guardar
function applyStyleVar(name, value) {
  document.documentElement.style.setProperty(name, value);
  // Guardar inmediatamente este cambio específico
  const saved = localStorage.getItem(STYLE_KEY);
  const styles = saved ? JSON.parse(saved) : {};
  styles[name] = value;
  localStorage.setItem(STYLE_KEY, JSON.stringify(styles));
}

// Guardar todos los estilos actuales en localStorage
function saveCurrentStyles() {
  const styles = {};
  const root = document.documentElement;

  // Guardar todos los DEFAULTS keys
  Object.keys(DEFAULTS).forEach(name => {
    const val = root.style.getPropertyValue(name).trim();
    if (val) styles[name] = val;
  });

  localStorage.setItem(STYLE_KEY, JSON.stringify(styles));
}

// Obtener valor actual de una var CSS
function getStyleValue(name) {
  const val = document.documentElement.style.getPropertyValue(name).trim();
  return val || DEFAULTS[name] || '';
}

// Abrir panel de ajustes
function openSettingsPanel() {
  const modal = document.getElementById('style-panel-wrap');
  if (!modal) return console.error('style-panel-wrap no encontrado');
  renderStyleInputs();
  switchSettingsTab('estilos');
  setTimeout(() => { modal.classList.add('open'); }, 0);
}

// Compat alias
function openStylePanel() { openSettingsPanel(); }

// Cerrar panel
function closeSettingsPanel() {
  const modal = document.getElementById('style-panel-wrap');
  if (modal) modal.classList.remove('open');
}
function closeStylePanel() { closeSettingsPanel(); }

// Cambiar tab
function switchSettingsTab(tab) {
  const tabs = ['estilos', 'conectores'];
  tabs.forEach(t => {
    const panel = document.getElementById('settings-tab-' + t);
    const btn = document.getElementById('tab-' + t);
    if (panel) panel.style.display = t === tab ? 'block' : 'none';
    if (btn) {
      btn.style.background = t === tab ? 'var(--accent)' : 'var(--surface2)';
      btn.style.color = t === tab ? '#fff' : 'var(--text)';
    }
  });
  if (tab === 'conectores') checkAllConnectors();
}

// ── CONECTORES ──────────────────────────────────────────────
const CONNECTORS = [
  { id: 'supabase', name: 'Supabase (Base de datos)', check: checkSupabase, reconnect: reconnectSupabase },
  { id: 'bc', name: 'Business Central (ERP)', check: checkBC, reconnect: reconnectBC },
  { id: 'google', name: 'Google Sheets', check: checkGoogle },
  { id: 'onedrive', name: 'OneDrive / SharePoint', check: checkOneDrive, reconnect: reconnectOneDrive },
  { id: 'ocr', name: 'OCR (Hugging Face)', check: checkOCR },
];

let _connectorResults = [];

function renderConnectorStatus(results) {
  _connectorResults = results;
  const list = document.getElementById('connectors-list');
  if (!list) return;
  list.innerHTML = results.map((r, i) => {
    const color = r.ok ? '#27ae60' : '#c0392b';
    const statusText = r.ok ? 'Operativo' : (r.error || 'Sin conexión');
    const connector = CONNECTORS[i];
    const reconnectBtn = !r.ok && connector.reconnect
      ? `<button onclick="reconnectSingle(${i})" style="padding:4px 10px;background:var(--accent);color:#fff;border:none;border-radius:6px;font-size:.68rem;font-weight:700;cursor:pointer;white-space:nowrap;flex-shrink:0">Conectar</button>`
      : '';
    return `<div style="display:flex;align-items:center;gap:10px;padding:10px 12px;background:var(--surface2);border-radius:8px;border:1.5px solid ${r.ok ? 'var(--border)' : 'rgba(192,57,43,.3)'}">
      <span style="color:${color};font-size:1.1rem;flex-shrink:0">●</span>
      <div style="flex:1;min-width:0">
        <div style="font-size:.8rem;font-weight:700;color:var(--text)">${r.name}</div>
        <div style="font-size:.7rem;color:${color};font-weight:600">${statusText}</div>
      </div>
      ${reconnectBtn}
    </div>`;
  }).join('');
}

async function reconnectSingle(idx) {
  const c = CONNECTORS[idx];
  if (!c || !c.reconnect) return;
  // Show loading on that item
  const list = document.getElementById('connectors-list');
  const items = list?.children;
  if (items && items[idx]) {
    const btn = items[idx].querySelector('button');
    if (btn) { btn.textContent = '...'; btn.disabled = true; }
  }
  try {
    await c.reconnect();
    // Re-check after reconnect
    const ok = await c.check();
    _connectorResults[idx] = { name: c.name, ok };
  } catch (e) {
    _connectorResults[idx] = { name: c.name, ok: false, error: e.message || 'Error' };
  }
  renderConnectorStatus(_connectorResults);
}

async function checkAllConnectors() {
  const list = document.getElementById('connectors-list');
  if (list) list.innerHTML = '<div style="text-align:center;padding:20px;color:var(--muted);font-size:.85rem">Comprobando conectores...</div>';

  const results = await Promise.all(CONNECTORS.map(async c => {
    try {
      const ok = await c.check();
      return { name: c.name, ok };
    } catch (e) {
      return { name: c.name, ok: false, error: e.message || 'Sin conexión' };
    }
  }));
  renderConnectorStatus(results);
}

// ── Checks ──
async function checkSupabase() {
  if (!_sessionToken) return false;
  const res = await fetch('/api/supabase', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + _sessionToken },
    body: JSON.stringify({ action: 'select', table: 'tblConfig', options: { limit: 1 } })
  });
  return res.ok;
}

async function checkBC() {
  try {
    const res = await fetch('/api/config');
    if (!res.ok) return false;
    const cfg = await res.json();
    return !!(cfg.bc && cfg.bc.tenant && cfg.bc.client && cfg.bc.env);
  } catch { return false; }
}

async function checkGoogle() {
  try {
    const res = await fetch('/api/google-sheet', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ spreadsheetId: '1fxHwVEgcIrRdyPh-TJ-k84QFBHXX-P3mNRCiWYaeDTQ', sheet: 'STOCK', range: 'STOCK!A1:A1' })
    });
    return res.ok;
  } catch { return false; }
}

async function checkOneDrive() {
  try {
    const res = await fetch('/api/config');
    if (!res.ok) return false;
    const cfg = await res.json();
    return !!(cfg.compras && cfg.compras.clientId);
  } catch { return false; }
}

async function checkOCR() {
  try {
    const res = await fetch('/api/config');
    return res.ok; // OCR depends on HF_TOKEN env var, we just check API is reachable
  } catch { return false; }
}

// ── Reconnect actions ──
async function reconnectSupabase() {
  if (typeof _refreshToken === 'function') await _refreshToken();
}

async function reconnectBC() {
  if (typeof getBCToken === 'function') {
    try { await getBCToken(); } catch {}
  }
}

async function reconnectOneDrive() {
  if (typeof comprasGetToken === 'function') {
    try { await comprasGetToken(); } catch {}
  }
}

// Renderizar inputs del panel con valores actuales
function renderStyleInputs() {
  // Colores
  const colorContainer = document.getElementById('style-colors-list');
  if (colorContainer) {
    colorContainer.innerHTML = COLOR_VARS.map(({ name, label }) => {
      const val = getStyleValue(name);
      return `
        <div style="margin-bottom:10px">
          <label style="display:flex;align-items:center;gap:8px;font-size:.75rem;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:var(--muted);margin-bottom:4px">
            ${label}
            <input type="color" value="${val}" onchange="applyStyleVar('${name}', this.value)" style="width:32px;height:32px;cursor:pointer;border:none;border-radius:6px">
          </label>
          <input type="text" value="${val}" onchange="applyStyleVar('${name}', this.value)" style="width:100%;background:var(--surface2);border:1.5px solid var(--border);border-radius:8px;color:var(--text);font-family:'DM Mono',monospace;font-size:.9rem;padding:8px;outline:none" placeholder="#000000">
        </div>
      `;
    }).join('');
  }

  // Font size
  const fontBaseInput = document.getElementById('style-font-base');
  const fontBaseDisplay = document.getElementById('style-font-base-display');
  if (fontBaseInput && fontBaseDisplay) {
    const val = getStyleValue('--font-base');
    const px = parseInt(val);
    fontBaseInput.value = px;
    fontBaseDisplay.textContent = val;

    // Remover handler anterior si existe
    fontBaseInput.removeEventListener('input', fontBaseSizeHandler);
    // Agregar nuevo handler
    fontBaseInput.addEventListener('input', fontBaseSizeHandler);
  }

  // Border radius
  const radiusInput = document.getElementById('style-radius');
  const radiusDisplay = document.getElementById('style-radius-display');
  if (radiusInput && radiusDisplay) {
    const val = getStyleValue('--radius');
    const px = parseInt(val);
    radiusInput.value = px;
    radiusDisplay.textContent = val;

    // Remover handler anterior si existe
    radiusInput.removeEventListener('input', radiusSizeHandler);
    // Agregar nuevo handler
    radiusInput.addEventListener('input', radiusSizeHandler);
  }
}

// Handlers para sliders (funciones nombradas para poder removerlas)
function fontBaseSizeHandler(e) {
  const val = e.target.value + 'px';
  applyStyleVar('--font-base', val);
  document.getElementById('style-font-base-display').textContent = val;
}

function radiusSizeHandler(e) {
  const val = e.target.value + 'px';
  applyStyleVar('--radius', val);
  document.getElementById('style-radius-display').textContent = val;
}

// Resetear a valores por defecto
function resetStyles() {
  if (!confirm('¿Resetear todos los estilos a los valores por defecto?')) return;
  localStorage.removeItem(STYLE_KEY);
  location.reload();
}

// Exportar configuración como JSON
function exportStyles() {
  const saved = localStorage.getItem(STYLE_KEY);
  const data = saved ? JSON.parse(saved) : DEFAULTS;

  const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `arifoma-estilos-${new Date().toISOString().slice(0, 10)}.json`;
  a.click();
  URL.revokeObjectURL(url);
}

// Importar configuración desde JSON
function importStyles(event) {
  const file = event.target.files?.[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = JSON.parse(e.target.result);

      // Validar que sea un objeto de estilos
      if (typeof data !== 'object' || data === null) {
        throw new Error('JSON inválido');
      }

      // Aplicar estilos
      Object.entries(data).forEach(([name, value]) => {
        if (typeof value === 'string') {
          document.documentElement.style.setProperty(name, value);
        }
      });

      // Guardar
      saveCurrentStyles();
      renderStyleInputs();
      alert('Estilos importados correctamente');
    } catch (err) {
      alert('Error importando archivo: ' + err.message);
    }
  };
  reader.readAsText(file);
}
