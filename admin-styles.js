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
  saveCurrentStyles();
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

// Abrir modal del panel
function openStylePanel() {
  const modal = document.getElementById('style-panel-wrap');
  if (!modal) return console.error('style-panel-wrap no encontrado');

  // Renderizar valores actuales
  renderStyleInputs();
  modal.classList.add('open');
}

// Cerrar modal
function closeStylePanel() {
  const modal = document.getElementById('style-panel-wrap');
  if (modal) modal.classList.remove('open');
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
  if (fontBaseInput) {
    const val = getStyleValue('--font-base');
    const px = parseInt(val);
    fontBaseInput.value = px;
    fontBaseInput.oninput = function() {
      applyStyleVar('--font-base', this.value + 'px');
      document.getElementById('style-font-base-display').textContent = this.value + 'px';
    };
    document.getElementById('style-font-base-display').textContent = val;
  }

  // Border radius
  const radiusInput = document.getElementById('style-radius');
  if (radiusInput) {
    const val = getStyleValue('--radius');
    const px = parseInt(val);
    radiusInput.value = px;
    radiusInput.oninput = function() {
      applyStyleVar('--radius', this.value + 'px');
      document.getElementById('style-radius-display').textContent = this.value + 'px';
    };
    document.getElementById('style-radius-display').textContent = val;
  }
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
