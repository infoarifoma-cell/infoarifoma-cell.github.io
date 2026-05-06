# GitHub Secrets Setup — ARIFOMA

Para que GitHub Actions inyecte las credenciales en producción, necesitas crear **Repository Secrets**.

## 🔐 Pasos:

### 1. Ir a GitHub Settings

URL: https://github.com/infoarifoma-cell/infoarifoma-cell.github.io/settings/secrets/actions

O manualmente:
- GitHub → Tu repo → Settings → Secrets and variables → Actions → New repository secret

### 2. Agregar cada secret (7 total)

| Name | Value |
|------|-------|
| `SUPABASE_URL` | `https://bnsfgzjqmibsrklllqxb.supabase.co` |
| `SUPABASE_KEY` | `eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImJuc2ZnempxbWlic3JrbGxzcXhiIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzQzNzYwNzksImV4cCI6MjA4OTk1MjA3OX0.8mTQHPdO954ICBd1Xam-kKmcA69CMyO2v3x1liFgWyk` |
| `BC_TENANT` | `5bd828f2-1899-48ba-a269-c37733f41806` |
| `BC_CLIENT` | `e2a57ff0-8ea7-433d-a2af-7335d3f01847` |
| `BC_SECRET` | `<tu-secret-actual-de-azure>` |
| `BC_ENV` | `Production` |
| `BC_COMPANY` | `ARIFOMA 25P.V06` |
| `SHEETS_API` | `https://script.google.com/macros/s/AKfycbwPIIgZCg03i4aJN8HIxKf20P5IPc-j3HOkoHmt2Jx0-vqiWrmq4Gz2WZmZvyopYJlv/exec` |

### 3. Para cada secret:

**Click "New repository secret"**

```
Name: SUPABASE_URL
Value: https://bnsfgzjqmibsrklllqxb.supabase.co
```

Luego "Add secret"

Repetir 8 veces (todas las vars).

---

## ✅ Verificación

Después de agregar todos los secrets:

1. Push un cambio a `main` (o trigger manual):
   ```bash
   git commit --allow-empty -m "trigger CI"
   git push
   ```

2. Ver workflow: https://github.com/infoarifoma-cell/infoarifoma-cell.github.io/actions
   - Si sale verde ✓ → secrets ok
   - Si sale rojo ✗ → revisar logs

3. Chequear que `_secrets.js` fue generado:
   - Ir a repo → `_secrets.js`
   - Debe tener valores, no vacío

---

## 🛠️ Local Development

En tu máquina local:

```bash
# .env (NO committed a Git — .gitignore lo ignora)
SUPABASE_URL=https://bnsfgzjqmibsrklllqxb.supabase.co
SUPABASE_KEY=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9...
BC_TENANT=5bd828f2-1899-48ba-a269-c37733f41806
BC_CLIENT=e2a57ff0-8ea7-433d-a2af-7335d3f01847
BC_SECRET=<tu-secret>
BC_ENV=Production
BC_COMPANY=ARIFOMA 25P.V06
SHEETS_API=https://script.google.com/macros/s/...
```

`_env-loader.js` cargará `.env` automáticamente al ejecutar localmente.

---

## ⚠️ IMPORTANTE

**NO subir `.env` a Git** — ya está en `.gitignore`

Si accidentalmente lo subes:
```bash
git rm --cached .env
git commit -m "Remove .env from git"
```

---

## 📝 Resumen flujo

```
Desarrollo (local)
├── .env (credenciales locales)
├── _env-loader.js (carga .env)
└── getEnvVar() (lee credenciales)

Producción (GitHub Pages)
├── GitHub Secrets (credenciales seguras)
├── GitHub Actions (inyecta en _secrets.js)
├── _secrets.js (auto-generado con valores)
└── getEnvVar() (lee credenciales)
```

---

Done. Adelante.
