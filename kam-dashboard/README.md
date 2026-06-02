# KAM BC — Dashboard de Cartera

Dashboard interactivo para análisis de cartera de clientes. Se actualiza automáticamente cada vez que subes un Excel nuevo.

---

## ¿Cómo funciona?

```
Tu Excel (.xlsx)  →  procesar_datos.py  →  data.json  →  Dashboard web
```

1. Subes un Excel nuevo a la carpeta `data/`
2. GitHub lo detecta y ejecuta el script automáticamente
3. El dashboard se actualiza en ~2 minutos

---

## Instalación inicial (una sola vez)

### 1. Fork o clona este repositorio en tu GitHub

### 2. Activa GitHub Pages
1. Ve a tu repo → **Settings** → **Pages**
2. En "Source" selecciona **Deploy from a branch**
3. Branch: `main` · Folder: `/docs`
4. Guarda. En ~1 minuto tendrás una URL pública como:
   `https://TU_USUARIO.github.io/kam-dashboard/`

### 3. Da permisos al workflow
1. Ve a **Settings** → **Actions** → **General**
2. En "Workflow permissions" selecciona **Read and write permissions**
3. Guarda.

---

## Uso diario

### Actualizar el dashboard con nuevos datos

**Opción A — Desde el navegador (más fácil):**
1. Ve a tu repo en GitHub
2. Entra a la carpeta `data/`
3. Haz clic en "Add file" → "Upload files"
4. Sube tu nuevo Excel (reemplaza el anterior)
5. Haz clic en "Commit changes"
6. En ~2 minutos el dashboard se actualiza solo ✅

**Opción B — Con Git (si lo usas):**
```bash
# Copia tu Excel a la carpeta data/
cp /ruta/a/LETRITAS_JUNIO_2026.xlsx data/

# Sube a GitHub
git add data/
git commit -m "Datos junio 2026"
git push
```

### Ver el resultado
Abre la URL de GitHub Pages de tu repositorio.

---

## Estructura del proyecto

```
kam-dashboard/
├── docs/
│   ├── index.html      ← El dashboard (no tocar)
│   └── data.json       ← Se genera automáticamente
├── data/
│   └── *.xlsx          ← Sube aquí tus archivos Excel
├── .github/
│   └── workflows/
│       └── actualizar.yml  ← El robot que procesa todo
├── procesar_datos.py   ← Script de procesamiento
└── README.md
```

---

## Formato esperado del Excel

El script detecta automáticamente el formato. Solo necesita:
- Hoja llamada **CUADRATURAS**
- Columna **Kam** (o "Codigo Vendedor") con tu código **BC**
- Columna **Precio** (o "Prrecio") con los montos
- Columna **Cliente** con el nombre del cliente
- Columna **Fecha Emision** con la fecha

Puedes subir **varios archivos Excel** (uno por mes, o uno acumulado). El script los combina automáticamente.

---

## Ejecutar localmente (opcional)

Si quieres probar sin subir a GitHub:

```bash
# Instalar dependencias
pip install pandas openpyxl

# Procesar un archivo
python procesar_datos.py data/LETRITAS_2025_2026.xlsx

# Abrir el dashboard
# Abre docs/index.html en tu navegador
# (necesitas un servidor local porque carga data.json por fetch)
# Opción simple:
python -m http.server 8000 --directory docs
# Luego abre: http://localhost:8000
```

---

## Preguntas frecuentes

**¿Puedo subir múltiples archivos Excel?**
Sí. Por ejemplo el histórico + el mes nuevo. El script los combina y si hay datos solapados, usa el más reciente.

**¿El dashboard es privado o público?**
GitHub Pages es público por defecto. Si quieres que sea privado necesitas GitHub Pro/Teams. Alternativamente puedes poner una contraseña básica en el HTML.

**¿Funciona en móvil?**
Sí, el dashboard es responsive.

**¿Quién puede ver el dashboard?**
Cualquier persona con la URL. Compártela con tu equipo directamente.
