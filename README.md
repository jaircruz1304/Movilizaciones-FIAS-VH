# FIAS · Inteligencia de Movilizaciones 2.0

Plataforma para integrar **movilizaciones institucionales de SharePoint** con reportes satelitales **InformeRecorridoPlus**, operando exclusivamente con **GitHub Pages + GitHub Actions + Microsoft 365/SharePoint**, sin servidor propio.

## Arquitectura 2.0

```text
Usuario autorizado
      │
      ▼
GitHub Pages (SPA estática)
      │ MSAL + Microsoft Graph
      ├──────────────► Lista SharePoint de movilizaciones
      │
Administrador GPS: jcruzg@fias.org.ec
      │ carga PDF
      ▼
SharePoint / Movilizaciones-FIAS/GPS/Entrada
      │
      │ GitHub Actions ~cada 5 min
      ▼
Python + Poppler
  extraer → validar → deduplicar → integrar
      │
      ▼
SharePoint / Movilizaciones-FIAS/GPS/Publicados
  manifest.json
  pdf-8770-YYYY-MM.json
  history.json
      │
      ▼
Dashboard autenticado lee los JSON mediante Microsoft Graph
```

### Decisión de seguridad importante

Los JSON GPS contienen coordenadas, horas, calles, eventos y kilometraje. Desde la versión 2.0 **no se publican dentro de GitHub Pages ni se versionan en GitHub**. El navegador autenticado los obtiene desde SharePoint mediante Microsoft Graph.

Esto es especialmente importante si el repositorio de GitHub es público: ocultar el dashboard detrás de un botón de login no protege archivos estáticos que hayan sido publicados como parte del sitio.

## Funcionalidades incorporadas

### Carga y procesamiento automático de GPS

- Carga de uno o varios PDF `InformeRecorridoPlus` desde la interfaz.
- Validación de tipo y tamaño antes de enviar.
- Almacenamiento inicial en SharePoint `GPS/Entrada`.
- Procesamiento automático en GitHub Actions con `pdftotext -layout` de Poppler.
- Extracción de fecha/hora, ubicación, evento, latitud, longitud, velocidad y kilometraje.
- Validación de estructura y coordenadas.
- Hash SHA-256 por PDF.
- Deduplicación por archivo y por punto GPS.
- Integración incremental por mes.
- Regeneración automática de `manifest.json`.
- Traslado del PDF a `Procesados` o `Rechazados`.
- Historial de auditoría protegido en SharePoint.

### Control de acceso

- Usuario administrador GPS: `jcruzg@fias.org.ec`.
- Usuarios normales: consulta, filtros, mapas, análisis y exportación.
- Usuarios normales solicitan a Microsoft Graph permisos de **lectura**.
- El permiso delegado `Files.ReadWrite` se solicita de forma incremental solo al administrador cuando usa la carga GPS; los demás usuarios permanecen con permisos de lectura.
- La carpeta `GPS/Entrada` debe tener permisos SharePoint de escritura únicamente para el administrador autorizado.
- GitHub Actions usa una App Registration separada con credenciales almacenadas solo en GitHub Secrets y se recomienda `Sites.Selected` con acceso `write` únicamente al sitio requerido.

### Reingeniería de interfaz

- Diseño responsive para escritorio, tablet y móvil.
- Modo claro/oscuro persistente.
- Logo oficial FIAS.
- Indicadores ejecutivos.
- Gráficos interactivos.
- Mapas de cobertura, rutas GPS, calor e histórico.
- Filtros bajo demanda por fecha, grupo/proyecto, usuario, vehículo y actividad.
- Comparación de períodos.
- Indicadores de kilometraje y horas de uso.
- Conteo de vehículos activos y métricas de utilización.
- Exportación CSV, Excel y PDF.
- Panel exclusivo de Administración GPS con cola e historial.

## Validación con los siete reportes suministrados

El parser 2.0 fue ejecutado de nuevo sobre los PDF de enero a julio de 2026:

| Mes | Puntos extraídos |
|---|---:|
| Enero | 2.601 |
| Febrero | 1.142 |
| Marzo | 3.998 |
| Abril | 9.384 |
| Mayo | 7.361 |
| Junio | 3.089 |
| Julio | 8.404 |
| **Total** | **35.979** |

Los siete archivos fueron procesados sin pérdida de registros respecto de la base histórica que contenía la versión anterior.

## Carpetas SharePoint

```text
Movilizaciones-FIAS/
└── GPS/
    ├── Entrada/
    ├── Procesados/
    ├── Rechazados/
    └── Publicados/
        ├── manifest.json
        ├── history.json
        └── pdf-8770-YYYY-MM.json
```

`Publicados` debe ser legible por los usuarios autorizados del dashboard, pero no necesita permisos de escritura para ellos. La App de GitHub Actions sí necesita escritura en el sitio.

## Estructura del repositorio

```text
.
├── .github/workflows/
│   ├── pages.yml
│   ├── sync-gps.yml
│   └── keepalive.yml
├── config/msal-config.js
├── css/styles.css
├── js/
│   ├── admin.js
│   ├── analytics.js
│   ├── app.js
│   ├── auth.js
│   ├── dashboard.js
│   ├── gps.js
│   ├── graph.js
│   ├── maps.js
│   ├── sharepoint.js
│   └── utils.js
├── scripts/
│   ├── extract_gps.py
│   ├── sync_sharepoint_gps.py
│   └── validate_gps_data.py
├── data/gps/README.md
├── index.html
└── DEPLOY.md
```

Los datos GPS productivos no forman parte del sitio estático. En repositorios públicos, `keepalive.yml` mantiene actividad mensual no sensible para evitar la desactivación automática de los workflows programados por inactividad.

## Prueba local del parser

Requiere Poppler (`pdftotext`).

```bash
mkdir -p /tmp/fias-gps-test
python scripts/extract_gps.py "InformeRecorridoPlus.pdf" --out /tmp/fias-gps-test
python scripts/validate_gps_data.py --data /tmp/fias-gps-test
```

## Operación normal en producción

No se ejecuta manualmente ningún conversor:

```text
Administrador carga PDF
→ SharePoint Entrada
→ GitHub Actions
→ JSON protegidos en SharePoint Publicados
→ usuario pulsa Actualizar o vuelve a abrir el dashboard
```

La configuración de despliegue y permisos está detallada en [DEPLOY.md](DEPLOY.md).
