# Changelog

## v2.1.1

- El panel **Conexión y estado de datos** se trasladó desde Control de datos a **Administración GPS**.
- El rótulo **Fuentes** fue eliminado y sustituido por **Control de datos**.
- El diagnóstico de conexión y la configuración de SharePoint quedan visibles únicamente para `jcruzg@fias.org.ec`.
- Los usuarios no administradores ya no reciben la ventana de configuración de fuente ante un error de sincronización.
- El módulo Control de datos conserva únicamente completitud y excepciones, mejorando su jerarquía visual.

# 2.1.0

- Rediseño del área de filtros: ahora permanece contraída y se abre bajo demanda desde un botón compacto.
- Resumen visible de filtros activos con chips removibles y conteo de resultados.
- Eliminación de Conductor como filtro, columna, KPI, indicador de calidad y campo de exportación para evitar redundancia con el solicitante.
- Simplificación del módulo Uso y demanda y de la tabla de Movilizaciones.
- Ajustes responsive y de modo oscuro para la nueva experiencia de filtros.

# 2.1.0

- Se separó el sitio SharePoint usado por datos GPS del sitio utilizado por la lista operativa de movilizaciones.
- GPS configurado en `https://fiasec.sharepoint.com/sites/FONDODEINVERSIONAMBIENTALSOSTENIBLE`.
- GitHub Actions y el procesador usan el mismo sitio GPS.

# Changelog

## v2.0.0

- Carga directa de PDF GPS desde la plataforma hacia SharePoint.
- Administración GPS restringida a `jcruzg@fias.org.ec`.
- Lectura normal y escritura incremental para el administrador mediante MSAL.
- Automatización con GitHub Actions y Poppler.
- Procesamiento, validación, hash SHA-256 y deduplicación.
- Integración incremental de JSON mensuales.
- Auditoría de cargas, duplicados y rechazos.
- JSON GPS de producción movidos a SharePoint protegido; ya no se publican en GitHub Pages.
- Workflow de Pages separado del workflow de datos GPS.
- Modo claro/oscuro y diseño responsive.
- Logo oficial FIAS.
- Filtro por conductor.
- Comparación de períodos.
- Indicadores de horas de uso, vehículos y conductores activos.
- Exportación CSV, Excel y PDF.
- Panel de cola e historial para administración GPS.
- Parser validado con enero–julio 2026: **35.979 puntos GPS**.

- Reducción del permiso delegado de carga a `Files.ReadWrite` y heartbeat mensual no sensible para mantener activo el cron de GitHub Actions en repositorios públicos.
