# Changelog

## v2.3.0

- Se corrige la interpretación territorial del GPS para que los nombres de calles no se confundan con provincias. En particular, `Av. Francisco de Orellana`, `Camino de Orellana` y `Fernando Sánchez de Orellana` ya no generan falsamente la provincia Orellana.
- La provincia GPS se obtiene del componente territorial final del reporte (`..., ciudad, provincia`) y no por coincidencia de palabras en toda la dirección.
- Se incorporan reparaciones para provincias partidas por el PDF, como `PICHI NCHA`, `TUNGURAHU A` o `PASTA ZA`.
- La ruta GPS se presenta en orden territorial de recorrido (por ejemplo, `Pichincha → Cotopaxi → Tungurahua`) en lugar de ordenar provincias por frecuencia.
- El destino consolidado prioriza el destino registrado en SharePoint cuando la traza GPS alcanza la provincia esperada.
- Se añade validación directa del destino registrado contra los puntos GPS. Por ejemplo, `Tababela` se valida con puntos cuyo texto GPS contiene `VIA COLLAS TABABELA`, aunque el proveedor clasifique territorialmente el punto como `QUITO,PICHINCHA`.
- Cuando el destino registrado no contiene una ubicación, se infiere una ciudad/provincia a partir de evidencia GPS estructurada y repetida, evitando decidir por un único punto aislado.
- Si SharePoint y GPS apuntan a territorios incompatibles, el sistema conserva el destino declarado y lo marca como `Revisar destino` en lugar de reemplazarlo silenciosamente.
- La vista de mapa utiliza, cuando existe, la referencia GPS coincidente con el destino registrado y diferencia ese punto del punto más alejado de la Matriz FIAS.
- Se amplía el catálogo de destinos de Pichincha con Tababela, Tumbaco, Cumbayá, Puembo, Pifo, Yaruquí, El Quinche, Guayllabamba, Calderón, Conocoto y Amaguaña.
- El parser fue validado nuevamente con los ocho reportes GPS de enero a agosto de 2026: **44.185 puntos**.

## v2.2.0

- Se elimina del Panorama el comparador manual de dos períodos.
- Se reemplaza “Actividades predominantes” por “Alcance territorial de las movilizaciones”.
- Se incorpora consolidación geográfica de destinos usando el texto registrado y evidencia GPS.
- Quito se trata como origen institucional cuando en el mismo texto consta otro destino.
- El ranking de destinos muestra ubicación consolidada, provincia y número de registros con evidencia GPS.
- Los registros no reconocidos se agrupan como “Por identificar”.
- El detalle y las exportaciones incorporan destino consolidado y provincia de destino.

## v2.1.1

- El panel **Conexión y estado de datos** se trasladó desde Control de datos a **Administración GPS**.
- El rótulo **Fuentes** fue eliminado y sustituido por **Control de datos**.
- El diagnóstico de conexión y la configuración de SharePoint quedan visibles únicamente para `jcruzg@fias.org.ec`.
- Los usuarios no administradores ya no reciben la ventana de configuración de fuente ante un error de sincronización.

## v2.1.0

- Rediseño del área de filtros: permanece contraída y se abre bajo demanda.
- Resumen de filtros activos con chips removibles y conteo de resultados.
- Eliminación de Conductor como filtro, columna, KPI, indicador de calidad y campo de exportación.
- Simplificación del módulo Uso y demanda y de la tabla de Movilizaciones.
- Ajustes responsive y de modo oscuro.

## v2.0.0

- Carga directa de PDF GPS desde la plataforma hacia SharePoint.
- Administración GPS restringida a `jcruzg@fias.org.ec`.
- Lectura normal y escritura incremental para el administrador mediante MSAL.
- Automatización con GitHub Actions y Poppler.
- Procesamiento, validación, hash SHA-256 y deduplicación.
- Integración incremental de JSON mensuales.
- Auditoría de cargas, duplicados y rechazos.
- JSON GPS de producción almacenados en SharePoint protegido.
- Exportación CSV, Excel y PDF.
