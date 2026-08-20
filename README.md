# FIAS · Inteligencia de Movilizaciones

Sistema web estático y modular para combinar la **lista institucional de movilizaciones en SharePoint** con los **reportes de rastreo satelital InformeRecorridoPlus**.

## Qué hace

- Autenticación Microsoft 365 mediante **MSAL**.
- Lectura en tiempo real de la lista de SharePoint mediante Microsoft Graph.
- Descubrimiento de listas y mapeo semántico de columnas.
- Procesamiento de fechas, solicitantes, grupos/proyectos, destinos y kilometrajes.
- Integración de rastreo GPS por fecha/hora.
- Trazas GPS reales, mapas de calor y cobertura territorial.
- Conciliación `RECORRIDO SharePoint` vs. odómetro GPS de referencia.
- Rankings de destinos, actividades, grupos/proyectos y usuarios.
- Detección de excepciones: kilometrajes, fechas, solicitudes, recorridos atípicos y diferencias GPS.
- Exportación de la vista filtrada a CSV.

## Arquitectura

```text
fias-movilizaciones/
├── index.html
├── css/
│   └── styles.css
├── config/
│   └── msal-config.js
├── js/
│   ├── app.js
│   ├── auth.js
│   ├── graph.js
│   ├── sharepoint.js
│   ├── gps.js
│   ├── analytics.js
│   ├── maps.js
│   ├── dashboard.js
│   └── utils.js
├── data/
│   └── gps/
│       ├── manifest.json
│       ├── pdf-8770-2026-01.json
│       ├── ...
│       └── pdf-8770-2026-07.json
├── scripts/
│   └── extract_gps.py
├── README.md
└── DEPLOY.md
```

## Fuentes actuales

### SharePoint

El sitio configurado es:

- Host: `fiasec.sharepoint.com`
- Sitio: `/sites/RecursosAdministrativo`
- La lista se descubre automáticamente por sus columnas; el usuario puede cambiarla desde **Control de datos → Configurar**.

### GPS

Se incorporaron los reportes del rastreador `PDF-8770` de **enero a julio de 2026**. Los PDF se convirtieron a JSON por mes para evitar cargar y reprocesar miles de páginas en cada visita al dashboard.

Datos extraídos por posición:

`fecha/hora · calle/ubicación · evento · latitud · longitud · velocidad · kilometraje`

Total actual: **35.979 puntos GPS**.

## Autenticación

`config/msal-config.js` conserva el patrón MSAL del HTML institucional de referencia:

- tenant institucional;
- client ID institucional;
- `loginPopup`;
- `acquireTokenSilent` con fallback a `acquireTokenPopup`;
- `sessionStorage`;
- redirect URI = URL actual de la aplicación, salvo que se configure una URI fija.

> La URL publicada debe registrarse exactamente en Microsoft Entra ID como **Single-page application (SPA) Redirect URI**.

## Ejecutar en local

Los módulos ES y los JSON deben servirse por HTTP, no abrirse con doble clic como `file://`.

```bash
cd fias-movilizaciones
python -m http.server 8080
```

Abrir:

`http://localhost:8080/`

Para probar autenticación local, esta dirección debe estar registrada en Entra si la política de la aplicación lo exige.

## Incorporar nuevos reportes GPS

Requiere `pdftotext` (Poppler) disponible en PATH.

```bash
python scripts/extract_gps.py \
  "InformeRecorridoPlus - Agosto 2026.pdf" \
  "InformeRecorridoPlus - Septiembre 2026.pdf" \
  --out data/gps
```

El script genera los JSON y actualiza `manifest.json`. Después solo se publican los archivos modificados.

## Relación SharePoint ↔ GPS

La correlación se realiza por la ventana temporal `FECHA INICIA USO` → `FECHA TERMINA USO`, con una tolerancia configurada en `GPS_CONFIG.matchPaddingMinutes`.

Si SharePoint contiene vehículo o placa, se puede configurar en `GPS_CONFIG.trackerAliases` qué unidad corresponde a cada rastreador. Si no existe ese campo, la correlación se considera temporal y debe interpretarse como evidencia de referencia.

El kilometraje GPS **no sustituye** el `RECORRIDO` institucional. Se muestra como contraste técnico porque ambos valores pueden diferir por ventana temporal, odómetro, puntos del dispositivo o forma de registro.


## Fuente SharePoint configurada

La fuente principal corresponde a la lista compartida del sitio `RecursosAdministrativo`:

`https://fiasec.sharepoint.com/:l:/s/RecursosAdministrativo/JABGnrxRvgAqSaT9aEcQZvFfAbrh1fNSrNx-i-ixKrm_kZw?e=Ryl4oq`

La aplicación identifica automáticamente esta lista por la firma de columnas esperada: FECHA INICIA USO, FECHA TERMINA, Usuario1, GRUPO, FECHA SOLICITUD, DESTINO, KM INICIAL, KM FINAL y RECORRIDO.

## Ajustes v1.2

- Los campos SharePoint de tipo **Persona** o **Lookup** ya no se muestran como identificadores numéricos cuando existe información de referencia. La aplicación intenta resolver `Usuario1LookupId` y otros lookups contra la lista interna de usuarios o la lista de referencia correspondiente.
- Se reconoce también el nombre interno habitual `UserInfo` de la lista **User Information List** de SharePoint.
- El mapa Leaflet quedó aislado en su propio contexto de apilamiento y el encabezado tiene prioridad visual, evitando que controles, popups o capas del mapa se superpongan al header al hacer zoom o desplazarse.

## Corrección v1.3 – campos Persona de SharePoint

La versión 1.3 resuelve los valores numéricos de columnas Persona/Grupo (por ejemplo `Usuario1 = 50`) contra la lista oculta **User Information List** de SharePoint. La aplicación intenta obtener esa lista directamente por título mediante Microsoft Graph y, si no puede precargarla completa, consulta bajo demanda únicamente los IDs presentes en los registros de movilización. Esto evita mostrar el `LookupId` como nombre del usuario.
