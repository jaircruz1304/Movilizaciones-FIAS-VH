# Despliegue en GitHub Pages

## 1. Crear un repositorio nuevo

Ejemplo: `FIAS-Movilizaciones`.

Subir **el contenido de esta carpeta** a la raíz del repositorio.

## 2. Activar GitHub Pages

En GitHub:

`Settings → Pages → Deploy from a branch → main / root`

La URL resultante tendrá una forma similar a:

`https://<usuario>.github.io/FIAS-Movilizaciones/`

## 3. Registrar Redirect URI en Microsoft Entra ID

La aplicación usa por defecto:

```js
window.location.origin + window.location.pathname
```

Por tanto, la URL que realmente abre el usuario debe constar exactamente en:

`Microsoft Entra ID → App registrations → aplicación → Authentication → Single-page application (SPA)`

Ejemplo:

`https://<usuario>.github.io/FIAS-Movilizaciones/`

No registrar una ruta distinta como `/index.html` si los usuarios abrirán la URL terminada en `/`, salvo que ambas se vayan a utilizar. Si se desean admitir ambas, registrar ambas direcciones.

## 4. Configuración del sitio SharePoint

Editar solo si cambia la fuente:

`config/msal-config.js`

```js
export const SHAREPOINT_CONFIG = {
  host: 'fiasec.sharepoint.com',
  sitePath: '/sites/RecursosAdministrativo',
  ...
};
```

No es necesario modificar los módulos de autenticación o analítica para cambiar la lista. La aplicación descubre las listas accesibles y permite seleccionar la correcta desde la interfaz.

## 5. Actualizaciones

Para nuevas versiones se reemplazan los archivos del repositorio **sin cambiar la URL pública**. De esta forma la Redirect URI permanece estable.
