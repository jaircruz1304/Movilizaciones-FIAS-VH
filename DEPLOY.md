# Despliegue de FIAS · Inteligencia de Movilizaciones 2.0

Objetivo: operar únicamente con **GitHub Pages + GitHub Actions + SharePoint/Microsoft Graph**.

## 1. Publicar el front-end con GitHub Actions

En GitHub:

`Settings → Pages → Build and deployment → Source → GitHub Actions`

El workflow `.github/workflows/pages.yml` publica únicamente:

```text
index.html
css/
config/
js/
assets/
```

No publica `data/gps`, porque las coordenadas permanecen en SharePoint.

## 2. Redirect URI de la SPA

En Microsoft Entra ID, en la App Registration usada por el navegador:

`Authentication → Single-page application`

Registrar la URL exacta de Pages, por ejemplo:

```text
https://jaircruz1304.github.io/Movilizaciones-FIAS-VH/
```

## 3. Permisos delegados de la SPA

La versión 2.0 utiliza:

```text
User.Read
Sites.Read.All
Files.ReadWrite
```

El código solicita `User.Read + Sites.Read.All` al inicio. `Files.ReadWrite` se solicita incrementalmente únicamente cuando el administrador autorizado intenta cargar un PDF. Es el permiso delegado de menor privilegio documentado por Microsoft Graph para cargar o reemplazar archivos en una unidad a la que el usuario ya tiene acceso.

La autorización de escritura no debe depender solo del correo comprobado en JavaScript. Debe reforzarse con permisos SharePoint.

## 4. Carpetas SharePoint

Sitio configurado:

```text
https://fiasec.sharepoint.com/sites/FONDODEINVERSIONAMBIENTALSOSTENIBLE
```

Rutas:

```text
Movilizaciones-FIAS/GPS/Entrada
Movilizaciones-FIAS/GPS/Procesados
Movilizaciones-FIAS/GPS/Rechazados
Movilizaciones-FIAS/GPS/Publicados
```

### Permisos recomendados

**Entrada**
- `jcruzg@fias.org.ec`: Editar.
- Usuarios normales: sin escritura.

**Publicados**
- Usuarios autorizados del dashboard: Lectura.
- `jcruzg@fias.org.ec`: Lectura o Edición según política interna.
- App de GitHub Actions: acceso de aplicación `write` al sitio mediante `Sites.Selected`.

**Procesados / Rechazados**
- Administradores: según política institucional.
- Usuarios normales: no necesitan acceso.

Si la biblioteca concede edición amplia por herencia, establecer permisos únicos en las carpetas sensibles.

## 5. App Registration separada para GitHub Actions

Crear una segunda aplicación, por ejemplo:

```text
FIAS GPS GitHub Actions
```

Microsoft Graph → **Application permissions**:

```text
Sites.Selected
```

Conceder Admin Consent.

Esta aplicación no se referencia en JavaScript y su secreto nunca llega al navegador.

## 6. Conceder `write` al sitio seleccionado

`Sites.Selected` requiere una concesión explícita sobre el sitio. Un administrador puede usar Microsoft Graph:

```http
POST https://graph.microsoft.com/v1.0/sites/{SITE_ID}/permissions
Content-Type: application/json

{
  "roles": ["write"],
  "grantedToIdentities": [
    {
      "application": {
        "id": "CLIENT_ID_DE_LA_APP_ACTIONS",
        "displayName": "FIAS GPS GitHub Actions"
      }
    }
  ]
}
```

## 7. GitHub Secrets

En:

`Settings → Secrets and variables → Actions`

crear:

```text
MS_TENANT_ID
MS_ACTION_CLIENT_ID
MS_ACTION_CLIENT_SECRET
```

- `MS_TENANT_ID`: tenant institucional.
- `MS_ACTION_CLIENT_ID`: App Registration exclusiva de Actions.
- `MS_ACTION_CLIENT_SECRET`: valor del secreto de cliente.

No colocar estos valores dentro de `config/msal-config.js`.

## 8. Workflow de sincronización GPS

`.github/workflows/sync-gps.yml`:

- se ejecuta aproximadamente cada cinco minutos;
- puede ejecutarse manualmente con `workflow_dispatch`;
- instala Poppler;
- lee `GPS/Entrada`;
- descarga los JSON protegidos existentes de `GPS/Publicados`;
- procesa incrementalmente los nuevos PDF;
- publica únicamente los JSON protegidos actualizados en `GPS/Publicados`;
- mueve los PDF a `Procesados` o `Rechazados`.

No hace commits de coordenadas al repositorio y no necesita volver a desplegar Pages cuando cambia un reporte GPS.

## 9. Primera puesta en producción

1. Publicar el código 2.0.
2. Verificar que GitHub Pages abra correctamente.
3. Iniciar con un usuario normal: no debe aparecer **Administración GPS**.
4. Iniciar con `jcruzg@fias.org.ec`: debe aparecer el módulo.
5. Cargar los siete PDF históricos de enero–julio de 2026 desde el módulo Administración GPS.
6. Ir a `Actions → Sincronizar GPS desde SharePoint → Run workflow` para la primera carga, sin esperar al cron.
7. Confirmar que `GPS/Publicados` contenga `manifest.json`, `history.json` y los JSON mensuales.
8. Pulsar **Actualizar** en la plataforma.
9. Confirmar un total histórico de **35.979 puntos GPS**.

Posteriormente solo se carga cada nuevo PDF; la base se integra de forma incremental.

## 10. Estados de auditoría

`history.json` puede registrar:

- `processed`: integrado correctamente;
- `duplicate`: el mismo SHA-256 ya se había procesado;
- `rejected`: error de validación o extracción;
- `baseline`: reservado para migraciones controladas.

## 11. Tiempo de actualización

GitHub Pages es estático y no puede ejecutar un backend seguro. Por eso la actualización es asíncrona:

```text
PDF → SharePoint → siguiente ejecución de Actions → JSON en SharePoint → Actualizar dashboard
```

GitHub permite schedules con un intervalo mínimo de 5 minutos y puede retrasar ejecuciones programadas durante períodos de alta carga.

## 12. Tamaño máximo

La interfaz admite hasta 200 MB por PDF. Microsoft Graph permite la carga directa de contenido de `driveItem` hasta 250 MB; para archivos mayores sería necesaria una upload session.

## 13. Mantener activo el cron en un repositorio público

GitHub deshabilita automáticamente los workflows programados de repositorios públicos cuando no existe actividad en el repositorio durante 60 días. Por ello se incluye `.github/workflows/keepalive.yml`, que una vez al mes actualiza un archivo **sin datos operativos** (`ops/actions-keepalive.txt`).

Para permitir ese heartbeat:

`Settings → Actions → General → Workflow permissions → Read and write permissions`

El workflow de Pages ignora `ops/**`, por lo que el heartbeat no vuelve a publicar el sitio ni expone información GPS. Si una regla de protección de rama impide el `git push` del workflow, debe autorizarse expresamente esa automatización o mantenerse actividad manual con una frecuencia menor a 60 días.

## 14. Importante si el repositorio actual es público

Las versiones anteriores almacenaban `data/gps/*.json` dentro del repositorio. Borrarlos en 2.0 evita que sigan formando parte del sitio actual, pero **no elimina automáticamente copias existentes del historial Git**.

Si las coordenadas deben considerarse información restringida, se recomienda una de estas opciones antes de continuar producción:

- migrar el código 2.0 a un repositorio nuevo sin historial de los JSON GPS; o
- aplicar un procedimiento aprobado de limpieza de historial Git y revisar forks/cachés existentes.

La versión 2.0 ya está preparada para que las futuras coordenadas permanezcan únicamente en SharePoint.
