# Seguridad de datos GPS

## Principio

GitHub Pages sirve contenido estático. Un formulario de login dentro de una SPA no convierte en privados los archivos estáticos incluidos en el despliegue.

Por esa razón, FIAS · Inteligencia de Movilizaciones 2.0 mantiene en GitHub únicamente el código del front-end y del workflow. Los archivos con coordenadas, horas, calles, eventos y kilometraje se almacenan en SharePoint y se consultan mediante Microsoft Graph después de autenticar al usuario.

## Controles

- Autenticación con Microsoft Entra ID / MSAL.
- Usuario normal: token de lectura.
- Administrador GPS: `jcruzg@fias.org.ec` y consentimiento incremental `Files.ReadWrite` para la carga a la carpeta autorizada.
- Permisos SharePoint en `GPS/Entrada` como control efectivo de carga.
- GitHub Actions con App Registration independiente.
- GitHub Secrets para credenciales de aplicación.
- `Sites.Selected` para limitar la App de Actions al sitio autorizado.
- SHA-256 y deduplicación de los PDF procesados.
- Auditoría en SharePoint `GPS/Publicados/history.json`.
- Datos GPS excluidos del repositorio por `.gitignore`.
