export const AUTH_CONFIG = {
  tenantId: '5e23e4af-237d-4d97-bf6e-dca808015787',
  clientId: 'c8a828c1-3a20-4876-96af-b4f30ce4abeb',
  // Permisos solicitados a todos los usuarios: solo lectura de SharePoint.
  scopes: ['User.Read','Sites.Read.All'],
  // Se solicita de forma incremental únicamente al administrador al cargar PDF.
  adminScopes: ['User.Read','Sites.Read.All','Files.ReadWrite'],
  // null = usar exactamente la URL publicada. Debe registrarse como SPA Redirect URI en Entra ID.
  redirectUri: null,
  cacheLocation: 'sessionStorage'
};

export const ACCESS_CONFIG = {
  gpsAdministrators: ['jcruzg@fias.org.ec']
};

export const SHAREPOINT_CONFIG = {
  host: 'fiasec.sharepoint.com',
  sitePath: '/sites/RecursosAdministrativo',
  listShareUrl: 'https://fiasec.sharepoint.com/:l:/s/RecursosAdministrativo/JABGnrxRvgAqSaT9aEcQZvFfAbrh1fNSrNx-i-ixKrm_kZw?e=Ryl4oq',
  preferredListId: '',
  expectedColumns: [
    'FECHA INICIA USO','FECHA TERMINA','Usuario1','GRUPO','FECHA SOLICITUD',
    'DESTINO','KM INICIAL','KM FINAL','RECORRIDO'
  ],
  lockToBestMatch: true,
  maxItems: 15000
};

// Carpetas relativas a la biblioteca de documentos predeterminada del sitio SharePoint.
// La carpeta Entrada debe tener permisos de escritura únicamente para el administrador autorizado.
export const SHAREPOINT_GPS_CONFIG = {
  // Sitio real donde se almacena la carpeta Movilizaciones-FIAS/GPS.
  host: 'fiasec.sharepoint.com',
  sitePath: '/sites/FONDODEINVERSIONAMBIENTALSOSTENIBLE',
  inboxFolder: 'Movilizaciones-FIAS/GPS/Entrada',
  processedFolder: 'Movilizaciones-FIAS/GPS/Procesados',
  rejectedFolder: 'Movilizaciones-FIAS/GPS/Rechazados',
  // JSON protegidos consumidos por el dashboard autenticado. Nunca se publican en GitHub Pages.
  publishedFolder: 'Movilizaciones-FIAS/GPS/Publicados',
  manifestName: 'manifest.json',
  historyName: 'history.json',
  maxUploadMb: 200,
  acceptedMimeTypes: ['application/pdf'],
  actionIntervalMinutes: 5
};

export const GPS_CONFIG = {
  origin: {
    name: 'Matriz FIAS · Quito',
    lat: -0.20465,
    lon: -78.48410,
    note: 'Punto institucional de referencia para análisis de salida y cobertura.'
  },
  trackerAliases: {
    'PDF-8770': ['PDF-8770']
  },
  matchPaddingMinutes: 15
};

export const APP_CONFIG = {
  name: 'FIAS · Inteligencia de Movilizaciones',
  timezone: 'America/Guayaquil',
  locale: 'es-EC',
  version: '2.1.1',
  logoUrl: 'https://fias.org.ec/wp-content/uploads/2021/11/Logo_FIAS_web.png'
};
