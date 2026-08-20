export const AUTH_CONFIG = {
  tenantId: '5e23e4af-237d-4d97-bf6e-dca808015787',
  clientId: 'c8a828c1-3a20-4876-96af-b4f30ce4abeb',
  scopes: ['User.Read','Files.ReadWrite.All','Sites.ReadWrite.All'],
  // null = usar exactamente la URL de la página publicada, igual al HTML de referencia.
  // En Microsoft Entra ID debe registrarse esa misma URL como SPA Redirect URI.
  redirectUri: null,
  cacheLocation: 'sessionStorage'
};

export const SHAREPOINT_CONFIG = {
  host: 'fiasec.sharepoint.com',
  sitePath: '/sites/RecursosAdministrativo',
  listShareUrl: 'https://fiasec.sharepoint.com/:l:/s/RecursosAdministrativo/JABGnrxRvgAqSaT9aEcQZvFfAbrh1fNSrNx-i-ixKrm_kZw?e=Ryl4oq',
  preferredListId: '',
  // Firma esperada de la lista de movilizaciones. Se usa para localizarla automáticamente
  // dentro del sitio RecursosAdministrativo aunque el enlace compartido no exponga el GUID.
  expectedColumns: [
    'FECHA INICIA USO','FECHA TERMINA','Usuario1','GRUPO','FECHA SOLICITUD',
    'DESTINO','KM INICIAL','KM FINAL','RECORRIDO'
  ],
  lockToBestMatch: true,
  maxItems: 15000
};

export const GPS_CONFIG = {
  manifestUrl: './data/gps/manifest.json',
  origin: {
    name: 'Matriz FIAS · Quito',
    lat: -0.20465,
    lon: -78.48410,
    note: 'Punto institucional de referencia para análisis de salida y cobertura.'
  },
  // Si SharePoint contiene un campo Vehículo/Placa, agregue aquí aliases que
  // correspondan al rastreador PDF-8770. Si el campo no existe, se correlaciona por tiempo.
  trackerAliases: {
    'PDF-8770': ['PDF-8770']
  },
  matchPaddingMinutes: 15
};

export const APP_CONFIG = {
  name: 'FIAS · Inteligencia de Movilizaciones',
  timezone: 'America/Guayaquil',
  locale: 'es-EC',
  version: '1.0.0'
};
