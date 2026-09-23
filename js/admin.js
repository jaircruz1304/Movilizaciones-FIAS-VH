import { AUTH_CONFIG, SHAREPOINT_GPS_CONFIG } from '../config/msal-config.js?v=2.1.0';
import { graph } from './graph.js?v=2.1.0';
import { resolveGpsSite } from './sharepoint.js?v=2.1.0';
import { getAuthenticatedEmail, isGpsAdministrator } from './auth.js?v=2.1.0';

let driveCache=null;

const adminAuth=()=>({authScopes:AUTH_CONFIG.adminScopes||AUTH_CONFIG.scopes});

function requireAdmin(){
  if(!isGpsAdministrator()) throw new Error('Esta función está restringida al administrador GPS autorizado.');
}

function encodePath(path=''){
  return String(path).split('/').filter(Boolean).map(encodeURIComponent).join('/');
}

function cleanFilename(name='reporte.pdf'){
  return String(name)
    .replace(/[\\/:*?"<>|#%]/g,'_')
    .replace(/\s+/g,' ')
    .trim()
    .slice(0,150) || 'reporte.pdf';
}

function uploadName(original){
  const stamp=new Date().toISOString().replace(/[-:]/g,'').replace(/\.\d{3}Z$/,'Z');
  return `${stamp}__${cleanFilename(original)}`;
}

export async function getGpsDrive(){
  if(driveCache) return driveCache;

  const site = await resolveGpsSite();

  const data = await graph(
    `/sites/${encodeURIComponent(site.id)}/drives?$select=id,name,webUrl,driveType`,
    adminAuth()
  );

  const drives = data?.value || [];

  if (!drives.length) {
    throw new Error(
      'No se encontraron bibliotecas de documentos accesibles en el sitio SharePoint.'
    );
  }

  // Biblioteca real utilizada por FIAS:
  // /Documentos compartidos/
  driveCache =
    drives.find(d => {
      const url = decodeURIComponent(String(d.webUrl || '')).toLowerCase();
      return url.includes('/documentos compartidos');
    }) ||
    drives.find(d =>
      ['documentos compartidos','documentos','documents','shared documents']
        .includes(String(d.name || '').trim().toLowerCase())
    ) ||
    drives[0];

  console.info('Biblioteca GPS seleccionada:', {
    id: driveCache.id,
    name: driveCache.name,
    webUrl: driveCache.webUrl
  });

  return driveCache;
}

async function getItemByPath(driveId,path){
  const encoded=encodePath(path);
  try{
    return await graph(`/drives/${encodeURIComponent(driveId)}/root:/${encoded}?$select=id,name,webUrl,parentReference,folder`,adminAuth());
  }catch(err){
    if(String(err?.message||err).includes('Microsoft Graph 404')) return null;
    throw err;
  }
}

export async function ensureFolderPath(path){
  requireAdmin();
  const drive=await getGpsDrive();
  const segments=String(path).split('/').filter(Boolean);
  let parent={id:'root'};
  let currentPath='';
  for(const segment of segments){
    currentPath=currentPath?`${currentPath}/${segment}`:segment;
    let item=await getItemByPath(drive.id,currentPath);
    if(!item){
      const endpoint=parent.id==='root'
        ? `/drives/${encodeURIComponent(drive.id)}/root/children`
        : `/drives/${encodeURIComponent(drive.id)}/items/${encodeURIComponent(parent.id)}/children`;
      item=await graph(endpoint,{
        ...adminAuth(),
        method:'POST',
        headers:{'Content-Type':'application/json'},
        body:JSON.stringify({name:segment,folder:{},'@microsoft.graph.conflictBehavior':'fail'})
      });
    }
    parent=item;
  }
  return parent;
}

export function validateGpsPdf(file){
  if(!file) throw new Error('Seleccione un archivo PDF.');
  const ext=String(file.name||'').toLowerCase().endsWith('.pdf');
  const mime=!file.type || (SHAREPOINT_GPS_CONFIG.acceptedMimeTypes||[]).includes(file.type);
  if(!ext || !mime) throw new Error('Solo se permiten reportes GPS en formato PDF.');
  const max=(Number(SHAREPOINT_GPS_CONFIG.maxUploadMb)||200)*1024*1024;
  if(file.size>max) throw new Error(`El PDF supera el límite de ${SHAREPOINT_GPS_CONFIG.maxUploadMb} MB.`);
  if(file.size<100) throw new Error('El archivo seleccionado está vacío o no es válido.');
  return true;
}

export async function uploadGpsPdf(file,onProgress=()=>{}){
  requireAdmin();
  validateGpsPdf(file);
  onProgress({stage:'preparing',pct:10,message:'Preparando carpeta de recepción…'});
  const drive=await getGpsDrive();
  await ensureFolderPath(SHAREPOINT_GPS_CONFIG.inboxFolder);
  const targetName=uploadName(file.name);
  const targetPath=`${SHAREPOINT_GPS_CONFIG.inboxFolder}/${targetName}`;
  onProgress({stage:'uploading',pct:35,message:'Cargando PDF a SharePoint…'});
  const item=await graph(`/drives/${encodeURIComponent(drive.id)}/root:/${encodePath(targetPath)}:/content`,{
    ...adminAuth(),
    method:'PUT',
    headers:{'Content-Type':'application/pdf'},
    body:file
  });
  onProgress({stage:'queued',pct:100,message:'Carga completada. El reporte quedó en cola para GitHub Actions.'});
  return {
    ...item,
    originalName:file.name,
    queuedName:targetName,
    uploadedBy:getAuthenticatedEmail(),
    queueFolder:SHAREPOINT_GPS_CONFIG.inboxFolder
  };
}

export async function listPendingUploads(){
  requireAdmin();

  const drive = await getGpsDrive();
  const folder = await ensureFolderPath(
    SHAREPOINT_GPS_CONFIG.inboxFolder
  );

  const data = await graph(
    `/drives/${encodeURIComponent(drive.id)}/items/${encodeURIComponent(folder.id)}/children?$select=id,name,size,createdDateTime,createdBy,webUrl,file&$top=100`,
    adminAuth()
  );

  return (data?.value || [])
    .filter(x =>
      x.file &&
      String(x.name || '').toLowerCase().endsWith('.pdf')
    )
    .sort((a,b) =>
      String(b.createdDateTime || '')
        .localeCompare(String(a.createdDateTime || ''))
    );
}

export async function loadGpsHistory(){
  requireAdmin();
  try{
    const drive=await getGpsDrive();
    const path=`${SHAREPOINT_GPS_CONFIG.publishedFolder}/${SHAREPOINT_GPS_CONFIG.historyName}`;
    const data=await graph(`/drives/${encodeURIComponent(drive.id)}/root:/${encodePath(path)}:/content`,adminAuth());
    const parsed=typeof data==='string'?JSON.parse(data):data;
    return Array.isArray(parsed)?parsed:(parsed?.entries||[]);
  }catch(err){
    if(String(err?.message||err).includes('Microsoft Graph 404'))return [];
    console.warn('No se pudo cargar historial GPS',err);
    return [];
  }
}

export function adminStatus(){
  return {
    email:getAuthenticatedEmail(),
    authorized:isGpsAdministrator(),
    inboxFolder:SHAREPOINT_GPS_CONFIG.inboxFolder,
    intervalMinutes:SHAREPOINT_GPS_CONFIG.actionIntervalMinutes
  };
}
