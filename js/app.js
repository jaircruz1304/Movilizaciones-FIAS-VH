import { APP_CONFIG, SHAREPOINT_GPS_CONFIG } from '../config/msal-config.js?v=2.2.0';
import { initAuth, login, logout, isAuthenticated, getAccount, authDiagnostics, isGpsAdministrator, getAuthenticatedEmail } from './auth.js?v=2.2.0';
import { discoverLists, chooseList, loadItems, normalizeItems, saveMapping, sharepointState } from './sharepoint.js?v=2.2.0';
import { loadGpsData, reconcileMovements, gpsState } from './gps.js?v=2.2.0';
import { filterMovements } from './analytics.js?v=2.2.0';
import { initMap, invalidateMap, showCoverage, showFrequentRoutes, showGpsHeat, showAllGpsTrace, showTrip } from './maps.js?v=2.2.0';
import { $, $$, debounce, escapeHtml, setLoading, toast, fmtInt, formatDateTime } from './utils.js?v=2.2.0';
import { renderAll, renderSourceModal, openSourceModal, closeSourceModal, closeDetail, collectMapping, exportCsv, exportXlsx, exportPdf, bindDashboardCallbacks } from './dashboard.js?v=2.2.0';
import { uploadGpsPdf, listPendingUploads, loadGpsHistory, adminStatus, validateGpsPdf } from './admin.js?v=2.2.0';

const state={all:[],filtered:[],mapMode:'coverage',ready:false,selectedGpsFiles:[],activeView:'overview'};
const VIEW_META={
  overview:['Panorama operativo','Indicadores y patrones relevantes del uso institucional de vehículos.'],
  territory:['Cobertura territorial','Rutas GPS, concentración espacial y puntos recurrentes de movilización.'],
  operations:['Explorador de movilizaciones','Cruce de registros SharePoint con evidencia satelital.'],
  fleet:['Uso y demanda','Distribución de uso por grupos, usuarios, vehículos y días de utilización.'],
  gps:['Rastreo satelital','Trazas, eventos y métricas derivadas de los reportes GPS.'],
  quality:['Control de datos','Completitud, conciliación y excepciones que requieren revisión.'],
  admin:['Administración GPS','Carga controlada, cola de procesamiento e historial de auditoría de reportes satelitales.']
};

function initTheme(){
  const stored=localStorage.getItem('fias.theme');
  const theme=stored || (window.matchMedia?.('(prefers-color-scheme: dark)').matches?'dark':'light');
  setTheme(theme,false);
}
function setTheme(theme,persist=true){
  const value=theme==='dark'?'dark':'light';
  document.documentElement.dataset.theme=value;
  if(persist)localStorage.setItem('fias.theme',value);
  const b=$('btnTheme');
  if(b){
    b.querySelector('i').className=`bi ${value==='dark'?'bi-sun':'bi-moon-stars'}`;
    const span=b.querySelector('span'); if(span)span.textContent=value==='dark'?'Claro':'Oscuro';
  }
}
function toggleTheme(){
  setTheme(document.documentElement.dataset.theme==='dark'?'light':'dark');
  if(state.ready){renderAll(state.filtered,state.all); if(state.activeView==='territory')renderMapMode();}
}

function setAuthUi(){
  const ok=isAuthenticated();
  $('loginView').classList.toggle('hidden',ok);
  $('app').classList.toggle('hidden',!ok);
  const d=authDiagnostics();
  $('loginDiagnostics').textContent=`Redirect URI: ${d.redirectUri}`;
  const admin=ok&&isGpsAdministrator();
  $('navAdmin').classList.toggle('hidden',!admin);
  if(ok){
    $('connectionText').textContent=getAccount()?.name||getAccount()?.username||'Sesión Microsoft 365';
    $('adminIdentity').textContent=admin?`Administrador · ${getAuthenticatedEmail()}`:'';
  }
  if(!admin && state.activeView==='admin') switchView('overview');
}

async function bootstrap(){
  initTheme();
  $('loginDiagnostics').textContent='Preparando autenticación Microsoft 365…';
  try{
    await initAuth(); setAuthUi(); bindEvents();
    if(isAuthenticated()) await syncAll();
  }catch(err){
    console.error(err); $('loginDiagnostics').textContent=err.message||String(err); toast('Inicialización',err.message||String(err),'bad');
  }
}

async function doLogin(){
  setLoading(true,'Autenticando con Microsoft 365…');
  try{await login();setAuthUi();await syncAll();}
  catch(err){console.error(err);toast('No fue posible iniciar sesión',err.message||String(err),'bad');}
  finally{setLoading(false);}
}

async function syncAll(){
  setLoading(true,'Leyendo SharePoint y preparando rastreo GPS…');
  try{
    await Promise.all([discoverLists(),loadGpsData(true)]);
    await chooseList();
    await loadItems();
    const movements=normalizeItems();
    state.all=reconcileMovements(movements);
    state.ready=true;
    populateFilters(); applyFilters(false); updateConnection(true);
    if(isGpsAdministrator()&&state.activeView==='admin') await refreshAdmin(false);
    toast('Sincronización completa',`${fmtInt(state.all.length)} movilizaciones · ${fmtInt(gpsState.points.length)} puntos GPS.`);
  }catch(err){
    console.error(err);updateConnection(false);toast('Error de sincronización',err.message||String(err),'bad');
    if(isGpsAdministrator()) openSourceModal();
  }finally{setLoading(false);}
}

async function changeList(id){
  setLoading(true,'Cambiando lista de SharePoint…');
  try{
    await chooseList(id);await loadItems();
    state.all=reconcileMovements(normalizeItems());
    populateFilters();applyFilters(false);renderSourceModal();closeSourceModal();updateConnection(true);
    toast('Lista activada',sharepointState.activeList?.displayName||sharepointState.activeList?.name||'SharePoint');
  }catch(err){toast('No se pudo cambiar la lista',err.message||String(err),'bad');}
  finally{setLoading(false);}
}

function populateFilters(){
  fill('fGroup',unique(state.all.map(r=>r.project||r.group)),'Todos');
  fill('fRequester',unique(state.all.map(r=>r.requester)),'Todos');
  fill('fVehicle',unique(state.all.map(r=>r.vehicle||r.plate)),'Todos');
  fill('fActivity',unique(state.all.map(r=>r.activityCategory)),'Todas');
  const dates=state.all.map(r=>r.start).filter(Boolean).sort((a,b)=>a-b);
  if(dates.length){
    const min=dateInput(dates[0]),max=dateInput(dates.at(-1));
    $('fFrom').min=min;$('fFrom').max=max;$('fTo').min=min;$('fTo').max=max;
  }
}
function unique(arr){return [...new Set(arr.filter(Boolean))].sort((a,b)=>String(a).localeCompare(String(b),'es'));}
function fill(id,values,label){const el=$(id),old=el.value;el.innerHTML=`<option value="">${escapeHtml(label)}</option>`+values.map(v=>`<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`).join('');if(values.includes(old))el.value=old;}
function dateInput(d){return new Intl.DateTimeFormat('en-CA',{timeZone:APP_CONFIG.timezone,year:'numeric',month:'2-digit',day:'2-digit'}).format(d);}

function currentFilters(){return{search:$('fSearch').value,from:$('fFrom').value,to:$('fTo').value,group:$('fGroup').value,requester:$('fRequester').value,vehicle:$('fVehicle').value,activity:$('fActivity').value};}

const FILTER_META={
  search:['fSearch','Búsqueda'],from:['fFrom','Desde'],to:['fTo','Hasta'],group:['fGroup','Grupo'],requester:['fRequester','Usuario'],vehicle:['fVehicle','Vehículo'],activity:['fActivity','Actividad']
};
function activeFilterEntries(){
  const values=currentFilters();
  return Object.entries(values).filter(([,value])=>String(value||'').trim()).map(([key,value])=>({key,value:String(value),id:FILTER_META[key][0],label:FILTER_META[key][1]}));
}
function setFiltersOpen(open){
  const panel=$('globalFilters'),button=$('btnToggleFilters');
  if(!panel||!button)return;
  const visible=Boolean(open)&&state.activeView!=='admin';
  panel.classList.toggle('collapsed',!visible);panel.setAttribute('aria-hidden',String(!visible));button.setAttribute('aria-expanded',String(visible));
  button.classList.toggle('active',visible);
}
function updateFilterUi(){
  const entries=activeFilterEntries(),count=entries.length;
  const badge=$('filterCount'),summary=$('filterSummary');
  if(badge){badge.textContent=String(count);badge.classList.toggle('hidden',count===0);}
  if(!summary)return;
  if(!count){summary.innerHTML=`<span class="filter-summary-neutral"><i class="bi bi-check2-circle"></i> Todos los registros · ${fmtInt(state.filtered.length||state.all.length)} movilizaciones</span>`;return;}
  summary.innerHTML=`<span class="filter-summary-label">${fmtInt(state.filtered.length)} resultados</span>`+entries.map(x=>`<button type="button" class="filter-chip" data-clear-filter="${escapeHtml(x.id)}" title="Quitar ${escapeHtml(x.label)}"><span>${escapeHtml(x.label)}</span><strong>${escapeHtml(x.value)}</strong><i class="bi bi-x"></i></button>`).join('');
}
function applyFilters(showToast=false){
  state.filtered=filterMovements(state.all,currentFilters());
  renderAll(state.filtered,state.all); renderMapMode(); updateFilterUi();
  if(showToast)toast('Filtros aplicados',`${fmtInt(state.filtered.length)} movilizaciones visibles.`);
}
function resetFilters(){['fSearch','fFrom','fTo','fGroup','fRequester','fVehicle','fActivity'].forEach(id=>$(id).value='');applyFilters();}

function updateConnection(ok){
  $('connectionPill').classList.toggle('ok',ok);
  $('connectionText').textContent=ok?`${sharepointState.activeList?.displayName||sharepointState.activeList?.name||'SharePoint'} · ${fmtInt(state.all.length)}`:'Revisar conexión';
  $('sourceSubtitle').textContent=`${sharepointState.site?.displayName||'Microsoft 365'} · ${fmtInt(gpsState.points.length)} puntos GPS · ${APP_CONFIG.version}`;
}

function switchView(view){
  if(view==='admin'&&!isGpsAdministrator()){toast('Acceso restringido','La administración GPS está habilitada únicamente para jcruzg@fias.org.ec.','warn');return;}
  state.activeView=view;
  $$('.nav-tab').forEach(b=>b.classList.toggle('active',b.dataset.view===view));
  $$('.view').forEach(v=>v.classList.toggle('active',v.id===`view-${view}`));
  const meta=VIEW_META[view]||VIEW_META.overview;$('viewTitle').textContent=meta[0];$('viewSubtitle').textContent=meta[1];
  $('filterActions').classList.toggle('hidden',view==='admin');
  $('filterSummary').classList.toggle('hidden',view==='admin');
  setFiltersOpen(false);
  if(view==='territory'){initMap();invalidateMap();renderMapMode();}
  if(view==='admin')refreshAdmin(false);
}

function setMapMode(mode){state.mapMode=mode;$$('#mapModes button').forEach(b=>b.classList.toggle('active',b.dataset.map===mode));renderMapMode();}
function gpsPeriodLabel(){
  const files=gpsState.manifest?.files||[]; if(!files.length)return 'histórico publicado';
  const months=files.map(f=>f.month).filter(Boolean).sort(); return `${months[0]} a ${months.at(-1)}`;
}
function renderMapMode(){
  if(!$('view-territory').classList.contains('active'))return;
  try{
    if(state.mapMode==='coverage'){showCoverage(state.filtered);$('mapCaption').textContent='Destinos derivados de los puntos más alejados de la Matriz FIAS dentro de cada movilización relacionada.';}
    else if(state.mapMode==='routes'){showFrequentRoutes(state.filtered);$('mapCaption').textContent='Trazas GPS reales de hasta 30 movilizaciones con mayor recorrido dentro del filtro actual.';}
    else if(state.mapMode==='heat'){showGpsHeat(state.filtered.flatMap(r=>r.gpsTrace||[]));$('mapCaption').textContent='Concentración de posiciones GPS pertenecientes a movilizaciones relacionadas.';}
    else {showAllGpsTrace(gpsState.points);$('mapCaption').textContent=`Histórico satelital completo incorporado: ${gpsPeriodLabel()}.`;}
  }catch(err){console.error(err);toast('Mapa',err.message||String(err),'warn');}
}

function openTripOnMap(movement){switchView('territory');state.mapMode='trip';$$('#mapModes button').forEach(b=>b.classList.remove('active'));showTrip(movement);$('mapCaption').textContent=`Ruta GPS relacionada con: ${movement.destinationLabel||movement.destination||'movilización'}.`;}

function fileSize(bytes=0){const n=Number(bytes)||0;if(n<1024)return `${n} B`;if(n<1024**2)return `${(n/1024).toFixed(1)} KB`;return `${(n/1024**2).toFixed(1)} MB`;}
function selectedFiles(files){
  const accepted=[];
  for(const f of files||[]){
    try{validateGpsPdf(f);accepted.push(f);}catch(err){toast(f.name,err.message||String(err),'warn');}
  }
  const byKey=new Map([...state.selectedGpsFiles,...accepted].map(f=>[`${f.name}|${f.size}|${f.lastModified}`,f]));
  state.selectedGpsFiles=[...byKey.values()];renderSelectedGpsFiles();
}
function renderSelectedGpsFiles(){
  $('selectedGpsFiles').innerHTML=state.selectedGpsFiles.length?state.selectedGpsFiles.map((f,i)=>`<div class="selected-file"><div><strong>${escapeHtml(f.name)}</strong><small>${escapeHtml(fileSize(f.size))}</small></div><button class="icon-action" data-remove-file="${i}" title="Quitar"><i class="bi bi-x-lg"></i></button></div>`).join(''):'';
  $('selectedGpsFiles').querySelectorAll('[data-remove-file]').forEach(b=>b.addEventListener('click',()=>{state.selectedGpsFiles.splice(Number(b.dataset.removeFile),1);renderSelectedGpsFiles();}));
}
function setUploadProgress(pct,text){$('uploadProgress').classList.remove('hidden');$('uploadProgressBar').style.width=`${Math.max(0,Math.min(100,pct))}%`;$('uploadProgressText').textContent=text;}
async function uploadSelectedGps(){
  if(!isGpsAdministrator())return;
  if(!state.selectedGpsFiles.length){toast('Carga GPS','Seleccione al menos un reporte PDF.','warn');return;}
  $('btnUploadGps').disabled=true;
  try{
    const total=state.selectedGpsFiles.length;
    for(let i=0;i<total;i++){
      const f=state.selectedGpsFiles[i];
      await uploadGpsPdf(f,p=>{const overall=((i+(p.pct||0)/100)/total)*100;setUploadProgress(overall,`${i+1}/${total} · ${f.name} · ${p.message}`);});
    }
    state.selectedGpsFiles=[];renderSelectedGpsFiles();$('gpsPdfInput').value='';
    setUploadProgress(100,`Carga completada. GitHub Actions procesará la bandeja aproximadamente cada ${SHAREPOINT_GPS_CONFIG.actionIntervalMinutes} minutos.`);
    toast('Reportes en cola','Los PDF fueron cargados a SharePoint. No se requiere reprocesamiento manual.');
    await refreshAdmin(false);
  }catch(err){console.error(err);toast('Carga GPS',err.message||String(err),'bad');setUploadProgress(0,err.message||String(err));}
  finally{$('btnUploadGps').disabled=false;}
}
function historyStatus(status){
  const labels={processed:'Procesado',baseline:'Base validada',duplicate:'Duplicado',rejected:'Rechazado'};
  return `<span class="pill ${status==='rejected'?'bad':status==='duplicate'?'warn':'ok'}"><span class="status-dot ${escapeHtml(status||'')}"></span>${escapeHtml(labels[status]||status||'—')}</span>`;
}
async function refreshAdmin(showNotice=true){
  if(!isGpsAdministrator())return;
  try{
    const status=adminStatus();$('adminIdentity').textContent=`Administrador · ${status.email}`;
    const [pending,history]=await Promise.all([listPendingUploads(),loadGpsHistory()]);
    $('pendingGpsFiles').innerHTML=pending.length?pending.map(x=>`<div class="source-card"><div><strong>${escapeHtml((x.name||'').split('__',2).at(-1)||x.name||'PDF')}</strong><small>${escapeHtml(fileSize(x.size))} · ${escapeHtml(formatDateTime(x.createdDateTime?new Date(x.createdDateTime):null))}</small></div><span class="pill warn">En cola</span></div>`).join(''):'<div class="empty">No existen PDF pendientes de procesamiento.</div>';
    const rows=[...history].sort((a,b)=>String(b.processedAt||b.uploadedAt||'').localeCompare(String(a.processedAt||a.uploadedAt||''))).slice(0,150);
    $('gpsHistoryRows').innerHTML=rows.length?rows.map(x=>{
      const period=(x.months||[]).join(', ')||x.month||'—'; const tracker=x.tracker||'—';
      return `<tr><td>${escapeHtml(formatDateTime(new Date(x.processedAt||x.uploadedAt||Date.now())))}</td><td>${escapeHtml(x.sourceFile||x.queuedName||'—')}</td><td>${escapeHtml(x.uploadedBy||'—')}</td><td>${escapeHtml(`${tracker} · ${period}`)}</td><td>${escapeHtml(fmtInt(x.addedPoints??x.inputPoints??0))}</td><td>${historyStatus(x.status)}</td><td class="history-detail">${escapeHtml(x.message||x.error||'—')}</td></tr>`;
    }).join(''):'<tr><td colspan="7"><div class="empty">Aún no existe historial publicado.</div></td></tr>';
    if(showNotice)toast('Administración GPS',`${fmtInt(pending.length)} PDF pendientes · ${fmtInt(rows.length)} registros de auditoría.`);
  }catch(err){console.error(err);toast('Administración GPS',err.message||String(err),'bad');}
}

function safeExport(fn,label){try{if(!state.filtered.length){toast(label,'No hay registros para exportar.','warn');return;}fn(state.filtered);}catch(err){toast(label,err.message||String(err),'bad');}}

function bindEvents(){
  $('btnLogin').addEventListener('click',doLogin);
  $('btnTheme').addEventListener('click',toggleTheme);
  $('btnSync').addEventListener('click',syncAll);
  $('btnLogout').addEventListener('click',async()=>{setLoading(true,'Cerrando sesión…');try{await logout();state.all=[];state.filtered=[];state.ready=false;setAuthUi();}catch(err){toast('Cierre de sesión',err.message||String(err),'warn')}finally{setLoading(false)}});
  $$('.nav-tab').forEach(b=>b.addEventListener('click',()=>switchView(b.dataset.view)));
  const debounced=debounce(()=>applyFilters(),220);
  $('fSearch').addEventListener('input',debounced);
  ['fFrom','fTo','fGroup','fRequester','fVehicle','fActivity'].forEach(id=>$(id).addEventListener('change',()=>applyFilters()));
  $('btnResetFilters').addEventListener('click',resetFilters);
  $('btnToggleFilters').addEventListener('click',()=>setFiltersOpen($('globalFilters').classList.contains('collapsed')));
  $('btnCloseFilters').addEventListener('click',()=>setFiltersOpen(false));
  $('filterSummary').addEventListener('click',e=>{const chip=e.target.closest('[data-clear-filter]');if(!chip)return;const input=$(chip.dataset.clearFilter);if(input){input.value='';applyFilters();}});
  document.addEventListener('keydown',e=>{if(e.key==='Escape'&&!$('globalFilters').classList.contains('collapsed'))setFiltersOpen(false);});
  $('mapModes').addEventListener('click',e=>{const b=e.target.closest('button[data-map]');if(b)setMapMode(b.dataset.map);});
  $('btnExportCsv').addEventListener('click',()=>safeExport(exportCsv,'Exportación CSV'));
  $('btnExportXlsx').addEventListener('click',()=>safeExport(exportXlsx,'Exportación Excel'));
  $('btnExportPdf').addEventListener('click',()=>safeExport(exportPdf,'Exportación PDF'));
  $('btnOpenSource').addEventListener('click',()=>{if(isGpsAdministrator())openSourceModal();});
  $('btnCloseSource').addEventListener('click',closeSourceModal);
  $('btnCloseDrawer').addEventListener('click',closeDetail);
  $('detailDrawer').addEventListener('click',e=>{if(e.target===$('detailDrawer'))closeDetail();});
  $('sourceModal').addEventListener('click',e=>{if(e.target===$('sourceModal'))closeSourceModal();});
  $('btnApplyMapping').addEventListener('click',()=>{
    saveMapping(collectMapping());
    state.all=reconcileMovements(normalizeItems());populateFilters();applyFilters(false);closeSourceModal();toast('Mapeo aplicado','La analítica fue reprocesada con los campos seleccionados.');
  });
  $('gpsPdfInput').addEventListener('change',e=>selectedFiles([...e.target.files]));
  const dz=$('gpsDropzone');
  ['dragenter','dragover'].forEach(type=>dz.addEventListener(type,e=>{e.preventDefault();dz.classList.add('dragover');}));
  ['dragleave','drop'].forEach(type=>dz.addEventListener(type,e=>{e.preventDefault();dz.classList.remove('dragover');}));
  dz.addEventListener('drop',e=>selectedFiles([...e.dataTransfer.files]));
  $('btnUploadGps').addEventListener('click',uploadSelectedGps);
  $('btnRefreshAdmin').addEventListener('click',()=>refreshAdmin(true));
  bindDashboardCallbacks({onOpenTrip:openTripOnMap,onChooseList:changeList});
}

bootstrap();
