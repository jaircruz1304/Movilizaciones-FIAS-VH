import { APP_CONFIG } from '../config/msal-config.js?v=1.5.0';
import { initAuth, login, logout, isAuthenticated, getAccount, authDiagnostics } from './auth.js?v=1.5.0';
import { discoverLists, chooseList, loadItems, normalizeItems, saveMapping, sharepointState } from './sharepoint.js?v=1.5.0';
import { loadGpsData, reconcileMovements, gpsState } from './gps.js?v=1.5.0';
import { filterMovements } from './analytics.js?v=1.5.0';
import { initMap, invalidateMap, showCoverage, showFrequentRoutes, showGpsHeat, showAllGpsTrace, showTrip } from './maps.js?v=1.5.0';
import { $, $$, debounce, escapeHtml, setLoading, toast, fmtInt } from './utils.js?v=1.5.0';
import { renderAll, renderSourceModal, openSourceModal, closeSourceModal, closeDetail, collectMapping, exportCsv, bindDashboardCallbacks } from './dashboard.js?v=1.5.0';

const state={all:[],filtered:[],mapMode:'coverage',ready:false};
const VIEW_META={
  overview:['Panorama operativo','Indicadores y patrones relevantes del uso institucional de vehículos.'],
  territory:['Cobertura territorial','Rutas GPS, concentración espacial y puntos recurrentes de movilización.'],
  operations:['Explorador de movilizaciones','Cruce de registros SharePoint con evidencia satelital.'],
  fleet:['Uso y demanda','Comparación por grupos, usuarios, vehículos y períodos de utilización.'],
  gps:['Rastreo satelital','Trazas, eventos y métricas derivadas de los reportes GPS.'],
  quality:['Control de datos','Completitud, conciliación y excepciones que requieren revisión.']
};

function setAuthUi(){
  const ok=isAuthenticated();
  $('loginView').classList.toggle('hidden',ok);
  $('app').classList.toggle('hidden',!ok);
  const d=authDiagnostics();
  $('loginDiagnostics').textContent=`Redirect URI: ${d.redirectUri}`;
  if(ok){
    $('connectionText').textContent=getAccount()?.name||getAccount()?.username||'Sesión Microsoft 365';
  }
}

async function bootstrap(){
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
    await Promise.all([discoverLists(),loadGpsData()]);
    await chooseList();
    await loadItems();
    const movements=normalizeItems();
    state.all=reconcileMovements(movements);
    state.ready=true;
    populateFilters(); applyFilters(false); updateConnection(true);
    toast('Sincronización completa',`${fmtInt(state.all.length)} movilizaciones · ${fmtInt(gpsState.points.length)} puntos GPS.`);
  }catch(err){
    console.error(err);updateConnection(false);toast('Error de sincronización',err.message||String(err),'bad');
    openSourceModal();
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
  if(dates.length){$('fFrom').min=dateInput(dates[0]);$('fTo').max=dateInput(dates.at(-1));}
}
function unique(arr){return [...new Set(arr.filter(Boolean))].sort((a,b)=>String(a).localeCompare(String(b),'es'));}
function fill(id,values,label){const el=$(id),old=el.value;el.innerHTML=`<option value="">${escapeHtml(label)}</option>`+values.map(v=>`<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`).join('');if(values.includes(old))el.value=old;}
function dateInput(d){const z=new Intl.DateTimeFormat('en-CA',{timeZone:APP_CONFIG.timezone,year:'numeric',month:'2-digit',day:'2-digit'}).format(d);return z;}

function currentFilters(){return{search:$('fSearch').value,from:$('fFrom').value,to:$('fTo').value,group:$('fGroup').value,requester:$('fRequester').value,vehicle:$('fVehicle').value,activity:$('fActivity').value};}
function applyFilters(showToast=false){
  state.filtered=filterMovements(state.all,currentFilters());
  renderAll(state.filtered,state.all); renderMapMode();
  if(showToast)toast('Filtros aplicados',`${fmtInt(state.filtered.length)} movilizaciones visibles.`);
}
function resetFilters(){['fSearch','fFrom','fTo','fGroup','fRequester','fVehicle','fActivity'].forEach(id=>$(id).value='');applyFilters();}

function updateConnection(ok){
  $('connectionPill').classList.toggle('ok',ok);
  $('connectionText').textContent=ok?`${sharepointState.activeList?.displayName||sharepointState.activeList?.name||'SharePoint'} · ${fmtInt(state.all.length)}`:'Revisar conexión';
  $('sourceSubtitle').textContent=`${sharepointState.site?.displayName||'Microsoft 365'} · ${fmtInt(gpsState.points.length)} puntos GPS · ${APP_CONFIG.version}`;
}

function switchView(view){
  $$('.nav-tab').forEach(b=>b.classList.toggle('active',b.dataset.view===view));
  $$('.view').forEach(v=>v.classList.toggle('active',v.id===`view-${view}`));
  const meta=VIEW_META[view]||VIEW_META.overview;$('viewTitle').textContent=meta[0];$('viewSubtitle').textContent=meta[1];
  if(view==='territory'){initMap();invalidateMap();renderMapMode();}
}

function setMapMode(mode){state.mapMode=mode;$$('#mapModes button').forEach(b=>b.classList.toggle('active',b.dataset.map===mode));renderMapMode();}
function renderMapMode(){
  if(!$('view-territory').classList.contains('active'))return;
  try{
    if(state.mapMode==='coverage'){showCoverage(state.filtered);$('mapCaption').textContent='Destinos derivados de los puntos más alejados de la Matriz FIAS dentro de cada movilización relacionada.';}
    else if(state.mapMode==='routes'){showFrequentRoutes(state.filtered);$('mapCaption').textContent='Trazas GPS reales de hasta 30 movilizaciones con mayor recorrido dentro del filtro actual.';}
    else if(state.mapMode==='heat'){showGpsHeat(state.filtered.flatMap(r=>r.gpsTrace||[]));$('mapCaption').textContent='Concentración de posiciones GPS pertenecientes a movilizaciones relacionadas.';}
    else {showAllGpsTrace(gpsState.points);$('mapCaption').textContent='Histórico satelital completo incorporado de enero a julio de 2026.';}
  }catch(err){console.error(err);toast('Mapa',err.message||String(err),'warn');}
}

function openTripOnMap(movement){switchView('territory');state.mapMode='trip';$$('#mapModes button').forEach(b=>b.classList.remove('active'));showTrip(movement);$('mapCaption').textContent=`Ruta GPS relacionada con: ${movement.destinationLabel||movement.destination||'movilización'}.`;}

function bindEvents(){
  $('btnLogin').addEventListener('click',doLogin);
  $('btnSync').addEventListener('click',syncAll);
  $('btnLogout').addEventListener('click',async()=>{setLoading(true,'Cerrando sesión…');try{await logout();state.all=[];state.filtered=[];setAuthUi();}catch(err){toast('Cierre de sesión',err.message||String(err),'warn')}finally{setLoading(false)}});
  $$('.nav-tab').forEach(b=>b.addEventListener('click',()=>switchView(b.dataset.view)));
  const debounced=debounce(()=>applyFilters(),220);
  $('fSearch').addEventListener('input',debounced);
  ['fFrom','fTo','fGroup','fRequester','fVehicle','fActivity'].forEach(id=>$(id).addEventListener('change',()=>applyFilters()));
  $('btnResetFilters').addEventListener('click',resetFilters);
  $('mapModes').addEventListener('click',e=>{const b=e.target.closest('button[data-map]');if(b)setMapMode(b.dataset.map);});
  $('btnExport').addEventListener('click',()=>exportCsv(state.filtered));
  $('btnOpenSource').addEventListener('click',openSourceModal);
  $('btnCloseSource').addEventListener('click',closeSourceModal);
  $('btnCloseDrawer').addEventListener('click',closeDetail);
  $('detailDrawer').addEventListener('click',e=>{if(e.target===$('detailDrawer'))closeDetail();});
  $('sourceModal').addEventListener('click',e=>{if(e.target===$('sourceModal'))closeSourceModal();});
  $('btnApplyMapping').addEventListener('click',()=>{
    saveMapping(collectMapping());
    state.all=reconcileMovements(normalizeItems());populateFilters();applyFilters(false);closeSourceModal();toast('Mapeo aplicado','La analítica fue reprocesada con los campos seleccionados.');
  });
  bindDashboardCallbacks({onOpenTrip:openTripOnMap,onChooseList:changeList});
}

bootstrap();
