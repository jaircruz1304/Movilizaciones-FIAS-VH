import { SHAREPOINT_CONFIG } from '../config/msal-config.js';
import { graph, graphPaged } from './graph.js';
import { compactKey, normalizeText, textValue, toDate, toNumber, hoursBetween } from './utils.js';

const SEMANTICS={
  start:{label:'Fecha inicio uso',aliases:['fecha inicia uso','fecha inicio uso','fechainiciauso','fechainiciouso','fecha inicio','inicio uso','fecha salida','fecha movilizacion','fecha viaje']},
  end:{label:'Fecha termina uso',aliases:['fecha termina uso','fecha termino uso','fechaterminauso','fecha fin uso','fecha retorno','fecha llegada','fecha termina','fecha fin']},
  requestDate:{label:'Fecha solicitud',aliases:['fecha solicitud','fechasolicitud','solicitud fecha','fecha de solicitud','created']},
  requester:{label:'Usuario solicitante',aliases:['usuario1','usuario','solicitante','usuario solicitante','funcionario','servidor','requirente','peticionario']},
  group:{label:'Grupo / programa',aliases:['grupo','programa','proyecto','fondo','componente','area','dependencia','unidad']},
  destination:{label:'Destino / finalidad',aliases:['destino','lugar destino','ubicacion destino','ruta','lugar','descripcion destino','finalidad','actividad']},
  kmStart:{label:'KM inicial',aliases:['km inicial','kminicial','kilometraje inicial','odometro inicial','kmsalida']},
  kmEnd:{label:'KM final',aliases:['km final','kmfinal','kilometraje final','odometro final','kmllegada']},
  distance:{label:'Recorrido',aliases:['recorrido','km recorrido','kilometros recorridos','distancia','kilometraje recorrido','total km']},
  vehicle:{label:'Vehículo',aliases:['vehiculo','vehículo','unidad vehicular','automotor','activo','vehicle']},
  plate:{label:'Placa',aliases:['placa','matricula','matrícula']},
  driver:{label:'Conductor',aliases:['conductor','chofer','driver','responsable conduccion','responsable conducción']},
  activity:{label:'Actividad',aliases:['actividad','motivo','objeto movilizacion','objeto comisión','trabajo a realizar','finalidad']},
  project:{label:'Proyecto',aliases:['proyecto','programa','grupo','centro costo','fondo','componente']},
  status:{label:'Estado',aliases:['estado','status','estado movilizacion']}
};

const state={site:null,lists:[],activeList:null,columns:[],items:[],mapping:{},userMap:new Map(),lookupMaps:new Map()};
export const sharepointState=state;

export async function resolveSite(){
  if(state.site) return state.site;
  const path=SHAREPOINT_CONFIG.sitePath.replace(/^\//,'');
  state.site=await graph(`/sites/${SHAREPOINT_CONFIG.host}:/${path}?$select=id,displayName,webUrl`);
  return state.site;
}

function semanticScore(columns,title=''){
  const names=columns.flatMap(c=>[compactKey(c.name),compactKey(c.displayName)]).filter(Boolean);
  let score=0;
  for(const [key,meta] of Object.entries(SEMANTICS)){
    if(meta.aliases.some(a=>names.some(n=>n===compactKey(a)||n.includes(compactKey(a))||compactKey(a).includes(n)))) score += ['start','destination','distance'].includes(key)?5:2;
  }
  if(/moviliz|vehicul|transporte|salida|comision|recorrido/i.test(normalizeText(title))) score+=10;

  // Prioridad fuerte para la lista que coincide con la estructura real reportada por FIAS.
  const expected=(SHAREPOINT_CONFIG.expectedColumns||[]).map(compactKey).filter(Boolean);
  let signatureHits=0;
  for(const expectedName of expected){
    if(names.some(n=>n===expectedName || n.includes(expectedName) || expectedName.includes(n))){
      signatureHits++;
      score+=12;
    }
  }
  if(expected.length && signatureHits>=Math.max(6,expected.length-2)) score+=60;
  if(expected.length && signatureHits===expected.length) score+=80;
  return score;
}

export async function discoverLists(){
  const site=await resolveSite();
  const path=`/sites/${encodeURIComponent(site.id)}/lists?$select=id,name,displayName,webUrl,list&$expand=columns&$top=200`;
  const data=await graph(path);
  const all=(data?.value||[]).map(l=>({...l,score:semanticScore(l.columns||[],l.displayName||l.name)}));
  state.lists=all.sort((a,b)=>b.score-a.score);
  return state.lists;
}

export async function chooseList(id=''){
  if(!state.lists.length) await discoverLists();
  const visible=state.lists.filter(x=>!(x.list?.hidden));
  const remembered=localStorage.getItem('fias.movilizaciones.listId')||'';
  let list=null;

  // 1) Selección explícita desde la interfaz.
  if(id) list=visible.find(x=>x.id===id);
  // 2) GUID fijo, si en el futuro se configura uno.
  if(!list && SHAREPOINT_CONFIG.preferredListId) list=visible.find(x=>x.id===SHAREPOINT_CONFIG.preferredListId);
  // 3) Para esta solución se prioriza la mejor coincidencia por firma de columnas.
  if(!list && SHAREPOINT_CONFIG.lockToBestMatch) list=visible[0];
  // 4) Solo como respaldo se usa la lista recordada.
  if(!list && remembered) list=visible.find(x=>x.id===remembered);
  if(!list) list=visible[0];
  if(!list) throw new Error('No se encontró una lista de SharePoint accesible para movilizaciones.');
  state.activeList=list;
  state.columns=(list.columns||[]).filter(c=>!c.hidden);
  localStorage.setItem('fias.movilizaciones.listId',list.id);
  state.mapping=detectMapping(state.columns);
  await loadReferenceMaps().catch(err=>console.warn('No se pudieron resolver referencias de Persona/Lookup',err));
  return list;
}

export function detectMapping(columns=state.columns){
  const savedKey=state.activeList?`fias.movilizaciones.mapping.${state.activeList.id}`:'';
  let saved={}; try{saved=JSON.parse(savedKey?localStorage.getItem(savedKey)||'{}':'{}')}catch{}
  const result={}; const used=new Set();
  for(const [key,meta] of Object.entries(SEMANTICS)){
    if(saved[key] && columns.some(c=>c.name===saved[key])){result[key]=saved[key];used.add(saved[key]);continue;}
    let best='',bestScore=0;
    for(const c of columns){
      if(used.has(c.name)) continue;
      const cn=compactKey(c.name), cd=compactKey(c.displayName);
      for(const alias of meta.aliases){
        const a=compactKey(alias); let s=0;
        if(cn===a||cd===a)s=100;
        else if(cn.includes(a)||cd.includes(a))s=80;
        else if(a.includes(cn)||a.includes(cd))s=50;
        if(s>bestScore){bestScore=s;best=c.name;}
      }
    }
    result[key]=bestScore>=50?best:'';
    if(result[key])used.add(result[key]);
  }
  return result;
}

export function saveMapping(mapping){
  state.mapping={...mapping};
  if(state.activeList) localStorage.setItem(`fias.movilizaciones.mapping.${state.activeList.id}`,JSON.stringify(state.mapping));
}

function cleanLookupLabel(value){
  if(value===null||value===undefined) return '';
  if(Array.isArray(value)) return value.map(cleanLookupLabel).filter(Boolean).join(', ');
  if(typeof value==='object') return textValue(value);
  const s=String(value).trim();
  // Formatos clásicos de SharePoint: "12;#Nombre" o "12;#Nombre;#18;#Otro".
  if(s.includes(';#')){
    const parts=s.split(';#').map(x=>x.trim()).filter(Boolean);
    const labels=parts.filter(x=>!/^\d+$/.test(x));
    if(labels.length) return labels.join(', ');
  }
  return s;
}

function resolveLookupValue(value,map){
  if(value===null||value===undefined||value==='') return '';
  if(Array.isArray(value)) return value.map(v=>resolveLookupValue(v,map)).filter(Boolean).join(', ');
  if(typeof value==='object') return textValue(value);
  const raw=String(value).trim();
  if(raw.includes(';#')){
    const parts=raw.split(';#').map(x=>x.trim()).filter(Boolean);
    const out=[];
    for(let i=0;i<parts.length;i++){
      const part=parts[i];
      if(/^\d+$/.test(part)){
        const found=map?.get(String(part));
        if(found) out.push(found);
        else if(parts[i+1] && !/^\d+$/.test(parts[i+1])) out.push(parts[++i]);
      }else out.push(part);
    }
    return [...new Set(out.filter(Boolean))].join(', ');
  }
  if(/^\d+$/.test(raw)) return map?.get(raw)||'';
  return cleanLookupLabel(raw);
}

async function findUserInformationList(){
  const rx=/user\s*information\s*list|userinfo|user\s*info|informaci[oó]n\s*del\s*usuario|lista\s*de\s*informaci[oó]n/i;
  let found=state.lists.find(l=>rx.test(`${l.displayName||''} ${l.name||''}`));
  if(found) return found;
  // Algunos sitios no devuelven la lista del sistema en la primera consulta expandida.
  const data=await graph(`/sites/${encodeURIComponent(state.site.id)}/lists?$select=id,name,displayName,webUrl,list&$top=500`);
  const more=data?.value||[];
  found=more.find(l=>rx.test(`${l.displayName||''} ${l.name||''}`));
  if(found && !state.lists.some(x=>x.id===found.id)) state.lists.push(found);
  return found||null;
}

async function loadUserMap(){
  if(!state.site) return;
  const userList=await findUserInformationList();
  if(!userList){ state.userMap=new Map(); return; }
  const items=await graphPaged(`/sites/${encodeURIComponent(state.site.id)}/lists/${encodeURIComponent(userList.id)}/items?$expand=fields&$top=999`,10000);
  const map=new Map();
  for(const it of items){
    const f=it.fields||{};
    const label=cleanLookupLabel(f.Title||f.Name||f.UserName||f.EMail||f.Email||f.SipAddress||'');
    if(label) map.set(String(it.id),label);
  }
  state.userMap=map;
}

async function loadGenericLookupMaps(){
  state.lookupMaps=new Map();
  if(!state.site||!state.columns?.length) return;
  const relevant=new Set(Object.values(state.mapping||{}).filter(Boolean));
  for(const col of state.columns){
    if(!relevant.has(col.name) || !col.lookup?.listId) continue;
    try{
      const targetColumn=col.lookup.lookupColumn||'Title';
      const items=await graphPaged(`/sites/${encodeURIComponent(state.site.id)}/lists/${encodeURIComponent(col.lookup.listId)}/items?$expand=fields&$top=999`,10000);
      const map=new Map();
      for(const it of items){
        const f=it.fields||{};
        const label=cleanLookupLabel(f[targetColumn]??f.Title??f.Name??f.Value??'');
        if(label) map.set(String(it.id),label);
      }
      if(map.size) state.lookupMaps.set(col.name,map);
    }catch(err){
      console.warn(`No se pudo resolver Lookup ${col.displayName||col.name}`,err);
    }
  }
}

async function loadReferenceMaps(){
  await Promise.all([
    loadUserMap().catch(err=>console.warn('No se pudo resolver User Information List',err)),
    loadGenericLookupMaps().catch(err=>console.warn('No se pudieron resolver Lookup de la lista',err))
  ]);
}

export async function loadItems(){
  if(!state.site) await resolveSite();
  if(!state.activeList) await chooseList();
  const path=`/sites/${encodeURIComponent(state.site.id)}/lists/${encodeURIComponent(state.activeList.id)}/items?$expand=fields&$top=999`;
  state.items=await graphPaged(path,SHAREPOINT_CONFIG.maxItems);
  return state.items;
}

function fieldRaw(fields,key){
  const name=state.mapping[key]; if(!name) return '';
  const direct=fields[name];
  const normalized=compactKey(name);
  const lookupKey=Object.keys(fields).find(k=>compactKey(k)===`${normalized}lookupid`);
  const lookup=lookupKey?fields[lookupKey]:fields[`${name}LookupId`];
  const column=state.columns.find(c=>c.name===name);
  const genericMap=state.lookupMaps.get(name);
  const preferredMap=genericMap || ((column?.personOrGroup || key==='requester' || key==='driver') ? state.userMap : null);

  // Para Persona/Lookup se prioriza resolver el LookupId. Así evitamos mostrar "17" en Usuario1.
  if(lookup!==undefined && lookup!==null && lookup!==''){
    const resolved=resolveLookupValue(lookup,preferredMap||state.userMap);
    if(resolved) return resolved;
  }

  if(direct!==undefined && direct!==null && direct!==''){
    // Algunas respuestas de Graph entregan el ID directamente en el nombre del campo.
    const resolved=resolveLookupValue(direct,preferredMap);
    if(resolved) return resolved;
    const cleaned=cleanLookupLabel(direct);
    if(cleaned && !(/^\d+$/.test(cleaned) && preferredMap)) return cleaned;
  }

  const alt=Object.keys(fields).find(k=>compactKey(k)===normalized||compactKey(k)===`${normalized}lookupid`);
  if(alt){
    const resolved=resolveLookupValue(fields[alt],preferredMap||state.userMap);
    return resolved||cleanLookupLabel(fields[alt]);
  }
  return '';
}

const PLACE_WORDS=['quito','cayambe','cotacachi','ibarra','manta','portoviejo','machala','guayaquil','cuenca','latacunga','salcedo','ambato','riobamba','macas','puyo','tena','archidona','loreto','otavalo','atuntaqui','machachi','sangolqui','tabacundo','pedro moncayo','esmeraldas','loja','zamora','el pangui','sucumbios','lago agrio','orellana','coca','manabi','imbabura','cotopaxi','tungurahua','chimborazo','napo','pastaza','pichincha','morona santiago'];
export function extractDestinationLabel(text){
  const raw=String(text||'').trim(); if(!raw)return 'Sin destino';
  const n=normalizeText(raw); const hits=[];
  for(const p of PLACE_WORDS) if(n.includes(normalizeText(p))) hits.push(p.replace(/\b\w/g,c=>c.toUpperCase()));
  return hits.length?[...new Set(hits)].slice(0,3).join(' · '):(raw.length>58?`${raw.slice(0,55)}…`:raw);
}

export function categorizeActivity(text){
  const n=normalizeText(text);
  const rules=[
    ['Seguimiento / visita técnica',['seguimiento','visita','campo','inspeccion','monitoreo']],
    ['Reunión / coordinación',['reunion','junta','coordinacion','mesa','asamblea']],
    ['Evento / intercambio',['evento','intercambio','taller','foro','seminario']],
    ['Firma / gestión institucional',['firma','convenio','carta compromiso','tramite','gestion']],
    ['Mantenimiento / soporte',['mantenimiento','reparacion','bateria','taller mecanico']],
    ['Entrevista / comunicación',['entrevista','grabacion','prensa','comunicacion']],
    ['Capacitación',['capacitacion','curso','induccion']],
    ['Logística / traslado',['traslado','movilizar','ruta','transporte']],
  ];
  for(const [label,keys] of rules) if(keys.some(k=>n.includes(normalizeText(k)))) return label;
  return 'Actividad institucional';
}

export function normalizeItems(items=state.items){
  return items.map(it=>{
    const f=it.fields||{};
    const start=toDate(fieldRaw(f,'start'));
    const end=toDate(fieldRaw(f,'end'));
    const requestDate=toDate(fieldRaw(f,'requestDate'));
    const kmStart=toNumber(fieldRaw(f,'kmStart'));
    const kmEnd=toNumber(fieldRaw(f,'kmEnd'));
    let distance=toNumber(fieldRaw(f,'distance'));
    if(!distance && kmEnd>=kmStart && kmEnd>0) distance=kmEnd-kmStart;
    const destination=textValue(fieldRaw(f,'destination'));
    const activity=textValue(fieldRaw(f,'activity'))||destination;
    const group=textValue(fieldRaw(f,'group'));
    const project=textValue(fieldRaw(f,'project'))||group;
    const vehicle=textValue(fieldRaw(f,'vehicle'))||textValue(fieldRaw(f,'plate'));
    return {
      id:String(it.id),
      start,end,requestDate,
      requester:textValue(fieldRaw(f,'requester')),
      group,
      destination,
      destinationLabel:extractDestinationLabel(destination),
      kmStart,kmEnd,distance,
      vehicle,
      plate:textValue(fieldRaw(f,'plate')),
      driver:textValue(fieldRaw(f,'driver')),
      activity,
      activityCategory:categorizeActivity(activity),
      project,
      status:textValue(fieldRaw(f,'status')),
      durationHours:start&&end?Math.max(0,hoursBetween(start,end)):0,
      leadHours:requestDate&&start?hoursBetween(requestDate,start):0,
      raw:f,
      webUrl:it.webUrl||''
    };
  }).sort((a,b)=>(a.start?.getTime()||0)-(b.start?.getTime()||0));
}

export function listDiagnostics(){
  return {
    site:state.site,
    activeList:state.activeList,
    columns:state.columns,
    mapping:state.mapping,
    totalItems:state.items.length,
    lists:state.lists
  };
}

export { SEMANTICS };
