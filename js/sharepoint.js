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

const state={site:null,lists:[],activeList:null,columns:[],items:[],mapping:{},userMap:new Map()};
export const sharepointState=state;

export async function resolveSite(){
  if(state.site) return state.site;
  const path=SHAREPOINT_CONFIG.sitePath.replace(/^\//,'');
  state.site=await graph(`/sites/${SHAREPOINT_CONFIG.host}:/${path}?$select=id,displayName,webUrl`);
  return state.site;
}

function semanticScore(columns,title=''){
  const names=columns.flatMap(c=>[compactKey(c.name),compactKey(c.displayName)]);
  let score=0;
  for(const [key,meta] of Object.entries(SEMANTICS)){
    if(meta.aliases.some(a=>names.some(n=>n===compactKey(a)||n.includes(compactKey(a))||compactKey(a).includes(n)))) score += ['start','destination','distance'].includes(key)?5:2;
  }
  if(/moviliz|vehicul|transporte|salida|comision|recorrido/i.test(normalizeText(title))) score+=10;
  return score;
}

export async function discoverLists(){
  const site=await resolveSite();
  const path=`/sites/${encodeURIComponent(site.id)}/lists?$select=id,name,displayName,webUrl,list&$expand=columns($select=id,name,displayName,hidden,readOnly)&$top=200`;
  const data=await graph(path);
  const all=(data?.value||[]).map(l=>({...l,score:semanticScore(l.columns||[],l.displayName||l.name)}));
  state.lists=all.sort((a,b)=>b.score-a.score);
  return state.lists;
}

export async function chooseList(id=''){
  if(!state.lists.length) await discoverLists();
  const remembered=localStorage.getItem('fias.movilizaciones.listId')||'';
  const wanted=id||SHAREPOINT_CONFIG.preferredListId||remembered;
  let list=state.lists.find(x=>x.id===wanted);
  if(!list) list=state.lists.filter(x=>!(x.list?.hidden))[0];
  if(!list) throw new Error('No se encontró una lista de SharePoint accesible para movilizaciones.');
  state.activeList=list;
  state.columns=(list.columns||[]).filter(c=>!c.hidden);
  localStorage.setItem('fias.movilizaciones.listId',list.id);
  state.mapping=detectMapping(state.columns);
  await loadUserMap().catch(err=>console.warn('No se pudo resolver User Information List',err));
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

async function loadUserMap(){
  if(!state.site) return;
  const userList=state.lists.find(l=>/user information list|informaci[oó]n del usuario/i.test(l.displayName||l.name||''));
  if(!userList) return;
  const items=await graphPaged(`/sites/${encodeURIComponent(state.site.id)}/lists/${encodeURIComponent(userList.id)}/items?$expand=fields&$top=999`,5000);
  state.userMap=new Map(items.map(it=>[String(it.id),it.fields?.Title||it.fields?.Name||it.fields?.EMail||`Usuario ${it.id}`]));
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
  if(fields[name]!==undefined) return fields[name];
  const lookup=fields[`${name}LookupId`];
  if(lookup!==undefined){
    if(Array.isArray(lookup)) return lookup.map(x=>state.userMap.get(String(x))||x);
    return state.userMap.get(String(lookup))||lookup;
  }
  const normalized=compactKey(name);
  const alt=Object.keys(fields).find(k=>compactKey(k)===normalized||compactKey(k)===`${normalized}lookupid`);
  return alt?fields[alt]:'';
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
