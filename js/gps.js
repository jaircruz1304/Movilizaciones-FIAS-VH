import { GPS_CONFIG, SHAREPOINT_GPS_CONFIG } from '../config/msal-config.js?v=2.3.0';
import { haversineKm, normalizeText, percentile } from './utils.js?v=2.3.0';
import { graph } from './graph.js?v=2.3.0';
import { resolveGpsSite, extractDestinationGeo } from './sharepoint.js?v=2.3.0';

const gpsState={manifest:null,points:[],trackers:new Map(),loaded:false,source:'SharePoint protegido'};
export { gpsState };

let gpsDrive=null;
function encodePath(path=''){return String(path).split('/').filter(Boolean).map(encodeURIComponent).join('/');}
async function getGpsReadDrive(){
  if(gpsDrive) return gpsDrive;

  const site = await resolveGpsSite();

  const data = await graph(
    `/sites/${encodeURIComponent(site.id)}/drives?$select=id,name,webUrl,driveType`
  );

  const drives = data?.value || [];

  if (!drives.length) {
    throw new Error(
      'No se encontraron bibliotecas de documentos accesibles para los datos GPS.'
    );
  }

  gpsDrive =
    drives.find(d => {
      const url = decodeURIComponent(String(d.webUrl || '')).toLowerCase();
      return url.includes('/documentos compartidos');
    }) ||
    drives.find(d =>
      ['documentos compartidos','documentos','documents','shared documents']
        .includes(String(d.name || '').trim().toLowerCase())
    ) ||
    drives[0];

  console.info('Biblioteca GPS de lectura:', {
    id: gpsDrive.id,
    name: gpsDrive.name,
    webUrl: gpsDrive.webUrl
  });

  return gpsDrive;
}
async function loadProtectedJson(filename){
  const drive=await getGpsReadDrive();
  const path=`${SHAREPOINT_GPS_CONFIG.publishedFolder}/${filename}`;
  try{
    const data=await graph(`/drives/${encodeURIComponent(drive.id)}/root:/${encodePath(path)}:/content`);
    return typeof data==='string'?JSON.parse(data):data;
  }catch(err){
    if(String(err?.message||err).includes('Microsoft Graph 404'))return null;
    throw err;
  }
}

function decodePoint(row,eventCodes){
  return {
    t:Number(row[0]),
    date:new Date(Number(row[0])*1000),
    lat:Number(row[1]),lon:Number(row[2]),
    speed:Number(row[3])||0,
    odometer:Number(row[4])||0,
    event:eventCodes[String(row[5])]||'OTRO',
    place:String(row[6]||'')
  };
}

export async function loadGpsData(force=false){
  if(gpsState.loaded && !force) return gpsState;
  if(force){gpsState.manifest=null;gpsState.points=[];gpsState.trackers=new Map();gpsState.loaded=false;}
  const manifest=await loadProtectedJson(SHAREPOINT_GPS_CONFIG.manifestName);
  if(!manifest){
    gpsState.manifest={version:2,files:[],eventCodes:{},totalPoints:0,updatedAt:null};
    gpsState.points=[];gpsState.trackers=new Map();gpsState.loaded=true;
    return gpsState;
  }
  gpsState.manifest=manifest;
  const all=[];
  for(const item of manifest.files||[]){
    const filename=item.file || String(item.url||'').split('/').pop();
    if(!filename)continue;
    const data=await loadProtectedJson(filename);
    if(!data)continue;
    const pts=(data.points||[]).map(row=>({...decodePoint(row,manifest.eventCodes||{}),tracker:data.tracker||item.tracker,month:item.month||''}));
    all.push(...pts);
    if(!gpsState.trackers.has(item.tracker))gpsState.trackers.set(item.tracker,[]);
    gpsState.trackers.get(item.tracker).push(...pts);
  }
  all.sort((a,b)=>a.t-b.t);
  for(const [,arr] of gpsState.trackers) arr.sort((a,b)=>a.t-b.t);
  gpsState.points=all; gpsState.loaded=true;
  return gpsState;
}

function lowerBound(points,t){let lo=0,hi=points.length;while(lo<hi){const m=(lo+hi)>>1;if(points[m].t<t)lo=m+1;else hi=m;}return lo;}
function upperBound(points,t){let lo=0,hi=points.length;while(lo<hi){const m=(lo+hi)>>1;if(points[m].t<=t)lo=m+1;else hi=m;}return lo;}

export function pointsInRange(start,end,tracker=''){
  const points=tracker?(gpsState.trackers.get(tracker)||[]):gpsState.points;
  if(!points.length||!start||!end)return [];
  const pad=(GPS_CONFIG.matchPaddingMinutes||0)*60;
  const a=Math.floor(start.getTime()/1000)-pad,b=Math.floor(end.getTime()/1000)+pad;
  return points.slice(lowerBound(points,a),upperBound(points,b));
}

export function trackerCompatible(movement,tracker){
  if(!movement.vehicle && !movement.plate) return true;
  const aliases=(GPS_CONFIG.trackerAliases?.[tracker]||[]).map(normalizeText);
  if(!aliases.length) return true;
  const hay=normalizeText(`${movement.vehicle||''} ${movement.plate||''}`);
  return aliases.some(a=>hay.includes(a));
}

export function pathDistance(points){
  let total=0;
  for(let i=1;i<points.length;i++){
    const a=points[i-1],b=points[i];
    const d=haversineKm(a,b);
    // descarta saltos espurios extremos entre registros consecutivos
    if(Number.isFinite(d)&&d<=25) total+=d;
  }
  return total;
}

const GEO_PROVINCES=[
  ['Azuay',['azuay']],['Bolívar',['bolivar']],['Cañar',['canar']],['Carchi',['carchi']],['Chimborazo',['chimborazo']],
  ['Cotopaxi',['cotopaxi']],['El Oro',['el oro']],['Esmeraldas',['esmeraldas']],['Guayas',['guayas']],['Imbabura',['imbabura']],
  ['Loja',['loja']],['Los Ríos',['los rios']],['Manabí',['manabi']],['Morona Santiago',['morona santiago']],['Napo',['napo']],
  ['Orellana',['orellana']],['Pastaza',['pastaza']],['Pichincha',['pichincha']],['Santa Elena',['santa elena']],
  ['Santo Domingo de los Tsáchilas',['santo domingo de los tsachilas','santo domingo']],['Sucumbíos',['sucumbios']],
  ['Tungurahua',['tungurahua']],['Zamora Chinchipe',['zamora chinchipe']]
].map(([label,aliases])=>({label,aliases}));

const GPS_CITIES=[
  ['Quito','Pichincha',['quito']],['Tababela','Pichincha',['tababela']],['Tumbaco','Pichincha',['tumbaco']],
  ['Cumbayá','Pichincha',['cumbaya']],['Puembo','Pichincha',['puembo']],['Pifo','Pichincha',['pifo']],
  ['Yaruquí','Pichincha',['yaruqui']],['El Quinche','Pichincha',['el quinche','quinche']],['Guayllabamba','Pichincha',['guayllabamba']],
  ['Cayambe','Pichincha',['cayambe']],['Machachi','Pichincha',['machachi']],['Sangolquí','Pichincha',['sangolqui','ruminahui']],
  ['Ibarra','Imbabura',['ibarra']],['Otavalo','Imbabura',['otavalo']],['Tulcán','Carchi',['tulcan']],
  ['Latacunga','Cotopaxi',['latacunga']],['Ambato','Tungurahua',['ambato']],['Baños','Tungurahua',['banos de agua santa','banos']],
  ['Riobamba','Chimborazo',['riobamba']],['Puyo','Pastaza',['puyo']],['Tena','Napo',['tena']],['Archidona','Napo',['archidona']],
  ['El Chaco','Napo',['el chaco']],['Nueva Loja','Sucumbíos',['nueva loja','lago agrio']],['Shushufindi','Sucumbíos',['shushufindi']],
  ['Puerto Francisco de Orellana','Orellana',['puerto francisco de orellana','el coca','coca']],['Loreto','Orellana',['loreto']],
  ['Macas','Morona Santiago',['macas']],['Sucúa','Morona Santiago',['sucua']],['Cuenca','Azuay',['cuenca']],['Azogues','Cañar',['azogues']],
  ['Loja','Loja',['loja']],['Zamora','Zamora Chinchipe',['zamora']],['El Pangui','Zamora Chinchipe',['el pangui','pangui']],
  ['Yantzaza','Zamora Chinchipe',['yantzaza']],['Guayaquil','Guayas',['guayaquil']],['Durán','Guayas',['duran']],
  ['Portoviejo','Manabí',['portoviejo']],['Manta','Manabí',['manta']],['Machala','El Oro',['machala']],['Esmeraldas','Esmeraldas',['esmeraldas']],
  ['Santo Domingo','Santo Domingo de los Tsáchilas',['santo domingo']]
].map(([label,province,aliases])=>({label,province,aliases}));

function compactGeo(value=''){
  return normalizeText(value).replace(/\s+/g,'');
}

function titleGeo(value=''){
  return String(value||'').toLowerCase().replace(/\b\p{L}/gu,c=>c.toUpperCase()).trim();
}

export function canonicalProvince(value=''){
  const key=compactGeo(value);
  if(!key)return '';
  for(const province of GEO_PROVINCES){
    if(province.aliases.some(a=>compactGeo(a)===key)) return province.label;
  }
  return '';
}

function cityDefinition(value=''){
  const key=normalizeText(value);
  if(!key)return null;
  return GPS_CITIES.find(city=>city.aliases.some(a=>normalizeText(a)===key))||null;
}

function inferProvinceFromCity(value=''){
  return cityDefinition(value)?.province||'';
}

function canonicalCity(value='',province=''){
  const raw=String(value||'').replace(/\s+/g,' ').trim();
  const def=cityDefinition(raw);
  if(def && (!province || normalizeText(def.province)===normalizeText(province))) return def.label;
  if(!raw || !province)return '';
  const n=normalizeText(raw);
  if(/\b(?:avenida|calle|via|ruta|autopista|carretera|camino)\b/.test(n))return '';
  if(/^[-+]?\d/.test(raw) || raw.length>70)return '';
  // Con provincia estructurada, el penúltimo componente del proveedor corresponde
  // normalmente a ciudad/cantón. Se conserva como ubicación útil aunque no esté en catálogo.
  return titleGeo(raw);
}

/**
 * Interpreta el campo CALLE del proveedor como una estructura jerárquica.
 * La provincia se toma exclusivamente del componente territorial final (o se
 * infiere por una ciudad reconocida), nunca de palabras incluidas en el nombre
 * de una vía. Así "Av. Francisco de Orellana, ..., Quito, Pichincha" no genera
 * la provincia Orellana.
 */
export function parseGpsPlace(place=''){
  const raw=String(place||'').replace(/\s+/g,' ').trim();
  const parts=raw.split(',').map(x=>x.trim()).filter(Boolean);
  let province=parts.length?canonicalProvince(parts.at(-1)):'';
  let cityRaw='';
  if(province && parts.length>=2) cityRaw=parts.at(-2);
  if(!province){
    // Algunos PDF dividen el nombre de la provincia entre líneas (PICHI NCHA,
    // TUNGURAHU A, PASTA ZA). canonicalProvince ya repara espacios internos.
    // Si el último componente quedó contaminado por un evento, la ciudad permite
    // recuperar la provincia sin utilizar el nombre de la calle.
    for(let i=parts.length-1;i>=Math.max(0,parts.length-3);i--){
      const inferred=inferProvinceFromCity(parts[i]);
      if(inferred){province=inferred;cityRaw=parts[i];break;}
    }
  }
  let city=canonicalCity(cityRaw,province);
  if(!city && province && parts.length>=2){
    const candidate=parts.at(-2);
    city=canonicalCity(candidate,province);
  }
  const provinceIndex=province?parts.length-1:-1;
  const cityIndex=cityRaw?parts.lastIndexOf(cityRaw):-1;
  const locality=(cityIndex>0?parts[cityIndex-1]:'')||'';
  const street=(cityIndex>0?parts.slice(0,Math.max(1,cityIndex-1)).join(', '):(parts[0]||''));
  return {raw,street,locality,city,province,label:city||province||''};
}

export function provincesFromPoints(points){
  const out=[];
  const seen=new Set();
  for(const p of points){
    const province=parseGpsPlace(p.place).province;
    if(!province)continue;
    const key=normalizeText(province);
    if(seen.has(key))continue;
    seen.add(key);out.push(province);
  }
  return out;
}

export function remotePoint(points){
  const origin=GPS_CONFIG.origin; let best=null,bestD=-1;
  for(const p of points){const d=haversineKm(origin,p);if(d>bestD){bestD=d;best=p;}}
  return best?{...best,distanceFromOrigin:bestD}:null;
}

function registeredDestinationEvidence(points,textGeo){
  if(!points?.length || textGeo?.kind!=='place' || !textGeo.label)return null;
  const key=normalizeText(textGeo.label);
  if(!key)return null;
  const expectedProvince=textGeo.province||'';
  const matches=[];
  for(const p of points){
    const geo=parseGpsPlace(p.place);
    if(expectedProvince && geo.province && !sameProvince(geo.province,expectedProvince))continue;
    const placeKey=normalizeText(p.place);
    const directCity=geo.city && normalizeText(geo.city)===key;
    const textualHit=placeKey.includes(key);
    if(directCity||textualHit)matches.push({...p,distanceFromOrigin:haversineKm(GPS_CONFIG.origin,p),geo});
  }
  // Dos o más puntos reducen el riesgo de validar un destino por una coincidencia aislada.
  if(matches.length<2)return null;
  const point=matches.reduce((best,p)=>!best||p.distanceFromOrigin>best.distanceFromOrigin?p:best,null);
  return {
    label:textGeo.label,province:textGeo.province||point?.geo?.province||'',count:matches.length,
    point,confidence:Math.min(1,.55+Math.log10(matches.length+1)/2),source:'GPS + destino registrado'
  };
}

function gpsDestinationEvidence(points){
  if(!points.length)return null;
  const origin=GPS_CONFIG.origin;
  const enriched=points.map(p=>{
    const geo=parseGpsPlace(p.place);
    return {p,geo,distance:haversineKm(origin,p)};
  }).filter(x=>x.geo.province);
  if(!enriched.length)return null;
  const maxRadius=Math.max(...enriched.map(x=>x.distance));
  const threshold=maxRadius<=5?maxRadius*.45:Math.max(5,maxRadius*.85);
  let candidates=enriched.filter(x=>x.distance>=threshold);
  if(candidates.length<3)candidates=enriched;
  const groups=new Map();
  for(const x of candidates){
    const label=x.geo.city||x.geo.province;
    const key=`${normalizeText(label)}|${normalizeText(x.geo.province)}`;
    if(!groups.has(key))groups.set(key,{label,city:x.geo.city,province:x.geo.province,count:0,score:0,maxDistance:0,point:null});
    const g=groups.get(key);
    const radiusWeight=maxRadius?x.distance/maxRadius:0;
    const dwellWeight=(Number(x.p.speed)||0)<=5?.35:0;
    g.count++;g.score+=1+radiusWeight*1.5+dwellWeight;
    if(x.distance>g.maxDistance){g.maxDistance=x.distance;g.point=x.p;}
  }
  const grouped=[...groups.values()];
  const eligible=grouped.some(x=>x.count>=3)?grouped.filter(x=>x.count>=3):grouped;
  // El destino se aproxima con el grupo territorial que alcanza el mayor radio
  // de forma repetida; el conteo evita que un único punto espurio defina el viaje.
  const ranked=eligible.sort((a,b)=>b.maxDistance-a.maxDistance||b.score-a.score||b.count-a.count);
  const best=ranked[0];
  if(!best)return null;
  const totalScore=ranked.reduce((n,x)=>n+x.score,0)||1;
  return {...best,confidence:best.score/totalScore,candidatePoints:candidates.length,maxRadiusKm:maxRadius};
}

export function gpsMetrics(points){
  if(!points.length)return null;
  const validOdo=points.filter(p=>p.odometer>0).map(p=>p.odometer);
  const odoMin=validOdo.length?Math.min(...validOdo):0;
  const odoMax=validOdo.length?Math.max(...validOdo):0;
  const speeds=points.map(p=>p.speed).filter(Number.isFinite);
  const events=new Map(); for(const p of points)events.set(p.event,(events.get(p.event)||0)+1);
  const remote=remotePoint(points);
  const provinces=provincesFromPoints(points);
  const destinationEvidence=gpsDestinationEvidence(points);
  return {
    points:points.length,
    start:points[0].date,end:points.at(-1).date,
    odometerKm:odoMax>=odoMin?odoMax-odoMin:0,
    pathKm:pathDistance(points),
    maxSpeed:speeds.length?Math.max(...speeds):0,
    p95Speed:speeds.length?percentile(speeds,.95):0,
    speedingEvents:events.get('INICIA EXCESO VELOCIDAD')||0,
    ignitionOn:events.get('VEHÍCULO ENCENDIDO')||0,
    ignitionOff:events.get('VEHÍCULO APAGADO')||0,
    events:[...events.entries()].sort((a,b)=>b[1]-a[1]),
    provinces,
    remote,
    destinationEvidence,
    maxRadiusKm:remote?.distanceFromOrigin||0,
    startPoint:points[0],endPoint:points.at(-1)
  };
}

function sameProvince(a='',b=''){
  return !!a && !!b && normalizeText(a)===normalizeText(b);
}

function routeHasProvince(metrics,province=''){
  return !!province && (metrics?.provinces||[]).some(p=>sameProvince(p,province));
}

function consolidateDestination(movement,metrics){
  const textGeo=extractDestinationGeo(movement.destination||'');
  const evidence=metrics?.destinationEvidence;
  const gpsGeo=evidence?{
    label:evidence.label,province:evidence.province,kind:evidence.city?'place':'province',confidence:evidence.confidence,source:'GPS'
  }:{label:'Por identificar',province:'',kind:'unknown',confidence:0,source:'none'};
  const hasGps=!!metrics;

  // 1) Un destino específico registrado conserva prioridad cuando la traza GPS
  // alcanza la provincia esperada. Esto permite, por ejemplo, mantener Tababela
  // aunque el proveedor clasifique sus coordenadas como QUITO,PICHINCHA.
  if(textGeo.kind==='place'){
    if(!hasGps)return {...textGeo,source:'SharePoint',validation:'Sin GPS'};
    if(metrics?.registeredDestinationEvidence){
      return {...textGeo,source:'SharePoint + GPS',validation:'Validado GPS'};
    }
    if(routeHasProvince(metrics,textGeo.province)){
      return {...textGeo,source:'SharePoint + GPS',validation:'Validado GPS territorial'};
    }
    // No se reemplaza silenciosamente un destino declarado por otra provincia:
    // se conserva y se marca la discrepancia para revisión humana.
    return {...textGeo,source:'SharePoint',validation:'Revisar destino'};
  }

  // 2) Si SharePoint solo registra una provincia, el GPS puede precisar la ciudad
  // siempre que ambas fuentes sean territorialmente coherentes.
  if(textGeo.kind==='province'){
    if(gpsGeo.kind==='place' && sameProvince(gpsGeo.province,textGeo.province)){
      return {...gpsGeo,source:'SharePoint + GPS',validation:'Validado GPS'};
    }
    if(hasGps && routeHasProvince(metrics,textGeo.province)){
      return {...textGeo,source:'SharePoint + GPS',validation:'Validado GPS'};
    }
    return {...textGeo,source:'SharePoint',validation:hasGps?'Revisar destino':'Sin GPS'};
  }

  // 3) Cuando el texto no contiene una ubicación, se utiliza la evidencia GPS
  // estructurada (ciudad/provincia) obtenida de las coordenadas y no de la calle.
  if(gpsGeo.kind!=='unknown'){
    return {...gpsGeo,source:'GPS',validation:'Inferido por GPS'};
  }
  const fallbackProvince=(metrics?.provinces||[]).at(-1)||'';
  if(fallbackProvince){
    return {label:fallbackProvince,province:fallbackProvince,kind:'province',confidence:1,source:'GPS',validation:'Inferido por GPS'};
  }
  return {label:'Por identificar',province:'',kind:'unknown',confidence:0,source:'none',validation:'Sin evidencia'};
}

export function reconcileMovements(movements){
  const trackers=[...gpsState.trackers.keys()];
  return movements.map(m=>{
    if(!m.start) return {...m,gps:null,gpsTrace:[],gpsStatus:'Sin fecha',destinationValidation:m.destinationLabel==='Por identificar'?'Sin evidencia':'Sin GPS'};
    const end=m.end || new Date(m.start.getTime()+24*3600000);
    let best=null;
    for(const tracker of trackers){
      if(!trackerCompatible(m,tracker)) continue;
      const pts=pointsInRange(m.start,end,tracker);
      if(!pts.length) continue;
      const metrics=gpsMetrics(pts);
      const score=pts.length + (metrics.odometerKm>0?100:0);
      if(!best||score>best.score)best={tracker,pts,metrics,score};
    }
    if(!best) return {...m,gps:null,gpsTrace:[],gpsStatus:'Sin coincidencia GPS',destinationValidation:m.destinationLabel==='Por identificar'?'Sin evidencia':'Sin GPS'};
    const spKm=Number(m.distance)||0;
    const gpsKm=best.metrics.odometerKm||0;
    const diff=spKm&&gpsKm?Math.abs(spKm-gpsKm):0;
    const diffPct=spKm&&gpsKm?diff/Math.max(spKm,gpsKm)*100:0;
    const agreement=!spKm||!gpsKm?'Referencia':diffPct<=10?'Alta':diffPct<=25?'Media':'Revisar';
    const declaredGeo=extractDestinationGeo(m.destination||'');
    const registeredEvidence=registeredDestinationEvidence(best.pts,declaredGeo);
    const metrics={...best.metrics,registeredDestinationEvidence:registeredEvidence};
    const destinationGeo=consolidateDestination(m,metrics);
    return {
      ...m,
      destinationLabel:destinationGeo.label,
      destinationProvince:destinationGeo.province,
      destinationKind:destinationGeo.kind,
      destinationSource:destinationGeo.source,
      destinationConfidence:destinationGeo.confidence,
      destinationValidation:destinationGeo.validation,
      gps:{tracker:best.tracker,...metrics,differenceKm:diff,differencePct:diffPct,agreement},
      gpsTrace:best.pts,
      gpsStatus:'Relacionado'
    };
  });
}

export function downsample(points,max=1200){
  if(points.length<=max)return points;
  const step=Math.ceil(points.length/max); const out=[];
  for(let i=0;i<points.length;i+=step)out.push(points[i]);
  if(out.at(-1)!==points.at(-1))out.push(points.at(-1));
  return out;
}

export function gpsSummary(){
  const p=gpsState.points;
  if(!p.length)return {points:0,trackers:0,months:0,events:[],maxSpeed:0,odometerKm:0};
  const events=new Map();for(const x of p)events.set(x.event,(events.get(x.event)||0)+1);
  const odo=p.filter(x=>x.odometer>0).map(x=>x.odometer);
  return {
    points:p.length,
    trackers:gpsState.trackers.size,
    months:new Set(p.map(x=>x.month)).size,
    start:p[0].date,end:p.at(-1).date,
    maxSpeed:Math.max(...p.map(x=>x.speed)),
    odometerKm:odo.length?Math.max(...odo)-Math.min(...odo):0,
    events:[...events.entries()].sort((a,b)=>b[1]-a[1]),
    provinces:provincesFromPoints(p)
  };
}
