import { GPS_CONFIG } from '../config/msal-config.js';
import { haversineKm, normalizeText, percentile } from './utils.js';

const gpsState={manifest:null,points:[],trackers:new Map(),loaded:false};
export { gpsState };

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

export async function loadGpsData(){
  if(gpsState.loaded) return gpsState;
  const manifest=await fetch(GPS_CONFIG.manifestUrl).then(r=>{if(!r.ok)throw new Error(`No se pudo cargar ${GPS_CONFIG.manifestUrl}`);return r.json();});
  gpsState.manifest=manifest;
  const all=[];
  for(const item of manifest.files||[]){
    const data=await fetch(item.url).then(r=>{if(!r.ok)throw new Error(`No se pudo cargar ${item.url}`);return r.json();});
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

const PROVINCES=['PICHINCHA','IMBABURA','COTOPAXI','TUNGURAHUA','CHIMBORAZO','PASTAZA','NAPO','ORELLANA','SUCUMBIOS','SUCUMBÍOS','MANABI','MANABÍ','MORONA SANTIAGO','ESMERALDAS','GUAYAS','AZUAY','LOJA','ZAMORA CHINCHIPE','SANTO DOMINGO','CARCHI'];
export function provincesFromPoints(points){
  const found=new Map();
  for(const p of points){
    const u=p.place.toUpperCase();
    for(const province of PROVINCES){
      if(u.includes(province)){const key=province.normalize('NFD').replace(/[\u0300-\u036f]/g,'');found.set(key,(found.get(key)||0)+1);}
    }
  }
  return [...found.entries()].sort((a,b)=>b[1]-a[1]).map(x=>x[0]);
}

export function remotePoint(points){
  const origin=GPS_CONFIG.origin; let best=null,bestD=-1;
  for(const p of points){const d=haversineKm(origin,p);if(d>bestD){bestD=d;best=p;}}
  return best?{...best,distanceFromOrigin:bestD}:null;
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
    maxRadiusKm:remote?.distanceFromOrigin||0,
    startPoint:points[0],endPoint:points.at(-1)
  };
}

export function reconcileMovements(movements){
  const trackers=[...gpsState.trackers.keys()];
  return movements.map(m=>{
    if(!m.start) return {...m,gps:null,gpsTrace:[],gpsStatus:'Sin fecha'};
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
    if(!best) return {...m,gps:null,gpsTrace:[],gpsStatus:'Sin coincidencia GPS'};
    const spKm=Number(m.distance)||0;
    const gpsKm=best.metrics.odometerKm||0;
    const diff=spKm&&gpsKm?Math.abs(spKm-gpsKm):0;
    const diffPct=spKm&&gpsKm?diff/Math.max(spKm,gpsKm)*100:0;
    const agreement=!spKm||!gpsKm?'Referencia':diffPct<=10?'Alta':diffPct<=25?'Media':'Revisar';
    return {
      ...m,
      gps:{tracker:best.tracker,...best.metrics,differenceKm:diff,differencePct:diffPct,agreement},
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
