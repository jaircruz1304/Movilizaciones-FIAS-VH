import { GPS_CONFIG } from '../config/msal-config.js';
import { downsample } from './gps.js';
import { escapeHtml, haversineKm } from './utils.js';

let map=null;
let layers=[];
let heatLayer=null;

function clearLayers(){
  for(const l of layers){try{map.removeLayer(l)}catch{}}
  layers=[];
  if(heatLayer){try{map.removeLayer(heatLayer)}catch{} heatLayer=null;}
}

export function initMap(){
  if(map) return map;
  if(!window.L) throw new Error('Leaflet no está disponible.');
  map=L.map('mainMap',{zoomControl:true,preferCanvas:true}).setView([GPS_CONFIG.origin.lat,GPS_CONFIG.origin.lon],7);
  L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png',{
    maxZoom:19,attribution:'&copy; OpenStreetMap contributors'
  }).addTo(map);
  const origin=L.circleMarker([GPS_CONFIG.origin.lat,GPS_CONFIG.origin.lon],{
    radius:9,weight:3,color:'#ffffff',fillColor:'#103f2d',fillOpacity:1
  }).bindPopup(`<strong>${escapeHtml(GPS_CONFIG.origin.name)}</strong><br>${escapeHtml(GPS_CONFIG.origin.note||'')}`).addTo(map);
  layers.push(origin);
  return map;
}

export function invalidateMap(){ if(map)setTimeout(()=>map.invalidateSize(),50); }

export function showGpsHeat(points){
  initMap();clearLayers();addOrigin();
  const pts=points.filter(p=>Number.isFinite(p.lat)&&Number.isFinite(p.lon));
  if(window.L.heatLayer && pts.length){
    const sample=downsample(pts,7000);
    heatLayer=L.heatLayer(sample.map(p=>[p.lat,p.lon,Math.max(.25,Math.min(1,(p.speed||0)/100+.25))]),{radius:18,blur:24,maxZoom:11}).addTo(map);
    fitPoints(sample);
  }
}

function addOrigin(){
  const l=L.circleMarker([GPS_CONFIG.origin.lat,GPS_CONFIG.origin.lon],{radius:9,weight:3,color:'#fff',fillColor:'#103f2d',fillOpacity:1})
    .bindPopup(`<strong>${escapeHtml(GPS_CONFIG.origin.name)}</strong>`).addTo(map);layers.push(l);
}

export function showCoverage(movements){
  initMap();clearLayers();addOrigin();
  const bounds=[];
  for(const m of movements){
    const p=m.gps?.remote;if(!p)continue;
    const marker=L.circleMarker([p.lat,p.lon],{radius:7,weight:2,color:'#fff',fillColor:'#2f7d5a',fillOpacity:.85})
      .bindPopup(`<strong>${escapeHtml(m.destinationLabel||m.destination||'Destino')}</strong><br>${escapeHtml(p.place||'')}<br>${(p.distanceFromOrigin||0).toFixed(1)} km desde Matriz FIAS`)
      .addTo(map);layers.push(marker);bounds.push([p.lat,p.lon]);
    const line=L.polyline([[GPS_CONFIG.origin.lat,GPS_CONFIG.origin.lon],[p.lat,p.lon]],{weight:1.4,opacity:.26,dashArray:'5 6'}).addTo(map);layers.push(line);
  }
  if(bounds.length)fitLatLngs(bounds); else map.setView([GPS_CONFIG.origin.lat,GPS_CONFIG.origin.lon],7);
}

export function showFrequentRoutes(movements){
  initMap();clearLayers();addOrigin();
  const all=[];
  const sorted=[...movements].filter(m=>m.gpsTrace?.length).sort((a,b)=>(b.gps?.odometerKm||0)-(a.gps?.odometerKm||0)).slice(0,30);
  for(const m of sorted){
    const pts=downsample(m.gpsTrace,450).map(p=>[p.lat,p.lon]);if(pts.length<2)continue;
    const line=L.polyline(pts,{weight:2.4,opacity:.42}).bindTooltip(`${m.destinationLabel||'Movilización'} · ${(m.gps?.odometerKm||0).toFixed(0)} km`).addTo(map);
    layers.push(line);all.push(...pts);
  }
  if(all.length)fitLatLngs(all);
}

export function showTrip(movement){
  initMap();clearLayers();addOrigin();
  const trace=downsample(movement?.gpsTrace||[],1500);
  if(!trace.length){map.setView([GPS_CONFIG.origin.lat,GPS_CONFIG.origin.lon],9);return;}
  const pts=trace.map(p=>[p.lat,p.lon]);
  const line=L.polyline(pts,{weight:4,opacity:.78}).addTo(map);layers.push(line);
  const start=trace[0],end=trace.at(-1),remote=movement.gps?.remote;
  layers.push(L.circleMarker([start.lat,start.lon],{radius:7,weight:2,color:'#fff',fillColor:'#2563eb',fillOpacity:1}).bindPopup(`Inicio GPS<br>${escapeHtml(start.place)}`).addTo(map));
  layers.push(L.circleMarker([end.lat,end.lon],{radius:7,weight:2,color:'#fff',fillColor:'#7c3aed',fillOpacity:1}).bindPopup(`Fin GPS<br>${escapeHtml(end.place)}`).addTo(map));
  if(remote)layers.push(L.circleMarker([remote.lat,remote.lon],{radius:8,weight:2,color:'#fff',fillColor:'#dc2626',fillOpacity:1}).bindPopup(`<strong>Punto más alejado</strong><br>${escapeHtml(remote.place)}<br>${remote.distanceFromOrigin.toFixed(1)} km desde FIAS`).addTo(map));
  fitLatLngs(pts);
}

export function showAllGpsTrace(points){
  initMap();clearLayers();addOrigin();
  const sample=downsample(points,3500); const pts=sample.map(p=>[p.lat,p.lon]);
  if(pts.length>1){const l=L.polyline(pts,{weight:2,opacity:.55}).addTo(map);layers.push(l);fitLatLngs(pts);}
}

function fitPoints(points){fitLatLngs(points.map(p=>[p.lat,p.lon]));}
function fitLatLngs(pts){if(!pts.length)return;const b=L.latLngBounds(pts);if(b.isValid())map.fitBounds(b.pad(.08),{maxZoom:13});}

export function nearestGpsPoint(points,target){
  let best=null,dist=Infinity;
  for(const p of points){const d=haversineKm(p,target);if(d<dist){dist=d;best=p;}}
  return best?{point:best,distanceKm:dist}:null;
}
