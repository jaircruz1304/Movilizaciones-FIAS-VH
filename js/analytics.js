import { groupCounts, sum, average, iqrOutlierThreshold, formatDateKey } from './utils.js?v=2.3.0';

export function filterMovements(rows,filters={}){
  const q=(filters.search||'').trim().toLowerCase();
  return rows.filter(r=>{
    if(filters.from && (!r.start || formatDateKey(r.start)<filters.from))return false;
    if(filters.to && (!r.start || formatDateKey(r.start)>filters.to))return false;
    if(filters.group && r.group!==filters.group)return false;
    if(filters.requester && r.requester!==filters.requester)return false;
    if(filters.vehicle && (r.vehicle||r.plate)!==filters.vehicle)return false;
    if(filters.activity && r.activityCategory!==filters.activity)return false;
    if(q){
      const hay=[r.requester,r.group,r.destination,r.destinationLabel,r.destinationProvince,r.vehicle,r.plate,r.activity,r.activityCategory,r.project,r.gps?.remote?.place,r.gps?.provinces?.join(' ')].join(' ').toLowerCase();
      if(!hay.includes(q))return false;
    }
    return true;
  });
}

export function kpis(rows){
  const matched=rows.filter(r=>r.gps);
  const spKm=sum(rows,r=>r.distance);
  const gpsKm=sum(matched,r=>r.gps?.odometerKm||0);
  const durations=rows.filter(r=>r.durationHours>0);
  const leads=rows.filter(r=>Number.isFinite(r.leadHours)&&r.leadHours!==0);
  const destinations=destinationSummary(rows,Math.max(10,rows.length)).filter(x=>x.mapped);
  const projects=groupCounts(rows,r=>r.project||r.group);
  return {
    movements:rows.length,
    sharepointKm:spKm,
    gpsKm,
    gpsMatchRate:rows.length?matched.length/rows.length*100:0,
    totalUseHours:sum(durations,r=>r.durationHours),
    avgDuration:average(durations,r=>r.durationHours),
    avgLeadHours:average(leads,r=>r.leadHours),
    topDestination:destinations[0]?.label||'—',
    topDestinationCount:destinations[0]?.count||0,
    topProject:projects[0]?.[0]||'—',
    topProjectCount:projects[0]?.[1]||0,
    provinces:new Set(matched.flatMap(r=>r.gps?.provinces||[])).size,
    speedEvents:sum(matched,r=>r.gps?.speedingEvents||0)
  };
}

export function monthlyTrend(rows){
  const m=new Map();
  for(const r of rows){if(!r.start)continue;const k=formatDateKey(r.start).slice(0,7);if(!m.has(k))m.set(k,{movements:0,km:0,gpsKm:0});const x=m.get(k);x.movements++;x.km+=r.distance||0;x.gpsKm+=r.gps?.odometerKm||0;}
  return [...m.entries()].sort((a,b)=>a[0].localeCompare(b[0])).map(([month,v])=>({month,...v}));
}

export function weekdayDemand(rows){
  const names=['Dom','Lun','Mar','Mié','Jue','Vie','Sáb'],vals=Array(7).fill(0);
  for(const r of rows)if(r.start)vals[r.start.getDay()]++;
  return [1,2,3,4,5,6,0].map(i=>({label:names[i],value:vals[i]}));
}

export function top(rows,key,n=10){
  return groupCounts(rows,r=>typeof key==='function'?key(r):r[key]).slice(0,n);
}

export function destinationSummary(rows,n=10){
  const map=new Map();
  for(const r of rows){
    const label=r.destinationLabel||'Por identificar';
    if(!map.has(label))map.set(label,{label,count:0,gps:0,provinces:new Map(),mapped:label!=='Por identificar'});
    const x=map.get(label);x.count++;if(r.gps)x.gps++;
    const province=r.destinationProvince||'';
    if(province)x.provinces.set(province,(x.provinces.get(province)||0)+1);
  }
  return [...map.values()].map(x=>({
    label:x.label,count:x.count,gps:x.gps,mapped:x.mapped,
    province:[...x.provinces.entries()].sort((a,b)=>b[1]-a[1])[0]?.[0]||''
  })).sort((a,b)=>(Number(b.mapped)-Number(a.mapped))||(b.count-a.count)||a.label.localeCompare(b.label,'es')).slice(0,n);
}

export function territorialScope(rows){
  const counts=new Map([
    ['Quito y entorno',0],['Resto de Pichincha',0],['Fuera de Pichincha',0],['Sin ubicación suficiente',0]
  ]);
  for(const r of rows){
    const province=String(r.destinationProvince||'');
    const label=String(r.destinationLabel||'');
    const gpsProvinces=(r.gps?.provinces||[]).map(String);
    const hasOutside=gpsProvinces.some(p=>normalizeProvince(p)!=='pichincha');
    const radius=Number(r.gps?.maxRadiusKm)||0;
    let bucket='Sin ubicación suficiente';
    if(label==='Por identificar'&&!province&&!gpsProvinces.length)bucket='Sin ubicación suficiente';
    else if(hasOutside || (province && normalizeProvince(province)!=='pichincha'))bucket='Fuera de Pichincha';
    else if(label==='Quito' || (normalizeProvince(province)==='pichincha' && radius>0 && radius<=35))bucket='Quito y entorno';
    else if(normalizeProvince(province)==='pichincha' || gpsProvinces.some(p=>normalizeProvince(p)==='pichincha'))bucket='Resto de Pichincha';
    counts.set(bucket,(counts.get(bucket)||0)+1);
  }
  return [...counts.entries()].map(([label,value])=>({label,value}));
}

function normalizeProvince(v){
  return String(v||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().trim();
}

export function provinceCoverage(rows){
  const map=new Map();
  for(const r of rows){
    const destinations=r.destinationProvince?[r.destinationProvince]:(r.gps?.provinces||[]);
    for(const p of destinations){if(!p)continue;map.set(p,(map.get(p)||0)+1);}
  }
  return [...map.entries()].sort((a,b)=>b[1]-a[1]);
}

export function anomalies(rows){
  const distances=rows.map(r=>r.distance).filter(x=>x>0);
  const durations=rows.map(r=>r.durationHours).filter(x=>x>0);
  const dThr=iqrOutlierThreshold(distances),hThr=iqrOutlierThreshold(durations);
  return rows.map(r=>{
    const issues=[];
    if(r.start&&r.end&&r.end<r.start)issues.push('Fecha final anterior al inicio');
    if(r.kmEnd&&r.kmStart&&r.kmEnd<r.kmStart)issues.push('KM final menor al inicial');
    if(r.requestDate&&r.start&&r.requestDate>r.start)issues.push('Solicitud posterior al inicio');
    if(r.distance>0&&distances.length>=4&&r.distance>dThr.high)issues.push('Recorrido SharePoint atípicamente alto');
    if(r.durationHours>0&&durations.length>=4&&r.durationHours>hThr.high)issues.push('Duración atípicamente alta');
    if(r.gps?.agreement==='Revisar')issues.push('Diferencia relevante entre recorrido SharePoint y GPS');
    if(r.destinationValidation==='Revisar destino')issues.push('Destino registrado no coincide con la cobertura GPS');
    if((r.gps?.maxSpeed||0)>=120)issues.push('Velocidad GPS máxima elevada');
    if(!r.destination)issues.push('Destino no informado');
    else if(r.destinationLabel==='Por identificar')issues.push('Destino geográfico no identificado');
    return {...r,issues};
  }).filter(r=>r.issues.length).sort((a,b)=>b.issues.length-a.issues.length);
}

export function dataQuality(rows){
  const total=rows.length||1;
  const fields=[
    ['Fecha inicio',r=>!!r.start],['Fecha fin',r=>!!r.end],['Solicitante',r=>!!r.requester],['Grupo/proyecto',r=>!!r.group||!!r.project],
    ['Destino registrado',r=>!!r.destination],['Destino geográfico',r=>!!r.destinationLabel&&r.destinationLabel!=='Por identificar'],['Kilometraje',r=>r.distance>0],['Vehículo/placa',r=>!!r.vehicle||!!r.plate],['GPS relacionado',r=>!!r.gps]
  ];
  return fields.map(([label,test])=>{const n=rows.filter(test).length;return{label,count:n,pct:n/total*100}});
}

export function routeClusters(rows){
  const clusters=new Map();
  for(const r of rows){
    const p=r.gps?.remote; if(!p)continue;
    const lat=Math.round(p.lat/0.12)*0.12, lon=Math.round(p.lon/0.12)*0.12;
    const gpsProvinces=r.gps?.provinces||[];
    const province=r.destinationProvince||gpsProvinces.find(p=>normalizeProvince(p)!=='pichincha')||gpsProvinces[0]||'';
    const key=`${lat.toFixed(2)}|${lon.toFixed(2)}|${province}`;
    if(!clusters.has(key))clusters.set(key,{count:0,labels:new Map(),km:0,lat:0,lon:0,province});
    const c=clusters.get(key);c.count++;c.km+=p.distanceFromOrigin||0;c.lat+=p.lat;c.lon+=p.lon;
    const label=r.destinationLabel||p.place||province||'Corredor';c.labels.set(label,(c.labels.get(label)||0)+1);
  }
  return [...clusters.values()].map(c=>{
    const label=[...c.labels.entries()].sort((a,b)=>b[1]-a[1])[0]?.[0]||c.province||'Corredor';
    return {...c,label,avgKm:c.count?c.km/c.count:0,lat:c.lat/c.count,lon:c.lon/c.count};
  }).sort((a,b)=>b.count-a.count);
}

export function executiveInsights(rows){
  if(!rows.length)return ['No existen movilizaciones para los filtros actuales.'];
  const K=kpis(rows),dest=destinationSummary(rows,4).filter(x=>x.mapped),grp=top(rows,r=>r.project||r.group,3),prov=provinceCoverage(rows),scope=territorialScope(rows);
  const insights=[];
  const mappedCount=rows.filter(r=>r.destinationLabel&&r.destinationLabel!=='Por identificar').length;
  if(dest[0]) insights.push(`${dest[0].label}${dest[0].province?` (${dest[0].province})`:''} concentra ${Math.round(dest[0].count/rows.length*100)}% de las movilizaciones del período analizado.`);
  const outside=scope.find(x=>x.label==='Fuera de Pichincha')?.value||0;
  if(outside) insights.push(`${outside} movilizaciones (${Math.round(outside/rows.length*100)}%) presentan alcance fuera de Pichincha según el destino consolidado y/o la evidencia GPS.`);
  if(grp[0]) insights.push(`${grp[0][0]} es el grupo/proyecto con mayor demanda vehicular (${grp[0][1]} movilizaciones).`);
  if(mappedCount<rows.length) insights.push(`Se logró consolidar geográficamente ${mappedCount} de ${rows.length} destinos; ${rows.length-mappedCount} registros requieren una descripción territorial más precisa.`);
  if(K.gpsMatchRate>0) insights.push(`El ${K.gpsMatchRate.toFixed(0)}% de los registros filtrados tiene evidencia GPS temporalmente relacionada.`);
  if(prov.length) insights.push(`La cobertura GPS alcanza ${prov.length} provincias; ${prov[0][0]} aparece con mayor frecuencia en las trazas relacionadas.`);
  if(K.speedEvents) insights.push(`Se identifican ${K.speedEvents} inicios de eventos de exceso de velocidad dentro de movilizaciones relacionadas.`);
  const avgKm=rows.length?K.sharepointKm/rows.length:0;if(avgKm)insights.push(`El recorrido medio registrado en SharePoint es de ${avgKm.toFixed(1)} km por movilización.`);
  return insights;
}
