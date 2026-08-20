import { groupCounts, sum, average, iqrOutlierThreshold, formatDateKey } from './utils.js';

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
      const hay=[r.requester,r.group,r.destination,r.destinationLabel,r.vehicle,r.plate,r.driver,r.activity,r.activityCategory,r.project,r.gps?.remote?.place,r.gps?.provinces?.join(' ')].join(' ').toLowerCase();
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
  const destinations=groupCounts(rows,r=>r.destinationLabel);
  const projects=groupCounts(rows,r=>r.project||r.group);
  return {
    movements:rows.length,
    sharepointKm:spKm,
    gpsKm,
    gpsMatchRate:rows.length?matched.length/rows.length*100:0,
    avgDuration:average(durations,r=>r.durationHours),
    avgLeadHours:average(leads,r=>r.leadHours),
    topDestination:destinations[0]?.[0]||'—',
    topDestinationCount:destinations[0]?.[1]||0,
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

export function provinceCoverage(rows){
  const map=new Map();
  for(const r of rows){for(const p of r.gps?.provinces||[])map.set(p,(map.get(p)||0)+1);}
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
    if((r.gps?.maxSpeed||0)>=120)issues.push('Velocidad GPS máxima elevada');
    if(!r.destination)issues.push('Destino no informado');
    return {...r,issues};
  }).filter(r=>r.issues.length).sort((a,b)=>b.issues.length-a.issues.length);
}

export function dataQuality(rows){
  const total=rows.length||1;
  const fields=[
    ['Fecha inicio',r=>!!r.start],['Fecha fin',r=>!!r.end],['Solicitante',r=>!!r.requester],['Grupo/proyecto',r=>!!r.group||!!r.project],
    ['Destino',r=>!!r.destination],['Kilometraje',r=>r.distance>0],['Vehículo/placa',r=>!!r.vehicle||!!r.plate],['Conductor',r=>!!r.driver],['GPS relacionado',r=>!!r.gps]
  ];
  return fields.map(([label,test])=>{const n=rows.filter(test).length;return{label,count:n,pct:n/total*100}});
}

export function routeClusters(rows){
  const clusters=new Map();
  for(const r of rows){
    const p=r.gps?.remote; if(!p)continue;
    const lat=Math.round(p.lat/0.12)*0.12, lon=Math.round(p.lon/0.12)*0.12;
    const province=(r.gps?.provinces||[])[0]||'';
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
  const K=kpis(rows),dest=top(rows,'destinationLabel',3),act=top(rows,'activityCategory',3),grp=top(rows,r=>r.project||r.group,3),prov=provinceCoverage(rows);
  const insights=[];
  if(dest[0]) insights.push(`${dest[0][0]} concentra ${Math.round(dest[0][1]/rows.length*100)}% de las movilizaciones del período analizado.`);
  if(act[0]) insights.push(`La finalidad predominante es “${act[0][0]}”, con ${act[0][1]} registros.`);
  if(grp[0]) insights.push(`${grp[0][0]} es el grupo/proyecto con mayor demanda vehicular (${grp[0][1]} movilizaciones).`);
  if(K.gpsMatchRate>0) insights.push(`El ${K.gpsMatchRate.toFixed(0)}% de los registros filtrados tiene evidencia GPS temporalmente relacionada.`);
  if(prov.length) insights.push(`La cobertura GPS alcanza ${prov.length} provincias; ${prov[0][0]} aparece con mayor frecuencia en las trazas relacionadas.`);
  if(K.speedEvents) insights.push(`Se identifican ${K.speedEvents} inicios de eventos de exceso de velocidad dentro de movilizaciones relacionadas.`);
  const avgKm=rows.length?K.sharepointKm/rows.length:0;if(avgKm)insights.push(`El recorrido medio registrado en SharePoint es de ${avgKm.toFixed(1)} km por movilización.`);
  return insights;
}
