import { $, $$, escapeHtml, fmtInt, fmtKm, fmtPct, fmt1, formatDateTime, formatDate, groupCounts, csvEscape, downloadText, safeJson } from './utils.js?v=1.4.0';
import { kpis, monthlyTrend, top, weekdayDemand, provinceCoverage, routeClusters, anomalies, dataQuality, executiveInsights } from './analytics.js?v=1.4.0';
import { gpsSummary, gpsState } from './gps.js?v=1.4.0';
import { sharepointState, SEMANTICS } from './sharepoint.js?v=1.4.0';
import { authDiagnostics } from './auth.js?v=1.4.0';
import { showTrip } from './maps.js?v=1.4.0';

const charts={};
let currentRows=[];
let allRows=[];
let onOpenTrip=null;
let onApplyMapping=null;
let onChooseList=null;

export function bindDashboardCallbacks(cb={}){onOpenTrip=cb.onOpenTrip||null;onApplyMapping=cb.onApplyMapping||null;onChooseList=cb.onChooseList||null;}
export function setRows(filtered,all){currentRows=filtered;allRows=all||filtered;}

function chart(id,type,labels,datasets,opts={}){
  if(charts[id])charts[id].destroy(); const canvas=$(id);if(!canvas||!window.Chart)return;
  const common={responsive:true,maintainAspectRatio:false,interaction:{mode:'index',intersect:false},plugins:{legend:{display:opts.legend!==false,position:'bottom'}},scales:type==='doughnut'?{}:{x:{grid:{display:false}},y:{beginAtZero:true,grid:{color:'#eef2ef'}}}};
  charts[id]=new Chart(canvas,{type,data:{labels,datasets},options:{...common,...(opts.options||{})}});
}
function rankHtml(items,subFn=(x)=>`${x[1]} registros`,valueFn=(x)=>fmtInt(x[1])){
  if(!items.length)return '<div class="empty">Sin datos para el filtro actual.</div>';
  return items.map((x,i)=>`<div class="rank-row"><div class="rank-no">${i+1}</div><div class="rank-main"><strong title="${escapeHtml(x[0])}">${escapeHtml(x[0])}</strong><small>${escapeHtml(subFn(x))}</small></div><div class="rank-val">${escapeHtml(valueFn(x))}</div></div>`).join('');
}
function kpiCard(label,value,sub,icon){return `<article class="kpi-card"><div><div class="kpi-icon"><i class="bi ${icon}"></i></div><span class="kpi-label">${escapeHtml(label)}</span><strong>${escapeHtml(value)}</strong></div><small>${escapeHtml(sub)}</small></article>`;}

export function renderOverview(rows=currentRows){
  const K=kpis(rows);
  $('kpiGrid').innerHTML=[
    kpiCard('Movilizaciones',fmtInt(K.movements),'Registros en el período filtrado','bi-signpost-2'),
    kpiCard('Km registrados',fmtKm(K.sharepointKm),'Recorrido declarado en SharePoint','bi-speedometer2'),
    kpiCard('Conciliación GPS',fmtPct(K.gpsMatchRate),'Registros relacionados temporalmente','bi-broadcast-pin'),
    kpiCard('Cobertura',`${fmtInt(K.provinces)} provincias`,'Según trazas GPS relacionadas','bi-map'),
    kpiCard('Destino recurrente',K.topDestination,`${fmtInt(K.topDestinationCount)} movilizaciones`,'bi-geo-alt'),
    kpiCard('Exceso velocidad',fmtInt(K.speedEvents),'Eventos GPS dentro de movilizaciones','bi-exclamation-triangle')
  ].join('');
  const trend=monthlyTrend(rows);
  chart('chartTrend','bar',trend.map(x=>x.month),[
    {label:'Movilizaciones',data:trend.map(x=>x.movements),borderWidth:0,borderRadius:6,yAxisID:'y'},
    {label:'Km SharePoint',data:trend.map(x=>Math.round(x.km)),type:'line',borderWidth:2,tension:.3,pointRadius:3,yAxisID:'y1'}
  ],{options:{scales:{x:{grid:{display:false}},y:{beginAtZero:true,title:{display:true,text:'Movilizaciones'}},y1:{beginAtZero:true,position:'right',grid:{drawOnChartArea:false},title:{display:true,text:'Km'}}}}});
  $('destinationRanking').innerHTML=rankHtml(top(rows,'destinationLabel',8));
  $('insightList').innerHTML=executiveInsights(rows).map(x=>`<div class="insight"><i class="bi bi-stars"></i> ${escapeHtml(x)}</div>`).join('');
  const acts=top(rows,'activityCategory',8);
  chart('chartActivities','doughnut',acts.map(x=>x[0]),[{label:'Movilizaciones',data:acts.map(x=>x[1]),borderWidth:0}],{legend:true});
}

export function renderTerritory(rows=currentRows){
  const prov=provinceCoverage(rows).slice(0,12);
  $('provinceRanking').innerHTML=rankHtml(prov,x=>`${x[1]} movilizaciones relacionadas`);
  const routes=routeClusters(rows).slice(0,10);
  $('remoteRanking').innerHTML=routes.length?routes.map((r,i)=>`<div class="rank-row"><div class="rank-no">${i+1}</div><div class="rank-main"><strong>${escapeHtml(r.label)}</strong><small>${escapeHtml(r.province||'Cobertura GPS')} · radio medio ${fmt1(r.avgKm)} km</small></div><div class="rank-val">${fmtInt(r.count)}</div></div>`).join(''):'<div class="empty">Sin rutas GPS relacionadas.</div>';
}

function pill(text,type='neutral'){return `<span class="pill ${type}">${escapeHtml(text)}</span>`;}
export function renderOperations(rows=currentRows){
  const body=$('movementRows');
  body.innerHTML=rows.length?rows.slice().sort((a,b)=>(b.start?.getTime()||0)-(a.start?.getTime()||0)).map(r=>{
    const agreement=r.gps?.agreement||'Sin GPS';
    const cls=agreement==='Alta'?'ok':agreement==='Media'||agreement==='Referencia'?'warn':agreement==='Revisar'?'bad':'neutral';
    return `<tr data-id="${escapeHtml(r.id)}"><td>${escapeHtml(formatDateTime(r.start))}<br><span class="muted">${escapeHtml(formatDateTime(r.end))}</span></td><td>${escapeHtml(r.requester||'—')}</td><td>${escapeHtml(r.project||r.group||'—')}</td><td><strong>${escapeHtml(r.destinationLabel||'—')}</strong><br><span class="muted">${escapeHtml((r.destination||'').slice(0,95))}</span></td><td>${r.distance?fmtKm(r.distance):'—'}</td><td>${r.gps?pill('Relacionado','ok'):pill('Sin relación','neutral')}</td><td>${r.gps?.odometerKm?fmtKm(r.gps.odometerKm):'—'}</td><td>${escapeHtml((r.gps?.provinces||[]).slice(0,3).join(' · ')||'—')}</td><td>${pill(agreement,cls)}</td></tr>`;
  }).join(''):'<tr><td colspan="9"><div class="empty">No existen movilizaciones para los filtros aplicados.</div></td></tr>';
  body.querySelectorAll('tr[data-id]').forEach(tr=>tr.addEventListener('click',()=>openDetail(rows.find(r=>r.id===tr.dataset.id))));
}

export function renderFleet(rows=currentRows){
  const groups=top(rows,r=>r.project||r.group,12),users=top(rows,'requester',12),vehicles=top(rows,r=>r.vehicle||r.plate,12),week=weekdayDemand(rows);
  chart('chartGroups','bar',groups.map(x=>x[0]),[{label:'Movilizaciones',data:groups.map(x=>x[1]),borderWidth:0,borderRadius:5}],{legend:false,options:{indexAxis:'y',scales:{x:{beginAtZero:true},y:{grid:{display:false}}}}});
  chart('chartRequesters','bar',users.map(x=>x[0]),[{label:'Movilizaciones',data:users.map(x=>x[1]),borderWidth:0,borderRadius:5}],{legend:false,options:{indexAxis:'y',scales:{x:{beginAtZero:true},y:{grid:{display:false}}}}});
  chart('chartVehicles','bar',vehicles.map(x=>x[0]),[{label:'Movilizaciones',data:vehicles.map(x=>x[1]),borderWidth:0,borderRadius:5}],{legend:false});
  chart('chartWeekday','bar',week.map(x=>x.label),[{label:'Movilizaciones',data:week.map(x=>x.value),borderWidth:0,borderRadius:5}],{legend:false});
}

export function renderGps(rows=currentRows){
  const S=gpsSummary();
  $('gpsKpis').innerHTML=[
    kpiCard('Puntos GPS',fmtInt(S.points),`${fmtInt(S.months)} meses incorporados`,'bi-crosshair'),
    kpiCard('Rastreador',fmtInt(S.trackers),'Identificador actual: PDF-8770','bi-router'),
    kpiCard('Odómetro histórico',fmtKm(S.odometerKm),'Diferencia global de lecturas válidas','bi-speedometer'),
    kpiCard('Velocidad máxima',`${fmtInt(S.maxSpeed)} km/h`,'Máximo observado en reportes','bi-lightning'),
    kpiCard('Cobertura GPS',`${fmtInt(S.provinces?.length||0)} provincias`,'Histórico de todos los puntos','bi-globe-americas'),
    kpiCard('Registros relacionados',fmtInt(rows.filter(r=>r.gps).length),'Movilizaciones del filtro actual','bi-link-45deg')
  ].join('');
  const ev=(S.events||[]).filter(x=>x[0]!=='REP PERIÓDICO').slice(0,12);
  chart('chartGpsEvents','bar',ev.map(x=>x[0]),[{label:'Eventos',data:ev.map(x=>x[1]),borderWidth:0,borderRadius:5}],{legend:false,options:{indexAxis:'y',scales:{x:{beginAtZero:true},y:{grid:{display:false}}}}});
  const files=gpsState.manifest?.files||[];
  $('gpsFiles').innerHTML=files.map(f=>`<div class="source-card"><div><strong>${escapeHtml(f.month)} · ${escapeHtml(f.tracker)}</strong><small>${fmtInt(f.points)} puntos · odómetro ${fmtInt((f.odoMax||0)-(f.odoMin||0))} km</small></div><span class="pill ok">Integrado</span></div>`).join('');
  const comp=rows.filter(r=>r.distance>0&&r.gps?.odometerKm>0).slice(-24);
  chart('chartKmCompare','bar',comp.map(r=>formatDate(r.start)),[
    {label:'SharePoint',data:comp.map(r=>r.distance),borderWidth:0,borderRadius:4},
    {label:'GPS referencia',data:comp.map(r=>r.gps.odometerKm),borderWidth:0,borderRadius:4}
  ],{legend:true});
}

export function renderQuality(rows=currentRows){
  const q=dataQuality(rows);
  $('qualityGrid').innerHTML=q.map(x=>`<div class="quality-item"><span>${escapeHtml(x.label)}</span><strong>${fmtPct(x.pct)}</strong><div class="progress"><div style="width:${Math.min(100,x.pct)}%"></div></div><small>${fmtInt(x.count)} de ${fmtInt(rows.length)}</small></div>`).join('');
  const d=sharepointState;
  const a=authDiagnostics();
  $('sourceDiagnostics').innerHTML=[
    ['Usuario',a.account||'—'],['Sitio',d.site?.displayName||'—'],['Lista',d.activeList?.displayName||d.activeList?.name||'—'],['Registros',fmtInt(d.items.length)],['GPS',`${fmtInt(gpsState.points.length)} puntos`],['Redirect URI',a.redirectUri]
  ].map(([k,v])=>`<div class="diag-row"><span>${escapeHtml(k)}</span><strong title="${escapeHtml(v)}">${escapeHtml(v)}</strong></div>`).join('');
  const an=anomalies(rows).slice(0,100);
  $('anomalyRows').innerHTML=an.length?an.map(r=>`<tr data-id="${escapeHtml(r.id)}"><td>${escapeHtml(formatDateTime(r.start))}</td><td>${escapeHtml(r.destinationLabel||r.destination||'—')}</td><td>${escapeHtml(r.requester||'—')}</td><td>${r.issues.map(x=>pill(x,'warn')).join(' ')}</td></tr>`).join(''):'<tr><td colspan="4"><div class="empty">No se detectaron excepciones con las reglas actuales.</div></td></tr>';
  $('anomalyRows').querySelectorAll('tr[data-id]').forEach(tr=>tr.addEventListener('click',()=>openDetail(rows.find(r=>r.id===tr.dataset.id))));
}

export function renderSourceModal(){
  const lists=sharepointState.lists.filter(l=>!(l.list?.hidden));
  $('listChooser').innerHTML=lists.slice(0,25).map(l=>`<div class="source-card ${l.id===sharepointState.activeList?.id?'active':''}"><div><strong>${escapeHtml(l.displayName||l.name)}</strong><small>${escapeHtml(l.webUrl||'')} · puntuación ${fmtInt(l.score||0)}</small></div><button data-list="${escapeHtml(l.id)}">${l.id===sharepointState.activeList?.id?'Activa':'Usar'}</button></div>`).join('');
  $('listChooser').querySelectorAll('button[data-list]').forEach(b=>b.addEventListener('click',()=>onChooseList?.(b.dataset.list)));
  const cols=sharepointState.columns||[];
  $('mappingEditor').innerHTML=Object.entries(SEMANTICS).map(([key,meta])=>`<label class="mapping-item"><span>${escapeHtml(meta.label)}</span><select data-map-key="${escapeHtml(key)}"><option value="">No utilizar</option>${cols.map(c=>`<option value="${escapeHtml(c.name)}" ${sharepointState.mapping[key]===c.name?'selected':''}>${escapeHtml(c.displayName||c.name)}</option>`).join('')}</select></label>`).join('');
}

function openDetail(r){
  if(!r)return;
  $('detailTitle').textContent=r.destinationLabel||'Detalle de movilización';
  const gps=r.gps;
  $('detailBody').innerHTML=`
    <div class="detail-summary">
      ${detailMetric('Inicio',formatDateTime(r.start))}${detailMetric('Fin',formatDateTime(r.end))}${detailMetric('Solicitante',r.requester||'—')}${detailMetric('Grupo / proyecto',r.project||r.group||'—')}${detailMetric('Km SharePoint',r.distance?fmtKm(r.distance):'—')}${detailMetric('Km GPS',gps?.odometerKm?fmtKm(gps.odometerKm):'—')}
      ${detailMetric('Velocidad máx.',gps?`${fmtInt(gps.maxSpeed)} km/h`:'—')}${detailMetric('Provincias',(gps?.provinces||[]).join(' · ')||'—')}${detailMetric('Punto más alejado',gps?.remote?fmtKm(gps.remote.distanceFromOrigin):'—')}${detailMetric('Conciliación',gps?.agreement||'Sin GPS')}${detailMetric('Anticipación solicitud',Number.isFinite(r.leadHours)?`${fmt1(r.leadHours)} h`:'—')}${detailMetric('Duración',r.durationHours?`${fmt1(r.durationHours)} h`:'—')}
    </div>
    <article class="panel compact"><span class="eyebrow">Destino / finalidad registrada</span><p>${escapeHtml(r.destination||'Sin información')}</p></article>
    ${gps?`<article class="panel compact" style="margin-top:10px"><span class="eyebrow">Evidencia GPS</span><p><strong>${escapeHtml(gps.tracker)}</strong> · ${fmtInt(gps.points)} puntos · ${fmtKm(gps.pathKm)} de trayectoria geométrica · ${fmtInt(gps.speedingEvents)} eventos de exceso de velocidad.</p><p class="muted">Punto remoto: ${escapeHtml(gps.remote?.place||'—')}</p><button id="btnDetailMap" class="primary-action small"><i class="bi bi-map"></i> Ver ruta GPS en el mapa</button></article>`:''}
    <details style="margin-top:12px"><summary>Registro completo de SharePoint</summary><pre style="white-space:pre-wrap;background:#f5f7f5;padding:12px;border-radius:12px;overflow:auto">${escapeHtml(safeJson(r.raw))}</pre></details>`;
  $('detailDrawer').classList.add('open');$('detailDrawer').setAttribute('aria-hidden','false');
  $('btnDetailMap')?.addEventListener('click',()=>{closeDetail();onOpenTrip?.(r);});
}
function detailMetric(label,value){return `<div class="detail-metric"><span>${escapeHtml(label)}</span><strong>${escapeHtml(value)}</strong></div>`;}
export function closeDetail(){$('detailDrawer').classList.remove('open');$('detailDrawer').setAttribute('aria-hidden','true');}

export function openSourceModal(){renderSourceModal();$('sourceModal').classList.add('open');$('sourceModal').setAttribute('aria-hidden','false');}
export function closeSourceModal(){$('sourceModal').classList.remove('open');$('sourceModal').setAttribute('aria-hidden','true');}
export function collectMapping(){const map={};$$('[data-map-key]',$('mappingEditor')).forEach(s=>map[s.dataset.mapKey]=s.value);return map;}

export function exportCsv(rows=currentRows){
  const head=['ID','Inicio','Fin','Solicitante','Grupo/Proyecto','Destino','Actividad','Vehículo','Conductor','Km SharePoint','GPS tracker','Km GPS','Diferencia km','Velocidad máxima','Provincias','Conciliación'];
  const body=rows.map(r=>[r.id,formatDateTime(r.start),formatDateTime(r.end),r.requester,r.project||r.group,r.destination,r.activityCategory,r.vehicle||r.plate,r.driver,r.distance,r.gps?.tracker||'',r.gps?.odometerKm||'',r.gps?.differenceKm||'',r.gps?.maxSpeed||'',(r.gps?.provinces||[]).join('|'),r.gps?.agreement||''].map(csvEscape).join(';'));
  downloadText('movilizaciones-fias.csv','\ufeff'+[head.map(csvEscape).join(';'),...body].join('\n'),'text/csv;charset=utf-8');
}

export function renderAll(rows=currentRows,all=allRows){setRows(rows,all);renderOverview(rows);renderTerritory(rows);renderOperations(rows);renderFleet(rows);renderGps(rows);renderQuality(rows);}
