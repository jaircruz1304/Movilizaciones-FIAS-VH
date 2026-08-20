import { APP_CONFIG } from '../config/msal-config.js?v=1.5.0';

export const $ = (id) => document.getElementById(id);
export const $$ = (sel, root=document) => [...root.querySelectorAll(sel)];
export const escapeHtml = (v='') => String(v ?? '').replace(/[&<>"']/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[ch]));
export const clamp = (v,min,max) => Math.min(max,Math.max(min,v));
export const sleep = (ms) => new Promise(r=>setTimeout(r,ms));

export function normalizeText(value){
  return String(value ?? '')
    .normalize('NFD').replace(/[\u0300-\u036f]/g,'')
    .replace(/_x([0-9a-fA-F]{4})_/g,(_,h)=>String.fromCharCode(parseInt(h,16)))
    .toLowerCase().replace(/[^a-z0-9]+/g,' ').trim();
}

export function compactKey(value){ return normalizeText(value).replace(/\s+/g,''); }

export function toNumber(value){
  if(value===null || value===undefined || value==='') return 0;
  if(typeof value==='number') return Number.isFinite(value) ? value : 0;
  let s=String(value).trim().replace(/\s+/g,'').replace(/[^0-9,.-]/g,'');
  if(!s) return 0;
  if(s.includes(',') && s.includes('.')){
    if(s.lastIndexOf(',') > s.lastIndexOf('.')) s=s.replace(/\./g,'').replace(',','.');
    else s=s.replace(/,/g,'');
  } else if(s.includes(',')) s=s.replace(',','.');
  else if(/^\d{1,3}(\.\d{3})+$/.test(s)) s=s.replace(/\./g,'');
  const n=Number(s); return Number.isFinite(n) ? n : 0;
}

export function toDate(value){
  if(!value) return null;
  if(value instanceof Date) return Number.isNaN(value) ? null : value;
  if(typeof value==='number' && value>20000 && value<100000){
    const d=new Date(Math.round((value-25569)*86400*1000));
    return Number.isNaN(d) ? null : d;
  }
  const s=String(value).trim();
  let m=s.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if(m){
    const [,dd,mm,yyyy,hh='0',mi='0',ss='0']=m;
    return new Date(`${yyyy}-${mm.padStart(2,'0')}-${dd.padStart(2,'0')}T${hh.padStart(2,'0')}:${mi}:${ss}-05:00`);
  }
  const d=new Date(s); return Number.isNaN(d) ? null : d;
}

export function isoDate(value){ const d=toDate(value); return d ? formatDateKey(d) : ''; }
export function formatDateKey(d){
  return new Intl.DateTimeFormat('en-CA',{timeZone:APP_CONFIG.timezone,year:'numeric',month:'2-digit',day:'2-digit'}).format(d);
}
export function formatDateTime(value){
  const d=value instanceof Date?value:toDate(value);
  if(!d) return '—';
  return new Intl.DateTimeFormat(APP_CONFIG.locale,{timeZone:APP_CONFIG.timezone,dateStyle:'medium',timeStyle:'short'}).format(d);
}
export function formatDate(value){
  const d=value instanceof Date?value:toDate(value);
  if(!d) return '—';
  return new Intl.DateTimeFormat(APP_CONFIG.locale,{timeZone:APP_CONFIG.timezone,dateStyle:'medium'}).format(d);
}
export const fmtInt = n => new Intl.NumberFormat(APP_CONFIG.locale,{maximumFractionDigits:0}).format(Number(n)||0);
export const fmt1 = n => new Intl.NumberFormat(APP_CONFIG.locale,{maximumFractionDigits:1}).format(Number(n)||0);
export const fmtKm = n => `${fmt1(n)} km`;
export const fmtPct = n => `${fmt1(n)}%`;

export function hoursBetween(a,b){
  const da=toDate(a), db=toDate(b); if(!da||!db) return 0; return (db-da)/3600000;
}
export function minutesBetween(a,b){ return hoursBetween(a,b)*60; }

export function haversineKm(a,b){
  if(!a||!b) return 0;
  const R=6371, rad=x=>x*Math.PI/180;
  const dLat=rad(b.lat-a.lat), dLon=rad(b.lon-a.lon);
  const la1=rad(a.lat), la2=rad(b.lat);
  const h=Math.sin(dLat/2)**2 + Math.cos(la1)*Math.cos(la2)*Math.sin(dLon/2)**2;
  return 2*R*Math.asin(Math.sqrt(h));
}

export function groupCounts(rows, getter){
  const m=new Map();
  for(const row of rows){ const v=getter(row); if(v===null||v===undefined||v==='') continue; const key=String(v); m.set(key,(m.get(key)||0)+1); }
  return [...m.entries()].sort((a,b)=>b[1]-a[1]);
}
export function sum(rows,getter){ return rows.reduce((s,r)=>s+(Number(getter(r))||0),0); }
export function average(rows,getter){ const vals=rows.map(getter).map(Number).filter(Number.isFinite); return vals.length ? vals.reduce((a,b)=>a+b,0)/vals.length : 0; }
export function median(values){ const a=values.map(Number).filter(Number.isFinite).sort((x,y)=>x-y); if(!a.length)return 0; const m=Math.floor(a.length/2); return a.length%2?a[m]:(a[m-1]+a[m])/2; }
export function percentile(values,p){ const a=values.map(Number).filter(Number.isFinite).sort((x,y)=>x-y); if(!a.length)return 0; const i=(a.length-1)*p, lo=Math.floor(i), hi=Math.ceil(i); return a[lo]+(a[hi]-a[lo])*(i-lo); }
export function iqrOutlierThreshold(values){ const q1=percentile(values,.25),q3=percentile(values,.75),iqr=q3-q1; return {low:q1-1.5*iqr,high:q3+1.5*iqr,q1,q3}; }

export function downloadText(filename,text,type='text/plain;charset=utf-8'){
  const blob=new Blob([text],{type}); const url=URL.createObjectURL(blob); const a=document.createElement('a'); a.href=url;a.download=filename;a.click();setTimeout(()=>URL.revokeObjectURL(url),1500);
}

export function csvEscape(v){ const s=String(v??''); return /[",\n;]/.test(s)?`"${s.replace(/"/g,'""')}"`:s; }
export function debounce(fn,ms=200){ let t; return (...args)=>{clearTimeout(t);t=setTimeout(()=>fn(...args),ms)}; }

export function toast(title,body='',type='ok'){
  const zone=$('toastZone'); if(!zone)return;
  const el=document.createElement('div'); el.className=`toast ${type}`;
  el.innerHTML=`<strong>${escapeHtml(title)}</strong><span>${escapeHtml(body)}</span>`;
  zone.appendChild(el); setTimeout(()=>el.remove(),6500);
}
export function setLoading(active,text='Procesando...'){
  const el=$('loading'); if(!el)return; el.classList.toggle('show',!!active); const t=$('loadingText'); if(t)t.textContent=text;
}

export function textValue(v){
  if(v===null||v===undefined) return '';
  if(Array.isArray(v)) return v.map(textValue).filter(Boolean).join(', ');
  if(typeof v==='object'){
    return String(v.displayName||v.DisplayName||v.LookupValue||v.Label||v.Title||v.Email||Object.values(v).find(x=>typeof x==='string')||'');
  }
  return String(v).trim();
}

export function safeJson(v){ try{return JSON.stringify(v,null,2)}catch{return String(v)} }
