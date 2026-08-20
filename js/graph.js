import { acquireToken } from './auth.js?v=1.4.0';
import { sleep } from './utils.js?v=1.4.0';

async function fetchWithTimeout(url,options={},timeout=45000){
  const controller=new AbortController(); const timer=setTimeout(()=>controller.abort(),timeout);
  try{return await fetch(url,{...options,signal:controller.signal});}
  catch(err){if(err?.name==='AbortError')throw new Error('Tiempo de espera agotado al contactar Microsoft Graph.');throw err;}
  finally{clearTimeout(timer);}
}

export async function graph(path,options={}){
  const absolute=String(path).startsWith('http');
  const url=absolute?path:`https://graph.microsoft.com/v1.0${path}`;
  let attempt=0;
  while(attempt<3){
    const token=await acquireToken();
    const headers={Authorization:`Bearer ${token}`,...(options.headers||{})};
    if(options.body && !headers['Content-Type']) headers['Content-Type']='application/json';
    const res=await fetchWithTimeout(url,{...options,headers});
    if([409,423,429,503,504].includes(res.status)){ attempt++; await sleep(1000*attempt); continue; }
    if(res.status===204) return null;
    if(!res.ok){
      const body=await res.text().catch(()=>res.statusText);
      throw new Error(`Microsoft Graph ${res.status}: ${body}`);
    }
    const ct=res.headers.get('content-type')||'';
    return ct.includes('json')?res.json():res.text();
  }
  throw new Error('Microsoft Graph no respondió después de varios intentos.');
}

export async function graphPaged(path,max=20000){
  const rows=[]; let url=path;
  while(url && rows.length<max){
    const data=await graph(url); rows.push(...(data?.value||[])); url=data?.['@odata.nextLink']||'';
  }
  return rows.slice(0,max);
}
