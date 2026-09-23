#!/usr/bin/env python3
"""Procesa la bandeja GPS de SharePoint y publica JSON protegidos en SharePoint.

Arquitectura de producción:
  GitHub Pages (SPA) -> SharePoint/GPS/Entrada -> GitHub Actions
  -> SharePoint/GPS/Publicados (manifest + JSON mensuales + history)

Los JSON de coordenadas NO se publican en GitHub Pages ni se guardan en el
repositorio. El workflow usa una App Registration de Microsoft Entra con
Sites.Selected y permiso write únicamente sobre el sitio autorizado.
"""
from __future__ import annotations

import json
import os
import sys
import tempfile
import time
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path
from typing import Any

from extract_gps import integrate_report, parse_report, rebuild_manifest, utc_now_iso

TENANT_ID=os.getenv('MS_TENANT_ID','').strip()
CLIENT_ID=os.getenv('MS_ACTION_CLIENT_ID','').strip()
CLIENT_SECRET=os.getenv('MS_ACTION_CLIENT_SECRET','').strip()
SP_HOST=os.getenv('SP_HOST','fiasec.sharepoint.com').strip()
SP_SITE_PATH=os.getenv('SP_SITE_PATH','/sites/FONDODEINVERSIONAMBIENTALSOSTENIBLE').strip()
INBOX=os.getenv('SP_GPS_INBOX','Movilizaciones-FIAS/GPS/Entrada').strip('/')
PROCESSED=os.getenv('SP_GPS_PROCESSED','Movilizaciones-FIAS/GPS/Procesados').strip('/')
REJECTED=os.getenv('SP_GPS_REJECTED','Movilizaciones-FIAS/GPS/Rechazados').strip('/')
PUBLISHED=os.getenv('SP_GPS_PUBLISHED','Movilizaciones-FIAS/GPS/Publicados').strip('/')
MAX_HISTORY=int(os.getenv('GPS_HISTORY_LIMIT','500'))


def require_env() -> None:
    missing=[k for k,v in {
        'MS_TENANT_ID':TENANT_ID,
        'MS_ACTION_CLIENT_ID':CLIENT_ID,
        'MS_ACTION_CLIENT_SECRET':CLIENT_SECRET,
    }.items() if not v]
    if missing:
        raise RuntimeError('Faltan secretos/variables: '+', '.join(missing))


def http_data(url: str, *, method='GET', headers=None, data: bytes|None=None, retries=4) -> Any:
    headers={'Accept':'application/json',**(headers or {})}
    for attempt in range(retries):
        req=urllib.request.Request(url,data=data,headers=headers,method=method)
        try:
            with urllib.request.urlopen(req,timeout=90) as res:
                raw=res.read()
                if not raw:
                    return None
                ctype=res.headers.get('Content-Type','')
                if 'json' in ctype or raw[:1] in (b'{',b'['):
                    return json.loads(raw.decode('utf-8'))
                return raw
        except urllib.error.HTTPError as exc:
            body=exc.read().decode('utf-8','replace')
            if exc.code in (429,500,502,503,504) and attempt<retries-1:
                retry_after=exc.headers.get('Retry-After')
                wait=float(retry_after) if retry_after and retry_after.isdigit() else 2**attempt
                time.sleep(min(wait,20))
                continue
            raise RuntimeError(f'HTTP {exc.code} {url}: {body[:1200]}') from exc
        except urllib.error.URLError as exc:
            if attempt<retries-1:
                time.sleep(2**attempt)
                continue
            raise RuntimeError(f'Error de red en {url}: {exc}') from exc
    raise RuntimeError(f'No se pudo completar la solicitud: {url}')


def token() -> str:
    url=f'https://login.microsoftonline.com/{urllib.parse.quote(TENANT_ID)}/oauth2/v2.0/token'
    body=urllib.parse.urlencode({
        'client_id':CLIENT_ID,
        'client_secret':CLIENT_SECRET,
        'scope':'https://graph.microsoft.com/.default',
        'grant_type':'client_credentials',
    }).encode()
    data=http_data(url,method='POST',headers={'Content-Type':'application/x-www-form-urlencoded'},data=body)
    access=data.get('access_token') if isinstance(data,dict) else None
    if not access:
        raise RuntimeError('Microsoft Entra no devolvió access_token para GitHub Actions.')
    return access


class Graph:
    def __init__(self,access_token: str):
        self.access_token=access_token
        self.base='https://graph.microsoft.com/v1.0'

    def _url(self,path: str) -> str:
        return path if path.startswith('http') else self.base+path

    def request(self,path: str,*,method='GET',body: Any=None) -> Any:
        headers={'Authorization':f'Bearer {self.access_token}'}
        data=None
        if body is not None:
            data=json.dumps(body,ensure_ascii=False).encode('utf-8')
            headers['Content-Type']='application/json'
        return http_data(self._url(path),method=method,headers=headers,data=data)

    def paged(self,path: str) -> list[dict[str,Any]]:
        rows=[]; url=path
        while url:
            data=self.request(url)
            rows.extend((data or {}).get('value') or [])
            url=(data or {}).get('@odata.nextLink') or ''
        return rows

    def download(self,path: str) -> bytes:
        data=http_data(self._url(path),headers={'Authorization':f'Bearer {self.access_token}'})
        if isinstance(data,bytes): return data
        if isinstance(data,(dict,list)): return json.dumps(data,ensure_ascii=False,indent=2).encode('utf-8')
        if isinstance(data,str): return data.encode('utf-8')
        raise RuntimeError('Microsoft Graph no devolvió contenido descargable.')

    def put_content(self,drive_id: str,path: str,data: bytes,content_type='application/octet-stream') -> dict[str,Any]:
        url=self._url(f'/drives/{urllib.parse.quote(drive_id,safe="")}/root:/{enc_path(path)}:/content')
        result=http_data(url,method='PUT',headers={
            'Authorization':f'Bearer {self.access_token}',
            'Content-Type':content_type,
            'Accept':'application/json'
        },data=data)
        if not isinstance(result,dict):
            raise RuntimeError(f'No se confirmó la publicación de {path}.')
        return result


def enc_path(path: str) -> str:
    return '/'.join(urllib.parse.quote(s,safe='') for s in path.split('/') if s)


def graph_item_by_path(g: Graph,drive_id: str,path: str) -> dict[str,Any]|None:
    try:
        return g.request(f'/drives/{urllib.parse.quote(drive_id,safe="")}/root:/{enc_path(path)}?$select=id,name,webUrl,folder,parentReference,file,size')
    except RuntimeError as exc:
        if 'HTTP 404' in str(exc): return None
        raise


def ensure_folder(g: Graph,drive_id: str,path: str) -> dict[str,Any]:
    segments=[s for s in path.split('/') if s]
    parent={'id':'root'}; current=''
    for segment in segments:
        current=f'{current}/{segment}'.strip('/')
        found=graph_item_by_path(g,drive_id,current)
        if found:
            parent=found; continue
        endpoint=(
            f'/drives/{urllib.parse.quote(drive_id,safe="")}/root/children'
            if parent['id']=='root' else
            f'/drives/{urllib.parse.quote(drive_id,safe="")}/items/{urllib.parse.quote(parent["id"],safe="")}/children'
        )
        parent=g.request(endpoint,method='POST',body={
            'name':segment,'folder':{},'@microsoft.graph.conflictBehavior':'fail'
        })
    return parent


def download_if_exists(g: Graph,drive_id: str,path: str) -> bytes|None:
    item=graph_item_by_path(g,drive_id,path)
    if not item or not item.get('file'): return None
    return g.download(f'/drives/{urllib.parse.quote(drive_id,safe="")}/items/{urllib.parse.quote(item["id"],safe="")}/content')


def materialize_published(g: Graph,drive_id: str,data_dir: Path) -> dict[str,Any]|None:
    data_dir.mkdir(parents=True,exist_ok=True)
    raw=download_if_exists(g,drive_id,f'{PUBLISHED}/manifest.json')
    if not raw: return None
    manifest=json.loads(raw.decode('utf-8'))
    (data_dir/'manifest.json').write_bytes(raw)
    for entry in manifest.get('files') or []:
        filename=entry.get('file') or str(entry.get('url') or '').split('/')[-1]
        if not filename: continue
        content=download_if_exists(g,drive_id,f'{PUBLISHED}/{filename}')
        if content: (data_dir/filename).write_bytes(content)
    history=download_if_exists(g,drive_id,f'{PUBLISHED}/history.json')
    if history: (data_dir/'history.json').write_bytes(history)
    return manifest


def load_history(path: Path) -> dict[str,Any]:
    try:
        raw=json.loads(path.read_text(encoding='utf-8'))
        if isinstance(raw,list): return {'version':1,'entries':raw}
        if isinstance(raw,dict): raw.setdefault('entries',[]); return raw
    except (FileNotFoundError,json.JSONDecodeError):
        pass
    return {'version':2,'entries':[]}


def save_history(path: Path,history: dict[str,Any]) -> None:
    history['version']=2
    history['updatedAt']=utc_now_iso()
    history['entries']=(history.get('entries') or [])[-MAX_HISTORY:]
    path.write_text(json.dumps(history,ensure_ascii=False,indent=2),encoding='utf-8')


def created_by(item: dict[str,Any]) -> str:
    user=(item.get('createdBy') or {}).get('user') or {}
    app=(item.get('createdBy') or {}).get('application') or {}
    return user.get('email') or user.get('displayName') or app.get('displayName') or ''


def original_name(queued_name: str) -> str:
    return queued_name.split('__',1)[1] if '__' in queued_name else queued_name


def history_entry(item: dict[str,Any],status: str,**extra: Any) -> dict[str,Any]:
    return {
        'id':f"{item.get('id','')}:{utc_now_iso()}",
        'sharePointItemId':item.get('id',''),
        'queuedName':item.get('name',''),
        'sourceFile':original_name(item.get('name','')),
        'uploadedAt':item.get('createdDateTime',''),
        'uploadedBy':created_by(item),
        'size':item.get('size',0),
        'status':status,
        'processedAt':utc_now_iso(),
        **extra,
    }


def move_item(g: Graph,drive_id: str,item: dict[str,Any],folder: dict[str,Any]) -> None:
    g.request(
        f'/drives/{urllib.parse.quote(drive_id,safe="")}/items/{urllib.parse.quote(item["id"],safe="")}',
        method='PATCH',body={'parentReference':{'id':folder['id']},'name':item.get('name')}
    )


def main() -> int:
    require_env()
    g=Graph(token())
    site_path=SP_SITE_PATH.strip('/')
    site=g.request(f'/sites/{SP_HOST}:/{site_path}?$select=id,displayName,webUrl')
    drive=g.request(f'/sites/{urllib.parse.quote(site["id"],safe="")}/drive?$select=id,name,webUrl')
    drive_id=drive['id']
    inbox=ensure_folder(g,drive_id,INBOX)
    processed=ensure_folder(g,drive_id,PROCESSED)
    rejected=ensure_folder(g,drive_id,REJECTED)
    ensure_folder(g,drive_id,PUBLISHED)

    items=g.paged(
        f'/drives/{urllib.parse.quote(drive_id,safe="")}/items/{urllib.parse.quote(inbox["id"],safe="")}/children'
        '?$select=id,name,size,createdDateTime,createdBy,lastModifiedDateTime,file,webUrl&$top=200'
    )
    pdfs=[x for x in items if x.get('file') and str(x.get('name','')).lower().endswith('.pdf')]
    print(f'Sitio: {site.get("displayName")} | Biblioteca: {drive.get("name")} | PDFs pendientes: {len(pdfs)}')
    if not pdfs:
        return 0

    with tempfile.TemporaryDirectory(prefix='fias-gps-sync-') as td:
        data_dir=Path(td)/'gps'
        existing_manifest=materialize_published(g,drive_id,data_dir)
        history_path=data_dir/'history.json'
        history=load_history(history_path)
        known_success={x.get('sha256') for x in history.get('entries',[]) if x.get('sha256') and x.get('status') in {'processed','duplicate','baseline'}}

        processed_count=rejected_count=duplicate_count=0
        changed_monthly=set()
        moves=[]
        for item in sorted(pdfs,key=lambda x:x.get('createdDateTime','')):
            try:
                content=g.download(f'/drives/{urllib.parse.quote(drive_id,safe="")}/items/{urllib.parse.quote(item["id"],safe="")}/content')
                local=Path(td)/f'input-{item["id"]}.pdf'
                local.write_bytes(content)
                report=parse_report(local,source_name=original_name(item['name']))
                file_hash=report.get('sourceSha256','')
                if file_hash and file_hash in known_success:
                    duplicate_count+=1
                    history['entries'].append(history_entry(
                        item,'duplicate',sha256=file_hash,tracker=report.get('tracker',''),
                        message='El contenido ya había sido procesado previamente; no se duplicaron puntos.'
                    ))
                    moves.append((item,processed))
                    print(f'DUPLICADO: {original_name(item["name"])}')
                    continue

                integrations=integrate_report(
                    report,data_dir,uploaded_by=created_by(item),sharepoint_item_id=item.get('id','')
                )
                changed_monthly.update(x['file'] for x in integrations)
                months=[x['month'] for x in integrations]
                added=sum(x['addedPoints'] for x in integrations)
                total_in=sum(x['incomingPoints'] for x in integrations)
                history['entries'].append(history_entry(
                    item,'processed',sha256=file_hash,tracker=report.get('tracker',''),
                    months=months,inputPoints=total_in,addedPoints=added,integrations=integrations,
                    warnings=[],message='Reporte validado e integrado correctamente.'
                ))
                if file_hash: known_success.add(file_hash)
                moves.append((item,processed)); processed_count+=1
                print(f'OK: {original_name(item["name"])} -> {", ".join(months)} | +{added:,} puntos')
            except Exception as exc:
                rejected_count+=1
                history['entries'].append(history_entry(
                    item,'rejected',error=str(exc)[:2000],message='El reporte fue rechazado durante validación o procesamiento.'
                ))
                moves.append((item,rejected))
                print(f'ERROR: {item.get("name")}: {exc}',file=sys.stderr)

        if processed_count:
            manifest=rebuild_manifest(data_dir)
        else:
            manifest=existing_manifest or {'version':2,'files':[],'totalPoints':0,'eventCodes':{},'schema':[]}
        save_history(history_path,history)

        # Publicar primero; solo después mover los PDF fuera de Entrada. Si esta fase falla,
        # la siguiente ejecución puede reintentar sin perder el documento fuente.
        for filename in sorted(changed_monthly):
            g.put_content(drive_id,f'{PUBLISHED}/{filename}',(data_dir/filename).read_bytes(),'application/json')
        if processed_count:
            g.put_content(drive_id,f'{PUBLISHED}/manifest.json',(data_dir/'manifest.json').read_bytes(),'application/json')
        g.put_content(drive_id,f'{PUBLISHED}/history.json',history_path.read_bytes(),'application/json')

        for item,target in moves:
            move_item(g,drive_id,item,target)

        print(
            f'Resultado: {processed_count} procesados, {duplicate_count} duplicados, '
            f'{rejected_count} rechazados. Total protegido: {manifest.get("totalPoints",0):,} puntos.'
        )
        return 0 if rejected_count==0 else 2


if __name__=='__main__':
    raise SystemExit(main())
