#!/usr/bin/env python3
"""Parser e integrador incremental de reportes GPS InformeRecorridoPlus.

Convierte PDF -> JSON por rastreador/mes. Si un mes ya existe, fusiona los
puntos y elimina duplicados sin borrar el histórico previo.

Uso local:
    python scripts/extract_gps.py reporte1.pdf reporte2.pdf --out data/gps

Requisito del sistema: pdftotext (Poppler).
"""
from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import re
import subprocess
from pathlib import Path
from typing import Any

ECU_TZ = dt.timezone(dt.timedelta(hours=-5))

EVENTS = [
    'INICIA EXCESO VELOCIDAD','FINALIZA EXCESO VELOCIDAD',
    'VEHÍCULO APAGADO','VEHÍCULO ENCENDIDO',
    'REINICIO DISPOSITIVO','LLAMADA ENTRANTE',
    'AUTO-CALIBRACIÓN','FINALIZA DETENCIÓN','INICIA DETENCIÓN',
    'BOTÓN PÁNICO LIBERADO','BOTÓN PÁNICO PRESIONADO',
    'BATERÍA DESCARGÁNDOSE','BATERÍA CARGÁNDOSE',
    'BATERIA PRINCIPAL CONECTADA','BATERIA PRINCIPAL DESCONECTADA',
    'VEHÍCULO DESBLOQUEADO','VEHÍCULO BLOQUEADO',
    'PUERTAS CERRADAS','PUERTAS ABIERTAS',
    'DESCONEXION BLUETOOTH','CONEXION BLUETOOTH',
    'UBICACIÓN EN TIEMPO REAL','REP PERIÓDICO'
]
EVENT_CODE = {name: i for i, name in enumerate(EVENTS)}
SCHEMA = ['t','lat','lon','speed','odometer','event','place']

ROW_RE = re.compile(
    r'^\s*(\d{1,2}/\d{1,2}/\d{4}\s+\d{1,2}:\d{2}:\d{2})\s+'
    r'(.*?)\s+(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+'
    r'(\d+(?:\.\d+)?)\s+(\d+(?:\.\d+)?)\s*$'
)
TRACKER_RE = re.compile(r'INFORME DEL RECORRIDO DE:\s*([^\r\n]+)', re.I)
RANGE_RE = re.compile(r'DESDE:\s*([^\r\n]+?)\s+HASTA\s+([^\r\n]+)', re.I)
TYPE_RE = re.compile(r'TIPO:\s*([^\r\n]+)', re.I)
GEN_RE = re.compile(r'FECHA DE GENERACION:\s*([^\r\n]+)', re.I)


def utc_now_iso() -> str:
    return dt.datetime.now(dt.timezone.utc).replace(microsecond=0).isoformat().replace('+00:00','Z')


def sha256_file(path: Path) -> str:
    h=hashlib.sha256()
    with path.open('rb') as fh:
        for chunk in iter(lambda:fh.read(1024*1024),b''):
            h.update(chunk)
    return h.hexdigest()


def pdf_text(pdf: Path) -> str:
    try:
        proc = subprocess.run(
            ['pdftotext', '-layout', str(pdf), '-'],
            capture_output=True, text=True, check=True
        )
    except FileNotFoundError as exc:
        raise RuntimeError('No se encontró pdftotext (Poppler) en el sistema.') from exc
    except subprocess.CalledProcessError as exc:
        detail=(exc.stderr or '').strip()
        raise RuntimeError(f'pdftotext no pudo leer el PDF: {detail or exc.returncode}') from exc
    return proc.stdout


def cleanup_place(text: str) -> str:
    text = re.sub(r'\s+', ' ', text).strip(' |')
    fixes = {
        'PICHINCH A':'PICHINCHA','TUNGURAH UA':'TUNGURAHUA','COTOPA XI':'COTOPAXI',
        'ORELLAN A':'ORELLANA','CHIMBOR AZO':'CHIMBORAZO','PASTAZ A':'PASTAZA',
        'IMBABUR A':'IMBABURA','MORONA SANTIAG O':'MORONA SANTIAGO'
    }
    for a,b in fixes.items():
        text = text.replace(a,b)
    return text


def detect_event(block_text: str, middle: str) -> str:
    joined = f'{middle} {block_text}'
    for ev in sorted(EVENTS, key=len, reverse=True):
        if ev in joined:
            return ev
    return 'OTRO'


def parse_report(pdf: Path, source_name: str|None=None) -> dict[str, Any]:
    text = pdf_text(pdf)
    if not text.strip():
        raise ValueError('El PDF no contiene texto extraíble.')
    lines = text.splitlines()
    tracker_match = TRACKER_RE.search(text)
    tracker = tracker_match.group(1).strip() if tracker_match else ''
    range_match = RANGE_RE.search(text)
    report_range = [x.strip() for x in range_match.groups()] if range_match else []
    report_type = TYPE_RE.search(text)
    generated = GEN_RE.search(text)
    points: list[list[Any]] = []

    for i, line in enumerate(lines):
        m = ROW_RE.match(line)
        if not m:
            continue
        dt_s, middle, lat, lon, speed, odo = m.groups()
        a = i - 1
        while a >= 0 and lines[a].strip():
            a -= 1
        b = i + 1
        while b < len(lines) and lines[b].strip():
            b += 1
        block_lines = lines[a+1:b]
        block_text = ' '.join(x.strip() for x in block_lines)
        event = detect_event(block_text, middle)

        location_parts=[]
        for row in block_lines:
            st=row.strip()
            if not st or any(h in st for h in ('FECHA / HORA','INFORME DEL RECORRIDO','TIPO:','DESDE:','FECHA DE GENERACION')):
                continue
            if ROW_RE.match(row):
                frag=middle
                for ev in EVENTS:
                    frag=frag.replace(ev,'')
                if frag.strip():
                    location_parts.append(frag.strip())
                continue
            frag=st
            for ev in EVENTS:
                frag=frag.replace(ev,'')
            frag=re.sub(r'-?\d+(?:\.\d+)?','',frag).strip()
            if frag and not re.match(r'^\d+\.',frag):
                location_parts.append(frag)
        place=cleanup_place(' '.join(location_parts))

        local_dt=dt.datetime.strptime(dt_s,'%d/%m/%Y %H:%M:%S')
        epoch=int(local_dt.replace(tzinfo=ECU_TZ).timestamp())
        points.append([
            epoch,
            round(float(lat),6),
            round(float(lon),6),
            int(float(speed)),
            int(float(odo)),
            EVENT_CODE.get(event,255),
            place,
        ])

    points.sort(key=lambda p:p[0])
    report={
        'tracker':tracker,
        'source':source_name or pdf.name,
        'sourceSha256':sha256_file(pdf),
        'reportType':report_type.group(1).strip() if report_type else '',
        'reportRange':report_range,
        'generatedAtSource':generated.group(1).strip() if generated else '',
        'schema':SCHEMA,
        'points':points,
    }
    validate_report(report)
    return report


def validate_report(report: dict[str, Any]) -> dict[str, Any]:
    errors=[]
    warnings=[]
    tracker=str(report.get('tracker') or '').strip()
    points=report.get('points') or []
    if not tracker:
        errors.append('No se identificó el rastreador en “INFORME DEL RECORRIDO DE”.')
    if not points:
        errors.append('No se identificaron registros GPS con fecha, coordenadas, velocidad y kilometraje.')

    invalid_coords=0
    extreme_speed=0
    for p in points:
        try:
            _,lat,lon,speed,odo,_,_=p
            if not (-90<=float(lat)<=90 and -180<=float(lon)<=180):
                invalid_coords+=1
            if float(speed)>250:
                extreme_speed+=1
            if float(odo)<0:
                warnings.append('Se detectaron lecturas negativas de kilometraje.')
                break
        except Exception:
            errors.append('Existe al menos un registro con estructura inválida.')
            break
    if invalid_coords:
        errors.append(f'{invalid_coords} registros contienen coordenadas fuera de rango.')
    if extreme_speed:
        warnings.append(f'{extreme_speed} registros superan 250 km/h y deben revisarse.')
    if errors:
        raise ValueError(' '.join(errors))
    return {'errors':errors,'warnings':warnings}


def month_key(epoch: int) -> str:
    return dt.datetime.fromtimestamp(epoch,tz=ECU_TZ).strftime('%Y-%m')


def split_report_by_month(report: dict[str, Any]) -> dict[str, dict[str, Any]]:
    by_month: dict[str,list[list[Any]]] = {}
    for point in report.get('points') or []:
        by_month.setdefault(month_key(point[0]),[]).append(point)
    out={}
    for month,points in by_month.items():
        out[month]={**report,'month':month,'points':points}
    return out


def safe_tracker_slug(tracker: str) -> str:
    slug=re.sub(r'[^a-z0-9]+','-',tracker.lower()).strip('-')
    return slug or 'tracker'


def point_key(p: list[Any]) -> tuple[Any,...]:
    # Un mismo instante puede contener eventos distintos; el evento forma parte de la clave.
    return (int(p[0]),round(float(p[1]),6),round(float(p[2]),6),int(p[5]))


def merge_points(existing: list[list[Any]], incoming: list[list[Any]]) -> list[list[Any]]:
    merged={point_key(p):p for p in existing}
    # El reporte nuevo prevalece para el mismo instante/coordenada/evento.
    for p in incoming:
        merged[point_key(p)]=p
    return sorted(merged.values(),key=lambda p:(p[0],p[5],p[1],p[2]))


def read_json(path: Path, default: Any) -> Any:
    try:
        return json.loads(path.read_text(encoding='utf-8'))
    except FileNotFoundError:
        return default


def integrate_report(report: dict[str, Any], out: Path, *, uploaded_by: str='', sharepoint_item_id: str='') -> list[dict[str, Any]]:
    out.mkdir(parents=True,exist_ok=True)
    results=[]
    imported_at=utc_now_iso()
    for month,chunk in split_report_by_month(report).items():
        filename=f"{safe_tracker_slug(chunk['tracker'])}-{month}.json"
        target=out/filename
        existing=read_json(target,{})
        old_points=existing.get('points') or []
        merged=merge_points(old_points,chunk['points'])
        source_entry={
            'name':chunk['source'],
            'sha256':chunk.get('sourceSha256',''),
            'importedAt':imported_at,
            'uploadedBy':uploaded_by,
            'sharePointItemId':sharepoint_item_id,
        }
        source_entries=list(existing.get('sources') or [])
        if not any(x.get('sha256')==source_entry['sha256'] and source_entry['sha256'] for x in source_entries):
            source_entries.append(source_entry)
        source_entries=source_entries[-30:]
        payload={
            'tracker':chunk['tracker'],
            'month':month,
            'source':chunk['source'],
            'sources':source_entries,
            'reportRange':chunk.get('reportRange') or [],
            'schema':SCHEMA,
            'points':merged,
            'updatedAt':imported_at,
        }
        target.write_text(json.dumps(payload,ensure_ascii=False,separators=(',',':')),encoding='utf-8')
        results.append({
            'tracker':chunk['tracker'],
            'month':month,
            'file':filename,
            'incomingPoints':len(chunk['points']),
            'previousPoints':len(old_points),
            'points':len(merged),
            'addedPoints':max(0,len(merged)-len(old_points)),
        })
    return results


def rebuild_manifest(out: Path) -> dict[str, Any]:
    files=[]
    total=0
    for path in sorted(out.glob('*.json')):
        if path.name in {'manifest.json','history.json'}:
            continue
        data=read_json(path,{})
        pts=data.get('points') or []
        tracker=data.get('tracker')
        month=data.get('month') or (month_key(pts[0][0]) if pts else '')
        if not tracker or not month:
            continue
        odos=[x[4] for x in pts if len(x)>4 and x[4]>0]
        files.append({
            'tracker':tracker,
            'month':month,
            'file':path.name,
            'points':len(pts),
            'start':pts[0][0] if pts else None,
            'end':pts[-1][0] if pts else None,
            'odoMin':min(odos) if odos else None,
            'odoMax':max(odos) if odos else None,
            'source':data.get('source',''),
            'updatedAt':data.get('updatedAt',''),
        })
        total+=len(pts)
    files.sort(key=lambda x:(x['tracker'],x['month']))
    manifest={
        'version':2,
        'timezone':'America/Guayaquil',
        'updatedAt':utc_now_iso(),
        'eventCodes':{str(code):name for name,code in EVENT_CODE.items()},
        'schema':SCHEMA,
        'files':files,
        'totalPoints':total,
    }
    (out/'manifest.json').write_text(json.dumps(manifest,ensure_ascii=False,indent=2),encoding='utf-8')
    return manifest


def main() -> None:
    ap=argparse.ArgumentParser()
    ap.add_argument('pdfs',nargs='+')
    ap.add_argument('--out',default='data/gps')
    args=ap.parse_args()
    out=Path(args.out)
    out.mkdir(parents=True,exist_ok=True)
    for raw in args.pdfs:
        pdf=Path(raw)
        report=parse_report(pdf)
        results=integrate_report(report,out)
        for r in results:
            print(f"{r['file']}: {r['points']:,} puntos ({r['addedPoints']:+,} nuevos)")
    manifest=rebuild_manifest(out)
    print(f"Total: {manifest['totalPoints']:,} puntos")


if __name__=='__main__':
    main()
