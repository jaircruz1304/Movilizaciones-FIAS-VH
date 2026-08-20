#!/usr/bin/env python3
"""Convierte reportes InformeRecorridoPlus PDF en JSON compacto para la app FIAS.

Uso:
    python scripts/extract_gps.py <pdf1> <pdf2> ... --out data/gps

El parser usa pdftotext -layout. Los timestamps se interpretan como hora local
Ecuador (UTC-05:00). No modifica los PDF originales.
"""
from __future__ import annotations
import argparse
import datetime as dt
import json
import os
import re
import subprocess
from pathlib import Path

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
ROW_RE = re.compile(
    r'^\s*(\d{1,2}/\d{1,2}/\d{4}\s+\d{1,2}:\d{2}:\d{2})\s+'
    r'(.*?)\s+(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+'
    r'(\d+(?:\.\d+)?)\s+(\d+(?:\.\d+)?)\s*$'
)
TRACKER_RE = re.compile(r'INFORME DEL RECORRIDO DE:\s*([^\r\n]+)', re.I)
RANGE_RE = re.compile(r'DESDE:\s*([^\r\n]+?)\s+HASTA\s+([^\r\n]+)', re.I)


def pdf_text(pdf: Path) -> str:
    proc = subprocess.run(
        ['pdftotext', '-layout', str(pdf), '-'],
        capture_output=True, text=True, check=True
    )
    return proc.stdout


def cleanup_place(text: str) -> str:
    text = re.sub(r'\s+', ' ', text).strip(' |')
    fixes = {
        'PICHINCH A':'PICHINCHA','TUNGURAH UA':'TUNGURAHUA','COTOPA XI':'COTOPAXI',
        'ORELLAN A':'ORELLANA','CHIMBOR AZO':'CHIMBORAZO','PASTAZ A':'PASTAZA',
        'IMBABUR A':'IMBABURA','MORONA SANTIAG O':'MORONA SANTIAGO'
    }
    for a,b in fixes.items(): text = text.replace(a,b)
    return text


def detect_event(block_text: str, middle: str) -> str:
    joined = f'{middle} {block_text}'
    # evaluar primero eventos largos para evitar coincidencias parciales
    for ev in sorted(EVENTS, key=len, reverse=True):
        if ev in joined:
            return ev
    return 'OTRO'


def parse_report(pdf: Path):
    text = pdf_text(pdf)
    lines = text.splitlines()
    tracker_match = TRACKER_RE.search(text)
    tracker = tracker_match.group(1).strip() if tracker_match else pdf.stem
    range_match = RANGE_RE.search(text)
    report_range = [x.strip() for x in range_match.groups()] if range_match else []
    points = []

    for i, line in enumerate(lines):
        m = ROW_RE.match(line)
        if not m:
            continue
        dt_s, middle, lat, lon, speed, odo = m.groups()
        # cada registro visual está delimitado por líneas en blanco en el PDF
        a = i - 1
        while a >= 0 and lines[a].strip():
            a -= 1
        b = i + 1
        while b < len(lines) and lines[b].strip():
            b += 1
        block_lines = lines[a+1:b]
        block_text = ' '.join(x.strip() for x in block_lines)
        event = detect_event(block_text, middle)

        # reconstruir el texto de ubicación sin cabeceras, fecha, evento ni numéricos
        location_parts = []
        for row in block_lines:
            st = row.strip()
            if not st or any(h in st for h in ('FECHA / HORA','INFORME DEL RECORRIDO','TIPO:','DESDE:','FECHA DE GENERACION')):
                continue
            if ROW_RE.match(row):
                frag = middle
                for ev in EVENTS:
                    frag = frag.replace(ev, '')
                if frag.strip():
                    location_parts.append(frag.strip())
                continue
            frag = st
            for ev in EVENTS:
                frag = frag.replace(ev, '')
            frag = re.sub(r'-?\d+(?:\.\d+)?', '', frag).strip()
            if frag and not re.match(r'^\d+\.', frag):
                location_parts.append(frag)
        place = cleanup_place(' '.join(location_parts))

        local_dt = dt.datetime.strptime(dt_s, '%d/%m/%Y %H:%M:%S')
        ecu = dt.timezone(dt.timedelta(hours=-5))
        epoch = int(local_dt.replace(tzinfo=ecu).timestamp())
        points.append([
            epoch,
            round(float(lat), 6),
            round(float(lon), 6),
            int(float(speed)),
            int(float(odo)),
            EVENT_CODE.get(event, 255),
            place,
        ])

    points.sort(key=lambda p: p[0])
    return {
        'tracker': tracker,
        'source': pdf.name,
        'reportRange': report_range,
        'schema': ['t','lat','lon','speed','odometer','event','place'],
        'points': points,
    }


def month_key_from_points(points):
    if not points: return 'sin-fecha'
    d = dt.datetime.fromtimestamp(points[0][0], tz=dt.timezone(dt.timedelta(hours=-5)))
    return d.strftime('%Y-%m')


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('pdfs', nargs='+')
    ap.add_argument('--out', default='data/gps')
    args = ap.parse_args()
    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)

    manifest = {
        'version': 1,
        'timezone': 'America/Guayaquil',
        'eventCodes': {str(code): name for name, code in EVENT_CODE.items()},
        'schema': ['t','lat','lon','speed','odometer','event','place'],
        'files': []
    }

    total = 0
    for p in args.pdfs:
        report = parse_report(Path(p))
        month = month_key_from_points(report['points'])
        filename = f"{report['tracker'].lower().replace(' ','-')}-{month}.json"
        target = out / filename
        target.write_text(json.dumps(report, ensure_ascii=False, separators=(',',':')), encoding='utf-8')
        pts = report['points']
        total += len(pts)
        odos = [x[4] for x in pts if x[4] > 0]
        manifest['files'].append({
            'tracker': report['tracker'],
            'month': month,
            'url': f'./data/gps/{filename}',
            'points': len(pts),
            'start': pts[0][0] if pts else None,
            'end': pts[-1][0] if pts else None,
            'odoMin': min(odos) if odos else None,
            'odoMax': max(odos) if odos else None,
            'source': report['source'],
        })
        print(f'{target}: {len(pts):,} puntos')

    manifest['files'].sort(key=lambda x: (x['tracker'], x['month']))
    manifest['totalPoints'] = total
    (out / 'manifest.json').write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding='utf-8')
    print(f'Total: {total:,} puntos')


if __name__ == '__main__':
    main()
