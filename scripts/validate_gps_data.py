#!/usr/bin/env python3
"""Valida consistencia de un directorio GPS generado por extract_gps.py."""
from __future__ import annotations
import argparse
import json
import sys
from pathlib import Path


def fail(msg:str):
    print(f'ERROR: {msg}',file=sys.stderr)
    raise SystemExit(1)


def validate(data_dir: Path):
    manifest_path=data_dir/'manifest.json'
    if not manifest_path.exists():
        fail(f'No existe {manifest_path}')
    manifest=json.loads(manifest_path.read_text(encoding='utf-8'))
    total=0
    seen=set()
    for entry in manifest.get('files') or []:
        name=entry.get('file') or str(entry.get('url','')).rsplit('/',1)[-1]
        if not name: fail('Entrada de manifiesto sin nombre de archivo.')
        path=data_dir/name
        if not path.exists(): fail(f'El manifiesto referencia un archivo inexistente: {name}')
        data=json.loads(path.read_text(encoding='utf-8'))
        points=data.get('points') or []
        if len(points)!=entry.get('points'):
            fail(f'Conteo inconsistente en {name}: manifest={entry.get("points")} json={len(points)}')
        last=None
        for p in points:
            if not isinstance(p,list) or len(p)<7: fail(f'Registro inválido en {name}')
            t,lat,lon,speed,odo,event,place=p[:7]
            if not (-90<=float(lat)<=90 and -180<=float(lon)<=180): fail(f'Coordenadas fuera de rango en {name}: {lat}, {lon}')
            if last is not None and int(t)<last: fail(f'Orden temporal inválido en {name}')
            last=int(t)
            key=(int(t),round(float(lat),6),round(float(lon),6),int(event))
            if key in seen: fail(f'Duplicado global detectado en {name}: {key}')
            seen.add(key)
        total+=len(points)
    if total!=manifest.get('totalPoints'): fail(f'Total inconsistente: manifest={manifest.get("totalPoints")} real={total}')
    print(f'OK: {len(manifest.get("files") or [])} archivos GPS, {total:,} puntos válidos.')


def main():
    ap=argparse.ArgumentParser()
    ap.add_argument('--data',default='data/gps',help='Directorio que contiene manifest.json y JSON mensuales')
    args=ap.parse_args()
    validate(Path(args.data))

if __name__=='__main__':
    main()
