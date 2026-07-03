#!/usr/bin/env python3
"""Verifica credenciales + IDs de carpetas SIN tocar ningún archivo.
Lista lo que hay en cada carpeta de Drive."""
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import drive

BASE = os.path.dirname(os.path.abspath(__file__))


def main():
    if not os.path.exists(drive.CREDENTIALS):
        sys.exit(f"✗ Falta {drive.CREDENTIALS}\n  Descargá el JSON de OAuth y "
                 f"guardalo ahí (ver README, paso A).")
    cfg = json.load(open(os.path.join(BASE, "config.json")))
    svc = drive.service()
    print("✓ Autorización OK\n")
    for nombre in ("crudos", "normalizados", "fijos"):
        fid = cfg[nombre].strip()
        try:
            files = drive.list_files(svc, fid)
        except Exception as e:
            print(f"✗ {nombre}: error accediendo a la carpeta ({fid}) -> {e}")
            continue
        print(f"✓ {nombre} ({fid}): {len(files)} archivo(s)")
        for _id, name, _m in files[:10]:
            print(f"     - {name}")
    print("\nSi las 3 carpetas responden, ya podés correr:  python pipeline/run_drive.py")


if __name__ == "__main__":
    main()
