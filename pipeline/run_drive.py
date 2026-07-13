#!/usr/bin/env python3
"""
Runner del pipeline sobre Google Drive.

Flujo:
  1) baja los crudos de la carpeta CRUDOS a una carpeta temporal
  2) baja los maestros actuales de FIJOS (para poder actualizarlos en su lugar)
  3) corre el orquestador local (normaliza + crea/actualiza maestros)
  4) sube los NORMALIZADOS a Drive
  5) sube/reemplaza los MAESTROS en FIJOS (mismo archivo = en su lugar)

Config: pipeline/config.json  (ver config.example.json)
    { "crudos": "<folderId>", "normalizados": "<folderId>", "fijos": "<folderId>" }
"""
import json
import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import drive
import orchestrator

BASE = os.path.dirname(os.path.abspath(__file__))
CONFIG = os.path.join(BASE, "config.json")

RAW_EXT = (".xlsx", ".xlsb")


def main(keep_crudos=False):
    cfg = json.load(open(CONFIG))
    svc = drive.service()
    work = tempfile.mkdtemp(prefix="pipe_")
    d_crudos = os.path.join(work, "crudos")
    d_norm = os.path.join(work, "norm")
    d_fijos = os.path.join(work, "fijos")
    for d in (d_crudos, d_norm, d_fijos):
        os.makedirs(d, exist_ok=True)

    # 1) bajar crudos (guardando su id para poder eliminarlos al final)
    crudos = drive.list_files(svc, cfg["crudos"])
    procesados = []            # (file_id, nombre) de los crudos bajados
    for fid, name, mime in crudos:
        if name.startswith("~$"):
            continue
        if mime == drive.GSHEET_MIME or name.lower().endswith(RAW_EXT):
            ext = ".xlsx" if mime == drive.GSHEET_MIME else os.path.splitext(name)[1]
            dest = os.path.join(d_crudos, os.path.splitext(name)[0] + ext)
            drive.download(svc, fid, dest, mime)
            procesados.append((fid, name))
    print(f"Crudos bajados: {len(procesados)}")
    if not procesados:
        print("No hay crudos para procesar.")
        return

    # 2) bajar maestros actuales de FIJOS
    fijos_remote = drive.list_files(svc, cfg["fijos"])
    for fid, name, mime in fijos_remote:
        if name.lower().endswith(".xlsx") and not name.startswith("~$"):
            drive.download(svc, fid, os.path.join(d_fijos, name), mime)
    print(f"Maestros existentes: {len([f for f in fijos_remote if f[1].lower().endswith('.xlsx')])}")

    # 3) procesar local
    results = orchestrator.process_folder(d_crudos, d_norm, d_fijos)

    # 4) subir normalizados a una subcarpeta por MARCA (Reebok/Kappa/Crocs)
    BRAND_DIR = {"reebok": "Reebok", "kappa": "Kappa", "crocs": "Crocs"}
    brand_folders = {}          # marca -> folderId (cache)
    brand_listing = {}          # folderId -> archivos (para upsert)
    for r in results:
        norm_path = os.path.join(d_norm, r["norm"])
        if not os.path.exists(norm_path):
            continue
        carpeta = BRAND_DIR.get(r["brand"], "Otros")
        if carpeta not in brand_folders:
            fid = drive.get_or_create_folder(svc, cfg["normalizados"], carpeta)
            brand_folders[carpeta] = fid
            brand_listing[fid] = drive.list_files(svc, fid)
        fid = brand_folders[carpeta]
        _id, nuevo = drive.upsert_by_name(svc, norm_path, fid, brand_listing[fid])
        print(f"  norm -> {carpeta}/{r['norm']} ({'nuevo' if nuevo else 'reemplazado'})")

    # 5) subir/actualizar maestros en su lugar
    for f in sorted(os.listdir(d_fijos)):
        if f.lower().endswith(".xlsx") and not f.startswith("~$"):
            _id, nuevo = drive.upsert_by_name(svc, os.path.join(d_fijos, f),
                                              cfg["fijos"], fijos_remote)
            print(f"  maestro {'creado' if nuevo else 'actualizado en su lugar'}: {f}")

    # 6) recién ahora (todo subido OK) mandar los crudos a la Papelera
    if keep_crudos:
        print("\n(--keep) Los crudos se dejan en su carpeta.")
    else:
        for fid, name in procesados:
            drive.trash(svc, fid)
            print(f"  crudo a Papelera: {name}")

    print("\n✅ Pipeline completo.")


def main_cli():
    import argparse
    ap = argparse.ArgumentParser(description="Pipeline de stocks sobre Google Drive")
    ap.add_argument("--keep", action="store_true",
                    help="No eliminar los crudos al terminar (por defecto van a Papelera)")
    a = ap.parse_args()
    main(keep_crudos=a.keep)


if __name__ == "__main__":
    main_cli()
