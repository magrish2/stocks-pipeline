#!/usr/bin/env python3
"""
Orquestador del flujo (independiente de Drive: trabaja sobre carpetas).

Para cada crudo:
  1) lo normaliza -> carpeta NORMALIZADOS (copia completa)
  2) busca el maestro con la misma clave (marca_tipo) en FIJOS
       - si no existe: crea el maestro (normalizado, sin lo sin-stock)
       - si existe: lee su Pedido, regenera el maestro EN SU LUGAR desde el
         crudo (actualiza cantidades, quita sin-stock, agrega nuevos, conserva
         Pedido) y reporta altas/bajas/cambios.

La capa de Google Drive (drive.py) baja los crudos a una carpeta temporal,
llama a process_folder y sube los resultados.
"""
import glob
import os
import shutil
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import engine
import ghosts
import match
import sync


def _master_key(path):
    """Clave de un maestro. Los que crea el pipeline se llaman 'MAESTRO <key>.xlsx'
    (estable); si no, se infiere del nombre con match.key_for."""
    b = os.path.splitext(os.path.basename(path))[0]
    if b.lower().startswith("maestro "):
        return b[len("maestro "):].strip()
    return match.key_for(path)


def _find_master(fijos_dir, key):
    for f in glob.glob(os.path.join(fijos_dir, "*.xlsx")):
        if os.path.basename(f).startswith("~$"):
            continue
        if _master_key(f) == key:
            return f
    return None


def process_raw(raw, normalizados_dir, fijos_dir):
    os.makedirs(normalizados_dir, exist_ok=True)
    os.makedirs(fijos_dir, exist_ok=True)
    base = os.path.splitext(os.path.basename(raw))[0]
    key = match.key_for(raw)

    # 1) Normalizado completo -> NORMALIZADOS
    norm_out = os.path.join(normalizados_dir, base + " NORMALIZADO.xlsx")
    engine.normalize(raw, norm_out)
    print(f"  normalizado -> {os.path.basename(norm_out)}")

    # 2) Maestro en FIJOS. No se borra nada: lo sin stock queda en 0.
    master = _find_master(fijos_dir, key)
    if master is None:
        master_out = os.path.join(fijos_dir, f"MAESTRO {key}.xlsx")
        engine.normalize(raw, master_out)          # incluye los de stock 0
        print(f"  maestro NUEVO ({key}) -> {os.path.basename(master_out)}")
        return {"key": key, "master": master_out, "created": True}

    # existe: arrastrar Pedido, regenerar en su lugar y dejar en 0 los que
    # desaparecieron del crudo (no se borran).
    carry = sync.read_pedido_map(master)
    tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False).name
    shutil.copy2(master, tmp)                      # copia del maestro viejo
    engine.normalize(raw, master, carry=carry)     # presentes (incl. 0)
    n_ghost = ghosts.carry_zero(master, tmp)       # desaparecidos -> fila en 0
    d = sync.diff(tmp, master)
    os.remove(tmp)
    print(f"  maestro ACTUALIZADO ({key}): +{len(d['altas'])} altas, "
          f"{n_ghost} dejados en 0, ~{len(d['cambios'])} cambios de cantidad "
          f"(Pedido conservado: {len(carry)})")
    return {"key": key, "master": master, "created": False, "diff": d}


def process_folder(crudos_dir, normalizados_dir, fijos_dir):
    raws = [f for f in sorted(glob.glob(os.path.join(crudos_dir, "*.xls*")))
            if not os.path.basename(f).startswith("~$")]
    print(f"Crudos: {len(raws)}")
    results = []
    for raw in raws:
        print(f"\n## {os.path.basename(raw)}  [{match.key_for(raw)}]")
        results.append(process_raw(raw, normalizados_dir, fijos_dir))
    return results


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("--crudos", required=True)
    ap.add_argument("--normalizados", required=True)
    ap.add_argument("--fijos", required=True)
    a = ap.parse_args()
    process_folder(a.crudos, a.normalizados, a.fijos)
