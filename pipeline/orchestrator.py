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


def _blank_pedido(path):
    """Vacía la columna 'Pedido' de todas las hojas (el NORMALIZADO sale sin
    Pedido; el maestro sí lo conserva). Se usa al derivar el NORMALIZADO del
    maestro con una copia, para no re-normalizar el crudo una segunda vez."""
    import openpyxl
    wb = openpyxl.load_workbook(path)
    for ws in wb.worksheets:
        pcol = None
        for c in range(1, (ws.max_column or 1) + 1):
            if str(ws.cell(1, c).value or "").strip().lower() == "pedido":
                pcol = c
                break
        if pcol:
            for r in range(2, (ws.max_row or 1) + 1):
                ws.cell(r, pcol).value = None
    wb.save(path)


def process_raw(raw, normalizados_dir, fijos_dir):
    os.makedirs(normalizados_dir, exist_ok=True)
    os.makedirs(fijos_dir, exist_ok=True)
    base = os.path.splitext(os.path.basename(raw))[0]
    key = match.key_for(raw)
    brand = engine.detect_brand(raw)

    # 0) ¿Es un PENDIENTE (pedidos por cliente)? -> solo imágenes, sin maestro.
    if engine.is_pendiente(raw):
        out = os.path.join(normalizados_dir, base + " CON IMAGENES.xlsx")
        engine.add_images_pendiente(raw, out)
        print(f"  pendiente (con imágenes) -> {os.path.basename(out)}")
        return {"key": None, "brand": brand, "pendiente": True,
                "norm": os.path.basename(out), "master": None}

    # El NORMALIZADO y el maestro son el MISMO contenido salvo la columna Pedido
    # (el maestro la arrastra). Por eso se normaliza UNA sola vez (al maestro) y
    # el NORMALIZADO se deriva copiando + vaciando Pedido — antes se normalizaba
    # el crudo dos veces (reparseo del .xlsb + re-embebido de fotos duplicados).
    norm_out = os.path.join(normalizados_dir, base + " NORMALIZADO.xlsx")
    norm_name = os.path.basename(norm_out)
    master = _find_master(fijos_dir, key)

    if master is None:
        master_out = os.path.join(fijos_dir, f"MAESTRO {key}.xlsx")
        engine.normalize(raw, master_out)          # incluye los de stock 0
        shutil.copy2(master_out, norm_out)         # maestro nuevo no tiene Pedido
        print(f"  normalizado -> {norm_name}")
        print(f"  maestro NUEVO ({key}) -> {os.path.basename(master_out)}")
        return {"key": key, "brand": brand, "norm": norm_name,
                "master": master_out, "created": True}

    # existe: arrastrar Pedido y regenerar en su lugar. Los SKU que
    # desaparecieron del crudo se BORRAN (ya no se dejan en 0).
    carry = sync.read_pedido_map(master)
    tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False).name
    shutil.copy2(master, tmp)                      # copia del maestro viejo (para el diff)
    engine.normalize(raw, master, carry=carry)     # UNA sola normalización
    d = sync.diff(tmp, master)
    os.remove(tmp)
    shutil.copy2(master, norm_out)                 # NORMALIZADO = maestro...
    _blank_pedido(norm_out)                         # ...pero con Pedido en blanco
    print(f"  normalizado -> {norm_name}")
    print(f"  maestro ACTUALIZADO ({key}): +{len(d['altas'])} altas, "
          f"{len(d['bajas'])} borrados (desaparecidos del crudo), "
          f"~{len(d['cambios'])} cambios de cantidad "
          f"(Pedido conservado: {len(carry)})")
    return {"key": key, "brand": brand, "norm": norm_name,
            "master": master, "created": False, "diff": d}


def process_folder(crudos_dir, normalizados_dir, fijos_dir):
    raws = [f for f in sorted(glob.glob(os.path.join(crudos_dir, "*.xls*")))
            if not os.path.basename(f).startswith("~$")]
    print(f"Crudos: {len(raws)}")
    results = []
    for raw in raws:
        name = os.path.basename(raw)
        print(f"\n## {name}  [{match.key_for(raw)}]")
        try:
            r = process_raw(raw, normalizados_dir, fijos_dir)
            r["raw"] = name
        except Exception as e:
            import traceback
            # Un crudo roto NO frena la tanda: se marca y se sigue con el resto.
            print(f"  [!] ERROR en {name}: {e}")
            r = {"raw": name, "error": f"{type(e).__name__}: {e}",
                 "trace": traceback.format_exc()}
        results.append(r)
    return results


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("--crudos", required=True)
    ap.add_argument("--normalizados", required=True)
    ap.add_argument("--fijos", required=True)
    a = ap.parse_args()
    process_folder(a.crudos, a.normalizados, a.fijos)
