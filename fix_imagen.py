#!/usr/bin/env python3
"""
Corrige la foto de un modelo Reebok: busca en reebok.com.ar el producto correcto
(por nombre, matcheando el código de modelo en el handle) y guarda su URL en
image_overrides.json, que tiene prioridad sobre el nombre "adivinado" del CDN.

    python fix_imagen.py RBK1100239759
    python fix_imagen.py RBK1100239759 RBK1100208957      # varios
    python fix_imagen.py RBK1100239759 --name "PRIME SERVE"
    python fix_imagen.py RBK1100239759 --url "https://.../foto.jpg"   # manual

Después de correrlo, volvé a normalizar los stocks para que tome la foto buena.
"""
import argparse
import glob
import json
import os
import re
import sys

import requests

import normalizar_stock as ns

UA = ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/124 Safari/537.36")


def model_base(code):
    """Código de modelo base (saca talle/pack). RBK1100239759---M12U... -> RBK1100239759."""
    c = str(code).strip()
    m = re.match(r"^([A-Za-z]{2,4}\d{6,})", c)
    return m.group(1) if m else re.split(r"[-\s]", c)[0]


def buscar_nombre_en_stocks(model):
    """Busca la 'Descripción sin talle' (o descripción) del modelo en los stocks."""
    pats = ["*.xlsb", "*.xlsx", "*/*.xlsb", "*/*.xlsx", "*/*/*.xlsb", "*/*/*.xlsx"]
    vistos = set()
    for pat in pats:
        for f in glob.glob(pat):
            b = os.path.basename(f)
            if b.startswith("~$") or "NORMALIZADO" in b.upper() or " HD.xlsx" in b:
                continue
            if f in vistos:
                continue
            vistos.add(f)
            try:
                sheets = ns.list_sheets(f)
            except Exception:
                continue
            for sh in sheets:
                try:
                    rows = ns.read_sheet_rows(f, sh)
                except Exception:
                    continue
                h = ns.find_header_index(rows)
                if h is None:
                    continue
                cmap = ns.col_map(rows[h])
                si = ns.pick(cmap, "número de artículo", "numero de articulo", "sku")
                ni = ns.pick(cmap, "descripción sin talle", "descripcion sin talle")
                di = ns.pick(cmap, "descripción del artículo", "descripcion del articulo")
                if si is None:
                    continue
                for r in rows[h + 1:]:
                    v = str(r[si]).strip() if si < len(r) and r[si] else ""
                    if model_base(v) == model:
                        nombre = (r[ni] if ni is not None and ni < len(r) else None) \
                            or (r[di] if di is not None and di < len(r) else None)
                        if nombre:
                            return str(nombre).strip(), os.path.basename(f)
    return None, None


def buscar_url_reebok(model, nombre, session):
    """Devuelve la URL de la foto correcta: producto cuyo handle contiene el modelo."""
    queries = []
    if nombre:
        queries.append(nombre)
        queries.append(" ".join(nombre.split()[:3]))
    queries.append(model)
    ml = model.lower()
    for q in queries:
        try:
            r = session.get("https://reebok.com.ar/search/suggest.json",
                            params={"q": q, "resources[type]": "product",
                                    "resources[limit]": 10}, timeout=15)
            prods = r.json().get("resources", {}).get("results", {}).get("products", [])
        except Exception:
            continue
        for p in prods:
            if ml in str(p.get("handle", "")).lower():
                fi = p.get("featured_image")
                url = fi.get("url") if isinstance(fi, dict) else (fi or p.get("image"))
                if url:
                    return url, q
    return None, None


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("modelos", nargs="+", help="Código(s) de modelo o SKU")
    ap.add_argument("--name", help="Nombre del producto (si no, lo busca en los stocks)")
    ap.add_argument("--url", help="URL exacta a usar (solo con 1 modelo)")
    a = ap.parse_args()

    session = requests.Session()
    session.headers.update({"User-Agent": UA})

    overrides = ns.load_json(ns.IMG_OVERRIDES_PATH, default={}) or {}
    cambios = 0

    for code in a.modelos:
        model = model_base(code)
        if a.url and len(a.modelos) == 1:
            url = a.url
            print(f"[{model}] URL manual.")
        else:
            nombre = a.name
            fuente = "--name"
            if not nombre:
                nombre, fuente = buscar_nombre_en_stocks(model)
            if not nombre:
                print(f"[{model}] ✗ no encontré el nombre en los stocks. "
                      f"Pasá --name \"NOMBRE\" o --url URL.")
                continue
            url, q = buscar_url_reebok(model, nombre, session)
            if not url:
                print(f"[{model}] ✗ no encontré el producto en reebok "
                      f"(nombre '{nombre}'). Pasá --url URL a mano.")
                continue
            print(f"[{model}] nombre '{nombre}' ({fuente}) -> match con query '{q}'")

        # verificar que baje
        dest = os.path.join(ns.IMG_CACHE_DIR, f"{re.sub(r'[^A-Za-z0-9_-]', '_', model)}.jpg")
        os.makedirs(ns.IMG_CACHE_DIR, exist_ok=True)
        if not ns.download_thumb(url, dest, session):
            print(f"[{model}] ✗ la URL no bajó una imagen válida: {url}")
            continue

        overrides[model] = url
        cambios += 1
        print(f"[{model}] ✓ override guardado -> {url}")

    if cambios:
        ns.save_json(ns.IMG_OVERRIDES_PATH, overrides)
        print(f"\n✅ {cambios} override(s) en {os.path.basename(ns.IMG_OVERRIDES_PATH)}. "
              f"Volvé a normalizar los stocks para aplicar la foto buena.")
    else:
        print("\nNo se guardó ningún override.")


if __name__ == "__main__":
    main()
