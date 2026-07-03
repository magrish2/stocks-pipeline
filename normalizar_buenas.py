#!/usr/bin/env python3
"""
Normaliza stocks al layout de venta con imágenes 'buenas' pero livianas
(thumbnails nítidos desde la fuente oficial, SIN ESRGAN → rápido) y deja
todo en una carpeta aparte.

Uso (corré cuando quieras):
    python normalizar_buenas.py                      # auto: *con referencia*.xlsx
    python normalizar_buenas.py "archivo1.xlsx" ...  # archivos puntuales
    python normalizar_buenas.py --out-dir "Salida" --thumb-px 512

Defaults pensados para rapidez: 512px de thumbnail desde fuente 1024px.
Subí --thumb-px para más nitidez (más lento/pesado), bajalo para más rápido.
"""
import argparse
import glob
import os
from types import SimpleNamespace

import normalizar_stock as ns


def encontrar_stocks(carpeta="."):
    """Stocks candidatos: '*con referencia*.xlsx', sin los ya generados."""
    out = []
    for f in sorted(glob.glob(os.path.join(carpeta, "*con referencia*.xlsx"))):
        base = os.path.basename(f).upper()
        if "CON IMAGENES" in base or "NORMALIZADO" in base or base.startswith("~$"):
            continue
        out.append(f)
    return out


def normalizar_buenas(files=None, out_dir="Stocks normalizados",
                      thumb_px=512, width=1024, quality=85, source="reebok"):
    """
    Genera el Excel normalizado (Foto, SKU, Descripción, Disponibilidad,
    precios, Pedido) para cada stock, con fotos de reebok.com.ar + fallback a
    las imágenes embebidas del propio archivo. Devuelve la lista de salidas.
    """
    files = files or encontrar_stocks()
    if not files:
        print("No encontré stocks 'con referencia'. Pasá los archivos como argumento.")
        return []

    # Calidad de imagen (globals que usa normalizar_stock al descargar).
    ns.THUMB_PX = thumb_px
    ns.THUMB_QUALITY = quality
    ns.REEBOK_CDN = "https://reebok.com.ar/cdn/shop/files/{name}?width=%d" % width

    os.makedirs(out_dir, exist_ok=True)
    salidas = []
    for f in files:
        if not os.path.exists(f):
            print(f"[!] no existe: {f}")
            continue
        base = os.path.splitext(os.path.basename(f))[0]
        base = base.replace(" + con referencia", "").replace(" con referencia", "")
        out = os.path.join(out_dir, base + " NORMALIZADO.xlsx")
        print(f"\n######## {f}\n   -> {out}")
        args = SimpleNamespace(input=f, output=out, source=source,
                               no_images=False, image_mode="embed",
                               limit=None, sheets=None)
        ns.cmd_run(args)
        salidas.append(out)

    print(f"\n=== LISTO: {len(salidas)} stock(s) en la carpeta '{out_dir}'")
    return salidas


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("files", nargs="*", help="Stocks a normalizar (default: auto)")
    ap.add_argument("--out-dir", default="Stocks normalizados",
                    help="Carpeta de salida (default: 'Stocks normalizados')")
    ap.add_argument("--thumb-px", type=int, default=512,
                    help="Lado máx. del thumbnail (default 512; más=nítido/lento)")
    ap.add_argument("--width", type=int, default=1024,
                    help="Ancho que se pide a reebok.com.ar (default 1024)")
    ap.add_argument("--quality", type=int, default=85, help="Calidad JPEG (1-95)")
    ap.add_argument("--source", choices=["reebok", "meli", "both"], default="reebok")
    a = ap.parse_args()
    normalizar_buenas(a.files or None, a.out_dir, a.thumb_px, a.width,
                      a.quality, a.source)


if __name__ == "__main__":
    main()
