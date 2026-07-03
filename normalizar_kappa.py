#!/usr/bin/env python3
"""
Normaliza los stocks de KAPPA al mismo layout de venta que Reebok
(Foto, SKU, Descripción, Género, Disponibilidad, precios, Pedido).

Kappa no tiene CDN online: las fotos salen de las imágenes embebidas de cada
archivo (modo --no-online → rápido, sin red). Salta hojas auxiliares
('BASE SKU', 'Hoja2') y la hoja duplicada 'KAPPA CALZADO (2)'.

Uso:
    python normalizar_kappa.py                 # auto: todos los kappa/*.xlsx
    python normalizar_kappa.py "kappa/X.xlsx"  # archivos puntuales
"""
import argparse
import glob
import os
from types import SimpleNamespace

import normalizar_stock as ns

IN_DIR = "kappa"
OUT_DIR = os.path.join("kappa", "Stocks normalizados")

# Tienda oficial Kappa Argentina (Shopify). La imagen se nombra por
# "Modelo Color" (ej. K1322835W-K001 -> .../K1322835W-K001-1.jpg).
KAPPA_CDN = "https://www.kappastore.com.ar/cdn/shop/files/{name}?width=1024"

# Hojas que NO son stock (lookups, duplicados).
SKIP_SHEETS = {"base sku", "hoja2", "kappa calzado (2)"}


def stock_sheets(path):
    """Hojas de stock de un archivo (excluye auxiliares/duplicadas)."""
    return [s for s in ns.list_sheets(path) if s.strip().lower() not in SKIP_SHEETS]


def normalizar_kappa(files=None, out_dir=OUT_DIR, thumb_px=512, quality=85,
                     online=False):
    files = files or sorted(glob.glob(os.path.join(IN_DIR, "*.xlsx")))
    files = [f for f in files if not os.path.basename(f).startswith("~$")
             and "NORMALIZADO" not in os.path.basename(f).upper()]
    if not files:
        print("No encontré archivos en kappa/.")
        return []

    ns.THUMB_PX = thumb_px
    ns.THUMB_QUALITY = quality
    if online:
        # Fuente online = tienda oficial Kappa (misma mecánica que Reebok, con
        # otro CDN). El código de imagen es la col "Modelo Color".
        ns.REEBOK_CDN = KAPPA_CDN
        print("Modo ONLINE: bajando de kappastore.com.ar (fallback embebido)")

    os.makedirs(out_dir, exist_ok=True)
    salidas = []
    for f in files:
        sheets = stock_sheets(f)
        if not sheets:
            print(f"[!] sin hojas de stock: {f}")
            continue
        base = os.path.splitext(os.path.basename(f))[0]
        out = os.path.join(out_dir, base + " NORMALIZADO.xlsx")
        print(f"\n######## {os.path.basename(f)}  (hojas: {sheets})\n   -> {out}")
        args = SimpleNamespace(input=f, output=out, source="reebok",
                               no_images=False, image_mode="embed",
                               online=online, limit=None, sheets=sheets)
        ns.cmd_run(args)
        salidas.append(out)

    print(f"\n=== LISTO: {len(salidas)} archivo(s) en '{out_dir}'")
    return salidas


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("files", nargs="*", help="Archivos a normalizar (default: auto)")
    ap.add_argument("--out-dir", default=OUT_DIR)
    ap.add_argument("--thumb-px", type=int, default=512)
    ap.add_argument("--quality", type=int, default=85)
    ap.add_argument("--online", action="store_true",
                    help="Bajar fotos de kappastore.com.ar (fallback embebido)")
    a = ap.parse_args()
    normalizar_kappa(a.files or None, a.out_dir, a.thumb_px, a.quality, a.online)


if __name__ == "__main__":
    main()
