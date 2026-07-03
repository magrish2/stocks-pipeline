#!/usr/bin/env python3
"""Motor de normalización: envuelve normalizar_stock para Reebok y Kappa,
con soporte de sync de maestro (saltear sin stock + arrastrar Pedido)."""
import os
import sys
from types import SimpleNamespace

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import normalizar_stock as ns
import normalizar_kappa as nk

REEBOK_CDN = "https://reebok.com.ar/cdn/shop/files/{name}?width=1024"
KAPPA_CDN = "https://www.kappastore.com.ar/cdn/shop/files/{name}?width=1024"


def detect_brand(path):
    n = os.path.basename(path).lower()
    if "kappa" in n:
        return "kappa"
    if "croc" in n:
        return "crocs"
    return "reebok"


def _sheets_for(path, brand):
    if brand == "kappa":
        return nk.stock_sheets(path)          # excluye BASE SKU / Hoja2 / (2)
    return None                                # Reebok: cmd_run elige CALZADO/INDUMENTARIA


def normalize(raw, out_path, skip_out_of_stock=False, carry=None, thumb_px=512):
    """Normaliza `raw` a `out_path`. Devuelve out_path."""
    brand = detect_brand(raw)
    ns.THUMB_PX = thumb_px
    ns.IMAGE_URL_FINDER = None
    if brand == "kappa":
        ns.REEBOK_CDN = KAPPA_CDN
    elif brand == "crocs":
        # Crocs no tiene un CDN "adivinable" al 100%: usamos finder multi-fuente.
        import requests
        import crocs
        sess = requests.Session()
        sess.headers.update({"User-Agent": "Mozilla/5.0"})
        finder, _n = crocs.make_finder(sess)
        ns.IMAGE_URL_FINDER = finder
    else:
        ns.REEBOK_CDN = REEBOK_CDN
    args = SimpleNamespace(
        input=raw, output=out_path, source="reebok", no_images=False,
        image_mode="embed", online=True, limit=None,
        sheets=_sheets_for(raw, brand),
        skip_out_of_stock=skip_out_of_stock, carry=carry)
    ns.cmd_run(args)
    return out_path
