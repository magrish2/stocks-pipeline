#!/usr/bin/env python3
"""Resolver de imágenes de Crocs (multi-fuente), portado de Stock_imagenes.py.

Orden: CDN Shopify de crocs.com.ar por Modelo-Color -> catálogo products.json
-> CDN media.crocs.com. Devuelve una URL de imagen o None."""
import re

CDN_MEDIA = ("https://media.crocs.com/images/"
             "f_auto%2Cq_auto%2Cw_900%2Ch_900%2Cc_pad%2Cb_transparent/products")


def build_catalog(session, log=print, max_pages=60):
    """Índice {SKU|ModeloColor(upper): url} del catálogo de crocs.com.ar."""
    cat, page = {}, 1
    while page <= max_pages:
        url = ("https://www.crocs.com.ar/collections/all/products.json"
               f"?limit=250&page={page}")
        try:
            r = session.get(url, timeout=15)
            if r.status_code != 200:
                break
            data = r.json()
        except Exception:
            break
        prods = data.get("products", [])
        if not prods:
            break
        for p in prods:
            imgs = p.get("images", [])
            fallback = imgs[0].get("src", "") if imgs else ""
            by_id = {i["id"]: i.get("src", "") for i in imgs if i.get("id")}
            for v in p.get("variants", []):
                sku = str(v.get("sku") or "").strip().upper()
                if not sku:
                    continue
                fi = v.get("featured_image") or {}
                u = fi.get("src", "") or by_id.get(fi.get("id", ""), "") or fallback
                if not u:
                    continue
                cat.setdefault(sku, u)
                m = re.match(r"^(C\d+-C\w+)-\S", sku)
                if m:
                    cat.setdefault(m.group(1), u)
        page += 1
    log(f"  catálogo crocs.com.ar: {len(cat)} claves")
    return cat


def make_finder(session, log=print):
    """Devuelve un finder(model_code, session) -> url|None (usa el catálogo)."""
    cat = build_catalog(session, log)

    def finder(mc, sess):
        if not mc:
            return None
        up = str(mc).strip().upper()
        # 1) CDN Shopify directo
        for n in ("1", "2", "3"):
            url = f"https://www.crocs.com.ar/cdn/shop/files/{up}-{n}.jpg"
            try:
                r = sess.head(url, timeout=10, allow_redirects=True)
                if r.status_code == 200 and \
                        r.headers.get("content-type", "").startswith("image"):
                    return url
            except Exception:
                pass
        # 2) catálogo
        if up in cat:
            return cat[up]
        # 3) media.crocs.com
        m = re.match(r"C(\d+)-C(\w+)", up)
        if m:
            prod, color = m.group(1), m.group(2)
            for suf in ("_ALT100", "_ALT1", ""):
                for fn in ("crocs.jpg", "jibbitz.jpg"):
                    url = f"{CDN_MEDIA}/{prod}_{color}{suf}/{fn}"
                    try:
                        r = sess.head(url, timeout=8, allow_redirects=True)
                        if r.status_code == 200 and \
                                r.headers.get("content-type", "").startswith("image"):
                            return url
                    except Exception:
                        pass
        return None

    return finder, len(cat)
