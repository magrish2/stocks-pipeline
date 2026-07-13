#!/usr/bin/env python3
"""Matcheo crudo <-> maestro por nombre de archivo.

La clave = marca + tipo (+ variante de promo). Dos archivos con la misma clave
son el mismo stock (aunque cambie la fecha del nombre)."""
import os
import re
import unicodedata


def _norm(s):
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", s.lower()).strip()


def key_for(filename):
    """Devuelve una clave canónica (marca_tipo) a partir del nombre de archivo."""
    n = _norm(os.path.basename(filename))
    n = re.sub(r"\.(xlsx|xlsb)$", "", n)

    if "kappa" in n:
        marca = "kappa"
    elif "croc" in n:
        marca = "crocs"
    elif "reebok" in n or "rbk" in n:
        marca = "reebok"
    else:
        marca = "otro"

    # Tipo (orden = prioridad)
    if "postdatado" in n:
        tipo = "postdatado"
    elif "3 x 1" in n or "3x1" in n:
        tipo = "promo_3x1"
    elif "dia del padre" in n or "día del padre" in n:
        tipo = "promo_diapadre"
    elif re.search(r"\b40\b", n) or "40 %" in n or "40%" in n:
        tipo = "promo_40"
    elif "50" in n and "promo" in n or "indumentaria 50" in n:
        tipo = "promo_50"
    elif "temp" in n and "anterior" in n:
        tipo = "promo_tempant"
    elif "promo" in n and "calzado" in n:
        tipo = "promo_calzado"
    elif "promo" in n and "indumentaria" in n:
        tipo = "promo_indumentaria"
    elif "promo" in n:
        tipo = "promo"
    elif "inmediato" in n:
        tipo = "inmediato"
    else:
        tipo = None

    # Marca desconocida en el nombre (ej. clubes: Atl. Tucumán, Huracán, Tigre):
    # usamos un slug del nombre para que cada producto tenga su propio maestro.
    if marca == "otro":
        tipo = _slug(n)
    elif tipo is None:
        tipo = "general"

    return f"{marca}_{tipo}"


def _slug(n):
    """Nombre distintivo a partir del filename (saca 'stock', fechas, símbolos)."""
    s = re.sub(r"\bstock\b", " ", n)
    s = re.sub(r"\d{1,2}[-/.]\d{1,2}([-/.]\d{2,4})?", " ", s)   # fechas
    s = re.sub(r"[^a-z0-9]+", "-", s).strip("-")
    return s or "general"


if __name__ == "__main__":
    import sys
    for f in sys.argv[1:]:
        print(f"{key_for(f):22s} <- {f}")
