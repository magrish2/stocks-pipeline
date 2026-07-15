#!/usr/bin/env python3
"""
Retoma la mejora HD de los Kappa: procesa SOLO los '… NORMALIZADO.xlsx' que
todavía no tienen su '… NORMALIZADO HD.xlsx' válido.

Es idempotente y seguro: si la Mac se durmió a mitad de un archivo, ese HD
no quedó escrito (se escribe atómico al final), así que lo rehace. Corré esto
las veces que quieras hasta que diga "TODO HD LISTO".

    python retomar_hd.py
"""
import glob
import os
import sys
import zipfile

import imagenes_hd as h

DIR = "kappa/Stocks normalizados"
MAX_PX = 800


def es_xlsx_valido(path):
    if not os.path.exists(path):
        return False
    try:
        with zipfile.ZipFile(path) as z:
            return z.testzip() is None
    except Exception:
        return False


def main():
    if not h.have_engine():
        sys.exit(f"Falta el motor Real-ESRGAN en {h.REALESRGAN_BIN}")

    files = sorted(glob.glob(os.path.join(DIR, "*NORMALIZADO.xlsx")))
    files = [f for f in files if " HD.xlsx" not in f
             and not os.path.basename(f).startswith("~$")]

    pendientes = []
    for f in files:
        hd = os.path.splitext(f)[0] + " HD.xlsx"
        if es_xlsx_valido(hd):
            print("✓ ya está:", os.path.basename(hd))
        else:
            pendientes.append((f, hd))

    print(f"\nPendientes: {len(pendientes)} de {len(files)}")
    for f, hd in pendientes:
        print(f"\n== {os.path.basename(f)} ==")
        tmp = hd + ".part"
        h.enhance_excel(f, tmp, max_px=MAX_PX)   # escribe a .part
        os.replace(tmp, hd)                      # rename atómico -> HD final
        print("  guardado:", os.path.basename(hd))

    print("\n=== TODO HD LISTO ===")


if __name__ == "__main__":
    main()
