#!/usr/bin/env python3
"""Watcher LOCAL de carpeta: arrastrás los crudos a una carpeta de la Mac y se
normalizan solos (local, rápido, con el cache de imágenes persistente), y el
resultado se sube a Drive (normalizados + maestros) con el reintento de drive.py.

- Sondea la carpeta cada POLL_SECS. Procesa un archivo recién cuando su tamaño
  quedó ESTABLE (terminó de copiarse).
- Los crudos ya procesados se mueven a <drop>/_procesados (no se borran).
- Notificación nativa de macOS al terminar (ok / error).
- Un lock evita solaparse con una corrida anterior.

Pensado para correr siempre vía launchd (KeepAlive). También se puede correr a
mano:  .venv/bin/python pipeline/watch_local.py
"""
import datetime
import fcntl
import json
import os
import shutil
import subprocess
import sys
import time
import traceback

BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)
os.chdir(os.path.dirname(BASE))              # raíz del proyecto (config/paths/cache)

import drive
import run_drive

CONFIG = os.path.join(BASE, "config.json")
LOG = os.path.join(BASE, "watch_local.log")
LOCK = os.path.join(BASE, ".watch_local.lock")
POLL_SECS = 3
RAW_EXT = (".xlsx", ".xlsb")


def log(msg):
    ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}"
    with open(LOG, "a") as fh:
        fh.write(line + "\n")
    print(line, flush=True)


def notify(title, msg):
    """Notificación nativa de macOS (no rompe si falla)."""
    try:
        subprocess.run(
            ["osascript", "-e",
             f'display notification {json.dumps(msg)} with title {json.dumps(title)}'],
            check=False, timeout=10)
    except Exception:
        pass


def _crudos_en(drop):
    return [f for f in os.listdir(drop)
            if f.lower().endswith(RAW_EXT) and not f.startswith("~$")
            and os.path.isfile(os.path.join(drop, f))]


def _estables(drop, prev_sizes):
    """Archivos cuyo tamaño no cambió desde el sondeo anterior (ya terminaron
    de copiarse). Devuelve (lista_estables, sizes_actuales)."""
    sizes = {}
    for f in _crudos_en(drop):
        try:
            sizes[f] = os.path.getsize(os.path.join(drop, f))
        except OSError:
            sizes[f] = -1
    estables = [f for f, s in sizes.items()
                if s > 0 and prev_sizes.get(f) == s]
    return estables, sizes


def procesar(cfg, svc, drop, procesados_dir, local_norm):
    archivos = _crudos_en(drop)
    if not archivos:
        return
    log(f"Procesando {len(archivos)} crudo(s): {', '.join(archivos)}")
    try:
        results, _ = run_drive.process_and_upload(cfg, svc, drop, local_norm_dir=local_norm)
    except Exception:
        log("ERROR en la normalización/subida:\n" + traceback.format_exc())
        notify("❌ Normalización falló", "Mirá watch_local.log")
        return
    # mover crudos ya procesados (todo subió OK)
    os.makedirs(procesados_dir, exist_ok=True)
    for f in archivos:
        try:
            shutil.move(os.path.join(drop, f), os.path.join(procesados_dir, f))
        except Exception:
            log(f"  no pude mover {f} a _procesados")
    n_norm = sum(1 for r in results if r.get("norm"))
    log(f"Listo: {n_norm} normalizado(s) subido(s) a Drive.")
    notify("✅ Normalización lista", f"{n_norm} archivo(s) subido(s) a Drive")


def main():
    lf = open(LOCK, "w")
    try:
        fcntl.flock(lf, fcntl.LOCK_EX | fcntl.LOCK_NB)
    except BlockingIOError:
        log("Ya hay un watcher corriendo; salgo.")
        return

    cfg = json.load(open(CONFIG))
    drop = cfg.get("crudos_local") or os.path.join(os.getcwd(), "_CRUDOS (arrastrar aca)")
    procesados_dir = os.path.join(drop, "_procesados")
    local_norm = cfg.get("normalizados_local") or os.path.join(os.getcwd(), "Normalizados")
    os.makedirs(drop, exist_ok=True)
    svc = drive.service()
    log(f"Watcher activo. Arrastrá crudos a: {drop}")

    prev = {}
    try:
        while True:
            estables, sizes = _estables(drop, prev)
            if estables:
                # esperar un ciclo más de quietud total antes de arrancar
                time.sleep(POLL_SECS)
                estables2, sizes2 = _estables(drop, sizes)
                if _crudos_en(drop):
                    procesar(cfg, svc, drop, procesados_dir, local_norm)
                prev = {}
            else:
                prev = sizes
            time.sleep(POLL_SECS)
    except KeyboardInterrupt:
        log("Watcher detenido.")
    finally:
        fcntl.flock(lf, fcntl.LOCK_UN)
        lf.close()


if __name__ == "__main__":
    main()
