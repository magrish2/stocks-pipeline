#!/usr/bin/env python3
"""Ejecuta el pipeline si hay crudos. Pensado para correr en loop vía launchd.

- Usa un lock para no solaparse con una corrida previa aún en curso.
- Loguea con fecha a pipeline/watch.log (además del stdout que capture launchd).
"""
import datetime
import fcntl
import os
import sys
import traceback

BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)
os.chdir(os.path.dirname(BASE))          # raíz del proyecto (para config/paths)

LOCK = os.path.join(BASE, ".watch.lock")
LOG = os.path.join(BASE, "watch.log")


def log(msg):
    ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(LOG, "a") as fh:
        fh.write(f"[{ts}] {msg}\n")
    print(f"[{ts}] {msg}", flush=True)


def main():
    lf = open(LOCK, "w")
    try:
        fcntl.flock(lf, fcntl.LOCK_EX | fcntl.LOCK_NB)
    except BlockingIOError:
        log("Ya hay una corrida en curso; salteo.")
        return
    try:
        import run_drive
        log("Chequeando crudos...")
        run_drive.main()
        log("Fin del ciclo.")
    except Exception:
        log("ERROR:\n" + traceback.format_exc())
    finally:
        fcntl.flock(lf, fcntl.LOCK_UN)
        lf.close()


if __name__ == "__main__":
    main()
