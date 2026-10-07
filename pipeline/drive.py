#!/usr/bin/env python3
"""Cliente mínimo de Google Drive (OAuth de usuario) para el pipeline.

Necesita: credentials.json (OAuth client de escritorio) en la carpeta pipeline/.
La primera vez abre el navegador para autorizar y guarda token.json.
"""
import io
import os
import random
import time

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload


def _retry(fn, tries=6, base=2.0):
    """Ejecuta fn() reintentando ante cortes transitorios de red/SSL (el
    SSLEOFError que corta las subidas grandes a Drive desde GitHub Actions) o
    errores 5xx/429. Backoff exponencial con jitter. Los 4xx no se reintentan."""
    last = None
    for i in range(tries):
        try:
            return fn()
        except HttpError as e:
            status = getattr(getattr(e, "resp", None), "status", None)
            if status and int(status) < 500 and int(status) != 429:
                raise                       # error permanente (4xx) → no reintentar
            last = e
        except OSError as e:                # ssl.SSLError/EOF, ConnectionError, etc.
            last = e
        if i == tries - 1:
            raise last
        time.sleep(min(30.0, base * (2 ** i)) + random.uniform(0, 1))

BASE = os.path.dirname(os.path.abspath(__file__))
SCOPES = ["https://www.googleapis.com/auth/drive"]
CREDENTIALS = os.path.join(BASE, "credentials.json")
TOKEN = os.path.join(BASE, "token.json")

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
XLSB_MIME = "application/vnd.ms-excel.sheet.binary.macroEnabled.12"
GSHEET_MIME = "application/vnd.google-apps.spreadsheet"


def service():
    creds = None
    if os.path.exists(TOKEN):
        creds = Credentials.from_authorized_user_file(TOKEN, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())        # headless: renueva sin navegador
        elif os.environ.get("GITHUB_ACTIONS") or os.environ.get("PIPELINE_HEADLESS"):
            raise RuntimeError(
                "Sin token válido en modo headless. Regenerá pipeline/token.json "
                "localmente (app en 'producción') y actualizá el secret GOOGLE_TOKEN.")
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS, SCOPES)
            creds = flow.run_local_server(port=0)
        try:
            with open(TOKEN, "w") as fh:
                fh.write(creds.to_json())   # en CI es efímero; el refresh_token no cambia
        except OSError:
            pass
    return build("drive", "v3", credentials=creds)


def list_files(svc, folder_id):
    """[(id, name, mimeType)] de los archivos (no carpetas) de una carpeta."""
    out, page = [], None
    q = f"'{folder_id}' in parents and trashed=false"
    while True:
        resp = svc.files().list(
            q=q, fields="nextPageToken, files(id,name,mimeType)",
            pageSize=1000, pageToken=page,
            supportsAllDrives=True, includeItemsFromAllDrives=True).execute()
        out += [(f["id"], f["name"], f["mimeType"]) for f in resp.get("files", [])]
        page = resp.get("nextPageToken")
        if not page:
            break
    return out


def get_or_create_folder(svc, parent_id, name):
    """Devuelve el id de la subcarpeta `name` dentro de `parent_id` (la crea si no está)."""
    q = (f"'{parent_id}' in parents and name = '{name}' and "
         "mimeType = 'application/vnd.google-apps.folder' and trashed = false")
    r = svc.files().list(q=q, fields="files(id,name)", supportsAllDrives=True,
                         includeItemsFromAllDrives=True).execute()
    if r.get("files"):
        return r["files"][0]["id"]
    meta = {"name": name, "parents": [parent_id],
            "mimeType": "application/vnd.google-apps.folder"}
    return svc.files().create(body=meta, fields="id",
                              supportsAllDrives=True).execute()["id"]


def download(svc, file_id, dest, mime=None):
    """Baja un archivo. Si es Google Sheet, lo exporta a xlsx."""
    if mime == GSHEET_MIME:
        req = svc.files().export_media(fileId=file_id, mimeType=XLSX_MIME)
    else:
        req = svc.files().get_media(fileId=file_id, supportsAllDrives=True)
    with io.FileIO(dest, "wb") as fh:
        dl = MediaIoBaseDownload(fh, req)
        done = False
        while not done:
            _, done = dl.next_chunk()
    return dest


def upload_new(svc, path, folder_id, mime=XLSX_MIME):
    meta = {"name": os.path.basename(path), "parents": [folder_id]}

    def _do():
        media = MediaFileUpload(path, mimetype=mime, resumable=True)
        return svc.files().create(body=meta, media_body=media, fields="id",
                                  supportsAllDrives=True).execute()["id"]
    return _retry(_do)


def update_content(svc, file_id, path, mime=XLSX_MIME):
    """Reemplaza el contenido de un archivo existente (mismo id = en su lugar)."""
    def _do():
        media = MediaFileUpload(path, mimetype=mime, resumable=True)
        return svc.files().update(fileId=file_id, media_body=media,
                                  supportsAllDrives=True).execute()["id"]
    return _retry(_do)


def trash(svc, file_id):
    """Manda el archivo a la Papelera. Tolerante a que ya no esté (corridas
    solapadas): si falla, no rompe."""
    try:
        return svc.files().update(fileId=file_id, body={"trashed": True},
                                  supportsAllDrives=True).execute()["id"]
    except Exception:
        return None


def upsert_by_name(svc, path, folder_id, existing=None):
    """Sube si no existe (por nombre en la carpeta) o reemplaza contenido si sí.
    Si hubiera varias copias con el mismo nombre, actualiza UNA y manda el resto
    a la Papelera (evita que se acumulen duplicados)."""
    name = os.path.basename(path)
    existing = existing if existing is not None else list_files(svc, folder_id)
    matches = [fid for fid, fname, _m in existing if fname == name]
    if matches:
        update_content(svc, matches[0], path)
        for extra in matches[1:]:
            trash(svc, extra)              # dedupe: borra copias sobrantes
        return matches[0], False
    return upload_new(svc, path, folder_id), True
