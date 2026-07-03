# Pipeline de stocks (Google Drive)

Toma los Excel **crudos** de una carpeta de Drive, los **normaliza**, los deja en
la carpeta **normalizados** y mantiene un **maestro** por stock en la carpeta
**fijos** (actualizándolo en su lugar: cantidades, altas, bajas; conservando el
Pedido cargado a mano).

## Flujo

```
[CRUDOS] --normaliza--> [NORMALIZADOS]   (copia completa de cada crudo)
                 └─ 1ra vez  --> [FIJOS]  crea "MAESTRO <marca_tipo>.xlsx"
                 └─ 2da+     --> actualiza ese maestro EN SU LUGAR:
                                  • cantidades de lo que sigue (del crudo nuevo)
                                  • sin stock (0/vacío)  -> se elimina
                                  • SKU nuevo            -> se agrega
                                  • conserva la columna Pedido (manual) por SKU
```

El match crudo↔maestro es por **nombre** (marca + tipo): p. ej. cualquier
`...promo Reebok 40 %...` cae en `reebok_promo_40`, sin importar la fecha.

**Marcas soportadas** (se detectan por el nombre del archivo):
- **Reebok** → fotos de reebok.com.ar (CDN por Modelo).
- **Kappa** → fotos de kappastore.com.ar (CDN por Modelo-Color).
- **Crocs** → fotos de crocs.com.ar (CDN por Modelo-Color) + catálogo
  `products.json` + `media.crocs.com` (multi-fuente, ~94% de cobertura).

## Setup (una sola vez)

1. **Credenciales de Google** (OAuth de escritorio):
   - Entrá a https://console.cloud.google.com/ → creá un proyecto.
   - *APIs y servicios* → *Biblioteca* → activá **Google Drive API**.
   - *APIs y servicios* → *Pantalla de consentimiento OAuth* → tipo **Externo**,
     completá lo mínimo y agregá tu mail como **usuario de prueba**.
   - *Credenciales* → *Crear credenciales* → **ID de cliente OAuth** →
     tipo **App de escritorio** → descargá el JSON.
   - Guardalo como `pipeline/credentials.json`.

2. **IDs de las 3 carpetas de Drive**:
   - Abrí cada carpeta en Drive; el ID es lo que va después de `/folders/` en la URL.
   - Copiá `config.example.json` a `config.json` y pegá los 3 IDs.

## Uso

```bash
source .venv/bin/activate
python pipeline/run_drive.py
```

La **primera** corrida abre el navegador para autorizar (se guarda `token.json`
y no vuelve a pedirlo). Cada corrida procesa todos los crudos y sube resultados.

## Automático (se dispara solo al agregar crudos)

Un servicio de macOS (`launchd`) revisa la carpeta CRUDOS cada 5 min y corre el
pipeline si hay algo nuevo. Corre **mientras la Mac esté prendida y con sesión
iniciada**.

```bash
# Instalar / activar
cp pipeline/com.stocks.pipeline.plist ~/Library/LaunchAgents/
launchctl load ~/Library/LaunchAgents/com.stocks.pipeline.plist

# Estado
launchctl list | grep stocks

# Parar / desactivar
launchctl unload ~/Library/LaunchAgents/com.stocks.pipeline.plist

# Logs
tail -f pipeline/watch.log        # resumen con fecha
tail -f pipeline/watch.out.log    # detalle de cada corrida
```

Cambiar el intervalo: editá `StartInterval` (segundos) en el plist y recargá
(`unload` + `load`). Para 24/7 sin depender de la Mac hay que ir a la nube.

### Probar sin Drive (carpetas locales)

```bash
python pipeline/orchestrator.py --crudos <dir> --normalizados <dir> --fijos <dir>
```

## Archivos

- `engine.py` — normaliza (Reebok/Kappa), con saltear-sin-stock y arrastrar-Pedido.
- `match.py` — clave marca_tipo desde el nombre de archivo.
- `sync.py` — lee el Pedido del maestro y calcula altas/bajas/cambios.
- `orchestrator.py` — flujo sobre carpetas (crudos→normalizados→fijos).
- `drive.py` — cliente Google Drive (bajar/subir/reemplazar en su lugar).
- `run_drive.py` — runner completo sobre Drive (usa `config.json`).

## Notas

- Lo "manual" que se conserva es la columna **Pedido**. Si querés conservar más
  (notas, precios editados a mano), se agrega a `sync.read_pedido_map`.
- Las fotos se regeneran desde la web (reebok.com.ar / kappastore.com.ar) con
  fallback a las embebidas; las correcciones puntuales van en `image_overrides.json`.
- Abrí los `.xlsx` con **Excel** (no Numbers) para no romper las imágenes.
```
