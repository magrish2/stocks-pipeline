# Normalizador de stock Reebok + fotos MercadoLibre

Lee el archivo de stock (`.xlsb` o `.xlsx`), busca la foto de cada modelo en
MercadoLibre Argentina (API oficial) y genera un Excel normalizado con:
**Foto, SKU, Descripción, Disponibilidad, Precio mayorista, Precio público y
una columna Pedido** para completar. Procesa las hojas **CALZADO** e
**INDUMENTARIA**, una fila por talle/SKU.

## Instalación (ya hecha en esta carpeta)

```bash
python3 -m venv .venv
.venv/bin/pip install openpyxl pandas requests xlsxwriter pyxlsb pillow
```

Usá siempre `.venv/bin/python` para correr el script.

## 1) Credenciales de MercadoLibre (una sola vez)

1. Entrá a <https://developers.mercadolibre.com.ar/> con tu cuenta MELI y
   creá una aplicación ("Crear aplicación").
2. En **Redirect URI** poné `https://localhost:8080` (cualquier URL https sirve;
   no hace falta que exista, solo vas a copiar el `code` de la barra del navegador).
3. Copiá el **Client ID** y **Client Secret**.
4. En esta carpeta:
   ```bash
   cp meli_config.example.json meli_config.json
   ```
   y editá `meli_config.json` con tu `client_id`, `client_secret` y el mismo
   `redirect_uri` que cargaste en la app.

## 2) Autorizar la app (una sola vez)

```bash
.venv/bin/python normalizar_stock.py auth-url
```

Abrí la URL que imprime, autorizá, y el navegador te redirige a algo como
`https://localhost:8080/?code=TG-xxxxxxxx`. Copiá ese `code` y ejecutá:

```bash
.venv/bin/python normalizar_stock.py auth-exchange TG-xxxxxxxx
```

Esto guarda `meli_tokens.json`. El token se **refresca solo** después; no
tenés que repetir esto salvo que pase mucho tiempo sin usarlo.

## 3) Generar el Excel

```bash
.venv/bin/python normalizar_stock.py run
```

Toma automáticamente el archivo `Stock REEBOK inmediato*` más reciente
(carpeta actual, Escritorio o Documentos) y genera
`Stock_REEBOK_normalizado_AAAA-MM-DD.xlsx`.

### Opciones útiles

| Flag | Para qué |
|------|----------|
| `--input "ruta/archivo.xlsb"` | Usar un archivo puntual |
| `--output "salida.xlsx"` | Nombre del Excel de salida |
| `--no-images` | No consultar MELI; deja la columna Foto vacía (rápido) |
| `--limit 20` | Procesar solo 20 filas por hoja (prueba) |
| `--sheets CALZADO` | Procesar solo una hoja (repetible) |
| `--image-mode first-row` | Embeber la foto solo en la 1ª fila de cada modelo (Excel más liviano) |
| `--image-mode url` | En vez de foto, poner un hipervínculo de búsqueda en MELI |

Prueba recomendada la primera vez:
```bash
.venv/bin/python normalizar_stock.py run --limit 10
```

## Notas

- **Cache de fotos:** se guardan en `meli_cache.json` (modelo → URL) y en
  `.cache_img/`. Si volvés a correr el script no re-consulta MELI para modelos
  ya resueltos. Borrá esos para forzar una búsqueda nueva.
- **Búsqueda:** prioriza el código de fabricante de la descripción
  (`CN4107`, `GY0952`, `HS7769`). Si no está, busca por nombre del modelo.
  Es una búsqueda automática: revisá visualmente las fotos, alguna puede no
  ser exacta.
- **No compartas** `meli_config.json` ni `meli_tokens.json`: tienen tus
  credenciales.
- El `.xlsb` debe tener la fila de encabezados con "Número de artículo"
  (formato del archivo actual). Si cambia el formato, avisá.
```
