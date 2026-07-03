# Correr el pipeline 24/7 con GitHub Actions

El pipeline corre en la nube de GitHub cada ~15 min (sin depender de tu Mac).
Usa **tu propio login de Google** (no cuenta de servicio, que en Drive personal
no puede crear archivos). Pasos, una sola vez:

---

## A) Hacer permanente el token de Google (publicar la app)

Mientras la app OAuth está "En prueba", el token **vence a los 7 días** → no sirve
para 24/7. Hay que publicarla.

1. https://console.cloud.google.com/ → proyecto **stocks-501314**.
2. ☰ → **APIs y servicios** → **Pantalla de consentimiento de OAuth**
   (o *Google Auth Platform → Público*).
3. Botón **"Publicar app" / "Publish app"** → confirmar (queda "En producción").
   - No hace falta verificación de Google para uso personal; puede seguir
     mostrando el cartel de "no verificada", está OK.

4. **Regenerar el token** (ahora será permanente). En la terminal:
   ```bash
   cd "/Users/ginomagris/Desktop/Stocks DISTRINANDO"
   source .venv/bin/activate
   rm pipeline/token.json
   python pipeline/test_conexion.py     # abre el navegador, autorizá de nuevo
   ```
   Debe volver a listar las 3 carpetas. Ya tenés un `token.json` que no vence.

---

## B) Subir el código a un repo privado de GitHub

Lo más simple es con **GitHub Desktop** (https://desktop.github.com/):

1. Instalá y logueate con tu cuenta de GitHub (si no tenés, creala gratis).
2. **File → Add Local Repository** → elegí la carpeta
   `/Users/ginomagris/Desktop/Stocks DISTRINANDO` (ya está inicializada como repo).
3. Arriba, **Publish repository** → **dejá tildado "Keep this code private"** →
   Publish.

> Alternativa por terminal (si preferís): creá un repo vacío en github.com y luego
> `git remote add origin <url>` y `git push -u origin main`. Requiere un
> Personal Access Token como contraseña.

El `.gitignore` ya evita subir Excels, cachés y **los secretos**.

---

## C) Cargar los 2 secretos en GitHub

En el repo, en la web: **Settings → Secrets and variables → Actions → New
repository secret**. Creá dos:

1. **`GOOGLE_CREDENTIALS`** → pegá **todo** el contenido de
   `pipeline/credentials.json`.
2. **`GOOGLE_TOKEN`** → pegá **todo** el contenido de `pipeline/token.json`.

Para copiar el contenido al portapapeles:
```bash
cd "/Users/ginomagris/Desktop/Stocks DISTRINANDO"
cat pipeline/credentials.json | pbcopy     # y pegás en GOOGLE_CREDENTIALS
cat pipeline/token.json | pbcopy           # y pegás en GOOGLE_TOKEN
```

---

## D) Activar y probar

1. En el repo → pestaña **Actions** → si pide, **Enable workflows**.
2. Entrá al workflow **"Stock pipeline"** → **Run workflow** (botón) para probarlo
   ya mismo sin esperar el cron.
3. Miralo correr; al final debería decir "Pipeline completo". Subí un crudo a la
   carpeta CRUDOS de Drive y verificá que aparezca normalizado + maestro.

Desde ahí corre solo cada ~15 min. Para cambiar la frecuencia, editá el `cron`
en `.github/workflows/pipeline.yml`.

---

## Notas

- El servicio local de macOS (launchd) ya lo **apagamos**; ahora manda GitHub.
- Si alguna vez revocás el acceso o cambiás la contraseña de Google, regenerás el
  `token.json` (paso A.4) y actualizás el secret `GOOGLE_TOKEN`.
- Cada corrida arranca "en limpio" (sin caché de imágenes), así que un stock
  grande tarda unos minutos la primera vez. No pasa nada, corre en la nube.
