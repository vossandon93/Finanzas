# Instrucciones para Claude

- App personal de Vicente (solo él). Sigue las skills `crear-app`, `formato-app` y `cambios-apps-script` del repo `vossandon93/skills-vicente`.
- Cuando Vicente pide un cambio y ya está hecho y probado: crear el PR a `main` y hacer merge sin preguntar.
- Si el cambio toca `apps-script/Code.gs`: subir `VERSION` y recordar al final que hay que pegar el script nuevo y publicar una "Nueva versión".

## Cartolas (datos bancarios)

- **Nunca** subir cartolas, CSV ni capturas con datos reales a GitHub, ni copiarlos a archivos del repo. El `.gitignore` las bloquea; no forzarlo.
- Trabajar las cartolas en el scratchpad de la sesión, no dentro del repo.
- Para cada cartola: `python3 importar/normalizar.py movimientos <archivo> ...` (ver `importar/README.md`). Revisar el total de entradas y salidas con Vicente antes de cargar.
- Cuando se conozca el formato fijo de un banco, agregarlo a `FORMATOS` en `normalizar.py`.
- El CSV resultante se deja en la carpeta de Google Drive **"Finanzas – importar"** (conector de Google Drive). El script la revisa cada hora, o Vicente toca **Importar ahora** en la app.
- Las reglas de categorías viven en la hoja **Reglas** de la planilla, no en el código.
