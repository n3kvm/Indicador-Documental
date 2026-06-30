# Indicador Documental en GitHub Actions

Sube el contenido de esta carpeta al repositorio GitHub.

## Archivos incluidos

- `.github/workflows/dashboard-github-pages.yml`
- `github_actions/generate_dashboard.py`
- `github_actions/requirements.txt`
- `github_actions/README.md`
- `work/*.py`

No se incluyen PDFs, Excels descargados, caches ni resultados locales.

## Secretos requeridos

En GitHub:

`Settings > Secrets and variables > Actions > New repository secret`

Crea:

- `TENANT_ID`
- `CLIENT_ID`
- `CLIENT_SECRET`
- `SUPPORTS_FOLDER_URL`
- `CRONOGRAMA_URL`

La App Registration debe tener `Sites.Selected` y permiso `Read` sobre el sitio espejo de Brillaseo.

## Ejecutar

1. Entra a `Actions`.
2. Selecciona `Dashboard mantenimiento`.
3. Clic en `Run workflow`.
4. Indica anio y mes, o deja vacio para mes actual.
5. Descarga el artifact `dashboard-mantenimiento`.
