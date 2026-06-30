# Boton Actualizar Dashboard

Este Worker permite que el boton **Actualizar dashboard** de GitHub Pages ejecute el workflow sin exponer el token en el HTML.
La pagina envia el periodo seleccionado en formato:

```json
{
  "anio": "2026",
  "mes": "06"
}
```

El Worker pasa esos valores al workflow como inputs `anio` y `mes`.

## Variables del Worker

Configura estas variables/secrets en Cloudflare Worker:

- `GITHUB_TOKEN`: token fino de GitHub con permiso `Actions: Read and write` sobre el repositorio.
- `GITHUB_OWNER`: `n3kvm`
- `GITHUB_REPO`: `Indicador-Documental`
- `GITHUB_WORKFLOW`: `dashboard-github-pages.yml`
- `GITHUB_REF`: `main`
- `ALLOWED_ORIGIN`: URL de GitHub Pages, por ejemplo `https://n3kvm.github.io`

## Variable en GitHub

En el repositorio, agrega esta variable:

- `DASHBOARD_REFRESH_ENDPOINT`: URL publica del Worker.

Luego ejecuta nuevamente el workflow para que el boton quede conectado al endpoint.
