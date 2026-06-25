# Dashboard con GitHub Actions

Esta alternativa ejecuta el dashboard en la nube de GitHub, sin depender del equipo local ni de una suscripcion Azure.

## Como funciona

1. GitHub Actions se ejecuta manualmente o todos los dias a las 2:00 p. m. Colombia.
2. Consulta SharePoint/OneDrive con Microsoft Graph.
3. Descarga cronograma y PDFs.
4. Ejecuta los mismos scripts Python del proyecto.
5. Deja el HTML y reportes como artifact descargable.

## Requisitos

- Repositorio GitHub, idealmente privado.
- GitHub Actions habilitado.
- App Registration en Microsoft Entra ID con:
  - `Sites.Read.All`
  - `Files.Read.All`
  - Admin consent aprobado.

## Secretos del repositorio

En GitHub:

`Settings` -> `Secrets and variables` -> `Actions` -> `New repository secret`

Crea estos secretos:

| Secret | Que es |
| --- | --- |
| `TENANT_ID` | Tenant ID de Microsoft 365 |
| `CLIENT_ID` | Application ID del App Registration |
| `CLIENT_SECRET` | Secreto del App Registration |
| `SUPPORTS_FOLDER_URL` | URL de la carpeta de soportes |
| `CRONOGRAMA_URL` | URL del cronograma o carpeta del cronograma |

Opcionalmente crea esta variable:

`Settings` -> `Secrets and variables` -> `Actions` -> `Variables`

| Variable | Uso |
| --- | --- |
| `DASHBOARD_ROUTE_NAME` | Sufijo visual del dashboard, ejemplo `JUNIO` o `CONTINGENCIA` |

## Ejecutar manualmente

1. Ve a `Actions`
2. Selecciona `Dashboard mantenimiento`
3. Clic en `Run workflow`
4. Puedes dejar anio/mes vacios para usar el mes actual de Colombia
5. O puedes escribir un periodo especifico, por ejemplo `2026` y `6`
6. Cuando termine, entra a la ejecucion y descarga el artifact `dashboard-mantenimiento`

## Actualizacion automatica

El workflow corre todos los dias a las `19:00 UTC`, equivalente a `2:00 p. m.` Colombia.

Archivo del workflow:

`.github/workflows/dashboard-github-pages.yml`

## Descargar resultado

Despues de cada ejecucion:

1. Entra a `Actions`
2. Abre la ejecucion finalizada
3. Baja hasta `Artifacts`
4. Descarga `dashboard-mantenimiento`
5. Descomprime el ZIP
6. Abre `index.html`

El artifact incluye:

- `index.html`
- `metadata.json`
- `validacion_soportes_sharepoint.xlsx`
- `auditoria_soportes.xlsx`

## Seguridad

Los secretos no quedan dentro del HTML. GitHub Actions los usa en el servidor para consultar Microsoft Graph.

Importante: los PDFs se procesan temporalmente dentro de GitHub Actions. Confirma con la compania si esta permitido procesar esta informacion fuera de Microsoft 365.
