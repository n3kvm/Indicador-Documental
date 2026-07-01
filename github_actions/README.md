# Dashboard con GitHub Actions

Esta alternativa ejecuta el dashboard en la nube de GitHub, sin depender del equipo local ni de una suscripcion Azure.

## Como funciona

1. GitHub Actions se ejecuta manualmente o todos los dias a las 2:00 p. m. Colombia.
2. Consulta SharePoint/OneDrive con Microsoft Graph.
3. Descarga cronograma y PDFs.
4. Ejecuta los mismos scripts Python del proyecto.
5. Deja el HTML y reportes como artifact descargable en `public`.

## Requisitos

- Repositorio GitHub, idealmente privado.
- GitHub Actions habilitado.
- App Registration en Microsoft Entra ID con:
  - `Sites.Selected` de tipo aplicacion.
  - Admin consent aprobado.
- La App Registration debe estar asignada con rol `Read` sobre el sitio espejo de Brillaseo.

No se requiere `Sites.Read.All` ni `Files.Read.All` si `Sites.Selected` esta configurado correctamente sobre el sitio espejo.

## Secretos del repositorio

En GitHub:

`Settings` -> `Secrets and variables` -> `Actions` -> `New repository secret`

Crea estos secretos:

| Secret | Que es |
| --- | --- |
| `TENANT_ID` | Tenant ID de Microsoft 365 |
| `CLIENT_ID` | Application ID del App Registration |
| `CLIENT_SECRET` | Secreto del App Registration |
| `SUPPORTS_FOLDER_URL` | URL de la carpeta de soportes en el espejo Brillaseo |
| `CRONOGRAMA_URL` | URL del cronograma o carpeta del cronograma en el espejo Brillaseo |
| `ZOHO_REFRESH_TOKEN` | Refresh token OAuth de Zoho Creator |
| `ZOHO_CLIENT_ID` | Client ID OAuth de Zoho |
| `ZOHO_CLIENT_SECRET` | Client secret OAuth de Zoho |

Opcionalmente crea esta variable:

`Settings` -> `Secrets and variables` -> `Actions` -> `Variables`

| Variable | Uso |
| --- | --- |
| `DASHBOARD_ROUTE_NAME` | Sufijo visual del dashboard, ejemplo `JUNIO` o `CONTINGENCIA` |
| `ZOHO_CREATOR_BASE_URL` | Base API de Zoho. Por defecto `https://www.zohoapis.com` |
| `ZOHO_ACCOUNTS_URL` | URL de cuentas Zoho. Por defecto `https://accounts.zoho.com` |
| `ZOHO_ACCOUNT_OWNER` | Account owner name de la app Zoho Creator |
| `ZOHO_APP_LINK_NAME` | Link name de la aplicacion Zoho Creator |
| `ZOHO_REPORT_LINK_NAME` | Link name del reporte de autogestiones |
| `ZOHO_FIELD_AUTOGESTION` | Link name del campo No. autogestion. Por defecto `Autogestion` |
| `ZOHO_FIELD_SEDE` | Link name del campo sede. Por defecto `Sede` |
| `ZOHO_FIELD_UES` | Link name del campo UES. Por defecto `UES` |
| `ZOHO_FIELD_FECHA_REALIZACION` | Link name del campo fecha de realizacion. Por defecto `Fecha_Realizacion` |
| `ZOHO_FIELD_FECHA_REGISTRO` | Link name del campo fecha de registro. Por defecto `Added_Time` |

Si Zoho no esta configurado, el dashboard se genera igual, pero la vista `Falta subir aplicativo` queda sin datos del aplicativo.

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

## Refrescar el dashboard

La actualizacion se hace desde GitHub Actions o desde el Worker configurado para disparar el workflow. El HTML publicado no guarda tokens ni secretos.

## Publicar como GitHub Pages

Si el repositorio es publico, puedes publicar la carpeta `public` con GitHub Pages. En repositorios privados, GitHub Pages puede requerir plan pago o Enterprise.

La ruta mas simple es:

1. Ejecutar el workflow.
2. Descargar el artifact `dashboard-mantenimiento`.
3. Si quieres publicarlo fijo, subir el contenido de `public` a una rama o carpeta configurada para Pages.

Tambien se puede agregar un paso automatico de Pages, pero primero conviene validar que el workflow ya lee el espejo y genera `index.html`.

## Seguridad

Los secretos no quedan dentro del HTML. GitHub Actions los usa en el servidor para consultar Microsoft Graph.

Importante: los PDFs se procesan temporalmente dentro de GitHub Actions. Confirma con la compania si esta permitido procesar esta informacion fuera de Microsoft 365.
