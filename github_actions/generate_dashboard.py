import base64
import json
import os
import posixpath
import re
import shutil
import subprocess
import sys
import tempfile
import unicodedata
from datetime import datetime
from pathlib import Path
from urllib.parse import parse_qs, quote, unquote, urlparse
from zoneinfo import ZoneInfo

import msal
import requests


MONTHS = {
    1: "ENERO",
    2: "FEBRERO",
    3: "MARZO",
    4: "ABRIL",
    5: "MAYO",
    6: "JUNIO",
    7: "JULIO",
    8: "AGOSTO",
    9: "SEPTIEMBRE",
    10: "OCTUBRE",
    11: "NOVIEMBRE",
    12: "DICIEMBRE",
}


def required_env(name):
    value = os.environ.get(name)
    if not value or not value.strip():
        raise RuntimeError(f"Falta configurar el secreto o variable {name}")
    return value.strip()


def decode_jwt_payload(token):
    try:
        payload = token.split(".")[1]
        payload += "=" * (-len(payload) % 4)
        return json.loads(base64.urlsafe_b64decode(payload.encode("ascii")).decode("utf-8"))
    except Exception:
        return {}


def period_from_env():
    now = datetime.now(ZoneInfo("America/Bogota"))
    year = int(os.environ.get("DASHBOARD_YEAR") or now.year)
    month = int(os.environ.get("DASHBOARD_MONTH") or now.month)
    return year, month


class GraphClient:
    def __init__(self):
        tenant = required_env("TENANT_ID")
        client_id = required_env("CLIENT_ID")
        client_secret = required_env("CLIENT_SECRET")
        app = msal.ConfidentialClientApplication(
            client_id,
            authority=f"https://login.microsoftonline.com/{tenant}",
            client_credential=client_secret,
        )
        token = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
        if "access_token" not in token:
            raise RuntimeError(f"No se pudo obtener token Graph: {token}")
        self.headers = {"Authorization": f"Bearer {token['access_token']}"}
        payload = decode_jwt_payload(token["access_token"])
        roles = payload.get("roles") or []
        print(f"Token Graph obtenido. Tenant token: {payload.get('tid', 'N/D')}. App: {payload.get('appid', 'N/D')}.")
        print(f"Permisos Graph en token: {', '.join(roles) if roles else 'Sin roles de aplicacion'}")
        if "Sites.Selected" in roles:
            print("Modo Graph: Sites.Selected. La app solo podra leer sitios asignados explicitamente con rol Read.")
        else:
            print(
                "ADVERTENCIA: el token no muestra Sites.Selected. "
                "Si usas menor privilegio, revisa API permissions y Admin consent en Microsoft Entra."
            )

    def get_json(self, url):
        if not self.headers.get("Authorization", "").replace("Bearer", "").strip():
            raise RuntimeError("Token Graph vacio: revisa TENANT_ID, CLIENT_ID y CLIENT_SECRET en GitHub Secrets.")
        res = requests.get(url, headers=self.headers, timeout=90)
        if not res.ok:
            extra = ""
            if res.status_code in (401, 403):
                extra = (
                    "\n\nPosibles causas:"
                    "\n- TENANT_ID no corresponde al tenant donde vive el espejo."
                    "\n- La App Registration no tiene Sites.Selected concedido con admin consent."
                    "\n- Falta asignar la app con rol Read sobre el sitio espejo."
                    "\n- El enlace pertenece a otro tenant o al OneDrive de otra organizacion."
                    "\n- Si usas Sites.Selected, evita enlaces compartidos /shares y usa URL directa del sitio espejo."
                )
            raise RuntimeError(f"Graph HTTP {res.status_code}: {res.text[:1200]}{extra}")
        return res.json()
    def post_json(self, url, body):
        res = requests.post(url, headers={**self.headers, "Content-Type": "application/json"}, json=body, timeout=90)
        if not res.ok:
            raise RuntimeError(f"Graph POST HTTP {res.status_code}: {res.text[:1200]}")
        return res.json()

    def put_bytes(self, url, data, content_type="application/octet-stream"):
        res = requests.put(url, headers={**self.headers, "Content-Type": content_type}, data=data, timeout=300)
        if not res.ok:
            raise RuntimeError(f"Graph PUT HTTP {res.status_code}: {res.text[:1200]}")
        return res.json()

    def site_and_drive_from_url(self, site_url):
        parsed = urlparse(site_url)
        parts = [p for p in unquote(parsed.path).split("/") if p]
        site_index = next((i for i, p in enumerate(parts) if p.lower() in {"sites", "teams", "personal"}), None)
        if site_index is None or site_index + 1 >= len(parts):
            site_path = self.configured_site_path(parsed.netloc)
            if not site_path:
                raise RuntimeError(f"No pude detectar el sitio de publicacion en {site_url}")
        else:
            site_path = "/" + "/".join(parts[site_index:site_index + 2])
        site = self.get_json(f"https://graph.microsoft.com/v1.0/sites/{parsed.netloc}:{site_path}")
        drive = self.get_json(f"https://graph.microsoft.com/v1.0/sites/{site['id']}/drive")
        return site, drive

    def ensure_folder(self, site_url, folder_path):
        _, drive = self.site_and_drive_from_url(site_url)
        current = self.get_json(f"https://graph.microsoft.com/v1.0/drives/{drive['id']}/root")
        for segment in [p for p in folder_path.replace("\\", "/").split("/") if p.strip()]:
            children = self.children(current)
            match = next((item for item in children if "folder" in item and item.get("name", "").lower() == segment.lower()), None)
            if not match:
                match = self.post_json(
                    f"https://graph.microsoft.com/v1.0/drives/{drive['id']}/items/{current['id']}/children",
                    {"name": segment, "folder": {}, "@microsoft.graph.conflictBehavior": "replace"},
                )
            current = match
        return current

    def upload_to_folder(self, folder_item, local_file, remote_name=None):
        remote_name = remote_name or Path(local_file).name
        drive_id = folder_item["parentReference"]["driveId"]
        folder_id = folder_item["id"]
        encoded_name = quote(remote_name, safe="")
        data = Path(local_file).read_bytes()
        uploaded = self.put_bytes(
            f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{folder_id}:/{encoded_name}:/content",
            data,
            "text/html; charset=utf-8" if remote_name.lower().endswith(".html") else "application/octet-stream",
        )
        return uploaded

    def drive_item_from_share_url(self, share_url):
        parsed = urlparse(share_url or "")
        # Los enlaces compartidos de SharePoint/OneDrive vienen como /:f:/s/... o /:x:/r/...
        # Esos no son rutas reales del drive; se resuelven con el endpoint /shares.
        if parsed.path.startswith("/:"):
            return self.drive_item_from_encoded_share(share_url)
        if parsed.netloc.endswith(".sharepoint.com"):
            return self.drive_item_from_sharepoint_url(share_url)
        return self.drive_item_from_encoded_share(share_url)

    def drive_item_from_encoded_share(self, share_url):
        share_id = "u!" + base64.urlsafe_b64encode(share_url.encode("utf-8")).decode("ascii").rstrip("=")
        return self.get_json(f"https://graph.microsoft.com/v1.0/shares/{share_id}/driveItem")

    def configured_site_path(self, hostname):
        configured = os.environ.get("SHAREPOINT_SITE_PATH") or os.environ.get("SHAREPOINT_SITE_URL") or ""
        if configured:
            parsed = urlparse(configured)
            path_value = parsed.path if parsed.scheme else configured
            path_value = unquote(path_value).strip("/")
            parts = [p for p in path_value.split("/") if p]
            site_index = next((i for i, p in enumerate(parts) if p.lower() in {"sites", "teams", "personal"}), None)
            if site_index is not None and site_index + 1 < len(parts):
                return "/" + "/".join(parts[site_index:site_index + 2])
        if hostname.lower() == "brillaseo2.sharepoint.com":
            return "/sites/SoportesEspejo"
        return ""

    def drive_item_from_sharepoint_url(self, share_url):
        parsed = urlparse(share_url)
        raw_path = parse_qs(parsed.query).get("id", [""])[0]
        folder_path = unquote(raw_path) if raw_path else unquote(parsed.path)
        if "/Forms/" in folder_path:
            folder_path = folder_path.split("/Forms/", 1)[0]
        parts = [p for p in folder_path.split("/") if p]
        site_index = next((i for i, p in enumerate(parts) if p.lower() in {"sites", "teams", "personal"}), None)
        if site_index is None or site_index + 1 >= len(parts):
            site_path = self.configured_site_path(parsed.netloc)
            if not site_path:
                raise RuntimeError(
                    "No pude detectar el sitio en SUPPORTS_FOLDER_URL/CRONOGRAMA_URL. "
                    "Configura SHAREPOINT_SITE_PATH como /sites/SoportesEspejo o usa una URL directa del sitio."
                )
            rel_parts = parts
        else:
            site_path = "/" + "/".join(parts[site_index:site_index + 2])
            rel_parts = parts[site_index + 2:]
        site = self.get_json(f"https://graph.microsoft.com/v1.0/sites/{parsed.netloc}:{site_path}")
        drive = self.get_json(f"https://graph.microsoft.com/v1.0/sites/{site['id']}/drive")
        if rel_parts and rel_parts[0].lower() in {"shared documents", "documentos compartidos", "documents"}:
            rel_parts = rel_parts[1:]
        rel_path = "/".join(rel_parts).strip("/")
        root = self.get_json(f"https://graph.microsoft.com/v1.0/drives/{drive['id']}/root")
        if not rel_path:
            return root
        print(f"Resolviendo ruta en drive por segmentos: {rel_path}")
        return self.drive_item_by_segments(root, rel_parts)

    def drive_item_by_segments(self, root_item, rel_parts):
        current = root_item
        ignored = {"shared documents", "documentos compartidos", "documents"}
        for segment in [p for p in rel_parts if p and p.lower() not in ignored]:
            children = self.children(current)
            match = next((item for item in children if item.get("name", "").lower() == segment.lower()), None)
            if not match:
                available = " | ".join(sorted(item.get("name", "") for item in children)[:40])
                raise RuntimeError(f"No encontre '{segment}' dentro de '{current.get('name', 'root')}'. Disponibles: {available}")
            current = match
        return current

    def children(self, item):
        drive_id = item["parentReference"]["driveId"]
        item_id = item["id"]
        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/children?$top=200"
        out = []
        while url:
            data = self.get_json(url)
            out.extend(data.get("value", []))
            url = data.get("@odata.nextLink")
        return out

    def list_recursive(self, folder_item):
        out = []
        stack = [folder_item]
        while stack:
            current = stack.pop()
            for child in self.children(current):
                if "folder" in child:
                    stack.append(child)
                else:
                    out.append(child)
        return out

    def download(self, item, dest):
        url = item.get("@microsoft.graph.downloadUrl")
        if not url:
            drive_id = item["parentReference"]["driveId"]
            item_id = item["id"]
            item = self.get_json(f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}")
            url = item.get("@microsoft.graph.downloadUrl")
        if not url:
            raise RuntimeError(f"Graph no entrego downloadUrl para {item.get('name')}")
        res = requests.get(url, timeout=300)
        if not res.ok:
            raise RuntimeError(f"Descarga HTTP {res.status_code}: {res.text[:500]}")
        dest.parent.mkdir(parents=True, exist_ok=True)
        dest.write_bytes(res.content)
        return len(res.content)


def safe_name(name):
    bad = '<>:"/\\|?*'
    return " ".join("".join("_" if ch in bad else ch for ch in name).split())


def normalize_folder_name(value):
    text = unicodedata.normalize("NFKD", str(value or ""))
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = text.upper()
    text = re.sub(r"[^A-Z0-9]+", "", text)
    return text


def find_child_folder(graph, parent, expected_names, label):
    expected = {normalize_folder_name(name) for name in expected_names if str(name or "").strip()}
    children = graph.children(parent)
    folders = [item for item in children if "folder" in item]
    for item in folders:
        if normalize_folder_name(item.get("name", "")) in expected:
            return item
    available = " | ".join(item.get("name", "") for item in folders) or "sin subcarpetas"
    wanted = " | ".join(expected_names)
    raise RuntimeError(f"No encontre la carpeta {label}. Esperaba: {wanted}. Disponibles: {available}")


def folder_name_matches(item, expected_names):
    current = normalize_folder_name(item.get("name", ""))
    expected = {normalize_folder_name(name) for name in expected_names if str(name or "").strip()}
    return current in expected


def resolve_support_period_folder(graph, supports_root, year, month, month_name):
    month_candidates = [
        f"{month:02d}-{month_name}",
        f"{month:02d} - {month_name}",
        f"{month:02d}_{month_name}",
        f"{month:02d} {month_name}",
        month_name,
        str(month),
        f"{month:02d}",
    ]
    year_candidates = [str(year)]

    root_name = supports_root.get("name", "")
    if folder_name_matches(supports_root, month_candidates):
        print(f"La URL de soportes ya apunta al mes solicitado: {root_name}")
        return supports_root

    if folder_name_matches(supports_root, year_candidates):
        print(f"La URL de soportes apunta al anio solicitado: {root_name}")
        return find_child_folder(graph, supports_root, month_candidates, f"del mes {month:02d}-{month_name}")

    root_children = [item for item in graph.children(supports_root) if "folder" in item]

    direct_month = next((item for item in root_children if folder_name_matches(item, month_candidates)), None)
    if direct_month:
        print(f"Mes encontrado directamente dentro de la URL base: {direct_month.get('name')}")
        return direct_month

    year_folder = next((item for item in root_children if folder_name_matches(item, year_candidates)), None)
    if year_folder:
        print(f"Anio encontrado dentro de la URL base: {year_folder.get('name')}")
        return find_child_folder(graph, year_folder, month_candidates, f"del mes {month:02d}-{month_name}")

    available = " | ".join(item.get("name", "") for item in root_children) or "sin subcarpetas"
    raise RuntimeError(
        f"No pude ubicar la carpeta de soportes para {year}/{month:02d}-{month_name}. "
        f"La URL base apunta a '{root_name}' y contiene: {available}. "
        "La URL debe apuntar a Soportes, al anio o al mes especifico."
    )


def server_relative_from_web_url(web_url):
    parsed = urlparse(web_url or "")
    return unquote(parsed.path)


def graph_file_row(item):
    web_url = item.get("webUrl", "")
    parsed = urlparse(web_url)
    server_relative = server_relative_from_web_url(web_url)
    return {
        "Name": item.get("name", ""),
        "ParentFolder": posixpath.dirname(server_relative),
        "ServerRelativeUrl": server_relative,
        "SiteOrigin": f"{parsed.scheme}://{parsed.netloc}" if parsed.scheme and parsed.netloc else "",
        "TimeCreated": item.get("createdDateTime", ""),
        "TimeLastModified": item.get("lastModifiedDateTime", ""),
        "Length": str(item.get("size", 0)),
        "webUrl": web_url,
    }


def select_cronograma(items, year, month_name):
    candidates = []
    excel_items = [item for item in items if item.get("name", "").lower().endswith(".xlsx")]
    for item in excel_items:
        name = item.get("name", "")
        normalized = name.lower()
        if "cronograma general mtto prog" in normalized and str(year) in normalized and month_name.lower() in normalized:
            candidates.append(item)
    if not candidates:
        available = " | ".join(item.get("name", "") for item in excel_items)
        raise RuntimeError(f"No encontre cronograma Excel para {month_name} {year}. Excels disponibles: {available}")
    return sorted(candidates, key=lambda item: item.get("lastModifiedDateTime", ""), reverse=True)[0]


def run_step(script, cwd, env):
    result = subprocess.run([sys.executable, str(script)], cwd=cwd, env=env, text=True, capture_output=True, timeout=540)
    print(result.stdout)
    if result.returncode:
        raise RuntimeError(f"Fallo {script.name}\nSTDOUT:\n{result.stdout[-3000:]}\nSTDERR:\n{result.stderr[-3000:]}")
def build_portal_html(metadata):
    generated = metadata.get("generated_at", "")[:16].replace("T", " ")
    period = f"{metadata.get('month', '')}-{metadata.get('month_name', '')} {metadata.get('year', '')}"
    current_year = metadata.get("year", "")
    current_month = str(metadata.get("month", "")).zfill(2)
    support_files = metadata.get("support_files", 0)
    cronograma = metadata.get("cronograma") or "Cronograma consultado"
    actions_url = "https://github.com/n3kvm/Indicador-Documental/actions/workflows/dashboard-github-pages.yml"
    refresh_endpoint = os.environ.get("DASHBOARD_REFRESH_ENDPOINT", "").strip()
    refresh_target = refresh_endpoint or actions_url
    refresh_mode = "endpoint" if refresh_endpoint else "actions"
    return f'''<!doctype html>
<html lang="es">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Indicador Documental</title>
  <style>
    :root{{--blue:#071936;--green:#7ee000;--ink:#17202a;--muted:#5f7088;--line:#d9e2ec;--bg:#f4f7fb;--card:#fff;--red:#d33f49;--gold:#c58b00}}
    *{{box-sizing:border-box}}body{{margin:0;background:var(--bg);font-family:Arial,Helvetica,sans-serif;color:var(--ink)}}
    .page{{max-width:1180px;margin:0 auto;padding:28px 20px 44px}}.hero{{background:var(--blue);color:white;border-radius:10px;overflow:hidden;position:relative;box-shadow:0 16px 38px rgba(7,25,54,.18)}}
    .hero:after{{content:"";position:absolute;right:-72px;top:-100px;width:380px;height:380px;border:34px solid rgba(126,224,0,.22);transform:rotate(18deg);border-radius:34px}}.hero-inner{{position:relative;z-index:1;padding:34px 38px;display:grid;grid-template-columns:1fr auto;gap:24px;align-items:center}}.eyebrow{{font-size:12px;font-weight:800;text-transform:uppercase;letter-spacing:.16em;color:var(--green);margin-bottom:10px}}.hero h1{{margin:0;font-size:34px;line-height:1.08}}.hero p{{margin:12px 0 0;max-width:760px;color:#dce7f4;font-size:16px;line-height:1.45}}.mark{{border:1px solid rgba(255,255,255,.28);border-radius:999px;padding:10px 16px;font-size:12px;text-transform:uppercase;letter-spacing:.12em;font-weight:800;white-space:nowrap}}.actions{{display:flex;gap:12px;flex-wrap:wrap;margin-top:24px}}.btn{{display:inline-flex;align-items:center;justify-content:center;min-height:44px;padding:0 18px;border-radius:7px;text-decoration:none;font-weight:800}}.btn.primary{{background:var(--green);color:#071936}}.btn.secondary{{background:#123e66;color:white;border:1px solid rgba(255,255,255,.18)}}
    .grid{{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin:18px 0}}.metric,.panel{{background:var(--card);border:1px solid var(--line);border-radius:9px;box-shadow:0 10px 28px rgba(30,45,65,.07)}}.metric{{padding:18px}}.metric span{{display:block;color:var(--muted);font-size:13px;margin-bottom:8px}}.metric strong{{font-size:30px}}.metric.ok strong{{color:#21885f}}.metric.warn strong{{color:var(--gold)}}.metric.bad strong{{color:var(--red)}}
    .layout{{display:grid;grid-template-columns:1.15fr .85fr;gap:18px}}.panel{{padding:22px}}.panel h2{{margin:0 0 14px;font-size:21px;color:var(--blue)}}.panel p{{color:var(--muted);line-height:1.5}}.link-card{{display:grid;grid-template-columns:auto 1fr auto;gap:14px;align-items:center;padding:14px;border:1px solid var(--line);border-radius:8px;margin-top:10px;background:#fbfdff;text-decoration:none;color:inherit}}.icon{{width:38px;height:38px;border-radius:8px;background:#eaf5ff;color:#0a4d7a;display:flex;align-items:center;justify-content:center;font-weight:900}}.link-card b{{display:block}}.link-card small{{color:var(--muted)}}.pill{{font-size:12px;font-weight:800;border-radius:999px;padding:7px 10px;background:#eef5e8;color:#467500}}.steps{{counter-reset:item;display:grid;gap:10px}}.step{{display:grid;grid-template-columns:34px 1fr;gap:12px;align-items:start}}.step:before{{counter-increment:item;content:counter(item);width:34px;height:34px;border-radius:50%;background:var(--blue);color:white;display:flex;align-items:center;justify-content:center;font-weight:800}}.footer{{margin-top:18px;color:var(--muted);font-size:12px;text-align:center}}
    .refresh-status{{display:none;margin-top:14px;padding:12px 14px;border-radius:8px;background:#ecf7ff;color:#0a4d7a;font-weight:700}}.refresh-status.show{{display:block}}.refresh-status.error{{background:#fff1f1;color:#b4232b}}.refresh-status.ok{{background:#eefbe8;color:#2c7200}}
    .refresh-form{{display:flex;gap:10px;flex-wrap:wrap;align-items:end;margin-top:18px}}.field{{display:grid;gap:6px}}.field label{{font-size:12px;font-weight:800;text-transform:uppercase;letter-spacing:.08em;color:#dce7f4}}.field input{{height:42px;border-radius:7px;border:1px solid rgba(255,255,255,.24);background:white;color:var(--ink);font-weight:800;padding:0 12px;min-width:104px}}
    @media(max-width:900px){{.hero-inner,.layout{{grid-template-columns:1fr}}.grid{{grid-template-columns:repeat(2,1fr)}}.hero h1{{font-size:28px}}.mark{{width:max-content}}}}@media(max-width:560px){{.grid{{grid-template-columns:1fr}}.hero-inner{{padding:26px 22px}}}}
  </style>
</head>
<body>
  <main class="page">
    <section class="hero"><div class="hero-inner"><div><div class="eyebrow">Indicador documental</div><h1>Dashboard mantenimiento preventivo</h1><p>Consulta centralizada de soportes cargados, checklist esperados, pilares de mantenimiento, horas calculadas y alertas documentales.</p><div class="refresh-form"><div class="field"><label for="refreshYear">Año</label><input id="refreshYear" inputmode="numeric" maxlength="4" value="{current_year}" /></div><div class="field"><label for="refreshMonth">Mes</label><input id="refreshMonth" inputmode="numeric" maxlength="2" value="{current_month}" /></div><button class="btn secondary" id="refreshBtn" type="button" data-mode="{refresh_mode}" data-target="{refresh_target}">Actualizar dashboard</button></div><div class="actions"><a class="btn primary" href="dashboard.html">Abrir dashboard</a><a class="btn secondary" href="validacion_soportes_sharepoint.xlsx">Descargar validación</a><a class="btn secondary" href="auditoria_soportes.xlsx">Descargar auditoría</a></div><div class="refresh-status" id="refreshStatus"></div></div><div class="mark">Realizado por Bryan Martinez</div></div></section>
    <section class="grid" aria-label="Resumen del reporte"><div class="metric ok"><span>Estado</span><strong>Activo</strong></div><div class="metric"><span>Periodo</span><strong>{period}</strong></div><div class="metric warn"><span>PDFs leídos</span><strong>{support_files}</strong></div><div class="metric bad"><span>Actualizado</span><strong>{generated}</strong></div></section>
    <section class="layout"><article class="panel"><h2>Accesos principales</h2><a class="link-card" href="dashboard.html"><div class="icon">ID</div><div><b>Dashboard interactivo</b><small>Vista HTML con filtros, alertas, archivos, pilares y cálculo por días.</small></div><span class="pill">HTML</span></a><button class="link-card" id="refreshCard" type="button" data-mode="{refresh_mode}" data-target="{refresh_target}"><div class="icon">R</div><div><b>Actualizar dashboard</b><small>Usa el año y mes digitados arriba para ejecutar nuevamente la lectura de SharePoint.</small></div><span class="pill">Action</span></button><a class="link-card" href="validacion_soportes_sharepoint.xlsx"><div class="icon">XL</div><div><b>Validación de soportes</b><small>Comparativo cronograma, esperado y SharePoint por sede.</small></div><span class="pill">Excel</span></a><a class="link-card" href="auditoria_soportes.xlsx"><div class="icon">AU</div><div><b>Auditoría del cronograma</b><small>{cronograma}</small></div><span class="pill">Excel</span></a></article><aside class="panel"><h2>Cómo usar</h2><div class="steps"><div class="step"><div><b>Elegir periodo</b><p>Digita el año y el mes que quieres consultar antes de actualizar.</p></div></div><div class="step"><div><b>Actualizar datos</b><p>Pulsa Actualizar dashboard y espera unos minutos mientras GitHub Pages publica el nuevo resultado.</p></div></div><div class="step"><div><b>Abrir dashboard</b><p>Usa el botón principal para revisar el consolidado del mes publicado.</p></div></div><div class="step"><div><b>Descargar soporte</b><p>Usa los reportes Excel cuando necesites enviar evidencias o hacer revisión detallada.</p></div></div></div></aside></section>
    <div class="footer">Brillaseo SAS - Gestión Locativa Integral - Indicador documental</div>
  </main>
  <script>
    const statusEl = document.getElementById("refreshStatus");
    function showStatus(text, kind) {{
      statusEl.textContent = text;
      statusEl.className = "refresh-status show" + (kind ? " " + kind : "");
    }}

    async function waitForCompletion(endpoint, button) {{
      const statusEndpoint = endpoint.replace(/\/$/, "");

      for (let attempt = 1; attempt <= 90; attempt++) {{
        await new Promise(r => setTimeout(r, 5000));

        try {{
          const response = await fetch(statusEndpoint, {{
            method: "GET",
            cache: "no-store"
          }});

          const data = await response.json().catch(() => ({{ ok: false }}));

          if (!response.ok || !data.ok) {{
            showStatus("Actualización iniciada. Consultando estado de GitHub Actions...", "");
            continue;
          }}

          if (data.status === "queued") {{
            showStatus("🟡 Actualización en cola en GitHub Actions...", "");
            continue;
          }}

          if (data.status === "in_progress") {{
            showStatus("🔵 Generando dashboard. Leyendo SharePoint y publicando resultados...", "");
            continue;
          }}

          if (data.status === "completed") {{
            if (data.conclusion === "success") {{
              showStatus("✅ Dashboard actualizado. Recargando página...", "ok");
              setTimeout(() => location.reload(), 2000);
            }} else {{
              showStatus("❌ La actualización terminó con estado: " + (data.conclusion || "sin conclusión") + ". Revisa GitHub Actions.", "error");
              button.disabled = false;
            }}
            return;
          }}
        }} catch (error) {{
          showStatus("Actualización iniciada. No se pudo consultar el estado todavía...", "");
        }}
      }}

      showStatus("La actualización sigue tardando. Revisa GitHub Actions o recarga en unos minutos.", "error");
      button.disabled = false;
    }}
    async function refreshDashboard(button) {{
      const target = button.dataset.target;
      const mode = button.dataset.mode;
      const anio = document.getElementById("refreshYear").value.trim();
      const mes = document.getElementById("refreshMonth").value.trim().padStart(2, "0");
      if (!/^\\d{{4}}$/.test(anio) || !/^\\d{{2}}$/.test(mes) || Number(mes) < 1 || Number(mes) > 12) {{
        showStatus("Digite un año de 4 dígitos y un mes entre 01 y 12.", "error");
        return;
      }}
      if (!target) {{
        showStatus("No hay endpoint de actualización configurado.", "error");
        return;
      }}
      if (mode !== "endpoint") {{
        window.open(target, "_blank", "noopener");
        showStatus("Se abrió GitHub Actions. Configura DASHBOARD_REFRESH_ENDPOINT para actualizar con un solo clic.", "error");
        return;
      }}
      button.disabled = true;
      showStatus("Actualización enviada para " + anio + "-" + mes + ". GitHub está generando el nuevo dashboard...", "");
      try {{
        const response = await fetch(target, {{
          method: "POST",
          headers: {{ "Content-Type": "application/json" }},
          body: JSON.stringify({{ anio, mes }})
        }});
        const data = await response.json().catch(() => ({{ ok: false, error: "Respuesta no válida del endpoint." }}));
        if (!response.ok || !data.ok) throw new Error(data.error || data.message || data.detail || "HTTP " + response.status);

        showStatus("Solicitud enviada. Esperando a que GitHub termine...", "");
        waitForCompletion(target, button);
        return;
      }} catch (error) {{
        showStatus("No se pudo iniciar la actualización: " + error.message, "error");
        button.disabled = false;
      }}
    }}
    document.getElementById("refreshBtn").addEventListener("click", (event) => refreshDashboard(event.currentTarget));
    document.getElementById("refreshCard").addEventListener("click", (event) => refreshDashboard(event.currentTarget));
  </script>
</body>
</html>'''


def main():
    repo_root = Path(__file__).resolve().parents[1]
    work_scripts = repo_root / "work"
    public_dir = repo_root / "public"
    public_dir.mkdir(parents=True, exist_ok=True)

    year, month = period_from_env()
    month_name = MONTHS[month]
    folder_month = f"{month:02d}-{month_name}"
    route_name = os.environ.get("DASHBOARD_ROUTE_NAME", "GITHUB")
    print(f"DASHBOARD_REFRESH_ENDPOINT configurado: {'SI' if os.environ.get('DASHBOARD_REFRESH_ENDPOINT', '').strip() else 'NO'}")

    supports_url = required_env("SUPPORTS_FOLDER_URL")
    cronograma_url = required_env("CRONOGRAMA_URL")

    run_root = Path(tempfile.mkdtemp(prefix="dashboard_github_"))
    run_work = run_root / "work"
    run_outputs = run_root / "outputs"
    run_work.mkdir(parents=True, exist_ok=True)
    run_outputs.mkdir(parents=True, exist_ok=True)

    graph = GraphClient()

    print(f"Periodo: {year}-{month:02d} {month_name}")
    print("Listando soportes...")
    supports_root_item = graph.drive_item_from_share_url(supports_url)
    supports_item = resolve_support_period_folder(graph, supports_root_item, year, month, month_name)
    print(f"Carpeta de soportes seleccionada: {supports_item.get('name')}")
    support_items = [item for item in graph.list_recursive(supports_item) if item.get("name", "").lower().endswith(".pdf")]
    print(f"PDFs encontrados en {year}/{folder_month}: {len(support_items)}")

    print("Buscando cronograma...")
    cronograma_item = graph.drive_item_from_share_url(cronograma_url)
    cronograma_candidates = graph.list_recursive(cronograma_item) if "folder" in cronograma_item else [cronograma_item]
    selected_cron = select_cronograma(cronograma_candidates, year, month_name)

    cronograma_path = run_work / "cronograma_copia.xlsx"
    cron_bytes = graph.download(selected_cron, cronograma_path)
    (run_work / "cronograma_sharepoint.xlsx").write_bytes(cronograma_path.read_bytes())
    (run_work / "cronograma_sharepoint_meta.json").write_text(
        json.dumps({"selected": selected_cron, "downloadedBytes": cron_bytes}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    print(f"Descargando {len(support_items)} PDF(s)...")
    manifest = []
    file_rows = []
    pdf_dir = run_work / "sharepoint_pdfs"
    for idx, item in enumerate(support_items, start=1):
        row = graph_file_row(item)
        local_path = pdf_dir / safe_name(row["Name"])
        downloaded = graph.download(item, local_path)
        row["localPath"] = str(local_path)
        row["downloadedBytes"] = downloaded
        row["error"] = "" if downloaded > 0 else "Archivo descargado con 0 bytes."
        manifest.append(row)
        file_rows.append({k: v for k, v in row.items() if k != "localPath"})
        print(f"{idx}/{len(support_items)} {row['Name']} {downloaded}")

    (run_work / "sharepoint_files.json").write_text(
        json.dumps({"files": file_rows, "siteOrigin": file_rows[0].get("SiteOrigin", "") if file_rows else ""}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    (run_work / "sharepoint_pdf_manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")

    compare_output = run_outputs / "validacion_soportes_sharepoint.xlsx"
    audit_output = run_outputs / "auditoria_soportes.xlsx"
    pillars_output = run_work / "pillars_hours_validation.json"
    app_upload_output = run_work / "app_upload_validation.json"
    zoho_csv_output = run_work / "zoho_autogestiones.csv"
    dashboard_output = run_outputs / "index.html"

    env = os.environ.copy()
    env.update(
        {
            "SHAREPOINT_ANIO": str(year),
            "SHAREPOINT_MES": f"{month:02d}",
            "SHAREPOINT_MES_NOMBRE": month_name,
            "SHAREPOINT_CARPETA_MES": folder_month,
            "CRONOGRAMA_PATH": str(cronograma_path),
            "EXPECTED_SCHEDULE_PATH": str(run_work / "expected_schedule.json"),
            "SCHEDULE_AUDIT_OUTPUT_PATH": str(audit_output),
            "SHAREPOINT_FILES_PATH": str(run_work / "sharepoint_files.json"),
            "PDF_MANIFEST_PATH": str(run_work / "sharepoint_pdf_manifest.json"),
            "PDF_TEXT_DIR": str(run_work / "pdf_text"),
            "PDF_TEXT_SUMMARY_PATH": str(run_work / "pdf_text_summary.json"),
            "COMPARE_OUTPUT_PATH": str(compare_output),
            "PILLARS_VALIDATION_PATH": str(pillars_output),
            "APP_UPLOAD_VALIDATION_PATH": str(app_upload_output),
            "APP_AUTOGESTIONES_CSV_PATH": str(zoho_csv_output),
            "ZOHO_OUTPUT_CSV_PATH": str(zoho_csv_output),
            "ZOHO_OUTPUT_META_PATH": str(run_work / "zoho_autogestiones_meta.json"),
            "DASHBOARD_INCLUDE_APP_VIEW": "1",
            "DASHBOARD_OUTPUT_PATH": str(dashboard_output),
            "DASHBOARD_REFRESH_ENDPOINT": "",
        }
    )

    for script_name in [
        "extract_pdf_text.py",
        "extract_schedule.py",
        "compare_sharepoint_supports.py",
        "fetch_zoho_app_records.py",
        "compare_app_uploads.py",
        "validate_pillars_hours.py",
        "build_dashboard_pillars.py",
    ]:
        run_step(work_scripts / script_name, run_root, env)

    shutil.copy2(dashboard_output, public_dir / "dashboard.html")
    if compare_output.exists():
        shutil.copy2(compare_output, public_dir / "validacion_soportes_sharepoint.xlsx")
    if audit_output.exists():
        shutil.copy2(audit_output, public_dir / "auditoria_soportes.xlsx")

    metadata = {
        "generated_at": datetime.now(ZoneInfo("America/Bogota")).isoformat(),
        "year": year,
        "month": f"{month:02d}",
        "month_name": month_name,
        "route_name": route_name,
        "support_files": len(support_items),
        "cronograma": selected_cron.get("name"),
    }
    (public_dir / "index.html").write_text(build_portal_html(metadata), encoding="utf-8")
    metadata_path = public_dir / "metadata.json"
    metadata_path.write_text(json.dumps(metadata, ensure_ascii=False, indent=2), encoding="utf-8")

    publish_site_url = os.environ.get("PUBLISH_SITE_URL", "https://brillaseo2.sharepoint.com/sites/SoportesEspejo").strip()
    publish_folder = os.environ.get("PUBLISH_FOLDER_PATH", "Dashboard Mantenimiento").strip()
    if publish_site_url and publish_folder and os.environ.get("PUBLISH_TO_SHAREPOINT", "1") != "0":
        print(f"Publicando dashboard en SharePoint: {publish_site_url} / {publish_folder}")
        published = []
        try:
            folder = graph.ensure_folder(publish_site_url, publish_folder)
            for file_path in [public_dir / "index.html", public_dir / "dashboard.html", public_dir / "metadata.json", public_dir / "validacion_soportes_sharepoint.xlsx", public_dir / "auditoria_soportes.xlsx"]:
                if file_path.exists():
                    uploaded = graph.upload_to_folder(folder, file_path)
                    published.append({"name": file_path.name, "webUrl": uploaded.get("webUrl", "")})
                    print(f"Publicado: {file_path.name} -> {uploaded.get('webUrl', '')}")
            metadata["published_to_sharepoint"] = published
        except Exception as exc:
            metadata["publish_warning"] = str(exc)
            print("ADVERTENCIA: el dashboard se genero, pero no se pudo publicar en SharePoint.")
            print("Para publicar, la app necesita permiso de escritura sobre el sitio/carpeta destino.")
            print(f"Detalle publicacion: {exc}")
        metadata_path.write_text(json.dumps(metadata, ensure_ascii=False, indent=2), encoding="utf-8")

    print(json.dumps(metadata, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()










