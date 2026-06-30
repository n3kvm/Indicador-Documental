import base64
import json
import os
import posixpath
import shutil
import subprocess
import sys
import tempfile
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

    def drive_item_from_share_url(self, share_url):
        parsed = urlparse(share_url or "")
        if parsed.netloc.endswith(".sharepoint.com"):
            return self.drive_item_from_sharepoint_url(share_url)
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
        if not rel_path:
            return self.get_json(f"https://graph.microsoft.com/v1.0/drives/{drive['id']}/root")
        encoded = quote(rel_path, safe="/")
        return self.get_json(f"https://graph.microsoft.com/v1.0/drives/{drive['id']}/root:/{encoded}:")

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


def main():
    repo_root = Path(__file__).resolve().parents[1]
    work_scripts = repo_root / "work"
    public_dir = repo_root / "public"
    public_dir.mkdir(parents=True, exist_ok=True)

    year, month = period_from_env()
    month_name = MONTHS[month]
    folder_month = f"{month:02d}-{month_name}"
    route_name = os.environ.get("DASHBOARD_ROUTE_NAME", "GITHUB")

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
    supports_item = graph.drive_item_from_share_url(supports_url)
    support_items = [item for item in graph.list_recursive(supports_item) if item.get("name", "").lower().endswith(".pdf")]

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
            "DASHBOARD_OUTPUT_PATH": str(dashboard_output),
            "DASHBOARD_REFRESH_ENDPOINT": "",
        }
    )

    for script_name in [
        "extract_pdf_text.py",
        "extract_schedule.py",
        "compare_sharepoint_supports.py",
        "validate_pillars_hours.py",
        "build_dashboard_pillars.py",
    ]:
        run_step(work_scripts / script_name, run_root, env)

    shutil.copy2(dashboard_output, public_dir / "index.html")
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
    (public_dir / "metadata.json").write_text(json.dumps(metadata, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(metadata, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()



