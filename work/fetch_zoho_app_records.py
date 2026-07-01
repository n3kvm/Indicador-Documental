import csv
import json
import os
import sys
import time
import urllib.parse
import urllib.request
from pathlib import Path


OUTPUT_CSV = Path(os.environ.get("ZOHO_OUTPUT_CSV_PATH", "work/zoho_autogestiones.csv"))
META_PATH = Path(os.environ.get("ZOHO_OUTPUT_META_PATH", "work/zoho_autogestiones_meta.json"))

BASE_URL = os.environ.get("ZOHO_CREATOR_BASE_URL", "https://www.zohoapis.com").rstrip("/")
ACCOUNTS_URL = os.environ.get("ZOHO_ACCOUNTS_URL", "https://accounts.zoho.com").rstrip("/")
OWNER = os.environ.get("ZOHO_ACCOUNT_OWNER", "").strip()
APP_LINK = os.environ.get("ZOHO_APP_LINK_NAME", "").strip()
REPORT_LINK = os.environ.get("ZOHO_REPORT_LINK_NAME", "").strip()

ACCESS_TOKEN = os.environ.get("ZOHO_ACCESS_TOKEN", "").strip()
REFRESH_TOKEN = os.environ.get("ZOHO_REFRESH_TOKEN", "").strip()
CLIENT_ID = os.environ.get("ZOHO_CLIENT_ID", "").strip()
CLIENT_SECRET = os.environ.get("ZOHO_CLIENT_SECRET", "").strip()

FIELD_AUTOGESTION = os.environ.get("ZOHO_FIELD_AUTOGESTION", "Autogestion")
FIELD_SEDE = os.environ.get("ZOHO_FIELD_SEDE", "Sede")
FIELD_UES = os.environ.get("ZOHO_FIELD_UES", "UES")
FIELD_FECHA_REALIZACION = os.environ.get("ZOHO_FIELD_FECHA_REALIZACION", "Fecha_Realizacion")
FIELD_FECHA_REGISTRO = os.environ.get("ZOHO_FIELD_FECHA_REGISTRO", "Added_Time")


def configured():
    has_token = bool(ACCESS_TOKEN) or bool(REFRESH_TOKEN and CLIENT_ID and CLIENT_SECRET)
    return bool(OWNER and APP_LINK and REPORT_LINK and has_token)


def write_meta(**data):
    META_PATH.parent.mkdir(parents=True, exist_ok=True)
    META_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def request_json(url, headers=None, data=None, method=None):
    req = urllib.request.Request(
        url,
        data=data,
        headers=headers or {},
        method=method or ("POST" if data is not None else "GET"),
    )
    with urllib.request.urlopen(req, timeout=90) as res:
        body = res.read().decode("utf-8", errors="replace")
        return json.loads(body), res.headers


def get_access_token():
    if ACCESS_TOKEN:
        return ACCESS_TOKEN

    params = urllib.parse.urlencode(
        {
            "refresh_token": REFRESH_TOKEN,
            "client_id": CLIENT_ID,
            "client_secret": CLIENT_SECRET,
            "grant_type": "refresh_token",
        }
    ).encode("utf-8")
    token_url = f"{ACCOUNTS_URL}/oauth/v2/token"
    token, _ = request_json(token_url, data=params, headers={"Content-Type": "application/x-www-form-urlencoded"})
    access_token = token.get("access_token")
    if not access_token:
        raise RuntimeError(f"Zoho no entrego access_token: {token}")
    return access_token


def flatten_value(value):
    if value is None:
        return ""
    if isinstance(value, dict):
        return str(value.get("zc_display_value") or value.get("display_value") or value.get("name") or value.get("Name") or value.get("ID") or "")
    if isinstance(value, list):
        parts = [flatten_value(item) for item in value]
        return ", ".join(part for part in parts if part)
    return str(value)


def record_value(record, field):
    if field in record:
        return flatten_value(record.get(field))
    field_lower = field.lower()
    for key, value in record.items():
        if key.lower() == field_lower:
            return flatten_value(value)
    return ""


def fetch_records(token):
    encoded_owner = urllib.parse.quote(OWNER, safe="")
    encoded_app = urllib.parse.quote(APP_LINK, safe="")
    encoded_report = urllib.parse.quote(REPORT_LINK, safe="")
    url = f"{BASE_URL}/creator/v2.1/data/{encoded_owner}/{encoded_app}/report/{encoded_report}"
    params = {
        "max_records": os.environ.get("ZOHO_MAX_RECORDS", "1000"),
        "field_config": os.environ.get("ZOHO_FIELD_CONFIG", "all"),
    }
    criteria = os.environ.get("ZOHO_CRITERIA", "").strip()
    if criteria:
        params["criteria"] = criteria

    headers = {
        "Authorization": f"Zoho-oauthtoken {token}",
        "Accept": "application/json",
    }
    environment = os.environ.get("ZOHO_ENVIRONMENT", "").strip()
    if environment:
        headers["environment"] = environment

    records = []
    cursor = ""
    for _ in range(100):
        page_url = f"{url}?{urllib.parse.urlencode(params)}"
        page_headers = dict(headers)
        if cursor:
            page_headers["record_cursor"] = cursor
        payload, response_headers = request_json(page_url, headers=page_headers)
        code = str(payload.get("code", ""))
        if code and code != "3000":
            raise RuntimeError(f"Zoho respondio codigo {code}: {payload}")
        records.extend(payload.get("data") or [])
        cursor = response_headers.get("record_cursor") or response_headers.get("Record-Cursor") or ""
        if not cursor:
            break
        time.sleep(0.2)
    return records


def write_csv(records):
    OUTPUT_CSV.parent.mkdir(parents=True, exist_ok=True)
    with OUTPUT_CSV.open("w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(
            fh,
            fieldnames=[
                "Autogestion",
                "Nombre Sede",
                "UES",
                "Fecha Realizacion",
                "Fecha Registro",
                "Zoho ID",
            ],
        )
        writer.writeheader()
        for record in records:
            writer.writerow(
                {
                    "Autogestion": record_value(record, FIELD_AUTOGESTION),
                    "Nombre Sede": record_value(record, FIELD_SEDE),
                    "UES": record_value(record, FIELD_UES),
                    "Fecha Realizacion": record_value(record, FIELD_FECHA_REALIZACION),
                    "Fecha Registro": record_value(record, FIELD_FECHA_REGISTRO),
                    "Zoho ID": record_value(record, "ID"),
                }
            )


def main():
    if not configured():
        write_meta(
            enabled=False,
            message="Zoho no configurado. Se usara CSV local si existe.",
            required=[
                "ZOHO_ACCOUNT_OWNER",
                "ZOHO_APP_LINK_NAME",
                "ZOHO_REPORT_LINK_NAME",
                "ZOHO_REFRESH_TOKEN + ZOHO_CLIENT_ID + ZOHO_CLIENT_SECRET o ZOHO_ACCESS_TOKEN",
            ],
        )
        print("Zoho no configurado; se omite descarga del aplicativo.")
        return 0

    token = get_access_token()
    records = fetch_records(token)
    write_csv(records)
    write_meta(
        enabled=True,
        source="zoho_creator_api",
        output=str(OUTPUT_CSV),
        records=len(records),
        owner=OWNER,
        app_link_name=APP_LINK,
        report_link_name=REPORT_LINK,
        fields={
            "autogestion": FIELD_AUTOGESTION,
            "sede": FIELD_SEDE,
            "ues": FIELD_UES,
            "fecha_realizacion": FIELD_FECHA_REALIZACION,
            "fecha_registro": FIELD_FECHA_REGISTRO,
        },
    )
    print(f"Zoho Creator: {len(records)} registro(s) exportados a {OUTPUT_CSV}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        write_meta(enabled=False, source="zoho_creator_api", error=str(exc))
        print(f"ERROR Zoho Creator: {exc}", file=sys.stderr)
        raise
