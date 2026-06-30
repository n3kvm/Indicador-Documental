import json
import os
import re
from pathlib import Path

try:
    from pypdf import PdfReader
except Exception as exc:
    PdfReader = None
    IMPORT_ERROR = exc
else:
    IMPORT_ERROR = None

MANIFEST = Path(os.environ.get("SHAREPOINT_PDF_MANIFEST_PATH", "work/sharepoint_pdf_manifest.json"))
OUT_DIR = Path(os.environ.get("PDF_TEXT_DIR", "work/pdf_text"))
SUMMARY = Path(os.environ.get("PDF_TEXT_SUMMARY_PATH", "work/pdf_text_summary.json"))


def safe_stem(name):
    stem = re.sub(r"\.[^.]+$", "", str(name or "archivo"))
    stem = re.sub(r"[<>:\"/\\|?*\x00-\x1F]+", "_", stem).strip(" ._")
    return stem or "archivo"


def extract_text(pdf_path):
    if IMPORT_ERROR:
        raise RuntimeError(f"No se pudo importar pypdf: {IMPORT_ERROR}")
    reader = PdfReader(str(pdf_path))
    pages = []
    for page in reader.pages:
        try:
            pages.append(page.extract_text() or "")
        except Exception as exc:
            pages.append(f"\n[Error leyendo pagina: {exc}]\n")
    return "\n\n".join(pages), len(reader.pages)


def main():
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    if not MANIFEST.exists():
        SUMMARY.write_text("[]", encoding="utf-8")
        print(f"No existe {MANIFEST}. Resumen PDF vacio.")
        return
    items = json.loads(MANIFEST.read_text(encoding="utf-8"))
    summary = []
    for index, item in enumerate(items, start=1):
        name = item.get("Name") or item.get("name") or Path(item.get("localPath", "archivo.pdf")).name
        local_path = Path(item.get("localPath") or "")
        text_path = OUT_DIR / f"{index:03d}_{safe_stem(name)}.txt"
        row = dict(item)
        row["Name"] = name
        row["text_path"] = str(text_path)
        row["pages"] = 0
        row["chars"] = 0
        if not local_path.exists() or not local_path.is_file():
            row["text_error"] = row.get("text_error") or "PDF no descargado o ruta local inexistente"
            text_path.write_text("", encoding="utf-8")
        elif local_path.stat().st_size == 0:
            row["downloadedBytes"] = 0
            row["text_error"] = row.get("text_error") or "PDF vacio / 0 bytes"
            text_path.write_text("", encoding="utf-8")
        else:
            try:
                text, pages = extract_text(local_path)
                text_path.write_text(text, encoding="utf-8", errors="ignore")
                row["pages"] = pages
                row["chars"] = len(text)
                row["text_error"] = ""
            except Exception as exc:
                row["text_error"] = str(exc)
                text_path.write_text("", encoding="utf-8")
        summary.append(row)
        print(f"{index}/{len(items)} {name} paginas={row['pages']} caracteres={row['chars']} error={row.get('text_error') or '-'}")
    SUMMARY.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Resumen PDF generado: {SUMMARY}")


if __name__ == "__main__":
    main()
