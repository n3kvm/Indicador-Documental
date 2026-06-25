import json
import os
import re
from pathlib import Path

from pypdf import PdfReader


MANIFEST = Path(os.environ.get("PDF_MANIFEST_PATH", "work/sharepoint_pdf_manifest.json"))
OUT_DIR = Path(os.environ.get("PDF_TEXT_DIR", "work/pdf_text"))
SUMMARY_OUT = Path(os.environ.get("PDF_TEXT_SUMMARY_PATH", "work/pdf_text_summary.json"))


def clean_text(text):
    text = text or ""
    text = text.replace("\x00", " ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


manifest = json.loads(MANIFEST.read_text(encoding="utf-8")) if MANIFEST.exists() else []
if not manifest and not MANIFEST.exists():
    pdf_dir = Path("work/sharepoint_pdfs")
    manifest = [
        {
            "Name": pdf.name,
            "localPath": str(pdf),
            "error": "",
        }
        for pdf in sorted(pdf_dir.glob("*.pdf"))
    ]
OUT_DIR.mkdir(parents=True, exist_ok=True)
summary = []
for item in manifest:
    local = item.get("localPath")
    if item.get("error") or not local or not Path(local).exists():
        summary.append({**item, "pages": 0, "text_chars": 0, "text_path": "", "text_error": item.get("error", "missing local file")})
        continue
    try:
        reader = PdfReader(local)
        parts = []
        for i, page in enumerate(reader.pages):
            text = page.extract_text() or ""
            parts.append(f"\n\n--- PAGE {i + 1} ---\n{text}")
        full_text = clean_text("\n".join(parts))
        text_path = OUT_DIR / (Path(local).stem + ".txt")
        text_path.write_text(full_text, encoding="utf-8")
        summary.append({**item, "pages": len(reader.pages), "text_chars": len(full_text), "text_path": str(text_path), "text_error": ""})
    except Exception as exc:
        summary.append({**item, "pages": 0, "text_chars": 0, "text_path": "", "text_error": str(exc)})

SUMMARY_OUT.parent.mkdir(parents=True, exist_ok=True)
SUMMARY_OUT.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
print(json.dumps({
    "pdfs": len(summary),
    "readable": sum(1 for s in summary if s.get("text_chars", 0) > 500),
    "low_text": sum(1 for s in summary if 0 < s.get("text_chars", 0) <= 500),
    "no_text_or_error": sum(1 for s in summary if s.get("text_chars", 0) == 0),
    "total_chars": sum(s.get("text_chars", 0) for s in summary),
}, ensure_ascii=False, indent=2))
