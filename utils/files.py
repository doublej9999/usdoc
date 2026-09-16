import os
import re
import uuid
from datetime import datetime
from pathlib import Path


def sanitize_filename(filename: str) -> str:
    raw = str(filename or "").replace("\0", "")
    base = os.path.basename(raw.replace("\\", "/")).strip()
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", base).strip()
    cleaned = cleaned.rstrip(". ")
    if cleaned in (".", ".."):
        return ""
    return cleaned


def sanitize_output_filename(filename: str) -> str:
    return sanitize_filename(filename)


def normalize_output_docx_filename(filename: str) -> str:
    raw = (filename or "").strip()
    if not raw:
        return build_default_docx_filename()

    cleaned = sanitize_output_filename(raw)
    if not cleaned:
        return build_default_docx_filename()
    if not cleaned.lower().endswith(".docx"):
        cleaned = f"{cleaned}.docx"
    return cleaned


def build_default_docx_filename() -> str:
    now = datetime.now().strftime("%Y%m%d_%H%M%S")
    uid = uuid.uuid4().hex[:8]
    return f"generated_{now}_{uid}.docx"


def ensure_unique_docx_path(output_dir: Path, requested_filename: str) -> tuple[str, Path]:
    docx_filename = normalize_output_docx_filename(requested_filename)
    docx_path = output_dir / docx_filename
    if docx_path.exists():
        uid = uuid.uuid4().hex[:8]
        stem = Path(docx_filename).stem
        docx_filename = f"{stem}_{uid}.docx"
        docx_path = output_dir / docx_filename
    return docx_filename, docx_path
