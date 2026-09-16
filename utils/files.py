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


def reserve_output_path(output_dir: Path, filename: str) -> Path:
    """Atomically claim ``filename`` inside ``output_dir``.

    The placeholder file is created with ``O_EXCL`` so two concurrent jobs that
    request the same name can never be handed the same path. Callers write the
    real document over the placeholder (or drop it via :func:`discard_output_path`).
    """
    output_dir.mkdir(parents=True, exist_ok=True)
    stem = Path(filename).stem or "document"
    suffix = Path(filename).suffix or ".docx"

    for attempt in range(10):
        candidate = filename if attempt == 0 else f"{stem}_{uuid.uuid4().hex[:8]}{suffix}"
        path = output_dir / candidate
        try:
            descriptor = os.open(path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        except FileExistsError:
            continue
        os.close(descriptor)
        return path

    raise RuntimeError(f"无法为 {filename} 分配唯一输出路径")


def discard_output_path(path: Path) -> None:
    """Release a path previously claimed by :func:`reserve_output_path`."""
    try:
        path.unlink()
    except OSError:
        pass


def ensure_unique_docx_path(output_dir: Path, requested_filename: str) -> tuple[str, Path]:
    docx_filename = normalize_output_docx_filename(requested_filename)
    docx_path = reserve_output_path(output_dir, docx_filename)
    return docx_path.name, docx_path
