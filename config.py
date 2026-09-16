import os
import sys
from pathlib import Path


def resolve_base_dir() -> Path:
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parent


BASE_DIR = resolve_base_dir()
DATA_DIR = Path.cwd() if getattr(sys, "frozen", False) else BASE_DIR
UPLOAD_DIR = DATA_DIR / "uploads"
OUTPUT_DIR = DATA_DIR / "outputs"
TEMPLATES_DIR = BASE_DIR / "templates"
STATIC_DIR = BASE_DIR / "static"
GENERATION_RECORDS_FILE = DATA_DIR / "generation_records.json"
APP_TITLE = "AI 文档生成系统"

DEFAULT_API_URL = os.getenv("OPENAI_BASE_URL", "https://ai.927900.xyz/v1/chat/completions")
DEFAULT_MODEL = os.getenv("OPENAI_MODEL", "gpt-5.4")
DEFAULT_API_KEY = os.getenv("OPENAI_API_KEY", "")


def ensure_runtime_directories() -> None:
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
