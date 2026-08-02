import json
import logging
import os
import re
import shutil
import subprocess
import sys
import uuid
import zipfile
from datetime import datetime
from pathlib import Path
from threading import Lock
from typing import Dict, List
import time

import requests
from docx import Document
from docxtpl import DocxTemplate
from fastapi import BackgroundTasks, FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from fastapi import Request
from jinja2 import TemplateSyntaxError
from openpyxl import load_workbook
from pydantic import BaseModel

from starlette.concurrency import run_in_threadpool

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------

logger = logging.getLogger("usdoc")
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------

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

UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

MAX_UPLOAD_SIZE = 20 * 1024 * 1024  # 20 MB
DOCX_MAGIC_BYTES = b"PK"  # .docx is a ZIP archive
AI_REQUEST_TIMEOUT = 90  # seconds
AI_MAX_RETRIES = 3
AI_RETRY_BACKOFF = 2  # seconds, doubled each retry
LIBREOFFICE_TIMEOUT = 60  # seconds

app = FastAPI(title="AI 文档生成系统")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")
templates = Jinja2Templates(directory=str(TEMPLATES_DIR))


# ---------------------------------------------------------------------------
# Pydantic models
# ---------------------------------------------------------------------------

class GenerateRequest(BaseModel):
    prompt: str
    api_url: str
    api_key: str
    model: str
    template_name: str
    output_filename: str = ""

# ---------------------------------------------------------------------------
# Generation records (in-memory + JSON file persistence)
# ---------------------------------------------------------------------------

generation_records: Dict[str, Dict[str, str]] = {}
generation_records_lock = Lock()
libreoffice_lock = Lock()


def now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")


def dump_generation_records() -> None:
    with generation_records_lock:
        records = list(generation_records.values())
        records.sort(key=lambda x: x["created_at"], reverse=True)
        GENERATION_RECORDS_FILE.write_text(
            json.dumps(records, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )


def load_generation_records() -> None:
    if not GENERATION_RECORDS_FILE.exists():
        return

    try:
        records = json.loads(GENERATION_RECORDS_FILE.read_text(encoding="utf-8"))
        if not isinstance(records, list):
            return
    except Exception:
        return

    with generation_records_lock:
        generation_records.clear()
        for record in records:
            if isinstance(record, dict) and record.get("id"):
                generation_records[record["id"]] = record


def add_generation_record(req: GenerateRequest) -> Dict[str, str]:
    record_id = uuid.uuid4().hex
    current_time = now_iso()
    record = {
        "id": record_id,
        "prompt": req.prompt,
        # #2 fix: do not store api_url in generation records to avoid leaking endpoint info
        "model": req.model,
        "template_name": req.template_name,
        "generated_filename": req.output_filename,
        "status": "pending",
        "message": "任务已创建，等待后台生成",
        "filename": "",
        "download_url": "",
        "error": "",
        "created_at": current_time,
        "updated_at": current_time,
    }

    with generation_records_lock:
        generation_records[record_id] = record

    dump_generation_records()
    return record


def update_generation_record(record_id: str, **fields) -> None:
    updated = False
    with generation_records_lock:
        record = generation_records.get(record_id)
        if record:
            record.update(fields)
            record["updated_at"] = now_iso()
            updated = True

    if updated:
        dump_generation_records()


def get_generation_record(record_id: str) -> Dict[str, str]:
    with generation_records_lock:
        record = generation_records.get(record_id)
        if not record:
            raise HTTPException(status_code=404, detail="生成记录不存在")
        return dict(record)


def list_generation_records(limit: int) -> List[Dict[str, str]]:
    with generation_records_lock:
        records = [dict(item) for item in generation_records.values()]

    records.sort(key=lambda x: x["created_at"], reverse=True)
    return records[:limit]


load_generation_records()


# ---------------------------------------------------------------------------
# Filename sanitisation
# ---------------------------------------------------------------------------

def sanitize_filename(filename: str) -> str:
    """Sanitise a filename for safe storage. Strips path components and replaces
    non-ASCII-alphanumeric characters (except ``._-``) with underscores."""
    base = os.path.basename(filename)
    cleaned = re.sub(r"[^a-zA-Z0-9._-]", "_", base)
    # #5 fix: guard against edge cases like "..docx" or ".hidden"
    if cleaned in (".", "..") or cleaned.startswith(".") and not cleaned.startswith(".."):
        cleaned = f"_{cleaned}"
    elif cleaned.startswith(".."):
        cleaned = f"_{cleaned[1:]}"
    return cleaned


def sanitize_output_filename(filename: str) -> str:
    base = os.path.basename(filename)
    # 允许 Unicode 字母/数字（含中文）、下划线、点和短横线
    cleaned = re.sub(r"[^\w.\-]", "_", base, flags=re.UNICODE)
    # #5 fix: same edge-case guard
    if cleaned in (".", "..") or (cleaned.startswith(".") and not cleaned.startswith("..")):
        cleaned = f"_{cleaned}"
    elif cleaned.startswith(".."):
        cleaned = f"_{cleaned[1:]}"
    return cleaned


def normalize_output_docx_filename(filename: str) -> str:
    raw = (filename or "").strip()
    if not raw:
        now = datetime.now().strftime("%Y%m%d_%H%M%S")
        uid = uuid.uuid4().hex[:8]
        return f"generated_{now}_{uid}.docx"

    cleaned = sanitize_output_filename(raw)
    if not cleaned:
        now = datetime.now().strftime("%Y%m%d_%H%M%S")
        uid = uuid.uuid4().hex[:8]
        return f"generated_{now}_{uid}.docx"
    if not cleaned.lower().endswith(".docx"):
        cleaned = f"{cleaned}.docx"
    return cleaned


# ---------------------------------------------------------------------------
# DOCX text extraction (shared helper — #12 fix)
# ---------------------------------------------------------------------------

def _extract_docx_text(template_path: Path) -> str:
    """Read all text from a .docx file: paragraphs + table cells."""
    doc = Document(str(template_path))
    text_blocks: List[str] = []

    for paragraph in doc.paragraphs:
        text_blocks.append(paragraph.text)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    text_blocks.append(paragraph.text)

    return "\n".join(text_blocks)


def extract_template_variables(template_path: Path) -> List[str]:
    full_text = _extract_docx_text(template_path)
    vars_found = re.findall(r"{{\s*([a-zA-Z_][a-zA-Z0-9_]*)\s*}}", full_text)

    ordered: List[str] = []
    for key in vars_found:
        if key not in ordered:
            ordered.append(key)

    return ordered


def extract_invalid_template_placeholders(template_path: Path) -> List[str]:
    full_text = _extract_docx_text(template_path)
    all_placeholders = re.findall(r"{{\s*([^{}]+?)\s*}}", full_text)

    invalid: List[str] = []
    for placeholder in all_placeholders:
        if not re.fullmatch(r"[a-zA-Z_][a-zA-Z0-9_]*", placeholder):
            if placeholder not in invalid:
                invalid.append(placeholder)

    return invalid


def _validate_template(template_name: str) -> Path:
    """Shared template validation logic (#13 fix). Returns the template path."""
    template_name = sanitize_filename(template_name)
    template_path = UPLOAD_DIR / template_name

    if not template_path.exists():
        raise HTTPException(status_code=404, detail="模板不存在，请先上传模板")

    invalid_placeholders = extract_invalid_template_placeholders(template_path)
    if invalid_placeholders:
        raise HTTPException(
            status_code=400,
            detail="模板中存在非法占位符: "
            + str(invalid_placeholders)
            + "。请改为 {{story_acceptance_criteria}} 这类格式。",
        )

    variables = extract_template_variables(template_path)
    if not variables:
        raise HTTPException(
            status_code=400,
            detail="模板中未检测到有效占位符，请在 .docx 中使用 {{variable_name}} 格式。",
        )

    return template_path, variables


def _resolve_output_path(raw_filename: str) -> Path:
    """Shared output path resolution (#13 fix)."""
    docx_filename = normalize_output_docx_filename(raw_filename)
    docx_path = OUTPUT_DIR / docx_filename
    if docx_path.exists():
        uid = uuid.uuid4().hex[:8]
        stem = Path(docx_filename).stem
        docx_filename = f"{stem}_{uid}.docx"
        docx_path = OUTPUT_DIR / docx_filename
    return docx_path, docx_filename


def _validate_docx_magic_bytes(content: bytes) -> None:
    """#9 fix: verify file is a real ZIP / DOCX by checking magic bytes."""
    if not content[:2] == DOCX_MAGIC_BYTES:
        raise HTTPException(status_code=400, detail="文件内容不是有效的 .docx 格式（缺少 ZIP 头）")


# ---------------------------------------------------------------------------
# JSON extraction from AI response
# ---------------------------------------------------------------------------

def extract_json_from_text(text: str) -> Dict[str, str]:
    text = text.strip()
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass

    match = re.search(r"\{[\s\S]*\}", text)
    if not match:
        raise ValueError("AI 返回内容中未找到 JSON")

    return json.loads(match.group(0))


# ---------------------------------------------------------------------------
# AI call with retry + backoff (#16 fix)
# ---------------------------------------------------------------------------

def call_ai_generate_json(prompt: str, api_url: str, api_key: str, model: str, variables: List[str]) -> Dict[str, str]:
    schema = {k: "" for k in variables}

    system_prompt = (
        "你是一个严谨的 JSON 生成器。"
        "只允许输出一个 JSON 对象，不允许任何额外文字、Markdown、代码块。"
        "必须包含给定字段，且字段名完全一致。"
    )

    user_prompt = (
        f"用户提示词：{prompt}\n\n"
        f"请基于提示词生成 JSON，字段必须严格如下：\n"
        f"{json.dumps(schema, ensure_ascii=False, indent=2)}\n\n"
        f"要求：\n"
        f"1) 仅输出 JSON 对象\n"
        f"2) 所有字段必须存在\n"
        f"3) 字段值使用简洁中文文本"
    )

    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        "temperature": 0.2,
    }

    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }

    last_exception: Exception | None = None

    for attempt in range(1, AI_MAX_RETRIES + 1):
        try:
            resp = requests.post(api_url, headers=headers, json=payload, timeout=AI_REQUEST_TIMEOUT)
        except requests.RequestException as e:
            last_exception = e
            logger.warning("AI request attempt %d/%d failed: %s", attempt, AI_MAX_RETRIES, e)
            if attempt < AI_MAX_RETRIES:
                time.sleep(AI_RETRY_BACKOFF * attempt)
            continue

        if resp.status_code >= 500:
            # Server error — retryable
            last_exception = HTTPException(
                status_code=resp.status_code,
                detail=f"AI 接口服务端错误: {resp.text[:500]}",
            )
            logger.warning("AI 5xx on attempt %d/%d: %s", attempt, AI_MAX_RETRIES, resp.text[:200])
            if attempt < AI_MAX_RETRIES:
                time.sleep(AI_RETRY_BACKOFF * attempt)
            continue

        if resp.status_code >= 400:
            # Client error (4xx) — not retryable
            detail = resp.text[:1000]
            raise HTTPException(status_code=resp.status_code, detail=f"AI 接口错误: {detail}")

        try:
            data = resp.json()
            content = data["choices"][0]["message"]["content"]
        except Exception:
            raise HTTPException(status_code=500, detail=f"AI 响应格式异常: {resp.text[:1000]}")

        try:
            parsed = extract_json_from_text(content)
        except Exception as e:
            raise HTTPException(status_code=500, detail=f"AI 未返回合法 JSON: {str(e)}; 原始内容: {content[:800]}")

        missing = [k for k in variables if k not in parsed]
        if missing:
            raise HTTPException(status_code=500, detail=f"AI 返回缺少字段: {missing}")

        for k in variables:
            if parsed.get(k) is None:
                parsed[k] = ""
            else:
                parsed[k] = str(parsed[k])

        return parsed

    # All retries exhausted
    raise last_exception if last_exception else HTTPException(status_code=502, detail="AI 请求失败: 重试次数耗尽")


# ---------------------------------------------------------------------------
# DOCX rendering
# ---------------------------------------------------------------------------

def render_docx(template_path: Path, context: Dict[str, str], output_docx: Path) -> None:
    tpl = DocxTemplate(str(template_path))
    try:
        tpl.render(context)
    except TemplateSyntaxError as e:
        raise HTTPException(
            status_code=400,
            detail="模板语法错误: "
            + str(e)
            + "。请使用 {{变量名}}，变量名仅支持字母/数字/下划线，且不能包含空格。",
        )
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"模板渲染失败: {str(e)}")
    tpl.save(str(output_docx))


# ---------------------------------------------------------------------------
# PDF conversion (#3 fix: timeout, #17 fix: serialised via lock)
# ---------------------------------------------------------------------------

def convert_docx_to_pdf(docx_path: Path, pdf_path: Path) -> None:
    libreoffice = shutil.which("libreoffice")

    if libreoffice:
        with libreoffice_lock:  # #17 fix: serialise LibreOffice calls
            cmd = [
                libreoffice,
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                str(pdf_path.parent),
                str(docx_path),
            ]
            # #3 fix: add timeout to prevent hanging
            completed = subprocess.run(
                cmd, capture_output=True, text=True, timeout=LIBREOFFICE_TIMEOUT,
            )
            if completed.returncode != 0:
                raise HTTPException(status_code=500, detail=f"PDF 转换失败: {completed.stderr or completed.stdout}")

            converted = pdf_path.parent / f"{docx_path.stem}.pdf"
            if converted.exists() and converted != pdf_path:
                converted.rename(pdf_path)
            if not pdf_path.exists():
                raise HTTPException(status_code=500, detail="PDF 转换失败: 未生成目标文件")
            return

    docx2pdf_cmd = shutil.which("docx2pdf")
    if docx2pdf_cmd:
        cmd = [docx2pdf_cmd, str(docx_path), str(pdf_path)]
        completed = subprocess.run(cmd, capture_output=True, text=True, timeout=LIBREOFFICE_TIMEOUT)
        if completed.returncode != 0 or not pdf_path.exists():
            raise HTTPException(status_code=500, detail=f"PDF 转换失败(docx2pdf): {completed.stderr or completed.stdout}")
        return

    raise HTTPException(status_code=500, detail="未检测到 libreoffice 或 docx2pdf，无法生成 PDF")


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.get("/")
async def index(request: Request):
    return templates.TemplateResponse(request, "index.html", {"request": request})


@app.post("/upload-template")
async def upload_template(file: UploadFile = File(...)):
    if not file.filename.lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="仅支持 .docx 模板")

    filename = sanitize_filename(file.filename)
    save_path = UPLOAD_DIR / filename

    content = await file.read()

    # #8 fix: enforce file size limit
    if len(content) > MAX_UPLOAD_SIZE:
        raise HTTPException(status_code=413, detail=f"文件过大，最大支持 {MAX_UPLOAD_SIZE // (1024 * 1024)}MB")

    # #9 fix: verify file is a real ZIP/DOCX by checking magic bytes
    if not content[:2] == DOCX_MAGIC_BYTES:
        raise HTTPException(status_code=400, detail="文件内容不是有效的 .docx 格式（缺少 ZIP 头）")

    with open(save_path, "wb") as f:
        f.write(content)

    return JSONResponse(
        {
            "message": "模板上传成功",
            "template_name": filename,
        }
    )


def generate_common(req: GenerateRequest) -> Dict[str, str]:
    # #13 fix: use shared validation helper
    template_path, variables = _validate_template(req.template_name)

    ai_result = call_ai_generate_json(
        prompt=req.prompt,
        api_url=req.api_url,
        api_key=req.api_key,
        model=req.model,
        variables=variables,
    )

    docx_path, docx_filename = _resolve_output_path(req.output_filename)

    render_docx(template_path, ai_result, docx_path)

    return {
        "docx_filename": docx_filename,
        "docx_path": str(docx_path),
        "variables": variables,
        "ai_result": ai_result,
    }


def generate_default_docx(template_name: str, output_filename: str) -> str:
    # #13 fix: use shared validation helper
    template_path, variables = _validate_template(template_name)

    context = {k: "" for k in variables}

    docx_path, docx_filename = _resolve_output_path(output_filename)

    render_docx(template_path, context, docx_path)
    return docx_filename


def resolve_column_index(column: str, header_row: List[str]) -> int:
    value = (column or "").strip()
    if not value:
        raise HTTPException(status_code=400, detail="Excel 列不能为空")

    if re.fullmatch(r"[A-Za-z]+", value):
        idx = 0
        for ch in value.upper():
            idx = idx * 26 + (ord(ch) - ord("A") + 1)
        return idx - 1

    needle = value.lower()
    for i, name in enumerate(header_row):
        if str(name or "").strip().lower() == needle:
            return i

    raise HTTPException(status_code=400, detail=f"未找到列: {column}")


def build_row_prompt(base_prompt: str, cell_value: str) -> str:
    text = str(cell_value or "").strip()
    if not text:
        return ""

    if not str(base_prompt or "").strip():
        return text

    tokens = ("{{excel_value}}", "{excel_value}", "{{value}}", "{value}")
    row_prompt = base_prompt
    replaced = False
    for token in tokens:
        if token in row_prompt:
            row_prompt = row_prompt.replace(token, text)
            replaced = True

    if not replaced:
        row_prompt = f"{base_prompt}\n\nExcel内容: {text}"

    return row_prompt.strip()


# ---------------------------------------------------------------------------
# Retryable errors for Excel batch generation (#6 fix)
# ---------------------------------------------------------------------------

_RETRYABLE_STATUS_CODES = {429, 500, 502, 503, 504}


def _is_retryable(exc: Exception) -> bool:
    """Determine whether an exception is worth retrying."""
    if isinstance(exc, HTTPException):
        return exc.status_code in _RETRYABLE_STATUS_CODES
    if isinstance(exc, requests.RequestException):
        return True
    return False


def _generate_with_retry(req: GenerateRequest, max_retries: int = 3) -> Dict[str, str] | None:
    """Attempt to generate a docx with structured retry (#6 fix).

    Returns the result dict on success, or None if all retries exhausted.
    Non-retryable errors are re-raised immediately.
    """
    for attempt in range(1, max_retries + 1):
        try:
            return generate_common(req)
        except HTTPException as exc:
            if not _is_retryable(exc):
                # Non-retryable — re-raise to caller
                raise
            logger.warning(
                "Row generation attempt %d/%d failed (retryable): %s",
                attempt, max_retries, exc.detail,
            )
            if attempt < max_retries:
                time.sleep(AI_RETRY_BACKOFF * attempt)
        except requests.RequestException as exc:
            logger.warning(
                "Row generation attempt %d/%d failed (network): %s",
                attempt, max_retries, exc,
            )
            if attempt < max_retries:
                time.sleep(AI_RETRY_BACKOFF * attempt)

    return None


def run_docx_generation_task(record_id: str, request_payload: Dict[str, str]) -> None:
    update_generation_record(
        record_id,
        status="processing",
        message="后台正在生成 Word",
        error="",
    )

    try:
        req = GenerateRequest(**request_payload)
        result = generate_common(req)
        filename = result["docx_filename"]
        update_generation_record(
            record_id,
            status="completed",
            message="Word 生成完成",
            generated_filename=filename,
            filename=filename,
            download_url=f"/download/{filename}",
            error="",
        )
    except HTTPException as exc:
        detail = exc.detail if isinstance(exc.detail, str) else json.dumps(exc.detail, ensure_ascii=False)
        logger.error("Generation task %s failed (HTTP %d): %s", record_id, exc.status_code, detail)
        update_generation_record(
            record_id,
            status="failed",
            message="Word 生成失败",
            error=detail,
        )
    except Exception as exc:
        logger.exception("Generation task %s failed with unexpected error", record_id)
        update_generation_record(
            record_id,
            status="failed",
            message="Word 生成失败",
            error=str(exc),
        )


@app.post("/generate-docx")
async def generate_docx(req: GenerateRequest, background_tasks: BackgroundTasks):
    record = add_generation_record(req)
    # #19 fix: use model_dump() directly (Pydantic v2)
    payload = req.model_dump()
    background_tasks.add_task(run_docx_generation_task, record["id"], payload)

    return {
        "message": "Word 已提交后台生成",
        "record": record,
    }


@app.post("/generate-docx-from-excel")
async def generate_docx_from_excel(
    file: UploadFile = File(...),
    prompt: str = Form(""),
    api_url: str = Form(...),
    api_key: str = Form(...),
    model: str = Form(...),
    template_name: str = Form(...),
    sheet_name: str = Form("Sheet1"),
    prompt_column: str = Form(...),
    filename_column: str = Form(...),
    output_prefix: str = Form("excel_generated"),
):
    if not file.filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="仅支持 .xlsx Excel 文件")

    # #8 fix: read with size limit
    content = await file.read()
    if len(content) > MAX_UPLOAD_SIZE:
        raise HTTPException(status_code=413, detail=f"文件过大，最大支持 {MAX_UPLOAD_SIZE // (1024 * 1024)}MB")

    temp_excel_path = UPLOAD_DIR / f"excel_{uuid.uuid4().hex}.xlsx"
    with open(temp_excel_path, "wb") as f:
        f.write(content)

    workbook = None
    try:
        workbook = load_workbook(filename=str(temp_excel_path), read_only=True, data_only=True)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Excel 打开失败: {str(exc)}")
    try:
        if sheet_name not in workbook.sheetnames:
            raise HTTPException(status_code=400, detail=f"未找到 sheet: {sheet_name}")

        ws = workbook[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            raise HTTPException(status_code=400, detail="Excel sheet 为空")

        prompt_col_is_letter = bool(re.fullmatch(r"[A-Za-z]+", (prompt_column or "").strip()))
        filename_col_is_letter = bool(re.fullmatch(r"[A-Za-z]+", (filename_column or "").strip()))
        use_header = not (prompt_col_is_letter and filename_col_is_letter)

        if not use_header:
            prompt_col_index = resolve_column_index(prompt_column, [])
            filename_col_index = resolve_column_index(filename_column, [])
            data_rows = list(enumerate(rows, start=1))
        else:
            header_row = [str(item or "").strip() for item in rows[0]]
            prompt_col_index = resolve_column_index(prompt_column, header_row)
            filename_col_index = resolve_column_index(filename_column, header_row)
            data_rows = list(enumerate(rows[1:], start=2))

        generated_files = []
        processed_rows = 0
        fallback_count = 0
        safe_prefix = sanitize_filename(output_prefix or "excel_generated")

        for row_num, row in data_rows:
            prompt_text = ""
            file_name_text = ""
            if prompt_col_index < len(row):
                prompt_text = str(row[prompt_col_index] or "").strip()
            if filename_col_index < len(row):
                file_name_text = str(row[filename_col_index] or "").strip()

            if not prompt_text:
                continue

            row_prompt = build_row_prompt(prompt, prompt_text)
            if not row_prompt:
                continue

            output_name = sanitize_output_filename(file_name_text) if file_name_text else ""
            if output_name and not output_name.lower().endswith(".docx"):
                output_name = f"{output_name}.docx"
            if not output_name:
                output_name = f"{safe_prefix}_row{row_num}.docx"

            req = GenerateRequest(
                prompt=row_prompt,
                api_url=api_url,
                api_key=api_key,
                model=model,
                template_name=template_name,
                output_filename=output_name,
            )

            # #6 fix: structured retry with logging — non-retryable errors re-raise
            result = _generate_with_retry(req, max_retries=AI_MAX_RETRIES)

            if result is not None:
                generated_files.append(result["docx_filename"])
                processed_rows += 1
                continue

            # All retries exhausted — fall back to blank template
            logger.warning("Row %d: all retries failed, using fallback", row_num)
            fail_name = output_name
            if fail_name.lower().endswith(".docx"):
                fail_name = f"{fail_name[:-5]}_fail.docx"
            else:
                fail_name = f"{fail_name}_fail.docx"
            default_filename = generate_default_docx(template_name, fail_name)
            generated_files.append(default_filename)
            processed_rows += 1
            fallback_count += 1

        if not generated_files:
            raise HTTPException(status_code=400, detail="提示词列没有可用数据，未生成文件")

        zip_filename = f"{safe_prefix}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
        zip_path = OUTPUT_DIR / zip_filename
        with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for filename in generated_files:
                source = OUTPUT_DIR / filename
                if source.exists():
                    zf.write(source, arcname=filename)

        return {
            "message": f"批量生成完成，共 {processed_rows} 个 Word，失败回退 {fallback_count} 个",
            "count": processed_rows,
            "fallback_count": fallback_count,
            "sheet_name": sheet_name,
            "prompt_column": prompt_column,
            "filename_column": filename_column,
            "zip_filename": zip_filename,
            "download_url": f"/download/{zip_filename}",
            "files": generated_files,
        }
    finally:
        if workbook is not None:
            workbook.close()
        temp_excel_path.unlink(missing_ok=True)


@app.get("/generation-records")
async def get_generation_records(limit: int = 50):
    safe_limit = max(1, min(limit, 200))
    return {"records": list_generation_records(safe_limit)}


@app.get("/generation-records/{record_id}")
async def get_generation_record_by_id(record_id: str):
    return {"record": get_generation_record(record_id)}


@app.post("/generate-pdf")
async def generate_pdf(req: GenerateRequest):
    # #14 fix: offload blocking work to threadpool to avoid blocking the event loop
    result = await run_in_threadpool(generate_common, req)

    docx_filename = result["docx_filename"]
    docx_path = Path(result["docx_path"])
    pdf_filename = docx_filename.replace(".docx", ".pdf")
    pdf_path = OUTPUT_DIR / pdf_filename

    # PDF conversion is also blocking — offload to threadpool
    await run_in_threadpool(convert_docx_to_pdf, docx_path, pdf_path)

    return {
        "message": "PDF 生成成功",
        "filename": pdf_filename,
        "download_url": f"/download/{pdf_filename}",
        "variables": result["variables"],
    }


@app.get("/download/{filename}")
async def download_file(filename: str):
    filename = sanitize_filename(filename)
    file_path = OUTPUT_DIR / filename
    if not file_path.exists():
        raise HTTPException(status_code=404, detail="文件不存在")

    media_type = "application/octet-stream"
    if filename.lower().endswith(".pdf"):
        media_type = "application/pdf"
    elif filename.lower().endswith(".docx"):
        media_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    elif filename.lower().endswith(".zip"):
        media_type = "application/zip"

    return FileResponse(str(file_path), media_type=media_type, filename=filename)


# #4 fix: don't leak internal error details to clients
@app.exception_handler(Exception)
async def unhandled_exception_to_json(request: Request, exc: Exception):
    logger.exception("Unhandled exception on %s %s", request.method, request.url.path)
    return JSONResponse(status_code=500, content={"detail": "Internal Server Error"})
