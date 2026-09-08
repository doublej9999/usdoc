import json
import os
import re
import shutil
import subprocess
import sys
import uuid
import zipfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path
from threading import Lock
from typing import Dict, List, Optional

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


class GenerateRequest(BaseModel):
    prompt: str
    api_url: str
    api_key: str
    model: str
    template_name: str
    output_filename: str = ""


generation_records: Dict[str, Dict[str, str]] = {}
generation_records_lock = Lock()

# 批量任务全局存储与锁
BATCH_TASKS_FILE = DATA_DIR / "batch_tasks.json"
batch_tasks: Dict[str, Dict] = {}
batch_tasks_lock = Lock()


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
        "api_url": req.api_url,
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


def dump_batch_tasks() -> None:
    with batch_tasks_lock:
        tasks = []
        for item in batch_tasks.values():
            copy_item = dict(item)
            # 安全脱敏 api_key，不将明文 key 落盘
            if "api_key" in copy_item:
                key = copy_item["api_key"]
                copy_item["api_key"] = f"{key[:3]}***{key[-3:]}" if len(key) > 6 else "***"
            tasks.append(copy_item)
        tasks.sort(key=lambda x: x.get("created_at", ""), reverse=True)
        try:
            BATCH_TASKS_FILE.write_text(
                json.dumps(tasks, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
        except Exception:
            pass


def load_batch_tasks() -> None:
    if not BATCH_TASKS_FILE.exists():
        return

    try:
        tasks = json.loads(BATCH_TASKS_FILE.read_text(encoding="utf-8"))
        if not isinstance(tasks, list):
            return
    except Exception:
        return

    with batch_tasks_lock:
        batch_tasks.clear()
        for task in tasks:
            if isinstance(task, dict) and task.get("id"):
                batch_tasks[task["id"]] = task


load_batch_tasks()


def append_batch_log(batch_id: str, level: str, message: str) -> None:
    with batch_tasks_lock:
        task = batch_tasks.get(batch_id)
        if not task:
            return
        log_entry = {
            "time": datetime.now().strftime("%H:%M:%S"),
            "level": level,
            "message": message,
        }
        logs = task.setdefault("logs", [])
        logs.append(log_entry)
        if len(logs) > 500:
            task["logs"] = logs[-500:]


def update_batch_task_summary(batch_id: str) -> None:
    with batch_tasks_lock:
        task = batch_tasks.get(batch_id)
        if not task:
            return
        items = task.get("items", {})
        total = len(items)
        success_count = sum(1 for it in items.values() if it.get("status") == "success")
        failed_count = sum(1 for it in items.values() if it.get("status") == "failed")
        running_count = sum(1 for it in items.values() if it.get("status") == "running")
        pending_count = sum(1 for it in items.values() if it.get("status") == "pending")
        completed_count = success_count + failed_count

        task["total"] = total
        task["completed"] = completed_count
        task["success_count"] = success_count
        task["failed_count"] = failed_count
        task["running_count"] = running_count
        task["pending_count"] = pending_count

        if running_count > 0 or pending_count > 0:
            task["status"] = "running"
        elif failed_count > 0 and success_count > 0:
            task["status"] = "partial_failed"
        elif failed_count > 0 and success_count == 0:
            task["status"] = "failed"
        else:
            task["status"] = "completed"

        task["updated_at"] = now_iso()

    dump_batch_tasks()


def repack_batch_zip(batch_id: str) -> Optional[str]:
    with batch_tasks_lock:
        task = batch_tasks.get(batch_id)
        if not task:
            return None
        items = list(task.get("items", {}).values())
        safe_prefix = sanitize_filename(task.get("output_prefix") or "batch_generated")

    success_files = [it["generated_filename"] for it in items if it.get("status") == "success" and it.get("generated_filename")]
    if not success_files:
        return None

    zip_filename = f"{safe_prefix}_{batch_id[:8]}.zip"
    zip_path = OUTPUT_DIR / zip_filename
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for fname in success_files:
            file_path = OUTPUT_DIR / fname
            if file_path.exists():
                zf.write(file_path, arcname=fname)

    with batch_tasks_lock:
        if batch_id in batch_tasks:
            batch_tasks[batch_id]["zip_filename"] = zip_filename
            batch_tasks[batch_id]["download_url"] = f"/download/{zip_filename}"

    dump_batch_tasks()
    return zip_filename


def sanitize_filename(filename: str) -> str:
    base = os.path.basename(filename)
    return re.sub(r"[^a-zA-Z0-9._-]", "_", base)


def sanitize_output_filename(filename: str) -> str:
    base = os.path.basename(filename)
    # 允许 Unicode 字母/数字（含中文）、下划线、点和短横线
    return re.sub(r"[^\w.\-]", "_", base, flags=re.UNICODE)


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


def extract_template_variables(template_path: Path) -> List[str]:
    doc = Document(str(template_path))
    text_blocks = []

    for paragraph in doc.paragraphs:
        text_blocks.append(paragraph.text)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    text_blocks.append(paragraph.text)

    full_text = "\n".join(text_blocks)
    vars_found = re.findall(r"{{\s*([a-zA-Z_][a-zA-Z0-9_]*)\s*}}", full_text)

    ordered = []
    for key in vars_found:
        if key not in ordered:
            ordered.append(key)

    return ordered


def extract_invalid_template_placeholders(template_path: Path) -> List[str]:
    doc = Document(str(template_path))
    text_blocks = []

    for paragraph in doc.paragraphs:
        text_blocks.append(paragraph.text)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    text_blocks.append(paragraph.text)

    full_text = "\n".join(text_blocks)
    all_placeholders = re.findall(r"{{\s*([^{}]+?)\s*}}", full_text)

    invalid = []
    for placeholder in all_placeholders:
        if not re.fullmatch(r"[a-zA-Z_][a-zA-Z0-9_]*", placeholder):
            if placeholder not in invalid:
                invalid.append(placeholder)

    return invalid


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

    try:
        resp = requests.post(api_url, headers=headers, json=payload, timeout=90)
    except requests.RequestException as e:
        raise HTTPException(status_code=502, detail=f"AI 请求失败: {str(e)}")

    if resp.status_code >= 400:
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


def convert_docx_to_pdf(docx_path: Path, pdf_path: Path) -> None:
    libreoffice = shutil.which("libreoffice")

    if libreoffice:
        cmd = [
            libreoffice,
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            str(pdf_path.parent),
            str(docx_path),
        ]
        completed = subprocess.run(cmd, capture_output=True, text=True)
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
        completed = subprocess.run(cmd, capture_output=True, text=True)
        if completed.returncode != 0 or not pdf_path.exists():
            raise HTTPException(status_code=500, detail=f"PDF 转换失败(docx2pdf): {completed.stderr or completed.stdout}")
        return

    raise HTTPException(status_code=500, detail="未检测到 libreoffice 或 docx2pdf，无法生成 PDF")


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
    with open(save_path, "wb") as f:
        f.write(content)

    return JSONResponse(
        {
            "message": "模板上传成功",
            "template_name": filename,
        }
    )


def generate_common(req: GenerateRequest) -> Dict[str, str]:
    template_name = sanitize_filename(req.template_name)
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

    ai_result = call_ai_generate_json(
        prompt=req.prompt,
        api_url=req.api_url,
        api_key=req.api_key,
        model=req.model,
        variables=variables,
    )

    docx_filename = normalize_output_docx_filename(req.output_filename)
    docx_path = OUTPUT_DIR / docx_filename
    if docx_path.exists():
        uid = uuid.uuid4().hex[:8]
        stem = Path(docx_filename).stem
        docx_filename = f"{stem}_{uid}.docx"
    docx_path = OUTPUT_DIR / docx_filename

    render_docx(template_path, ai_result, docx_path)

    return {
        "docx_filename": docx_filename,
        "docx_path": str(docx_path),
        "variables": variables,
        "ai_result": ai_result,
    }


def generate_default_docx(template_name: str, output_filename: str) -> str:
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
    context = {k: "" for k in variables}

    docx_filename = normalize_output_docx_filename(output_filename)
    docx_path = OUTPUT_DIR / docx_filename
    if docx_path.exists():
        uid = uuid.uuid4().hex[:8]
        stem = Path(docx_filename).stem
        docx_filename = f"{stem}_{uid}.docx"
        docx_path = OUTPUT_DIR / docx_filename

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
        update_generation_record(
            record_id,
            status="failed",
            message="Word 生成失败",
            error=detail,
        )
    except Exception as exc:
        update_generation_record(
            record_id,
            status="failed",
            message="Word 生成失败",
            error=str(exc),
        )


@app.post("/generate-docx")
async def generate_docx(req: GenerateRequest, background_tasks: BackgroundTasks):
    record = add_generation_record(req)
    payload = req.model_dump() if hasattr(req, "model_dump") else req.dict()
    background_tasks.add_task(run_docx_generation_task, record["id"], payload)

    return {
        "message": "Word 已提交后台生成",
        "record": record,
    }


def execute_single_item(batch_id: str, item_id: str) -> None:
    with batch_tasks_lock:
        task = batch_tasks.get(batch_id)
        if not task:
            return
        item = task.get("items", {}).get(item_id)
        if not item:
            return
        item["status"] = "running"
        item["error"] = ""
        template_name = task["template_name"]
        api_url = task["api_url"]
        api_key = task["api_key"]
        model = task["model"]
        row_num = item["row_num"]
        row_prompt = item["row_prompt"]
        output_filename = item["output_filename"]

    append_batch_log(batch_id, "INFO", f"[行 {row_num}] 开始生成: {output_filename}")
    update_batch_task_summary(batch_id)

    req = GenerateRequest(
        prompt=row_prompt,
        api_url=api_url,
        api_key=api_key,
        model=model,
        template_name=template_name,
        output_filename=output_filename,
    )

    max_retries = 2
    success = False
    last_err = ""
    for attempt in range(1, max_retries + 1):
        try:
            res = generate_common(req)
            success = True
            with batch_tasks_lock:
                item = batch_tasks[batch_id]["items"][item_id]
                item["status"] = "success"
                item["generated_filename"] = res["docx_filename"]
                item["error"] = ""
            append_batch_log(batch_id, "SUCCESS", f"[行 {row_num}] 生成成功: {res['docx_filename']}")
            break
        except HTTPException as he:
            last_err = he.detail if isinstance(he.detail, str) else str(he.detail)
            if attempt < max_retries:
                append_batch_log(batch_id, "WARNING", f"[行 {row_num}] 第 {attempt} 次重试失败: {last_err}")
        except Exception as ex:
            last_err = str(ex)
            if attempt < max_retries:
                append_batch_log(batch_id, "WARNING", f"[行 {row_num}] 第 {attempt} 次重试失败: {last_err}")

    if not success:
        with batch_tasks_lock:
            item = batch_tasks[batch_id]["items"][item_id]
            item["status"] = "failed"
            item["error"] = last_err
        append_batch_log(batch_id, "ERROR", f"[行 {row_num}] 生成失败: {last_err}")

    update_batch_task_summary(batch_id)


def run_batch_generation_task(batch_id: str, concurrency: int = 5) -> None:
    with batch_tasks_lock:
        task = batch_tasks.get(batch_id)
        if not task:
            return
        item_ids = [item_id for item_id, it in task.get("items", {}).items() if it.get("status") in ("pending", "failed")]

    safe_concurrency = max(1, min(concurrency, 10))
    append_batch_log(batch_id, "INFO", f"启动批处理引擎，任务项: {len(item_ids)} 个，并发数: {safe_concurrency}")

    with ThreadPoolExecutor(max_workers=safe_concurrency) as executor:
        futures = {executor.submit(execute_single_item, batch_id, item_id): item_id for item_id in item_ids}
        for future in as_completed(futures):
            try:
                future.result()
            except Exception as exc:
                append_batch_log(batch_id, "ERROR", f"执行异常: {str(exc)}")

    # 尝试更新一次压缩包
    zip_name = repack_batch_zip(batch_id)
    with batch_tasks_lock:
        t = batch_tasks.get(batch_id)
        if t:
            succ = t.get("success_count", 0)
            fail = t.get("failed_count", 0)
            if zip_name:
                append_batch_log(batch_id, "SUCCESS", f"本轮处理完成！成功: {succ}, 失败: {fail}。打包完成: {zip_name}")
            else:
                append_batch_log(batch_id, "WARNING", f"本轮处理完成！成功: {succ}, 失败: {fail}。暂无成功文件打包。")

    update_batch_task_summary(batch_id)


@app.post("/batch-tasks")
async def create_batch_task(
    background_tasks: BackgroundTasks,
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
    concurrency: int = Form(5),
):
    if not file.filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="仅支持 .xlsx Excel 文件")

    temp_excel_path = UPLOAD_DIR / f"excel_{uuid.uuid4().hex}.xlsx"
    with open(temp_excel_path, "wb") as f:
        f.write(await file.read())

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

        safe_prefix = sanitize_filename(output_prefix or "excel_generated")
        items = {}

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

            item_id = f"item_{row_num}_{uuid.uuid4().hex[:6]}"
            items[item_id] = {
                "item_id": item_id,
                "row_num": row_num,
                "prompt_text": prompt_text,
                "row_prompt": row_prompt,
                "output_filename": output_name,
                "status": "pending",
                "generated_filename": "",
                "error": "",
            }

        if not items:
            raise HTTPException(status_code=400, detail="提示词列没有可用数据，未生成任务项")

        batch_id = uuid.uuid4().hex
        now_time = now_iso()
        task = {
            "id": batch_id,
            "created_at": now_time,
            "updated_at": now_time,
            "template_name": template_name,
            "sheet_name": sheet_name,
            "prompt_column": prompt_column,
            "filename_column": filename_column,
            "output_prefix": safe_prefix,
            "api_url": api_url,
            "api_key": api_key,
            "model": model,
            "concurrency": max(1, min(concurrency, 10)),
            "status": "running",
            "total": len(items),
            "completed": 0,
            "success_count": 0,
            "failed_count": 0,
            "running_count": 0,
            "pending_count": len(items),
            "zip_filename": "",
            "download_url": "",
            "logs": [],
            "items": items,
        }

        with batch_tasks_lock:
            batch_tasks[batch_id] = task

        dump_batch_tasks()
        append_batch_log(batch_id, "INFO", f"任务创建成功，共解析出 {len(items)} 条待处理行。")

        background_tasks.add_task(run_batch_generation_task, batch_id, task["concurrency"])

        return {
            "message": "批量任务已创建并启动并发生成",
            "batch_id": batch_id,
            "total": len(items),
        }
    finally:
        if workbook is not None:
            workbook.close()
        temp_excel_path.unlink(missing_ok=True)


@app.get("/batch-tasks/{batch_id}")
async def get_batch_task(batch_id: str):
    with batch_tasks_lock:
        task = batch_tasks.get(batch_id)
        if not task:
            raise HTTPException(status_code=404, detail="批量任务不存在")
        # 复制返回，并对 key 脱敏
        res = dict(task)
        if "api_key" in res:
            k = res["api_key"]
            res["api_key"] = f"{k[:3]}***{k[-3:]}" if len(k) > 6 else "***"
        return res


@app.post("/batch-tasks/{batch_id}/retry-failed")
async def retry_failed_batch_items(batch_id: str, background_tasks: BackgroundTasks):
    with batch_tasks_lock:
        task = batch_tasks.get(batch_id)
        if not task:
            raise HTTPException(status_code=404, detail="批量任务不存在")
        failed_items = [it for it in task.get("items", {}).values() if it.get("status") == "failed"]
        if not failed_items:
            return {"message": "当前没有失败的任务项需要重试", "retried_count": 0}

        for it in failed_items:
            it["status"] = "pending"
            it["error"] = ""

        task["status"] = "running"
        concurrency = task.get("concurrency", 5)

    update_batch_task_summary(batch_id)
    append_batch_log(batch_id, "INFO", f"触发重试失败项，共 {len(failed_items)} 项重试中...")
    background_tasks.add_task(run_batch_generation_task, batch_id, concurrency)

    return {"message": f"已将 {len(failed_items)} 个失败项加入重试队列", "retried_count": len(failed_items)}


@app.post("/batch-tasks/{batch_id}/items/{item_id}/retry")
async def retry_single_batch_item(batch_id: str, item_id: str, background_tasks: BackgroundTasks):
    with batch_tasks_lock:
        task = batch_tasks.get(batch_id)
        if not task:
            raise HTTPException(status_code=404, detail="批量任务不存在")
        item = task.get("items", {}).get(item_id)
        if not item:
            raise HTTPException(status_code=404, detail="任务项不存在")

        item["status"] = "pending"
        item["error"] = ""
        task["status"] = "running"

    update_batch_task_summary(batch_id)
    append_batch_log(batch_id, "INFO", f"[行 {item.get('row_num')}] 单独重试已加入队列")

    def _single_runner():
        execute_single_item(batch_id, item_id)
        repack_batch_zip(batch_id)
        update_batch_task_summary(batch_id)

    background_tasks.add_task(_single_runner)
    return {"message": f"行 {item.get('row_num')} 已加入单独重试队列"}


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

    temp_excel_path = UPLOAD_DIR / f"excel_{uuid.uuid4().hex}.xlsx"
    with open(temp_excel_path, "wb") as f:
        f.write(await file.read())

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

            success = False
            for _ in range(5):
                try:
                    result = generate_common(req)
                    generated_files.append(result["docx_filename"])
                    processed_rows += 1
                    success = True
                    break
                except Exception:
                    pass

            if success:
                continue

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
    result = generate_common(req)

    docx_filename = result["docx_filename"]
    docx_path = Path(result["docx_path"])
    pdf_filename = docx_filename.replace(".docx", ".pdf")
    pdf_path = OUTPUT_DIR / pdf_filename

    convert_docx_to_pdf(docx_path, pdf_path)

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


@app.exception_handler(Exception)
async def unhandled_exception_to_json(request: Request, exc: Exception):
    return JSONResponse(status_code=500, content={"detail": f"Internal Server Error: {str(exc)}"})
