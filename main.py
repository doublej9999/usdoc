import json
import os
import re
import shutil
import subprocess
import sys
import uuid
from datetime import datetime
from pathlib import Path
from threading import Lock
from typing import Dict, List

import requests
from docx import Document
from docxtpl import DocxTemplate
from fastapi import BackgroundTasks, FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from fastapi import Request
from jinja2 import TemplateSyntaxError
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


generation_records: Dict[str, Dict[str, str]] = {}
generation_records_lock = Lock()


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


def sanitize_filename(filename: str) -> str:
    base = os.path.basename(filename)
    return re.sub(r"[^a-zA-Z0-9._-]", "_", base)


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

    now = datetime.now().strftime("%Y%m%d_%H%M%S")
    uid = uuid.uuid4().hex[:8]
    docx_filename = f"generated_{now}_{uid}.docx"
    docx_path = OUTPUT_DIR / docx_filename

    render_docx(template_path, ai_result, docx_path)

    return {
        "docx_filename": docx_filename,
        "docx_path": str(docx_path),
        "variables": variables,
        "ai_result": ai_result,
    }


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

    return FileResponse(str(file_path), media_type=media_type, filename=filename)


@app.exception_handler(Exception)
async def unhandled_exception_to_json(request: Request, exc: Exception):
    return JSONResponse(status_code=500, content={"detail": f"Internal Server Error: {str(exc)}"})
