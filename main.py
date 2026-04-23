import json
import os
import re
import shutil
import subprocess
import uuid
from datetime import datetime
from pathlib import Path
from typing import Dict, List

import requests
from docx import Document
from docxtpl import DocxTemplate
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from fastapi import Request
from pydantic import BaseModel

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"
TEMPLATES_DIR = BASE_DIR / "templates"
STATIC_DIR = BASE_DIR / "static"

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
    tpl.render(context)
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

    variables = extract_template_variables(template_path)
    if not variables:
        variables = ["company_name", "service_plan", "price", "timeline"]

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


@app.post("/generate-docx")
async def generate_docx(req: GenerateRequest):
    result = generate_common(req)
    filename = result["docx_filename"]
    return {
        "message": "Word 生成成功",
        "filename": filename,
        "download_url": f"/download/{filename}",
        "variables": result["variables"],
    }


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
