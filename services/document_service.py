import shutil
import subprocess
from pathlib import Path
from typing import Any, Dict, List, Tuple

from docxtpl import DocxTemplate
from fastapi import HTTPException
from jinja2 import TemplateSyntaxError

from services.ai_service import call_ai_generate_json
from utils.files import discard_output_path, ensure_unique_docx_path, sanitize_filename
from utils.templates import extract_invalid_template_placeholders, extract_template_variables


def validate_template(template_name: str, upload_dir: Path) -> Tuple[Path, List[str]]:
    safe_template_name = sanitize_filename(template_name)
    template_path = upload_dir / safe_template_name

    if not template_path.exists():
        # Fallback to upload_dir.parent (e.g. project root / BASE_DIR for built-in templates like US.docx)
        fallback = upload_dir.parent / safe_template_name
        if fallback.exists():
            template_path = fallback
        else:
            raise HTTPException(status_code=404, detail="模板不存在，请先上传模板")

    invalid_placeholders = extract_invalid_template_placeholders(template_path)
    if invalid_placeholders:
        raise HTTPException(
            status_code=400,
            detail=f"模板中存在非法占位符: {invalid_placeholders}。请改为 {{story_acceptance_criteria}} 这类格式。",
        )

    variables = extract_template_variables(template_path)
    if not variables:
        raise HTTPException(
            status_code=400,
            detail="模板中未检测到有效占位符，请在 .docx 中使用 {{variable_name}} 格式。",
        )

    return template_path, variables


def render_docx(template_path: Path, context: Dict[str, Any], output_docx: Path) -> None:
    try:
        tpl = DocxTemplate(str(template_path))
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"无法加载模板文件: {str(exc)}")

    try:
        tpl.render(context)
    except TemplateSyntaxError as exc:
        raise HTTPException(
            status_code=400,
            detail=f"模板语法错误: {str(exc)}。请使用 {{变量名}}，变量名仅支持字母/数字/下划线，且不能包含空格。",
        )
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"模板渲染失败: {str(exc)}")

    try:
        output_docx.parent.mkdir(parents=True, exist_ok=True)
        tpl.save(str(output_docx))
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"保存生成的文档失败: {str(exc)}")


def _render_reserved(template_path: Path, context: Dict[str, Any], output_docx: Path) -> None:
    """Render into a path claimed by ``ensure_unique_docx_path``.

    The reservation is an empty file, so any failure must release it — otherwise
    a zero-byte ``.docx`` would linger in ``outputs/`` and stay downloadable.
    """
    try:
        render_docx(template_path, context, output_docx)
    except BaseException:
        discard_output_path(output_docx)
        raise


def generate_document(
    *,
    prompt: str,
    api_url: str,
    api_key: str,
    model: str,
    template_name: str,
    output_filename: str,
    upload_dir: Path,
    output_dir: Path,
    stream: bool = False,
    temperature: float = 0.2,
) -> Dict[str, Any]:
    template_path, variables = validate_template(template_name, upload_dir)
    ai_result = call_ai_generate_json(
        prompt=prompt,
        api_url=api_url,
        api_key=api_key,
        model=model,
        variables=variables,
        stream=stream,
        temperature=temperature,
    )
    docx_filename, docx_path = ensure_unique_docx_path(output_dir, output_filename)
    _render_reserved(template_path, ai_result, docx_path)

    return {
        "docx_filename": docx_filename,
        "docx_path": str(docx_path),
        "variables": variables,
        "ai_result": ai_result,
    }


def generate_default_document(
    *,
    template_name: str,
    output_filename: str,
    upload_dir: Path,
    output_dir: Path,
) -> str:
    template_path, variables = validate_template(template_name, upload_dir)
    context = {key: "" for key in variables}
    docx_filename, docx_path = ensure_unique_docx_path(output_dir, output_filename)
    _render_reserved(template_path, context, docx_path)
    return docx_filename


def convert_docx_to_pdf(docx_path: Path, pdf_path: Path) -> None:
    libreoffice = shutil.which("libreoffice") or shutil.which("soffice")

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
            raise HTTPException(
                status_code=500,
                detail=f"PDF 转换失败: {completed.stderr or completed.stdout}",
            )

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
            raise HTTPException(
                status_code=500,
                detail=f"PDF 转换失败 (docx2pdf): {completed.stderr or completed.stdout}",
            )
        return

    raise HTTPException(
        status_code=500,
        detail="未检测到 libreoffice 或 docx2pdf，无法生成 PDF",
    )
