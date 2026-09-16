import re
from typing import List

from fastapi import HTTPException


def resolve_column_index(column: str, header_row: List[str]) -> int:
    value = (column or "").strip()
    if not value:
        raise HTTPException(status_code=400, detail="Excel 列不能为空")

    needle = value.lower()
    for i, name in enumerate(header_row):
        if str(name or "").strip().lower() == needle:
            return i

    if re.fullmatch(r"[A-Za-z]+", value):
        idx = 0
        for ch in value.upper():
            idx = idx * 26 + (ord(ch) - ord("A") + 1)
        return idx - 1

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
