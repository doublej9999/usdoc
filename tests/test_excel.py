# -*- coding: utf-8 -*-
import sys
from pathlib import Path

# Ensure project root is in sys.path when running standalone
PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import pytest
from fastapi import HTTPException
from utils.excel import resolve_column_index, build_row_prompt


def test_resolve_column_index_letters():
    headers = ["Col1", "Col2", "Col3"]
    assert resolve_column_index("A", headers) == 0
    assert resolve_column_index("B", headers) == 1
    assert resolve_column_index("Z", headers) == 25
    assert resolve_column_index("AA", headers) == 26
    assert resolve_column_index("AB", headers) == 27
    assert resolve_column_index("a", headers) == 0
    assert resolve_column_index("aa", headers) == 26


def test_resolve_column_index_headers():
    headers = ["用户ID", "用户故事描述", "优先级", "Status"]
    assert resolve_column_index("用户ID", headers) == 0
    assert resolve_column_index("用户故事描述", headers) == 1
    assert resolve_column_index("优先级", headers) == 2
    assert resolve_column_index("Status", headers) == 3
    assert resolve_column_index("status", headers) == 3
    assert resolve_column_index("  用户ID  ", headers) == 0


def test_resolve_column_index_errors():
    headers = ["ColA", "ColB"]
    with pytest.raises(HTTPException) as exc_info_empty:
        resolve_column_index("", headers)
    assert exc_info_empty.value.status_code == 400

    with pytest.raises(HTTPException) as exc_info_missing:
        resolve_column_index("不存在的列名", headers)
    assert exc_info_missing.value.status_code == 400


def test_build_row_prompt_token_replacement():
    base_prompt = "请根据此需求生成用户故事：{excel_value}，要求详细"
    cell_value = "用户注册登录流程"
    result = build_row_prompt(base_prompt, cell_value)
    assert result == "请根据此需求生成用户故事：用户注册登录流程，要求详细"

    # Also test {{excel_value}}
    base_prompt_double = "请分析需求：{{excel_value}}"
    result_double = build_row_prompt(base_prompt_double, cell_value)
    assert result_double == "请分析需求：用户注册登录流程"


def test_build_row_prompt_fallback_concatenation():
    base_prompt = "请为以下模块生成测试用例"
    cell_value = "学生成绩导入功能"
    result = build_row_prompt(base_prompt, cell_value)
    assert "请为以下模块生成测试用例" in result
    assert "学生成绩导入功能" in result
    assert "Excel" in result


def test_build_row_prompt_empty_inputs():
    assert build_row_prompt("prompt", "") == ""
    assert build_row_prompt("prompt", "   ") == ""
    assert build_row_prompt("", "学生成绩导入功能") == "学生成绩导入功能"


if __name__ == "__main__":
    pytest.main([__file__])
