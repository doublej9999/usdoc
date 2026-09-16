# -*- coding: utf-8 -*-
import pytest
from fastapi import HTTPException

from utils.excel import build_row_prompt, resolve_column_index


def test_resolve_column_index_letters():
    assert resolve_column_index("A", []) == 0
    assert resolve_column_index("B", []) == 1
    assert resolve_column_index("Z", []) == 25
    assert resolve_column_index("AA", []) == 26
    assert resolve_column_index("AB", []) == 27
    assert resolve_column_index("a", []) == 0


def test_resolve_column_index_prefers_matching_header():
    headers = ["用户ID", "用户故事描述", "优先级", "Status"]
    assert resolve_column_index("用户ID", headers) == 0
    assert resolve_column_index("用户故事描述", headers) == 1
    assert resolve_column_index("status", headers) == 3
    assert resolve_column_index("  用户ID  ", headers) == 0


def test_resolve_column_index_rejects_unknown_column():
    with pytest.raises(HTTPException) as empty:
        resolve_column_index("", ["ColA"])
    assert empty.value.status_code == 400

    with pytest.raises(HTTPException) as missing:
        resolve_column_index("不存在的列名", ["ColA", "ColB"])
    assert missing.value.status_code == 400


def test_build_row_prompt_replaces_placeholder_token():
    base = "请根据此需求生成用户故事：{excel_value}，要求详细"
    assert build_row_prompt(base, "用户注册登录流程") == "请根据此需求生成用户故事：用户注册登录流程，要求详细"
    assert build_row_prompt("请分析需求：{{excel_value}}", "用户注册") == "请分析需求：用户注册"


def test_build_row_prompt_falls_back_to_concatenation():
    result = build_row_prompt("请为以下模块生成测试用例", "学生成绩导入功能")
    assert "请为以下模块生成测试用例" in result
    assert "学生成绩导入功能" in result


def test_build_row_prompt_handles_empty_inputs():
    assert build_row_prompt("prompt", "") == ""
    assert build_row_prompt("prompt", "   ") == ""
    assert build_row_prompt("", "学生成绩导入功能") == "学生成绩导入功能"


if __name__ == "__main__":
    pytest.main([__file__])
