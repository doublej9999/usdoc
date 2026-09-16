# -*- coding: utf-8 -*-
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import pytest
from fastapi.testclient import TestClient
from main import app

client = TestClient(app)


def test_index_route():
    response = client.get("/")
    assert response.status_code == 200
    assert "AI 文档生成系统" in response.text


def test_list_templates_route():
    response = client.get("/templates")
    assert response.status_code == 200
    data = response.json()
    assert "templates" in data
    assert isinstance(data["templates"], list)
    template_names = [t["name"] for t in data["templates"]]
    assert "US.docx" in template_names


def test_template_variables_route():
    response = client.get("/templates/US.docx/variables")
    assert response.status_code == 200
    data = response.json()
    assert data["template_name"] == "US.docx"
    assert "variables" in data
    assert "story_title" in data["variables"]


def test_generation_records_crud():
    # List records
    response = client.get("/generation-records")
    assert response.status_code == 200
    assert "records" in response.json()

    # Clear records
    clear_resp = client.delete("/generation-records")
    assert clear_resp.status_code == 200

    # List again -> empty
    response_after = client.get("/generation-records")
    assert response_after.status_code == 200
    assert response_after.json()["records"] == []


def test_download_chinese_filename(tmp_path: Path):
    from config import OUTPUT_DIR
    chinese_file = OUTPUT_DIR / "测试_中文文档.docx"
    chinese_file.write_bytes(b"dummy docx content")

    try:
        response = client.get("/download/测试_中文文档.docx")
        assert response.status_code == 200
        assert response.content == b"dummy docx content"
    finally:
        chinese_file.unlink(missing_ok=True)
