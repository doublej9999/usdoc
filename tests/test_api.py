# -*- coding: utf-8 -*-
import main
import pytest
from fastapi.testclient import TestClient

client = TestClient(main.app)


def test_index_route():
    response = client.get("/")
    assert response.status_code == 200
    assert "AI 文档生成系统" in response.text


def test_list_templates_route():
    response = client.get("/templates")
    assert response.status_code == 200
    data = response.json()
    assert isinstance(data["templates"], list)
    assert "US.docx" in [item["name"] for item in data["templates"]]


def test_template_variables_route():
    response = client.get("/templates/US.docx/variables")
    assert response.status_code == 200
    data = response.json()
    assert data["template_name"] == "US.docx"
    assert "story_title" in data["variables"]


def test_generation_records_start_empty_and_list():
    response = client.get("/generation-records")
    assert response.status_code == 200
    assert response.json()["records"] == []


def test_generation_record_delete_missing_returns_404():
    response = client.delete("/generation-records/does-not-exist")
    assert response.status_code == 404


def test_clear_generation_records():
    response = client.delete("/generation-records")
    assert response.status_code == 200
    assert response.json()["deleted_count"] == 0


def test_download_unicode_filename():
    target = main.OUTPUT_DIR / "测试_中文文档.docx"
    target.write_bytes(b"dummy docx content")

    response = client.get("/download/测试_中文文档.docx")
    assert response.status_code == 200
    assert response.content == b"dummy docx content"


def test_download_missing_file_returns_404():
    response = client.get("/download/not-here.docx")
    assert response.status_code == 404


def test_download_rejects_path_traversal():
    response = client.get("/download/..%2F..%2Fconfig.py")
    assert response.status_code == 404


def test_interrupted_jobs_are_failed_on_reload():
    """Jobs live in-process only: a restart must not leave records stuck forever."""
    import json

    main.GENERATION_RECORDS_FILE.write_text(
        json.dumps(
            [
                {
                    "id": "stuck",
                    "prompt": "p",
                    "status": "processing",
                    "message": "后台正在生成 Word",
                    "error": "",
                    "created_at": "2026-01-01T00:00:00",
                    "updated_at": "2026-01-01T00:00:00",
                },
                {
                    "id": "done",
                    "prompt": "p",
                    "status": "completed",
                    "message": "Word 生成完成",
                    "error": "",
                    "created_at": "2026-01-01T00:00:00",
                    "updated_at": "2026-01-01T00:00:00",
                },
            ],
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )

    main.load_generation_records()

    records = {record["id"]: record for record in main.list_generation_records(10)}
    assert records["stuck"]["status"] == "failed"
    assert records["stuck"]["error"]
    assert records["done"]["status"] == "completed"


def test_corrupt_records_file_is_ignored():
    main.GENERATION_RECORDS_FILE.write_text("{ not json", encoding="utf-8")

    main.load_generation_records()

    assert main.list_generation_records(10) == []


if __name__ == "__main__":
    pytest.main([__file__])
