# -*- coding: utf-8 -*-
import sys
from pathlib import Path

# Ensure project root is in sys.path when running standalone
PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import pytest
from utils.files import sanitize_filename, ensure_unique_docx_path


def test_sanitize_filename_chinese():
    result = sanitize_filename("US0001_学生管理.docx")
    assert result == "US0001_学生管理.docx"


def test_sanitize_filename_path_traversal():
    assert sanitize_filename("../../evil.docx") == "evil.docx"
    assert sanitize_filename("..\\..\\evil.docx") == "evil.docx"


def test_sanitize_filename_windows_invalid_chars():
    assert sanitize_filename("file:name*?.docx") == "file_name__.docx"


def test_ensure_unique_docx_path_nonexistent(tmp_path: Path):
    filename, path = ensure_unique_docx_path(tmp_path, "report.docx")
    assert filename == "report.docx"
    assert path == tmp_path / "report.docx"


def test_ensure_unique_docx_path_collision(tmp_path: Path):
    initial_file = tmp_path / "US0001_学生管理.docx"
    initial_file.touch()

    new_filename, new_path = ensure_unique_docx_path(tmp_path, "US0001_学生管理.docx")
    assert new_filename != "US0001_学生管理.docx"
    assert new_filename.startswith("US0001_学生管理_")
    assert new_filename.endswith(".docx")
    assert not new_path.exists()
    assert new_path.parent == tmp_path


if __name__ == "__main__":
    pytest.main([__file__])
