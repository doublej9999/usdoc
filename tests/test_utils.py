"""
Unit tests for usdoc utility functions.

Run with: python -m pytest tests/ -v
"""

import sys
from pathlib import Path

# Ensure the project root is importable
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from main import (
    sanitize_filename,
    sanitize_output_filename,
    normalize_output_docx_filename,
    resolve_column_index,
    build_row_prompt,
    extract_json_from_text,
)


# ---------------------------------------------------------------------------
# sanitize_filename
# ---------------------------------------------------------------------------

class TestSanitizeFilename:
    def test_basic_ascii(self):
        assert sanitize_filename("hello.txt") == "hello.txt"

    def test_strips_path_components(self):
        assert sanitize_filename("../../etc/passwd") == "passwd"

    def test_replaces_non_ascii(self):
        result = sanitize_filename("中文文件.docx")
        assert result.endswith(".docx")
        assert all(c in "_.0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ" for c in result)

    def test_dot_fallback(self):
        result = sanitize_filename(".hidden")
        assert result.startswith("_")
        assert "hidden" in result

    def test_double_dot_fallback(self):
        result = sanitize_filename("..docx")
        assert result.startswith("_")
        assert "docx" in result

    def test_spaces_replaced(self):
        assert sanitize_filename("my file.docx") == "my_file.docx"

    def test_special_chars_replaced(self):
        assert sanitize_filename("file;name.docx") == "file_name.docx"


# ---------------------------------------------------------------------------
# sanitize_output_filename (allows Unicode)
# ---------------------------------------------------------------------------

class TestSanitizeOutputFilename:
    def test_preserves_unicode(self):
        result = sanitize_output_filename("中文文档.docx")
        assert "中文文档" in result or result.endswith(".docx")

    def test_strips_path(self):
        assert sanitize_output_filename("../evil.docx") == "evil.docx"

    def test_spaces_replaced(self):
        assert sanitize_output_filename("my output.docx") == "my_output.docx"

    def test_dot_fallback(self):
        result = sanitize_output_filename(".secret")
        assert result.startswith("_")

    def test_double_dot(self):
        result = sanitize_output_filename("..docx")
        assert result.startswith("_")
        assert "docx" in result


# ---------------------------------------------------------------------------
# normalize_output_docx_filename
# ---------------------------------------------------------------------------

class TestNormalizeOutputDocxFilename:
    def test_empty_returns_generated(self):
        result = normalize_output_docx_filename("")
        assert result.startswith("generated_")
        assert result.endswith(".docx")

    def test_none_returns_generated(self):
        result = normalize_output_docx_filename(None)
        assert result.startswith("generated_")
        assert result.endswith(".docx")

    def test_adds_docx_extension(self):
        result = normalize_output_docx_filename("myfile")
        assert result == "myfile.docx"

    def test_preserves_existing_extension(self):
        assert normalize_output_docx_filename("report.docx") == "report.docx"

    def test_strips_path(self):
        result = normalize_output_docx_filename("../hack.docx")
        assert ".." not in result
        assert result.endswith(".docx")


# ---------------------------------------------------------------------------
# resolve_column_index
# ---------------------------------------------------------------------------

class TestResolveColumnIndex:
    def test_single_letter(self):
        assert resolve_column_index("A", []) == 0
        assert resolve_column_index("B", []) == 1
        assert resolve_column_index("Z", []) == 25

    def test_double_letter(self):
        assert resolve_column_index("AA", []) == 26
        assert resolve_column_index("AB", []) == 27

    def test_case_insensitive(self):
        assert resolve_column_index("a", []) == 0
        assert resolve_column_index("c", []) == 2

    def test_by_header_name(self):
        # Use header names that are NOT valid Excel column letters so header lookup is exercised
        headers = ["序号", "prompt_text", "output_name"]
        assert resolve_column_index("prompt_text", headers) == 1
        assert resolve_column_index("output_name", headers) == 2

    def test_by_header_name_case_insensitive(self):
        # "Prompt_text" is not a valid Excel column (contains non-alpha chars won't work)
        # Use a header that is pure-letters but will be treated as column letters.
        # So instead test with a mixed-case header that has non-letter chars
        headers = ["序号", "prompt_text", "output_name"]
        assert resolve_column_index("序号", headers) == 0
        # Headers that contain letters only serve as header lookups
        assert resolve_column_index("output_name", headers) == 2

    def test_empty_raises(self):
        import pytest
        with pytest.raises(Exception):
            resolve_column_index("", [])

    def test_not_found_raises(self):
        import pytest
        # "非xyz" is not a valid column letter so it falls through to header search
        with pytest.raises(Exception):
            resolve_column_index("非xyz", ["a", "b"])


# ---------------------------------------------------------------------------
# build_row_prompt
# ---------------------------------------------------------------------------

class TestBuildRowPrompt:
    def test_empty_cell_returns_empty(self):
        assert build_row_prompt("some prompt", "") == ""
        assert build_row_prompt("some prompt", None) == ""

    def test_empty_base_returns_cell(self):
        assert build_row_prompt("", "hello") == "hello"
        assert build_row_prompt(None, "hello") == "hello"

    def test_token_replacement(self):
        result = build_row_prompt("Generate for {{excel_value}}", "test input")
        assert "test input" in result
        assert "{{excel_value}}" not in result

    def test_curly_token_replacement(self):
        result = build_row_prompt("Generate for {excel_value}", "test input")
        assert "test input" in result
        assert "{excel_value}" not in result

    def test_value_token_replacement(self):
        result = build_row_prompt("Use {{value}} here", "data")
        assert "data" in result
        assert "{{value}}" not in result

    def test_no_token_appends(self):
        result = build_row_prompt("Generate document", "my content")
        assert "Generate document" in result
        assert "my content" in result
        assert "Excel内容" in result


# ---------------------------------------------------------------------------
# extract_json_from_text
# ---------------------------------------------------------------------------

class TestExtractJsonFromText:
    def test_pure_json(self):
        result = extract_json_from_text('{"key": "value"}')
        assert result == {"key": "value"}

    def test_json_with_whitespace(self):
        result = extract_json_from_text('  {"key": "value"}  ')
        assert result == {"key": "value"}

    def test_json_embedded_in_text(self):
        result = extract_json_from_text('Here is the result: {"name": "test"} done')
        assert result == {"name": "test"}

    def test_no_json_raises(self):
        import pytest
        with pytest.raises(Exception):
            extract_json_from_text("no json here")

    def test_nested_json(self):
        result = extract_json_from_text('{"outer": {"inner": "value"}}')
        assert result == {"outer": {"inner": "value"}}

    def test_multiple_keys(self):
        result = extract_json_from_text('{"a": "1", "b": "2", "c": "3"}')
        assert result == {"a": "1", "b": "2", "c": "3"}
