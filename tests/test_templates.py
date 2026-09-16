# -*- coding: utf-8 -*-
import sys
from pathlib import Path

# Ensure project root is in sys.path when running standalone
PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import pytest
from docx import Document
from utils.templates import extract_template_variables, extract_invalid_template_placeholders


def test_extract_template_variables_on_us_docx():
    template_path = PROJECT_ROOT / "US.docx"
    assert template_path.exists(), f"Template not found at {template_path}"

    variables = extract_template_variables(template_path)
    assert isinstance(variables, list)
    assert len(variables) > 0

    expected_vars = [
        "story_id",
        "story_title",
        "story_description",
        "story_acceptance_criteria",
    ]
    for var in expected_vars:
        assert var in variables, f"Expected variable '{var}' not found in {variables}"


def test_extract_invalid_template_placeholders_clean_template():
    template_path = PROJECT_ROOT / "US.docx"
    assert template_path.exists(), f"Template not found at {template_path}"

    invalid = extract_invalid_template_placeholders(template_path)
    assert invalid == []


def test_extract_invalid_template_placeholders_with_invalid_tags(tmp_path: Path):
    doc = Document()
    doc.add_paragraph("Hello {{ valid_var }} and {{ 123invalid }} and {{ bad-identifier }}")
    temp_docx = tmp_path / "test_invalid.docx"
    doc.save(str(temp_docx))

    invalid = extract_invalid_template_placeholders(temp_docx)
    assert "123invalid" in invalid or "bad-identifier" in invalid

def test_generate_default_document(tmp_path: Path):
    from services.document_service import generate_default_document
    from config import UPLOAD_DIR

    output_name = "test_default_gen.docx"
    generated_file = generate_default_document(
        template_name="US.docx",
        output_filename=output_name,
        upload_dir=UPLOAD_DIR,
        output_dir=tmp_path,
    )
    assert generated_file == output_name
    assert (tmp_path / output_name).exists()



if __name__ == "__main__":
    pytest.main([__file__])
