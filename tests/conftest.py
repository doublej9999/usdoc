# -*- coding: utf-8 -*-
"""Shared pytest fixtures.

Tests must never touch the real ``uploads/``, ``outputs/`` or
``generation_records.json`` of a developer/operator checkout.
"""
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import pytest


@pytest.fixture(autouse=True)
def isolated_runtime(tmp_path, monkeypatch):
    """Redirect every runtime path used by ``main`` into a per-test tmp dir."""
    import main
    from config import UPLOAD_DIR as REAL_UPLOAD_DIR

    uploads = tmp_path / "uploads"
    outputs = tmp_path / "outputs"
    uploads.mkdir()
    outputs.mkdir()

    # Copy built-in templates so they stay selectable but are never mutated.
    if REAL_UPLOAD_DIR.exists():
        for template in REAL_UPLOAD_DIR.glob("*.docx"):
            (uploads / template.name).write_bytes(template.read_bytes())

    monkeypatch.setattr(main, "UPLOAD_DIR", uploads)
    monkeypatch.setattr(main, "OUTPUT_DIR", outputs)
    monkeypatch.setattr(main, "GENERATION_RECORDS_FILE", tmp_path / "generation_records.json")

    with main.generation_records_lock:
        main.generation_records.clear()

    yield

    with main.generation_records_lock:
        main.generation_records.clear()
