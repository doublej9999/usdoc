# -*- coding: utf-8 -*-
import pytest

from utils.files import ensure_unique_docx_path, sanitize_filename


def test_sanitize_filename_keeps_unicode_word_characters():
    assert sanitize_filename("US0001_学生管理.docx") == "US0001_学生管理.docx"


def test_sanitize_filename_strips_directory_components():
    assert sanitize_filename("../../evil.docx") == "evil.docx"
    assert sanitize_filename("..\\..\\evil.docx") == "evil.docx"
    assert sanitize_filename("C:/windows/system32/cmd.exe") == "cmd.exe"


def test_sanitize_filename_replaces_windows_illegal_characters():
    assert sanitize_filename("file:name*?.docx") == "file_name__.docx"
    assert sanitize_filename("a\x00b.docx") == "ab.docx"


def test_sanitize_filename_rejects_directory_only_names():
    assert sanitize_filename("..") == ""
    assert sanitize_filename(".") == ""
    assert sanitize_filename("") == ""


def test_ensure_unique_docx_path_free_name(tmp_path):
    filename, path = ensure_unique_docx_path(tmp_path, "report.docx")
    assert filename == "report.docx"
    assert path == tmp_path / "report.docx"


def test_ensure_unique_docx_path_collision(tmp_path):
    (tmp_path / "US0001_学生管理.docx").touch()

    filename, path = ensure_unique_docx_path(tmp_path, "US0001_学生管理.docx")
    assert filename != "US0001_学生管理.docx"
    assert filename.startswith("US0001_学生管理_")
    assert filename.endswith(".docx")
    assert path.parent == tmp_path


def test_ensure_unique_docx_path_reservation_is_atomic(tmp_path):
    """Concurrent callers must never receive the same path (guards batch jobs)."""
    import threading

    threads_count = 12
    barrier = threading.Barrier(threads_count)
    results = [None] * threads_count
    guard = threading.Lock()

    def reserve(index: int) -> None:
        barrier.wait()  # make every thread check availability at the same instant
        name, path = ensure_unique_docx_path(tmp_path, "dup.docx")
        with guard:
            results[index] = (name, path)

    workers = [threading.Thread(target=reserve, args=(i,)) for i in range(threads_count)]
    for worker in workers:
        worker.start()
    for worker in workers:
        worker.join()

    names = [name for name, _ in results]
    assert len(set(names)) == threads_count


if __name__ == "__main__":
    pytest.main([__file__])
