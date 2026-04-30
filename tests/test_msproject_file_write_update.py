"""Test msproject_file write actions: update_task + save_as."""
import os
import shutil
import sys

import pytest

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from msproject_mcp_core import (  # noqa: E402
    _msp_file_update_task, _msp_file_save_as, _msp_file_read_tasks,
)

FIXTURE_DIR = os.path.join(os.path.dirname(__file__), "fixtures")
SOURCE_XML = os.path.join(FIXTURE_DIR, "sample_msp.xml")


@pytest.fixture
def writable_xml(tmp_path):
    dst = tmp_path / "writable.xml"
    shutil.copy(SOURCE_XML, dst)
    return str(dst)


def test_update_task_duration(writable_xml):
    """Update T1 (id=1) duration 1d -> 5d, re-read confirms 40h."""
    r = _msp_file_update_task(file_path=writable_xml, task_id=1,
                              fields={"duration": "5d"})
    assert r["status"] == "ok"
    r2 = _msp_file_read_tasks(file_path=writable_xml)
    t1 = next(t for t in r2["tasks"] if t["id"] == 1)
    assert t1["duration_h"] == 40.0


def test_update_task_missing_id(writable_xml):
    r = _msp_file_update_task(file_path=writable_xml, task_id=99999,
                              fields={"duration": "1d"})
    assert r["status"] == "error"
    assert "task" in r["error"].lower()


def test_update_task_invalid_file():
    r = _msp_file_update_task(file_path="/nonexistent.xml", task_id=1,
                              fields={"duration": "1d"})
    assert r["status"] == "error"


def test_update_task_mpp_rejected(tmp_path):
    fake_mpp = tmp_path / "test.mpp"
    fake_mpp.write_bytes(b"\x00" * 100)
    r = _msp_file_update_task(file_path=str(fake_mpp), task_id=1,
                              fields={"duration": "5d"})
    assert r["status"] == "error"
    msg = r["error"].lower()
    assert ".mpp" in msg or "binary" in msg or "not supported" in msg


def test_save_as_xml(tmp_path):
    """Save fixture to a new path; output file exists and has size."""
    dst = tmp_path / "renamed.xml"
    r = _msp_file_save_as(file_path=SOURCE_XML, output_path=str(dst))
    assert r["status"] == "ok"
    assert os.path.exists(str(dst))
    assert r["output_path"] == str(dst)
    assert r["size_bytes"] > 0


def test_save_as_invalid_output_extension(tmp_path):
    """output_path must end .xml or .mspdi."""
    dst = tmp_path / "bad.txt"
    r = _msp_file_save_as(file_path=SOURCE_XML, output_path=str(dst))
    assert r["status"] == "error"
    assert "extension" in r["error"].lower() or "xml" in r["error"].lower()


def test_save_as_invalid_input_file():
    r = _msp_file_save_as(file_path="/nonexistent.xml",
                          output_path="/tmp/out.xml")
    assert r["status"] == "error"
