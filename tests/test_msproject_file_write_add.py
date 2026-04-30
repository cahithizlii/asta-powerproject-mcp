"""Test msproject_file write actions: add_tasks + add_links + add_resources."""
import os
import shutil
import sys

import pytest

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from msproject_mcp_core import (  # noqa: E402
    _msp_file_add_tasks, _msp_file_add_links, _msp_file_add_resources,
    _msp_file_read_tasks, _msp_file_read_links, _msp_file_read_resources,
)

FIXTURE_DIR = os.path.join(os.path.dirname(__file__), "fixtures")
SOURCE_XML = os.path.join(FIXTURE_DIR, "sample_msp.xml")


@pytest.fixture
def writable_xml(tmp_path):
    dst = tmp_path / "writable.xml"
    shutil.copy(SOURCE_XML, dst)
    return str(dst)


def test_add_tasks_appends_to_xml(writable_xml):
    """Add 2 tasks -> re-read shows 5 tasks (3 base + 2 new)."""
    r = _msp_file_add_tasks(file_path=writable_xml, items=[
        {"name": "T4", "duration": "5d"},
        {"name": "T5", "duration": "2d"},
    ])
    assert r["status"] == "ok"
    assert r["count"] == 2
    r2 = _msp_file_read_tasks(file_path=writable_xml)
    assert r2["count"] == 5
    names = [t["name"] for t in r2["tasks"]]
    assert "T4" in names and "T5" in names


def test_add_links_appends(writable_xml):
    """Add T1 -> T2 FS link, count incremented."""
    r0 = _msp_file_read_links(file_path=writable_xml)
    base_count = r0["count"]
    r = _msp_file_add_links(file_path=writable_xml, items=[
        {"from_id": 1, "to_id": 2, "type": "FS", "lag": "0d"},
    ])
    assert r["status"] == "ok"
    assert r["count"] == 1
    r2 = _msp_file_read_links(file_path=writable_xml)
    assert r2["count"] == base_count + 1


def test_add_resources_appends(writable_xml):
    r = _msp_file_add_resources(file_path=writable_xml, items=[
        {"name": "R3", "type": "Work", "max_units": 1.0},
    ])
    assert r["status"] == "ok"
    assert r["count"] == 1
    r2 = _msp_file_read_resources(file_path=writable_xml)
    names = [res["name"] for res in r2["resources"]]
    assert "R3" in names


def test_add_tasks_mpp_rejected(tmp_path):
    """MPP write not supported -> clear error."""
    fake_mpp = tmp_path / "test.mpp"
    fake_mpp.write_bytes(b"\x00" * 100)
    r = _msp_file_add_tasks(file_path=str(fake_mpp),
                            items=[{"name": "X", "duration": "1d"}])
    assert r["status"] == "error"
    msg = r["error"].lower()
    assert ".mpp" in msg or "binary" in msg or "not supported" in msg


def test_add_tasks_returns_auto_imported_flag(writable_xml):
    """Result includes auto_imported key (T72 will populate; T70 placeholder)."""
    r = _msp_file_add_tasks(file_path=writable_xml,
                            items=[{"name": "T6", "duration": "1d"}])
    assert "auto_imported" in r
    assert isinstance(r["auto_imported"], bool)


def test_add_resources_invalid_file():
    r = _msp_file_add_resources(file_path="/nonexistent.xml",
                                items=[{"name": "X", "type": "Work"}])
    assert r["status"] == "error"


def test_add_links_invalid_file():
    r = _msp_file_add_links(file_path="/nonexistent.xml",
                            items=[{"from_id": 1, "to_id": 2}])
    assert r["status"] == "error"
