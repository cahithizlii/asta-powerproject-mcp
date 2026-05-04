"""Test FastMCP msproject_file dispatcher (T74)."""
import asyncio
import json
import os
import shutil
import sys

import pytest

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from msproject_mcp_core import msproject_file  # noqa: E402

MSP_XML = os.path.join(os.path.dirname(__file__), "fixtures", "sample_msp.xml")


def _run(coro):
    return asyncio.run(coro)


def _call(action, **params):
    raw = _run(msproject_file({"action": action, **params}))
    return json.loads(raw)


def test_dispatcher_read_tasks():
    p = _call("read_tasks", file_path=MSP_XML)
    assert p["status"] == "ok"
    assert "tasks" in p


def test_dispatcher_read_links():
    p = _call("read_links", file_path=MSP_XML)
    assert p["status"] == "ok"
    assert "links" in p


def test_dispatcher_read_resources():
    p = _call("read_resources", file_path=MSP_XML)
    assert p["status"] == "ok"
    assert "resources" in p


def test_dispatcher_read_assignments():
    p = _call("read_assignments", file_path=MSP_XML)
    assert p["status"] == "ok"
    assert "assignments" in p


def test_dispatcher_read_calendars():
    p = _call("read_calendars", file_path=MSP_XML)
    assert p["status"] == "ok"


def test_dispatcher_read_baselines():
    p = _call("read_baselines", file_path=MSP_XML)
    assert p["status"] == "ok"


def test_dispatcher_read_progress():
    p = _call("read_progress", file_path=MSP_XML)
    assert p["status"] == "ok"


def test_dispatcher_query():
    p = _call("query", file_path=MSP_XML, expression="duration_h > 8")
    assert p["status"] == "ok"
    assert p["count"] >= 1


def test_dispatcher_save_as(tmp_path):
    dst = tmp_path / "renamed.xml"
    p = _call("save_as", file_path=MSP_XML, output_path=str(dst))
    assert p["status"] == "ok"
    assert os.path.exists(str(dst))


def test_dispatcher_write_chain_add_tasks(tmp_path):
    """add_tasks via dispatcher round-trips."""
    dst = tmp_path / "writable.xml"
    shutil.copy(MSP_XML, str(dst))
    p = _call("add_tasks", file_path=str(dst),
              items=[{"name": "DispatcherT", "duration": "1d"}])
    assert p["status"] == "ok"
    assert p["count"] == 1


def test_dispatcher_unknown_action():
    p = _call("nonsense", file_path=MSP_XML)
    assert p["status"] == "error"
    assert "Unknown action" in p["error"]


def test_dispatcher_missing_file_path():
    """No file_path → error (TypeError caught and translated)."""
    p = _call("read_tasks")
    assert p["status"] == "error"
    assert "file_path" in p["error"].lower()


# =============================================================================
# Phase 11.2 — Edge Case + Negative Path tests (T142)
# =============================================================================
# Each test asserts {status: 'error', error: <key substring>} for an invalid
# input path. No production code change required.


def test_dispatcher_read_tasks_nonexistent_file_returns_error(tmp_path):
    """Path that doesn't exist on disk → error with 'not found' substring."""
    p = _call("read_tasks", file_path=str(tmp_path / "missing.xml"))
    assert p["status"] == "error"
    assert "not found" in p["error"].lower()


def test_dispatcher_read_tasks_unsupported_extension_returns_error(tmp_path):
    """Unsupported extension (.txt) → error mentioning ext or unsupported."""
    bad = tmp_path / "notaproject.txt"
    bad.write_text("hello")
    p = _call("read_tasks", file_path=str(bad))
    assert p["status"] == "error"
    assert "unsupported" in p["error"].lower() or ".txt" in p["error"]


def test_dispatcher_read_tasks_pp_extension_rejected(tmp_path):
    """.pp file (Asta-only) is unsupported in MS Project file MCP."""
    bad = tmp_path / "asta.pp"
    bad.write_text("not a real pp")
    p = _call("read_tasks", file_path=str(bad))
    assert p["status"] == "error"
    assert "unsupported" in p["error"].lower() or ".pp" in p["error"]


def test_dispatcher_corrupted_xml_returns_error(tmp_path):
    """Garbage XML content → error (not crash)."""
    bad = tmp_path / "garbage.xml"
    bad.write_bytes(b"\x00\x01garbage<not xml")
    p = _call("read_tasks", file_path=str(bad))
    assert p["status"] == "error"
    # parser rejects unknown XML → message names the schema mismatch
    err = p["error"].lower()
    assert "schema" in err or "ms project" in err or "asta" in err


def test_dispatcher_save_as_missing_extension_returns_error(tmp_path):
    """output_path without .xml/.mspdi extension → error."""
    p = _call("save_as", file_path=MSP_XML, output_path=str(tmp_path / "noext"))
    assert p["status"] == "error"
    assert "xml" in p["error"].lower() or "mspdi" in p["error"].lower()


def test_dispatcher_save_as_wrong_extension_returns_error(tmp_path):
    """output_path with .pdf extension → error."""
    p = _call("save_as", file_path=MSP_XML, output_path=str(tmp_path / "x.pdf"))
    assert p["status"] == "error"
    assert "xml" in p["error"].lower() or "mspdi" in p["error"].lower()


def test_dispatcher_write_baseline_missing_data_returns_error():
    """write_baseline without baseline_data → error."""
    p = _call("write_baseline", file_path=MSP_XML,
              output_path="/tmp/out.xml", baseline_data=[])
    assert p["status"] == "error"
    assert "baseline_data" in p["error"].lower()


def test_dispatcher_write_baseline_missing_file_path_returns_error():
    """write_baseline without file_path → error."""
    p = _call("write_baseline", output_path="/tmp/out.xml",
              baseline_data=[{"task_uid": 1}])
    assert p["status"] == "error"
    assert "file_path" in p["error"].lower()


def test_dispatcher_write_baseline_missing_output_path_returns_error():
    """write_baseline without output_path → error."""
    p = _call("write_baseline", file_path=MSP_XML,
              baseline_data=[{"task_uid": 1}])
    assert p["status"] == "error"
    assert "output_path" in p["error"].lower()


def test_dispatcher_write_baseline_xer_source_rejected(tmp_path):
    """write_baseline with .xer source extension → error."""
    bad = tmp_path / "src.xer"
    bad.write_text("hello")  # content irrelevant — extension rejected first
    p = _call("write_baseline", file_path=str(bad),
              output_path=str(tmp_path / "out.xml"),
              baseline_data=[{"task_uid": 1}])
    assert p["status"] == "error"
    assert ".xml" in p["error"].lower() or ".mspdi" in p["error"].lower() or ".xer" in p["error"].lower()
