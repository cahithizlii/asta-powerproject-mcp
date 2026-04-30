"""Test _auto_sync_to_open_msp helper (T72) — conservative match-by-file_path semantics.

Phase 4 SAFETY: auto-sync ONLY acts on a project whose FullName matches
the modified file_path. If MSP is closed, OR open with no matching
project, the helper returns auto_imported=False without touching the
user's active project. This prevents user-state contamination from
unrelated file MCP edits.
"""
import os
import sys
from unittest.mock import patch, MagicMock

import pytest

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from msproject_mcp_core import _auto_sync_to_open_msp  # noqa: E402

FIXTURE_DIR = os.path.join(os.path.dirname(__file__), "fixtures")
MSP_XML = os.path.join(FIXTURE_DIR, "sample_msp.xml")


def test_auto_sync_invalid_file_path():
    """Missing XML path -> auto_imported=False with error."""
    r = _auto_sync_to_open_msp("/nonexistent.xml")
    assert r["auto_imported"] is False
    # Either error or msg key carries the explanation
    assert "error" in r or "msg" in r


def test_auto_sync_msp_closed_returns_not_imported():
    """When GetActiveObject fails, auto_imported=False with msg."""
    with patch("msproject_mcp_core.win32com.client.GetActiveObject",
               side_effect=Exception("MSP not running")):
        r = _auto_sync_to_open_msp(MSP_XML)
        assert r["auto_imported"] is False
        msg = (r.get("msg") or r.get("error") or "").lower()
        assert "msp" in msg or "closed" in msg or "not running" in msg


def test_auto_sync_no_matching_project_skipped():
    """MSP open with NO project matching file_path -> auto_imported=False.

    This is the SAFETY behavior — never touch a project the user didn't
    mean to update via the file MCP.
    """
    fake_app = MagicMock()
    fake_app.Projects.Count = 1
    fake_proj = MagicMock()
    fake_proj.FullName = "C:/Users/Other/project_unrelated.mpp"
    fake_proj.Name = "project_unrelated"
    fake_app.Projects.return_value = fake_proj
    fake_app.Projects.__getitem__ = lambda self, idx: fake_proj
    # Mock Projects(i) call style
    fake_app.Projects = MagicMock()
    fake_app.Projects.Count = 1
    fake_app.Projects.side_effect = lambda i: fake_proj
    fake_app.ActiveProject = fake_proj

    with patch("msproject_mcp_core.win32com.client.GetActiveObject",
               return_value=fake_app):
        r = _auto_sync_to_open_msp(MSP_XML)
    assert r["auto_imported"] is False
    msg = (r.get("msg") or r.get("error") or "").lower()
    assert "matching" in msg or "no" in msg or "open" in msg


def test_auto_sync_matching_project_reloads(tmp_path):
    """MSP open with matching project -> close + reopen + Reschedule.

    Mocks the COM dance. Verifies FileClose + FileOpen + Reschedule
    are all invoked and auto_imported=True.
    """
    test_xml = tmp_path / "synctest.xml"
    test_xml.write_bytes(b"<Project xmlns=\"http://schemas.microsoft.com/project\"/>")
    file_path = str(test_xml)

    fake_app = MagicMock()
    matching_proj = MagicMock()
    matching_proj.FullName = file_path.replace("\\", "/")
    matching_proj.Name = "synctest"
    matching_window = MagicMock()
    matching_window.Caption = "synctest"
    matching_proj.Windows.return_value = matching_window
    matching_proj.Windows.__getitem__ = lambda self, idx: matching_window
    matching_proj.Windows = MagicMock()
    matching_proj.Windows.side_effect = lambda i: matching_window

    fake_app.Projects = MagicMock()
    fake_app.Projects.Count = 1
    fake_app.Projects.side_effect = lambda i: matching_proj
    fake_app.ActiveProject = matching_proj

    with patch("msproject_mcp_core.win32com.client.GetActiveObject",
               return_value=fake_app):
        r = _auto_sync_to_open_msp(file_path)

    assert r["auto_imported"] is True
    assert r.get("reschedule_ok") is True
    fake_app.FileClose.assert_called()
    fake_app.FileOpen.assert_called()
    matching_proj.Reschedule.assert_called()
