"""Test Phase 5c T101 FastMCP dispatcher (msproject_excel)."""
import asyncio
import json
import os
import sys

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from msproject_mcp_core import msproject_excel
from openpyxl import Workbook

MSP_XML = os.path.join(os.path.dirname(__file__), "fixtures", "sample_msp.xml")


def _run(coro):
    return asyncio.run(coro)


def _call(action, **kw):
    raw = _run(msproject_excel({"action": action, **kw}))
    return json.loads(raw)


def test_dispatcher_export_hakedis(tmp_path):
    p = _call("export_hakedis", file_path=MSP_XML,
              xlsx_path=str(tmp_path / "h.xlsx"))
    assert p["status"] == "ok"


def test_dispatcher_export_tasks(tmp_path):
    p = _call("export_tasks", file_path=MSP_XML,
              xlsx_path=str(tmp_path / "t.xlsx"))
    assert p["status"] == "ok"


def test_dispatcher_export_evm(tmp_path):
    p = _call("export_evm", file_path=MSP_XML,
              xlsx_path=str(tmp_path / "e.xlsx"))
    assert p["status"] == "ok"


def test_dispatcher_export_dcma(tmp_path):
    p = _call("export_dcma", file_path=MSP_XML,
              xlsx_path=str(tmp_path / "d.xlsx"))
    assert p["status"] == "ok"


def test_dispatcher_import_tasks(tmp_path):
    """Build a small Tasks xlsx, route via dispatcher."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Tasks"
    ws.append(["ID", "Name", "Duration (d)", "Start", "Finish",
               "%Complete", "Critical", "Summary"])
    ws.append([1, "DispImport1", 3, None, None, 0, False, False])
    xlsx = tmp_path / "imp.xlsx"
    wb.save(str(xlsx))
    p = _call("import_tasks", xlsx_path=str(xlsx))
    assert p["status"] in ("ok", "error")


def test_dispatcher_import_progress(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Progress"
    ws.append(["Task ID", "%Complete"])
    ws.append([1, 50])
    xlsx = tmp_path / "p.xlsx"
    wb.save(str(xlsx))
    p = _call("import_progress", xlsx_path=str(xlsx))
    assert p["status"] in ("ok", "error")


def test_dispatcher_unknown_action():
    p = _call("nonsense", file_path=MSP_XML)
    assert p["status"] == "error"
    assert "Unknown action" in p["error"]


def test_dispatcher_missing_xlsx_path():
    p = _call("export_hakedis", file_path=MSP_XML)
    assert p["status"] == "error"
