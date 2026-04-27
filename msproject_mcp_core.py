"""MS Project MCP Server - COM-based.

Hybrid speed strategy:
  1-5 items   -> COM direct (real-time, instant UI feedback)
  6-19 items  -> COM batch (Calculation manual + ScreenUpdating off)
  20+ items   -> MSPDI XML bulk import (~3-5s for 200 tasks)

Phase 1 tools: msproject_task, msproject_link, msproject_schedule.
"""
from __future__ import annotations
import atexit
import logging
import os
import threading
import time
from typing import Any, Dict, List, Optional

import pythoncom
import win32com.client
from mcp.server.fastmcp import FastMCP

# ---------- LOGGING ----------
log_dir = os.path.expanduser("~/.claude/logs")
os.makedirs(log_dir, exist_ok=True)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    handlers=[
        logging.FileHandler(os.path.join(log_dir, "msproject_mcp.log"), encoding="utf-8"),
        logging.StreamHandler(),
    ],
)
logger = logging.getLogger("msproject_mcp")

# ---------- MCP SERVER ----------
mcp = FastMCP(
    "msproject_mcp",
    instructions=(
        "MS Project COM-based MCP server. Connects to running MS Project (Application='MSProject.Application'). "
        "Hybrid speed: 1-5 items COM direct, 6-19 batch, 20+ MSPDI bulk import. "
        "Tools: msproject_task, msproject_link, msproject_schedule (Phase 1)."
    ),
)

# ---------- COM CONNECTION CACHE ----------
_app_lock = threading.RLock()
_app: Optional[Any] = None
_calc_modified = False  # track if we changed calc mode (need restore on exit)
_screenupdating_modified = False


def _connect_app() -> Any:
    """Connect to running MS Project. Cached singleton."""
    global _app
    with _app_lock:
        if _app is not None:
            try:
                _ = _app.Version  # ping
                return _app
            except Exception:
                _app = None  # invalidate
        pythoncom.CoInitialize()
        try:
            _app = win32com.client.GetActiveObject("MSProject.Application")
            logger.info(f"Connected to MS Project {_app.Version}")
        except Exception as e:
            raise RuntimeError(
                f"MS Project'e bağlanılamadı: {e}. "
                "(1) MS Project açık olduğundan emin olun. "
                "(2) Bir proje açık olmalı (boş Project bile yeterli). "
                "(3) Hala olmuyorsa MS Project'i yeniden başlatın."
            )
        return _app


def _validate_active_project() -> Any:
    """Validates ActiveProject is present."""
    app = _connect_app()
    if app.ActiveProject is None:
        raise RuntimeError("MS Project'te aktif proje yok. Boş bir proje açın veya File → New.")
    return app


def _route_operation(op_count: int) -> str:
    """Pick speed path based on operation count."""
    if op_count <= 5:
        return "com_direct"
    if op_count <= 19:
        return "com_batch"
    return "mspdi_bulk"


def _enter_batch_mode():
    """Enter COM batch mode: disable screen update, manual calc, no events."""
    global _calc_modified, _screenupdating_modified
    app = _connect_app()
    pj_manual = 0  # PjCalculation.pjManual
    if app.Calculation != pj_manual:
        app.Calculation = pj_manual
        _calc_modified = True
    if app.ScreenUpdating:
        app.ScreenUpdating = False
        _screenupdating_modified = True
    proj = app.ActiveProject
    try:
        proj.EventsEnabled = False
    except Exception:
        pass


def _exit_batch_mode():
    """Restore screen update + auto calc + events."""
    global _calc_modified, _screenupdating_modified
    try:
        app = _connect_app()
        if _calc_modified:
            pj_auto = 1  # PjCalculation.pjAutomatic
            app.Calculation = pj_auto
            _calc_modified = False
        if _screenupdating_modified:
            app.ScreenUpdating = True
            _screenupdating_modified = False
        proj = app.ActiveProject
        if proj:
            try:
                proj.EventsEnabled = True
            except Exception:
                pass
    except Exception as e:
        logger.warning(f"_exit_batch_mode error (non-fatal): {e}")


@atexit.register
def _restore_on_exit():
    """Critical: ensure MS Project never left in manual/screen-off state."""
    _exit_batch_mode()


# ---------- TOOL DISPATCHERS (filled in T6+) ----------
# (Placeholder - actual @mcp.tool functions added in T14)


def main():
    """Run MCP server (stdio)."""
    mcp.run()


if __name__ == "__main__":
    main()
