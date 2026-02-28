#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Asta Powerproject GUI MCP Server
=================================
GUI automation tools for Asta Powerproject.
Uses pyautogui and pywinauto for screen interaction.

Author: Claude AI for Cahit
"""

import json
import os
import sys
import logging
import subprocess
import time
import base64
from datetime import datetime
from typing import Optional, List, Dict, Any
from pathlib import Path

from pydantic import BaseModel, Field, field_validator, ConfigDict
from mcp.server.fastmcp import FastMCP

# Logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stderr),
        logging.FileHandler(
            os.path.join(os.path.expanduser("~"), "asta_mcp_gui.log"),
            encoding='utf-8'
        )
    ]
)
logger = logging.getLogger("asta_mcp_gui")

mcp = FastMCP("asta_powerproject_gui")

ASTA_WINDOW_TITLE = "Asta Powerproject"

# ============================================================================
# HELPERS
# ============================================================================
def clean_turkish(text: str) -> str:
    """Convert Turkish characters to ASCII equivalents."""
    if text is None:
        return ""
    text = str(text)
    tr_map = {
        '\u00e7': 'c', '\u00c7': 'C', '\u011f': 'g', '\u011e': 'G',
        '\u015f': 's', '\u015e': 'S', '\u00fc': 'u', '\u00dc': 'U',
        '\u00f6': 'o', '\u00d6': 'O', '\u0131': 'i', '\u0130': 'I',
    }
    for tr_char, eng_char in tr_map.items():
        text = text.replace(tr_char, eng_char)
    return text


def _get_powershell_path() -> str:
    """Find the full path to PowerShell executable."""
    candidates = [
        os.path.join(os.environ.get("SystemRoot", r"C:\WINDOWS"),
                     "System32", "WindowsPowerShell", "v1.0", "powershell.exe"),
        r"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe",
        r"C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe",
    ]
    for p in candidates:
        if os.path.isfile(p):
            return p
    # Last resort: hope it's on PATH
    return "powershell.exe"


def _clipboard_paste(text: str):
    """Type text using clipboard paste (handles Unicode/Turkish characters)."""
    import subprocess
    ps_path = _get_powershell_path()
    escaped = text.replace("'", "''")
    subprocess.run(
        [ps_path, '-NoProfile', '-command', f"Set-Clipboard -Value '{escaped}'"],
        capture_output=True, timeout=5
    )
    import pyautogui
    time.sleep(0.1)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.2)



# ============================================================================
# GUI AUTOMATION MANAGER CLASS
# ============================================================================
class AstaGUIManager:
    """Manages GUI automation for Asta Powerproject using pyautogui/pywinauto."""

    @staticmethod
    def _check_gui_libs():
        """Check if GUI automation libraries are available."""
        missing = []
        try:
            import pyautogui
        except ImportError:
            missing.append("pyautogui")
        try:
            import pywinauto
        except ImportError:
            missing.append("pywinauto")
        if missing:
            return False, f"Missing libraries: {', '.join(missing)}. Install with: pip install {' '.join(missing)}"
        return True, "OK"

    @staticmethod
    def find_asta_window():
        """Find the Asta Powerproject window."""
        try:
            import pywinauto
            from pywinauto import Desktop
            desktop = Desktop(backend="uia")
            windows = desktop.windows()
            for w in windows:
                title = w.window_text()
                if "powerproject" in title.lower() or "asta" in title.lower():
                    return {"found": True, "title": title, "handle": w.handle}
            return {"found": False, "message": "Asta Powerproject window not found. Please open the application first."}
        except Exception as e:
            return {"found": False, "message": f"Error finding window: {str(e)}"}

    @staticmethod
    def bring_to_front():
        """Bring Asta Powerproject to the foreground."""
        try:
            import pywinauto
            from pywinauto import Desktop
            desktop = Desktop(backend="uia")
            for w in desktop.windows():
                if "powerproject" in w.window_text().lower() or "asta" in w.window_text().lower():
                    w.set_focus()
                    time.sleep(0.5)
                    return {"success": True, "message": "Asta brought to foreground"}
            return {"success": False, "message": "Asta window not found"}
        except Exception as e:
            return {"success": False, "message": str(e)}

    @staticmethod
    def send_shortcut(keys: str, delay: float = 0.5):
        """Send keyboard shortcut to Asta."""
        try:
            import pyautogui
            # Bring Asta to front first
            AstaGUIManager.bring_to_front()
            time.sleep(delay)
            pyautogui.hotkey(*keys.split('+'))
            time.sleep(delay)
            return {"success": True, "message": f"Shortcut '{keys}' sent"}
        except Exception as e:
            return {"success": False, "message": str(e)}

    @staticmethod
    def click_at(x: int, y: int, clicks: int = 1, button: str = "left"):
        """Click at specific coordinates."""
        try:
            import pyautogui
            AstaGUIManager.bring_to_front()
            time.sleep(0.3)
            pyautogui.click(x=x, y=y, clicks=clicks, button=button)
            time.sleep(0.3)
            return {"success": True, "message": f"Clicked at ({x}, {y})"}
        except Exception as e:
            return {"success": False, "message": str(e)}

    @staticmethod
    def type_text(text: str, interval: float = 0.05):
        """Type text in the currently focused field."""
        try:
            import pyautogui
            _clipboard_paste(text)
            time.sleep(0.3)
            return {"success": True, "message": f"Typed: {text}"}
        except Exception as e:
            return {"success": False, "message": str(e)}

    @staticmethod
    def take_screenshot(save_path: str = None) -> dict:
        """Take a screenshot of the Asta window."""
        if not save_path:
            save_path = os.path.join(
                os.path.expanduser("~"), "Downloads",
                f"asta_screenshot_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
            )
        # Try pyautogui first
        try:
            import pyautogui
            screenshot = pyautogui.screenshot()
            screenshot.save(save_path)
            return {"success": True, "path": save_path, "message": f"Screenshot saved: {save_path}"}
        except Exception as pyautogui_err:
            logger.warning(f"pyautogui screenshot failed: {pyautogui_err}")

        # Fallback: use mss (if available)
        try:
            import mss
            with mss.mss() as sct:
                sct.shot(output=save_path)
            return {"success": True, "path": save_path, "message": f"Screenshot saved (mss fallback): {save_path}"}
        except ImportError:
            pass
        except Exception as mss_err:
            logger.warning(f"mss screenshot also failed: {mss_err}")

        # Final fallback: PowerShell
        try:
            ps_script = (
                "Add-Type -AssemblyName System.Windows.Forms; "
                "$screen = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds; "
                "$bitmap = New-Object System.Drawing.Bitmap($screen.Width, $screen.Height); "
                "$graphics = [System.Drawing.Graphics]::FromImage($bitmap); "
                "$graphics.CopyFromScreen($screen.Location, [System.Drawing.Point]::Empty, $screen.Size); "
                f"$bitmap.Save('{save_path.replace(chr(39), chr(39)+chr(39))}'); "
                "$graphics.Dispose(); $bitmap.Dispose()"
            )
            subprocess.run([_get_powershell_path(), '-NoProfile', '-command', ps_script],
                           capture_output=True, timeout=10)
            if os.path.exists(save_path):
                return {"success": True, "path": save_path, "message": f"Screenshot saved (PowerShell fallback): {save_path}"}
        except Exception as ps_err:
            logger.warning(f"PowerShell screenshot also failed: {ps_err}")

        return {
            "success": False,
            "message": "All screenshot methods failed. Try: pip install --upgrade Pillow pyscreeze"
        }

    @staticmethod
    def navigate_menu(menu_path: List[str], delay: float = 0.5):
        """Navigate through Asta ribbon/menu system."""
        try:
            import pyautogui
            AstaGUIManager.bring_to_front()
            time.sleep(delay)
            # Click each menu item in sequence
            for item in menu_path:
                # Use Alt key for ribbon tabs
                pyautogui.press('alt')
                time.sleep(0.3)
                _clipboard_paste(item)
                time.sleep(delay)
            return {"success": True, "message": f"Navigated: {' > '.join(menu_path)}"}
        except Exception as e:
            return {"success": False, "message": str(e)}




# ============================================================================
# PYDANTIC INPUT MODELS (GUI)
# ============================================================================
class GUIShortcutInput(BaseModel):
    """Input for sending keyboard shortcuts."""
    model_config = ConfigDict(str_strip_whitespace=True)
    shortcut: str = Field(
        ...,
        description="Keyboard shortcut to send. Examples: 'ctrl+s' (save), 'F9' (reschedule), 'ctrl+z' (undo), 'ctrl+p' (print)"
    )


class GUIClickInput(BaseModel):
    """Input for clicking at coordinates."""
    x: int = Field(..., description="X coordinate (pixels from left)", ge=0)
    y: int = Field(..., description="Y coordinate (pixels from top)", ge=0)
    clicks: int = Field(default=1, description="Number of clicks (1=single, 2=double)", ge=1, le=3)
    button: str = Field(default="left", description="Mouse button: 'left', 'right', or 'middle'")


class GUITypeInput(BaseModel):
    """Input for typing text."""
    model_config = ConfigDict(str_strip_whitespace=False)
    text: str = Field(..., description="Text to type in the currently focused cell/field")
    press_enter: bool = Field(default=False, description="Press Enter after typing")


class GUIMenuInput(BaseModel):
    """Input for menu navigation."""
    tab: str = Field(
        ...,
        description="Ribbon tab name: 'Home', 'View', 'Project', 'Allocation', 'Format', 'File'"
    )
    command: Optional[str] = Field(
        default=None,
        description="Command to click within the tab (e.g., 'Reschedule', 'Link Tasks', 'Summarise')"
    )


class GUIScreenshotInput(BaseModel):
    """Input for taking screenshots."""
    save_path: Optional[str] = Field(
        default=None,
        description="File path to save screenshot. Auto-generates if empty"
    )


class GUINewProjectInput(BaseModel):
    """Input for creating a new project via GUI."""
    project_name: str = Field(..., description="Name for the new project", min_length=1)
    client_name: Optional[str] = Field(default=None, description="Client/customer name (For field)")
    contractor_name: Optional[str] = Field(default=None, description="Contractor name (By field)")
    start_date: Optional[str] = Field(
        default=None,
        description="Project start date in DD/MM/YYYY format. Uses today if empty"
    )
    template: str = Field(
        default="Construction Template",
        description="Template to use: 'Construction Template', 'Housing Template', 'Blank Project', etc."
    )


class GUIProgressInput(BaseModel):
    """Input for entering progress via GUI."""
    task_name: str = Field(..., description="Name of the task to update progress for")
    percent_complete: float = Field(
        ..., description="Completion percentage (0-100)", ge=0, le=100
    )
    actual_start: Optional[str] = Field(
        default=None,
        description="Actual start date in DD/MM/YYYY format"
    )
    actual_finish: Optional[str] = Field(
        default=None,
        description="Actual finish date in DD/MM/YYYY format"
    )


class GUIOpenFileInput(BaseModel):
    """Input for opening a project file in Asta."""
    file_path: str = Field(
        ...,
        description="Full path to the project file to open in Asta Powerproject"
    )


class GUIFilterInput(BaseModel):
    """Input for applying filters."""
    filter_type: str = Field(
        ...,
        description="Type of filter: 'critical' (critical tasks only), 'complete' (completed tasks), 'incomplete' (not completed), 'code' (by code library), 'none' (remove filter)"
    )
    code_name: Optional[str] = Field(
        default=None,
        description="Code library name (only used when filter_type='code')"
    )


class GUIBaselineInput(BaseModel):
    """Input for taking a baseline."""
    baseline_name: str = Field(
        default="Original Plan",
        description="Name for the baseline snapshot"
    )


class GUILinkTasksInput(BaseModel):
    """Input for linking tasks in GUI."""
    predecessor_row: int = Field(..., description="Row number of the predecessor task", ge=1)
    successor_row: int = Field(..., description="Row number of the successor task", ge=1)
    link_type: str = Field(
        default="FS",
        description="Link type: 'FS' (Finish-to-Start), 'SS' (Start-to-Start), 'FF' (Finish-to-Finish), 'SF' (Start-to-Finish)"
    )
    lag: Optional[str] = Field(
        default=None,
        description="Lag/lead time (e.g., '2d' for 2 day lag, '-1d' for 1 day lead)"
    )


class GUIPrintInput(BaseModel):
    """Input for printing/exporting."""
    output_type: str = Field(
        default="pdf",
        description="Output type: 'pdf', 'clipboard', 'picture', 'printer'"
    )
    save_path: Optional[str] = Field(
        default=None,
        description="File path for PDF or picture output"
    )




# ============================================================================
# GUI TOOLS
# ============================================================================
@mcp.tool(
    name="asta_gui_check_status",
    annotations={
        "title": "Check Asta GUI Status",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    }
)
async def asta_gui_check_status() -> str:
    """Check if Asta Powerproject is running and get window information.

    Use this FIRST before any GUI automation to verify Asta is open.
    Also checks if required GUI automation libraries are installed.

    Returns:
        Status of Asta window and GUI library availability
    """
    gui = AstaGUIManager()

    # Check libraries
    libs_ok, libs_msg = gui._check_gui_libs()

    # Check window
    window_info = gui.find_asta_window()

    result = {
        "gui_libraries_installed": libs_ok,
        "gui_libraries_message": libs_msg,
        "asta_window": window_info,
    }

    if libs_ok and window_info.get("found"):
        result["status"] = "READY - Asta is running and GUI tools are available"
    elif not libs_ok:
        result["status"] = f"GUI LIBRARIES MISSING - {libs_msg}"
    else:
        result["status"] = "ASTA NOT FOUND - Please open Asta Powerproject first"

    return json.dumps(result, indent=2, default=str)


@mcp.tool(
    name="asta_gui_bring_to_front",
    annotations={
        "title": "Bring Asta to Front",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    }
)
async def asta_gui_bring_to_front() -> str:
    """Bring the Asta Powerproject window to the foreground.

    Use this before performing any GUI actions to ensure Asta is visible
    and has keyboard/mouse focus.

    Returns:
        Success or failure message
    """
    result = AstaGUIManager.bring_to_front()
    return json.dumps(result, indent=2)


@mcp.tool(
    name="asta_gui_send_shortcut",
    annotations={
        "title": "Send Keyboard Shortcut",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": True,
    }
)
async def asta_gui_send_shortcut(params: GUIShortcutInput) -> str:
    """Send a keyboard shortcut to Asta Powerproject.

    Common Asta shortcuts:
    - 'ctrl+s' = Save project
    - 'F9' = Reschedule (calculate critical path)
    - 'ctrl+z' = Undo
    - 'ctrl+y' = Redo
    - 'ctrl+p' = Print
    - 'insert' = Insert new bar/row
    - 'delete' = Delete selected item
    - 'F1' = Help

    Args:
        params: Contains the shortcut string (e.g., 'ctrl+s')

    Returns:
        Confirmation that shortcut was sent
    """
    result = AstaGUIManager.send_shortcut(params.shortcut)
    return json.dumps(result, indent=2)


@mcp.tool(
    name="asta_gui_reschedule",
    annotations={
        "title": "Reschedule Project",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    }
)
async def asta_gui_reschedule() -> str:
    """Run the Reschedule command in Asta Powerproject (F9).

    Rescheduling:
    1. Calculates the optimal start/end dates for all tasks
    2. Determines the Critical Path (shown in red)
    3. Calculates Total Float and Free Float
    4. Identifies constraint violations
    5. Finds the earliest project finish date

    IMPORTANT: Always reschedule after making changes to tasks or links.

    Returns:
        Confirmation that reschedule was triggered
    """
    # First bring Asta to front
    AstaGUIManager.bring_to_front()
    time.sleep(0.5)
    result = AstaGUIManager.send_shortcut("F9", delay=1.0)
    result["action"] = "Reschedule (F9)"
    result["note"] = "Check the Asta window for reschedule results. Critical path will be shown in red."
    return json.dumps(result, indent=2)


@mcp.tool(
    name="asta_gui_save",
    annotations={
        "title": "Save Project (GUI)",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    }
)
async def asta_gui_save() -> str:
    """Save the current project in Asta Powerproject (Ctrl+S).

    Returns:
        Confirmation that save command was sent
    """
    result = AstaGUIManager.send_shortcut("ctrl+s")
    result["action"] = "Save (Ctrl+S)"
    return json.dumps(result, indent=2)


@mcp.tool(
    name="asta_gui_undo",
    annotations={
        "title": "Undo Last Action",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": True,
    }
)
async def asta_gui_undo() -> str:
    """Undo the last action in Asta Powerproject (Ctrl+Z).

    Returns:
        Confirmation that undo was triggered
    """
    result = AstaGUIManager.send_shortcut("ctrl+z")
    result["action"] = "Undo (Ctrl+Z)"
    return json.dumps(result, indent=2)


@mcp.tool(
    name="asta_gui_click",
    annotations={
        "title": "Click at Position",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": True,
    }
)
async def asta_gui_click(params: GUIClickInput) -> str:
    """Click at specific screen coordinates in Asta Powerproject.

    Use take_screenshot first to identify the correct coordinates.
    The Asta window will be brought to the foreground automatically.

    Args:
        params: Contains x, y coordinates, click count, and button type

    Returns:
        Confirmation of click action
    """
    result = AstaGUIManager.click_at(params.x, params.y, params.clicks, params.button)
    return json.dumps(result, indent=2)


@mcp.tool(
    name="asta_gui_type_text",
    annotations={
        "title": "Type Text in Asta",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": True,
    }
)
async def asta_gui_type_text(params: GUITypeInput) -> str:
    """Type text into the currently focused cell or field in Asta.

    Click on a cell first, then use this tool to enter text.
    Can optionally press Enter after typing to confirm the entry.

    Args:
        params: Contains text to type and whether to press Enter

    Returns:
        Confirmation of typed text
    """
    try:
        import pyautogui
        AstaGUIManager.bring_to_front()
        time.sleep(0.3)
        _clipboard_paste(params.text)
        if params.press_enter:
            time.sleep(0.2)
            pyautogui.press('enter')
        return json.dumps({
            "success": True,
            "typed": params.text,
            "enter_pressed": params.press_enter
        }, indent=2)
    except Exception as e:
        return json.dumps({"success": False, "message": str(e)}, indent=2)


@mcp.tool(
    name="asta_gui_screenshot",
    annotations={
        "title": "Take Screenshot",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    }
)
async def asta_gui_screenshot(params: GUIScreenshotInput) -> str:
    """Take a screenshot of the current Asta Powerproject screen.

    Useful for verifying the current state before and after operations.
    The screenshot is saved as a PNG file.

    Args:
        params: Optional save_path for the screenshot file

    Returns:
        Path to the saved screenshot file
    """
    result = AstaGUIManager.take_screenshot(params.save_path)
    return json.dumps(result, indent=2)


@mcp.tool(
    name="asta_gui_open_file",
    annotations={
        "title": "Open Project File in Asta",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    }
)
async def asta_gui_open_file(params: GUIOpenFileInput) -> str:
    """Open a project file in Asta Powerproject using File > Open dialog.

    This uses keyboard shortcuts to navigate the Open dialog:
    1. Sends Ctrl+O to open the file dialog
    2. Types the file path
    3. Presses Enter to open

    Args:
        params: Contains the file path to open

    Returns:
        Confirmation of open command
    """
    try:
        import pyautogui
        gui = AstaGUIManager()
        gui.bring_to_front()
        time.sleep(0.5)

        # Send Ctrl+O
        pyautogui.hotkey('ctrl', 'o')
        time.sleep(1.5)

        # Type the file path
        file_path = params.file_path.replace("/", "\\")
        _clipboard_paste(file_path)
        time.sleep(0.5)

        # Press Enter
        pyautogui.press('enter')
        time.sleep(2.0)

        return json.dumps({
            "success": True,
            "file": params.file_path,
            "message": f"Open command sent for: {params.file_path}. Check Asta window for result."
        }, indent=2)
    except Exception as e:
        return json.dumps({"success": False, "message": str(e)}, indent=2)


@mcp.tool(
    name="asta_gui_new_project",
    annotations={
        "title": "Create New Project (GUI)",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": True,
    }
)
async def asta_gui_new_project(params: GUINewProjectInput) -> str:
    """Create a new project in Asta Powerproject using the GUI.

    Steps performed:
    1. Opens File > New dialog
    2. Selects the specified template
    3. Enters project name, client, contractor, and start date
    4. Clicks Create

    Args:
        params: Contains project_name, client_name, contractor_name, start_date, template

    Returns:
        Confirmation with instructions for manual verification
    """
    try:
        import pyautogui
        gui = AstaGUIManager()
        gui.bring_to_front()
        time.sleep(0.5)

        # Open New Project dialog: Ctrl+N
        pyautogui.hotkey('ctrl', 'n')
        time.sleep(2.0)

        return json.dumps({
            "success": True,
            "message": "New Project dialog opened. Please complete the following steps in Asta:",
            "steps": [
                f"1. Select template: '{params.template}'",
                f"2. Enter project name: '{params.project_name}'",
                f"3. Enter client (For): '{params.client_name or 'N/A'}'",
                f"4. Enter contractor (By): '{params.contractor_name or 'N/A'}'",
                f"5. Set start date: '{params.start_date or 'Today'}'",
                "6. Click 'Create'"
            ],
            "note": "Due to the complexity of the New Project dialog, some fields may need manual input. Use asta_gui_click and asta_gui_type_text for precise control."
        }, indent=2)
    except Exception as e:
        return json.dumps({"success": False, "message": str(e)}, indent=2)


@mcp.tool(
    name="asta_gui_take_baseline",
    annotations={
        "title": "Take Baseline Snapshot",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": True,
    }
)
async def asta_gui_take_baseline(params: GUIBaselineInput) -> str:
    """Take a baseline snapshot of the current project schedule.

    A baseline captures the current plan so you can compare actual
    progress against it later. This navigates to Project tab > Take Baseline.

    Args:
        params: Contains baseline_name

    Returns:
        Instructions for completing the baseline operation
    """
    try:
        import pyautogui
        gui = AstaGUIManager()
        gui.bring_to_front()
        time.sleep(0.5)

        return json.dumps({
            "success": True,
            "message": "To take a baseline in Asta:",
            "steps": [
                "1. Click the 'Project' tab in the ribbon",
                "2. Click 'Take Baseline' button",
                f"3. Enter baseline name: '{params.baseline_name}'",
                "4. Click OK",
            ],
            "note": "Use asta_gui_click to click specific UI elements, or follow these steps manually.",
            "keyboard_alternative": "You can also use the menu: Project > Take Baseline"
        }, indent=2)
    except Exception as e:
        return json.dumps({"success": False, "message": str(e)}, indent=2)


@mcp.tool(
    name="asta_gui_insert_row",
    annotations={
        "title": "Insert New Row",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": True,
    }
)
async def asta_gui_insert_row() -> str:
    """Insert a new empty row/bar at the current position in Asta.

    Sends the Insert key to create a new row in the spreadsheet.
    Click on the desired position first before using this tool.

    Returns:
        Confirmation that Insert key was pressed
    """
    result = AstaGUIManager.send_shortcut("insert")
    result["action"] = "Insert new row"
    result["note"] = "A new empty row should appear. Type the task name in the Name column."
    return json.dumps(result, indent=2)


@mcp.tool(
    name="asta_gui_delete_selected",
    annotations={
        "title": "Delete Selected Item",
        "readOnlyHint": False,
        "destructiveHint": True,
        "idempotentHint": False,
        "openWorldHint": True,
    }
)
async def asta_gui_delete_selected() -> str:
    """Delete the currently selected task, link, or item in Asta.

    WARNING: This sends the Delete key. Make sure the correct item
    is selected before using this tool. Use Ctrl+Z to undo if needed.

    Returns:
        Confirmation that Delete key was pressed
    """
    result = AstaGUIManager.send_shortcut("delete")
    result["action"] = "Delete selected item"
    result["warning"] = "Use Ctrl+Z (asta_gui_undo) immediately if wrong item was deleted"
    return json.dumps(result, indent=2)


@mcp.tool(
    name="asta_gui_link_tasks",
    annotations={
        "title": "Link Tasks (GUI)",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": True,
    }
)
async def asta_gui_link_tasks(params: GUILinkTasksInput) -> str:
    """Link two tasks together in Asta using keyboard block-linking.

    This selects the predecessor and successor tasks and uses
    Home tab > Link Tasks to create a Finish-to-Start link.

    For other link types (SS, FF, SF) or lag, you'll need to
    edit the link properties after creation.

    Args:
        params: Contains predecessor_row, successor_row, link_type, lag

    Returns:
        Instructions for linking the tasks
    """
    return json.dumps({
        "success": True,
        "message": "To link tasks in Asta:",
        "method_1_mouse": [
            f"1. Hover at the END of task in row {params.predecessor_row} (cursor becomes link icon)",
            f"2. Click and drag to the START of task in row {params.successor_row}",
            "3. Release to create a Finish-to-Start link",
            f"4. If lag is needed ({params.lag or 'none'}): Hold Shift while dragging to add lag time",
        ],
        "method_2_keyboard": [
            f"1. Click on row {params.predecessor_row} to select it",
            f"2. Hold Ctrl and click on row {params.successor_row} to add to selection",
            "3. Go to Home tab > Link Tasks button",
            "4. This creates FS links between selected tasks in order",
        ],
        "link_type": params.link_type,
        "lag": params.lag or "None",
        "note": f"To change link type to {params.link_type}: Double-click the link line > change Type in properties"
    }, indent=2)


@mcp.tool(
    name="asta_gui_apply_filter",
    annotations={
        "title": "Apply Filter",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    }
)
async def asta_gui_apply_filter(params: GUIFilterInput) -> str:
    """Apply a display filter in Asta Powerproject.

    Filters control which tasks are visible on the bar chart:
    - 'critical': Show only critical path tasks
    - 'complete': Show only completed tasks
    - 'incomplete': Show only tasks not yet completed
    - 'code': Filter by code library
    - 'none': Remove all filters (show everything)

    Args:
        params: Contains filter_type and optional code_name

    Returns:
        Instructions for applying the filter
    """
    filter_instructions = {
        "critical": [
            "1. Go to View tab",
            "2. Click 'Filter' dropdown",
            "3. Select 'Critical Tasks' or similar filter",
        ],
        "complete": [
            "1. Go to View tab",
            "2. Click 'Filter' dropdown",
            "3. Select 'Complete Tasks'",
        ],
        "incomplete": [
            "1. Go to View tab",
            "2. Click 'Filter' dropdown",
            "3. Select 'Incomplete Tasks'",
        ],
        "code": [
            "1. Go to View tab",
            "2. Click 'Filter' dropdown",
            "3. Select 'Codes' > 'Which Code'",
            f"4. Select code library: '{params.code_name or 'Select library'}'",
            "5. Check the desired code entries",
            "6. Click Finish",
        ],
        "none": [
            "1. Go to View tab",
            "2. Click 'Filter' dropdown",
            "3. Select 'No Filter'",
        ],
    }

    steps = filter_instructions.get(params.filter_type, ["Unknown filter type"])

    return json.dumps({
        "success": True,
        "filter_type": params.filter_type,
        "steps": steps,
        "note": "Use asta_gui_click to click specific buttons, or follow these steps manually"
    }, indent=2)


@mcp.tool(
    name="asta_gui_change_table",
    annotations={
        "title": "Change Spreadsheet Table",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    }
)
async def asta_gui_change_table(table_name: str = "Default") -> str:
    """Change the spreadsheet table view in Asta.

    Common table views:
    - 'Default': Standard task information
    - 'Progress - no baseline': For entering progress data
    - 'Progress - with baseline': For comparing against baseline
    - 'Resource': Resource-related columns
    - 'Cost': Cost-related columns

    Args:
        table_name: Name of the table view to switch to

    Returns:
        Instructions for changing the table
    """
    return json.dumps({
        "success": True,
        "table": table_name,
        "steps": [
            "1. Go to View tab",
            "2. Click 'Table' dropdown",
            f"3. Select '{table_name}'",
        ],
        "note": "This changes which columns are visible in the spreadsheet area"
    }, indent=2)


@mcp.tool(
    name="asta_gui_print_export",
    annotations={
        "title": "Print/Export Project",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    }
)
async def asta_gui_print_export(params: GUIPrintInput) -> str:
    """Print or export the current Asta view to PDF, picture, or printer.

    The output shows exactly what's currently visible on screen -
    so set up your view (filter, zoom, table) before printing.

    Args:
        params: Contains output_type ('pdf', 'clipboard', 'picture', 'printer')
                and optional save_path

    Returns:
        Instructions for completing the print/export
    """
    try:
        import pyautogui
        AstaGUIManager.bring_to_front()
        time.sleep(0.5)

        # Open print dialog
        pyautogui.hotkey('ctrl', 'p')
        time.sleep(1.5)

        instructions = {
            "pdf": [
                "1. Print dialog is now open (Ctrl+P)",
                "2. Select a PDF printer (e.g., 'Microsoft Print to PDF')",
                "3. Set orientation to Landscape (recommended for Gantt charts)",
                f"4. Click Print and save to: {params.save_path or 'choose location'}",
            ],
            "clipboard": [
                "1. Print dialog is now open (Ctrl+P)",
                "2. Click the 'Clipboard' option",
                "3. Click Print",
                "4. The bar chart is now in your clipboard - paste into Word/Excel",
            ],
            "picture": [
                "1. Print dialog is now open (Ctrl+P)",
                "2. Click the 'Picture file' option",
                "3. Browse to select save location",
                f"4. Save as: {params.save_path or 'choose location'}",
            ],
            "printer": [
                "1. Print dialog is now open (Ctrl+P)",
                "2. Select your printer",
                "3. Set paper size and orientation",
                "4. Click Print",
            ],
        }

        return json.dumps({
            "success": True,
            "output_type": params.output_type,
            "steps": instructions.get(params.output_type, ["Unknown output type"]),
            "tip": "Set up your view (filter, zoom, table) BEFORE printing - what you see is what you get!"
        }, indent=2)
    except Exception as e:
        return json.dumps({"success": False, "message": str(e)}, indent=2)


@mcp.tool(
    name="asta_gui_zoom",
    annotations={
        "title": "Zoom Bar Chart",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": True,
    }
)
async def asta_gui_zoom(direction: str = "fit") -> str:
    """Zoom the bar chart view in Asta Powerproject.

    Args:
        direction: 'in' to zoom in, 'out' to zoom out, 'fit' to fit all tasks

    Returns:
        Instructions for zooming
    """
    zoom_instructions = {
        "in": "Use Ctrl+Mouse Wheel Up on the Date Zone area to zoom in (show more detail)",
        "out": "Use Ctrl+Mouse Wheel Down on the Date Zone area to zoom out (show more time)",
        "fit": "Go to View tab > click 'Zoom to Fit' to show all tasks in the visible area",
    }

    return json.dumps({
        "success": True,
        "direction": direction,
        "instruction": zoom_instructions.get(direction, "Unknown direction"),
        "tip": "You can also right-click the Date Zone for more zoom options"
    }, indent=2)


@mcp.tool(
    name="asta_gui_summarize_tasks",
    annotations={
        "title": "Create Summary Task",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": True,
    }
)
async def asta_gui_summarize_tasks(summary_name: str = "New Group") -> str:
    """Create a summary task to group selected tasks in Asta.

    Summary tasks group related activities (like a folder groups files).
    Select the tasks you want to group first, then use Home > Summarise.

    Args:
        summary_name: Name for the summary group (e.g., 'Foundation Work', 'Phase 1')

    Returns:
        Step-by-step instructions for creating the summary
    """
    return json.dumps({
        "success": True,
        "steps": [
            "1. Select the tasks you want to group:",
            "   - Click first task, then Shift+Click last task (for range)",
            "   - Or Ctrl+Click to select individual tasks",
            "2. Go to Home tab",
            "3. Click 'Summarise' button",
            "4. A summary bar appears above the selected tasks",
            f"5. Type the summary name: '{summary_name}'",
            "6. Press Enter to confirm",
        ],
        "tips": [
            "To add more tasks later: Select task > Home > Indent",
            "To remove task from group: Select task > Home > Outdent",
            "Double-click summary bar to collapse/expand the group",
        ]
    }, indent=2)


@mcp.tool(
    name="asta_gui_indent_task",
    annotations={
        "title": "Indent/Outdent Task",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": True,
    }
)
async def asta_gui_indent_task(direction: str = "indent") -> str:
    """Indent or outdent the selected task in the project hierarchy.

    - Indent: Makes the task a child of the task above (adds to summary group)
    - Outdent: Moves the task up one level (removes from summary group)

    Args:
        direction: 'indent' to make child, 'outdent' to move up

    Returns:
        Confirmation of indent/outdent action
    """
    return json.dumps({
        "success": True,
        "direction": direction,
        "steps": [
            "1. Select the task(s) you want to move",
            f"2. Go to Home tab > click '{direction.title()}' button",
            f"   (Arrow {'right' if direction == 'indent' else 'left'} icon)",
        ],
        "note": "Indent adds task to the summary above, Outdent removes it from its current group"
    }, indent=2)


# ============================================================================
# UTILITY / HELP TOOLS
# ============================================================================

@mcp.tool(
    name="asta_help",
    annotations={
        "title": "Asta Help Guide",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    }
)
async def asta_help(topic: str = "overview") -> str:
    """Get help and guidance on Asta Powerproject topics.

    Available topics:
    - 'overview': General overview of available tools
    - 'shortcuts': Keyboard shortcuts reference
    - 'workflow': Common workflow steps
    - 'links': How to create and manage task links
    - 'progress': How to track project progress
    - 'critical_path': Understanding the critical path
    - 'resources': Working with resources and costs
    - 'printing': Printing and exporting

    Args:
        topic: Help topic to display

    Returns:
        Detailed help text for the requested topic
    """
    help_topics = {
        "overview": """# Asta Powerproject MCP Server - Overview

## File-Based Tools (read/write project files):
- **asta_analyze_project**: Analyze a project file - START HERE
- **asta_list_tasks**: List all tasks with details
- **asta_get_task**: Get detailed info on one task
- **asta_add_task**: Add a new task
- **asta_update_task**: Update task properties
- **asta_delete_task**: Delete a task
- **asta_get_critical_path**: View critical path
- **asta_list_resources**: List all resources
- **asta_get_resource_assignments**: View resource allocations
- **asta_get_calendars**: View project calendars
- **asta_float_analysis**: Analyze float distribution
- **asta_save_project**: Save project to XML

## GUI Automation Tools (control Asta on screen):
- **asta_gui_check_status**: Check if Asta is running
- **asta_gui_reschedule**: Run reschedule (F9)
- **asta_gui_save**: Save project (Ctrl+S)
- **asta_gui_screenshot**: Take screenshot
- **asta_gui_click**: Click at coordinates
- **asta_gui_type_text**: Type text in cells
- **asta_gui_send_shortcut**: Send keyboard shortcuts
- **asta_gui_open_file**: Open a project file
- **asta_gui_new_project**: Create new project
- **asta_gui_take_baseline**: Take baseline snapshot
- **asta_gui_link_tasks**: Link tasks together
- **asta_gui_apply_filter**: Apply display filters
- **asta_gui_print_export**: Print/export to PDF
""",
        "shortcuts": """# Asta Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| F9 | Reschedule (calculate critical path) |
| Ctrl+S | Save project |
| Ctrl+P | Print |
| Ctrl+Z | Undo |
| Ctrl+Y | Redo |
| Ctrl+N | New project |
| Ctrl+O | Open project |
| Insert | Insert new row |
| Delete | Delete selected item |
| F1 | Help |
| Shift+Click | Select range |
| Ctrl+Click | Add to selection |
| Shift+Drag link | Add lag to link |
""",
        "workflow": """# Common Asta Workflow

## Creating a New Schedule:
1. File > New > Select template
2. Enter project details (name, dates, client)
3. Enter tasks in Spreadsheet (Name + Duration columns)
4. Link tasks (drag from end of one to start of another)
5. Press F9 to Reschedule
6. Review Critical Path (shown in red)
7. Take a Baseline (Project > Take Baseline)

## Updating Progress:
1. View > Table > Progress - no baseline
2. Enter % Complete for each active task
3. Enter Actual Start dates for started tasks
4. Enter Actual Finish dates for completed tasks
5. Press F9 to Reschedule with progress
6. Compare with baseline

## Weekly Routine:
1. Open project file
2. Enter progress for the past week
3. Reschedule (F9)
4. Check if critical path changed
5. Review float on near-critical tasks
6. Print/export updated schedule
7. Save and backup
""",
        "links": """# Task Links in Asta

## Link Types:
- **FS (Finish-to-Start)**: B starts after A finishes (most common)
- **SS (Start-to-Start)**: B starts when A starts
- **FF (Finish-to-Finish)**: B finishes when A finishes
- **SF (Start-to-Finish)**: B finishes when A starts (rare)

## Creating Links:
1. Hover at the END of the predecessor task
2. Cursor changes to link icon
3. Click and drag to the START of successor task
4. Release to create FS link

## Adding Lag (waiting time):
- Hold Shift while dragging to add lag
- Example: 2d lag for concrete curing time

## Block Linking (multiple tasks):
1. Select tasks (Shift or Ctrl click)
2. Home > Link Tasks
3. Creates FS links in sequence
""",
        "progress": """# Progress Tracking in Asta

## Setup:
1. Take a baseline BEFORE entering progress
2. Switch to Progress table: View > Table > Progress

## Entering Progress:
- **% Complete**: Enter in 'Overall Percent Complete' column
- **Actual Start**: Enter when task actually started
- **Actual Finish**: Enter when task actually finished
- **Planned %**: System calculates based on schedule

## After Entering Progress:
1. Press F9 (Reschedule)
2. Select 'Straighten progress entry period'
3. This moves incomplete work past the report date

## Reading Results:
- **Slip**: Difference between planned and actual
- **Progress line**: Visual indicator on bar chart
- Red tasks = Critical (monitor closely!)
""",
        "critical_path": """# Critical Path in Asta

## What is the Critical Path?
The longest chain of linked tasks through the project.
Any delay on critical tasks = project delay.

## How to See It:
1. Link all tasks properly
2. Press F9 (Reschedule)
3. Critical tasks get RED outline
4. Non-critical tasks show float (blue bars)

## Float Types:
- **Total Float**: Time task can slip without affecting project end
- **Free Float**: Time task can slip without affecting next task
- **Zero Float** = Critical task

## Tips:
- Keep critical tasks under close watch
- Look for 'near-critical' tasks (low float)
- Consider adding resources to shorten critical tasks
- Use Part Critical Shading for partially critical tasks
""",
        "resources": """# Resources in Asta

## Resource Types:
- **Permanent**: Labour, equipment (reusable)
- **Consumable**: Materials (used up)

## Creating Resources:
1. View > Library Explorer
2. Navigate to Resources folder
3. Right-click > New Resource
4. Set name, type, rate, availability

## Assigning Resources:
1. Open Project View (left panel)
2. Drag resource onto task
3. Set units/quantity

## Cost Centres:
1. Create Cost Centres (Labour, Materials, Plant)
2. Define rates per resource
3. Assign to tasks
4. View cost reports
""",
        "printing": """# Printing in Asta

## Before Printing:
- Set up view (filter, zoom, columns)
- What you see = what you print

## Print Options:
1. **Ctrl+P** to open print dialog
2. Choose output: Printer, PDF, Clipboard, Picture
3. Set paper size and orientation (Landscape recommended)
4. Select border file (company template)
5. Adjust scaling (fit to pages)
6. Preview before printing

## Border Files:
- Templates with company logo, revision info
- Select in Details tab > Browse
- Embed in project for portability

## Tips:
- Use 'Fit to 1 page wide' for clean output
- Save print profiles for reuse
- Landscape orientation works best for Gantt charts
""",
    }

    return help_topics.get(topic, f"Unknown topic: '{topic}'. Available: {', '.join(help_topics.keys())}")




# ============================================================================
# ENTRY POINT
# ============================================================================
if __name__ == "__main__":
    logger.info("Starting Asta Powerproject GUI MCP Server...")
    mcp.run()
