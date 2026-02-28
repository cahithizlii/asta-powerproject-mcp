#!/usr/bin/env python3
"""Split asta_mcp_server.py into asta_mcp_core.py and asta_mcp_gui.py"""

import os

src = r'C:\Users\CahAsus\asta-powerproject-mcp\asta_mcp_server.py'
with open(src, 'r', encoding='utf-8') as f:
    lines = f.readlines()

print(f'Source: {len(lines)} lines')

# =====================================================================
# CORE MCP
# =====================================================================
core_lines = []

# Section 1: Imports, logging, constants, JVM, helpers, AstaFileManager (lines 1-1048)
for i in range(0, 1048):
    line = lines[i]
    if 'mcp = FastMCP("asta_powerproject_mcp")' in line:
        line = 'mcp = FastMCP("asta_powerproject_core")\n'
    if 'logger = logging.getLogger("asta_mcp")' in line:
        line = 'logger = logging.getLogger("asta_mcp_core")\n'
    if '"asta_mcp.log"' in line:
        line = line.replace("asta_mcp.log", "asta_mcp_core.log")
    core_lines.append(line)

# Section 2: Core Pydantic models (lines 1212-1411)
core_lines.append('\n')
core_lines.append('# ============================================================================\n')
core_lines.append('# PYDANTIC INPUT MODELS (Core)\n')
core_lines.append('# ============================================================================\n')
for i in range(1211, 1411):
    core_lines.append(lines[i])

# Section 3: File-based @mcp.tool functions (lines 1545-2821)
core_lines.append('\n')
core_lines.append('# ============================================================================\n')
core_lines.append('# FILE-BASED TOOLS (MPXJ) with COM-first when file_path is omitted\n')
core_lines.append('# ============================================================================\n')
for i in range(1544, 2815):
    core_lines.append(lines[i])

# Section 4: COM infrastructure (lines 3846-4505)
core_lines.append('\n')
for i in range(3845, 4505):
    core_lines.append(lines[i])

# Section 5: asta_reschedule_project (lines 4508-4840)
for i in range(4507, 4840):
    core_lines.append(lines[i])

# Section 6: Phase 3 tools (lines 4841-7148)
for i in range(4840, 7148):
    core_lines.append(lines[i])

# Entry point
core_lines.append('\n\n')
core_lines.append('# ============================================================================\n')
core_lines.append('# ENTRY POINT\n')
core_lines.append('# ============================================================================\n')
core_lines.append('if __name__ == "__main__":\n')
core_lines.append('    logger.info("Starting Asta Powerproject Core MCP Server...")\n')
core_lines.append('    mcp.run()\n')

core_path = r'C:\Users\CahAsus\asta-powerproject-mcp\asta_mcp_core.py'
with open(core_path, 'w', encoding='utf-8') as f:
    f.writelines(core_lines)
print(f'Core MCP: {len(core_lines)} lines -> {core_path}')

# =====================================================================
# GUI MCP
# =====================================================================
gui_lines = []

# Header + imports
gui_header = '''#!/usr/bin/env python3
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

'''
gui_lines.append(gui_header)

# Helper functions needed by GUI
gui_lines.append('# ============================================================================\n')
gui_lines.append('# HELPERS\n')
gui_lines.append('# ============================================================================\n')
# clean_turkish (lines 74-86)
for i in range(73, 87):
    gui_lines.append(lines[i])
gui_lines.append('\n')
# _get_powershell_path (lines 179-192)
for i in range(178, 192):
    gui_lines.append(lines[i])
gui_lines.append('\n')
# _clipboard_paste (lines 194-206)
for i in range(193, 207):
    gui_lines.append(lines[i])
gui_lines.append('\n')

# AstaGUIManager class (lines 1049-1206)
gui_lines.append('\n')
gui_lines.append('# ============================================================================\n')
gui_lines.append('# GUI AUTOMATION MANAGER CLASS\n')
gui_lines.append('# ============================================================================\n')
for i in range(1048, 1207):
    gui_lines.append(lines[i])
gui_lines.append('\n')

# GUI Pydantic models (lines 1412-1539)
gui_lines.append('\n')
gui_lines.append('# ============================================================================\n')
gui_lines.append('# PYDANTIC INPUT MODELS (GUI)\n')
gui_lines.append('# ============================================================================\n')
for i in range(1411, 1540):
    gui_lines.append(lines[i])
gui_lines.append('\n')

# GUI tool functions (lines 2821-3843) - includes asta_gui_* and asta_help
gui_lines.append('\n')
gui_lines.append('# ============================================================================\n')
gui_lines.append('# GUI TOOLS\n')
gui_lines.append('# ============================================================================\n')
for i in range(2820, 3844):
    gui_lines.append(lines[i])

# Entry point
gui_lines.append('\n\n')
gui_lines.append('# ============================================================================\n')
gui_lines.append('# ENTRY POINT\n')
gui_lines.append('# ============================================================================\n')
gui_lines.append('if __name__ == "__main__":\n')
gui_lines.append('    logger.info("Starting Asta Powerproject GUI MCP Server...")\n')
gui_lines.append('    mcp.run()\n')

gui_path = r'C:\Users\CahAsus\asta-powerproject-mcp\asta_mcp_gui.py'
with open(gui_path, 'w', encoding='utf-8') as f:
    f.writelines(gui_lines)
print(f'GUI MCP: {len(gui_lines)} lines -> {gui_path}')

print('\nDone! Both files created.')
