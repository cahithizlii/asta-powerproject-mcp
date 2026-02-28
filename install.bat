@echo off
echo ============================================
echo  Asta Powerproject MCP Server - Installer
echo ============================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH!
    echo Please install Python 3.10+ from https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during installation.
    pause
    exit /b 1
)

echo [1/4] Python found. Checking version...
python -c "import sys; assert sys.version_info >= (3, 10), 'Python 3.10+ required'" 2>nul
if errorlevel 1 (
    echo WARNING: Python 3.10 or newer is recommended.
    echo Your current version may still work.
)

echo.
echo [2/4] Installing required packages...
echo This may take a few minutes...
echo.
pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo ERROR: Failed to install some packages.
    echo Try running: pip install -r requirements.txt --break-system-packages
    pause
    exit /b 1
)

echo.
echo [3/4] Verifying installation...
python -c "from mcp.server.fastmcp import FastMCP; print('  MCP SDK: OK')"
python -c "import pydantic; print('  Pydantic: OK')"
python -c "import pyautogui; print('  PyAutoGUI: OK')" 2>nul || echo  PyAutoGUI: SKIPPED (optional for GUI)
python -c "import pywinauto; print('  PyWinAuto: OK')" 2>nul || echo  PyWinAuto: SKIPPED (optional for GUI)

echo.
echo [4/4] Testing server syntax...
python -c "import ast; ast.parse(open('asta_mcp_server.py', encoding='utf-8').read()); print('  Server code: OK')"

echo.
echo ============================================
echo  Installation Complete!
echo ============================================
echo.
echo NEXT STEPS:
echo.
echo 1. Open Claude Desktop settings:
echo    File ^> Settings ^> Developer ^> Edit Config
echo.
echo 2. Add this to your claude_desktop_config.json:
echo.
echo    "asta_powerproject_mcp": {
echo        "command": "python",
echo        "args": ["%~dp0asta_mcp_server.py"]
echo    }
echo.
echo 3. Restart Claude Desktop
echo.
echo 4. You should see "asta_powerproject_mcp" in the tools list
echo.
echo ============================================
pause
