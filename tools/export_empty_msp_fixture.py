"""One-shot helper: export the active MS Project document as an MSPDI XML fixture.

Run with MS Project open on an empty (or test) project.
The output ends up at tests/fixtures/empty_msp.xml.

Usage:
    python tools/export_empty_msp_fixture.py
"""
import os
import sys
import pythoncom
import win32com.client


def export_empty_fixture():
    pythoncom.CoInitialize()
    try:
        app = win32com.client.GetActiveObject("MSProject.Application")
    except Exception as exc:
        print(f"ERROR: Could not connect to MS Project: {exc}")
        sys.exit(1)

    if app.ActiveProject is None:
        print("ERROR: No active project in MS Project")
        sys.exit(1)

    repo_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    out = os.path.join(repo_root, "tests", "fixtures", "empty_msp.xml")
    os.makedirs(os.path.dirname(out), exist_ok=True)

    # Remove any old fixture so FileSaveAs does not get prompted to overwrite.
    if os.path.exists(out):
        os.remove(out)

    # MS Project FileSaveAs: Name (file path), FormatID (string).
    # FormatID = "MSProject.XML" → MSPDI XML export.
    try:
        app.FileSaveAs(Name=out, FormatID="MSProject.XML")
    except Exception as exc:
        print(f"ERROR: FileSaveAs failed: {exc}")
        sys.exit(1)

    if not os.path.exists(out):
        print("ERROR: FileSaveAs returned but file does not exist")
        sys.exit(1)

    size = os.path.getsize(out)
    print(f"Exported: {out} ({size} bytes)")


if __name__ == "__main__":
    export_empty_fixture()
