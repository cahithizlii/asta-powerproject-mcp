"""Phase 10.4 — shared helpers for samples/build_*_lifecycle.py scripts.

Extracted boilerplate previously copy-pasted across build_compare_lifecycle.py
and build_currency_validation_lifecycle.py:
- REPO_ROOT path setup
- write_synthetic_xer (UTF-16-LE BOM XER fixture builder)
- call_async_dispatcher (asyncio.run + json.loads wrapper for MCP tools)
- print_section (banner + truncated JSON dump)

Sample scripts import these helpers via:
    from samples._lifecycle_common import (
        write_synthetic_xer, call_async_dispatcher, print_section,
    )
"""
import asyncio
import json
import os
import sys
import tempfile

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)


def write_synthetic_xer(content: str, name: str) -> str:
    """Write a UTF-16-LE BOM XER fixture to the system tempdir.

    Args:
        content: full XER text (ERMHDR + table sections + %E).
        name: bare filename (no path); placed in tempfile.gettempdir().

    Returns:
        absolute path to the written file.
    """
    path = os.path.join(tempfile.gettempdir(), name)
    with open(path, "wb") as f:
        f.write(b"\xff\xfe")
        f.write(content.encode("utf-16-le"))
    return path


def call_async_dispatcher(dispatcher, action: str, **kw) -> dict:
    """Run an async @mcp.tool dispatcher and parse its JSON result.

    Args:
        dispatcher: the async function (e.g. msproject_compare).
        action: the action key.
        **kw: remaining params forwarded to the dispatcher.

    Returns:
        Parsed JSON dict (already through json.loads).
    """
    raw = asyncio.run(dispatcher({"action": action, **kw}))
    return json.loads(raw)


def print_section(title: str, payload: dict, max_chars: int = 800) -> None:
    """Pretty-print a labelled JSON section with separator banner.

    Truncates payloads larger than max_chars to keep stdout readable in
    CI/manual runs. Output goes through json.dumps(default=str) so dates
    and non-trivial types are still readable.
    """
    print(f"\n{'=' * 70}\n  {title}\n{'=' * 70}")
    s = json.dumps(payload, indent=2, ensure_ascii=False, default=str)
    if len(s) > max_chars:
        s = s[:max_chars] + "\n... [truncated]"
    print(s)
