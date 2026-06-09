"""P2 #7 — Asta bulk_add task/link (verifiable non-COM paths only).

The COM execution path requires a live Asta session and is NOT exercised
here (see asta_com_bulk_add_* docstrings: marked needs-live-verification).
These tests cover the input-validation / error paths that run without COM.
"""
import asyncio
import json
import pytest

import asta_mcp_core as a


def test_bulk_add_tasks_empty_items_error():
    out = json.loads(asyncio.run(a.asta_com_bulk_add_tasks([])))
    assert "error" in out
    assert "empty" in out["error"].lower()


def test_bulk_add_links_empty_items_error():
    out = json.loads(asyncio.run(a.asta_com_bulk_add_links([])))
    assert "error" in out
    assert "empty" in out["error"].lower()


def test_bulk_add_functions_exist_and_callable():
    assert callable(a.asta_com_bulk_add_tasks)
    assert callable(a.asta_com_bulk_add_links)
