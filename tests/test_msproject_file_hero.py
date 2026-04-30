"""🚀 T73 HERO test — file-based MSPDI bulk_add_assignments performance.

Phase 4 success gate: 2800 task-resource assignments in <5s via the
file MCP path (mspdi_parser.bulk_add_assignments + single XML write).

This test does NOT exercise MSP COM merge — it measures the pure
Python XML write path. See Phase 2b's test_bulk_assign_hero_2800_under_5s
for the COM-bound xfail (which is a different gate, deferred).
"""
import os
import shutil
import sys
import time

import pytest

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from msproject_mcp_core import (  # noqa: E402
    _msp_file_add_tasks, _msp_file_add_resources,
    _msp_file_bulk_add_assignments, _msp_file_read_assignments,
)

FIXTURE = os.path.join(os.path.dirname(__file__), "fixtures", "sample_msp.xml")


@pytest.fixture
def hero_xml(tmp_path):
    """Copy sample fixture for in-place editing."""
    dst = tmp_path / "hero.xml"
    shutil.copy(FIXTURE, dst)
    return str(dst)


def test_hero_2800_assignments_under_5s_via_file_xml(hero_xml):
    """200 task * 14 resource = 2800 assignments via file MCP path <5s."""
    # Build 200 tasks (sample fixture starts with 3, so add 200 more)
    task_items = [{"name": f"V{i:03d}", "duration": "2d"} for i in range(200)]
    r1 = _msp_file_add_tasks(file_path=hero_xml, items=task_items)
    assert r1["status"] == "ok"
    new_task_ids = r1["task_ids"]
    assert len(new_task_ids) == 200

    # Add 14 resources (sample fixture starts with R1, R2)
    res_items = [{"name": f"CR{i:02d}", "type": "Work"} for i in range(14)]
    r2 = _msp_file_add_resources(file_path=hero_xml, items=res_items)
    assert r2["status"] == "ok"
    new_res_ids = r2["resource_ids"]
    assert len(new_res_ids) == 14

    # 2800 assignment items
    assignment_items = [{"task_id": tid, "resource_id": rid}
                        for tid in new_task_ids for rid in new_res_ids]
    assert len(assignment_items) == 2800

    # 🚀 HERO measurement
    start = time.time()
    r3 = _msp_file_bulk_add_assignments(file_path=hero_xml, items=assignment_items)
    elapsed = time.time() - start

    assert r3["status"] == "ok"
    assert r3["count"] == 2800
    assert elapsed < 5.0, (
        f"HERO bulk_add_assignments took {elapsed:.2f}s (target <5s). "
        f"reported elapsed_s={r3.get('elapsed_s')}"
    )

    # Verify by re-reading
    r4 = _msp_file_read_assignments(file_path=hero_xml)
    assert r4["status"] == "ok"
    # Existing 6 + 2800 new = 2806
    assert r4["count"] >= 2800


def test_hero_empty_items_no_op(hero_xml):
    r = _msp_file_bulk_add_assignments(file_path=hero_xml, items=[])
    assert r["status"] == "ok"
    assert r["count"] == 0


def test_hero_invalid_file():
    r = _msp_file_bulk_add_assignments(file_path="/nonexistent.xml",
                                       items=[{"task_id": 1, "resource_id": 1}])
    assert r["status"] == "error"


def test_hero_skips_missing_ids(hero_xml):
    """Missing task or resource IDs are silently skipped (bulk semantics)."""
    r = _msp_file_bulk_add_assignments(file_path=hero_xml, items=[
        {"task_id": 1, "resource_id": 1},
        {"task_id": 99999, "resource_id": 1},  # missing task → skip
        {"task_id": 1, "resource_id": 99999},  # missing resource → skip
    ])
    assert r["status"] == "ok"
    assert r["count"] == 1  # only the first valid pair
