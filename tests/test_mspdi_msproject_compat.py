"""Verify mspdi_parser handles MS Project XML (not just Asta exports).

Both Asta Powerproject and MS Project speak MSPDI (Microsoft Project XML
Data Interchange) — same `http://schemas.microsoft.com/project` namespace.
This test confirms the parser written for Asta XML also round-trips MS
Project's own XML output without modification.
"""
import os
import sys

import pytest

# Repo root on sys.path so `import mspdi_parser` works no matter where pytest
# is invoked from.
REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from mspdi_parser import MspdiProject  # noqa: E402


@pytest.fixture
def empty_msp_xml(fixtures_dir):
    path = os.path.join(fixtures_dir, "empty_msp.xml")
    if not os.path.exists(path):
        pytest.skip(f"Fixture missing: {path}")
    return path


def test_empty_msp_xml_loads(empty_msp_xml):
    """Empty MS Project export should be parseable."""
    proj = MspdiProject(empty_msp_xml)
    assert proj is not None
    summary = proj.get_project_summary()
    assert "project_name" in summary
    # MSP export of an empty project always has at least the root summary task.
    assert summary["total_tasks"] >= 1


def test_msp_xml_round_trip(empty_msp_xml, tmp_path):
    """Read → modify → save → re-read symmetry."""
    proj = MspdiProject(empty_msp_xml)
    new_task = proj.add_task(name="Test Task", duration_str="3d")
    # mspdi_parser.add_task() returns "task_id" (not "id"); see mspdi_parser.py
    # ~line 1196.  Both keys are acceptable — accept whichever exists.
    assert "task_id" in new_task or "id" in new_task

    output = tmp_path / "modified.xml"
    proj.save(output_path=str(output))
    assert output.exists()

    proj2 = MspdiProject(str(output))
    tasks = proj2.get_all_tasks(include_summary=False)
    names = [t["name"] for t in tasks]
    assert "Test Task" in names


def test_msp_xml_namespace_preserved(empty_msp_xml, tmp_path):
    """After save, output must keep the MSPDI namespace declaration."""
    proj = MspdiProject(empty_msp_xml)
    proj.add_task(name="NS Probe", duration_str="1d")
    output = tmp_path / "ns_check.xml"
    proj.save(output_path=str(output))

    content = output.read_text(encoding="utf-8")
    assert "http://schemas.microsoft.com/project" in content
    # ElementTree sometimes emits ns0: prefixes — mspdi_parser.save() should
    # have stripped those during post-processing.
    assert "ns0:" not in content
