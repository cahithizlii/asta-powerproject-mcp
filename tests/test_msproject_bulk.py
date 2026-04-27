"""Test MSPDI bulk-write engine for MS Project."""
import pytest
from msproject_bulk import MsprojectBulkWriter


def test_writer_creates_empty_project(tmp_path):
    """Empty bulk writer should produce a valid MSPDI XML."""
    w = MsprojectBulkWriter(project_name="Test Bulk")
    out = tmp_path / "test_bulk.xml"
    w.save(str(out))
    assert out.exists()
    content = out.read_text(encoding="utf-8")
    assert "<?xml" in content
    assert "Test Bulk" in content


def test_writer_adds_tasks(tmp_path):
    """Bulk add 50 tasks → all present in output."""
    w = MsprojectBulkWriter(project_name="Bulk 50")
    items = [{"name": f"Task {i}", "duration": "1d"} for i in range(50)]
    w.bulk_add_tasks(items)
    out = tmp_path / "bulk50.xml"
    w.save(str(out))
    # Re-read with mspdi_parser to verify
    from mspdi_parser import MspdiProject
    p = MspdiProject(str(out))
    tasks = p.get_all_tasks(include_summary=False)
    assert len(tasks) == 50


def test_writer_adds_links(tmp_path):
    """Bulk add tasks + links → links present in output."""
    w = MsprojectBulkWriter(project_name="Bulk Links")
    items = [{"name": f"T{i}", "duration": "2d"} for i in range(5)]
    uids = w.bulk_add_tasks(items)
    # Chain: T0→T1→T2→T3→T4
    links = [{"pred_uid": uids[i], "succ_uid": uids[i+1], "type": "FS", "lag": "0d"}
             for i in range(4)]
    count = w.bulk_add_links(links)
    assert count == 4
    out = tmp_path / "bulk_links.xml"
    w.save(str(out))
    content = out.read_text(encoding="utf-8")
    assert "PredecessorLink" in content


def test_writer_handles_turkish_chars(tmp_path):
    """UTF-8 encoding for Turkish characters."""
    w = MsprojectBulkWriter(project_name="Şantiye İmalatları")
    items = [
        {"name": "Çatı İmalatı", "duration": "5d"},
        {"name": "Müşterek Boya", "duration": "3d"},
    ]
    w.bulk_add_tasks(items)
    out = tmp_path / "turkish.xml"
    w.save(str(out))
    content = out.read_text(encoding="utf-8")
    assert "Şantiye" in content
    assert "Çatı" in content
    assert "Müşterek" in content


def test_duration_subminute_does_not_truncate():
    """30m should produce PT0H30M0S not PT0H0M0S (regression test for C1)."""
    iso = MsprojectBulkWriter._duration_to_iso("30m")
    assert iso == "PT0H30M0S"
    iso2 = MsprojectBulkWriter._duration_to_iso("90m")
    assert iso2 == "PT1H30M0S"
    iso3 = MsprojectBulkWriter._duration_to_iso("5d")
    assert iso3 == "PT40H0M0S"


def test_resource_type_int_codes(tmp_path):
    """Resource Type must be int per MSPDI spec, accept various inputs (C2)."""
    w = MsprojectBulkWriter(project_name="Res Test")
    w.bulk_add_resources([
        {"name": "Worker", "type": "Work"},
        {"name": "Concrete", "type": "Material"},
        {"name": "Subcontractor", "type": "Cost"},
    ])
    out = tmp_path / "res.xml"
    w.save(str(out))
    content = out.read_text(encoding="utf-8")
    # Should contain integer types per MSPDI spec
    assert "<Type>1</Type>" in content  # Work
    assert "<Type>0</Type>" in content  # Material
    assert "<Type>2</Type>" in content  # Cost

    # Lowercase and int inputs should also work
    w2 = MsprojectBulkWriter(project_name="Res Test 2")
    w2.bulk_add_resources([
        {"name": "lower-work", "type": "work"},
        {"name": "int-material", "type": 0},
        {"name": "int-cost", "type": 2},
    ])
    out2 = tmp_path / "res2.xml"
    w2.save(str(out2))
    content2 = out2.read_text(encoding="utf-8")
    assert "<Type>1</Type>" in content2
    assert "<Type>0</Type>" in content2
    assert "<Type>2</Type>" in content2


def test_uid_separate_spaces_for_tasks_and_resources(tmp_path):
    """Task UIDs and Resource UIDs are in separate counters starting at 1 (C3)."""
    w = MsprojectBulkWriter(project_name="UID Test")
    res_uids = w.bulk_add_resources([{"name": "R1"}, {"name": "R2"}])
    task_uids = w.bulk_add_tasks([{"name": "T1", "duration": "1d"},
                                   {"name": "T2", "duration": "2d"}])
    assert res_uids == [1, 2], f"Resources should start at UID 1, got {res_uids}"
    assert task_uids == [1, 2], f"Tasks should start at UID 1 (separate space), got {task_uids}"
