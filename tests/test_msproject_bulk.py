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
