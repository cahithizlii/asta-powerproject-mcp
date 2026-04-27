"""Test msproject_task CRUD operations."""
import pytest
from msproject_mcp_core import (
    _msp_task_add_single, _msp_task_update, _msp_task_delete,
    _msp_task_get, _msp_task_list,
)


@pytest.fixture
def temp_task(msproject_app):
    """Create + cleanup a temp task."""
    result = _msp_task_add_single(name="Temp", duration="2d")
    task_id = result["task_id"]
    yield task_id
    try:
        _msp_task_delete(task_id=task_id)
    except Exception:
        pass


def test_get_task(temp_task):
    r = _msp_task_get(task_id=temp_task)
    assert r["status"] == "ok"
    assert r["task"]["name"] == "Temp"


def test_update_task(temp_task):
    r = _msp_task_update(task_id=temp_task, name="Renamed", duration="5d", notes="changed")
    assert r["status"] == "ok"
    g = _msp_task_get(task_id=temp_task)
    assert g["task"]["name"] == "Renamed"
    assert g["task"]["notes"] == "changed"


def test_delete_task(msproject_app):
    r = _msp_task_add_single(name="ToDelete", duration="1d")
    tid = r["task_id"]
    proj = msproject_app.ActiveProject
    before = proj.Tasks.Count
    d = _msp_task_delete(task_id=tid)
    assert d["status"] == "ok"
    assert proj.Tasks.Count == before - 1


def test_list_tasks(temp_task):
    r = _msp_task_list(limit=200)
    assert r["status"] == "ok"
    names = [t["name"] for t in r["tasks"]]
    assert "Temp" in names
