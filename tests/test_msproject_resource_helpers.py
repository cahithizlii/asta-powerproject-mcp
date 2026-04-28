"""Test resource helpers + RESOURCE_TYPES constant."""
import pytest
from msproject_mcp_core import (
    RESOURCE_TYPES, _find_resource_by_name, _find_resource_by_id, _serialize_resource,
)


def test_resource_types_constant():
    """3 types with COM enum codes."""
    assert RESOURCE_TYPES == {"Work": 0, "Material": 1, "Cost": 2}


def test_find_resource_by_name_missing(clean_test_project):
    proj = clean_test_project
    assert _find_resource_by_name(proj, "NonExistent-T32") is None


def test_find_resource_by_name_found(clean_test_project):
    """Add a Work resource via raw COM, find by name."""
    proj = clean_test_project
    raw = proj.Resources.Add("FoundRes-T32")
    res = _find_resource_by_name(proj, "FoundRes-T32")
    assert res is not None
    assert res.ID == raw.ID


def test_find_resource_by_id_existing(clean_test_project):
    """_find_resource_by_id (T23 helper) still works."""
    proj = clean_test_project
    raw = proj.Resources.Add("ByIdRes-T32")
    res = _find_resource_by_id(proj, raw.ID)
    assert res is not None
    assert res.Name == "ByIdRes-T32"


def test_serialize_resource_work(clean_test_project):
    """_serialize_resource returns dict with type-aware fields."""
    proj = clean_test_project
    raw = proj.Resources.Add("SerWork-T32")
    raw.Type = 0  # Work
    raw.MaxUnits = 1.0  # 100%
    raw.StandardRate = 50.0  # $50/h
    d = _serialize_resource(raw)
    assert d["name"] == "SerWork-T32"
    assert d["type"] == "Work"
    assert d["max_units"] == 100.0  # serialized as %
    assert d["standard_rate"] == 50.0
    assert "id" in d and "uid" in d
