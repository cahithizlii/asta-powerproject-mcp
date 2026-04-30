"""Test Phase 4 file MCP factory + format detection.

T65 foundations: schema sniff helper, factory dispatcher, MspMppFileManager
class shell. These are the prerequisites for T66+ (read actions, write
actions, hero msproject_file tool).
"""
import os
import sys

import pytest

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from msproject_mcp_core import (  # noqa: E402
    _detect_msp_xml_schema,
    _get_msp_file_manager,
    MspMppFileManager,
)
from mspdi_parser import MspdiProject  # noqa: E402

FIXTURE_DIR = os.path.join(os.path.dirname(__file__), "fixtures")
MSP_XML = os.path.join(FIXTURE_DIR, "sample_msp.xml")
EMPTY_MSP = os.path.join(FIXTURE_DIR, "empty_msp.xml")


def test_detect_msp_xml_schema_positive():
    """MSP XML with schemas.microsoft.com/project namespace -> True."""
    assert _detect_msp_xml_schema(MSP_XML) is True
    assert _detect_msp_xml_schema(EMPTY_MSP) is True


def test_detect_msp_xml_schema_missing_file_raises():
    with pytest.raises(FileNotFoundError):
        _detect_msp_xml_schema("/nonexistent/file.xml")


def test_get_manager_xml_returns_mspdi():
    mgr = _get_msp_file_manager(MSP_XML)
    assert isinstance(mgr, MspdiProject)


def test_get_manager_unsupported_extension():
    with pytest.raises(ValueError) as exc:
        _get_msp_file_manager("/path/file.pp")
    assert "extension" in str(exc.value).lower()


def test_msp_mpp_file_manager_init_smoke(tmp_path):
    """MspMppFileManager initializes with .mpp path (smoke - no read needed)."""
    fake_mpp = tmp_path / "fake.mpp"
    fake_mpp.write_bytes(b"\x00" * 100)  # not real MPP, file exists
    mgr = MspMppFileManager(str(fake_mpp))
    assert mgr.file_path.endswith("fake.mpp")
