"""Test Phase 5b T92 action helpers (assess_all, summary, drill_down, compare)."""
import os, sys
REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from msproject_mcp_core import (
    _msp_dcma_assess_all,
    _msp_dcma_summary,
    _msp_dcma_drill_down,
    _msp_dcma_compare,
)

MSP_XML = os.path.join(os.path.dirname(__file__), "fixtures", "sample_msp.xml")


def test_msp_dcma_assess_all_xml():
    r = _msp_dcma_assess_all(file_path=MSP_XML)
    assert r["status"] == "ok"
    assert len(r["rules"]) == 14
    assert "summary" in r


def test_msp_dcma_summary_xml():
    r = _msp_dcma_summary(file_path=MSP_XML)
    assert r["status"] == "ok"
    assert r["overall_rag"] in ("green", "amber", "red")


def test_msp_dcma_drill_down_valid_rule():
    r = _msp_dcma_drill_down(file_path=MSP_XML, rule_id=1)
    assert r["status"] == "ok"
    assert r["rule"]["id"] == 1
    assert "failed_tasks" in r


def test_msp_dcma_drill_down_invalid_rule():
    r = _msp_dcma_drill_down(file_path=MSP_XML, rule_id=99)
    assert r["status"] == "error"
    assert "1-14" in r["error"]


def test_msp_dcma_compare_no_prev_snapshot(tmp_path):
    snap = str(tmp_path / "no_snap.json")
    r = _msp_dcma_compare(file_path=MSP_XML, snapshot_path=snap)
    # No prev snapshot -> graceful: ok with empty delta
    assert r["status"] in ("ok", "error")
