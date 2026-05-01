import os, sys, json
REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)
from msproject_mcp_core import (
    _msp_evm_save_period_snapshot,
    _msp_evm_get_period_history,
    _msp_evm_trend,
)

MSP_XML = os.path.join(os.path.dirname(__file__), "fixtures", "sample_msp.xml")


def test_msp_evm_save_period_snapshot(tmp_path):
    snap_path = str(tmp_path / "snaps.json")
    r = _msp_evm_save_period_snapshot(file_path=MSP_XML,
                                      snapshot_path=snap_path,
                                      tag="test-week")
    assert r["status"] == "ok"
    assert os.path.exists(snap_path)
    data = json.loads(open(snap_path).read())
    assert len(data["snapshots"]) == 1
    assert data["snapshots"][0]["tag"] == "test-week"


def test_msp_evm_save_appends(tmp_path):
    snap_path = str(tmp_path / "snaps.json")
    _msp_evm_save_period_snapshot(file_path=MSP_XML,
                                  snapshot_path=snap_path, tag="w1")
    _msp_evm_save_period_snapshot(file_path=MSP_XML,
                                  snapshot_path=snap_path, tag="w2")
    data = json.loads(open(snap_path).read())
    assert len(data["snapshots"]) == 2


def test_msp_evm_get_period_history_filter(tmp_path):
    snap_path = str(tmp_path / "snaps.json")
    _msp_evm_save_period_snapshot(file_path=MSP_XML,
                                  snapshot_path=snap_path, tag="all")
    r = _msp_evm_get_period_history(snapshot_path=snap_path)
    assert r["status"] == "ok"
    assert len(r["snapshots"]) >= 1


def test_msp_evm_trend_returns_series(tmp_path):
    snap_path = str(tmp_path / "snaps.json")
    _msp_evm_save_period_snapshot(file_path=MSP_XML,
                                  snapshot_path=snap_path, tag="w1")
    _msp_evm_save_period_snapshot(file_path=MSP_XML,
                                  snapshot_path=snap_path, tag="w2")
    r = _msp_evm_trend(snapshot_path=snap_path)
    assert r["status"] == "ok"
    assert "series" in r
