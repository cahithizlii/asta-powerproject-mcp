import os, sys, json
REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)
from msproject_mcp_core import _msp_evm_period_delta

MSP_XML = os.path.join(os.path.dirname(__file__), "fixtures", "sample_msp.xml")


def test_msp_evm_period_delta_first_period(tmp_path):
    """No prev snapshot -> period values = current cum values."""
    snap_path = str(tmp_path / "snaps.json")
    r = _msp_evm_period_delta(file_path=MSP_XML, snapshot_path=snap_path)
    assert r["status"] == "ok"
    assert "period_pv" in r and "period_ev" in r and "period_ac" in r
    assert r["period_bac"] == 0


def test_msp_evm_period_delta_with_prev_snapshot(tmp_path):
    """prev snapshot exists -> period values = current - prev."""
    snap_path = tmp_path / "snaps.json"
    snap_path.write_text(json.dumps({
        "snapshots": [{
            "saved_at": "2026-01-01T00:00:00",
            "metrics": {"pv": 100, "ev": 80, "ac": 90, "bac": 1000},
        }]
    }))
    r = _msp_evm_period_delta(file_path=MSP_XML, snapshot_path=str(snap_path))
    assert r["status"] == "ok"
    # period_pv = current_pv - 100, etc. Just verify keys exist.
    assert "period_pv" in r
