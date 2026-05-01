import os, sys
REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)
from msproject_mcp_core import _msp_evm_time_phased_evm

MSP_XML = os.path.join(os.path.dirname(__file__), "fixtures", "sample_msp.xml")


def test_msp_evm_time_phased_week_buckets():
    r = _msp_evm_time_phased_evm(file_path=MSP_XML, bucket="week")
    assert r["status"] == "ok"
    assert "buckets" in r
    for b in r["buckets"]:
        assert "period_start" in b and "period_end" in b
        assert "pv" in b and "ev" in b and "ac" in b


def test_msp_evm_time_phased_day_buckets():
    r = _msp_evm_time_phased_evm(file_path=MSP_XML, bucket="day")
    assert r["status"] == "ok"


def test_msp_evm_time_phased_invalid_bucket():
    r = _msp_evm_time_phased_evm(file_path=MSP_XML, bucket="invalid")
    assert r["status"] == "error"
    assert "bucket" in r["error"].lower()
