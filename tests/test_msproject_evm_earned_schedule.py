import os, sys
REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)
from msproject_mcp_core import _msp_evm_earned_schedule

MSP_XML = os.path.join(os.path.dirname(__file__), "fixtures", "sample_msp.xml")


def test_msp_evm_earned_schedule_xml():
    r = _msp_evm_earned_schedule(file_path=MSP_XML, bucket="week")
    assert r["status"] == "ok"
    assert "at" in r and "es" in r and "sv_t" in r and "spi_t" in r
