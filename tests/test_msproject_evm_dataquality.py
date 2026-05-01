import os, sys
REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)
from msproject_mcp_core import _msp_evm_progress_data_quality

MSP_XML = os.path.join(os.path.dirname(__file__), "fixtures", "sample_msp.xml")


def test_msp_evm_pdq_xml():
    r = _msp_evm_progress_data_quality(file_path=MSP_XML)
    assert r["status"] == "ok"
    assert "warnings" in r
    assert isinstance(r["warnings"], list)
