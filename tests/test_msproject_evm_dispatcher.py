import asyncio, json, os, sys
REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)
from msproject_mcp_core import msproject_evm

MSP_XML = os.path.join(os.path.dirname(__file__), "fixtures", "sample_msp.xml")


def _run(coro):
    return asyncio.run(coro)


def _call(action, **kw):
    raw = _run(msproject_evm({"action": action, **kw}))
    return json.loads(raw)


def test_dispatcher_compute_metrics():
    p = _call("compute_metrics", file_path=MSP_XML)
    assert p["status"] == "ok"
    assert "spi" in p


def test_dispatcher_forecast():
    p = _call("forecast", file_path=MSP_XML)
    assert p["status"] == "ok"


def test_dispatcher_summary():
    p = _call("summary", file_path=MSP_XML)
    assert p["status"] == "ok"
    assert p["rag"] in ("RED", "AMBER", "GREEN")


def test_dispatcher_unknown_action():
    p = _call("nonsense", file_path=MSP_XML)
    assert p["status"] == "error"
    assert "Unknown action" in p["error"]


def test_dispatcher_detect_currency_mode():
    p = _call("detect_currency_mode", file_path=MSP_XML)
    assert p["status"] == "ok"
    assert p["mode"] in ("hours", "cost")
