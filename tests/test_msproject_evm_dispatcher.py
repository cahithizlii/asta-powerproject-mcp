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


# =============================================================================
# Phase 11.2 — Edge Case + Negative Path tests (T142)
# =============================================================================


def test_dispatcher_compute_metrics_missing_file_returns_error(tmp_path):
    """Missing source file → error from underlying loader."""
    p = _call("compute_metrics", file_path=str(tmp_path / "no_such.xml"))
    assert p["status"] == "error"
    assert p["error"]


def test_dispatcher_forecast_missing_file_returns_error(tmp_path):
    """Forecast with missing file → error."""
    p = _call("forecast", file_path=str(tmp_path / "ghost.xml"))
    assert p["status"] == "error"


def test_dispatcher_time_phased_evm_invalid_bucket_returns_error():
    """time_phased_evm with bucket='quarterly' (not day/week/month) → error."""
    p = _call("time_phased_evm", file_path=MSP_XML, bucket="quarterly")
    assert p["status"] == "error"
    assert "bucket" in p["error"].lower()


def test_dispatcher_time_phased_evm_invalid_bucket_string_returns_error():
    """time_phased_evm with arbitrary bucket string → error mentioning valid set."""
    p = _call("time_phased_evm", file_path=MSP_XML, bucket="zzz")
    assert p["status"] == "error"
    assert "day" in p["error"].lower() or "week" in p["error"].lower()


def test_dispatcher_compute_metrics_invalid_baseline_negative_returns_error():
    """baseline_number=-1 → error (must be 0-10)."""
    p = _call("compute_metrics", file_path=MSP_XML, baseline_number=-1)
    assert p["status"] == "error"
    assert "baseline" in p["error"].lower()


def test_dispatcher_compute_metrics_invalid_baseline_too_high_returns_error():
    """baseline_number=99 → error."""
    p = _call("compute_metrics", file_path=MSP_XML, baseline_number=99)
    assert p["status"] == "error"
    assert "baseline" in p["error"].lower()


def test_dispatcher_unknown_action_lists_valid_actions():
    """Unknown action error message lists valid actions."""
    p = _call("totally_made_up_action", file_path=MSP_XML)
    assert p["status"] == "error"
    assert "compute_metrics" in p["error"]


def test_dispatcher_compute_metrics_unsupported_extension_returns_error(tmp_path):
    """Unsupported file extension → error from loader."""
    bad = tmp_path / "x.docx"
    bad.write_text("not a project")
    p = _call("compute_metrics", file_path=str(bad))
    assert p["status"] == "error"


def test_dispatcher_validate_currency_mode_missing_file_returns_error(tmp_path):
    """validate_currency_mode with bad path → error."""
    p = _call("validate_currency_mode", file_path=str(tmp_path / "no.xml"))
    assert p["status"] == "error"


def test_dispatcher_period_delta_invalid_baseline_returns_error():
    """period_delta with baseline_number out of range → error."""
    p = _call("period_delta", file_path=MSP_XML, baseline_number=15)
    assert p["status"] == "error"
    assert "baseline" in p["error"].lower()


def test_dispatcher_earned_schedule_missing_file_returns_error(tmp_path):
    """earned_schedule with non-existent path → error."""
    p = _call("earned_schedule", file_path=str(tmp_path / "missing.xml"))
    assert p["status"] == "error"


def test_dispatcher_time_phased_evm_unsupported_extension_returns_error(tmp_path):
    """time_phased_evm with unsupported file → error."""
    bad = tmp_path / "x.docx"
    bad.write_text("not a project")
    p = _call("time_phased_evm", file_path=str(bad), bucket="week")
    assert p["status"] == "error"
