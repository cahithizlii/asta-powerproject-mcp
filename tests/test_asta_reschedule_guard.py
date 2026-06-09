"""P2 #8 — POLYBLMH reschedule risk guard (RULE 17). Pure, no COM."""
import pytest
from asta_reschedule_guard import assess_reschedule_risk, _parse


def test_polyblmh_exact_signature_high_risk():
    """ReportDate 187d after start + 0% progress = high risk."""
    r = assess_reschedule_risk("2026-02-28", "2025-08-25",
                               project_percent_complete=0.0)
    assert r["risk"] is True
    assert r["severity"] == "high"
    assert r["gap_days"] == 187
    assert "POLYBLMH" in r["message"]
    assert r["recommendation"] is not None


def test_late_report_date_unknown_progress_low_advisory():
    r = assess_reschedule_risk("2026-02-28", "2025-08-25",
                               project_percent_complete=None)
    assert r["risk"] is False           # cannot confirm -> not high
    assert r["severity"] == "low"
    assert r["message"] is not None


def test_late_report_date_with_progress_no_risk():
    """Normal in-progress project: late ReportDate is fine if work done."""
    r = assess_reschedule_risk("2026-02-28", "2025-08-25",
                               project_percent_complete=45.0)
    assert r["risk"] is False
    assert r["severity"] == "none"
    assert r["message"] is None


def test_report_date_near_start_no_warning():
    r = assess_reschedule_risk("2025-09-01", "2025-08-25",
                               project_percent_complete=0.0)
    assert r["gap_days"] == 7
    assert r["severity"] == "none"
    assert r["message"] is None


def test_gap_threshold_boundary():
    # exactly at threshold (30) -> no warning; 31 -> warning
    assert assess_reschedule_risk("2025-09-24", "2025-08-25", 0.0)["severity"] == "none"
    assert assess_reschedule_risk("2025-09-25", "2025-08-25", 0.0)["severity"] == "high"


def test_unparseable_dates_safe():
    r = assess_reschedule_risk(None, "2025-08-25", 0.0)
    assert r["risk"] is False
    assert r["gap_days"] is None
    r2 = assess_reschedule_risk("not-a-date", "also-bad", 0.0)
    assert r2["gap_days"] is None


def test_parse_multiple_formats():
    assert _parse("2025-08-25").isoformat() == "2025-08-25"
    assert _parse("25/08/2025").isoformat() == "2025-08-25"
    assert _parse(None) is None


def test_custom_gap_warn_days():
    r = assess_reschedule_risk("2025-09-10", "2025-08-25", 0.0,
                               gap_warn_days=5)
    assert r["severity"] == "high"      # 16 > 5
