"""Test msproject_calendar holidays_uzbek action."""
import pytest
import time
from msproject_mcp_core import (
    _msp_calendar_create, _msp_calendar_holidays_uzbek,
    _find_calendar_by_name, UZBEK_HOLIDAYS_2026,
)


def test_uzbek_holidays_added(clean_test_project):
    """All 9 Uzbek holidays added to a fresh calendar in <2s."""
    proj = clean_test_project
    _msp_calendar_create(name="UzbekCal-Phase2a", base_calendar="Standard")
    start = time.time()
    r = _msp_calendar_holidays_uzbek(calendar_name="UzbekCal-Phase2a", year=2026)
    elapsed = time.time() - start
    assert r["status"] == "ok"
    assert r["count"] == 9
    assert elapsed < 2.0, f"holidays_uzbek took {elapsed:.2f}s (target <2s)"
    cal = _find_calendar_by_name(proj, "UzbekCal-Phase2a")
    assert cal.Exceptions.Count == 9


def test_uzbek_holidays_dates_correct(clean_test_project):
    """Returned holiday dates match UZBEK_HOLIDAYS_2026 constant."""
    _msp_calendar_create(name="UzbekDateCal-Phase2a", base_calendar="Standard")
    r = _msp_calendar_holidays_uzbek(calendar_name="UzbekDateCal-Phase2a", year=2026)
    assert r["status"] == "ok"
    returned_dates = {(h["month"], h["day"]) for h in r["holidays"]}
    expected_dates = {(m, d) for _, m, d in UZBEK_HOLIDAYS_2026}
    assert returned_dates == expected_dates


def test_uzbek_holidays_missing_calendar_errors(clean_test_project):
    r = _msp_calendar_holidays_uzbek(calendar_name="NoSuch-Phase2a", year=2026)
    assert r["status"] == "error"
    assert "not found" in r["error"].lower()


def test_uzbek_holidays_idempotent(clean_test_project):
    """Calling twice on same calendar skips already-added (name-based dedup)."""
    proj = clean_test_project
    _msp_calendar_create(name="IdempotentCal-Phase2a", base_calendar="Standard")
    r1 = _msp_calendar_holidays_uzbek(calendar_name="IdempotentCal-Phase2a", year=2026)
    assert r1["status"] == "ok"
    assert r1["count"] == 9
    cal = _find_calendar_by_name(proj, "IdempotentCal-Phase2a")
    count_after_first = cal.Exceptions.Count
    assert count_after_first >= 9

    # Second call — all 9 should be skipped, no new exceptions added
    r2 = _msp_calendar_holidays_uzbek(calendar_name="IdempotentCal-Phase2a", year=2026)
    assert r2["status"] == "already_done"
    assert r2["count"] == 0
    assert r2["skipped_count"] == 9
    assert len(r2["skipped"]) == 9
    assert all("already exists" in s["reason"] for s in r2["skipped"])
    # Exception count unchanged
    assert cal.Exceptions.Count == count_after_first


def test_uzbek_holidays_partial_failure(clean_test_project, monkeypatch):
    """If add_exception fails for some holidays, status='partial' with failures listed."""
    import msproject_mcp_core as core
    _msp_calendar_create(name="PartialCal-Phase2a", base_calendar="Standard")

    real_add = core._msp_calendar_add_exception
    call_count = {"n": 0}
    def flaky_add(*args, **kwargs):
        call_count["n"] += 1
        # Fail on the 5th holiday (1 May - İşçi Bayramı)
        if call_count["n"] == 5:
            return {"status": "error", "error": "simulated COM failure"}
        return real_add(*args, **kwargs)
    monkeypatch.setattr(core, "_msp_calendar_add_exception", flaky_add)

    r = _msp_calendar_holidays_uzbek(calendar_name="PartialCal-Phase2a", year=2026)
    assert r["status"] == "partial"
    assert r["count"] == 8
    assert len(r["failures"]) == 1
    assert "simulated COM failure" in r["failures"][0]["error"]
