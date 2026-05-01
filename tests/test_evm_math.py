"""Test pure-math EVM algorithms (RULE 4-9). No fixtures, no COM, no MSP."""
import pytest
from evm_math import compute_metrics, forecast, rag_status


# ---------- compute_metrics (RULE 4) ----------

def test_compute_metrics_happy():
    r = compute_metrics(bac=1000, pv=500, ev=400, ac=450)
    assert r["spi"] == pytest.approx(0.8, rel=1e-3)
    assert r["cpi"] == pytest.approx(0.889, rel=1e-3)
    assert r["sv"] == -100  # ev - pv
    assert r["cv"] == -50   # ev - ac


def test_compute_metrics_pv_zero_returns_none_spi():
    r = compute_metrics(bac=1000, pv=0, ev=0, ac=0)
    assert r["spi"] is None
    assert r["cpi"] is None


def test_compute_metrics_passes_inputs_through():
    r = compute_metrics(bac=1000, pv=500, ev=400, ac=450)
    assert r["bac"] == 1000
    assert r["pv"] == 500
    assert r["ev"] == 400
    assert r["ac"] == 450


# ---------- forecast (RULE 9) ----------

def test_forecast_eac_t1_variance_one_time():
    """EAC1 = AC + (BAC - EV)."""
    r = forecast(bac=1000, ev=400, ac=450, cpi=0.889, spi=0.8)
    assert r["eac_t1"] == pytest.approx(1050.0, rel=1e-3)


def test_forecast_eac_t2_current_performance():
    """EAC2 = BAC / CPI."""
    r = forecast(bac=1000, ev=400, ac=450, cpi=0.889, spi=0.8)
    assert r["eac_t2"] == pytest.approx(1124.86, rel=1e-3)


def test_forecast_eac_t3_combined():
    """EAC3 = AC + (BAC - EV) / (CPI * SPI)."""
    r = forecast(bac=1000, ev=400, ac=450, cpi=0.889, spi=0.8)
    expected = 450 + (1000 - 400) / (0.889 * 0.8)
    assert r["eac_t3"] == pytest.approx(expected, rel=1e-3)


def test_forecast_etc_and_vac():
    r = forecast(bac=1000, ev=400, ac=450, cpi=0.889, spi=0.8)
    assert r["etc"] == pytest.approx(r["eac_t3"] - 450, rel=1e-3)
    assert r["vac"] == pytest.approx(1000 - r["eac_t3"], rel=1e-3)


def test_forecast_tcpi_bac():
    """TCPI(BAC) = (BAC - EV) / (BAC - AC)."""
    r = forecast(bac=1000, ev=400, ac=450, cpi=0.889, spi=0.8)
    expected = (1000 - 400) / (1000 - 450)
    assert r["tcpi_bac"] == pytest.approx(expected, rel=1e-3)


def test_forecast_cpi_zero_returns_none_eac_t2():
    r = forecast(bac=1000, ev=0, ac=0, cpi=0, spi=0)
    assert r["eac_t2"] is None
    assert r["eac_t3"] is None


# ---------- rag_status (RULE 12) ----------

def test_rag_status_green_above_07():
    assert rag_status(spi=0.85, completion_pct=50) == "GREEN"


def test_rag_status_amber_between():
    assert rag_status(spi=0.50, completion_pct=30) == "AMBER"


def test_rag_status_red_below_03():
    assert rag_status(spi=0.20, completion_pct=10) == "RED"


def test_rag_status_red_zero_progress():
    assert rag_status(spi=0.85, completion_pct=0) == "RED"


def test_rag_status_red_spi_none():
    assert rag_status(spi=None, completion_pct=10) == "RED"


import datetime as dt
from evm_math import time_phased_pv, time_phased_ev, period_delta


# ---------- time_phased_pv (RULE 5) ----------

def _make_task(name, bs, bf, bw):
    return {"name": name, "baseline_start": bs, "baseline_finish": bf, "baseline_work": bw}


def test_time_phased_pv_task_fully_within_bucket():
    """Task baseline completed before bucket end -> full BW counts."""
    tasks = [_make_task("T1", dt.date(2026, 1, 1), dt.date(2026, 1, 10), 80.0)]
    buckets = [(dt.date(2026, 1, 1), dt.date(2026, 1, 31))]
    pv = time_phased_pv(tasks, buckets)
    assert pv == [80.0]


def test_time_phased_pv_task_not_started():
    """Task baseline_start after bucket end -> 0 PV."""
    tasks = [_make_task("T2", dt.date(2026, 2, 1), dt.date(2026, 2, 10), 80.0)]
    buckets = [(dt.date(2026, 1, 1), dt.date(2026, 1, 31))]
    pv = time_phased_pv(tasks, buckets)
    assert pv == [0.0]


def test_time_phased_pv_task_partial():
    """Task spans bucket end - linear distribution."""
    # T3: 10 day baseline (Jan 5 - Jan 15), 100h. Bucket ends Jan 10 = 50% way -> 50h.
    tasks = [_make_task("T3", dt.date(2026, 1, 5), dt.date(2026, 1, 15), 100.0)]
    buckets = [(dt.date(2026, 1, 1), dt.date(2026, 1, 10))]
    pv = time_phased_pv(tasks, buckets)
    assert pv[0] == pytest.approx(50.0, rel=1e-2)


def test_time_phased_pv_multi_bucket_cumulative():
    """PV per bucket is cumulative as of bucket end (not per-period delta)."""
    tasks = [_make_task("T4", dt.date(2026, 1, 1), dt.date(2026, 1, 30), 300.0)]
    buckets = [
        (dt.date(2026, 1, 1), dt.date(2026, 1, 10)),
        (dt.date(2026, 1, 1), dt.date(2026, 1, 20)),
        (dt.date(2026, 1, 1), dt.date(2026, 1, 31)),
    ]
    pv = time_phased_pv(tasks, buckets)
    # Linear: ~10/29 days = 34.5%, ~20/29 = 69%, finished
    assert pv[0] == pytest.approx(300 * 9 / 29, rel=1e-2)
    assert pv[1] == pytest.approx(300 * 19 / 29, rel=1e-2)
    assert pv[2] == 300.0


# ---------- time_phased_ev ----------

def _make_task_ev(name, bs, bf, bw, pct):
    return {"name": name, "baseline_start": bs, "baseline_finish": bf,
            "baseline_work": bw, "percent_complete": pct}


def test_time_phased_ev_capped_at_data_date():
    """EV future buckets repeat current EV (capped at data_date)."""
    tasks = [_make_task_ev("T1", dt.date(2026, 1, 1), dt.date(2026, 1, 30), 100.0, 50)]
    data_date = dt.date(2026, 1, 15)
    buckets = [
        (dt.date(2026, 1, 1), dt.date(2026, 1, 10)),
        (dt.date(2026, 1, 1), dt.date(2026, 1, 20)),
        (dt.date(2026, 1, 1), dt.date(2026, 1, 31)),
    ]
    ev = time_phased_ev(tasks, buckets, data_date)
    # All 3 buckets see the same EV = 100 * 50% = 50 because data_date caps any bucket past it
    assert ev[0] == pytest.approx(50.0, rel=1e-2)
    assert ev[1] == pytest.approx(50.0, rel=1e-2)
    assert ev[2] == pytest.approx(50.0, rel=1e-2)


def test_time_phased_ev_excludes_unstarted_tasks():
    tasks = [_make_task_ev("T1", dt.date(2026, 2, 1), dt.date(2026, 2, 10), 100.0, 50)]
    buckets = [(dt.date(2026, 1, 1), dt.date(2026, 1, 31))]
    ev = time_phased_ev(tasks, buckets, data_date=dt.date(2026, 1, 31))
    assert ev[0] == 0.0  # task hasn't baseline-started yet


# ---------- period_delta (RULE 6) ----------

def test_period_delta_basic():
    snap_now = {"pv": 1000, "ev": 800, "ac": 850}
    snap_prev = {"pv": 700, "ev": 500, "ac": 550}
    r = period_delta(snap_now, snap_prev)
    assert r["period_pv"] == 300
    assert r["period_ev"] == 300
    assert r["period_ac"] == 300
    assert r["period_bac"] == 0  # BAC sabit


def test_period_delta_first_period():
    """No prev snapshot -> period values = current cum values."""
    snap_now = {"pv": 1000, "ev": 800, "ac": 850}
    r = period_delta(snap_now, None)
    assert r["period_pv"] == 1000
    assert r["period_ev"] == 800
    assert r["period_ac"] == 850


from evm_math import earned_schedule, progress_data_quality


# ---------- earned_schedule (RULE 8 Lipke 2003) ----------

def test_earned_schedule_on_track():
    """ev_now exactly matches a curve point -> es == that t."""
    project_start = dt.date(2026, 1, 1)
    pv_curve = [
        (dt.date(2026, 1, 8),  100.0),
        (dt.date(2026, 1, 15), 200.0),
        (dt.date(2026, 1, 22), 300.0),
    ]
    data_date = dt.date(2026, 1, 22)  # AT = 3 weeks
    r = earned_schedule(pv_curve, ev_now=200.0, project_start=project_start,
                        data_date=data_date)
    assert r["at"] == pytest.approx(3.0, rel=1e-2)
    assert r["es"] == pytest.approx(2.0, rel=1e-2)
    assert r["sv_t"] == pytest.approx(-1.0, rel=1e-2)
    assert r["spi_t"] == pytest.approx(2.0 / 3.0, rel=1e-2)


def test_earned_schedule_interpolation():
    """ev_now between two points -> linear interp."""
    project_start = dt.date(2026, 1, 1)
    pv_curve = [
        (dt.date(2026, 1, 8),  100.0),  # 1 week
        (dt.date(2026, 1, 15), 200.0),  # 2 weeks
    ]
    data_date = dt.date(2026, 1, 15)  # AT = 2 weeks
    # ev=150 -> halfway between 100 and 200 -> es = 1.5 weeks
    r = earned_schedule(pv_curve, ev_now=150.0, project_start=project_start,
                        data_date=data_date)
    assert r["es"] == pytest.approx(1.5, rel=1e-2)


def test_earned_schedule_ev_below_first_point():
    """ev_now < first PV -> es ~= 0 (clamped)."""
    project_start = dt.date(2026, 1, 1)
    pv_curve = [(dt.date(2026, 1, 8), 100.0), (dt.date(2026, 1, 15), 200.0)]
    r = earned_schedule(pv_curve, ev_now=10.0, project_start=project_start,
                        data_date=dt.date(2026, 1, 15))
    assert r["es"] is not None
    assert r["es"] >= 0.0


def test_earned_schedule_ev_above_last_point():
    """ev_now > last PV -> es clamped to last point."""
    project_start = dt.date(2026, 1, 1)
    pv_curve = [(dt.date(2026, 1, 8), 100.0), (dt.date(2026, 1, 15), 200.0)]
    r = earned_schedule(pv_curve, ev_now=999.0, project_start=project_start,
                        data_date=dt.date(2026, 1, 15))
    # es should be at least 2 weeks (last curve point)
    assert r["es"] >= 2.0


# ---------- progress_data_quality (RULE 7) ----------

def test_pdq_no_warnings_on_clean_data():
    warnings = progress_data_quality(spi_h=0.85, spi_t=0.83,
                                    completion_pct=50, has_resources=True)
    assert warnings == []


def test_pdq_spi_divergence_warning():
    """abs(spi_h - spi_t) > 0.15 -> warning."""
    warnings = progress_data_quality(spi_h=0.85, spi_t=0.50,
                                    completion_pct=50, has_resources=True)
    assert any("divergence" in w["warning"].lower() or "ev input" in w["warning"].lower()
               for w in warnings)


def test_pdq_no_resources_warning():
    """has_resources=False with progress entered -> silent EV pattern."""
    warnings = progress_data_quality(spi_h=0.85, spi_t=0.85,
                                    completion_pct=50, has_resources=False)
    assert any("resource" in w["warning"].lower() for w in warnings)
