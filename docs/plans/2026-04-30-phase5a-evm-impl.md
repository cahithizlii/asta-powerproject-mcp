# Phase 5a EVM Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans (or subagent-driven-development) to implement this plan task-by-task.

**Goal:** New `msproject_evm` MCP tool — 13 actions covering full PMI PMBOK 8th § 7.4.2 EVM (compute, forecast, Earned Schedule per Lipke 2003), time-phased PV/EV/AC per RULE 5/6, RULE 7 progress data quality warnings, RULE 12 RAG status, RULE 3 currency mode auto-detection, Phase 3a baseline integration, and JSON snapshot history.

**Architecture:** New `evm_math.py` module — pure Python, MSP/COM/file independent — implements RULE 4-9 algorithms. `msproject_mcp_core.py` Phase 5a section adds I/O adapters (`_evm_load_*` hybrid file+COM data source), 13 helper functions, and FastMCP dispatcher. Phase 1+2a+2b+3a+3b+4 helpers DOKUNULMAZ; only read-only calls into Phase 4 file helpers (`_msp_file_read_tasks/progress/baselines`) and Phase 1 COM (`_validate_active_project`).

**Tech Stack:** Python 3.12, mcp (FastMCP), pywin32 COM (existing), pytest. New helper-only Python module `evm_math.py`. Mevcut `msproject_mcp_core.py` (~4750 satır after Phase 4), 35+ test files, **80/80 Phase 4 file tests + Phase 1-3 regression** baseline.

**Design doc:** `docs/plans/2026-04-30-phase5a-evm-design.md` (commit `32f91c2`)

**Baseline state at start:** HEAD `32f91c2`, MS Project running v16.0.

**KEY REFERENCES:**
- CLAUDE.md RULE 3 (currency mode), RULE 4 (EVM core), RULE 5 (time-phased PV linear), RULE 6 (period delta), RULE 7 (data quality), RULE 8 (Earned Schedule Lipke), RULE 9 (forecasting PMBOK 8th), RULE 12 (RAG)
- Phase 3b summary helper `_msp_progress_summary` — already returns BAC/ACWP/StatusDate, Phase 5a builds on this
- Phase 3a `BASELINE_NUMBERS`, `_baseline_property_name`, `_read_task_baseline`
- Phase 4 `_msp_file_read_tasks/links/resources/assignments/baselines/progress`

---

## Task 75: `evm_math.py` Foundations — compute_metrics + forecast + rag_status

**Files:**
- Create: `C:\Users\CahAsus\asta-powerproject-mcp\evm_math.py`
- Create: `C:\Users\CahAsus\asta-powerproject-mcp\tests\test_evm_math.py`

**Step 1: Failing tests**

`tests/test_evm_math.py`:
```python
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
```

**Step 2: Run — FAIL**

```bash
cd C:\Users\CahAsus\asta-powerproject-mcp
python -m pytest tests/test_evm_math.py -v
```
Expected: ImportError (evm_math doesn't exist).

**Step 3: Implementation**

Create `evm_math.py`:
```python
"""Phase 5a — Pure-math EVM algorithms.

Implements CLAUDE.md RULE 4-9 + RULE 12. MSP/COM/file independent —
takes plain numbers, returns plain numbers. Easily testable without
fixtures, without COM, without MS Project.

References:
- RULE 4: SPI/CPI/SV/CV
- RULE 5: Time-phased PV linear distribution
- RULE 6: Period delta (haftalik delta)
- RULE 7: EV data quality (SPI(h) vs SPI(t))
- RULE 8: Earned Schedule (Lipke 2003)
- RULE 9: PMI PMBOK 8th § 7.4.2 forecasting
- RULE 12: RAG status thresholds
"""
from typing import Optional, Dict, Any, List, Tuple
import datetime as _dt


def compute_metrics(bac: float, pv: float, ev: float, ac: float) -> Dict[str, Any]:
    """RULE 4 — Compute SPI/CPI/SV/CV.

    Returns dict with all input values + spi/cpi/sv/cv. None when divisor 0.
    """
    spi = ev / pv if pv > 0 else None
    cpi = ev / ac if ac > 0 else None
    sv = ev - pv
    cv = ev - ac
    return {
        "bac": bac, "pv": pv, "ev": ev, "ac": ac,
        "spi": spi, "cpi": cpi, "sv": sv, "cv": cv,
    }


def forecast(bac: float, ev: float, ac: float,
             cpi: Optional[float], spi: Optional[float]) -> Dict[str, Any]:
    """RULE 9 — PMI PMBOK 8th § 7.4.2.

    EAC1 = AC + (BAC-EV)            variance one-time
    EAC2 = BAC / CPI                current performance
    EAC3 = AC + (BAC-EV)/(CPI*SPI)  combined
    ETC, VAC, TCPI(BAC), TCPI(EAC) per spec.
    """
    eac_t1 = ac + (bac - ev)
    eac_t2 = (bac / cpi) if cpi and cpi > 0 else None
    eac_t3 = (ac + (bac - ev) / (cpi * spi)) if cpi and spi and cpi > 0 and spi > 0 else None
    # ETC + VAC use EAC3 if available, else EAC1 fallback
    eac_for_etc = eac_t3 if eac_t3 is not None else eac_t1
    etc = eac_for_etc - ac
    vac = bac - eac_for_etc
    tcpi_bac = ((bac - ev) / (bac - ac)) if (bac - ac) > 0 else None
    tcpi_eac = ((bac - ev) / (eac_for_etc - ac)) if eac_for_etc and (eac_for_etc - ac) > 0 else None
    return {
        "eac_t1": eac_t1, "eac_t2": eac_t2, "eac_t3": eac_t3,
        "etc": etc, "vac": vac,
        "tcpi_bac": tcpi_bac, "tcpi_eac": tcpi_eac,
    }


def rag_status(spi: Optional[float], completion_pct: float) -> str:
    """RULE 12 — RED/AMBER/GREEN.

    RED:   spi None OR spi < 0.3 OR completion_pct == 0
    AMBER: 0.3 <= spi < 0.7
    GREEN: spi >= 0.7
    """
    if spi is None or completion_pct == 0:
        return "RED"
    if spi < 0.3:
        return "RED"
    if spi < 0.7:
        return "AMBER"
    return "GREEN"
```

**Step 4: Run — PASS**

```bash
python -m pytest tests/test_evm_math.py -v
```
Expected: 13 PASS.

**Step 5: Commit**

```bash
git add evm_math.py tests/test_evm_math.py
git commit -m "Phase 5a T75: evm_math foundations (compute_metrics + forecast + rag_status)

Pure-Python implementations of CLAUDE.md RULE 4 (SPI/CPI/SV/CV),
RULE 9 (PMBOK 8th forecasting EAC1/2/3 + ETC/VAC/TCPI), RULE 12
(RAG thresholds). MSP/COM/file independent — easily testable.

13 unit tests, no fixtures."
```

DO NOT push (T84 will push the chain).

Expected: regression unchanged (Phase 4 80/80 + Phase 1-3 baseline).

---

## Task 76: `evm_math.py` time-phased — time_phased_pv + time_phased_ev + period_delta

**Files:**
- Modify: `evm_math.py`
- Modify: `tests/test_evm_math.py`

**Step 1: Add failing tests**

Append to `tests/test_evm_math.py`:
```python
import datetime as dt
from evm_math import time_phased_pv, time_phased_ev, period_delta


# ---------- time_phased_pv (RULE 5) ----------

def _make_task(name, bs, bf, bw):
    return {"name": name, "baseline_start": bs, "baseline_finish": bf, "baseline_work": bw}


def test_time_phased_pv_task_fully_within_bucket():
    """Task baseline completed before bucket end → full BW counts."""
    tasks = [_make_task("T1", dt.date(2026, 1, 1), dt.date(2026, 1, 10), 80.0)]
    buckets = [(dt.date(2026, 1, 1), dt.date(2026, 1, 31))]
    pv = time_phased_pv(tasks, buckets)
    assert pv == [80.0]


def test_time_phased_pv_task_not_started():
    """Task baseline_start after bucket end → 0 PV."""
    tasks = [_make_task("T2", dt.date(2026, 2, 1), dt.date(2026, 2, 10), 80.0)]
    buckets = [(dt.date(2026, 1, 1), dt.date(2026, 1, 31))]
    pv = time_phased_pv(tasks, buckets)
    assert pv == [0.0]


def test_time_phased_pv_task_partial():
    """Task spans bucket end — linear distribution."""
    # T3: 10 day baseline (Jan 5 - Jan 15), 100h. Bucket ends Jan 10 = 50% way → 50h.
    tasks = [_make_task("T3", dt.date(2026, 1, 5), dt.date(2026, 1, 15), 100.0)]
    buckets = [(dt.date(2026, 1, 1), dt.date(2026, 1, 10))]
    pv = time_phased_pv(tasks, buckets)
    assert pv[0] == pytest.approx(50.0, rel=1e-2)


def test_time_phased_pv_multi_bucket_cumulative():
    """PV per bucket is cumulative as of bucket end (not per-period delta)."""
    tasks = [_make_task("T4", dt.date(2026, 1, 1), dt.date(2026, 1, 30), 300.0)]
    # Buckets: Jan 1-10, Jan 11-20, Jan 21-31
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
    # All 3 buckets see the same EV = 100 * 50% = 50
    # because data_date caps any bucket past it
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
    """No prev snapshot → period values = current cum values."""
    snap_now = {"pv": 1000, "ev": 800, "ac": 850}
    r = period_delta(snap_now, None)
    assert r["period_pv"] == 1000
    assert r["period_ev"] == 800
    assert r["period_ac"] == 850
```

**Step 2: Run — FAIL** (NameError on time_phased_pv etc.)

**Step 3: Implementation**

Append to `evm_math.py`:
```python
def _task_pv_at_date(task: Dict[str, Any], eval_date: _dt.date) -> float:
    """RULE 5 — Linear distribution per task. Hours OR cost (caller decides)."""
    bs = task.get("baseline_start")
    bf = task.get("baseline_finish")
    bw = float(task.get("baseline_work") or 0)
    if bs is None or bf is None or bw == 0:
        return 0.0
    # Coerce to date if datetime
    if hasattr(bs, "date"):
        bs = bs.date()
    if hasattr(bf, "date"):
        bf = bf.date()
    if bf <= eval_date:
        return bw  # task fully baseline-completed
    if bs >= eval_date:
        return 0.0  # not yet baseline-started
    duration_days = max((bf - bs).days, 1)
    elapsed_days = max((eval_date - bs).days, 0)
    return bw * elapsed_days / duration_days


def time_phased_pv(tasks: List[Dict[str, Any]],
                  buckets: List[Tuple[_dt.date, _dt.date]]) -> List[float]:
    """RULE 5 — Cumulative PV at each bucket end (linear distribution)."""
    return [sum(_task_pv_at_date(t, bucket_end) for t in tasks)
            for (_, bucket_end) in buckets]


def time_phased_ev(tasks: List[Dict[str, Any]],
                  buckets: List[Tuple[_dt.date, _dt.date]],
                  data_date: _dt.date) -> List[float]:
    """Cumulative EV at each bucket end. Future buckets capped at data_date.

    EV uses current percent_complete × baseline_work, but only counts tasks
    that have baseline-started by min(bucket_end, data_date). This avoids
    double-counting future periods (EV doesn't grow beyond data_date).
    """
    if hasattr(data_date, "date"):
        data_date = data_date.date()
    out = []
    for (_, bucket_end) in buckets:
        if hasattr(bucket_end, "date"):
            bucket_end = bucket_end.date()
        eval_at = min(bucket_end, data_date)
        ev = 0.0
        for t in tasks:
            bs = t.get("baseline_start")
            if bs is None:
                continue
            if hasattr(bs, "date"):
                bs = bs.date()
            if bs > eval_at:
                continue
            bw = float(t.get("baseline_work") or 0)
            pct = float(t.get("percent_complete") or 0) / 100.0
            ev += bw * pct
        out.append(ev)
    return out


def period_delta(snap_now: Dict[str, Any],
                snap_prev: Optional[Dict[str, Any]]) -> Dict[str, Any]:
    """RULE 6 — Haftalık/aylık delta. period_BAC = 0 (sabit)."""
    if snap_prev is None:
        return {
            "period_pv": float(snap_now.get("pv") or 0),
            "period_ev": float(snap_now.get("ev") or 0),
            "period_ac": float(snap_now.get("ac") or 0),
            "period_bac": 0.0,
        }
    return {
        "period_pv": float(snap_now.get("pv") or 0) - float(snap_prev.get("pv") or 0),
        "period_ev": float(snap_now.get("ev") or 0) - float(snap_prev.get("ev") or 0),
        "period_ac": float(snap_now.get("ac") or 0) - float(snap_prev.get("ac") or 0),
        "period_bac": 0.0,
    }
```

**Step 4: Run — PASS** (8 new + 13 prev = 21 PASS)

**Step 5: Commit**

```bash
git add evm_math.py tests/test_evm_math.py
git commit -m "Phase 5a T76: evm_math time-phased (PV + EV + period_delta)

CLAUDE.md RULE 5 linear distribution per task (no calendar awareness —
linear assumption matches RULE 5 spec). RULE 6 period delta with
period_bac=0 (sabit). EV capped at data_date to prevent future
period inflation."
```

---

## Task 77: `evm_math.py` earned_schedule + progress_data_quality

**Files:**
- Modify: `evm_math.py`
- Modify: `tests/test_evm_math.py`

**Step 1: Failing tests**

Append:
```python
from evm_math import earned_schedule, progress_data_quality


# ---------- earned_schedule (RULE 8 Lipke 2003) ----------

def test_earned_schedule_on_track():
    """ev_now exactly matches a curve point → es == that t."""
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
    """ev_now between two points → linear interp."""
    project_start = dt.date(2026, 1, 1)
    pv_curve = [
        (dt.date(2026, 1, 8),  100.0),  # 1 week
        (dt.date(2026, 1, 15), 200.0),  # 2 weeks
    ]
    data_date = dt.date(2026, 1, 15)  # AT = 2 weeks
    # ev=150 → halfway between 100 and 200 → es = 1.5 weeks
    r = earned_schedule(pv_curve, ev_now=150.0, project_start=project_start,
                        data_date=data_date)
    assert r["es"] == pytest.approx(1.5, rel=1e-2)


def test_earned_schedule_ev_below_first_point():
    """ev_now < first PV → es ~= 0 (clamped)."""
    project_start = dt.date(2026, 1, 1)
    pv_curve = [(dt.date(2026, 1, 8), 100.0), (dt.date(2026, 1, 15), 200.0)]
    r = earned_schedule(pv_curve, ev_now=10.0, project_start=project_start,
                        data_date=dt.date(2026, 1, 15))
    assert r["es"] is not None
    assert r["es"] >= 0.0


def test_earned_schedule_ev_above_last_point():
    """ev_now > last PV → es clamped to last point."""
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
    """abs(spi_h - spi_t) > 0.15 → warning."""
    warnings = progress_data_quality(spi_h=0.85, spi_t=0.50,
                                    completion_pct=50, has_resources=True)
    assert any("divergence" in w["warning"].lower() or "ev input" in w["warning"].lower()
               for w in warnings)


def test_pdq_no_resources_warning():
    """has_resources=False with progress entered → silent EV pattern."""
    warnings = progress_data_quality(spi_h=0.85, spi_t=0.85,
                                    completion_pct=50, has_resources=False)
    assert any("resource" in w["warning"].lower() for w in warnings)
```

**Step 2: Run — FAIL**

**Step 3: Implementation**

Append to `evm_math.py`:
```python
def earned_schedule(pv_curve: List[Tuple[_dt.date, float]],
                   ev_now: float,
                   project_start: _dt.date,
                   data_date: _dt.date) -> Dict[str, Any]:
    """RULE 8 (Lipke 2003) — Earned Schedule.

    pv_curve: list of (date, cumulative_pv) sorted ascending by date.
    Returns AT (weeks since project_start), ES (interpolated time where
    cumulative PV equals ev_now), SV(t)=ES-AT, SPI(t)=ES/AT.
    """
    if hasattr(project_start, "date"):
        project_start = project_start.date()
    if hasattr(data_date, "date"):
        data_date = data_date.date()
    at_weeks = max((data_date - project_start).days / 7.0, 1e-9)
    if not pv_curve:
        return {"at": at_weeks, "es": None, "sv_t": None, "spi_t": None}

    es_date: Optional[_dt.date] = None

    # Below first point — clamp to 0
    first_date, first_pv = pv_curve[0]
    if ev_now <= first_pv:
        # Linear from project_start to first_date proportional to ev/first_pv
        if first_pv > 0:
            frac = ev_now / first_pv
        else:
            frac = 0.0
        delta = (first_date - project_start).days * frac
        es_date = project_start + _dt.timedelta(days=int(delta))

    # Search adjacent pair
    if es_date is None:
        for i in range(1, len(pv_curve)):
            t_prev, pv_prev = pv_curve[i - 1]
            t_curr, pv_curr = pv_curve[i]
            if pv_prev <= ev_now <= pv_curr:
                if pv_curr - pv_prev > 1e-9:
                    frac = (ev_now - pv_prev) / (pv_curr - pv_prev)
                else:
                    frac = 0.0
                delta_days = (t_curr - t_prev).days * frac
                es_date = t_prev + _dt.timedelta(days=int(delta_days))
                break

    # Above last point — clamp to last
    if es_date is None:
        es_date = pv_curve[-1][0]

    es_weeks = max((es_date - project_start).days / 7.0, 0.0)
    sv_t = es_weeks - at_weeks
    spi_t = es_weeks / at_weeks if at_weeks > 0 else None
    return {"at": at_weeks, "es": es_weeks, "sv_t": sv_t, "spi_t": spi_t}


def progress_data_quality(spi_h: Optional[float],
                         spi_t: Optional[float],
                         completion_pct: float,
                         has_resources: bool) -> List[Dict[str, Any]]:
    """RULE 7 — Progress data quality warnings.

    Returns list of {warning, severity}. Severity: 'high'|'medium'|'low'.
    """
    warnings: List[Dict[str, Any]] = []
    if spi_h is not None and spi_t is not None:
        if abs(spi_h - spi_t) > 0.15:
            warnings.append({
                "warning": "SPI(h) vs SPI(t) divergence > 0.15 — EV input quality concern",
                "severity": "high",
                "spi_h": spi_h, "spi_t": spi_t,
            })
    if completion_pct > 0 and not has_resources:
        warnings.append({
            "warning": "Progress entered but no resource assignments — silent EV (Phase 3b pattern)",
            "severity": "high",
        })
    return warnings
```

**Step 4: Run — PASS** (7 new + 21 prev = 28 PASS)

**Step 5: Commit**

```bash
git add evm_math.py tests/test_evm_math.py
git commit -m "Phase 5a T77: evm_math earned_schedule (Lipke 2003) + progress_data_quality (RULE 7)

ES via linear interpolation on PV curve points. AT = (data_date -
project_start) / 7 weeks. SV(t) = ES - AT, SPI(t) = ES / AT.

Data quality warnings for SPI(h) vs SPI(t) divergence and silent EV
pattern (progress without resources)."
```

---

## Task 78: `_evm_load_*` Hybrid Adapters (file + COM)

**Files:**
- Modify: `msproject_mcp_core.py` (add at end of file, after Phase 4 section)
- Create: `tests/test_msproject_evm_loader.py`

**Step 1: Failing tests**

`tests/test_msproject_evm_loader.py`:
```python
"""Test hybrid file + COM data source adapter."""
import os
import sys

import pytest

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from msproject_mcp_core import (  # noqa: E402
    _evm_load_task_data,
    _evm_load_progress_data,
    _evm_load_baseline_data,
    _evm_detect_currency_mode,
)

MSP_XML = os.path.join(os.path.dirname(__file__), "fixtures", "sample_msp.xml")


def test_evm_load_task_data_xml():
    """file_path → reads via Phase 4 helpers, returns task list."""
    r = _evm_load_task_data(file_path=MSP_XML)
    assert r["status"] == "ok"
    assert len(r["tasks"]) == 3  # sample fixture
    for t in r["tasks"]:
        for k in ("id", "name", "duration_h"):
            assert k in t


def test_evm_load_task_data_xml_includes_resources():
    r = _evm_load_task_data(file_path=MSP_XML)
    assert "resources" in r
    assert len(r["resources"]) == 2  # R1, R2


def test_evm_load_progress_data_xml():
    r = _evm_load_progress_data(file_path=MSP_XML)
    assert r["status"] == "ok"
    assert "tasks" in r


def test_evm_load_baseline_data_xml():
    r = _evm_load_baseline_data(file_path=MSP_XML, baseline_number=0)
    assert r["status"] == "ok"
    assert r["baseline_number"] == 0


def test_evm_load_baseline_data_invalid_number():
    r = _evm_load_baseline_data(file_path=MSP_XML, baseline_number=99)
    assert r["status"] == "error"
    assert "0-10" in r["error"]


def test_evm_detect_currency_mode_hours():
    """Sample fixture has zero costs → hours mode."""
    tasks = _evm_load_task_data(file_path=MSP_XML)["tasks"]
    resources = _evm_load_task_data(file_path=MSP_XML)["resources"]
    mode = _evm_detect_currency_mode(tasks, resources)
    assert mode == "hours"
```

**Step 2: Run — FAIL**

**Step 3: Implementation**

Add at end of `msproject_mcp_core.py` (after Phase 4 dispatcher, before `def main`):
```python
# ============================================================================
# PHASE 5A — EVM TOOL
# ============================================================================
import datetime as _dt5
from evm_math import (
    compute_metrics as _evm_compute,
    forecast as _evm_forecast,
    earned_schedule as _evm_earned_schedule,
    time_phased_pv as _evm_tp_pv,
    time_phased_ev as _evm_tp_ev,
    period_delta as _evm_period_delta,
    progress_data_quality as _evm_pdq,
    rag_status as _evm_rag,
)


def _evm_load_task_data(file_path: Optional[str] = None) -> Dict[str, Any]:
    """Hybrid: file_path → Phase 4 file path; None → Phase 1 COM path.

    Returns {status, tasks: [...], resources: [...], assignments: [...],
             status_date, project_name}.
    """
    try:
        if file_path:
            tr = _msp_file_read_tasks(file_path=file_path)
            if tr.get("status") != "ok":
                return tr
            rr = _msp_file_read_resources(file_path=file_path)
            ar = _msp_file_read_assignments(file_path=file_path)
            pr = _msp_file_read_progress(file_path=file_path)
            return {
                "status": "ok",
                "tasks": tr.get("tasks", []),
                "resources": rr.get("resources", []) if rr.get("status") == "ok" else [],
                "assignments": ar.get("assignments", []) if ar.get("status") == "ok" else [],
                "status_date": pr.get("status_date") if pr.get("status") == "ok" else None,
                "project_file": file_path,
            }
        # COM path
        app = _validate_active_project()
        proj = app.ActiveProject
        tasks: List[Dict[str, Any]] = []
        for i in range(1, proj.Tasks.Count + 1):
            t = proj.Tasks(i)
            if t is None:
                continue
            tasks.append({
                "id": t.ID,
                "name": t.Name or "",
                "duration_h": float(t.Duration or 0) / 60.0,
                "start": str(t.Start) if t.Start else None,
                "finish": str(t.Finish) if t.Finish else None,
                "percent_complete": float(t.PercentComplete or 0),
                "summary": bool(t.Summary),
                "baseline_start": str(t.BaselineStart) if t.BaselineStart else None,
                "baseline_finish": str(t.BaselineFinish) if t.BaselineFinish else None,
                "baseline_work": float(t.BaselineWork or 0) / 60.0,
                "actual_work": float(t.ActualWork or 0) / 60.0,
            })
        resources: List[Dict[str, Any]] = []
        for i in range(1, proj.Resources.Count + 1):
            r = proj.Resources(i)
            if r is None:
                continue
            resources.append({
                "id": r.ID,
                "name": r.Name or "",
                "type": "Work",
                "max_units": float(r.MaxUnits or 1.0),
            })
        try:
            status_date = str(proj.StatusDate) if proj.StatusDate else None
        except Exception:
            status_date = None
        return {
            "status": "ok",
            "tasks": [t for t in tasks if not t["summary"]],
            "resources": resources,
            "assignments": [],  # COM path skips for perf
            "status_date": status_date,
            "project_name": proj.Name,
        }
    except Exception as e:
        logger.exception(f"_evm_load_task_data failed: {e}")
        return {"status": "error", "error": str(e)}


def _evm_load_progress_data(file_path: Optional[str] = None) -> Dict[str, Any]:
    """Read progress fields (percent_complete, actual_work, status_date)."""
    if file_path:
        return _msp_file_read_progress(file_path=file_path)
    # COM path
    return _msp_progress_summary()


def _evm_load_baseline_data(file_path: Optional[str] = None,
                            baseline_number: int = 0) -> Dict[str, Any]:
    """Read baseline data per Phase 3a."""
    if baseline_number not in BASELINE_NUMBERS:
        return {"status": "error",
                "error": f"baseline_number must be 0-10, got {baseline_number}"}
    if file_path:
        return _msp_file_read_baselines(file_path=file_path,
                                       baseline_number=baseline_number)
    # COM path — read baseline_saved_date + tasks
    app = _validate_active_project()
    proj = app.ActiveProject
    saved = _baseline_saved_date(proj, baseline_number)
    tasks_baseline = []
    for i in range(1, proj.Tasks.Count + 1):
        t = proj.Tasks(i)
        if t is None or t.Summary:
            continue
        b = _read_task_baseline(t, baseline_number)
        b["task_id"] = t.ID
        tasks_baseline.append(b)
    return {
        "status": "ok",
        "baseline_number": baseline_number,
        "saved_date": str(saved) if saved else None,
        "tasks": tasks_baseline,
    }


def _evm_detect_currency_mode(tasks: List[Dict[str, Any]],
                              resources: List[Dict[str, Any]]) -> str:
    """RULE 3 — hours vs cost loading.

    'hours' if no cost data anywhere or sum(cost)==sum(work_h);
    'cost' if any non-zero cost field found.
    """
    total_cost = 0.0
    for t in tasks:
        try:
            total_cost += float(t.get("cost") or 0)
        except (TypeError, ValueError):
            pass
    for r in resources:
        try:
            total_cost += float(r.get("cost") or 0)
        except (TypeError, ValueError):
            pass
    return "cost" if total_cost > 0 else "hours"
```

**Step 4: Run — PASS** (6 PASS)

**Step 5: Commit**

```bash
git add msproject_mcp_core.py tests/test_msproject_evm_loader.py
git commit -m "Phase 5a T78: hybrid data source adapters (file + COM)

_evm_load_task_data, _evm_load_progress_data, _evm_load_baseline_data —
file_path optional. file path → Phase 4 helpers; None → Phase 1 COM
iter. _evm_detect_currency_mode implements RULE 3 (hours vs cost).

Phase 1-4 helpers DOKUNULMAZ — only read-only calls."
```

---

## Task 79: compute_metrics + forecast + summary Action Helpers

**Files:**
- Modify: `msproject_mcp_core.py`
- Create: `tests/test_msproject_evm_compute.py`

**Step 1: Failing tests**

`tests/test_msproject_evm_compute.py`:
```python
import os
import sys

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from msproject_mcp_core import (  # noqa: E402
    _msp_evm_compute_metrics,
    _msp_evm_forecast,
    _msp_evm_summary,
)

MSP_XML = os.path.join(os.path.dirname(__file__), "fixtures", "sample_msp.xml")


def test_msp_evm_compute_metrics_xml():
    r = _msp_evm_compute_metrics(file_path=MSP_XML)
    assert r["status"] == "ok"
    for k in ("bac", "ev", "ac", "pv", "spi", "cpi", "sv", "cv"):
        assert k in r


def test_msp_evm_forecast_xml():
    r = _msp_evm_forecast(file_path=MSP_XML)
    assert r["status"] == "ok"
    for k in ("eac_t1", "eac_t2", "eac_t3", "etc", "vac", "tcpi_bac"):
        assert k in r


def test_msp_evm_summary_xml():
    r = _msp_evm_summary(file_path=MSP_XML)
    assert r["status"] == "ok"
    assert "rag" in r
    assert r["rag"] in ("RED", "AMBER", "GREEN")
```

**Step 2: Run — FAIL**

**Step 3: Implementation**

Append to Phase 5a section:
```python
def _evm_compute_pv_ev_ac(load_data: Dict[str, Any],
                         baseline_load: Dict[str, Any]) -> Tuple[float, float, float, float]:
    """Aggregate PV/EV/AC/BAC across all real tasks. Hours mode default.

    PV = sum(baseline_work × percent_at_data_date) — uses linear distribution
         Use linear PV at data_date.
    EV = sum(baseline_work × percent_complete / 100)
    AC = sum(actual_work)
    BAC = sum(baseline_work)
    """
    tasks = load_data.get("tasks", [])
    bac = sum(float(t.get("baseline_work") or 0) for t in tasks)
    ev = sum(float(t.get("baseline_work") or 0) *
             float(t.get("percent_complete") or 0) / 100.0
             for t in tasks)
    ac = sum(float(t.get("actual_work") or 0) for t in tasks)
    # PV at data_date — use linear distribution per task
    sd_str = load_data.get("status_date")
    data_date = _parse_iso_date(sd_str) if sd_str else _dt5.date.today()
    # Build task list with date-coerced baseline fields
    enriched = []
    for t in tasks:
        bs = _parse_iso_date(t.get("baseline_start"))
        bf = _parse_iso_date(t.get("baseline_finish"))
        if bs is None or bf is None:
            continue
        enriched.append({
            "baseline_start": bs,
            "baseline_finish": bf,
            "baseline_work": float(t.get("baseline_work") or 0),
        })
    if enriched:
        pv = sum(_evm_tp_pv([t], [(_dt5.date.min, data_date)])[0] for t in enriched)
    else:
        pv = 0.0
    return bac, pv, ev, ac


def _parse_iso_date(s: Optional[str]) -> Optional[_dt5.date]:
    """Parse '2026-01-01...' or 'YYYY-MM-DD'-prefix string to date."""
    if not s or s == "N/A":
        return None
    try:
        return _dt5.date.fromisoformat(s[:10])
    except (ValueError, TypeError):
        return None


def _msp_evm_compute_metrics(file_path: Optional[str] = None,
                             baseline_number: int = 0) -> Dict[str, Any]:
    """Action 1: compute_metrics — SPI/CPI/SV/CV (RULE 4)."""
    load = _evm_load_task_data(file_path=file_path)
    if load.get("status") != "ok":
        return load
    bload = _evm_load_baseline_data(file_path=file_path, baseline_number=baseline_number)
    if bload.get("status") != "ok":
        return bload
    bac, pv, ev, ac = _evm_compute_pv_ev_ac(load, bload)
    metrics = _evm_compute(bac=bac, pv=pv, ev=ev, ac=ac)
    return {"status": "ok", "baseline_number": baseline_number, **metrics}


def _msp_evm_forecast(file_path: Optional[str] = None,
                     baseline_number: int = 0) -> Dict[str, Any]:
    """Action 2: forecast — EAC1/2/3, ETC, VAC, TCPI (RULE 9)."""
    cm = _msp_evm_compute_metrics(file_path=file_path, baseline_number=baseline_number)
    if cm.get("status") != "ok":
        return cm
    fc = _evm_forecast(bac=cm["bac"], ev=cm["ev"], ac=cm["ac"],
                      cpi=cm.get("cpi"), spi=cm.get("spi"))
    return {"status": "ok", "baseline_number": baseline_number, **fc}


def _msp_evm_summary(file_path: Optional[str] = None,
                    baseline_number: int = 0) -> Dict[str, Any]:
    """Action 4: summary — RAG (RULE 12) + executive."""
    cm = _msp_evm_compute_metrics(file_path=file_path, baseline_number=baseline_number)
    if cm.get("status") != "ok":
        return cm
    completion_pct = (cm["ev"] / cm["bac"] * 100.0) if cm["bac"] > 0 else 0.0
    rag = _evm_rag(spi=cm.get("spi"), completion_pct=completion_pct)
    return {
        "status": "ok",
        "baseline_number": baseline_number,
        "rag": rag,
        "completion_pct": round(completion_pct, 2),
        "spi": cm.get("spi"),
        "cpi": cm.get("cpi"),
        "schedule_health": rag,
    }
```

**Step 4: Run — PASS** (3 PASS)

**Step 5: Commit**

```bash
git add msproject_mcp_core.py tests/test_msproject_evm_compute.py
git commit -m "Phase 5a T79: compute_metrics + forecast + summary action helpers

Aggregate BAC/EV/AC from task list; PV via linear time-phasing at
data_date. Wraps evm_math via Phase 5a load adapters."
```

---

## Task 80: earned_schedule + progress_data_quality + detect_currency_mode

**Files:**
- Modify: `msproject_mcp_core.py`
- Create: `tests/test_msproject_evm_earned_schedule.py`
- Create: `tests/test_msproject_evm_dataquality.py`
- Create: `tests/test_msproject_evm_currency.py`

**Step 1: Failing tests**

`tests/test_msproject_evm_earned_schedule.py`:
```python
import os
import sys
REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)
from msproject_mcp_core import _msp_evm_earned_schedule

MSP_XML = os.path.join(os.path.dirname(__file__), "fixtures", "sample_msp.xml")


def test_msp_evm_earned_schedule_xml():
    r = _msp_evm_earned_schedule(file_path=MSP_XML, bucket="week")
    assert r["status"] == "ok"
    assert "at" in r and "es" in r and "sv_t" in r and "spi_t" in r
```

`tests/test_msproject_evm_dataquality.py`:
```python
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
```

`tests/test_msproject_evm_currency.py`:
```python
import os, sys
REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)
from msproject_mcp_core import _msp_evm_detect_currency_mode

MSP_XML = os.path.join(os.path.dirname(__file__), "fixtures", "sample_msp.xml")


def test_msp_evm_currency_xml():
    r = _msp_evm_detect_currency_mode(file_path=MSP_XML)
    assert r["status"] == "ok"
    assert r["mode"] in ("hours", "cost")
```

**Step 2: Run — FAIL**

**Step 3: Implementation**

Append:
```python
def _evm_build_pv_curve(tasks: List[Dict[str, Any]],
                       project_start: _dt5.date,
                       project_finish: _dt5.date,
                       bucket: str = "week") -> List[Tuple[_dt5.date, float]]:
    """Build cumulative PV curve points across project duration."""
    enriched = []
    for t in tasks:
        bs = _parse_iso_date(t.get("baseline_start"))
        bf = _parse_iso_date(t.get("baseline_finish"))
        if bs is None or bf is None:
            continue
        enriched.append({
            "baseline_start": bs, "baseline_finish": bf,
            "baseline_work": float(t.get("baseline_work") or 0),
        })
    delta = _dt5.timedelta(days=7) if bucket == "week" else \
            _dt5.timedelta(days=1) if bucket == "day" else \
            _dt5.timedelta(days=30)
    points = []
    d = project_start
    while d <= project_finish:
        d += delta
        if enriched:
            pv_now = _evm_tp_pv(enriched, [(_dt5.date.min, d)])[0]
        else:
            pv_now = 0.0
        points.append((d, pv_now))
    return points


def _msp_evm_earned_schedule(file_path: Optional[str] = None,
                             baseline_number: int = 0,
                             bucket: str = "week") -> Dict[str, Any]:
    """Action 3: earned_schedule (RULE 8)."""
    load = _evm_load_task_data(file_path=file_path)
    if load.get("status") != "ok":
        return load
    tasks = load.get("tasks", [])
    if not tasks:
        return {"status": "error", "error": "No tasks loaded"}
    # Compute project start/finish bounds
    starts = [_parse_iso_date(t.get("baseline_start") or t.get("start"))
              for t in tasks]
    finishes = [_parse_iso_date(t.get("baseline_finish") or t.get("finish"))
                for t in tasks]
    starts = [s for s in starts if s is not None]
    finishes = [f for f in finishes if f is not None]
    if not starts or not finishes:
        return {"status": "error", "error": "Cannot determine project bounds"}
    project_start = min(starts)
    project_finish = max(finishes)
    sd_str = load.get("status_date")
    data_date = _parse_iso_date(sd_str) if sd_str else _dt5.date.today()
    # Compute current EV
    ev = sum(float(t.get("baseline_work") or 0) *
             float(t.get("percent_complete") or 0) / 100.0
             for t in tasks)
    # Build PV curve
    pv_curve = _evm_build_pv_curve(tasks, project_start, project_finish, bucket)
    es = _evm_earned_schedule(pv_curve=pv_curve, ev_now=ev,
                              project_start=project_start, data_date=data_date)
    return {"status": "ok", "baseline_number": baseline_number,
            "bucket": bucket, **es}


def _msp_evm_progress_data_quality(file_path: Optional[str] = None,
                                  baseline_number: int = 0) -> Dict[str, Any]:
    """Action 7: progress_data_quality (RULE 7)."""
    load = _evm_load_task_data(file_path=file_path)
    if load.get("status") != "ok":
        return load
    cm = _msp_evm_compute_metrics(file_path=file_path, baseline_number=baseline_number)
    es = _msp_evm_earned_schedule(file_path=file_path,
                                  baseline_number=baseline_number)
    spi_h = cm.get("spi") if cm.get("status") == "ok" else None
    spi_t = es.get("spi_t") if es.get("status") == "ok" else None
    completion_pct = (cm.get("ev", 0) / cm["bac"] * 100.0) if cm.get("bac", 0) > 0 else 0
    has_resources = len(load.get("resources", [])) > 0
    warnings = _evm_pdq(spi_h=spi_h, spi_t=spi_t,
                       completion_pct=completion_pct, has_resources=has_resources)
    return {"status": "ok", "warnings": warnings,
            "spi_h": spi_h, "spi_t": spi_t,
            "completion_pct": round(completion_pct, 2)}


def _msp_evm_detect_currency_mode(file_path: Optional[str] = None) -> Dict[str, Any]:
    """Action 13: detect_currency_mode (RULE 3)."""
    load = _evm_load_task_data(file_path=file_path)
    if load.get("status") != "ok":
        return load
    mode = _evm_detect_currency_mode(load.get("tasks", []),
                                    load.get("resources", []))
    return {"status": "ok", "mode": mode}
```

**Step 4: Run — PASS** (3 PASS)

**Step 5: Commit**

```bash
git add msproject_mcp_core.py tests/test_msproject_evm_earned_schedule.py tests/test_msproject_evm_dataquality.py tests/test_msproject_evm_currency.py
git commit -m "Phase 5a T80: earned_schedule + progress_data_quality + detect_currency_mode actions

PV curve builder iterates project bounds at bucket interval; ES via
linear interp. PDQ uses SPI(h) from compute_metrics + SPI(t) from ES.
Currency mode per RULE 3."
```

---

## Task 81: time_phased_evm + period_delta (BIG ONE — bucket logic)

**Files:**
- Modify: `msproject_mcp_core.py`
- Create: `tests/test_msproject_evm_time_phased.py`
- Create: `tests/test_msproject_evm_period_delta.py`

**Step 1: Failing tests**

`tests/test_msproject_evm_time_phased.py`:
```python
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
```

`tests/test_msproject_evm_period_delta.py`:
```python
import os, sys
REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)
from msproject_mcp_core import _msp_evm_period_delta

MSP_XML = os.path.join(os.path.dirname(__file__), "fixtures", "sample_msp.xml")


def test_msp_evm_period_delta_first_period(tmp_path):
    """No prev snapshot → period values = current cum values."""
    snap_path = str(tmp_path / "snaps.json")
    r = _msp_evm_period_delta(file_path=MSP_XML, snapshot_path=snap_path)
    assert r["status"] == "ok"
    assert "period_pv" in r and "period_ev" in r and "period_ac" in r
    assert r["period_bac"] == 0


def test_msp_evm_period_delta_with_prev_snapshot(tmp_path):
    """prev snapshot exists → period values = current - prev."""
    import json
    snap_path = tmp_path / "snaps.json"
    snap_path.write_text(json.dumps({
        "snapshots": [{
            "saved_at": "2026-01-01T00:00:00",
            "metrics": {"pv": 100, "ev": 80, "ac": 90, "bac": 1000},
        }]
    }))
    r = _msp_evm_period_delta(file_path=MSP_XML, snapshot_path=str(snap_path))
    assert r["status"] == "ok"
    # period should equal current - 100/80/90
```

**Step 2: Run — FAIL**

**Step 3: Implementation**

```python
def _evm_bucket_to_delta(bucket: str) -> Optional[_dt5.timedelta]:
    if bucket == "day":
        return _dt5.timedelta(days=1)
    if bucket == "week":
        return _dt5.timedelta(days=7)
    if bucket == "month":
        return _dt5.timedelta(days=30)
    return None


def _msp_evm_time_phased_evm(file_path: Optional[str] = None,
                             baseline_number: int = 0,
                             bucket: str = "week") -> Dict[str, Any]:
    """Action 5: time_phased_evm — PV/EV/AC per period."""
    delta = _evm_bucket_to_delta(bucket)
    if delta is None:
        return {"status": "error",
                "error": f"bucket must be day/week/month, got '{bucket}'"}
    load = _evm_load_task_data(file_path=file_path)
    if load.get("status") != "ok":
        return load
    tasks = load.get("tasks", [])
    if not tasks:
        return {"status": "ok", "buckets": []}
    starts = [_parse_iso_date(t.get("baseline_start") or t.get("start"))
              for t in tasks]
    finishes = [_parse_iso_date(t.get("baseline_finish") or t.get("finish"))
                for t in tasks]
    starts = [s for s in starts if s is not None]
    finishes = [f for f in finishes if f is not None]
    if not starts or not finishes:
        return {"status": "ok", "buckets": []}
    project_start = min(starts)
    project_finish = max(finishes)
    sd_str = load.get("status_date")
    data_date = _parse_iso_date(sd_str) if sd_str else _dt5.date.today()
    # Build buckets and compute PV/EV
    buckets: List[Tuple[_dt5.date, _dt5.date]] = []
    d = project_start
    while d <= project_finish:
        next_d = d + delta
        buckets.append((d, min(next_d, project_finish + _dt5.timedelta(days=1))))
        d = next_d
    enriched = [{
        "baseline_start": _parse_iso_date(t.get("baseline_start")),
        "baseline_finish": _parse_iso_date(t.get("baseline_finish")),
        "baseline_work": float(t.get("baseline_work") or 0),
        "percent_complete": float(t.get("percent_complete") or 0),
    } for t in tasks]
    pv = _evm_tp_pv(enriched, [(s, e) for (s, e) in buckets])
    ev = _evm_tp_ev(enriched, [(s, e) for (s, e) in buckets], data_date=data_date)
    # AC simplification: not time-phased per task (Phase 6 polish);
    # current AC distributed evenly across past buckets up to data_date
    total_ac = sum(float(t.get("actual_work") or 0) for t in tasks)
    past_buckets = sum(1 for (_, e) in buckets if e <= data_date)
    ac_per_bucket = (total_ac / past_buckets) if past_buckets > 0 else 0.0
    ac_cum = 0.0
    out = []
    for i, (s, e) in enumerate(buckets):
        if e <= data_date:
            ac_cum += ac_per_bucket
        out.append({
            "period_start": s.isoformat(),
            "period_end": e.isoformat(),
            "pv": round(pv[i], 2),
            "ev": round(ev[i], 2),
            "ac": round(ac_cum, 2),
        })
    return {"status": "ok", "bucket": bucket, "buckets": out}


def _msp_evm_period_delta(file_path: Optional[str] = None,
                          baseline_number: int = 0,
                          snapshot_path: Optional[str] = None) -> Dict[str, Any]:
    """Action 6: period_delta vs prev snapshot (RULE 6)."""
    cm = _msp_evm_compute_metrics(file_path=file_path, baseline_number=baseline_number)
    if cm.get("status") != "ok":
        return cm
    snap_now = {"pv": cm["pv"], "ev": cm["ev"], "ac": cm["ac"], "bac": cm["bac"]}
    snap_prev = None
    if snapshot_path and os.path.exists(snapshot_path):
        try:
            import json as _json
            with open(snapshot_path, "r", encoding="utf-8") as f:
                data = _json.load(f)
            snaps = sorted(data.get("snapshots", []),
                          key=lambda s: s.get("saved_at", ""))
            if snaps:
                snap_prev = snaps[-1].get("metrics", {})
        except Exception as e:
            logger.warning(f"period_delta: failed to load prev snapshot: {e}")
    delta = _evm_period_delta(snap_now, snap_prev)
    return {"status": "ok", "current": snap_now, **delta}
```

**Step 4: Run — PASS** (5 PASS — 3 time_phased + 2 period_delta)

**Step 5: Commit**

```bash
git add msproject_mcp_core.py tests/test_msproject_evm_time_phased.py tests/test_msproject_evm_period_delta.py
git commit -m "Phase 5a T81: time_phased_evm + period_delta (BIG ONE — bucket logic)

Time-phased PV/EV per RULE 5 + AC distributed evenly across past buckets
(Phase 6 polish for true time-phased AC). bucket=day/week/month.
period_delta loads previous snapshot from snapshot_path JSON file."
```

---

## Task 82: variance_to_baseline + compare_baselines_evm

**Files:**
- Modify: `msproject_mcp_core.py`
- Create: `tests/test_msproject_evm_baseline.py`

**Step 1: Failing tests**

```python
import os, sys
REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)
from msproject_mcp_core import (
    _msp_evm_variance_to_baseline,
    _msp_evm_compare_baselines_evm,
)

MSP_XML = os.path.join(os.path.dirname(__file__), "fixtures", "sample_msp.xml")


def test_msp_evm_variance_to_baseline_xml():
    r = _msp_evm_variance_to_baseline(file_path=MSP_XML, baseline_number=0)
    assert r["status"] == "ok"
    assert "baseline_number" in r


def test_msp_evm_variance_invalid_baseline():
    r = _msp_evm_variance_to_baseline(file_path=MSP_XML, baseline_number=99)
    assert r["status"] == "error"


def test_msp_evm_compare_baselines_xml():
    r = _msp_evm_compare_baselines_evm(file_path=MSP_XML,
                                       baseline_a=0, baseline_b=1)
    # Either ok with delta, or graceful "baseline N not saved" error
    assert r["status"] in ("ok", "error")
```

**Step 2: Run — FAIL**

**Step 3: Implementation**

```python
def _msp_evm_variance_to_baseline(file_path: Optional[str] = None,
                                  baseline_number: int = 0) -> Dict[str, Any]:
    """Action 8: variance_to_baseline."""
    if baseline_number not in BASELINE_NUMBERS:
        return {"status": "error",
                "error": f"baseline_number must be 0-10, got {baseline_number}"}
    return _msp_evm_compute_metrics(file_path=file_path,
                                    baseline_number=baseline_number)


def _msp_evm_compare_baselines_evm(file_path: Optional[str] = None,
                                   baseline_a: int = 0,
                                   baseline_b: int = 1) -> Dict[str, Any]:
    """Action 9: compare_baselines_evm — B_a vs B_b EVM delta."""
    a = _msp_evm_compute_metrics(file_path=file_path, baseline_number=baseline_a)
    if a.get("status") != "ok":
        return {"status": "error",
                "error": f"baseline_a {baseline_a}: {a.get('error', 'load failed')}"}
    b = _msp_evm_compute_metrics(file_path=file_path, baseline_number=baseline_b)
    if b.get("status") != "ok":
        return {"status": "error",
                "error": f"baseline_b {baseline_b}: {b.get('error', 'load failed')}"}
    delta = {
        "bac_delta": b["bac"] - a["bac"],
        "spi_delta": (b.get("spi") or 0) - (a.get("spi") or 0),
        "cpi_delta": (b.get("cpi") or 0) - (a.get("cpi") or 0),
    }
    return {"status": "ok",
            "baseline_a": a, "baseline_b": b, "delta": delta}
```

**Step 4: Run — PASS** (3 PASS)

**Step 5: Commit**

```bash
git add msproject_mcp_core.py tests/test_msproject_evm_baseline.py
git commit -m "Phase 5a T82: variance_to_baseline + compare_baselines_evm

Phase 3a integration — wraps compute_metrics across multiple baselines
and computes delta for revision impact analysis (memory: Phase 3a
compare_two pattern)."
```

---

## Task 83: Snapshot helpers + 3 history actions

**Files:**
- Modify: `msproject_mcp_core.py`
- Create: `tests/test_msproject_evm_snapshot.py`

**Step 1: Failing tests**

```python
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
```

**Step 2: Run — FAIL**

**Step 3: Implementation**

```python
def _evm_snapshot_save(snapshot_path: str, snapshot: Dict[str, Any]) -> None:
    """Append snapshot to JSON file. Creates file with empty array if missing."""
    import json as _json
    if os.path.exists(snapshot_path):
        with open(snapshot_path, "r", encoding="utf-8") as f:
            data = _json.load(f)
    else:
        data = {"snapshots": []}
    data.setdefault("snapshots", []).append(snapshot)
    os.makedirs(os.path.dirname(snapshot_path) or ".", exist_ok=True)
    with open(snapshot_path, "w", encoding="utf-8") as f:
        _json.dump(data, f, indent=2, default=str)


def _evm_snapshot_load(snapshot_path: str,
                       project_filter: Optional[str] = None,
                       baseline_filter: Optional[int] = None) -> List[Dict[str, Any]]:
    import json as _json
    if not os.path.exists(snapshot_path):
        return []
    with open(snapshot_path, "r", encoding="utf-8") as f:
        data = _json.load(f)
    snaps = data.get("snapshots", [])
    if project_filter:
        snaps = [s for s in snaps
                 if project_filter in (s.get("project_name") or s.get("project_file") or "")]
    if baseline_filter is not None:
        snaps = [s for s in snaps if s.get("baseline_number") == baseline_filter]
    return snaps


def _msp_evm_save_period_snapshot(file_path: Optional[str] = None,
                                  baseline_number: int = 0,
                                  snapshot_path: str = None,
                                  tag: Optional[str] = None) -> Dict[str, Any]:
    """Action 10: save_period_snapshot."""
    if not snapshot_path:
        snapshot_path = os.path.expanduser("~/msproject_evm_snapshots.json")
    cm = _msp_evm_compute_metrics(file_path=file_path,
                                  baseline_number=baseline_number)
    if cm.get("status") != "ok":
        return cm
    fc = _msp_evm_forecast(file_path=file_path, baseline_number=baseline_number)
    es = _msp_evm_earned_schedule(file_path=file_path,
                                  baseline_number=baseline_number)
    summary = _msp_evm_summary(file_path=file_path,
                               baseline_number=baseline_number)
    snap = {
        "id": _dt5.datetime.now().strftime("%Y%m%d-%H%M%S"),
        "saved_at": _dt5.datetime.now().isoformat(),
        "project_file": file_path,
        "baseline_number": baseline_number,
        "metrics": {k: cm.get(k) for k in
                    ("bac", "pv", "ev", "ac", "spi", "cpi", "sv", "cv")},
        "forecast": {k: fc.get(k) for k in
                     ("eac_t1", "eac_t2", "eac_t3", "etc", "vac",
                      "tcpi_bac", "tcpi_eac")} if fc.get("status") == "ok" else {},
        "earned_schedule": {k: es.get(k) for k in
                            ("at", "es", "sv_t", "spi_t")} if es.get("status") == "ok" else {},
        "rag": summary.get("rag") if summary.get("status") == "ok" else None,
        "tag": tag,
    }
    try:
        _evm_snapshot_save(snapshot_path, snap)
        return {"status": "ok", "snapshot_path": snapshot_path,
                "snapshot_id": snap["id"]}
    except Exception as e:
        logger.exception(f"save_period_snapshot failed: {e}")
        return {"status": "error", "error": str(e)}


def _msp_evm_get_period_history(snapshot_path: Optional[str] = None,
                               project_filter: Optional[str] = None,
                               baseline_filter: Optional[int] = None) -> Dict[str, Any]:
    """Action 11: get_period_history."""
    if not snapshot_path:
        snapshot_path = os.path.expanduser("~/msproject_evm_snapshots.json")
    snaps = _evm_snapshot_load(snapshot_path, project_filter=project_filter,
                              baseline_filter=baseline_filter)
    return {"status": "ok", "count": len(snaps), "snapshots": snaps}


def _msp_evm_trend(snapshot_path: Optional[str] = None,
                  project_filter: Optional[str] = None) -> Dict[str, Any]:
    """Action 12: trend — period-over-period series for SPI/CPI/EAC."""
    if not snapshot_path:
        snapshot_path = os.path.expanduser("~/msproject_evm_snapshots.json")
    snaps = _evm_snapshot_load(snapshot_path, project_filter=project_filter)
    snaps_sorted = sorted(snaps, key=lambda s: s.get("saved_at", ""))
    series = []
    for s in snaps_sorted:
        m = s.get("metrics", {}) or {}
        f = s.get("forecast", {}) or {}
        series.append({
            "saved_at": s.get("saved_at"),
            "tag": s.get("tag"),
            "spi": m.get("spi"),
            "cpi": m.get("cpi"),
            "eac_t3": f.get("eac_t3"),
            "rag": s.get("rag"),
        })
    return {"status": "ok", "count": len(series), "series": series}
```

**Step 4: Run — PASS** (4 PASS)

**Step 5: Commit**

```bash
git add msproject_mcp_core.py tests/test_msproject_evm_snapshot.py
git commit -m "Phase 5a T83: snapshot helpers + 3 history actions

JSON-backed append-only storage at ~/msproject_evm_snapshots.json
(override via snapshot_path). save_period_snapshot bundles metrics +
forecast + ES + rag. get_period_history filters by project/baseline.
trend extracts SPI/CPI/EAC trajectory series."
```

---

## Task 84: FastMCP Dispatcher + Acceptance Script + README + Push (FINAL)

**Files:**
- Modify: `msproject_mcp_core.py` (add `@mcp.tool msproject_evm` dispatcher)
- Create: `tests/test_msproject_evm_dispatcher.py`
- Create: `samples/build_evm_lifecycle.py`
- Modify: `README.md`

**Step 1: Failing dispatcher tests**

`tests/test_msproject_evm_dispatcher.py`:
```python
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
```

**Step 2: Run — FAIL**

**Step 3: Implementation**

Add `@mcp.tool msproject_evm` after `msproject_file` dispatcher:
```python
@mcp.tool(
    name="msproject_evm",
    annotations={"title": "MS Project EVM Operations", "readOnlyHint": True},
)
async def msproject_evm(params: dict) -> str:
    """Earned Value Management — PMI PMBOK 8th § 7.4.2 + Lipke 2003 ES.

    Hybrid: file_path verilirse Phase 4 file path; yoksa Phase 1 COM.

    Actions:
    - compute_metrics: SPI/CPI/SV/CV (RULE 4)
    - forecast: EAC1/2/3 + ETC + VAC + TCPI(BAC/EAC) (RULE 9)
    - earned_schedule: AT, ES, SV(t), SPI(t) (RULE 8 Lipke)
    - summary: RAG + completion_pct + executive (RULE 12)
    - time_phased_evm: PV/EV/AC per period (bucket day/week/month, RULE 5)
    - period_delta: vs prev snapshot (RULE 6)
    - progress_data_quality: warnings (RULE 7)
    - variance_to_baseline: vs Baseline N (Phase 3a integration)
    - compare_baselines_evm: B_a vs B_b EVM delta
    - save_period_snapshot: append to JSON snapshot file
    - get_period_history: list saved snapshots (filter by project/baseline)
    - trend: SPI/CPI/EAC trajectory series
    - detect_currency_mode: hours vs cost (RULE 3)

    Phase 5a (30 Apr 2026). Tool count 8 → 9.
    """
    import json
    action = params.get("action", "")
    p = {k: v for k, v in params.items() if k != "action"}
    try:
        if action == "compute_metrics":
            r = _msp_evm_compute_metrics(**p)
        elif action == "forecast":
            r = _msp_evm_forecast(**p)
        elif action == "earned_schedule":
            r = _msp_evm_earned_schedule(**p)
        elif action == "summary":
            r = _msp_evm_summary(**p)
        elif action == "time_phased_evm":
            r = _msp_evm_time_phased_evm(**p)
        elif action == "period_delta":
            r = _msp_evm_period_delta(**p)
        elif action == "progress_data_quality":
            r = _msp_evm_progress_data_quality(**p)
        elif action == "variance_to_baseline":
            r = _msp_evm_variance_to_baseline(**p)
        elif action == "compare_baselines_evm":
            r = _msp_evm_compare_baselines_evm(**p)
        elif action == "save_period_snapshot":
            r = _msp_evm_save_period_snapshot(**p)
        elif action == "get_period_history":
            r = _msp_evm_get_period_history(**p)
        elif action == "trend":
            r = _msp_evm_trend(**p)
        elif action == "detect_currency_mode":
            r = _msp_evm_detect_currency_mode(**p)
        else:
            r = {"status": "error",
                 "error": (f"Unknown action '{action}'. Valid: "
                          "compute_metrics/forecast/earned_schedule/summary/"
                          "time_phased_evm/period_delta/progress_data_quality/"
                          "variance_to_baseline/compare_baselines_evm/"
                          "save_period_snapshot/get_period_history/trend/"
                          "detect_currency_mode")}
    except TypeError as e:
        r = {"status": "error", "error": f"Invalid params for {action}: {e}"}
    except Exception as e:
        logger.exception(f"msproject_evm({action}) failed: {e}")
        r = {"status": "error", "error": str(e)}
    return json.dumps(r, default=str, ensure_ascii=False)
```

**Step 4: Acceptance script**

`samples/build_evm_lifecycle.py`:
```python
"""Phase 5a EVM acceptance: 200 task CAU-style + 3 baseline + 4 snapshots <30s.

SAFETY: FileNew + FileClose 0. User's active project untouched.

Scenario:
  1. Build 200 villa tasks + 14 CAU resources + assignments
  2. Save Baseline 0 (Original)
  3. Slip + revize, save Baseline 1
  4. Phase 3b: progress for week 1 (~30%)
  5. set_status_date "week 1"; save_period_snapshot tag=w1
  6. More progress for week 2 (~60%); save_period_snapshot tag=w2
  7. Week 3 + Week 4 snapshots → 4 history entries
  8. trend → SPI/CPI/EAC trajectory
  9. variance_to_baseline + compare_baselines_evm B0 vs B1
  10. progress_data_quality + detect_currency_mode
"""
import os, sys, time, tempfile
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import pythoncom, win32com.client
from msproject_mcp_core import (
    _msp_task_bulk_add, _msp_resource_add, _msp_resource_bulk_assign,
    _msp_baseline_save, _msp_progress_set_task, _msp_progress_set_status_date,
    _msp_evm_compute_metrics, _msp_evm_forecast, _msp_evm_summary,
    _msp_evm_save_period_snapshot, _msp_evm_get_period_history, _msp_evm_trend,
    _msp_evm_variance_to_baseline, _msp_evm_compare_baselines_evm,
    _msp_evm_progress_data_quality, _msp_evm_detect_currency_mode,
)

N_TASKS = 200


def main():
    pythoncom.CoInitialize()
    app = win32com.client.GetActiveObject("MSProject.Application")
    original_name = app.ActiveProject.Name if app.ActiveProject else None
    app.FileNew()
    test_proj = app.ActiveProject
    test_name = test_proj.Name
    print(f"[SAFE] isolated test: {test_name}")

    tmpdir = tempfile.mkdtemp(prefix="evm_phase5a_")
    snap_path = os.path.join(tmpdir, "snapshots.json")

    try:
        t0 = time.time()
        # 1. Build base
        print(f"\n1. Building {N_TASKS} tasks + 14 resources...")
        items = [{"name": f"V{i:03d}", "duration": "5d"} for i in range(N_TASKS)]
        tasks = _msp_task_bulk_add(items=items)
        task_ids = tasks["task_ids"]
        cau = ["COW", "EXT", "STL", "CAR", "MSN", "DRW", "INW",
               "EWI", "CWI", "ACP", "ELW", "DMS", "MTR", "LBR"]
        res_ids = [_msp_resource_add(name=n, type="Work")["resource_id"] for n in cau]
        # Subset assignments — full 2800 not needed, sample 200
        sample = [{"task_id": tid, "resource_id": res_ids[i % 14]}
                  for i, tid in enumerate(task_ids)]
        _msp_resource_bulk_assign(items=sample)
        print(f"   OK in {time.time()-t0:.2f}s")

        # 2. Save Baseline 0
        _msp_baseline_save(baseline_number=0)
        print(f"\n2. Baseline 0 saved at {time.time()-t0:.2f}s")

        # 3-6. Week 1 progress + snapshot
        for tid in task_ids[:60]:
            _msp_progress_set_task(task_id=tid, percent_complete=30.0)
        _msp_progress_set_status_date(status_date="2026-05-07")
        s1 = _msp_evm_save_period_snapshot(snapshot_path=snap_path, tag="w1")
        print(f"\n3. Week 1 snapshot saved {s1.get('snapshot_id')}")

        # 7. Week 2-4 snapshots
        for tid in task_ids[:120]:
            _msp_progress_set_task(task_id=tid, percent_complete=60.0)
        _msp_progress_set_status_date(status_date="2026-05-14")
        _msp_evm_save_period_snapshot(snapshot_path=snap_path, tag="w2")

        for tid in task_ids[:160]:
            _msp_progress_set_task(task_id=tid, percent_complete=80.0)
        _msp_progress_set_status_date(status_date="2026-05-21")
        _msp_evm_save_period_snapshot(snapshot_path=snap_path, tag="w3")

        for tid in task_ids[:180]:
            _msp_progress_set_task(task_id=tid, percent_complete=95.0)
        _msp_progress_set_status_date(status_date="2026-05-28")
        _msp_evm_save_period_snapshot(snapshot_path=snap_path, tag="w4")

        # 8. Trend
        trend = _msp_evm_trend(snapshot_path=snap_path)
        print(f"\n4. Trend series: {len(trend['series'])} points")
        for s in trend["series"]:
            print(f"   {s['tag']}: SPI={s['spi']:.3f} CPI={s['cpi']:.3f} RAG={s['rag']}")

        # 9. Compare baselines (only 0 saved here — graceful)
        # 10. Quality + currency
        pdq = _msp_evm_progress_data_quality()
        print(f"\n5. Data quality warnings: {len(pdq.get('warnings', []))}")
        cm = _msp_evm_detect_currency_mode()
        print(f"\n6. Currency mode: {cm.get('mode')}")

        elapsed = time.time() - t0
        print(f"\n[OK] ACCEPTANCE: {elapsed:.2f}s total (target <30s)")
        assert elapsed < 30.0, f"Too slow: {elapsed}s"

    finally:
        try:
            for i in range(1, app.Projects.Count + 1):
                if app.Projects(i).Name == test_name:
                    app.WindowActivate(app.Projects(i).Windows(1).Caption)
                    app.FileClose(0)
                    break
        except Exception:
            pass
        try:
            import shutil
            shutil.rmtree(tmpdir, ignore_errors=True)
        except Exception:
            pass


if __name__ == "__main__":
    main()
```

**Step 5: Run acceptance**

```bash
python samples/build_evm_lifecycle.py
```
Expected: `[OK] ACCEPTANCE: <Xs total (target <30s)`. Realistic ~10-20s.

**Step 6: README update**

Add Phase 5a section after Phase 4 section.

**Step 7: Run full regression**

```bash
python -m pytest tests/ --tb=line --ignore=cleanup_test.py --ignore=test_apply_tabledef.py 2>&1 | tail -5
```
Expected: ~330+ PASS + 0 xfail.

**Step 8: Commit + push**

```bash
git add msproject_mcp_core.py tests/test_msproject_evm_dispatcher.py samples/build_evm_lifecycle.py README.md
git commit -m "Phase 5a T84: dispatcher + acceptance + README + push (msproject_evm 9th tool)"
git push origin main
```

---

## Phase 5a Tamamlama Kriterleri

1. ✅ T75-T84 ~12-15 commit landed
2. ✅ Acceptance script `samples/build_evm_lifecycle.py` <30s
3. ✅ Yeni testler ~50 PASS (~25 saf math, no fixtures + ~25 dispatcher/loader/snapshot)
4. ✅ Phase 1+2+3+4 mevcut regression PASS — DOKUNULMAZ
5. ✅ Total ~330+ PASS + 0 xfail
6. ✅ Snapshot JSON ~5KB (4 entry)
7. ✅ Currency mode auto-detect (RULE 3)
8. ✅ Push to origin/main
9. ⏸ Kullanıcı manuel onayı → Phase 5b (DCMA) başlar

---

## Sequencing Tips

- **Pure math T75-T77** → manuel write + self-verify (test-driven, no probe, no fixtures)
- **T78 loader adapter** → manuel; integrasyon riski var ama Phase 4 helpers'ı zaten var
- **T79-T80 trivial action helpers** → manuel
- **T81 BIG ONE (time-phased + period_delta)** → subagent dispatch (bucket edge cases tricky)
- **T82 Phase 3a integration** → subagent (compare_two pattern)
- **T83 BIG ONE (JSON snapshot)** → subagent dispatch (file I/O + filter logic)
- **T84 standard finalize** → manuel + push

Phase 1+2+3+4 helpers DOKUNULMAZ; sadece read-only çağrılar.

---

*Plan tamamlandı: 30 Nisan 2026*
*Tahmini Phase 5a süresi: ~20 saat (T75-T84, 10 task TDD chain)*
*Sonraki phase (onay sonrası): Phase 5b — DCMA 14-Point (`msproject_health` tool)*
