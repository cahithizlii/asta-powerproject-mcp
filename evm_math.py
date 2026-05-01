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


def _task_pv_at_date(task: Dict[str, Any], eval_date: _dt.date) -> float:
    """RULE 5 — Linear distribution per task. Hours OR cost (caller decides)."""
    bs = task.get("baseline_start")
    bf = task.get("baseline_finish")
    bw = float(task.get("baseline_work") or 0)
    if bs is None or bf is None or bw == 0:
        return 0.0
    # Coerce datetime to date if needed
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

    EV uses current percent_complete x baseline_work, but only counts tasks
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
    """RULE 6 — Haftalik/aylik delta. period_BAC = 0 (sabit)."""
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

    es_days: Optional[float] = None  # days since project_start

    # Below first point — clamp to 0
    first_date, first_pv = pv_curve[0]
    if ev_now <= first_pv:
        if first_pv > 0:
            frac = ev_now / first_pv
        else:
            frac = 0.0
        es_days = (first_date - project_start).days * frac

    # Search adjacent pair
    if es_days is None:
        for i in range(1, len(pv_curve)):
            t_prev, pv_prev = pv_curve[i - 1]
            t_curr, pv_curr = pv_curve[i]
            if pv_prev <= ev_now <= pv_curr:
                if pv_curr - pv_prev > 1e-9:
                    frac = (ev_now - pv_prev) / (pv_curr - pv_prev)
                else:
                    frac = 0.0
                delta_days = (t_curr - t_prev).days * frac
                es_days = (t_prev - project_start).days + delta_days
                break

    # Above last point — clamp to last
    if es_days is None:
        es_days = (pv_curve[-1][0] - project_start).days

    es_weeks = max(es_days / 7.0, 0.0)
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
