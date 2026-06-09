"""P2 #8 — Progress-Period reschedule risk guard (POLYBLMH lesson, RULE 17).

Pure decision function. No COM. Detects the exact failure signature that
broke POLYBLMH: a ReportDate set well AFTER the project start while the
project has ~zero progress — `proj.Reschedule()` then pushes every
zero-progress task past the data date (Root EE 2026-10-15 -> 2027-03-22).

The guard is advisory: callers decide whether to warn or block.
"""
import datetime as _dt


def _parse(s):
    if not s:
        return None
    s = str(s)
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"):
        try:
            return _dt.datetime.strptime(s[:10], fmt).date()
        except ValueError:
            continue
    try:
        return _dt.date.fromisoformat(s[:10])
    except (ValueError, TypeError):
        return None


def assess_reschedule_risk(report_date, project_start,
                           project_percent_complete=None,
                           gap_warn_days=30):
    """Assess POLYBLMH reschedule risk.

    Args:
        report_date: target ReportDate / data date (str or date).
        project_start: project start date (str or date).
        project_percent_complete: overall % complete (0-100) if known; None
            when the caller could not determine it (guard is then advisory
            only, not high-confidence).
        gap_warn_days: how far ReportDate must be after start to be suspect.

    Returns dict:
        risk: bool  (True = high-confidence POLYBLMH signature)
        severity: 'none' | 'low' | 'high'
        gap_days: int | None  (report_date - project_start)
        message: human warning (None when no risk)
        recommendation: short action text (None when no risk)
    """
    rd = _parse(report_date)
    ps = _parse(project_start)
    out = {"risk": False, "severity": "none", "gap_days": None,
           "message": None, "recommendation": None}
    if rd is None or ps is None:
        return out
    gap = (rd - ps).days
    out["gap_days"] = gap
    if gap <= gap_warn_days:
        return out  # ReportDate at/near start -> normal, no POLYBLMH risk

    pct = project_percent_complete
    rec = (f"Tools → Progress Periods: ReportDate'i proje başlangıcına "
           f"({ps.isoformat()}) çekmeyi düşünün; reschedule tüm progress=0 "
           f"görevleri data date'e (POLYBLMH) ötelemesin.")
    if pct is not None and float(pct) <= 0.0:
        # Exact POLYBLMH signature: late ReportDate + zero progress.
        out.update({
            "risk": True, "severity": "high",
            "message": (
                f"YÜKSEK RİSK (POLYBLMH/RULE 17): ReportDate ({rd.isoformat()}) "
                f"proje başlangıcından {gap} gün sonra ve proje ilerlemesi %0. "
                f"Reschedule TÜM görevleri data date sonrasına öteleyebilir."),
            "recommendation": rec,
        })
    elif pct is None:
        # Cannot confirm progress -> advisory low warning only.
        out.update({
            "risk": False, "severity": "low",
            "message": (
                f"DİKKAT: ReportDate ({rd.isoformat()}) proje başlangıcından "
                f"{gap} gün sonra. Proje ilerlemesi doğrulanamadı — eğer "
                f"görevlerde progress yoksa reschedule tarihleri öteleyebilir "
                f"(POLYBLMH/RULE 17)."),
            "recommendation": rec,
        })
    return out
