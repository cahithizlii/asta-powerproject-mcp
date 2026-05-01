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
