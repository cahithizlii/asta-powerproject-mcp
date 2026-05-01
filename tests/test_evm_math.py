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
