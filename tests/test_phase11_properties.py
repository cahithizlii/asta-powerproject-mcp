"""Phase 11.4 — Hypothesis-based property tests.

10 invariants on pure modules. Discovers bug classes invisible to
example-based testing. Each property runs ~100 random examples
internally (Hypothesis default).

Pure modules covered:
    - evm_math (time_phased_ac, time_phased_ac_increments, compute_metrics)
    - currency_validator (extract_currency_code, cross_validate_modes,
                          detect_mode_from_xer_assignments)
    - xer_compare (diff_tasks, diff_links, summarize_compare)
    - xer_parser (_parse_clndr_data)
"""
import datetime as dt
import math
import os
import sys

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from hypothesis import given, strategies as st, settings, HealthCheck

from evm_math import (
    time_phased_ac,
    time_phased_ac_increments,
    compute_metrics,
)
from currency_validator import (
    detect_mode_from_xer_assignments,
    extract_currency_code,
    cross_validate_modes,
)
from xer_compare import diff_tasks, diff_links, summarize_compare
from xer_parser import _parse_clndr_data


# ---------------------------------------------------------------------------
# Strategy helpers
# ---------------------------------------------------------------------------
DATE_STRATEGY = st.dates(min_value=dt.date(2020, 1, 1),
                         max_value=dt.date(2030, 12, 31))


@st.composite
def task_with_actuals(draw):
    """Generate a task dict with consistent actual_start <= actual_finish."""
    a_start = draw(DATE_STRATEGY)
    a_finish = draw(st.dates(min_value=a_start,
                             max_value=dt.date(2030, 12, 31)))
    aw = draw(st.floats(min_value=0, max_value=10000,
                        allow_nan=False, allow_infinity=False))
    return {"actual_start": a_start,
            "actual_finish": a_finish,
            "actual_work": aw}


@st.composite
def task_id_dict(draw):
    return {"id": draw(st.integers(min_value=1, max_value=1000)),
            "name": draw(st.text(min_size=0, max_size=20)),
            "percent_complete": draw(st.floats(min_value=0, max_value=100,
                                               allow_nan=False,
                                               allow_infinity=False))}


@st.composite
def link_dict(draw):
    return {"from_id": draw(st.integers(min_value=1, max_value=100)),
            "to_id": draw(st.integers(min_value=1, max_value=100)),
            "type": draw(st.sampled_from(["FS", "FF", "SS", "SF"])),
            "lag_days": draw(st.floats(min_value=-10, max_value=10,
                                       allow_nan=False,
                                       allow_infinity=False))}


# ---------------------------------------------------------------------------
# Property 1 — time_phased_ac cumulative non-decreasing
# ---------------------------------------------------------------------------
@settings(suppress_health_check=[HealthCheck.too_slow], deadline=None)
@given(tasks=st.lists(task_with_actuals(), min_size=1, max_size=5),
       n_buckets=st.integers(min_value=1, max_value=5))
def test_property_time_phased_ac_monotonic(tasks, n_buckets):
    """Cumulative AC must never decrease across consecutive buckets."""
    starts = [t["actual_start"] for t in tasks]
    finishes = [t["actual_finish"] for t in tasks]
    project_start = min(starts)
    project_finish = max(finishes)
    span_days = max((project_finish - project_start).days, 1)
    delta = max(span_days // n_buckets, 1)
    buckets = []
    d = project_start
    for _ in range(n_buckets):
        end = d + dt.timedelta(days=delta)
        buckets.append((d, end))
        d = end
    cum = time_phased_ac(tasks, buckets, project_finish)
    for i in range(1, len(cum)):
        assert cum[i] >= cum[i - 1] - 0.01, (
            f"AC decreased between bucket {i-1} and {i}: "
            f"{cum[i-1]} -> {cum[i]} (full curve: {cum})"
        )


# ---------------------------------------------------------------------------
# Property 2 — increments sum to cumulative final (round-trip)
# ---------------------------------------------------------------------------
@settings(suppress_health_check=[HealthCheck.too_slow], deadline=None)
@given(tasks=st.lists(task_with_actuals(), min_size=1, max_size=5),
       n_buckets=st.integers(min_value=1, max_value=5))
def test_property_increments_sum_to_cumulative_final(tasks, n_buckets):
    """sum(increments) must equal cumulative[-1] within tolerance."""
    starts = [t["actual_start"] for t in tasks]
    finishes = [t["actual_finish"] for t in tasks]
    project_start = min(starts)
    project_finish = max(finishes)
    span_days = max((project_finish - project_start).days, 1)
    delta = max(span_days // n_buckets, 1)
    buckets = []
    d = project_start
    for _ in range(n_buckets):
        end = d + dt.timedelta(days=delta)
        buckets.append((d, end))
        d = end
    cum = time_phased_ac(tasks, buckets, project_finish)
    inc = time_phased_ac_increments(tasks, buckets, project_finish)
    if cum:
        assert abs(sum(inc) - cum[-1]) < 0.01, (
            f"Increment round-trip violation: "
            f"sum(inc)={sum(inc)} vs cum[-1]={cum[-1]}"
        )


# ---------------------------------------------------------------------------
# Property 3 — diff_tasks partitions full ID union
# ---------------------------------------------------------------------------
@settings(deadline=None)
@given(a=st.lists(task_id_dict(), min_size=0, max_size=10,
                  unique_by=lambda d: d["id"]),
       b=st.lists(task_id_dict(), min_size=0, max_size=10,
                  unique_by=lambda d: d["id"]))
def test_property_diff_tasks_partition(a, b):
    """added + removed + changed + unchanged_count == |union of ids|."""
    r = diff_tasks(a, b)
    union_ids = {t["id"] for t in a} | {t["id"] for t in b}
    total = (len(r["added"]) + len(r["removed"])
             + len(r["changed"]) + r["unchanged_count"])
    assert total == len(union_ids), (
        f"Partition violated: total={total} vs union={len(union_ids)}, "
        f"r={r}"
    )


# ---------------------------------------------------------------------------
# Property 4 — diff_links symmetric (A->B.removed == B->A.added by identity)
# ---------------------------------------------------------------------------
@settings(deadline=None)
@given(a=st.lists(link_dict(), min_size=0, max_size=10),
       b=st.lists(link_dict(), min_size=0, max_size=10))
def test_property_diff_links_symmetric(a, b):
    """diff(A, B).removed identity-equal to diff(B, A).added (set-wise).

    Identity tuple is (from_id, to_id, type) — lag changes are reflected
    in 'changed', not removed/added.
    """
    fwd = diff_links(a, b)
    rev = diff_links(b, a)
    fwd_removed_keys = {(l["from_id"], l["to_id"], l["type"])
                        for l in fwd["removed"]}
    rev_added_keys = {(l["from_id"], l["to_id"], l["type"])
                      for l in rev["added"]}
    assert fwd_removed_keys == rev_added_keys, (
        f"Asymmetric diff: fwd.removed={fwd_removed_keys} "
        f"vs rev.added={rev_added_keys}"
    )


# ---------------------------------------------------------------------------
# Property 5 — cross_validate_modes idempotent (deterministic)
# ---------------------------------------------------------------------------
@settings(deadline=None)
@given(sources=st.lists(st.tuples(
    st.text(min_size=1, max_size=10),
    st.sampled_from(["cost", "hours", "mixed", "uncertain"]),
), min_size=0, max_size=5))
def test_property_cross_validate_modes_idempotent(sources):
    """Same input always yields same output (no random / no global state)."""
    r1 = cross_validate_modes(sources)
    r2 = cross_validate_modes(sources)
    assert r1 == r2, f"Non-deterministic: r1={r1} vs r2={r2}"


# ---------------------------------------------------------------------------
# Property 6 — extract_currency_code never raises for arbitrary dict
# ---------------------------------------------------------------------------
@settings(deadline=None)
@given(d=st.dictionaries(st.text(min_size=0, max_size=10),
                         st.text(min_size=0, max_size=10),
                         min_size=0, max_size=5))
def test_property_extract_currency_code_no_exception(d):
    """extract_currency_code never raises for any dict-shaped input."""
    extract_currency_code(d)
    extract_currency_code(None)
    extract_currency_code({})


# ---------------------------------------------------------------------------
# Property 7 — compute_metrics SPI/CPI finite for positive inputs
# ---------------------------------------------------------------------------
@settings(deadline=None)
@given(bac=st.floats(min_value=1, max_value=1e9,
                     allow_nan=False, allow_infinity=False),
       pv=st.floats(min_value=0.001, max_value=1e9,
                    allow_nan=False, allow_infinity=False),
       ev=st.floats(min_value=0, max_value=1e9,
                    allow_nan=False, allow_infinity=False),
       ac=st.floats(min_value=0.001, max_value=1e9,
                    allow_nan=False, allow_infinity=False))
def test_property_evm_compute_finite(bac, pv, ev, ac):
    """compute_metrics returns finite SPI/CPI when inputs are positive."""
    r = compute_metrics(bac=bac, pv=pv, ev=ev, ac=ac)
    spi = r.get("spi")
    cpi = r.get("cpi")
    # With pv > 0 and ac > 0, both should be defined and finite
    assert spi is not None, f"SPI None for pv={pv} ev={ev}"
    assert cpi is not None, f"CPI None for ac={ac} ev={ev}"
    assert math.isfinite(spi), f"SPI not finite: {spi}"
    assert math.isfinite(cpi), f"CPI not finite: {cpi}"


# ---------------------------------------------------------------------------
# Property 8 — _parse_clndr_data accepts any string, returns list
# ---------------------------------------------------------------------------
@settings(deadline=None)
@given(blob=st.text(min_size=0, max_size=200))
def test_property_parse_clndr_data_no_exception(blob):
    """_parse_clndr_data never raises and always returns a list."""
    r = _parse_clndr_data(blob)
    assert isinstance(r, list), f"Expected list, got {type(r)}"


# ---------------------------------------------------------------------------
# Property 9 — detect_mode_from_xer_assignments([]) -> 'uncertain'
# ---------------------------------------------------------------------------
@settings(deadline=None)
@given(assignments=st.just([]))
def test_property_xer_assignments_empty_uncertain(assignments):
    """Empty assignments deterministically classifies as 'uncertain'."""
    result = detect_mode_from_xer_assignments(assignments)
    assert result == "uncertain", (
        f"Expected 'uncertain' for empty input, got {result!r}"
    )


# ---------------------------------------------------------------------------
# Property 10 — summarize_compare empty diffs -> 'no changes detected'
# ---------------------------------------------------------------------------
def test_property_summarize_compare_no_changes_headline():
    """All-empty diffs always produce the canonical 'no changes' headline."""
    empty_task = {"added": [], "removed": [], "changed": [],
                  "unchanged_count": 0}
    empty_link = {"added": [], "removed": [], "changed": [],
                  "unchanged_count": 0}
    empty_progress = {"tasks": [],
                      "summary": {"count_moved": 0,
                                  "total_pct_delta": 0,
                                  "total_aw_delta": 0}}
    empty_evm = {"spi_a": None, "spi_b": None,
                 "spi_delta": None, "cpi_delta": None,
                 "ev_delta": None}
    r = summarize_compare(empty_task, empty_link, empty_progress, empty_evm)
    assert r["headline"] == "no changes detected", (
        f"Unexpected headline: {r['headline']!r}"
    )
