# Phase 11 — Detailed Test Suite Expansion Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Expand the MS Project MCP test suite from 569 PASS to ~720 across 4 atomic sub-phase commits (coverage gap-fill → edge cases → end-to-end scenarios → property-based) with **98% line coverage** on test-eligible code.

**Architecture:** Sub-phased atomic delivery. Each sub-phase is one commit + one push. Pure addition (zero production code change). Hypothesis (property tests) and pytest-cov (coverage measurement) added as dev deps. COM-only branches excluded via `# pragma: no cover` and `.coveragerc`.

**Tech Stack:** pytest, pytest-cov, hypothesis, existing test fixtures (`tests/conftest.py`, `tests/_xer_fixture_builders.py`, `tests/fixtures/sample_msp.xml`).

**Design doc:** `docs/plans/2026-05-04-phase11-detailed-tests-design.md` (committed `6d71bf9`).

**Baseline:** HEAD `6d71bf9`, suite 569 PASS @ ~24s.

---

## Task 1 — Phase 11.1 Coverage Gap-Fill

**Goal:** Install pytest-cov, write `.coveragerc`, measure baseline coverage, fill gaps to reach **98% line coverage** on pure modules + non-COM adapters.

**Files:**
- Create: `.coveragerc`
- Create/extend: tests for `evm_math.py`, `currency_validator.py`, `dcma_checks.py`, `xer_parser.py`, `xer_compare.py`, `mspdi_parser.py`
- Modify (if gaps found): `tests/test_*.py` per pure module

**Step 1: Install pytest-cov**

Run: `pip install pytest-cov`
Expected: successful install, version printed

**Step 2: Create `.coveragerc`**

Content:
```ini
[run]
source = .
omit =
    tests/*
    samples/*
    build_*.py
    asta_mcp_*.py
    nuke_cleanup.py
    docs/*
    setup.py

[report]
exclude_lines =
    pragma: no cover
    raise NotImplementedError
    if __name__ == .__main__.:
    pythoncom.CoInitialize
    win32com.client
    proj\.Application
    \.ActiveProject
    pywintypes\.Time
    cal\.Exceptions\.Add
    if TYPE_CHECKING:
    @overload
```

**Step 3: Run baseline coverage on pure modules**

Run:
```bash
cd /c/Users/CahAsus/asta-powerproject-mcp && python -m pytest \
  tests/test_evm_math*.py tests/test_currency_validator.py \
  tests/test_dcma_checks.py tests/test_xer_parser.py \
  tests/test_xer_compare.py tests/test_mspdi_baseline_write.py \
  --cov=evm_math --cov=currency_validator --cov=dcma_checks \
  --cov=xer_parser --cov=xer_compare --cov=mspdi_parser \
  --cov-report=term-missing 2>&1 | tail -30
```
Expected: per-module coverage table with "Missing" line numbers.

**Step 4: For each pure module below 98%, add tests targeting `Missing` lines**

For each module:
1. Read the missing line numbers
2. Add a test in the corresponding test file that exercises the path
3. Re-run coverage; iterate until module ≥ 98%

Pattern: write the test (failing or covering uncovered branch) → run targeted file → confirm coverage % delta → next.

**Step 5: Verify all pure modules ≥ 98%**

Run the same coverage command; assert each module ≥ 98%.

**Step 6: Run extended coverage on adapter helpers (non-COM)**

Run:
```bash
python -m pytest tests/ --cov=msproject_mcp_core \
  --cov-report=term-missing 2>&1 | tail -20
```
Expected: msproject_mcp_core coverage. COM-only branches will show as missing — those are excluded via `.coveragerc`. Add tests for adapter helpers (`_evm_load_*`, `_dcma_load_*`, `_excel_collect_*`, `_xer_to_*`) until coverage ≥ 98% on **non-excluded lines**.

**Step 7: Run full regression — must still PASS**

Run:
```bash
python -m pytest tests/test_msproject_file_*.py tests/test_evm_math*.py \
  tests/test_msproject_evm_*.py tests/test_dcma_*.py \
  tests/test_msproject_dcma_*.py tests/test_excel_io.py \
  tests/test_msproject_excel_*.py tests/test_xer_parser.py \
  tests/test_msproject_xer_*.py tests/test_phase5e_*.py \
  tests/test_phase5f_*.py tests/test_currency_validator.py \
  tests/test_mspdi_baseline_write.py tests/test_xer_calendar_holidays.py \
  tests/test_xer_compare.py tests/test_msproject_compare_dispatcher.py \
  -q --tb=line 2>&1 | tail -3
```
Expected: PASS count ≥ 569 + (added tests). Zero regression.

**Step 8: Commit**

```bash
git add .coveragerc tests/test_*.py
git commit -m "Phase 11.1 T141: coverage gap-fill to 98% on pure modules + adapters

NEW: .coveragerc — excludes COM-only branches via patterns
NEW: pytest-cov dev dependency

Coverage uplift on pure modules:
  evm_math.py            : <baseline>% -> 98%+
  currency_validator.py  : <baseline>% -> 98%+
  dcma_checks.py         : <baseline>% -> 98%+
  xer_parser.py          : <baseline>% -> 98%+
  xer_compare.py         : <baseline>% -> 98%+
  mspdi_parser.py (R+W)  : <baseline>% -> 98%+
  msproject_mcp_core     : <baseline>% -> 98%+ (non-COM lines)

Tests added to existing per-module test files; no new test files.

Regression: 569 -> NNN PASS, zero regression."
```

**Step 9: Push**

```bash
git push origin main
```

Expected: clean push.

**Verify:** Suite > 569 PASS, all pure modules ≥ 98% coverage on
non-COM lines.

---

## Task 2 — Phase 11.2 Edge Case + Negative Path

**Goal:** Each dispatcher action gains 5-10 negative/boundary tests. Distributes ~80 tests across existing per-tool test files.

**Files:** Modify existing per-tool test files:
- `tests/test_msproject_task_*.py`
- `tests/test_msproject_link.py`
- `tests/test_msproject_calendar_*.py`
- `tests/test_msproject_resource_*.py`
- `tests/test_msproject_baseline_*.py`
- `tests/test_msproject_progress_*.py`
- `tests/test_msproject_file_*.py`
- `tests/test_msproject_evm_*.py`
- `tests/test_msproject_dcma_*.py`
- `tests/test_msproject_excel_*.py`
- `tests/test_msproject_xer_*.py`
- `tests/test_msproject_compare_dispatcher.py`

**Step 1: For each tool, write negative/boundary tests in this pattern**

Test name: `test_<action>_<bad_input_kind>_returns_error`

```python
def test_<action>_<bad_input_kind>_returns_error():
    r = _call("<action>", <bad_field>=<bad_value>)
    assert r["status"] == "error"
    assert "<key_substring>" in r["error"].lower()
```

For each tool, add ~5-10 such tests. Priority list:

#### msproject_task (~5 tests)
- invalid task_id (-1, 0, max-int)
- duration string "" / None / "abc"
- malformed start date

#### msproject_link (~5)
- self-loop (from_id == to_id)
- type='XX'
- non-existent task_id

#### msproject_calendar (~10)
- bad date format ('2026/01/01', '2026-13-01')
- day_of_week typo ('xyz')
- recurrence='quarterly'
- working_hours bad HH:MM
- weekly without days_of_week

#### msproject_resource (~5)
- duplicate name (same call twice)
- negative max_units
- type='XX'

#### msproject_baseline (~5)
- invalid baseline_number (-1, 99)
- missing source/target

#### msproject_progress (~5)
- pct > 100 / < 0
- negative actual_work
- future actual_start vs status_date

#### msproject_file (~10)
- corrupted XML (random bytes)
- .pp/.xer rejection
- file_path None
- output_path missing extension
- write_baseline empty data

#### msproject_evm (~10)
- missing baseline file
- zero BAC
- missing required action params (file_path None)
- bucket invalid string

#### msproject_health DCMA (~5)
- file_path None
- missing tasks
- threshold override invalid

#### msproject_excel (~5)
- unwritable path
- output ext mismatch
- missing baseline_number

#### msproject_xer (~5)
- malformed ERMHDR (no \t separator)
- missing tables
- file_path .xml rejection

#### msproject_compare (~5)
- identical files
- one-side empty
- non-existent file_a/file_b

**Step 2: Run per-tool test files individually**

After each tool's tests added, run that tool's test file:
```bash
python -m pytest tests/test_<tool>_*.py -q --tb=short 2>&1 | tail -10
```
Expected: all new + old PASS.

**Step 3: Run full regression**

Same command as Task 1 Step 7. Expected: PASS ≥ Task 1 PASS + ~80.

**Step 4: Commit**

```bash
git add tests/
git commit -m "Phase 11.2 T142: edge case + negative path coverage across all 13 tools

~80 new negative/boundary tests distributed across existing per-tool
test files. Pattern: invalid input returns {status:'error', error:<msg>}
with key substring assertion.

Per-tool counts:
  msproject_task        : ~5  (id, duration, date)
  msproject_link        : ~5  (self-loop, bad type)
  msproject_calendar    : ~10 (date, day_of_week, recurrence, hours)
  msproject_resource    : ~5  (duplicate, negative units)
  msproject_baseline    : ~5  (number, missing)
  msproject_progress    : ~5  (pct bounds, work, dates)
  msproject_file        : ~10 (xml corruption, ext, params)
  msproject_evm         : ~10 (BAC, baseline, params)
  msproject_health      : ~5  (DCMA empty inputs)
  msproject_excel       : ~5  (path, ext)
  msproject_xer         : ~5  (header, tables, encoding)
  msproject_compare     : ~5  (files, params)

Zero regression."
```

**Step 5: Push**

```bash
git push origin main
```

**Verify:** Suite +~80 PASS over Task 1 baseline.

---

## Task 3 — Phase 11.3 End-to-End Acceptance Scenarios

**Goal:** 8 multi-tool workflow tests in one new file. Each scenario chains 3-5 dispatcher actions and asserts on the composite result.

**Files:**
- Create: `tests/test_phase11_e2e_acceptance.py`

**Step 1: Write the new test file with all 8 scenarios**

Use existing helpers:
- `tests._xer_fixture_builders.write_synthetic_xer`
- `samples._lifecycle_common.call_async_dispatcher` (re-import in test)
- Existing dispatcher imports from `msproject_mcp_core`

Test functions (one per scenario):

```python
"""Phase 11.3 — End-to-end acceptance scenarios.

Multi-tool workflow tests. Each scenario exercises 3-5 dispatcher
actions and asserts on the composite result.
"""
import asyncio
import json
import os
import sys
import tempfile

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from msproject_mcp_core import (
    msproject_compare, msproject_evm, msproject_health,
    msproject_excel, msproject_xer, msproject_file,
    _msp_evm_validate_currency_mode,
    _msp_file_update_task,
)
from mspdi_parser import MspdiProject
from tests._xer_fixture_builders import write_synthetic_xer


def _call(dispatcher, action, **kw):
    raw = asyncio.run(dispatcher({"action": action, **kw}))
    return json.loads(raw)


# Scenario 1: CAU monthly hakediş chain
def test_e2e_cau_monthly_hakedis_chain(tmp_path):
    """XER A → XER B → monthly_report w/ Excel → assert all outputs."""
    a = write_synthetic_xer(<SNAPSHOT_A_CONTENT>, "e2e_a.xer")
    b = write_synthetic_xer(<SNAPSHOT_B_CONTENT>, "e2e_b.xer")
    xlsx = str(tmp_path / "monthly.xlsx")
    r = _call(msproject_compare, "monthly_report",
              file_a=a, file_b=b, output_excel=xlsx)
    assert r["status"] == "ok"
    assert r["counts"]["tasks_added"] >= 1
    assert r["evm_a"]["rag"] in ("Red", "Amber", "Green")
    assert os.path.exists(xlsx)
    for p in (a, b):
        os.remove(p)


# Scenario 2: Baseline lifecycle (MSPDI parse → write → re-read)
def test_e2e_baseline_lifecycle(tmp_path):
    """write_baseline via file MCP → reload → variance asserted."""
    src = os.path.join(os.path.dirname(__file__), "fixtures", "sample_msp.xml")
    out = str(tmp_path / "baselined.xml")
    proj = MspdiProject(src)
    uid = next(iter(proj._task_elems.keys()))
    del proj
    r = _call(msproject_file, "write_baseline",
              file_path=src, baseline_number=0,
              baseline_data=[{
                  "task_uid": uid,
                  "baseline_start": "2026-09-01T08:00:00",
                  "baseline_finish": "2026-09-30T17:00:00",
              }],
              output_path=out)
    assert r["status"] == "ok"
    assert r["tasks_written"] == 1
    reloaded = MspdiProject(out)
    bls = reloaded.read_baselines(0)
    assert any(b["baseline_start"] == "2026-09-01T08:00:00" for b in bls)


# Scenario 3: Currency validation pipeline
def test_e2e_currency_validation_pipeline(tmp_path):
    """XER load → validate_currency_mode → consensus assertions."""
    xer = write_synthetic_xer(<HOURS_PATTERN_CONTENT>, "e2e_curr.xer")
    r = _msp_evm_validate_currency_mode(file_path=xer)
    assert r["status"] == "ok"
    assert r["primary_mode"] in ("hours", "cost", "mixed", "uncertain")
    assert r["currency_code"] == "USD"
    assert "consensus_mode" in r["cross_validation"]
    os.remove(xer)


# Scenario 4: Time-phased EVM full chain
def test_e2e_time_phased_evm_chain(tmp_path):
    """compute_metrics → time_phased_evm → period_delta → coherence."""
    xer = write_synthetic_xer(<STAGGERED_CONTENT>, "e2e_tp.xer")
    cm = _call(msproject_evm, "compute_metrics", file_path=xer)
    tp = _call(msproject_evm, "time_phased_evm", file_path=xer, bucket="month")
    assert cm["status"] == "ok"
    assert tp["status"] == "ok"
    # Sum of ac_increment in tp.buckets must equal final cum AC
    increments = [b["ac_increment"] for b in tp["buckets"]]
    assert abs(sum(increments) - max(b["ac"] for b in tp["buckets"])) < 0.5
    os.remove(xer)


# Scenario 5: DCMA + EVM combined health
def test_e2e_dcma_plus_evm_health(tmp_path):
    """assess_all + progress_data_quality consistency."""
    xer = write_synthetic_xer(<STAGGERED_CONTENT>, "e2e_health.xer")
    dcma = _call(msproject_health, "assess_all", file_path=xer)
    evm_q = _call(msproject_evm, "progress_data_quality", file_path=xer)
    assert dcma["status"] == "ok"
    assert evm_q["status"] == "ok"
    os.remove(xer)


# Scenario 6: Calendar recurring + assignment effect
def test_e2e_calendar_recurring_assignment(clean_test_project):
    """create → add_exception(weekly) → assign to task → schedule effect."""
    from msproject_mcp_core import (
        _msp_calendar_create, _msp_calendar_add_exception,
        _msp_calendar_assign_to_task, _msp_task_add,
    )
    _msp_calendar_create(name="E2ECal", base_calendar="Standard")
    _msp_calendar_add_exception(
        calendar_name="E2ECal", exception_name="Weekly Off",
        start="2026-01-01", finish="2026-12-31",
        recurrence="weekly", days_of_week=["fri", "sat"],
    )
    # Assign + verify (depends on existing helpers)
    # ...


# Scenario 7: Update_task baseline awareness (Phase 9.3 + 10.1)
def test_e2e_update_task_baseline_aware(tmp_path):
    """update_task with baseline_* → baseline_after assertion."""
    import shutil
    src = os.path.join(os.path.dirname(__file__), "fixtures", "sample_msp.xml")
    dst = str(tmp_path / "e2e_upd.xml")
    shutil.copy(src, dst)
    r = _msp_file_update_task(file_path=dst, task_id=1, fields={
        "duration": "5d",
        "baseline_start": "2026-12-01T08:00:00",
        "baseline_finish": "2026-12-31T17:00:00",
    })
    assert r["status"] == "ok"
    assert r["schedule_updated"] is True
    assert r["baseline_written"] == 1
    assert r["baseline_after"]["baseline_start"] == "2026-12-01T08:00:00"


# Scenario 8: Compare with baseline write back
def test_e2e_compare_after_baseline_write(tmp_path):
    """original XER → modify task → compare task_delta → diff captures change.

    Use synthetic XER pair where snapshot B has updated dates for one task.
    """
    a = write_synthetic_xer(<DATES_A>, "e2e_cmp_a.xer")
    b = write_synthetic_xer(<DATES_B>, "e2e_cmp_b.xer")  # one task date changed
    r = _call(msproject_compare, "task_delta", file_a=a, file_b=b)
    assert r["status"] == "ok"
    assert len(r["changed"]) >= 1
    for p in (a, b):
        os.remove(p)
```

(SNAPSHOT_A_CONTENT, etc., are written inline in the file as multi-line string constants — copy patterns from `test_msproject_compare_dispatcher.py` and `test_msproject_evm_time_phased_ac_integration.py`.)

**Step 2: Run new file**

Run:
```bash
python -m pytest tests/test_phase11_e2e_acceptance.py -v --tb=short 2>&1 | tail -25
```
Expected: 8 PASS (some may skip if `clean_test_project` requires MSP COM).

**Step 3: Full regression**

Same as Task 1 Step 7. Expected: PASS ≥ Task 2 + 8.

**Step 4: Commit**

```bash
git add tests/test_phase11_e2e_acceptance.py
git commit -m "Phase 11.3 T143: end-to-end acceptance scenarios (8 multi-tool workflows)

NEW: tests/test_phase11_e2e_acceptance.py with 8 scenarios:
  1. CAU monthly hakediş chain (compare → monthly_report → Excel)
  2. Baseline lifecycle (file MCP write_baseline → reload → assert)
  3. Currency validation pipeline (XER → validate_currency_mode)
  4. Time-phased EVM full chain (compute → time_phased → period_delta)
  5. DCMA + EVM combined health (assess_all + progress_data_quality)
  6. Calendar recurring + task assignment effect (Phase 10.2)
  7. Update_task baseline awareness (Phase 9.3 + 10.1 baseline_after)
  8. Compare with baseline write back

Zero regression. Suite size growth ~+8."
```

**Step 5: Push**

```bash
git push origin main
```

**Verify:** Suite +8 PASS over Task 2 baseline.

---

## Task 4 — Phase 11.4 Property-Based / Fuzz

**Goal:** Hypothesis-based property tests — discover invariant violations.

**Files:**
- Create: `tests/test_phase11_properties.py`

**Step 1: Install hypothesis**

Run: `pip install hypothesis`
Expected: successful install.

**Step 2: Write property test file**

```python
"""Phase 11.4 — Hypothesis-based property tests.

10 invariants on pure modules. Discovers bug classes invisible to
example-based testing.
"""
import datetime as dt
import os
import sys

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from hypothesis import given, strategies as st, assume, settings, HealthCheck

from evm_math import (
    time_phased_ac, time_phased_ac_increments, compute_metrics,
)
from currency_validator import (
    detect_mode_from_xer_assignments,
    detect_mode_from_tasks_resources,
    extract_currency_code,
    cross_validate_modes,
)
from xer_compare import diff_tasks, diff_links, summarize_compare
from xer_parser import _parse_clndr_data


# Strategy helpers
DATE_STRATEGY = st.dates(min_value=dt.date(2020, 1, 1),
                         max_value=dt.date(2030, 12, 31))


@st.composite
def task_with_actuals(draw):
    a_start = draw(DATE_STRATEGY)
    a_finish = draw(st.dates(min_value=a_start,
                             max_value=dt.date(2030, 12, 31)))
    aw = draw(st.floats(min_value=0, max_value=10000,
                        allow_nan=False, allow_infinity=False))
    return {"actual_start": a_start, "actual_finish": a_finish,
            "actual_work": aw}


@settings(suppress_health_check=[HealthCheck.too_slow])
@given(tasks=st.lists(task_with_actuals(), min_size=0, max_size=5),
       n_buckets=st.integers(min_value=1, max_value=5))
def test_property_time_phased_ac_monotonic(tasks, n_buckets):
    """Cumulative AC must never decrease across buckets."""
    if not tasks:
        return
    starts = [t["actual_start"] for t in tasks]
    finishes = [t["actual_finish"] for t in tasks]
    project_start = min(starts)
    project_finish = max(finishes)
    delta = max((project_finish - project_start).days // n_buckets, 1)
    buckets = []
    d = project_start
    for _ in range(n_buckets):
        end = d + dt.timedelta(days=delta)
        buckets.append((d, end))
        d = end
    cum = time_phased_ac(tasks, buckets, project_finish)
    for i in range(1, len(cum)):
        assert cum[i] >= cum[i-1] - 0.01, f"AC decreased: {cum}"


@given(tasks=st.lists(task_with_actuals(), min_size=1, max_size=5),
       n_buckets=st.integers(min_value=1, max_value=5))
def test_property_increments_sum_to_cumulative_final(tasks, n_buckets):
    """sum(increments) == cumulative[-1] (round-trip invariant)."""
    starts = [t["actual_start"] for t in tasks]
    finishes = [t["actual_finish"] for t in tasks]
    project_start = min(starts)
    project_finish = max(finishes)
    delta = max((project_finish - project_start).days // n_buckets, 1)
    buckets = []
    d = project_start
    for _ in range(n_buckets):
        end = d + dt.timedelta(days=delta)
        buckets.append((d, end))
        d = end
    cum = time_phased_ac(tasks, buckets, project_finish)
    inc = time_phased_ac_increments(tasks, buckets, project_finish)
    if cum:
        assert abs(sum(inc) - cum[-1]) < 0.01


@st.composite
def task_id_dict(draw):
    return {"id": draw(st.integers(min_value=1, max_value=1000)),
            "name": draw(st.text(min_size=0, max_size=20)),
            "percent_complete": draw(st.floats(min_value=0, max_value=100,
                                               allow_nan=False))}


@given(a=st.lists(task_id_dict(), min_size=0, max_size=10, unique_by=lambda d: d["id"]),
       b=st.lists(task_id_dict(), min_size=0, max_size=10, unique_by=lambda d: d["id"]))
def test_property_diff_tasks_partition(a, b):
    """len(added)+len(removed)+len(changed)+unchanged_count == |union of ids|."""
    r = diff_tasks(a, b)
    union_ids = {t["id"] for t in a} | {t["id"] for t in b}
    total = (len(r["added"]) + len(r["removed"])
             + len(r["changed"]) + r["unchanged_count"])
    assert total == len(union_ids)


@st.composite
def link_dict(draw):
    return {"from_id": draw(st.integers(min_value=1, max_value=100)),
            "to_id": draw(st.integers(min_value=1, max_value=100)),
            "type": draw(st.sampled_from(["FS", "FF", "SS", "SF"])),
            "lag_days": draw(st.floats(min_value=-10, max_value=10,
                                       allow_nan=False))}


@given(a=st.lists(link_dict(), min_size=0, max_size=10),
       b=st.lists(link_dict(), min_size=0, max_size=10))
def test_property_diff_links_symmetric(a, b):
    """diff(A, B).removed identity-equal to diff(B, A).added (set-wise)."""
    fwd = diff_links(a, b)
    rev = diff_links(b, a)
    fwd_removed_keys = {(l["from_id"], l["to_id"], l["type"])
                        for l in fwd["removed"]}
    rev_added_keys = {(l["from_id"], l["to_id"], l["type"])
                      for l in rev["added"]}
    assert fwd_removed_keys == rev_added_keys


@given(sources=st.lists(st.tuples(
    st.text(min_size=1, max_size=10),
    st.sampled_from(["cost", "hours", "mixed", "uncertain"]),
), min_size=0, max_size=5))
def test_property_cross_validate_modes_idempotent(sources):
    """cross_validate_modes(s) == cross_validate_modes(s) (deterministic)."""
    r1 = cross_validate_modes(sources)
    r2 = cross_validate_modes(sources)
    assert r1 == r2


@given(d=st.dictionaries(st.text(min_size=0, max_size=10),
                         st.text(min_size=0, max_size=10),
                         min_size=0, max_size=5))
def test_property_extract_currency_code_no_exception(d):
    """extract_currency_code never raises for arbitrary dict."""
    extract_currency_code(d)
    extract_currency_code(None)


@given(bac=st.floats(min_value=1, max_value=1e9, allow_nan=False, allow_infinity=False),
       pv=st.floats(min_value=0, max_value=1e9, allow_nan=False, allow_infinity=False),
       ev=st.floats(min_value=0, max_value=1e9, allow_nan=False, allow_infinity=False),
       ac=st.floats(min_value=0, max_value=1e9, allow_nan=False, allow_infinity=False))
def test_property_evm_compute_finite(bac, pv, ev, ac):
    """compute_metrics returns finite SPI/CPI when inputs are positive numbers."""
    import math
    r = compute_metrics(bac=bac, pv=pv, ev=ev, ac=ac)
    spi = r.get("spi")
    cpi = r.get("cpi")
    if spi is not None:
        assert math.isfinite(spi)
    if cpi is not None:
        assert math.isfinite(cpi)


@given(blob=st.text(min_size=0, max_size=200))
def test_property_parse_clndr_data_no_exception(blob):
    """_parse_clndr_data returns a list for any input string."""
    r = _parse_clndr_data(blob)
    assert isinstance(r, list)


@given(assignments=st.just([]))
def test_property_xer_assignments_empty_uncertain(assignments):
    """detect_mode_from_xer_assignments([]) deterministically returns 'uncertain'."""
    assert detect_mode_from_xer_assignments(assignments) == "uncertain"


def test_property_summarize_compare_no_changes_headline():
    """Empty diffs always produce 'no changes detected' headline."""
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
    assert r["headline"] == "no changes detected"
```

**Step 3: Run property tests**

Run:
```bash
python -m pytest tests/test_phase11_properties.py -v --tb=short 2>&1 | tail -20
```
Expected: 10 PASS. Each property runs ~100 random examples internally (Hypothesis default).

**Step 4: Full regression**

Same as Task 1 Step 7. Expected: PASS ≥ Task 3 baseline + ~10-18 (Hypothesis test functions).

**Step 5: Commit**

```bash
git add tests/test_phase11_properties.py
git commit -m "Phase 11.4 T144: hypothesis-based property tests (10 invariants)

NEW: hypothesis dev dependency
NEW: tests/test_phase11_properties.py — 10 property tests

Invariants asserted:
  1. time_phased_ac cumulative non-decreasing
  2. ac_increments sum equals cumulative final (round-trip)
  3. diff_tasks partitions full ID union (added+removed+changed+unchanged)
  4. diff_links symmetric (A→B.removed == B→A.added)
  5. cross_validate_modes idempotent (deterministic)
  6. extract_currency_code never raises for arbitrary dict
  7. compute_metrics SPI/CPI finite for positive inputs
  8. _parse_clndr_data accepts any string, returns list
  9. detect_mode_from_xer_assignments([]) deterministic 'uncertain'
  10. summarize_compare empty diffs -> 'no changes detected'

Zero regression. Each property runs ~100 random examples (Hypothesis
default). Discovers bugs example-based tests miss.

Suite end: ~720 PASS."
```

**Step 6: Push**

```bash
git push origin main
```

**Verify:** Suite reaches ~720 PASS, all property tests green.

---

## Task 5 — Phase 11 Wrap-up

**Goal:** Memory + final verification.

**Step 1: Run final regression**

```bash
python -m pytest tests/test_msproject_file_*.py tests/test_evm_math*.py \
  tests/test_msproject_evm_*.py tests/test_dcma_*.py \
  tests/test_msproject_dcma_*.py tests/test_excel_io.py \
  tests/test_msproject_excel_*.py tests/test_xer_parser.py \
  tests/test_msproject_xer_*.py tests/test_phase5e_*.py \
  tests/test_phase5f_*.py tests/test_currency_validator.py \
  tests/test_mspdi_baseline_write.py tests/test_xer_calendar_holidays.py \
  tests/test_xer_compare.py tests/test_msproject_compare_dispatcher.py \
  tests/test_phase11_e2e_acceptance.py tests/test_phase11_properties.py \
  tests/test_msproject_calendar_exception.py \
  -q --tb=line 2>&1 | tail -3
```
Expected: ~720 PASS.

**Step 2: Run final coverage report**

```bash
python -m pytest tests/ --cov=. --cov-report=term-missing 2>&1 | tail -30
```
Expected: pure modules ≥ 98%, msproject_mcp_core non-COM lines ≥ 98%.

**Step 3: Update memory `MEMORY.md` index entry**

File: `C:/Users/CahAsus/.claude/projects/C--Users-CahAsus-Downloads/memory/MEMORY.md`

Update the project_msproject_mcp_phase2 line entry with Phase 11 wrap details (HEAD, PASS count, sub-phase commits).

**Step 4: No additional commit needed** (memory is in claude-config repo, not msproject repo).

---

## Verification checklist (final)

- [ ] All 4 sub-phase commits pushed to origin/main
- [ ] Suite size ~720 PASS, zero regression
- [ ] Pure modules ≥ 98% line coverage on test-eligible code
- [ ] `.coveragerc` committed in Task 1
- [ ] `pytest-cov` and `hypothesis` in dev deps (no setup.py change needed for now; documented in plan)
- [ ] `tests/test_phase11_e2e_acceptance.py` exercises 8 multi-tool workflows
- [ ] `tests/test_phase11_properties.py` exercises 10 invariants via hypothesis
- [ ] All Phase 1-10 DOKUNULMAZ contracts honored
- [ ] Memory MEMORY.md index entry updated

---

## DRY / YAGNI / TDD notes

- **DRY:** Reuses existing fixtures (`tests/conftest.py`,
  `tests/_xer_fixture_builders.py`, `tests/fixtures/sample_msp.xml`)
  and helpers (`samples/_lifecycle_common.py`).
- **YAGNI:** No new tools, no new dispatcher actions, no new pure
  modules. Pure addition of tests.
- **TDD:** Property tests follow TDD spirit (write invariant assertion,
  then verify on existing implementation). Coverage gap-fill tests are
  written to fail first (uncovered branch) then pass (after exercising
  the path).
- **Frequent commits:** 4 atomic sub-phase commits, each independently
  shippable + revertible.
