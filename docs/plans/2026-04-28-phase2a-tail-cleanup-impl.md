# Phase 2a TAIL Cleanup Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Address 9 deferred TAIL items + 1 untracked file from Phase 2a per-task code reviews. 4 ardışık commit, all on `main`.

**Architecture:** Modify `msproject_mcp_core.py` (calendar functions only — Phase 1 functions untouched). Add 1-2 small new test files. Each commit independently revertable. Phase 1 SAFETY pattern preserved (`clean_test_project` fixture).

**Tech Stack:** Python 3.12, mcp (FastMCP), pywin32 COM, pytest. Mevcut `msproject_mcp_core.py` (~1280 satır), 11 test dosyası, 83 test PASS baseline.

**Design doc:** `docs/plans/2026-04-28-phase2a-tail-cleanup-design.md` (commit `ffa8576`)

**Baseline state at start:** HEAD `ffa8576` (design doc commit), 83/83 tests PASS, MS Project running.

---

## Task 28: Contract/Behavior Cleanup (3 sub-items)

**Files:**
- Modify: `msproject_mcp_core.py` (lines 307-365 add_exception, lines 415-437 list)
- Modify: `tests/test_msproject_calendar_list.py`
- Modify: `tests/test_msproject_calendar_exception.py` (if any test asserts on `working` field)

### Sub-item A: `uid` → `calendar_uid` in `_msp_calendar_list`

**Step 1: Read the current list function and tests**

```bash
cd C:\Users\CahAsus\asta-powerproject-mcp
```

In `msproject_mcp_core.py:430`, current code reads:
```python
out.append({
    "uid": cal.Guid,
    "name": cal.Name,
    "exception_count": ex_count,
})
```

**Step 2: Edit msproject_mcp_core.py:430**

Change `"uid": cal.Guid,` to `"calendar_uid": cal.Guid,`. Use Edit tool — exact replacement.

**Step 3: Verify no test currently asserts on `c["uid"]`**

```bash
python -c "import subprocess; r = subprocess.run(['grep', '-rn', 'c\\[.uid.\\]', 'tests/'], capture_output=True, text=True); print(r.stdout, r.stderr)"
```

If grep finds matches, update those tests. If empty, no test changes needed for sub-item A (the Sub-item C `count` assertion will exercise the new key indirectly).

### Sub-item B: Remove dead shift loop in `_msp_calendar_add_exception`

**Step 4: Locate and delete the dead loop**

In `msproject_mcp_core.py`, find the block around line 343-355 that looks like:
```python
        # Mark non-working: zero out all shifts (each setattr in its own try/except
        # since some MSP versions reject int 0 — the Type=PJ_EXCEPTION_DAILY default
        # already implies non-working but we belt-and-suspenders set shifts too)
        for prop in ("Shift1Start", "Shift1Finish", "Shift2Start", "Shift2Finish",
                     "Shift3Start", "Shift3Finish"):
            try:
                setattr(ex, prop, 0)
            except Exception:
                pass
```

Replace with a single-line comment:
```python
        # Type=PJ_EXCEPTION_DAILY=7 already implies non-working in MSP semantics;
        # MSP 16.0 exposes shift times via ex.Shift1.Start sub-objects (not flat
        # ShiftNStart props), so any zeroing is best handled if/when working=True
        # support arrives in Phase 3+.
```

Use Edit tool with the full multi-line `for prop in (...)` block as `old_string` to ensure uniqueness.

### Sub-item C: Drop vestigial `working=working` field from add_exception payload

**Step 5: Edit msproject_mcp_core.py:360**

Locate the success return in `_msp_calendar_add_exception` (around lines 356-360):
```python
        return {"status": "ok",
                "calendar_name": calendar_name,
                "exception_name": exception_name,
                "start": start,
                "finish": finish or start,
                "working": working}
```

Remove the `"working": working` line entirely. New return:
```python
        return {"status": "ok",
                "calendar_name": calendar_name,
                "exception_name": exception_name,
                "start": start,
                "finish": finish or start}
```

**Step 6: Check for tests that assert on `working` in payload**

```bash
python -c "import subprocess; r = subprocess.run(['grep', '-rn', 'working.*]', 'tests/test_msproject_calendar_exception.py'], capture_output=True, text=True); print(r.stdout)"
```

If any test asserts `assert r["working"] == False` or similar, remove that assertion. The `test_add_exception_working_true_rejected` test asserts on the ERROR path — it stays unchanged.

### Sub-item Test: count assertion + key check

**Step 7: Edit `tests/test_msproject_calendar_list.py`**

In `test_list_includes_standard`, after `assert "Standard" in names`, add:
```python
    assert r["count"] == len(r["calendars"]) >= 1
    # Verify new key name (renamed from "uid" in T28)
    assert all("calendar_uid" in c for c in r["calendars"])
```

This both validates Sub-item A's rename AND adds the count assertion (T30 Sub-item B is now folded in here for atomicity — see note below).

NOTE: This `count` assertion was originally planned as a T30 polish item but landing it now alongside the `uid` rename keeps the test in sync with the contract change. T30 will skip the redundant count assertion.

### Sub-item Test: ensure existing tests still pass with shift-loop removed

**Step 8: Run target test files**

```bash
python -m pytest tests/test_msproject_calendar_exception.py tests/test_msproject_calendar_list.py -v
```

Expected: ALL PASS. The existing `test_add_exception_actually_non_working` (T21 fix) verifies non-working contract via `cal.Period(date).Working is False` — independent of the deleted loop, so it must still pass.

If `test_add_exception_actually_non_working` fails, the loop wasn't dead after all (unlikely — review confirmed it). In that case, STOP and investigate.

**Step 9: Run full regression**

```bash
python -m pytest tests/ -v --tb=short -q
```

Expected: **83 PASSED** (same as baseline — these are contract-preserving changes).

**Step 10: Commit**

```bash
git add msproject_mcp_core.py tests/test_msproject_calendar_list.py tests/test_msproject_calendar_exception.py
git commit -m "Phase 2a T28 (TAIL): contract cleanup (uid->calendar_uid, drop dead shift loop + vestigial working field)"
```

---

## Task 29: DX Cleanup (2 sub-items)

**Files:**
- Modify: `msproject_mcp_core.py` (add `_format_com_error` helper, replace 7 `str(e)` sites in calendar fns, modify dispatcher around lines 1239-1246)
- Create: `tests/test_msproject_format_com_error.py` (unit test)
- Modify: `tests/test_msproject_calendar_dispatcher.py` (add alias test)

### Sub-item A: `_format_com_error(e)` helper

**Step 1: Failing test — write the unit test FIRST**

Create `tests/test_msproject_format_com_error.py`:
```python
"""Unit test for _format_com_error helper (no MS Project required)."""
import pytest
from msproject_mcp_core import _format_com_error


def test_format_pywintypes_com_error_tuple():
    """pywintypes.com_error has args = (hresult, msg, excepinfo, argerr).
    excepinfo[2] is the human-readable description."""
    class FakeComError(Exception):
        def __init__(self):
            self.args = (
                -2147352567,
                "Exception occurred.",
                (0, "Microsoft Project", "The calendar name already exists.",
                 None, 0, -2147352567),
                None,
            )
    result = _format_com_error(FakeComError())
    assert result == "The calendar name already exists."


def test_format_falls_back_to_args1_when_excepinfo_empty_description():
    """If excepinfo[2] (description) is None/empty, fall back to args[1]."""
    class FakeComError(Exception):
        def __init__(self):
            self.args = (
                -2147352567,
                "Exception occurred.",
                (0, "Microsoft Project", None, None, 0, -2147352567),
                None,
            )
    result = _format_com_error(FakeComError())
    assert result == "Exception occurred."


def test_format_plain_exception_uses_str():
    """Plain Python exception falls through to str(e)."""
    e = ValueError("bad input")
    assert _format_com_error(e) == "bad input"


def test_format_empty_args_does_not_crash():
    """Edge: exception with no args."""
    class Empty(Exception):
        pass
    e = Empty()
    result = _format_com_error(e)
    assert isinstance(result, str)  # whatever it returns, must be a string


def test_format_args_too_short_for_excepinfo():
    """args has only 2 elements, no excepinfo tuple."""
    class Short(Exception):
        def __init__(self):
            self.args = ("just one message",)
    result = _format_com_error(Short())
    assert "just one message" in result
```

**Step 2: Run — FAIL** (ImportError)

```bash
python -m pytest tests/test_msproject_format_com_error.py -v
```

Expected: ImportError on `_format_com_error`.

**Step 3: Add helper to `msproject_mcp_core.py`**

Place the helper at the top of the file's helper section, right after `_route_operation` and before the section divider for task helpers (around line 100-105 — use Grep to find `_route_operation` to locate). Insert:

```python
def _format_com_error(e: Exception) -> str:
    """Extract a human-readable message from pywintypes.com_error or fallback to str(e).

    pywintypes.com_error.args = (hresult, msg, excepinfo, argerr) where
    excepinfo = (wCode, source, description, helpFile, helpContext, scode).
    Description (excepinfo[2]) is the user-friendly message; fall back to
    msg (args[1]) or str(e) if unavailable.
    """
    try:
        if hasattr(e, "args") and len(e.args) >= 3 and isinstance(e.args[2], tuple):
            excepinfo = e.args[2]
            if len(excepinfo) >= 3 and excepinfo[2]:
                return str(excepinfo[2]).strip()
            if len(e.args) >= 2 and e.args[1]:
                return str(e.args[1]).strip()
    except Exception:
        pass
    return str(e)
```

**Step 4: Run unit test — PASS**

```bash
python -m pytest tests/test_msproject_format_com_error.py -v
```

Expected: 5 PASSED.

**Step 5: Replace `str(e)` in 7 calendar function sites**

Use Edit tool to replace each. The 7 lines are at (approximate):
- Line 258: `_msp_calendar_create`
- Line 295: `_msp_calendar_update`
- Line 364 (or near): `_msp_calendar_add_exception` (the line shifts due to T28's dead loop removal)
- Line 384: `_msp_calendar_assign_to_task`
- Line 412: `_msp_calendar_assign_to_resource`
- Line 437: `_msp_calendar_list`
- And 1 inside `_msp_calendar_holidays_uzbek` (use Grep first)

For each site, replace `str(e)` with `_format_com_error(e)` in the error return. Pattern:
```python
# Before
return {"status": "error", "error": str(e)}
# After
return {"status": "error", "error": _format_com_error(e)}
```

**IMPORTANT:** Phase 1 functions also have `str(e)` on many lines (530, 544, 904, 945, 1046, 1065, 1076, 1102) and dispatcher try/except (1152, 1189, 1223, 1268). **DO NOT TOUCH THESE.** Only the calendar `_msp_*` functions are in scope. The dispatcher line 1268 (`msproject_calendar` dispatcher) is also OUT of scope — the underlying helpers already format properly via `_format_com_error`, the dispatcher catches its own internal errors which are different.

After Edit, verify only 7 calendar sites changed:
```bash
git diff msproject_mcp_core.py | grep -E "^\\+|^\\-" | grep -E "str\\(e\\)|_format_com_error" | head -20
```

Expected: 7 `-str(e)` removals + 7 `+_format_com_error(e)` additions.

**Step 6: Full regression**

```bash
python -m pytest tests/ -v --tb=short -q
```

Expected: **88 PASSED** (83 baseline + 5 new format tests). All existing tests pass because the error message format change is backward-compatible (tests use `.lower()` substring matching).

### Sub-item B: Dispatcher `name`/`calendar_name` alias

**Step 7: Failing test**

Append to `tests/test_msproject_calendar_dispatcher.py`:
```python
def test_dispatcher_calendar_name_alias_for_add_exception(clean_test_project):
    """Dispatcher accepts 'name' as alias for 'calendar_name' on actions
    that natively expect calendar_name."""
    _run(msproject_calendar({
        "action": "create",
        "name": "AliasCal-T29",
        "base_calendar": "Standard",
    }))
    # Now use 'name' instead of 'calendar_name' for add_exception
    r = _run(msproject_calendar({
        "action": "add_exception",
        "name": "AliasCal-T29",         # alias for calendar_name
        "exception_name": "Aliased Holiday",
        "start": "2026-06-15",
    }))
    parsed = json.loads(r)
    assert parsed["status"] == "ok"
    assert parsed["calendar_name"] == "AliasCal-T29"


def test_dispatcher_name_alias_for_holidays_uzbek(clean_test_project):
    """holidays_uzbek accepts 'name' instead of 'calendar_name'."""
    _run(msproject_calendar({
        "action": "create",
        "name": "AliasUzbek-T29",
        "base_calendar": "Standard",
    }))
    r = _run(msproject_calendar({
        "action": "holidays_uzbek",
        "name": "AliasUzbek-T29",       # alias
        "year": 2026,
    }))
    parsed = json.loads(r)
    assert parsed["status"] in ("ok", "partial")
    assert parsed["count"] == 9


def test_dispatcher_calendar_name_alias_reverse(clean_test_project):
    """Reverse direction: 'create' accepts 'calendar_name' as alias for 'name'."""
    r = _run(msproject_calendar({
        "action": "create",
        "calendar_name": "ReverseAlias-T29",   # alias for name
        "base_calendar": "Standard",
    }))
    parsed = json.loads(r)
    assert parsed["status"] == "ok"
```

**Step 8: Run — FAIL** (alias not implemented)

```bash
python -m pytest tests/test_msproject_calendar_dispatcher.py -v
```

Expected: 3 new tests FAIL with "unexpected keyword argument 'name'" (or similar TypeError).

**Step 9: Add alias logic to dispatcher**

In `msproject_mcp_core.py`, locate `async def msproject_calendar(params: dict) -> str:` (around line 1231). After `p = {k: v for k, v in params.items() if k != "action"}` (around line 1248), insert:

```python
    # Alias: accept 'name' / 'calendar_name' interchangeably across actions
    NAME_ALIAS_ACTIONS = {"add_exception", "assign_to_task", "assign_to_resource",
                          "list", "holidays_uzbek"}
    NAME_NATIVE_ACTIONS = {"create", "update"}
    if action in NAME_ALIAS_ACTIONS and "name" in p and "calendar_name" not in p:
        p["calendar_name"] = p.pop("name")
    elif action in NAME_NATIVE_ACTIONS and "calendar_name" in p and "name" not in p:
        p["name"] = p.pop("calendar_name")
```

**Step 10: Run dispatcher tests — PASS**

```bash
python -m pytest tests/test_msproject_calendar_dispatcher.py -v
```

Expected: 7 PASSED (4 original + 3 new alias tests).

**Step 11: Full regression**

```bash
python -m pytest tests/ -v --tb=short -q
```

Expected: **91 PASSED** (88 + 3 new alias tests).

**Step 12: Commit T29**

```bash
git add msproject_mcp_core.py tests/test_msproject_format_com_error.py tests/test_msproject_calendar_dispatcher.py
git commit -m "Phase 2a T29 (TAIL): DX cleanup (_format_com_error helper + dispatcher name/calendar_name alias)"
```

---

## Task 30: Test/Doc Polish (3 sub-items remaining; count assertion was folded into T28)

**Files:**
- Modify: `msproject_mcp_core.py` (docstring on `_msp_calendar_list`, debug log in `_msp_calendar_holidays_uzbek`)
- Modify: `tests/test_msproject_calendar_create.py` (monkeypatch test)

### Sub-item A: T19 monkeypatch test for "succeeded but not found" guard

**Step 1: Append failing test to `tests/test_msproject_calendar_create.py`**

```python
def test_create_succeeded_but_not_found_guard(clean_test_project, monkeypatch):
    """Cover the 'BaseCalendarCreate succeeded but not found' branch.

    Stubs app.BaseCalendarCreate to a no-op (does nothing). The pre-flight
    name check passes (calendar doesn't exist), then BaseCalendarCreate
    silently returns, then the post-create lookup fails -> guard fires.
    """
    import msproject_mcp_core as core
    real_connect = core._connect_app
    real_app = real_connect()

    class _NoOp:
        def BaseCalendarCreate(self, **kwargs):
            return None  # silently no-op
        def __getattr__(self, name):
            return getattr(real_app, name)

    fake_app = _NoOp()
    monkeypatch.setattr(core, "_connect_app", lambda: fake_app)

    r = core._msp_calendar_create(name="GuardCal-T30", base_calendar="Standard")
    assert r["status"] == "error"
    assert "succeeded but" in r["error"].lower() or "not found" in r["error"].lower()
```

NOTE on the `__getattr__` proxy: `_msp_calendar_create` calls `_validate_active_project()` which calls `_connect_app()` then accesses `app.ActiveProject`. The proxy delegates all OTHER calls to the real app so validation passes; only `BaseCalendarCreate` is stubbed.

**Step 2: Run — FAIL** (test should produce the guard error message)

If the implementation already covers the branch correctly, the test will PASS immediately on first run (which is also fine — it means the guard works, we just lacked coverage). If it FAILs unexpectedly, debug:
```bash
python -m pytest tests/test_msproject_calendar_create.py::test_create_succeeded_but_not_found_guard -v -s
```

Expected: PASS (the guard exists at line 252-254 of `_msp_calendar_create`; this test exercises it).

If the proxy approach has issues with COM (e.g., `_validate_active_project` re-invokes `_connect_app` from a different code path), simplify by patching `app.BaseCalendarCreate` directly:
```python
def fake_create(**kwargs):
    return None
monkeypatch.setattr(real_app, "BaseCalendarCreate", fake_create)
```
Note: `monkeypatch.setattr` on a COM object may not work — fall back to the proxy approach if needed.

### Sub-item B: T24 docstring — list ordering note

**Step 3: Edit `_msp_calendar_list` docstring**

In `msproject_mcp_core.py:415-417`, current:
```python
def _msp_calendar_list() -> Dict[str, Any]:
    """List all base calendars in the active project with exception counts."""
```

Update to:
```python
def _msp_calendar_list() -> Dict[str, Any]:
    """List all base calendars in the active project with exception counts.

    Order: matches `proj.BaseCalendars` enumeration (typically insertion
    order, not sorted). If callers need lexicographic ordering, sort
    client-side on the `name` field.
    """
```

### Sub-item C: T25 debug log — pre-scan exception swallow

**Step 4: Edit `_msp_calendar_holidays_uzbek` pre-scan loop**

In `msproject_mcp_core.py`, locate the pre-scan loop in `_msp_calendar_holidays_uzbek` (around lines 453-461). Current:
```python
    # Pre-scan existing exception names for dedup
    existing_names = set()
    try:
        for i in range(1, cal.Exceptions.Count + 1):
            ex = cal.Exceptions(i)
            if ex is not None and ex.Name:
                existing_names.add(ex.Name)
    except Exception:
        pass  # if we can't read exceptions, treat as none-existing
```

Update the bare `except Exception: pass` to add a debug log:
```python
    except Exception as e:
        logger.debug(f"holidays_uzbek pre-scan failed (treating calendar as empty): {_format_com_error(e)}")
```

(Use `_format_com_error` from T29 — DRY.)

**Step 5: Run target tests**

```bash
python -m pytest tests/test_msproject_calendar_create.py tests/test_msproject_calendar_list.py tests/test_msproject_calendar_uzbek.py -v
```

Expected: ALL PASS, including the new `test_create_succeeded_but_not_found_guard`.

**Step 6: Full regression**

```bash
python -m pytest tests/ -v --tb=short -q
```

Expected: **92 PASSED** (91 + 1 new monkeypatch test).

**Step 7: Commit T30**

```bash
git add msproject_mcp_core.py tests/test_msproject_calendar_create.py
git commit -m "Phase 2a T30 (TAIL): test/doc polish (T19 guard test, T24 ordering docstring, T25 debug log)"
```

---

## Task 31: Add Untracked Tools Helper

**Files:**
- Add (existing untracked): `tools/export_empty_msp_fixture.py`

**Step 1: Verify the file exists and is the expected helper**

```bash
cd C:\Users\CahAsus\asta-powerproject-mcp
git status --short
```

Expected output includes: `?? tools/export_empty_msp_fixture.py`

```bash
head -10 tools/export_empty_msp_fixture.py
```

Expected: docstring `"""One-shot helper: export the active MS Project document as an MSPDI XML fixture."""`

**Step 2: Add and commit**

```bash
git add tools/export_empty_msp_fixture.py
git commit -m "Add tools/export_empty_msp_fixture.py (Phase 1 T3 fixture regenerator)"
```

**Step 3: Verify clean working tree**

```bash
git status
```

Expected: `nothing to commit, working tree clean`

---

## Push to GitHub (Final Step After All 4 Commits)

**Step 1: Verify commit chain**

```bash
git log --oneline -6
```

Expected (top to bottom):
```
<sha> Add tools/export_empty_msp_fixture.py (Phase 1 T3 fixture regenerator)
<sha> Phase 2a T30 (TAIL): test/doc polish (T19 guard test, T24 ordering docstring, T25 debug log)
<sha> Phase 2a T29 (TAIL): DX cleanup (_format_com_error helper + dispatcher name/calendar_name alias)
<sha> Phase 2a T28 (TAIL): contract cleanup (uid->calendar_uid, drop dead shift loop + vestigial working field)
ffa8576 Phase 2a TAIL cleanup design doc (approved)
42574f2 Phase 2a T27: end-to-end Uzbekistan calendar acceptance + README
```

**Step 2: Push**

```bash
git push origin main
```

Expected: 5 commits pushed (4 TAIL + 1 design doc), `42574f2..<new_head>  main -> main`.

**Step 3: Verify GitHub state**

```bash
git status
```

Expected: `Your branch is up to date with 'origin/main'.`

---

## TAIL Cleanup Tamamlama Kriterleri

1. ✅ T28: contract cleanup committed (`uid → calendar_uid`, dead shift loop removed, vestigial `working` field dropped)
2. ✅ T29: DX cleanup committed (`_format_com_error` helper + 7 site replacements + dispatcher alias)
3. ✅ T30: test/doc polish committed (monkeypatch guard test, ordering docstring, debug log)
4. ✅ T31: untracked helper file committed
5. ✅ Full regression: ~92 PASSED (83 baseline + 5 com_error + 3 alias + 1 guard test)
6. ✅ All 4 commits pushed to origin/main
7. ✅ Phase 2a tertemiz, ready for Phase 2b brainstorm

---

*Plan tamamlandı: 28 Nisan 2026*
*Tahmini TAIL cleanup süresi: ~30-45 dk (T28-T31, 4 commit)*
*Sonraki phase: Phase 2b — Resource Management brainstorm*
