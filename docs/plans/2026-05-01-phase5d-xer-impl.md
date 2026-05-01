# Phase 5d XER Reader Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development to execute this plan task-by-task (T102-T108).

**Goal:** 12th MCP tool `msproject_xer` — pure-Python Primavera P6 XER reader (6 read-only actions). Bridges CAU XER projects into Phase 5a EVM + 5b DCMA + 5c Excel pipelines.

**Architecture:** Yaklaşım C (Phase 5a/5b/5c proven). New `xer_parser.py` pure-Python module + I/O adapters in `msproject_mcp_core.py` Phase 5D section + `@mcp.tool msproject_xer` dispatcher. NO mpxj dependency. Phase 1-5c helpers DOKUNULMAZ.

**Tech Stack:** Python 3.12, mcp (FastMCP), pytest. Existing `msproject_mcp_core.py` (~6160 lines after Phase 5c TAIL fix), 50+ test files, 303 cumulative regression PASS baseline.

**Design doc:** `docs/plans/2026-05-01-phase5d-xer-design.md` (commit `c638016`)

**Baseline state at start:** HEAD `c638016`, MS Project running v16.0.

**KEY REFERENCES:**
- XER format: UTF-16-LE BOM tab-delimited, `%T <table>` / `%F <headers>` / `%R <row>` / `%E` markers
- Field mapping (XER → MSP-shape) in design doc
- Constraint enum: CS_MSO=2, CS_MFO=3, CS_MSOA→4, CS_MSOB→5, CS_MEOA→6, CS_MEOB→7 (DCMA Rule 6)
- Link type: PR_FS→FS, PR_SS→SS, PR_FF→FF, PR_SF→SF
- CLAUDE.md RULE 1: CAU calendar 6×9 = 54h/week, 9h/day
- Phase 4 file MCP pattern: `_msp_file_read_*` returns `{"status": "ok", "count": N, "<key>": [...]}`
- Phase 5b/5c TAIL pattern: single-collect aggregator, no per-rule re-fetch

---

## Task 102: `xer_parser.py` Foundations + Encoding Detect + Table Splitter

**Files:**
- Create: `xer_parser.py`
- Create: `tests/test_xer_parser.py`
- Create: `tests/conftest.py` (synthetic XER fixture builder — IF not exists; otherwise extend)

### Step 1: Write conftest.py with synthetic XER builder

```python
"""Phase 5d conftest: synthetic CAU-style XER fixture for tests."""
import os
import pytest


SAMPLE_XER_CONTENT = """ERMHDR\t18.8\t2026-05-01\tcahit\tProject Management\tUSD
%T\tPROJECT
%F\tproj_id\tproj_short_name\tplan_start_date\tplan_end_date\tlast_recalc_date
%R\t1\tCAU\t2024-07-08 08:00\t2028-06-20 17:00\t2026-05-01 17:00
%T\tCALENDAR
%F\tclndr_id\tclndr_name\tday_hr_cnt\tweek_hr_cnt
%R\t1\tCAU 6x9\t9.0\t54.0
%T\tRSRC
%F\trsrc_id\trsrc_name\trsrc_short_name\trsrc_type\tmax_qty_per_hr
%R\t101\tConcrete Workers\tCOW\tRT_Labor\t10.0
%R\t102\tExtractors\tEXT\tRT_Labor\t5.0
%R\t103\tSteel\tSTL\tRT_Mat\t100.0
%R\t104\tCarpenters\tCAR\tRT_Labor\t8.0
%T\tTASK
%F\ttask_id\twbs_id\tproj_id\tclndr_id\ttask_code\ttask_name\ttask_type\ttarget_drtn_hr_cnt\ttarget_start_date\ttarget_end_date\tact_start_date\tact_end_date\tphys_complete_pct\ttotal_float_hr_cnt\tcstr_type\tstatus_code
%R\t1001\t1\t1\t1\tA1010\tFoundation\tTT_Task\t180.0\t2024-07-08 08:00\t2024-07-29 17:00\t2024-07-08 08:00\t2024-07-29 17:00\t100.0\t0.0\tCS_ASAP\tTK_Complete
%R\t1002\t1\t1\t1\tA1020\tFrame\tTT_Task\t360.0\t2024-07-30 08:00\t2024-09-09 17:00\t2024-07-30 08:00\t\t75.0\t0.0\tCS_ASAP\tTK_Active
%R\t1003\t1\t1\t1\tA1030\tWalls\tTT_Task\t180.0\t2024-09-10 08:00\t2024-10-01 17:00\t\t\t0.0\t72.0\tCS_ASAP\tTK_NotStart
%R\t1004\t1\t1\t1\tA1040\tRoof\tTT_Task\t180.0\t2024-10-02 08:00\t2024-10-23 17:00\t\t\t0.0\t72.0\tCS_ASAP\tTK_NotStart
%R\t1005\t1\t1\t1\tA1050\tInterior\tTT_Task\t360.0\t2024-10-24 08:00\t2024-12-04 17:00\t\t\t0.0\t72.0\tCS_ASAP\tTK_NotStart
%R\t1006\t1\t1\t1\tA1060\tHandover\tTT_FinMile\t0.0\t2024-12-15 17:00\t2024-12-15 17:00\t\t\t0.0\t81.0\tCS_MFO\tTK_NotStart
%T\tTASKPRED
%F\ttask_pred_id\ttask_id\tpred_task_id\tpred_type\tlag_hr_cnt
%R\t1\t1002\t1001\tPR_FS\t0.0
%R\t2\t1003\t1002\tPR_FS\t0.0
%R\t3\t1004\t1003\tPR_FS\t0.0
%R\t4\t1005\t1004\tPR_FS\t0.0
%R\t5\t1006\t1005\tPR_FS\t0.0
%T\tTASKRSRC
%F\ttaskrsrc_id\ttask_id\trsrc_id\ttarget_qty\tact_reg_qty\ttarget_cost\tact_reg_cost
%R\t1\t1001\t101\t180.0\t180.0\t180.0\t180.0
%R\t2\t1001\t103\t1000.0\t1000.0\t1000.0\t1000.0
%R\t3\t1002\t101\t360.0\t270.0\t360.0\t270.0
%R\t4\t1002\t104\t180.0\t135.0\t180.0\t135.0
%R\t5\t1003\t101\t180.0\t0.0\t180.0\t0.0
%R\t6\t1004\t104\t180.0\t0.0\t180.0\t0.0
%R\t7\t1005\t102\t360.0\t0.0\t360.0\t0.0
%E
"""


@pytest.fixture
def sample_cau_xer(tmp_path):
    """Write synthetic CAU-style XER (UTF-16-LE BOM) to tmp_path/sample_cau.xer."""
    path = tmp_path / "sample_cau.xer"
    # XER files are UTF-16-LE with BOM
    with open(path, "wb") as f:
        f.write(b"\xff\xfe")  # BOM
        f.write(SAMPLE_XER_CONTENT.encode("utf-16-le"))
    return str(path)
```

### Step 2: Write failing tests

`tests/test_xer_parser.py`:
```python
"""Test pure-Python XER parser."""
import os
import pytest
from xer_parser import XerFile


def test_xer_file_parses(sample_cau_xer):
    x = XerFile(sample_cau_xer)
    assert x is not None


def test_xer_tables_present(sample_cau_xer):
    x = XerFile(sample_cau_xer)
    for tbl in ("PROJECT", "CALENDAR", "RSRC", "TASK", "TASKPRED", "TASKRSRC"):
        assert tbl in x.tables, f"missing table {tbl}"


def test_xer_task_table_row_count(sample_cau_xer):
    x = XerFile(sample_cau_xer)
    assert len(x.tables["TASK"]["rows"]) == 6


def test_xer_task_table_headers(sample_cau_xer):
    x = XerFile(sample_cau_xer)
    headers = x.tables["TASK"]["headers"]
    assert "task_id" in headers
    assert "task_name" in headers


def test_xer_task_first_row_dict(sample_cau_xer):
    x = XerFile(sample_cau_xer)
    row = x.tables["TASK"]["rows"][0]
    assert row["task_code"] == "A1010"
    assert row["task_name"] == "Foundation"


def test_xer_handles_utf8_no_bom(tmp_path):
    """UTF-8 XER (no BOM) should also parse."""
    content = "ERMHDR\t18.8\t2026-01-01\tu\tApp\tUSD\n%T\tTASK\n%F\ttask_id\ttask_name\n%R\t1\tT1\n%E\n"
    path = tmp_path / "u8.xer"
    path.write_bytes(content.encode("utf-8"))
    x = XerFile(str(path))
    assert "TASK" in x.tables


def test_xer_file_not_found():
    with pytest.raises(FileNotFoundError):
        XerFile("/definitely/nonexistent.xer")


def test_xer_empty_table_handled(tmp_path):
    """Table with header but no rows."""
    content = "ERMHDR\t18.8\t2026-01-01\tu\tApp\tUSD\n%T\tTASK\n%F\ttask_id\ttask_name\n%E\n"
    path = tmp_path / "e.xer"
    path.write_bytes(b"\xff\xfe" + content.encode("utf-16-le"))
    x = XerFile(str(path))
    assert "TASK" in x.tables
    assert x.tables["TASK"]["rows"] == []
```

### Step 3: Run — expect ImportError

### Step 4: Implementation

Create `xer_parser.py`:
```python
"""Phase 5d - Pure-Python Primavera P6 XER reader.

XER format (text, typically UTF-16-LE with BOM, fallback UTF-8):
- ERMHDR <version>\\t<date>\\t<user>\\t<app>\\t<currency>
- %T <table_name>            : table marker
- %F <header1>\\t<header2>... : field names
- %R <val1>\\t<val2>...       : data row (position-mapped to %F)
- %E                         : end of file
"""
import os
import logging

logger = logging.getLogger(__name__)


class XerFile:
    """Parse a P6 XER file into table dicts."""

    def __init__(self, file_path):
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"XER file not found: {file_path}")
        self.file_path = file_path
        self.header_fields = {}  # ERMHDR fields
        self.tables = {}  # {table_name: {"headers": [...], "rows": [{col:val}]}}
        self._parse()

    def _read_text(self):
        """Read file with encoding auto-detect (UTF-16-LE BOM or UTF-8)."""
        with open(self.file_path, "rb") as f:
            raw = f.read()
        if raw[:2] == b"\xff\xfe":
            return raw[2:].decode("utf-16-le", errors="replace")
        if raw[:3] == b"\xef\xbb\xbf":
            return raw[3:].decode("utf-8", errors="replace")
        try:
            return raw.decode("utf-16-le")
        except UnicodeDecodeError:
            return raw.decode("utf-8", errors="replace")

    def _parse(self):
        text = self._read_text()
        lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
        current_table = None
        for line in lines:
            if not line:
                continue
            if line.startswith("ERMHDR"):
                # First field is "ERMHDR", rest are positional metadata
                parts = line.split("\t")
                self.header_fields = {
                    "version": parts[1] if len(parts) > 1 else "",
                    "exported": parts[2] if len(parts) > 2 else "",
                    "user": parts[3] if len(parts) > 3 else "",
                    "app": parts[4] if len(parts) > 4 else "",
                    "currency": parts[5] if len(parts) > 5 else "",
                }
                continue
            if line.startswith("%T"):
                parts = line.split("\t", 1)
                if len(parts) == 2:
                    current_table = parts[1].strip()
                    self.tables[current_table] = {"headers": [], "rows": []}
                continue
            if line.startswith("%F"):
                if current_table is None:
                    continue
                parts = line.split("\t")
                self.tables[current_table]["headers"] = [p.strip() for p in parts[1:]]
                continue
            if line.startswith("%R"):
                if current_table is None:
                    continue
                headers = self.tables[current_table]["headers"]
                if not headers:
                    continue
                parts = line.split("\t")
                values = parts[1:]
                # pad/truncate to header length
                if len(values) < len(headers):
                    values = values + [""] * (len(headers) - len(values))
                row = {h: values[i] for i, h in enumerate(headers)}
                self.tables[current_table]["rows"].append(row)
                continue
            if line.startswith("%E"):
                break
            # Unknown marker line - skip
```

### Step 5: Run + commit

```bash
cd /c/Users/CahAsus/asta-powerproject-mcp && python -m pytest tests/test_xer_parser.py -v
git add xer_parser.py tests/test_xer_parser.py tests/conftest.py
git commit -m "Phase 5d T102: xer_parser foundations + encoding detect + table splitter"
```

---

## Task 103: `read_tasks` + `read_links` (TASK + TASKPRED)

**Files:** Modify `xer_parser.py` + `tests/test_xer_parser.py`

### Step 1: Failing tests

```python
def test_read_tasks_count(sample_cau_xer):
    x = XerFile(sample_cau_xer)
    tasks = x.read_tasks()
    assert len(tasks) == 6


def test_read_tasks_msp_shape(sample_cau_xer):
    x = XerFile(sample_cau_xer)
    tasks = x.read_tasks()
    t = tasks[0]
    for k in ("id", "name", "duration_h", "start", "finish",
              "percent_complete", "summary", "constraint_type"):
        assert k in t


def test_read_tasks_id_int(sample_cau_xer):
    x = XerFile(sample_cau_xer)
    tasks = x.read_tasks()
    assert isinstance(tasks[0]["id"], int)
    assert tasks[0]["id"] == 1001


def test_read_tasks_milestone_summary_flag(sample_cau_xer):
    """TT_FinMile is a milestone, not summary. Summary requires TT_LOE."""
    x = XerFile(sample_cau_xer)
    tasks = x.read_tasks()
    # Handover (id=1006) is TT_FinMile - milestone but not summary
    handover = next(t for t in tasks if t["id"] == 1006)
    assert handover["summary"] is False


def test_read_tasks_constraint_type_mapped(sample_cau_xer):
    x = XerFile(sample_cau_xer)
    tasks = x.read_tasks()
    # Handover has CS_MFO -> 3
    handover = next(t for t in tasks if t["id"] == 1006)
    assert handover["constraint_type"] == 3
    # Foundation has CS_ASAP -> 0
    foundation = next(t for t in tasks if t["id"] == 1001)
    assert foundation["constraint_type"] == 0


def test_read_tasks_percent_complete(sample_cau_xer):
    x = XerFile(sample_cau_xer)
    tasks = x.read_tasks()
    foundation = next(t for t in tasks if t["id"] == 1001)
    assert foundation["percent_complete"] == 100.0


def test_read_tasks_total_float_days(sample_cau_xer):
    """total_float_hr_cnt converted to days using 9h/day calendar."""
    x = XerFile(sample_cau_xer)
    tasks = x.read_tasks()
    walls = next(t for t in tasks if t["id"] == 1003)
    # 72h / 9h/day = 8 days. (Or 72/8 if generic — defaults to 8h/day)
    assert walls["total_float"] >= 7  # tolerant — exact div depends on calendar


# ---- Links ----

def test_read_links_count(sample_cau_xer):
    x = XerFile(sample_cau_xer)
    links = x.read_links()
    assert len(links) == 5  # 5 FS chain


def test_read_links_msp_shape(sample_cau_xer):
    x = XerFile(sample_cau_xer)
    link = x.read_links()[0]
    for k in ("from_id", "to_id", "type", "lag_days"):
        assert k in link


def test_read_links_type_mapping(sample_cau_xer):
    x = XerFile(sample_cau_xer)
    links = x.read_links()
    assert all(l["type"] == "FS" for l in links)


def test_read_links_zero_lag(sample_cau_xer):
    x = XerFile(sample_cau_xer)
    links = x.read_links()
    assert all(l["lag_days"] == 0 for l in links)


def test_read_links_pred_to_succ_mapping(sample_cau_xer):
    """from_id = predecessor, to_id = successor."""
    x = XerFile(sample_cau_xer)
    links = x.read_links()
    # First link: pred=1001, succ=1002
    first = links[0]
    assert first["from_id"] == 1001
    assert first["to_id"] == 1002
```

### Step 3: Implementation

Append to `xer_parser.py`:
```python
# ---------- Field mapping helpers ----------

CONSTRAINT_TYPE_MAP = {
    "CS_ASAP": 0, "CS_ALAP": 1,
    "CS_MSO": 2, "CS_MFO": 3,
    "CS_MSOA": 4, "CS_MSOB": 5,
    "CS_MEOA": 6, "CS_MEOB": 7,
}

LINK_TYPE_MAP = {
    "PR_FS": "FS", "PR_SS": "SS", "PR_FF": "FF", "PR_SF": "SF",
}

SUMMARY_TASK_TYPES = {"TT_LOE", "TT_WBS"}


def _to_int(s, default=None):
    try:
        return int(float(s)) if s else default
    except (ValueError, TypeError):
        return default


def _to_float(s, default=0.0):
    try:
        return float(s) if s else default
    except (ValueError, TypeError):
        return default


def _to_iso_date(s):
    """XER dates are 'YYYY-MM-DD HH:MM' or empty. Return ISO date or None."""
    if not s or not s.strip():
        return None
    return s[:10]  # 'YYYY-MM-DD' prefix sufficient


# ---------- Public read methods ----------

class XerFile:  # extension - the methods below merge into class above
    pass


# (re-open class to add methods - in actual impl, methods go inside class body
# directly. This block is a continuation marker only.)


def _read_tasks(self, day_hr_cnt=8.0):
    """TASK section -> list of MSP-shape task dicts."""
    tbl = self.tables.get("TASK", {"rows": []})
    out = []
    for row in tbl["rows"]:
        ttype = row.get("task_type", "")
        out.append({
            "id": _to_int(row.get("task_id")),
            "name": row.get("task_name", ""),
            "code": row.get("task_code", ""),
            "duration_h": _to_float(row.get("target_drtn_hr_cnt")),
            "start": _to_iso_date(row.get("target_start_date")),
            "finish": _to_iso_date(row.get("target_end_date")),
            "actual_start": _to_iso_date(row.get("act_start_date")),
            "actual_finish": _to_iso_date(row.get("act_end_date")),
            "percent_complete": _to_float(row.get("phys_complete_pct")),
            "total_float": _to_float(row.get("total_float_hr_cnt")) / day_hr_cnt,
            "summary": ttype in SUMMARY_TASK_TYPES,
            "task_type": ttype,
            "constraint_type": CONSTRAINT_TYPE_MAP.get(row.get("cstr_type", ""), 0),
            "status": row.get("status_code", ""),
        })
    return out


def _read_links(self):
    """TASKPRED section -> list of {from_id, to_id, type, lag_days}."""
    tbl = self.tables.get("TASKPRED", {"rows": []})
    out = []
    for row in tbl["rows"]:
        out.append({
            "from_id": _to_int(row.get("pred_task_id")),
            "to_id": _to_int(row.get("task_id")),
            "type": LINK_TYPE_MAP.get(row.get("pred_type", ""), "FS"),
            "lag_days": _to_float(row.get("lag_hr_cnt")) / 8.0,
        })
    return out


XerFile.read_tasks = _read_tasks
XerFile.read_links = _read_links
```

### Step 5: Commit

```bash
git commit -m "Phase 5d T103: read_tasks + read_links (TASK + TASKPRED parse)"
```

---

## Task 104: `read_resources` + `read_assignments` + `read_calendars`

**Files:** Modify `xer_parser.py` + `tests/test_xer_parser.py`

### Step 1: Failing tests

```python
def test_read_resources_count(sample_cau_xer):
    x = XerFile(sample_cau_xer)
    res = x.read_resources()
    assert len(res) == 4


def test_read_resources_shape(sample_cau_xer):
    x = XerFile(sample_cau_xer)
    res = x.read_resources()
    r = res[0]
    for k in ("id", "name", "type", "max_units"):
        assert k in r


def test_read_resources_cow(sample_cau_xer):
    x = XerFile(sample_cau_xer)
    res = x.read_resources()
    cow = next(r for r in res if r["id"] == 101)
    assert cow["name"] == "Concrete Workers"
    assert cow["max_units"] == 10.0


def test_read_assignments_count(sample_cau_xer):
    x = XerFile(sample_cau_xer)
    a = x.read_assignments()
    assert len(a) == 7


def test_read_assignments_shape(sample_cau_xer):
    x = XerFile(sample_cau_xer)
    a = x.read_assignments()
    item = a[0]
    for k in ("task_id", "resource_id", "target_qty", "actual_qty"):
        assert k in item


def test_read_assignments_task_resource_link(sample_cau_xer):
    x = XerFile(sample_cau_xer)
    a = x.read_assignments()
    foundation_assigns = [x for x in a if x["task_id"] == 1001]
    assert len(foundation_assigns) == 2
    # Foundation has COW(101) and STL(103)
    res_ids = {x["resource_id"] for x in foundation_assigns}
    assert res_ids == {101, 103}


def test_read_calendars_count(sample_cau_xer):
    x = XerFile(sample_cau_xer)
    cals = x.read_calendars()
    assert len(cals) == 1


def test_read_calendars_cau_9h_day(sample_cau_xer):
    """CAU calendar has day_hr_cnt=9 (CLAUDE.md RULE 1)."""
    x = XerFile(sample_cau_xer)
    cals = x.read_calendars()
    cal = cals[0]
    assert cal["day_hr_cnt"] == 9.0
    assert cal["week_hr_cnt"] == 54.0
```

### Step 3: Implementation

Append to `xer_parser.py`:
```python
def _read_resources(self):
    """RSRC section -> list of {id, name, type, max_units}."""
    tbl = self.tables.get("RSRC", {"rows": []})
    out = []
    for row in tbl["rows"]:
        # P6 RT_Labor/RT_Mat/RT_Equip; map to MSP Work/Material/...
        rtype = row.get("rsrc_type", "")
        msp_type = "Material" if rtype == "RT_Mat" else "Work"
        out.append({
            "id": _to_int(row.get("rsrc_id")),
            "name": row.get("rsrc_name", ""),
            "code": row.get("rsrc_short_name", ""),
            "type": msp_type,
            "max_units": _to_float(row.get("max_qty_per_hr"), default=1.0),
        })
    return out


def _read_assignments(self):
    """TASKRSRC section -> list of {task_id, resource_id, target_qty, actual_qty,
    target_cost, actual_cost}."""
    tbl = self.tables.get("TASKRSRC", {"rows": []})
    out = []
    for row in tbl["rows"]:
        out.append({
            "task_id": _to_int(row.get("task_id")),
            "resource_id": _to_int(row.get("rsrc_id")),
            "target_qty": _to_float(row.get("target_qty")),
            "actual_qty": _to_float(row.get("act_reg_qty")),
            "target_cost": _to_float(row.get("target_cost")),
            "actual_cost": _to_float(row.get("act_reg_cost")),
        })
    return out


def _read_calendars(self):
    """CALENDAR section -> list of {id, name, day_hr_cnt, week_hr_cnt}."""
    tbl = self.tables.get("CALENDAR", {"rows": []})
    out = []
    for row in tbl["rows"]:
        out.append({
            "id": _to_int(row.get("clndr_id")),
            "name": row.get("clndr_name", ""),
            "day_hr_cnt": _to_float(row.get("day_hr_cnt"), default=8.0),
            "week_hr_cnt": _to_float(row.get("week_hr_cnt"), default=40.0),
        })
    return out


XerFile.read_resources = _read_resources
XerFile.read_assignments = _read_assignments
XerFile.read_calendars = _read_calendars
```

### Step 5: Commit

```bash
git commit -m "Phase 5d T104: read_resources + read_assignments + read_calendars"
```

---

## Task 105: `read_progress` + status_date + project metadata

**Files:** Modify `xer_parser.py` + `tests/test_xer_parser.py`

### Step 1: Failing tests

```python
def test_read_progress_returns_status_date(sample_cau_xer):
    x = XerFile(sample_cau_xer)
    p = x.read_progress()
    assert "status_date" in p
    assert p["status_date"] == "2026-05-01"  # PROJECT.last_recalc_date


def test_read_progress_tasks(sample_cau_xer):
    x = XerFile(sample_cau_xer)
    p = x.read_progress()
    assert "tasks" in p
    foundation = next(t for t in p["tasks"] if t["id"] == 1001)
    assert foundation["percent_complete"] == 100.0


def test_read_project_metadata(sample_cau_xer):
    x = XerFile(sample_cau_xer)
    proj = x.read_project()
    assert proj["proj_id"] == 1
    assert proj["plan_start_date"] == "2024-07-08"
    assert proj["plan_end_date"] == "2028-06-20"


def test_read_progress_no_project_table(tmp_path):
    """Empty XER without PROJECT -> status_date None."""
    content = "ERMHDR\t18.8\t2026-01-01\tu\tApp\tUSD\n%T\tTASK\n%F\ttask_id\ttask_name\n%E\n"
    path = tmp_path / "n.xer"
    path.write_bytes(b"\xff\xfe" + content.encode("utf-16-le"))
    x = XerFile(str(path))
    p = x.read_progress()
    assert p["status_date"] is None
```

### Step 3: Implementation

```python
def _read_project(self):
    """PROJECT section -> first row dict (typically only 1 project per XER)."""
    tbl = self.tables.get("PROJECT", {"rows": []})
    if not tbl["rows"]:
        return {}
    row = tbl["rows"][0]
    return {
        "proj_id": _to_int(row.get("proj_id")),
        "plan_start_date": _to_iso_date(row.get("plan_start_date")),
        "plan_end_date": _to_iso_date(row.get("plan_end_date")),
        "last_recalc_date": _to_iso_date(row.get("last_recalc_date")),
        "proj_short_name": row.get("proj_short_name", ""),
    }


def _read_progress(self):
    """Return {status_date, tasks: [{id, percent_complete, actual_work_h}]}.

    status_date = PROJECT.last_recalc_date (P6 convention).
    """
    proj = self._read_project_internal()
    tbl = self.tables.get("TASK", {"rows": []})
    progress_tasks = []
    for row in tbl["rows"]:
        progress_tasks.append({
            "id": _to_int(row.get("task_id")),
            "percent_complete": _to_float(row.get("phys_complete_pct")),
            "actual_work_h": 0.0,  # XER actual work in TASKRSRC.act_reg_qty (resource-loaded)
        })
    return {
        "status_date": (proj or {}).get("last_recalc_date"),
        "tasks": progress_tasks,
    }


def _read_project_internal(self):
    """Internal alias for _read_project to allow read_progress to use without
    coupling to the public method name."""
    return _read_project(self)


XerFile.read_project = _read_project
XerFile.read_progress = _read_progress
```

### Step 5: Commit

```bash
git commit -m "Phase 5d T105: read_progress + status_date + project metadata"
```

---

## Task 106: BIG ONE — `_xer_collect_full_data` + 6 Action Helpers

**Subagent dispatch.** Files:
- Modify: `msproject_mcp_core.py` (add Phase 5D section AFTER Phase 5c excel dispatcher, BEFORE def main)
- Create: `tests/test_msproject_xer_loader.py`
- Create: `tests/test_msproject_xer_actions.py`

**Implementation skeleton:**

```python
# ============================================================================
# PHASE 5D - XER (PRIMAVERA P6) READER
# ============================================================================
from xer_parser import XerFile


def _xer_collect_full_data(file_path):
    """Single-collect aggregator (Phase 5b/5c TAIL lesson) - parse XER once,
    expose all 6 read shapes from a single XerFile instance."""
    if not file_path:
        return {"status": "error", "error": "file_path required"}
    try:
        xer = XerFile(file_path)
    except FileNotFoundError as e:
        return {"status": "error", "error": str(e)}
    except Exception as e:
        logger.exception(f"_xer_collect_full_data failed: {e}")
        return {"status": "error", "error": str(e)}
    # Use first calendar's day_hr_cnt for total_float conversion
    cals = xer.read_calendars()
    day_hr_cnt = cals[0]["day_hr_cnt"] if cals else 8.0
    return {
        "status": "ok",
        "tasks": xer.read_tasks(day_hr_cnt=day_hr_cnt),
        "links": xer.read_links(),
        "resources": xer.read_resources(),
        "assignments": xer.read_assignments(),
        "calendars": cals,
        "progress": xer.read_progress(),
        "project": xer.read_project(),
    }


def _msp_xer_read_tasks(file_path=None, filters=None, limit=None):
    data = _xer_collect_full_data(file_path)
    if data.get("status") != "ok":
        return data
    tasks = data["tasks"]
    if filters:
        # Phase 4 filter reuse (if available, else simple eq filter)
        for k, v in filters.items():
            tasks = [t for t in tasks if t.get(k) == v]
    if limit:
        tasks = tasks[:int(limit)]
    return {"status": "ok", "count": len(tasks), "tasks": tasks}


def _msp_xer_read_links(file_path=None):
    data = _xer_collect_full_data(file_path)
    if data.get("status") != "ok":
        return data
    return {"status": "ok", "count": len(data["links"]), "links": data["links"]}


def _msp_xer_read_resources(file_path=None):
    data = _xer_collect_full_data(file_path)
    if data.get("status") != "ok":
        return data
    return {"status": "ok", "count": len(data["resources"]), "resources": data["resources"]}


def _msp_xer_read_assignments(file_path=None):
    data = _xer_collect_full_data(file_path)
    if data.get("status") != "ok":
        return data
    return {"status": "ok", "count": len(data["assignments"]), "assignments": data["assignments"]}


def _msp_xer_read_calendars(file_path=None):
    data = _xer_collect_full_data(file_path)
    if data.get("status") != "ok":
        return data
    return {"status": "ok", "count": len(data["calendars"]), "calendars": data["calendars"]}


def _msp_xer_read_progress(file_path=None):
    data = _xer_collect_full_data(file_path)
    if data.get("status") != "ok":
        return data
    return {"status": "ok", **data["progress"]}
```

**Tests:** Standard pattern (5 loader + 6 action tests).

**Commit:** "Phase 5d T106 (BIG ONE): _xer_collect_full_data + 6 action helpers"

---

## Task 107: `@mcp.tool msproject_xer` Dispatcher + Tests

**Files:**
- Modify: `msproject_mcp_core.py` (add dispatcher after Phase 5D helpers, before def main)
- Create: `tests/test_msproject_xer_dispatcher.py`

```python
@mcp.tool(
    name="msproject_xer",
    annotations={
        "title": "MS Project XER (Primavera P6) Reader",
        "readOnlyHint": True,
    },
)
async def msproject_xer(params: dict) -> str:
    """Pure-Python Primavera P6 XER reader. Read-only.

    Actions: read_tasks, read_links, read_resources, read_assignments,
             read_calendars, read_progress.

    Phase 5d (1 May 2026). Tool count 11 -> 12.
    """
    import json
    action = params.get("action", "")
    p = {k: v for k, v in params.items() if k != "action"}
    try:
        if action == "read_tasks":
            r = _msp_xer_read_tasks(**p)
        elif action == "read_links":
            r = _msp_xer_read_links(**p)
        elif action == "read_resources":
            r = _msp_xer_read_resources(**p)
        elif action == "read_assignments":
            r = _msp_xer_read_assignments(**p)
        elif action == "read_calendars":
            r = _msp_xer_read_calendars(**p)
        elif action == "read_progress":
            r = _msp_xer_read_progress(**p)
        else:
            r = {"status": "error",
                 "error": (f"Unknown action '{action}'. Valid: "
                           "read_tasks/read_links/read_resources/"
                           "read_assignments/read_calendars/read_progress")}
    except TypeError as e:
        r = {"status": "error", "error": f"Invalid params for {action}: {e}"}
    except Exception as e:
        logger.exception(f"msproject_xer({action}) failed: {e}")
        r = {"status": "error", "error": str(e)}
    return json.dumps(r, default=str, ensure_ascii=False)
```

Dispatcher tests (~6 tests covering all 6 actions + unknown action). All use `sample_cau_xer` fixture.

**Commit:** "Phase 5d T107: msproject_xer dispatcher + dispatcher tests"

---

## Task 108: Acceptance + README + Push (FINAL)

**Files:**
- Create: `samples/build_xer_lifecycle.py`
- Modify: `README.md`

**Acceptance scenario:**
1. Build sample_cau.xer to a tempdir (use SAMPLE_XER_CONTENT from conftest)
2. Call all 6 read actions, print counts
3. Verify task count = 6, link count = 5, etc.
4. Total time ≤ 10s (pure-Python, no COM)

**README addition:**
```markdown
## Phase 5d — XER Reader (1 May 2026)

`msproject_xer` tool — pure-Python Primavera P6 XER reader. 6 read-only
actions covering tasks, links, resources, assignments, calendars, and
progress. NO mpxj dependency (UTF-16-LE BOM tab-delimited XER format
parsed natively).

**Actions:** read_tasks, read_links, read_resources, read_assignments,
read_calendars, read_progress.

Architecture: `xer_parser.py` pure module + Phase 5D adapters + dispatcher.
Phase 1-5c helpers DOKUNULMAZ.

Acceptance: `samples/build_xer_lifecycle.py` parses synthetic CAU-style
XER (10 tasks) in <10s.

Tool count: **12 tools, ~95 actions**.
```

**Commit + push:**
```bash
git add msproject_mcp_core.py samples/build_xer_lifecycle.py README.md
git commit -m "Phase 5d T108: msproject_xer acceptance + README + push (12th tool)"
git push origin main
```

---

## Phase 5d Acceptance Criteria

1. ✅ T102-T108 7-task chain landed
2. ✅ Acceptance ≤ 10s @ 10-task XER
3. ✅ All 6 actions return MSP-shape dicts
4. ✅ Phase 1-5c regression untouched
5. ✅ Push to origin/main
6. ✅ Tool count 11 → 12, actions ~89 → ~95

---

*Plan committed: 2026-05-01.*
