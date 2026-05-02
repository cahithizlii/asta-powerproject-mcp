# Phase 5f — Phase 4 File MCP XER Symmetry (1 May 2026)

## Goal

Extend Phase 4 `_msp_file_read_*` helpers (7 readers) with `.xer` extension routing so `msproject_file.read_tasks(file_path="cau.xer")` works directly. Mirrors Phase 5e additive guard pattern; no new tool surface.

## Background

Phase 5e shipped XER native integration but only via Phase 5a/5b loaders (`_evm_load_task_data`, `_evm_load_baseline_data`, `_dcma_load_links`). Phase 4 file MCP `msproject_file` tool errors on `.xer` files because its 7 read helpers don't recognize the format.

**Discoverability:** users find `msproject_file` first (Phase 4 was the file MCP). They expect it to work on any project file format. Currently they must know to use `msproject_xer` for XER.

## Decision (1 May 2026)

- **Option A chosen** (vs B msproject_pdf / C session stop): smallest scope, fastest win, Phase 5e additive guard pattern reused.
- 2-task chain T112-T113. ~30 min.
- NO new tool. `msproject_file` becomes XER-aware.

## Architecture (additive routing — Phase 4 helpers DOKUNULMAZ preserved)

7 single-line guards added (additive only):

| Helper | Phase | Returns | XER source |
|---|---|---|---|
| `_msp_file_read_tasks(file_path, filters?, limit?)` | 4 | `{status, count, tasks}` | `XerFile.read_tasks()` |
| `_msp_file_read_links(file_path)` | 4 | `{status, count, links}` | `XerFile.read_links()` |
| `_msp_file_read_resources(file_path)` | 4 | `{status, count, resources}` | `XerFile.read_resources()` |
| `_msp_file_read_assignments(file_path, filters?)` | 4 | `{status, count, assignments}` | `XerFile.read_assignments()` |
| `_msp_file_read_calendars(file_path)` | 4 | `{status, count, calendars}` | `XerFile.read_calendars()` |
| `_msp_file_read_baselines(file_path, baseline_number?)` | 4 | `{status, count, tasks}` | `XerFile.read_tasks()` (CAU baseline = target) |
| `_msp_file_read_progress(file_path, filters?)` | 4 | `{status, status_date, tasks}` | `XerFile.read_progress()` |

All Phase 4 helpers gain pattern:
```python
if file_path and isinstance(file_path, str) and file_path.lower().endswith(".xer"):
    from xer_parser import XerFile
    try:
        xer = XerFile(file_path)
        # ... return appropriate shape
    except Exception as e:
        return {"status": "error", "error": str(e)}
# existing logic unchanged
```

## Sequencing — T112-T113

| Task | Content | Pattern |
|---|---|---|
| T112 | All 7 `_msp_file_read_*` `.xer` guards + per-helper tests | Manuel |
| T113 | Acceptance via `msproject_file` dispatcher (XER routing E2E) + README + push | Manuel finalize |

## Acceptance criteria

1. ✅ T112-T113 chain landed
2. ✅ `msproject_file.read_tasks(file_path="cau.xer")` returns 6 tasks
3. ✅ All 7 file MCP actions work on .xer
4. ✅ Phase 1-5e regression untouched (393 PASS)
5. ✅ Push to origin/main

## Out of scope

- XER write via msproject_file (P6 not user's authoring tool)
- Phase 4 read_baselines XER multi-baseline (XER has 1 implicit baseline)

---

*Design committed: 2026-05-01.*
