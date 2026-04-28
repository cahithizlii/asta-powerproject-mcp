# MS Project MCP — Phase 2b Resource Design

**Versiyon:** 1.0
**Tarih:** 28 Nisan 2026
**Yazar:** MCS Mühendis (Cahit Hızlı) — brainstorming with Claude (delegated all sub-decisions)
**Status:** Approved → writing-plans next
**Phase 2a HEAD:** `738167a` (TAIL cleanup complete, on origin/main)

---

## 1. Hedef

`msproject_mcp` server'ına `msproject_resource` tool'u (5. tool) ekle. Phase 2b sonunda kullanıcı tek mesajla:

> "14 CAU ekibini (COW, EXT, STL, CAR, MSN, DRW, INW, EWI, CWI, ACP, ELW, DMS, MTR, LBR) tanımla, 200 villa task'ına bulk ata."

talimatını verebilmeli ve MS Project UI'da resources + 14×200 = 2800 assignment görünür olmalı, hepsi <5 sn'de.

## 2. Karar Geçmişi

### Q1 — Phase 2b scope: A
Tam paket — 7 action, tüm 3 resource türü (Work/Material/Cost). Original design doc'taki yaklaşım korunur.

**Tüm sub-decision'ları kullanıcı delege etti** ("A ile başla, devam et ve bitir"). Aşağıdaki kararlar bu mandate altında alındı.

## 3. Tool Surface

**Tool:** `msproject_resource`

| # | Action | Parametreler | Çıktı |
|---|---|---|---|
| 1 | `add` | `name` (str), `type` (str: "Work"/"Material"/"Cost"), [`max_units` (Work, default 100), `standard_rate` (float, $/h or $/unit or $/use), `overtime_rate` (Work only), `material_label` (Material only, e.g. "kg")] | `{status, resource_id, resource_uid, name, type}` |
| 2 | `update` | `resource_id` (int), [`name`, `max_units`, `standard_rate`, `overtime_rate`, `material_label`] | `{status, resource_id, changes}` |
| 3 | `delete` | `resource_id` (int) | `{status, deleted_id, deleted_name}` |
| 4 | `list` | (yok) | `{status, count, resources: [{id, uid, name, type, max_units, standard_rate, assignment_count}]}` |
| 5 | `assign` | `task_id` (int), `resource_id` (int), [`units` (Work, default 100), `work_hours` (Work only)] | `{status, assignment_uid, task_id, resource_id, units}` |
| 6 | `unassign` | `task_id` (int), `resource_id` (int) | `{status, task_id, resource_id}` |
| 7 | `bulk_assign` | `items=[{task_id, resource_id, units}, ...]` | Hybrid: `{status, path: "com_direct"\|"com_batch"\|"mspdi_bulk", count, assignments: [...]}` |

## 4. Resource Type Behavior

**Work** (default — işçi, ekip):
- `Type=0` (PjResourceType.pjResourceTypeWork)
- `MaxUnits` (% capacity, e.g. 100 = 1 person, 500 = 5-person crew)
- `StandardRate` ($/hour), `OvertimeRate` ($/hour)
- `MaterialLabel` not applicable

**Material** (malzeme):
- `Type=1` (pjResourceTypeMaterial)
- `MaterialLabel` (e.g. "kg", "m³", "ton")
- `StandardRate` ($/unit)
- `MaxUnits`, `OvertimeRate` not applicable

**Cost** (sabit gider):
- `Type=2` (pjResourceTypeCost)
- `StandardRate` ($/use, set per-assignment)
- No units, no rate fields meaningful at resource level

## 5. Hibrit Speed Strategy (bulk_assign)

Reuse Phase 1 `_route_operation()`:
- 1-5 assignments: COM direct (`task.AssignResource` per item)
- 6-19: COM batch (Calculation manual + ScreenUpdating off)
- 20+: MSPDI XML bulk import (extend `MsprojectBulkWriter` with `bulk_add_assignments`)

**Performance hedefi:** 14 resources × 200 tasks = 2800 assignments **<5 sn** (MSPDI bulk path).

## 6. Implementation Architecture

```
msproject_mcp_core.py
├── (existing Phase 1+2a)
└── (NEW Phase 2b — RESOURCE section)
    ├── RESOURCE_TYPES = {"Work": 0, "Material": 1, "Cost": 2}
    ├── _find_resource_by_name(proj, name) -> Resource | None
    ├── _serialize_resource(res) -> dict
    ├── _msp_resource_add(name, type, ...)
    ├── _msp_resource_update(resource_id, ...)
    ├── _msp_resource_delete(resource_id)
    ├── _msp_resource_list()
    ├── _msp_resource_assign(task_id, resource_id, units, work_hours)
    ├── _msp_resource_unassign(task_id, resource_id)
    ├── _msp_resource_bulk_assign(items)  # hybrid routing
    └── @mcp.tool msproject_resource dispatcher

msproject_bulk.py
└── (extend MsprojectBulkWriter)
    └── bulk_add_assignments(items)  # writes <Assignment> elements
```

**Phase 1 SAFETY pattern preserved:** all integration tests use `clean_test_project` fixture.

**`_format_com_error` (T29) used everywhere:** new error returns use `_format_com_error(e)` from day 1.

**`_find_resource_by_id` (T23) reused** + new `_find_resource_by_name` for lookups by name.

## 7. Test Strategy

**Reuse from Phase 2a:**
- `clean_test_project` fixture (SAFETY)
- TDD discipline (failing test → impl → green → commit per task)
- Subagent-driven execution (implementer + spec compliance review + code quality review per task)

**New test count target:** ~30-40 tests across 7 tasks (T32-T38).

**Performance assertions:**
- T37 bulk_assign 14×200=2800 assignments **<5 sn** (hero)
- T36 single assign **<200ms**
- Full Phase 2b regression suite **<30s**

## 8. Refactor opportunity (deferred)

T23's `_msp_calendar_assign_to_resource` test uses raw `proj.Resources.Add(name)`. Once Phase 2b lands, this can be refactored to use `_msp_resource_add(name, "Work")`. **Defer to Phase 2c or later** — keeps T32-T38 focused.

## 9. Acceptance Kriterleri (Phase 2b Tamam)

1. ✅ `msproject_resource` tool 7 action ile çalışıyor (T32-T38)
2. ✅ Acceptance script `samples/build_villa_resources.py`:
   - 14 CAU ekibini Work resource olarak ekle (named after CAU labels)
   - 200 villa task oluştur (reuse Phase 1 acceptance pattern)
   - bulk_assign 14 ekibi tüm task'lara → 2800 assignment
   - List resources → 14 görünür, her birinin assignment_count=200
   - Total time **<5 sn** (hero performance)
3. ✅ Phase 2b yeni testleri (~35) PASS
4. ✅ Phase 1 + Phase 2a mevcut 94 testi regression PASS — total **~125-130 PASS**
5. ✅ Phase 1 SAFETY: kullanıcının aktif projesi DOKUNULMAZ
6. ✅ Commit + push GitHub'a

## 10. Out of Scope (Phase 2b'de YOK)

- Resource leveling (Phase 3+)
- Engagements (resource workload requests)
- Resource calendars (foundation hazır via T23 `_msp_calendar_assign_to_resource`, ama bu tool'da yok)
- `bulk_add` for resources (sadece `bulk_assign` — resource oluşturma genelde küçük volumda, single COM yeterli)
- Resource sheet view manipulation
- Cost rate tables (variable rates by date) — Phase 3+
- T23 raw COM test refactor — Phase 2c

## 11. Sonraki Adım

`writing-plans` skill → bu design'ı baz alarak T32-T38 (7 task) bite-sized TDD implementation plan oluştur. Plan dosyası: `docs/plans/2026-04-28-phase2b-resource-impl.md`.

---

*Approved by user (delegated): 28 Nisan 2026*
*Next: writing-plans skill → Phase 2b implementation plan*
