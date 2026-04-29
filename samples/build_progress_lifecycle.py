"""Phase 3b acceptance: full progress lifecycle.

SAFETY: Uses isolated FileNew project, never touches user's active project.

Scenario:
  1. Create 18 villa tasks + 3 resources + assignments
  2. Save Baseline 0 ('Original') (Phase 3a integration)
  3. set_task_progress on first 5 tasks (percent_complete=50)
  4. set_assignment_progress on next 5 tasks (per-resource man-hours)
  5. time_phased_actual_write on 1 task (today's week, varying hours)
  6. time_phased_actual_read verification
  7. set_status_date to today
  8. set_progress_by_date for older tasks (plan=actual catch-up)
  9. bulk_progress_update with 8 remaining items (com_batch path)
  10. summary -> BAC, ACWP, project_pct
  11. clear_all_progress -> reset

Target: end-to-end <15s.

NOTE: N=18 chosen to avoid Phase 2b TAIL bug where _msp_task_bulk_add
mspdi_bulk path (N>=20) silently drops Duration field. Lifecycle still
demonstrates all 12 actions plus Phase 3a baseline integration.
"""
import os, sys, time
import datetime as dt
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pythoncom
import win32com.client
from msproject_mcp_core import (
    _msp_task_bulk_add, _msp_resource_add, _msp_resource_bulk_assign,
    _msp_baseline_save,
    _msp_progress_set_task, _msp_progress_get_task,
    _msp_progress_set_assignment, _msp_progress_get_assignments,
    _msp_progress_set_by_date, _msp_progress_set_status_date,
    _msp_progress_time_phased_write, _msp_progress_time_phased_read,
    _msp_progress_bulk_update, _msp_progress_summary,
    _msp_progress_clear_all,
)


N_TASKS = 18  # Stay below mspdi_bulk threshold (20) to avoid duration-drop


def _today_iso(offset_days=0):
    return (dt.date.today() + dt.timedelta(days=offset_days)).isoformat()


def main():
    pythoncom.CoInitialize()
    app = win32com.client.GetActiveObject("MSProject.Application")
    original_name = app.ActiveProject.Name if app.ActiveProject else None
    app.FileNew()
    test_proj = app.ActiveProject
    test_name = test_proj.Name
    print(f"[SAFE] isolated test: {test_name}, user's: {original_name}")

    try:
        t0 = time.time()
        # 1. Tasks + resources + assignments
        print(f"\n1. Building {N_TASKS} villa tasks + 3 resources + assignments...")
        tasks = _msp_task_bulk_add(items=[
            {"name": f"Villa T{i:03d}", "duration": "4d"} for i in range(N_TASKS)])
        task_ids = tasks["task_ids"]
        res_ids = []
        for name in ["COW", "STL", "MSN"]:
            res_ids.append(_msp_resource_add(name=name, type="Work",
                                             max_units=300)["resource_id"])
        items = [{"task_id": tid, "resource_id": rid}
                 for tid in task_ids for rid in res_ids]
        _msp_resource_bulk_assign(items=items)
        print(f"   OK {len(task_ids)} tasks, 3 resources, "
              f"{len(items)} assignments in {time.time()-t0:.2f}s")

        # 2. Save Baseline 0 (Phase 3a integration)
        print("2. Saving Baseline 0 'Original'...")
        b0 = _msp_baseline_save(baseline_number=0, name="Original")
        assert b0["status"] == "ok"

        # 3. set_task_progress on first 5
        print("3. set_task_progress on first 5 tasks (50%)...")
        for tid in task_ids[:5]:
            _msp_progress_set_task(task_id=tid, percent_complete=50)

        # 4. set_assignment_progress on next 5 (per-resource man-hours)
        print("4. set_assignment_progress on next 5 tasks (COW=24h, STL=18h, MSN=10h)...")
        for tid in task_ids[5:10]:
            _msp_progress_set_assignment(task_id=tid, resource_id=res_ids[0],
                                         actual_work_h=24)
            _msp_progress_set_assignment(task_id=tid, resource_id=res_ids[1],
                                         actual_work_h=18)
            _msp_progress_set_assignment(task_id=tid, resource_id=res_ids[2],
                                         actual_work_h=10)

        # 5. time_phased_actual_write (1 task x current week)
        print("5. time_phased_actual_write on T010 (5 days varying hours)...")
        periods = [
            {"start": _today_iso(0), "end": _today_iso(1), "actual_work_h": 6},
            {"start": _today_iso(1), "end": _today_iso(2), "actual_work_h": 8},
            {"start": _today_iso(2), "end": _today_iso(3), "actual_work_h": 8},
            {"start": _today_iso(3), "end": _today_iso(4), "actual_work_h": 4},
            {"start": _today_iso(4), "end": _today_iso(5), "actual_work_h": 7},
        ]
        tpw = _msp_progress_time_phased_write(
            task_id=task_ids[10], resource_id=res_ids[0],
            periods=periods, unit="day")
        print(f"   written_count={tpw['written_count']}, status={tpw['status']}")

        # 6. time_phased_actual_read
        print("6. time_phased_actual_read verification...")
        tpr = _msp_progress_time_phased_read(
            task_id=task_ids[10], resource_id=res_ids[0],
            start_date=_today_iso(0), end_date=_today_iso(7), unit="day")
        for p in tpr["periods"][:5]:
            print(f"   {p['period_start'][:10] if p['period_start'] else '?'}: "
                  f"{p['actual_work_h']}h")

        # 7. set_status_date
        today_iso = _today_iso(0)
        print(f"7. set_status_date to today ({today_iso})...")
        sd = _msp_progress_set_status_date(status_date=today_iso)
        assert sd["status"] == "ok"

        # 8. set_progress_by_date (project beginning + 7d catch-up)
        catch_up = _today_iso(7)
        print(f"8. set_progress_by_date {catch_up} (plan=actual catch-up)...")
        sbd = _msp_progress_set_by_date(progress_date=catch_up)
        print(f"   {sbd['status']}, scope=all, mode={sbd.get('mode')}")

        # 9. bulk_progress_update on 8 remaining items (com_batch path)
        remaining = task_ids[10:]
        print(f"9. bulk_progress_update on {len(remaining)} items (com_batch path)...")
        bulk_items = [{"task_id": tid, "percent_complete": 30}
                       for tid in remaining]
        bu = _msp_progress_bulk_update(items=bulk_items)
        print(f"   {bu['status']}, path={bu['path']}, count={bu['count']}")

        # 10. summary
        print("10. summary (EVM-ready)...")
        summ = _msp_progress_summary()
        p = summ["project"]
        print(f"   BAC={p['bac_h']}h, ACWP={p['acwp_h']}h")
        print(f"   project_pct={p['project_percent_complete']}%, "
              f"completed={p['completed_count']}/{p['task_count']}")

        # 11. clear_all_progress
        print("11. clear_all_progress (reset)...")
        cl = _msp_progress_clear_all()
        print(f"   cleared_count={cl['cleared_count']}")

        elapsed = time.time() - t0
        print(f"\nOK ACCEPTANCE: {elapsed:.2f}s total (target <15s)")
        assert elapsed < 15.0, f"Too slow: {elapsed}s"

    finally:
        try:
            for i in range(1, app.Projects.Count + 1):
                if app.Projects(i).Name == test_name:
                    app.WindowActivate(app.Projects(i).Windows(1).Caption)
                    app.FileClose(0)
                    break
            if original_name:
                for i in range(1, app.Projects.Count + 1):
                    if app.Projects(i).Name == original_name:
                        app.WindowActivate(app.Projects(i).Windows(1).Caption)
                        break
        except Exception as e:
            print(f"[WARN] cleanup error: {e}")


if __name__ == "__main__":
    main()
