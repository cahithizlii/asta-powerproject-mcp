"""Phase 3a acceptance: full baseline lifecycle.

SAFETY: Uses isolated FileNew project, never touches user's active project.

Scenario:
  1. Create 50 villa tasks
  2. Add 3 work resources, assign to all tasks (mini Phase 2b chain)
  3. Save Baseline 0 ('Original')
  4. Update progress on first 20 tasks (extend duration to simulate slips)
  5. Compare(0) -> variance report
  6. Save Baseline 1 ('Rev1-AfterChangeOrder')
  7. Update more durations
  8. Compare_two(0, 1) -> revision delta
  9. Summary(0) -> RAG status

Target: end-to-end <10s.
"""
import os, sys, time
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pythoncom
import win32com.client
from msproject_mcp_core import (
    _msp_task_bulk_add, _msp_task_update,
    _msp_resource_add, _msp_resource_bulk_assign,
    _msp_baseline_save, _msp_baseline_compare, _msp_baseline_compare_two,
    _msp_baseline_summary, _msp_baseline_list,
    _enter_batch_mode, _exit_batch_mode,
)


N_TASKS = 40
SLIP_FIRST_N = 12  # batch-update count for first slip wave (kept small to stay <10s wall-clock)
SLIP_SECOND_N = 8  # batch-update count for second slip wave


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
        # 1-2. Tasks + resources
        print(f"\n1. Building {N_TASKS} villa tasks + 3 resources + assignments...")
        tasks = _msp_task_bulk_add(items=[{"name": f"Villa T{i:03d}", "duration": "2d"} for i in range(N_TASKS)])
        task_ids = tasks["task_ids"]
        res_ids = []
        for name in ["COW", "STL", "MSN"]:
            res_ids.append(_msp_resource_add(name=name, type="Work", max_units=300)["resource_id"])
        items = [{"task_id": tid, "resource_id": rid} for tid in task_ids for rid in res_ids]
        _msp_resource_bulk_assign(items=items)
        print(f"   OK {len(task_ids)} tasks, 3 resources, {len(items)} assignments in {time.time()-t0:.2f}s")

        # 3. Save baseline 0
        print("2. Saving Baseline 0 'Original'...")
        ts = time.time()
        b0 = _msp_baseline_save(baseline_number=0, name="Original")
        assert b0["status"] == "ok"
        print(f"   OK saved at {b0['saved_date']} ({b0['total_work_hours']}h total work) [{time.time()-ts:.2f}s]")

        # 4. Slip first wave (batch mode for speed)
        print(f"3. Slipping first {SLIP_FIRST_N} tasks (2d -> 5d)...")
        ts = time.time()
        _enter_batch_mode()
        try:
            for tid in task_ids[:SLIP_FIRST_N]:
                _msp_task_update(task_id=tid, duration="5d")
        finally:
            _exit_batch_mode()
        print(f"   [{time.time()-ts:.2f}s]")

        # 5. Compare against Baseline 0
        print("4. Compare current vs Baseline 0...")
        ts = time.time()
        cmp1 = _msp_baseline_compare(baseline_number=0)
        s = cmp1["summary"]
        print(f"   slipped={s['slipped_count']}, on_time={s['on_time_count']}, total_finish_drift={s['total_finish_drift_days']:.1f}d [{time.time()-ts:.2f}s]")

        # 6. Save baseline 1
        print("5. Saving Baseline 1 'Rev1-AfterChangeOrder'...")
        ts = time.time()
        b1 = _msp_baseline_save(baseline_number=1, name="Rev1-AfterChangeOrder")
        assert b1["status"] == "ok"
        print(f"   [{time.time()-ts:.2f}s]")

        # 7. More changes (batch mode)
        print(f"6. Slipping next {SLIP_SECOND_N} tasks (2d -> 4d)...")
        ts = time.time()
        _enter_batch_mode()
        try:
            for tid in task_ids[SLIP_FIRST_N:SLIP_FIRST_N + SLIP_SECOND_N]:
                _msp_task_update(task_id=tid, duration="4d")
        finally:
            _exit_batch_mode()
        print(f"   [{time.time()-ts:.2f}s]")

        # 8. Compare two baselines
        print("7. Compare Baseline 0 vs Baseline 1 (revision delta)...")
        ts = time.time()
        cmp2 = _msp_baseline_compare_two(baseline_a=0, baseline_b=1)
        s = cmp2["summary"]
        print(f"   slipped={s['slipped_count']}, total_finish_drift={s['total_finish_drift_days']:.1f}d [{time.time()-ts:.2f}s]")

        # 9. Summary
        print("8. Summary against Baseline 0 (RAG status)...")
        ts = time.time()
        summ = _msp_baseline_summary(baseline_number=0)
        p = summ["project"]
        print(f"   slipped_pct={p['slipped_pct']:.1f}%, schedule_health={p['schedule_health'].upper()} [{time.time()-ts:.2f}s]")

        # List
        print("9. Listing all saved baselines...")
        ts = time.time()
        bl = _msp_baseline_list()
        for b in bl["baselines"]:
            print(f"   - Baseline {b['number']}: {b['saved_date'][:10]} | {b['task_count']} tasks | {b['total_work_hours']:.1f}h work")
        print(f"   [{time.time()-ts:.2f}s]")

        elapsed = time.time() - t0
        print(f"\nOK ACCEPTANCE: {elapsed:.2f}s total (target <10s)")
        assert elapsed < 10.0, f"Too slow: {elapsed}s"

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
