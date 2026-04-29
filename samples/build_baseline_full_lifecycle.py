"""Phase 3a TAIL T51 — Extended acceptance: exercises ALL 9 msproject_baseline actions.

SAFETY: Uses isolated FileNew project, never touches user's active project.

Walks every public action exactly once:
  save -> get_task_baseline -> list -> compare -> (slip) -> compare -> save (B1) ->
  compare_two -> summary -> set_active -> list -> clear -> list -> clear_all -> list

Companion to build_baseline_lifecycle.py which is the perf benchmark; this one
prioritizes completeness over perf. Still aims for <10s wall-clock.
"""
import os, sys, time
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pythoncom
import win32com.client
from msproject_mcp_core import (
    _msp_task_bulk_add, _msp_task_update,
    _msp_baseline_save, _msp_baseline_clear, _msp_baseline_clear_all,
    _msp_baseline_list, _msp_baseline_get_task_baseline,
    _msp_baseline_compare, _msp_baseline_compare_two,
    _msp_baseline_summary, _msp_baseline_set_active,
)


N_TASKS = 15  # below MSPDI bulk-add threshold (20) so com_batch path preserves durations


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
        # 1. Build tasks
        print(f"\n1. Building {N_TASKS} tasks...")
        tasks = _msp_task_bulk_add(items=[{"name": f"FullT{i:02d}", "duration": "2d"} for i in range(N_TASKS)])
        task_ids = tasks["task_ids"]
        print(f"   OK {len(task_ids)} tasks in {time.time()-t0:.2f}s")

        # 2. SAVE Baseline 0 ('Original')
        print("2. action=save (Baseline 0 'Original')...")
        b0 = _msp_baseline_save(baseline_number=0, name="Original")
        assert b0["status"] == "ok"
        print(f"   OK B0 saved: {b0['saved_date']}, total {b0['total_work_hours']}h")

        # 3. GET_TASK_BASELINE — read first task's B0 values
        print("3. action=get_task_baseline (T0 in B0)...")
        gt = _msp_baseline_get_task_baseline(task_id=task_ids[0], baseline_number=0)
        assert gt["status"] == "ok"
        assert gt["baseline"]["start"] is not None
        assert gt["baseline"]["duration_h"] > 0
        print(f"   OK T0 B0: start={gt['baseline']['start']}, dur={gt['baseline']['duration_h']:.1f}h")

        # 4. LIST — should see B0 with name 'Original' (TAIL #3)
        print("4. action=list (after B0 save)...")
        bl1 = _msp_baseline_list()
        assert bl1["status"] == "ok"
        assert bl1["count_saved"] == 1
        assert bl1["baselines"][0]["name"] == "Original"
        print(f"   OK count_saved=1, name='Original'")

        # 5. COMPARE — no slip yet
        print("5. action=compare (current vs B0, no slip)...")
        cmp1 = _msp_baseline_compare(baseline_number=0)
        assert cmp1["status"] == "ok"
        assert cmp1["summary"]["slipped_count"] == 0
        print(f"   OK slipped=0, on_time={cmp1['summary']['on_time_count']}")

        # 6. Slip 5 tasks
        print(f"6. Slipping first 5 tasks (2d -> 5d)...")
        for tid in task_ids[:5]:
            _msp_task_update(task_id=tid, duration="5d")

        # 7. COMPARE again — should detect slips
        print("7. action=compare (after slip)...")
        cmp2 = _msp_baseline_compare(baseline_number=0)
        assert cmp2["summary"]["slipped_count"] == 5
        print(f"   OK slipped={cmp2['summary']['slipped_count']}, finish_drift={cmp2['summary']['total_finish_drift_days']:.1f}d")

        # 8. SAVE Baseline 1 ('Revised')
        print("8. action=save (Baseline 1 'Revised')...")
        b1 = _msp_baseline_save(baseline_number=1, name="Revised")
        assert b1["status"] == "ok"
        print(f"   OK B1 saved")

        # 9. COMPARE_TWO — B0 vs B1 delta
        print("9. action=compare_two (B0 vs B1)...")
        cmp_two = _msp_baseline_compare_two(baseline_a=0, baseline_b=1)
        assert cmp_two["status"] == "ok"
        # B1 captures the slipped state, so B1 - B0 shows positive finish drift
        assert cmp_two["summary"]["slipped_count"] == 5
        print(f"   OK delta slipped={cmp_two['summary']['slipped_count']}, drift={cmp_two['summary']['total_finish_drift_days']:.1f}d")

        # 10. SUMMARY — RAG against B0
        print("10. action=summary (RAG vs B0)...")
        summ = _msp_baseline_summary(baseline_number=0)
        assert summ["status"] == "ok"
        p = summ["project"]
        # 5 of 15 slipped = 33% -> red
        assert p["schedule_health"] == "red"
        print(f"   OK slipped_pct={p['slipped_pct']:.1f}%, schedule_health={p['schedule_health'].upper()}")

        # 11. SET_ACTIVE — set B1 active (happy path on MSP 16.0)
        print("11. action=set_active (B1)...")
        sa = _msp_baseline_set_active(baseline_number=1)
        # On MSP 16.0 should succeed; older versions may fall through to "not yet supported"
        if sa["status"] == "ok":
            print(f"   OK active baseline now {sa['active_baseline']} (via {sa.get('method', '?')})")
        else:
            print(f"   FALLBACK: {sa['error'][:80]}...")

        # 12. LIST — both B0 and B1
        print("12. action=list (both B0 and B1)...")
        bl2 = _msp_baseline_list()
        assert bl2["count_saved"] == 2
        names = sorted([b["name"] for b in bl2["baselines"]], key=lambda x: x or "")
        print(f"   OK count_saved=2, names={names}")

        # 13. CLEAR — single (B1)
        print("13. action=clear (B1 only)...")
        cl1 = _msp_baseline_clear(baseline_number=1)
        assert cl1["status"] == "ok"
        assert cl1["was_saved_date"] is not None
        print(f"   OK B1 cleared (was_saved_date={cl1['was_saved_date'][:10]})")

        # 14. LIST — only B0 remains
        print("14. action=list (after clear B1)...")
        bl3 = _msp_baseline_list()
        assert bl3["count_saved"] == 1
        assert bl3["baselines"][0]["number"] == 0
        # TAIL #3: B0's name should still be 'Original' (not evicted by clearing B1)
        assert bl3["baselines"][0]["name"] == "Original"
        print(f"   OK count_saved=1, B0 still 'Original'")

        # 15. CLEAR_ALL — wipe everything
        print("15. action=clear_all...")
        ca = _msp_baseline_clear_all()
        assert ca["status"] == "ok"
        assert ca["count"] == 1  # Only B0 was left
        print(f"   OK cleared {ca['count']} baseline(s): {ca['cleared']}")

        # 16. LIST — empty
        print("16. action=list (after clear_all)...")
        bl4 = _msp_baseline_list()
        assert bl4["count_saved"] == 0
        print(f"   OK count_saved=0")

        elapsed = time.time() - t0
        print(f"\nOK FULL ACCEPTANCE: all 9 actions exercised in {elapsed:.2f}s (target <10s)")
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
