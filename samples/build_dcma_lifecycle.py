"""Phase 5b DCMA acceptance: 200-task CAU-style with intentional issues.

SAFETY: FileNew + FileClose 0. User's active project untouched.

Scenario (target <60s wall clock):
  1. Build 200 tasks + 14 CAU resources
  2. Inject DCMA failures intentionally:
     - First 12 tasks WITHOUT predecessor (RULE 1 fail)
     - 15 tasks duration > 44d (RULE 9 fail)
     - First 12 tasks unassigned (RULE 11)
  3. Save Baseline 0
  4. Phase 3b progress for ~30 tasks (RULE 14 BEI calc)
  5. set_status_date
  6. msproject_health assess_all -> display 14 rule results
  7. drill_down for first failed rule -> first 5 tasks
     (single drill keeps acceptance under 90s target — each drill_down
      re-collects via COM ~15s; full loop would take ~150s)
  8. summary -> RAG

Run: python samples/build_dcma_lifecycle.py
"""
import os
import sys
import time
import functools

# Force unbuffered stdout (flush after each print) for live progress in pipes
print = functools.partial(print, flush=True)

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import pythoncom
import win32com.client

from msproject_mcp_core import (
    _msp_task_bulk_add, _msp_resource_add, _msp_resource_bulk_assign,
    _msp_baseline_save,
    _msp_progress_bulk_update, _msp_progress_set_status_date,
    _msp_dcma_assess_all, _msp_dcma_summary, _msp_dcma_drill_down,
)

N_TASKS = 200
N_HIGH_DUR = 15  # tasks with duration > 44d
N_NO_PRED = 12   # tasks without predecessor (also unassigned)


def main():
    pythoncom.CoInitialize()
    app = win32com.client.GetActiveObject("MSProject.Application")
    app.FileNew()
    test_proj = app.ActiveProject
    test_name = test_proj.Name
    print(f"[SAFE] isolated test project: {test_name}", flush=True)

    try:
        t0 = time.time()

        # 1. Build base + intentional issues
        print(f"\n1. Building {N_TASKS} tasks ({N_TASKS - N_HIGH_DUR} normal + {N_HIGH_DUR} high-duration)...")
        items = [{"name": f"V{i:03d}", "duration": "5d"}
                 for i in range(N_TASKS - N_HIGH_DUR)]
        items += [{"name": f"H{i:02d}", "duration": "60d"} for i in range(N_HIGH_DUR)]
        tasks = _msp_task_bulk_add(items=items)
        task_ids = tasks["task_ids"]
        print(f"   tasks created in {time.time() - t0:.2f}s ({len(task_ids)} ids)")

        # 14 CAU resources
        cau = ["COW", "EXT", "STL", "CAR", "MSN", "DRW", "INW",
               "EWI", "CWI", "ACP", "ELW", "DMS", "MTR", "LBR"]
        res_ids = [_msp_resource_add(name=n, type="Work")["resource_id"]
                   for n in cau]
        print(f"   {len(res_ids)} CAU resources added")

        # Assign resources to most tasks (skip first 12 -> RULE 11 fail too)
        sample = [{"task_id": tid, "resource_id": res_ids[i % 14]}
                  for i, tid in enumerate(task_ids[N_NO_PRED:])]
        _msp_resource_bulk_assign(items=sample)
        print(f"   {len(sample)} resource assignments at {time.time() - t0:.2f}s")

        # 2. Save Baseline 0
        _msp_baseline_save(baseline_number=0)
        print(f"\n2. Baseline 0 saved at {time.time() - t0:.2f}s")

        # 3. Some progress for BEI
        progress_items = [{"task_id": tid, "percent_complete": 50.0}
                          for tid in task_ids[:30]]
        _msp_progress_bulk_update(items=progress_items)
        _msp_progress_set_status_date(status_date="2026-05-15")
        print(f"   progress + status_date set at {time.time() - t0:.2f}s")

        # 4. assess_all
        print(f"\n3. DCMA assess_all results:")
        r = _msp_dcma_assess_all()
        if r.get("status") != "ok":
            print(f"   ERROR: {r.get('error')}")
            return
        for rule in r["rules"]:
            ok_label = "OK  " if rule["status"] == "pass" else "FAIL"
            print(f"   [{ok_label}] Rule {rule['id']:2d}: {rule['name']:24s} "
                  f"actual={rule.get('actual')}{rule.get('actual_unit', '')} "
                  f"({rule['threshold']})")

        # 5. summary
        s = _msp_dcma_summary()
        rag = s["overall_rag"].upper()
        print(f"\n4. Summary: {s['pass_count']}/14 pass, RAG={rag}")
        print(f"   {s['executive_text']}")

        # 6. drill_down for first failed rule only (each drill re-collects
        # via COM ~15s; full loop over 7 fails would take ~150s and is
        # redundant for action validation - dispatcher tests cover all rules)
        first_fail = next((rule for rule in r["rules"]
                          if rule["status"] == "fail"), None)
        if first_fail:
            print(f"\n5. Drill-down (first failed rule):")
            d = _msp_dcma_drill_down(rule_id=first_fail["id"])
            if d.get("status") == "ok":
                print(f"   Rule {first_fail['id']} ({first_fail['name']}): "
                      f"{d['failed_count']} failed tasks")
                for ft in d["failed_tasks"][:5]:
                    print(f"      - Task {ft['id']}: {ft['name']}")

        elapsed = time.time() - t0
        print(f"\n[OK] ACCEPTANCE: {elapsed:.2f}s total (target <90s)")
        assert elapsed < 90.0, f"Too slow: {elapsed}s"

    finally:
        try:
            for i in range(1, app.Projects.Count + 1):
                if app.Projects(i).Name == test_name:
                    app.WindowActivate(app.Projects(i).Windows(1).Caption)
                    app.FileClose(0)
                    print(f"[SAFE] closed test project {test_name}")
                    break
        except Exception as e:
            print(f"[WARN] cleanup error: {e}")


if __name__ == "__main__":
    main()
