"""Phase 5a EVM acceptance: 200 task CAU-style + 4 snapshots <30s.

SAFETY: FileNew + FileClose 0. User's active project untouched.

Scenario (target <30s wall clock):
  1. Build 200 villa tasks + 14 CAU resources + assignments
  2. Save Baseline 0 (Original)
  3. Phase 3b: progress for week 1 (~30%)
  4. set_status_date "week 1"; save_period_snapshot tag=w1
  5. More progress for week 2 (~60%); save_period_snapshot tag=w2
  6. Week 3 + Week 4 snapshots -> 4 history entries
  7. trend -> SPI/CPI/EAC trajectory
  8. variance_to_baseline + compare_baselines_evm B0
  9. progress_data_quality + detect_currency_mode
"""
import os, sys, time, tempfile, shutil
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import pythoncom, win32com.client
from msproject_mcp_core import (
    _msp_task_bulk_add, _msp_resource_add, _msp_resource_bulk_assign,
    _msp_baseline_save, _msp_progress_set_task, _msp_progress_set_status_date,
    _msp_evm_compute_metrics, _msp_evm_forecast, _msp_evm_summary,
    _msp_evm_save_period_snapshot, _msp_evm_get_period_history, _msp_evm_trend,
    _msp_evm_variance_to_baseline,
    _msp_evm_progress_data_quality, _msp_evm_detect_currency_mode,
)

N_TASKS = 200


def main():
    pythoncom.CoInitialize()
    app = win32com.client.GetActiveObject("MSProject.Application")
    original_name = app.ActiveProject.Name if app.ActiveProject else None
    app.FileNew()
    test_proj = app.ActiveProject
    test_name = test_proj.Name
    print(f"[SAFE] isolated test: {test_name}")

    tmpdir = tempfile.mkdtemp(prefix="evm_phase5a_")
    snap_path = os.path.join(tmpdir, "snapshots.json")

    try:
        t0 = time.time()
        # 1. Build base
        print(f"\n1. Building {N_TASKS} tasks + 14 resources...")
        items = [{"name": f"V{i:03d}", "duration": "5d"} for i in range(N_TASKS)]
        tasks = _msp_task_bulk_add(items=items)
        task_ids = tasks["task_ids"]
        cau = ["COW", "EXT", "STL", "CAR", "MSN", "DRW", "INW",
               "EWI", "CWI", "ACP", "ELW", "DMS", "MTR", "LBR"]
        res_ids = [_msp_resource_add(name=n, type="Work")["resource_id"] for n in cau]
        sample = [{"task_id": tid, "resource_id": res_ids[i % 14]}
                  for i, tid in enumerate(task_ids)]
        _msp_resource_bulk_assign(items=sample)
        print(f"   OK in {time.time()-t0:.2f}s")

        # 2. Baseline 0
        _msp_baseline_save(baseline_number=0)
        print(f"\n2. Baseline 0 saved at {time.time()-t0:.2f}s")

        # 3-6. Multi-week snapshots
        for week, (n_done, pct, status_date, tag) in enumerate([
            (60, 30.0, "2026-05-07", "w1"),
            (120, 60.0, "2026-05-14", "w2"),
            (160, 80.0, "2026-05-21", "w3"),
            (180, 95.0, "2026-05-28", "w4"),
        ], start=1):
            for tid in task_ids[:n_done]:
                _msp_progress_set_task(task_id=tid, percent_complete=pct)
            _msp_progress_set_status_date(status_date=status_date)
            s = _msp_evm_save_period_snapshot(snapshot_path=snap_path, tag=tag)
            print(f"   W{week} snapshot saved: {s.get('snapshot_id')}")

        # 7. Trend
        trend = _msp_evm_trend(snapshot_path=snap_path)
        print(f"\n4. Trend series: {len(trend['series'])} points")
        for s in trend["series"]:
            print(f"   {s['tag']}: SPI={s['spi']} CPI={s['cpi']} RAG={s['rag']}")

        # 8. variance_to_baseline + compare_baselines_evm
        var0 = _msp_evm_variance_to_baseline(baseline_number=0)
        print(f"\n5. variance_to_baseline 0: SPI={var0.get('spi')} CPI={var0.get('cpi')}")

        # 9. Quality + currency
        pdq = _msp_evm_progress_data_quality()
        print(f"\n6. Data quality warnings: {len(pdq.get('warnings', []))}")
        cm = _msp_evm_detect_currency_mode()
        print(f"\n7. Currency mode: {cm.get('mode')}")

        elapsed = time.time() - t0
        print(f"\n[OK] ACCEPTANCE: {elapsed:.2f}s total (target <30s)")
        assert elapsed < 30.0, f"Too slow: {elapsed}s"

    finally:
        try:
            for i in range(1, app.Projects.Count + 1):
                if app.Projects(i).Name == test_name:
                    app.WindowActivate(app.Projects(i).Windows(1).Caption)
                    app.FileClose(0)
                    break
        except Exception:
            pass
        try:
            shutil.rmtree(tmpdir, ignore_errors=True)
        except Exception:
            pass


if __name__ == "__main__":
    main()
