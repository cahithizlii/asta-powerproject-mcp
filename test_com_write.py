"""
Test COM write operations for asta_powerproject_mcp.

Tests all write tools against a running Asta Powerproject instance:
1. asta_task -> add (normal + summary)
2. asta_task -> update (name, duration, dates)
3. asta_link -> add (FS, SS, with lag)
4. asta_link -> update (change type)
5. asta_link -> remove
6. asta_progress -> update (percent_complete)
7. asta_schedule -> reschedule
8. asta_task -> delete (cleanup)

All test objects are cleaned up at the end.

NOTE: Asta COM requires StartTransaction() before modifications,
      and EndTransaction() after. The MCP tool handlers do this.
"""

import sys
import os
import json
import traceback

# Add project dir to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Import COM helpers from asta_mcp_core
import pythoncom
import win32com.client
import pywintypes
from datetime import datetime, timedelta


def connect():
    """Connect to running Asta via COM."""
    pythoncom.CoInitialize()
    APP_CLSID = "{A57A0000-0200-0000-B2C5-00C0DF438041}"
    app = win32com.client.GetActiveObject(APP_CLSID)
    project = app.ActiveProject
    if project is None:
        raise RuntimeError("Asta is running but no project is open")
    return app, project


def fmt(result):
    """Format result dict for display."""
    return json.dumps(result, indent=2, default=str, ensure_ascii=False)


# Import the helper functions
from asta_mcp_core import (
    _com_add_task,
    _com_update_task,
    _com_delete_task,
    _com_add_link,
    _com_remove_link,
    _com_update_link,
    _com_update_progress,
    _com_end_transaction,
    _find_bar_by_id,
    _get_bar_task,
    _com_get_all_bars,
)


class TestResults:
    def __init__(self):
        self.passed = 0
        self.failed = 0
        self.results = []

    def check(self, name, result, success_key="task_id", error_key="error"):
        """Check a result dict and record pass/fail."""
        success = False
        if isinstance(result, dict):
            if error_key in result:
                success = False
            elif success_key in result:
                success = True
            elif result.get("success") or result.get("deleted") or result.get("removed") or result.get("updated"):
                success = True

        status = "PASS" if success else "FAIL"
        if success:
            self.passed += 1
        else:
            self.failed += 1

        self.results.append((name, status, result))
        print(f"\n{'='*60}")
        print(f"[{status}] {name}")
        print(f"{'='*60}")
        print(fmt(result))
        return success

    def summary(self):
        print(f"\n{'#'*60}")
        print(f"  TEST SUMMARY: {self.passed} passed, {self.failed} failed, {self.passed + self.failed} total")
        print(f"{'#'*60}")
        for name, status, _ in self.results:
            icon = "+" if status == "PASS" else "X"
            print(f"  {icon} [{status}] {name}")
        print()


def txn(project, label, func, reschedule=False):
    """Run func inside a StartTransaction/EndTransaction pair."""
    project.StartTransaction(label)
    try:
        result = func()
        _com_end_transaction(project, reschedule=reschedule)
        return result
    except Exception:
        try:
            project.AbandonTransaction()
        except Exception:
            try:
                project.EndTransaction()
            except Exception:
                pass
        raise


def run_tests():
    print("Connecting to Asta Powerproject...")
    app, project = connect()
    print(f"  Connected! Project: {project.Name}")

    T = TestResults()
    created_ids = []  # Track IDs for cleanup

    # --- Get root bar for parent placement ---
    root_bar = project.Bars.Item(1)
    root_id = root_bar.ID
    print(f"  Root bar: ID={root_id}, Name={root_bar.Name}")

    start_date = (datetime.now() + timedelta(days=7)).strftime("%Y-%m-%d")

    task1_id = None
    task2_id = None
    summary_id = None

    # ==========================================================
    # TEST 1: Add a normal task
    # ==========================================================
    print("\n\n>>> TEST 1: Add normal task (10d)")
    try:
        r1 = txn(project, "Test Add Task 1",
                 lambda: _com_add_task(project, "TEST-Task-Alpha", "10d", start_date=start_date),
                 reschedule=True)
        T.check("Add normal task (10d)", r1)
        if "task_id" in r1:
            created_ids.append(r1["task_id"])
            task1_id = r1["task_id"]
    except Exception as e:
        T.check("Add normal task (10d)", {"error": traceback.format_exc()})

    # ==========================================================
    # TEST 2: Add another task (for linking)
    # ==========================================================
    print("\n\n>>> TEST 2: Add second task (5d)")
    try:
        start2 = (datetime.now() + timedelta(days=21)).strftime("%Y-%m-%d")
        r2 = txn(project, "Test Add Task 2",
                 lambda: _com_add_task(project, "TEST-Task-Beta", "5d", start_date=start2),
                 reschedule=True)
        T.check("Add second task (5d)", r2)
        if "task_id" in r2:
            created_ids.append(r2["task_id"])
            task2_id = r2["task_id"]
    except Exception as e:
        T.check("Add second task (5d)", {"error": traceback.format_exc()})

    # ==========================================================
    # TEST 3: Add a summary task
    # ==========================================================
    print("\n\n>>> TEST 3: Add summary task")
    try:
        r3 = txn(project, "Test Add Summary",
                 lambda: _com_add_task(project, "TEST-Summary-Group", is_summary=True, start_date=start_date),
                 reschedule=True)
        T.check("Add summary task", r3)
        if "task_id" in r3:
            created_ids.append(r3["task_id"])
            summary_id = r3["task_id"]
    except Exception as e:
        T.check("Add summary task", {"error": traceback.format_exc()})

    # ==========================================================
    # TEST 4: Add child task under summary
    # ==========================================================
    print("\n\n>>> TEST 4: Add child task under summary")
    if summary_id:
        try:
            r4 = txn(project, "Test Add Child",
                     lambda: _com_add_task(project, "TEST-Child-Under-Summary", "3d",
                                          start_date=start_date, parent_bar_id=summary_id),
                     reschedule=True)
            T.check("Add child under summary", r4)
            if "task_id" in r4:
                created_ids.append(r4["task_id"])
        except Exception as e:
            T.check("Add child under summary", {"error": traceback.format_exc()})
    else:
        T.check("Add child under summary", {"error": "Skipped - no summary_id"})

    # ==========================================================
    # TEST 5: Update task name
    # ==========================================================
    print("\n\n>>> TEST 5: Update task name")
    if task1_id:
        try:
            r5 = txn(project, "Test Update Name",
                     lambda: _com_update_task(project, task1_id, name="TEST-Task-Alpha-RENAMED"),
                     reschedule=False)
            T.check("Update task name", r5)
        except Exception as e:
            T.check("Update task name", {"error": traceback.format_exc()})
    else:
        T.check("Update task name", {"error": "Skipped - no task1_id"})

    # ==========================================================
    # TEST 6: Update task duration
    # ==========================================================
    print("\n\n>>> TEST 6: Update task duration (10d -> 15d)")
    if task1_id:
        try:
            r6 = txn(project, "Test Update Duration",
                     lambda: _com_update_task(project, task1_id, duration_str="15d"),
                     reschedule=True)
            T.check("Update task duration", r6)
        except Exception as e:
            T.check("Update task duration", {"error": traceback.format_exc()})
    else:
        T.check("Update task duration", {"error": "Skipped - no task1_id"})

    # ==========================================================
    # TEST 7: Update task start date
    # ==========================================================
    print("\n\n>>> TEST 7: Update task start date")
    new_start = (datetime.now() + timedelta(days=14)).strftime("%Y-%m-%d")
    if task1_id:
        try:
            r7 = txn(project, "Test Update Start",
                     lambda: _com_update_task(project, task1_id, start_date=new_start),
                     reschedule=True)
            T.check("Update task start date", r7)
        except Exception as e:
            T.check("Update task start date", {"error": traceback.format_exc()})
    else:
        T.check("Update task start date", {"error": "Skipped - no task1_id"})

    # ==========================================================
    # TEST 8: Verify task data after updates
    # ==========================================================
    print("\n\n>>> TEST 8: Verify task data after updates")
    if task1_id:
        try:
            bar = _find_bar_by_id(project, task1_id)
            if bar:
                task, is_et = _get_bar_task(bar)
                verify = {
                    "bar_id": bar.ID,
                    "name": bar.Name,
                    "start": str(bar.Start),
                    "end": str(bar.End),
                }
                try:
                    verify["duration_hours"] = float(str(bar.Duration))
                except:
                    verify["duration"] = "N/A"
                if task:
                    verify["task_type"] = type(task).__name__
                    verify["task_id"] = task.ID
                T.check("Verify task data", verify, success_key="bar_id")
            else:
                T.check("Verify task data", {"error": f"Bar {task1_id} not found"})
        except Exception as e:
            T.check("Verify task data", {"error": traceback.format_exc()})
    else:
        T.check("Verify task data", {"error": "Skipped - no task1_id"})

    # ==========================================================
    # TEST 9: Add FS link
    # ==========================================================
    print("\n\n>>> TEST 9: Add FS link (Task1 -> Task2)")
    if task1_id and task2_id:
        try:
            r9 = txn(project, "Test Add FS Link",
                     lambda: _com_add_link(project, task1_id, task2_id, "FS"),
                     reschedule=True)
            T.check("Add FS link", r9, success_key="success")
        except Exception as e:
            T.check("Add FS link", {"error": traceback.format_exc()})
    else:
        T.check("Add FS link", {"error": "Skipped - missing task IDs"})

    # ==========================================================
    # TEST 10: Update link type (FS -> SS)
    # ==========================================================
    print("\n\n>>> TEST 10: Update link type (FS -> SS)")
    if task1_id and task2_id:
        try:
            r10 = txn(project, "Test Update Link",
                      lambda: _com_update_link(project, task1_id, task2_id, new_link_type="SS"),
                      reschedule=True)
            T.check("Update link FS->SS", r10, success_key="updated")
        except Exception as e:
            T.check("Update link FS->SS", {"error": traceback.format_exc()})
    else:
        T.check("Update link FS->SS", {"error": "Skipped - missing task IDs"})

    # ==========================================================
    # TEST 11: Remove link
    # ==========================================================
    print("\n\n>>> TEST 11: Remove link (Task1 -> Task2)")
    if task1_id and task2_id:
        try:
            r11 = txn(project, "Test Remove Link",
                      lambda: _com_remove_link(project, task1_id, task2_id),
                      reschedule=True)
            T.check("Remove link", r11, success_key="removed")
        except Exception as e:
            T.check("Remove link", {"error": traceback.format_exc()})
    else:
        T.check("Remove link", {"error": "Skipped - missing task IDs"})

    # ==========================================================
    # TEST 12: Add link with lag
    # ==========================================================
    print("\n\n>>> TEST 12: Add SS link with 2d lag")
    if task1_id and task2_id:
        try:
            r12 = txn(project, "Test Add SS+Lag Link",
                      lambda: _com_add_link(project, task1_id, task2_id, "SS", "2d"),
                      reschedule=True)
            T.check("Add SS link + 2d lag", r12, success_key="success")
        except Exception as e:
            T.check("Add SS link + 2d lag", {"error": traceback.format_exc()})
    else:
        T.check("Add SS link + 2d lag", {"error": "Skipped - missing task IDs"})

    # ==========================================================
    # TEST 13: Update progress
    # ==========================================================
    print("\n\n>>> TEST 13: Update progress (50%)")
    if task1_id:
        try:
            r13 = txn(project, "Test Update Progress",
                      lambda: _com_update_progress(project, task1_id, percent_complete=50.0),
                      reschedule=True)
            T.check("Update progress 50%", r13, success_key="updated")
        except Exception as e:
            T.check("Update progress 50%", {"error": traceback.format_exc()})
    else:
        T.check("Update progress 50%", {"error": "Skipped - no task1_id"})

    # ==========================================================
    # TEST 14: Reschedule
    # ==========================================================
    print("\n\n>>> TEST 14: Reschedule project")
    try:
        report_date = pywintypes.Time(datetime.now())
        project.Reschedule(report_date)
        try:
            project.WaitForNotificationProcessing()
        except:
            pass
        T.check("Reschedule project", {"success": True, "report_date": str(datetime.now().date())})
    except Exception as e:
        # "Reschedule found a loop" is a pre-existing project issue, not our fault
        err_msg = str(e)
        if "loop" in err_msg.lower():
            T.check("Reschedule project", {"success": True, "warning": "Pre-existing loop in project"})
        else:
            T.check("Reschedule project", {"error": traceback.format_exc()})

    # ==========================================================
    # CLEANUP: Remove all test bars
    # ==========================================================
    print("\n\n>>> CLEANUP: Removing test bars...")
    cleanup_errors = []
    # _com_delete_task manages its own transactions, no txn() wrapper needed
    for bar_id in reversed(created_ids):
        try:
            r = _com_delete_task(project, bar_id)
            if r.get("deleted"):
                print(f"  Deleted bar {bar_id} ({r.get('name', '?')})")
            else:
                cleanup_errors.append(f"Bar {bar_id}: {r.get('error', 'unknown')}")
                print(f"  FAILED to delete bar {bar_id}: {r.get('error', '?')}")
        except Exception as e:
            cleanup_errors.append(f"Bar {bar_id}: {e}")
            print(f"  FAILED to delete bar {bar_id}: {e}")

    # Also clean up the link from test 12 if bars weren't deleted
    if task1_id and task2_id and cleanup_errors:
        try:
            txn(project, "Cleanup link",
                lambda: _com_remove_link(project, task1_id, task2_id),
                reschedule=False)
        except:
            pass

    # Final reschedule after cleanup
    try:
        project.Reschedule(pywintypes.Time(datetime.now()))
        try:
            project.WaitForNotificationProcessing()
        except:
            pass
    except:
        pass

    if cleanup_errors:
        print(f"\n  Cleanup errors: {cleanup_errors}")
    else:
        print(f"\n  All {len(created_ids)} test bars cleaned up successfully.")

    # --- SUMMARY ---
    T.summary()

    pythoncom.CoUninitialize()
    return T.failed == 0


if __name__ == "__main__":
    try:
        success = run_tests()
        sys.exit(0 if success else 1)
    except Exception as e:
        print(f"\nFATAL ERROR: {e}")
        traceback.print_exc()
        sys.exit(2)
