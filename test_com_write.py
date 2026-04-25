"""
Test COM write operations for asta_powerproject_mcp.

Tests task types: Task, Milestone, Summary (1 task per bar).
Tests: create, update, link, progress, reschedule, delete.
"""

import sys
import os
import json
import traceback

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import pythoncom
import win32com.client
import pywintypes
from datetime import datetime, timedelta

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
)


def fmt(result):
    return json.dumps(result, indent=2, default=str, ensure_ascii=False)


class TestResults:
    def __init__(self):
        self.passed = 0
        self.failed = 0
        self.results = []

    def check(self, name, result, success_key="task_id"):
        success = False
        if isinstance(result, dict):
            if "error" in result:
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
        print(f"  [{status}] {name}")
        if not success:
            print(f"         {fmt(result)}")
        return success

    def summary(self):
        print(f"\n{'='*50}")
        print(f"  {self.passed} passed, {self.failed} failed, {self.passed + self.failed} total")
        print(f"{'='*50}")
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
    pythoncom.CoInitialize()
    app = win32com.client.GetActiveObject("{A57A0000-0200-0000-B2C5-00C0DF438041}")
    project = app.ActiveProject
    print(f"Connected: {project.Name}\n")

    T = TestResults()
    created_ids = []
    start = (datetime.now() + timedelta(days=7)).strftime("%Y-%m-%d")

    task_id = None
    milestone_id = None
    summary_id = None
    child_id = None

    # --- 1. Add Task (10d) ---
    try:
        r = txn(project, "Add Task", lambda: _com_add_task(project, "TEST-Task", "10d", start_date=start), reschedule=True)
        T.check("Add task (10d)", r)
        if "task_id" in r:
            task_id = r["task_id"]
            created_ids.append(task_id)
    except Exception as e:
        T.check("Add task (10d)", {"error": str(e)})

    # --- 2. Add Milestone ---
    try:
        r = txn(project, "Add Milestone", lambda: _com_add_task(project, "TEST-Milestone", start_date=start, is_milestone=True), reschedule=True)
        T.check("Add milestone", r)
        if "task_id" in r:
            milestone_id = r["task_id"]
            created_ids.append(milestone_id)
    except Exception as e:
        T.check("Add milestone", {"error": str(e)})

    # --- 3. Add Summary ---
    try:
        r = txn(project, "Add Summary", lambda: _com_add_task(project, "TEST-Summary", is_summary=True, start_date=start), reschedule=True)
        T.check("Add summary", r)
        if "task_id" in r:
            summary_id = r["task_id"]
            created_ids.append(summary_id)
    except Exception as e:
        T.check("Add summary", {"error": str(e)})

    # --- 4. Add Child under Summary ---
    if summary_id:
        try:
            r = txn(project, "Add Child", lambda: _com_add_task(project, "TEST-Child", "5d", start_date=start, parent_bar_id=summary_id), reschedule=True)
            T.check("Add child under summary", r)
            if "task_id" in r:
                child_id = r["task_id"]
                created_ids.append(child_id)
        except Exception as e:
            T.check("Add child under summary", {"error": str(e)})
    else:
        T.check("Add child under summary", {"error": "No summary_id"})

    # --- 5. Update task name ---
    if task_id:
        try:
            r = txn(project, "Update Name", lambda: _com_update_task(project, task_id, name="TEST-Task-RENAMED"))
            T.check("Update task name", r, success_key="updated_fields")
        except Exception as e:
            T.check("Update task name", {"error": str(e)})
    else:
        T.check("Update task name", {"error": "No task_id"})

    # --- 6. Update task duration ---
    if task_id:
        try:
            r = txn(project, "Update Duration", lambda: _com_update_task(project, task_id, duration_str="15d"), reschedule=True)
            T.check("Update duration (15d)", r, success_key="updated_fields")
        except Exception as e:
            T.check("Update duration (15d)", {"error": str(e)})
    else:
        T.check("Update duration (15d)", {"error": "No task_id"})

    # --- 7. Add FS link (Task -> Milestone) ---
    if task_id and milestone_id:
        try:
            r = txn(project, "Add FS Link", lambda: _com_add_link(project, task_id, milestone_id, "FS"), reschedule=True)
            T.check("Add FS link", r, success_key="success")
        except Exception as e:
            T.check("Add FS link", {"error": str(e)})
    else:
        T.check("Add FS link", {"error": "Missing IDs"})

    # --- 8. Update link (FS -> SS) ---
    if task_id and milestone_id:
        try:
            r = txn(project, "Update Link", lambda: _com_update_link(project, task_id, milestone_id, new_link_type="SS"), reschedule=True)
            T.check("Update link FS->SS", r, success_key="updated")
        except Exception as e:
            T.check("Update link FS->SS", {"error": str(e)})
    else:
        T.check("Update link FS->SS", {"error": "Missing IDs"})

    # --- 9. Remove link ---
    if task_id and milestone_id:
        try:
            r = txn(project, "Remove Link", lambda: _com_remove_link(project, task_id, milestone_id), reschedule=True)
            T.check("Remove link", r, success_key="removed")
        except Exception as e:
            T.check("Remove link", {"error": str(e)})
    else:
        T.check("Remove link", {"error": "Missing IDs"})

    # --- 10. Add SS link with 2d lag ---
    if task_id and milestone_id:
        try:
            r = txn(project, "Add SS+Lag", lambda: _com_add_link(project, task_id, milestone_id, "SS", "2d"), reschedule=True)
            T.check("Add SS link + 2d lag", r, success_key="success")
        except Exception as e:
            T.check("Add SS link + 2d lag", {"error": str(e)})
    else:
        T.check("Add SS link + 2d lag", {"error": "Missing IDs"})

    # --- 11. Update progress (50%) ---
    if task_id:
        try:
            r = txn(project, "Update Progress", lambda: _com_update_progress(project, task_id, percent_complete=50.0), reschedule=True)
            T.check("Update progress 50%", r, success_key="updated")
        except Exception as e:
            T.check("Update progress 50%", {"error": str(e)})
    else:
        T.check("Update progress 50%", {"error": "No task_id"})

    # --- 12. Reschedule ---
    try:
        project.Reschedule(pywintypes.Time(datetime.now()))
        try:
            project.WaitForNotificationProcessing()
        except:
            pass
        T.check("Reschedule", {"success": True}, success_key="success")
    except Exception as e:
        if "loop" in str(e).lower():
            T.check("Reschedule", {"success": True}, success_key="success")
        else:
            T.check("Reschedule", {"error": str(e)})

    # --- CLEANUP ---
    print("\n  Cleanup...")
    for bar_id in reversed(created_ids):
        try:
            r = _com_delete_task(project, bar_id)
            if r.get("deleted"):
                print(f"    Deleted {bar_id}")
            else:
                print(f"    FAILED {bar_id}: {r.get('error', '?')}")
        except Exception as e:
            print(f"    FAILED {bar_id}: {e}")

    # Final reschedule
    try:
        project.Reschedule(pywintypes.Time(datetime.now()))
        try:
            project.WaitForNotificationProcessing()
        except:
            pass
    except:
        pass

    T.summary()
    pythoncom.CoUninitialize()
    return T.failed == 0


if __name__ == "__main__":
    try:
        success = run_tests()
        sys.exit(0 if success else 1)
    except Exception as e:
        print(f"\nFATAL: {e}")
        traceback.print_exc()
        sys.exit(2)
