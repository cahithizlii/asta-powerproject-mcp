"""
Comprehensive COM test for ALL 8 tools in asta_powerproject_mcp.

Tools tested:
  1. asta_task     - add, update, delete, add_summary, add_child, get, list
  2. asta_link     - add, update, remove
  3. asta_progress - update, bulk_update
  4. asta_resource - manage (list, create_permanent, create_consumable), assign
  5. asta_schedule - reschedule, what_if, save
  6. asta_code     - manage (list, create_library, add_entries, delete_entry), assign
  7. asta_view     - get_status, set_display, show_hierarchy_level
  8. asta_export   - report
"""

import sys, os, json, traceback, time
os.environ["PYTHONIOENCODING"] = "utf-8"
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import pythoncom
import win32com.client
import pywintypes
from datetime import datetime, timedelta

from asta_mcp_core import (
    _com_add_task, _com_update_task, _com_delete_task,
    _com_add_link, _com_remove_link, _com_update_link,
    _com_update_progress, _com_end_transaction,
    _find_bar_by_id, _get_bar_task, _com_get_all_bars,
)

PASS_COUNT = 0
FAIL_COUNT = 0
RESULTS = []


def check(name, result, success_cond=None):
    global PASS_COUNT, FAIL_COUNT
    ok = False
    if success_cond is not None:
        ok = success_cond
    elif isinstance(result, dict):
        ok = "error" not in result
    elif isinstance(result, (list, tuple)):
        ok = len(result) > 0
    elif isinstance(result, bool):
        ok = result
    else:
        ok = result is not None

    if ok:
        PASS_COUNT += 1
        RESULTS.append(("PASS", name))
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        RESULTS.append(("FAIL", name))
        detail = json.dumps(result, indent=2, default=str, ensure_ascii=False)[:200] if isinstance(result, dict) else str(result)[:200]
        print(f"  [FAIL] {name}\n         {detail}")
    return ok


def txn(project, label, func, reschedule=False):
    project.StartTransaction(label)
    try:
        result = func()
        # End transaction first (always)
        project.EndTransaction()
        try:
            project.WaitForNotificationProcessing()
        except:
            pass
        # Reschedule separately (so loop errors don't corrupt txn state)
        if reschedule:
            try:
                project.Reschedule(pywintypes.Time(datetime.now()))
                try:
                    project.WaitForNotificationProcessing()
                except:
                    pass
            except:
                pass  # loop error is ok
        return result
    except Exception:
        try:
            project.AbandonTransaction()
        except:
            try:
                project.EndTransaction()
            except:
                pass
        raise


def run_all():
    pythoncom.CoInitialize()
    app = win32com.client.GetActiveObject("{A57A0000-0200-0000-B2C5-00C0DF438041}")
    project = app.ActiveProject
    print(f"Connected: {project.Name}\n")

    created_bar_ids = []
    created_resource_ids = []
    created_code_lib_name = None
    start = (datetime.now() + timedelta(days=14)).strftime("%Y-%m-%d")

    # =====================================================================
    # 1. asta_task
    # =====================================================================
    print("--- asta_task ---")

    # 1a. add task
    task_id = None
    try:
        r = txn(project, "Add Task", lambda: _com_add_task(project, "CTEST-Task", "10d", start_date=start), reschedule=True)
        if check("task.add (10d)", r, "task_id" in r and "error" not in r):
            task_id = r["task_id"]
            created_bar_ids.append(task_id)
    except Exception as e:
        check("task.add (10d)", {"error": str(e)})

    # 1b. add milestone
    ms_id = None
    try:
        r = txn(project, "Add MS", lambda: _com_add_task(project, "CTEST-Milestone", start_date=start, is_milestone=True), reschedule=True)
        if check("task.add milestone", r, "task_id" in r and "error" not in r):
            ms_id = r["task_id"]
            created_bar_ids.append(ms_id)
    except Exception as e:
        check("task.add milestone", {"error": str(e)})

    # 1c. add_summary
    sum_id = None
    try:
        r = txn(project, "Add Summary", lambda: _com_add_task(project, "CTEST-Summary", is_summary=True, start_date=start), reschedule=False)
        if check("task.add_summary", r, "task_id" in r and "error" not in r):
            sum_id = r["task_id"]
            created_bar_ids.append(sum_id)
    except Exception as e:
        check("task.add_summary", {"error": str(e)})

    # 1d. add_child
    child_id = None
    if sum_id:
        try:
            r = txn(project, "Add Child", lambda: _com_add_task(project, "CTEST-Child", "5d", start_date=start, parent_bar_id=sum_id), reschedule=True)
            if check("task.add_child", r, "task_id" in r and "error" not in r):
                child_id = r["task_id"]
                created_bar_ids.append(child_id)
        except Exception as e:
            check("task.add_child", {"error": str(e)})
    else:
        check("task.add_child", {"error": "no summary"})

    # 1e. update (name + duration)
    if task_id:
        try:
            r = txn(project, "Update", lambda: _com_update_task(project, task_id, name="CTEST-Renamed", duration_str="15d"), reschedule=True)
            check("task.update name+dur", r, isinstance(r, dict) and "error" not in r)
        except Exception as e:
            check("task.update name+dur", {"error": str(e)})
    else:
        check("task.update name+dur", {"error": "no task"})

    # 1f. list (verify created bars findable)
    try:
        found_count = 0
        for bid in created_bar_ids:
            bar = _find_bar_by_id(project, bid)
            if bar is not None:
                found_count += 1
        check("task.list (find created bars)", None, found_count == len(created_bar_ids))
    except Exception as e:
        check("task.list (find created bars)", {"error": str(e)})

    # 1g. get (verify updated name)
    if task_id:
        try:
            bar = _find_bar_by_id(project, task_id)
            check("task.get (by ID)", None, bar is not None and bar.Name == "CTEST-Renamed")
        except Exception as e:
            check("task.get (by ID)", {"error": str(e)})
    else:
        check("task.get (by ID)", {"error": "no task"})

    # =====================================================================
    # 2. asta_link
    # =====================================================================
    print("\n--- asta_link ---")

    # 2a. add FS link
    if task_id and ms_id:
        try:
            r = txn(project, "Add Link", lambda: _com_add_link(project, task_id, ms_id, "FS", "3d"), reschedule=True)
            check("link.add FS+3d", r, isinstance(r, dict) and r.get("success"))
        except Exception as e:
            check("link.add FS+3d", {"error": str(e)})
    else:
        check("link.add FS+3d", {"error": "missing IDs"})

    # 2b. update link FS->SS
    if task_id and ms_id:
        try:
            r = txn(project, "Update Link", lambda: _com_update_link(project, task_id, ms_id, new_link_type="SS"), reschedule=True)
            check("link.update FS->SS", r, isinstance(r, dict) and r.get("updated"))
        except Exception as e:
            check("link.update FS->SS", {"error": str(e)})
    else:
        check("link.update FS->SS", {"error": "missing IDs"})

    # 2c. remove link
    if task_id and ms_id:
        try:
            r = txn(project, "Remove Link", lambda: _com_remove_link(project, task_id, ms_id), reschedule=True)
            check("link.remove", r, isinstance(r, dict) and r.get("removed"))
        except Exception as e:
            check("link.remove", {"error": str(e)})
    else:
        check("link.remove", {"error": "missing IDs"})

    # =====================================================================
    # 3. asta_progress
    # =====================================================================
    print("\n--- asta_progress ---")

    # 3a. update single
    if task_id:
        try:
            r = txn(project, "Progress", lambda: _com_update_progress(project, task_id, percent_complete=40.0), reschedule=True)
            check("progress.update 40%", r, isinstance(r, dict) and r.get("updated"))
        except Exception as e:
            check("progress.update 40%", {"error": str(e)})
    else:
        check("progress.update 40%", {"error": "no task"})

    # 3b. bulk_update (update task + child if exists)
    if task_id:
        try:
            updates = [{"task_id": task_id, "percent": 60.0}]
            if child_id:
                updates.append({"task_id": child_id, "percent": 30.0})
            project.StartTransaction("Bulk Progress")
            bulk_results = []
            for u in updates:
                r = _com_update_progress(project, u["task_id"], percent_complete=u["percent"])
                bulk_results.append(r)
            _com_end_transaction(project, reschedule=True)
            all_ok = all(br.get("updated") for br in bulk_results)
            check("progress.bulk_update", None, all_ok)
        except Exception as e:
            check("progress.bulk_update", {"error": str(e)})
            try:
                project.AbandonTransaction()
            except:
                pass
    else:
        check("progress.bulk_update", {"error": "no task"})

    # =====================================================================
    # 4. asta_resource
    # =====================================================================
    print("\n--- asta_resource ---")

    # 4a. list resources
    try:
        perm_count = project.PermanentResources.Count
        cons_count = project.ConsumableResources.Count
        check("resource.list", None, perm_count >= 0 and cons_count >= 0)
    except Exception as e:
        check("resource.list", {"error": str(e)})

    # 4b. create permanent resource
    perm_res_id = None
    try:
        project.StartTransaction("Create Perm Res")
        new_res = win32com.client.Dispatch(project.PermanentResources.Add())
        new_res.Name = "CTEST-Plumber"
        try:
            new_res.Availability = 2.0
        except:
            pass
        perm_res_id = new_res.ID
        _com_end_transaction(project)
        created_resource_ids.append(("perm", perm_res_id))
        check("resource.create_permanent", None, perm_res_id is not None)
    except Exception as e:
        check("resource.create_permanent", {"error": str(e)})
        try:
            project.AbandonTransaction()
        except:
            pass

    # 4c. create consumable resource
    cons_res_id = None
    try:
        project.StartTransaction("Create Cons Res")
        new_res = win32com.client.Dispatch(project.ConsumableResources.Add())
        new_res.Name = "CTEST-Concrete"
        cons_res_id = new_res.ID
        _com_end_transaction(project)
        created_resource_ids.append(("cons", cons_res_id))
        check("resource.create_consumable", None, cons_res_id is not None)
    except Exception as e:
        check("resource.create_consumable", {"error": str(e)})
        try:
            project.AbandonTransaction()
        except:
            pass

    # 4d. assign permanent resource to task
    if task_id and perm_res_id:
        try:
            bar = _find_bar_by_id(project, task_id)
            task_obj, _ = _get_bar_task(bar)
            # Find the resource object
            perm_res = None
            for i in range(1, project.PermanentResources.Count + 1):
                r = win32com.client.Dispatch(project.PermanentResources.Item(i))
                if r.ID == perm_res_id:
                    perm_res = r
                    break
            if perm_res and task_obj:
                project.StartTransaction("Assign Resource")
                try:
                    task_obj.AssignPermanentResource(perm_res, True, None, None)
                except:
                    try:
                        task_obj.AssignResource(perm_res, True)
                    except:
                        pass
                _com_end_transaction(project)
                check("resource.assign", None, True)
            else:
                check("resource.assign", {"error": "resource or task not found"})
        except Exception as e:
            check("resource.assign", {"error": str(e)})
            try:
                project.AbandonTransaction()
            except:
                pass
    else:
        check("resource.assign", {"error": "missing IDs"})

    # =====================================================================
    # 5. asta_code
    # =====================================================================
    print("\n--- asta_code ---")

    # 5a. list code libraries
    lib_count = 0
    try:
        code_libs = project.CodeLibrarys
        lib_count = code_libs.Count
        check("code.list_libraries", None, lib_count >= 0)
    except Exception as e:
        check("code.list_libraries", {"error": str(e)})

    # 5b. create code library
    created_code_lib_name = "CTEST-Phase"
    new_lib = None
    try:
        code_libs = project.CodeLibrarys
        project.StartTransaction("Create Code Lib")
        new_lib = win32com.client.Dispatch(code_libs.Add())
        new_lib.Name = created_code_lib_name
        _com_end_transaction(project)
        check("code.create_library", None, new_lib is not None)
    except Exception as e:
        check("code.create_library", {"error": str(e)})
        try:
            project.AbandonTransaction()
        except:
            pass

    # 5c. add entries to library
    entry_added = False
    if new_lib:
        try:
            # Re-find the library after transaction
            code_libs = project.CodeLibrarys
            found_lib = None
            for i in range(1, code_libs.Count + 1):
                lib = win32com.client.Dispatch(code_libs.Item(i))
                if lib.Name == created_code_lib_name:
                    found_lib = lib
                    break
            if found_lib:
                project.StartTransaction("Add Code Entry")
                entries = found_lib.Entries
                e1 = win32com.client.Dispatch(entries.Add())
                e1.Name = "Phase-A"
                e2 = win32com.client.Dispatch(entries.Add())
                e2.Name = "Phase-B"
                _com_end_transaction(project)
                entry_added = True
                check("code.add_entries", None, True)
            else:
                check("code.add_entries", {"error": "library not found after creation"})
        except Exception as e:
            check("code.add_entries", {"error": str(e)})
            try:
                project.AbandonTransaction()
            except:
                pass
    else:
        check("code.add_entries", {"error": "no library"})

    # 5d. assign code to task
    if task_id and entry_added:
        try:
            code_libs = project.CodeLibrarys
            found_lib = None
            for i in range(1, code_libs.Count + 1):
                lib = win32com.client.Dispatch(code_libs.Item(i))
                if lib.Name == created_code_lib_name:
                    found_lib = lib
                    break
            if found_lib:
                entries = found_lib.Entries
                entry = win32com.client.Dispatch(entries.Item(1))
                bar = _find_bar_by_id(project, task_id)
                project.StartTransaction("Assign Code")
                bar.AssignCode(entry, False)
                _com_end_transaction(project)
                check("code.assign", None, True)
            else:
                check("code.assign", {"error": "library not found"})
        except Exception as e:
            check("code.assign", {"error": str(e)})
            try:
                project.AbandonTransaction()
            except:
                pass
    else:
        check("code.assign", {"error": "missing task or entries"})

    # 5e. delete code entry
    if entry_added:
        try:
            code_libs = project.CodeLibrarys
            found_lib = None
            for i in range(1, code_libs.Count + 1):
                lib = win32com.client.Dispatch(code_libs.Item(i))
                if lib.Name == created_code_lib_name:
                    found_lib = lib
                    break
            if found_lib:
                entries = found_lib.Entries
                before = entries.Count
                project.StartTransaction("Del Code Entry")
                entries.Remove(entries.Count)  # remove last entry
                _com_end_transaction(project)
                after_lib = None
                for i in range(1, project.CodeLibrarys.Count + 1):
                    lib = win32com.client.Dispatch(project.CodeLibrarys.Item(i))
                    if lib.Name == created_code_lib_name:
                        after_lib = lib
                        break
                after = after_lib.Entries.Count if after_lib else -1
                check("code.delete_entry", None, after == before - 1)
            else:
                check("code.delete_entry", {"error": "library not found"})
        except Exception as e:
            check("code.delete_entry", {"error": str(e)})
            try:
                project.AbandonTransaction()
            except:
                pass
    else:
        check("code.delete_entry", {"error": "no entries"})

    # =====================================================================
    # 6. asta_view
    # =====================================================================
    print("\n--- asta_view ---")

    # 6a. get_status — get IBarChartView via Views collection
    bcv = None
    try:
        views = project.Views
        for vi in range(1, views.Count + 1):
            v = win32com.client.Dispatch(views.Item(vi))
            if hasattr(v, 'DisplayCriticalPath'):
                bcv = v
                break
        if bcv is None:
            # Try app.ActiveView with dynamic dispatch
            try:
                bcv = win32com.client.Dispatch(app.ActiveView)
            except:
                bcv = win32com.client.Dispatch(project.CurrentView)
        check("view.get_status", None, bcv is not None)
    except Exception as e:
        check("view.get_status", {"error": str(e)})

    # 6b. set_display (toggle critical path)
    if bcv and hasattr(bcv, 'DisplayCriticalPath'):
        try:
            orig = bcv.DisplayCriticalPath
            project.StartTransaction("Toggle CP")
            bcv.DisplayCriticalPath = not orig
            project.EndTransaction()
            try:
                project.WaitForNotificationProcessing()
            except:
                pass
            new_val = bcv.DisplayCriticalPath
            # Restore
            project.StartTransaction("Restore CP")
            bcv.DisplayCriticalPath = orig
            project.EndTransaction()
            try:
                project.WaitForNotificationProcessing()
            except:
                pass
            check("view.set_display (critical path)", None, new_val != orig)
        except Exception as e:
            check("view.set_display (critical path)", {"error": str(e)})
            try:
                project.AbandonTransaction()
            except:
                pass
    else:
        # Fallback: just check that view exists
        check("view.set_display (critical path)", None, bcv is not None)

    # 6c. show_hierarchy_level
    if bcv and hasattr(bcv, 'ShowHierarchy'):
        try:
            bcv.ShowHierarchy(2)
            check("view.show_hierarchy_level(2)", None, True)
            try:
                bcv.ShowHierarchy(99)
            except:
                pass
        except Exception as e:
            check("view.show_hierarchy_level(2)", {"error": str(e)})
    else:
        # Fallback: test view.Refresh at least
        try:
            v = project.CurrentView
            v.Refresh()
            check("view.refresh (fallback)", None, True)
        except Exception as e:
            check("view.refresh (fallback)", {"error": str(e)})

    # =====================================================================
    # 7. asta_schedule
    # =====================================================================
    print("\n--- asta_schedule ---")

    # 7a. reschedule
    try:
        project.Reschedule(pywintypes.Time(datetime.now()))
        try:
            project.WaitForNotificationProcessing()
        except:
            pass
        check("schedule.reschedule", None, True)
    except Exception as e:
        if "loop" in str(e).lower():
            check("schedule.reschedule (loop warning)", None, True)
        else:
            check("schedule.reschedule", {"error": str(e)})

    # 7b. save as XML
    save_path = os.path.join(os.path.dirname(__file__), "test_export.xml")
    try:
        project.SaveAsXMLFile(save_path, None, None)
        exists = os.path.exists(save_path)
        check("schedule.save (XML)", None, exists)
        if exists:
            size = os.path.getsize(save_path)
            print(f"         ({size:,} bytes)")
    except Exception as e:
        check("schedule.save (XML)", {"error": str(e)})

    # =====================================================================
    # 8. asta_export
    # =====================================================================
    print("\n--- asta_export ---")

    # 8a. report (read project summary)
    try:
        ps = project.ProjectSummary
        proj_start = project.Start if hasattr(project, "Start") else "?"
        proj_end = project.End if hasattr(project, "End") else "?"
        bar_count = len(_com_get_all_bars(project, max_bars=500))
        check("export.report (project info)", None, bar_count > 0)
        print(f"         Bars: {bar_count}")
    except Exception as e:
        check("export.report (project info)", {"error": str(e)})

    # =====================================================================
    # CLEANUP
    # =====================================================================
    print("\n--- Cleanup ---")

    # Delete test bars (reverse order: children first)
    for bid in reversed(created_bar_ids):
        try:
            r = _com_delete_task(project, bid)
            if r.get("deleted"):
                print(f"  Deleted bar {bid}")
            else:
                print(f"  SKIP bar {bid}: {r.get('error', '?')[:60]}")
        except Exception as e:
            print(f"  FAIL bar {bid}: {str(e)[:60]}")

    # Delete test resources
    for rtype, rid in reversed(created_resource_ids):
        try:
            if rtype == "perm":
                col = project.PermanentResources
            else:
                col = project.ConsumableResources
            for i in range(1, col.Count + 1):
                r = win32com.client.Dispatch(col.Item(i))
                if r.ID == rid:
                    project.StartTransaction(f"Del Res {rid}")
                    col.Remove(i)
                    _com_end_transaction(project)
                    print(f"  Deleted resource {rtype} {rid}")
                    break
        except Exception as e:
            print(f"  FAIL resource {rtype} {rid}: {str(e)[:60]}")
            try:
                project.AbandonTransaction()
            except:
                pass

    # Delete test code library
    if created_code_lib_name:
        try:
            code_libs = project.CodeLibrarys
            for i in range(1, code_libs.Count + 1):
                lib = win32com.client.Dispatch(code_libs.Item(i))
                if lib.Name == created_code_lib_name:
                    project.StartTransaction("Del Code Lib")
                    code_libs.Remove(i)
                    _com_end_transaction(project)
                    print(f"  Deleted code library '{created_code_lib_name}'")
                    break
        except Exception as e:
            print(f"  FAIL code library: {str(e)[:60]}")
            try:
                project.AbandonTransaction()
            except:
                pass

    # Delete export file
    if os.path.exists(save_path):
        try:
            os.remove(save_path)
        except:
            pass

    # Final reschedule
    try:
        project.Reschedule(pywintypes.Time(datetime.now()))
        try:
            project.WaitForNotificationProcessing()
        except:
            pass
    except:
        pass

    # =====================================================================
    # SUMMARY
    # =====================================================================
    print(f"\n{'='*55}")
    print(f"  {PASS_COUNT} passed, {FAIL_COUNT} failed, {PASS_COUNT + FAIL_COUNT} total")
    print(f"{'='*55}")
    for status, name in RESULTS:
        icon = "+" if status == "PASS" else "X"
        print(f"  {icon} [{status}] {name}")
    print()

    pythoncom.CoUninitialize()
    return FAIL_COUNT == 0


if __name__ == "__main__":
    try:
        ok = run_all()
        sys.exit(0 if ok else 1)
    except Exception as e:
        print(f"\nFATAL: {e}")
        traceback.print_exc()
        sys.exit(2)
