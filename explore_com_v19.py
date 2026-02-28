"""
COM Explorer v19 — CORRECT architecture understanding
KEY FINDINGS:
  - bar.ExpandedTask ALWAYS returns ROOT task (ID=1083) — WRONG for children!
  - bar.Tasks(1) returns the bar's own task — CORRECT
  - bar.Tasks(1) returns IExpandedTask for summaries, ITask for leaves
  - Satinalma bar has NO Tasks(1) entry (different bar type?)
  - ImposedStart only on IExpandedTask, not ITask

APPROACH:
  - Use Insaat (index 3) as parent — has working Tasks(1)
  - After creating new bar + EndTransaction, test Tasks(1)
  - For dates: use ImposedStart if IExpandedTask, else StartConstraintDate
"""
import pythoncom
import win32com.client
import pywintypes
from datetime import datetime
import traceback


def connect():
    pythoncom.CoInitialize()
    clsid = "{A57A0000-0200-0000-B2C5-00C0DF438041}"
    app = win32com.client.GetActiveObject(clsid)
    project = app.ActiveProject
    print(f"Connected to: {project.Name}")
    return app, project


def wait(project):
    try:
        project.WaitForNotificationProcessing()
    except Exception:
        pass


def get_root_task(project):
    bar = project.Bars.Item(1)
    return win32com.client.Dispatch(bar.Tasks(1))


def get_summary_bar(project, index):
    """Get Nth summary bar under root (1-based). Fresh fetch."""
    root_task = get_root_task(project)
    return win32com.client.Dispatch(root_task.ChildBars.Item(index))


def get_summary_task(project, index):
    """Get Nth summary bar's task. Fresh fetch."""
    bar = get_summary_bar(project, index)
    return win32com.client.Dispatch(bar.Tasks(1))


def find_bar_by_id(project, bar_id):
    """Find bar by ID. Fresh traversal."""
    root_task = get_root_task(project)
    return _search(root_task, bar_id, 0, 5)


def _search(parent_task, target_id, depth, max_depth):
    try:
        cb = parent_task.ChildBars
        for i in range(1, cb.Count + 1):
            b = win32com.client.Dispatch(cb.Item(i))
            if b.ID == target_id:
                return b
            if depth < max_depth:
                try:
                    ct = win32com.client.Dispatch(b.Tasks(1))
                    r = _search(ct, target_id, depth + 1, max_depth)
                    if r:
                        return r
                except Exception:
                    pass
    except Exception:
        pass
    return None


def get_bar_task(bar):
    """Get the task for a bar (Tasks(1) first, ExpandedTask fallback).
    Returns (task, is_expanded_task)."""
    # Tasks(1) first
    try:
        t = win32com.client.Dispatch(bar.Tasks(1))
        is_et = type(t).__name__ == 'IExpandedTask'
        return t, is_et
    except Exception:
        pass
    # ExpandedTask fallback — BUT this returns root's task!
    # Only use if we can verify it's the correct one
    return None, False


def test_create_under_insaat(project):
    """Create bars under Insaat (index 3)."""
    print("\n" + "=" * 80)
    print("TEST 1: CREATE BARS UNDER INSAAT")
    print("=" * 80)

    # List Insaat children
    insaat = get_summary_bar(project, 3)
    insaat_task = win32com.client.Dispatch(insaat.Tasks(1))
    print(f"Insaat: ID={insaat.ID}, Name={insaat.Name[:40]}")
    print(f"  Task: ID={insaat_task.ID}, type={type(insaat_task).__name__}")
    print(f"  ChildBars: {insaat_task.ChildBars.Count}")

    # Show first 5 children
    for i in range(1, min(insaat_task.ChildBars.Count + 1, 6)):
        cb = win32com.client.Dispatch(insaat_task.ChildBars.Item(i))
        try:
            ct = win32com.client.Dispatch(cb.Tasks(1))
            print(f"  [{i}] ID={cb.ID}, Name={cb.Name[:30]}, "
                  f"TaskID={ct.ID}, Type={type(ct).__name__}")
        except Exception:
            print(f"  [{i}] ID={cb.ID}, Name={cb.Name[:30]}, Tasks(1)=FAILED")

    # Create bar A
    print(f"\n--- Create A ---")
    project.StartTransaction("Create A")
    bar_a_id = None
    try:
        insaat_task = get_summary_task(project, 3)
        new_a = win32com.client.Dispatch(insaat_task.ChildBars.Add())
        new_a.Name = "V19_TEST_A"
        bar_a_id = new_a.ID
        print(f"  Created: BarID={bar_a_id}")

        # In-txn checks
        try:
            t = win32com.client.Dispatch(new_a.Tasks(1))
            print(f"  In-txn Tasks(1): ID={t.ID}, type={type(t).__name__}")
        except Exception as e:
            print(f"  In-txn Tasks(1): FAILED - {str(e)[:50]}")

        project.EndTransaction()
        wait(project)
    except Exception as e:
        print(f"  Error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass
        return []

    # After-txn: find bar and test Tasks(1)
    print(f"\n--- After commit: check bar A ---")
    bar_a = find_bar_by_id(project, bar_a_id)
    if bar_a:
        print(f"  Found: ID={bar_a.ID}, Name={bar_a.Name[:30]}")
        try:
            ta = win32com.client.Dispatch(bar_a.Tasks(1))
            print(f"  Tasks(1): ID={ta.ID}, type={type(ta).__name__}")
            print(f"  Start={ta.Start}, End={ta.End}")
            try:
                print(f"  Duration={ta.GetUserDuration().Hours}h")
            except Exception:
                pass
            # Check if IExpandedTask
            if hasattr(ta, 'ImposedStart'):
                print(f"  HAS ImposedStart (IExpandedTask)")
            else:
                print(f"  NO ImposedStart (ITask)")
                # Check for ConvertToExpandedTask
                if hasattr(ta, 'ConvertToExpandedTask'):
                    print(f"  HAS ConvertToExpandedTask")
        except Exception as e:
            print(f"  Tasks(1) FAILED: {str(e)[:60]}")
    else:
        print(f"  NOT FOUND!")
        return []

    # Create bar B
    print(f"\n--- Create B ---")
    project.StartTransaction("Create B")
    bar_b_id = None
    try:
        insaat_task = get_summary_task(project, 3)
        new_b = win32com.client.Dispatch(insaat_task.ChildBars.Add())
        new_b.Name = "V19_TEST_B"
        bar_b_id = new_b.ID
        print(f"  Created: BarID={bar_b_id}")
        project.EndTransaction()
        wait(project)
    except Exception as e:
        print(f"  Error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass
        return [bar_a_id]

    # Check bar B
    bar_b = find_bar_by_id(project, bar_b_id)
    if bar_b:
        try:
            tb = win32com.client.Dispatch(bar_b.Tasks(1))
            print(f"  B Tasks(1): ID={tb.ID}, type={type(tb).__name__}")
        except Exception as e:
            print(f"  B Tasks(1) FAILED: {str(e)[:60]}")

    return [bar_a_id, bar_b_id]


def test_duration(project, bar_id):
    """Set duration on a new bar."""
    print("\n" + "=" * 80)
    print("TEST 2: SET DURATION")
    print("=" * 80)

    bar = find_bar_by_id(project, bar_id)
    task, is_et = get_bar_task(bar)
    if not task:
        print(f"  No task for bar {bar_id}!")
        return

    print(f"  Task: ID={task.ID}, type={'IExpandedTask' if is_et else 'ITask'}")
    try:
        print(f"  Current duration: {task.GetUserDuration().Hours}h")
    except Exception:
        pass

    project.StartTransaction("Dur")
    try:
        bar = find_bar_by_id(project, bar_id)
        task = win32com.client.Dispatch(bar.Tasks(1))
        dur = task.GetDurationFromString("10d")
        task.SetUserDuration(dur)
        print(f"  SetUserDuration(10d) => OK")
        project.EndTransaction()
        wait(project)
    except Exception as e:
        print(f"  Error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass
        return

    project.Reschedule(pywintypes.Time(datetime.now()))
    wait(project)

    bar = find_bar_by_id(project, bar_id)
    task = win32com.client.Dispatch(bar.Tasks(1))
    dur_h = task.GetUserDuration().Hours
    print(f"  After: Duration={dur_h}h, Start={task.Start}, End={task.End}")
    if dur_h == 80.0:
        print(f"  *** DURATION CORRECT (10d = 80h) ***")
    else:
        print(f"  Duration unexpected: expected 80h, got {dur_h}h")


def test_dates(project, bar_id):
    """Set dates using various methods."""
    print("\n" + "=" * 80)
    print("TEST 3: SET DATES")
    print("=" * 80)

    bar = find_bar_by_id(project, bar_id)
    task = win32com.client.Dispatch(bar.Tasks(1))
    is_et = type(task).__name__ == 'IExpandedTask'
    print(f"  Task: ID={task.ID}, is_et={is_et}")
    print(f"  Start: {task.Start}")

    # Method A: ImposedStart (only on IExpandedTask)
    if is_et:
        print(f"\n--- A: ImposedStart ---")
        project.StartTransaction("ImposedStart")
        try:
            bar = find_bar_by_id(project, bar_id)
            task = win32com.client.Dispatch(bar.Tasks(1))
            task.ImposedStart = pywintypes.Time(datetime(2026, 7, 1))
            print(f"  ImposedStart = 2026-07-01 => OK")
            project.EndTransaction()
            wait(project)
        except Exception as e:
            print(f"  Error: {e}")
            try:
                project.AbandonTransaction()
            except Exception:
                pass

        project.Reschedule(pywintypes.Time(datetime.now()))
        wait(project)

        bar = find_bar_by_id(project, bar_id)
        task = win32com.client.Dispatch(bar.Tasks(1))
        print(f"  After: Start={task.Start}, End={task.End}, Constraint={task.Constraint}")

    # Method B: StartConstraintDate (available on both?)
    print(f"\n--- B: StartConstraintDate ---")
    project.StartTransaction("StartConstraint")
    try:
        bar = find_bar_by_id(project, bar_id)
        task = win32com.client.Dispatch(bar.Tasks(1))
        # Remove existing constraint first
        try:
            task.RemoveConstraint()
        except Exception:
            pass
        task.StartConstraintDate = pywintypes.Time(datetime(2026, 8, 3))
        print(f"  StartConstraintDate = 2026-08-03 => OK")
        print(f"  Readback: {task.StartConstraintDate}")
        print(f"  Constraint: {task.Constraint}")
        project.EndTransaction()
        wait(project)
    except Exception as e:
        print(f"  Error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass

    project.Reschedule(pywintypes.Time(datetime.now()))
    wait(project)

    bar = find_bar_by_id(project, bar_id)
    task = win32com.client.Dispatch(bar.Tasks(1))
    print(f"  After: Start={task.Start}, End={task.End}, Constraint={task.Constraint}")

    # Method C: Restore with ImposedStart
    if is_et:
        print(f"\n--- C: Restore ImposedStart 2026-07-01 ---")
        project.StartTransaction("Restore")
        try:
            bar = find_bar_by_id(project, bar_id)
            task = win32com.client.Dispatch(bar.Tasks(1))
            task.RemoveConstraint()
            task.ImposedStart = pywintypes.Time(datetime(2026, 7, 1))
            project.EndTransaction()
            wait(project)
        except Exception as e:
            print(f"  Error: {e}")
            try:
                project.AbandonTransaction()
            except Exception:
                pass

        project.Reschedule(pywintypes.Time(datetime.now()))
        wait(project)
        bar = find_bar_by_id(project, bar_id)
        task = win32com.client.Dispatch(bar.Tasks(1))
        print(f"  Restored: Start={task.Start}, Constraint={task.Constraint}")


def test_linking(project, bar_a_id, bar_b_id):
    """Link two bars."""
    print("\n" + "=" * 80)
    print("TEST 4: LINK A -> B")
    print("=" * 80)

    # Set duration on B
    project.StartTransaction("Dur B")
    try:
        bar_b = find_bar_by_id(project, bar_b_id)
        tb = win32com.client.Dispatch(bar_b.Tasks(1))
        dur = tb.GetDurationFromString("5d")
        tb.SetUserDuration(dur)
        print(f"  B SetUserDuration(5d) => OK")
        project.EndTransaction()
        wait(project)
    except Exception as e:
        print(f"  B dur error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass

    project.Reschedule(pywintypes.Time(datetime.now()))
    wait(project)

    # Link
    project.StartTransaction("Link")
    try:
        bar_a = find_bar_by_id(project, bar_a_id)
        bar_b = find_bar_by_id(project, bar_b_id)
        ta = win32com.client.Dispatch(bar_a.Tasks(1))
        tb = win32com.client.Dispatch(bar_b.Tasks(1))
        print(f"  A: TaskID={ta.ID}")
        print(f"  B: TaskID={tb.ID}")
        link = ta.LinkTo(tb)
        if link:
            ld = win32com.client.Dispatch(link)
            print(f"  LINK CREATED! ID={ld.ID}, type={ld.type}")
        else:
            print(f"  LinkTo returned None!")
        project.EndTransaction()
        wait(project)
    except Exception as e:
        print(f"  Link error: {e}")
        traceback.print_exc()
        try:
            project.AbandonTransaction()
        except Exception:
            pass
        return

    project.Reschedule(pywintypes.Time(datetime.now()))
    wait(project)

    # Verify
    print(f"\n--- VERIFY ---")
    bar_a = find_bar_by_id(project, bar_a_id)
    bar_b = find_bar_by_id(project, bar_b_id)
    ta = win32com.client.Dispatch(bar_a.Tasks(1))
    tb = win32com.client.Dispatch(bar_b.Tasks(1))
    print(f"  A: Start={ta.Start}, End={ta.End}, Dur={ta.GetUserDuration().Hours}h, LinksOut={ta.LinksOut.Count}")
    print(f"  B: Start={tb.Start}, End={tb.End}, Dur={tb.GetUserDuration().Hours}h, LinksIn={tb.LinksIn.Count}")

    if tb.Start >= ta.End:
        print(f"\n  *** SUCCESS: B starts after A ends! ***")
    else:
        print(f"\n  B.Start={tb.Start} vs A.End={ta.End}")

    # Test link type change
    print(f"\n--- Change to SS ---")
    project.StartTransaction("SS")
    try:
        bar_a = find_bar_by_id(project, bar_a_id)
        ta = win32com.client.Dispatch(bar_a.Tasks(1))
        link = win32com.client.Dispatch(ta.LinksOut.Item(1))
        link.type = 1  # SS
        print(f"  type=1 (SS) => readback: {link.type}")
        project.EndTransaction()
        wait(project)
        project.Reschedule(pywintypes.Time(datetime.now()))
        wait(project)
        bar_a = find_bar_by_id(project, bar_a_id)
        bar_b = find_bar_by_id(project, bar_b_id)
        ta = win32com.client.Dispatch(bar_a.Tasks(1))
        tb = win32com.client.Dispatch(bar_b.Tasks(1))
        print(f"  A Start={ta.Start}, B Start={tb.Start}")
    except Exception as e:
        print(f"  Error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass

    # Set lag
    print(f"\n--- Set FS + 2d lag ---")
    project.StartTransaction("Lag")
    try:
        bar_a = find_bar_by_id(project, bar_a_id)
        ta = win32com.client.Dispatch(bar_a.Tasks(1))
        link = win32com.client.Dispatch(ta.LinksOut.Item(1))
        link.type = 0  # FS
        lag = ta.GetDurationFromString("2d")
        link.StartLagTime = lag
        print(f"  FS + 2d lag => readback: type={link.type}, lag={link.StartLagTime.Hours}h")
        project.EndTransaction()
        wait(project)
        project.Reschedule(pywintypes.Time(datetime.now()))
        wait(project)
        bar_a = find_bar_by_id(project, bar_a_id)
        bar_b = find_bar_by_id(project, bar_b_id)
        ta = win32com.client.Dispatch(bar_a.Tasks(1))
        tb = win32com.client.Dispatch(bar_b.Tasks(1))
        print(f"  A End={ta.End}, B Start={tb.Start}")
    except Exception as e:
        print(f"  Error: {e}")
        traceback.print_exc()
        try:
            project.AbandonTransaction()
        except Exception:
            pass


def cleanup(project, bar_ids, parent_index=3):
    """Remove test bars under parent_index."""
    if not bar_ids:
        return
    print(f"\n--- Cleanup ---")
    for bid in reversed(bar_ids):
        try:
            parent_task = get_summary_task(project, parent_index)
            project.StartTransaction(f"Del {bid}")
            cb = parent_task.ChildBars
            for i in range(cb.Count, 0, -1):
                b = win32com.client.Dispatch(cb.Item(i))
                if b.ID == bid:
                    cb.Remove(i)
                    print(f"  Removed bar ID={bid}")
                    break
            project.EndTransaction()
            wait(project)
        except Exception as e:
            print(f"  Error: {e}")
            try:
                project.AbandonTransaction()
            except Exception:
                pass


if __name__ == "__main__":
    print("COM Explorer v19 — Correct Architecture")
    print("=" * 80)
    created_ids = []
    try:
        app, project = connect()

        # Pre-cleanup under Insaat (index 3)
        print("\n--- Pre-cleanup ---")
        try:
            insaat_task = get_summary_task(project, 3)
            stale = []
            for j in range(1, insaat_task.ChildBars.Count + 1):
                sb = win32com.client.Dispatch(insaat_task.ChildBars.Item(j))
                if sb.Name.startswith(("V19_", "V18", "V17_", "V16_", "V15_", "V14_", "TEST_")):
                    stale.append(sb.ID)
                    print(f"  Stale: ID={sb.ID}, Name={sb.Name}")
            if stale:
                cleanup(project, stale, 3)
        except Exception as e:
            print(f"  Pre-cleanup error: {e}")

        # Test 1: Create bars
        bar_ids = test_create_under_insaat(project)
        created_ids = list(bar_ids)

        # Test 2: Duration on A
        if len(bar_ids) >= 1:
            test_duration(project, bar_ids[0])

        # Test 3: Dates on A
        if len(bar_ids) >= 1:
            test_dates(project, bar_ids[0])

        # Test 4: Link A -> B
        if len(bar_ids) >= 2:
            test_linking(project, bar_ids[0], bar_ids[1])

        # Cleanup
        cleanup(project, created_ids, 3)

        print("\n" + "=" * 80)
        print("ALL TESTS COMPLETE")
        print("=" * 80)
    except Exception as e:
        print(f"Fatal: {e}")
        traceback.print_exc()
        if created_ids:
            try:
                cleanup(project, created_ids, 3)
            except Exception:
                pass
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
