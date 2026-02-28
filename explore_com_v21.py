"""
COM Explorer v21 — AddTask with correct parameters!
DISCOVERY: AddTask(date start_date, variant duration_or_end_date)
  - start_date: pywintypes.Time
  - duration_or_end_date: variant (duration string or end date?)

WORKFLOW:
  1. parent_task.ChildBars.Add() → empty bar (Tasks.Count=0)
  2. new_bar.Tasks.AddTask(start_date, duration) → task in bar
  3. Tasks(1) should now work
  4. Set properties, link, etc.
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


def find_bar_by_id(project, bar_id):
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


def test_addtask_with_params(project):
    """Create bar + AddTask with start date and duration."""
    print("\n" + "=" * 80)
    print("TEST 1: ChildBars.Add() + AddTask(start, duration)")
    print("=" * 80)

    # Use Insaat (index 3) as parent
    root_task = get_root_task(project)
    insaat = win32com.client.Dispatch(root_task.ChildBars.Item(3))
    insaat_task = win32com.client.Dispatch(insaat.Tasks(1))
    print(f"Parent: {insaat.Name[:40]}, ChildBars={insaat_task.ChildBars.Count}")

    start_date = pywintypes.Time(datetime(2026, 7, 1))
    end_date = pywintypes.Time(datetime(2026, 7, 15))

    created_bar_ids = []

    # Method A: AddTask(start_date, end_date)
    print(f"\n--- Method A: AddTask(start, end_date) ---")
    project.StartTransaction("Add A")
    try:
        insaat_task = win32com.client.Dispatch(insaat.Tasks(1))
        new_bar = win32com.client.Dispatch(insaat_task.ChildBars.Add())
        new_bar.Name = "V21_A_enddate"
        bar_id_a = new_bar.ID
        print(f"  Bar created: ID={bar_id_a}, Tasks.Count={new_bar.Tasks.Count}")

        # AddTask with start + end date
        task = new_bar.Tasks.AddTask(start_date, end_date)
        if task:
            td = win32com.client.Dispatch(task)
            print(f"  AddTask => ID={td.ID}, Name={td.Name[:30]}")
            print(f"  type: {type(td).__name__}")
            print(f"  Tasks.Count now: {new_bar.Tasks.Count}")
        else:
            print(f"  AddTask returned None")

        project.EndTransaction()
        wait(project)
        created_bar_ids.append(bar_id_a)
    except Exception as e:
        print(f"  Error: {e}")
        traceback.print_exc()
        try:
            project.AbandonTransaction()
        except Exception:
            pass

    # Check Tasks(1) after commit
    if created_bar_ids:
        print(f"\n--- After commit: check bar A ---")
        bar_a = find_bar_by_id(project, bar_id_a)
        if bar_a:
            try:
                ta = win32com.client.Dispatch(bar_a.Tasks(1))
                print(f"  Tasks(1): ID={ta.ID}, type={type(ta).__name__}")
                print(f"  Start={ta.Start}, End={ta.End}")
                print(f"  Duration={ta.GetUserDuration().Hours}h")
                print(f"  Constraint={ta.Constraint}")
                print(f"  LinksIn={ta.LinksIn.Count}, LinksOut={ta.LinksOut.Count}")
                if hasattr(ta, 'ImposedStart'):
                    print(f"  HAS ImposedStart")
                if hasattr(ta, 'LinkTo'):
                    print(f"  HAS LinkTo")
            except Exception as e:
                print(f"  Tasks(1) FAILED: {str(e)[:60]}")
        else:
            print(f"  Bar {bar_id_a} not found!")

    # Method B: AddTask(start_date, "10d") — duration string
    print(f"\n--- Method B: AddTask(start, '10d') ---")
    project.StartTransaction("Add B")
    try:
        root_task = get_root_task(project)
        insaat = win32com.client.Dispatch(root_task.ChildBars.Item(3))
        insaat_task = win32com.client.Dispatch(insaat.Tasks(1))
        new_bar = win32com.client.Dispatch(insaat_task.ChildBars.Add())
        new_bar.Name = "V21_B_durstr"
        bar_id_b = new_bar.ID
        print(f"  Bar created: ID={bar_id_b}")

        task = new_bar.Tasks.AddTask(start_date, "10d")
        if task:
            td = win32com.client.Dispatch(task)
            print(f"  AddTask => ID={td.ID}, type={type(td).__name__}")
            print(f"  Tasks.Count: {new_bar.Tasks.Count}")
        project.EndTransaction()
        wait(project)
        created_bar_ids.append(bar_id_b)
    except Exception as e:
        print(f"  Error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass

    # Method C: AddExpandedTask(start_date)
    print(f"\n--- Method C: AddExpandedTask(start) ---")
    project.StartTransaction("Add C")
    try:
        root_task = get_root_task(project)
        insaat = win32com.client.Dispatch(root_task.ChildBars.Item(3))
        insaat_task = win32com.client.Dispatch(insaat.Tasks(1))
        new_bar = win32com.client.Dispatch(insaat_task.ChildBars.Add())
        new_bar.Name = "V21_C_expanded"
        bar_id_c = new_bar.ID
        print(f"  Bar created: ID={bar_id_c}")

        task = new_bar.Tasks.AddExpandedTask(start_date)
        if task:
            td = win32com.client.Dispatch(task)
            print(f"  AddExpandedTask => ID={td.ID}, type={type(td).__name__}")
            print(f"  Tasks.Count: {new_bar.Tasks.Count}")
        project.EndTransaction()
        wait(project)
        created_bar_ids.append(bar_id_c)
    except Exception as e:
        print(f"  Error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass

    # Method D: AddMilestone(start_date)
    print(f"\n--- Method D: AddMilestone(start) ---")
    project.StartTransaction("Add D")
    try:
        root_task = get_root_task(project)
        insaat = win32com.client.Dispatch(root_task.ChildBars.Item(3))
        insaat_task = win32com.client.Dispatch(insaat.Tasks(1))
        new_bar = win32com.client.Dispatch(insaat_task.ChildBars.Add())
        new_bar.Name = "V21_D_milestone"
        bar_id_d = new_bar.ID
        print(f"  Bar created: ID={bar_id_d}")

        task = new_bar.Tasks.AddMilestone(start_date)
        if task:
            td = win32com.client.Dispatch(task)
            print(f"  AddMilestone => ID={td.ID}, type={type(td).__name__}")
            print(f"  Tasks.Count: {new_bar.Tasks.Count}")
        project.EndTransaction()
        wait(project)
        created_bar_ids.append(bar_id_d)
    except Exception as e:
        print(f"  Error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass

    # Verify all created bars
    print(f"\n--- Verify all created bars ---")
    for bid in created_bar_ids:
        bar = find_bar_by_id(project, bid)
        if bar:
            print(f"\n  Bar ID={bid}, Name={bar.Name[:30]}")
            try:
                t = win32com.client.Dispatch(bar.Tasks(1))
                print(f"    Tasks(1): ID={t.ID}, type={type(t).__name__}")
                print(f"    Start={t.Start}, End={t.End}")
                print(f"    Duration={t.GetUserDuration().Hours}h")
            except Exception as e:
                print(f"    Tasks(1) FAILED: {str(e)[:60]}")
        else:
            print(f"\n  Bar {bid} NOT FOUND")

    return created_bar_ids


def test_link_between_created(project, bar_ids):
    """Link two bars created with AddTask."""
    print("\n" + "=" * 80)
    print("TEST 2: LINK BETWEEN CREATED BARS")
    print("=" * 80)

    if len(bar_ids) < 2:
        print("  Need 2+ bars")
        return

    # Use first two bars with working Tasks(1)
    working = []
    for bid in bar_ids:
        bar = find_bar_by_id(project, bid)
        if bar:
            try:
                t = win32com.client.Dispatch(bar.Tasks(1))
                working.append(bid)
            except Exception:
                pass

    if len(working) < 2:
        print("  Need 2+ bars with Tasks(1)")
        return

    bid_1, bid_2 = working[0], working[1]

    project.StartTransaction("Link")
    try:
        bar_1 = find_bar_by_id(project, bid_1)
        bar_2 = find_bar_by_id(project, bid_2)
        t1 = win32com.client.Dispatch(bar_1.Tasks(1))
        t2 = win32com.client.Dispatch(bar_2.Tasks(1))
        print(f"  T1: ID={t1.ID}, Name={bar_1.Name[:20]}")
        print(f"  T2: ID={t2.ID}, Name={bar_2.Name[:20]}")

        link = t1.LinkTo(t2)
        if link:
            ld = win32com.client.Dispatch(link)
            print(f"  LINK CREATED! ID={ld.ID}, type={ld.type}")
        project.EndTransaction()
        wait(project)
    except Exception as e:
        print(f"  Link error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass
        return

    project.Reschedule(pywintypes.Time(datetime.now()))
    wait(project)

    # Verify
    bar_1 = find_bar_by_id(project, bid_1)
    bar_2 = find_bar_by_id(project, bid_2)
    t1 = win32com.client.Dispatch(bar_1.Tasks(1))
    t2 = win32com.client.Dispatch(bar_2.Tasks(1))
    print(f"  T1: Start={t1.Start}, End={t1.End}, LinksOut={t1.LinksOut.Count}")
    print(f"  T2: Start={t2.Start}, End={t2.End}, LinksIn={t2.LinksIn.Count}")

    if t2.Start >= t1.End:
        print(f"  *** SUCCESS: T2 starts after T1 ends! ***")


def cleanup(project, bar_ids):
    """Remove test bars."""
    if not bar_ids:
        return
    print(f"\n--- Cleanup ---")
    for bid in reversed(bar_ids):
        try:
            root_task = get_root_task(project)
            insaat = win32com.client.Dispatch(root_task.ChildBars.Item(3))
            insaat_task = win32com.client.Dispatch(insaat.Tasks(1))
            project.StartTransaction(f"Del {bid}")
            cb = insaat_task.ChildBars
            for i in range(cb.Count, 0, -1):
                b = win32com.client.Dispatch(cb.Item(i))
                if b.ID == bid:
                    cb.Remove(i)
                    print(f"  Removed bar ID={bid}")
                    break
            project.EndTransaction()
            wait(project)
        except Exception as e:
            print(f"  Error removing {bid}: {e}")
            try:
                project.AbandonTransaction()
            except Exception:
                pass


if __name__ == "__main__":
    print("COM Explorer v21 — AddTask with Parameters")
    print("=" * 80)
    created_ids = []
    try:
        app, project = connect()

        # Pre-cleanup
        print("--- Pre-cleanup ---")
        root_task = get_root_task(project)
        insaat = win32com.client.Dispatch(root_task.ChildBars.Item(3))
        insaat_task = win32com.client.Dispatch(insaat.Tasks(1))
        stale = []
        for j in range(1, insaat_task.ChildBars.Count + 1):
            sb = win32com.client.Dispatch(insaat_task.ChildBars.Item(j))
            if sb.Name.startswith(("V19_", "V20_", "V21_")):
                stale.append(sb.ID)
                print(f"  Stale: ID={sb.ID}, Name={sb.Name}")
        if stale:
            cleanup(project, stale)

        # Test 1: AddTask with params
        created_ids = test_addtask_with_params(project)

        # Test 2: Link
        if len(created_ids) >= 2:
            test_link_between_created(project, created_ids)

        # Cleanup
        cleanup(project, created_ids)

        print("\n" + "=" * 80)
        print("ALL TESTS COMPLETE")
        print("=" * 80)
    except Exception as e:
        print(f"Fatal: {e}")
        traceback.print_exc()
        if created_ids:
            cleanup(project, created_ids)
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
