"""
COM Explorer v10b — Refined tests:
1. ImposedStart + ImposedEnd in SAME transaction
2. Task type investigation
3. LinkTo test
4. Bar.EditTokenV for duration
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


def find_bar(project, bar_id):
    bars = project.Bars
    for i in range(1, bars.Count + 1):
        try:
            b = bars.Item(i)
            if b.ID == bar_id:
                return b
        except Exception:
            pass
    return None


def wait(project):
    try:
        project.WaitForNotificationProcessing()
    except Exception:
        pass


def test_dates_together(project):
    """Test setting ImposedStart AND ImposedEnd in same transaction."""
    print("\n" + "=" * 80)
    print("TEST 1: ImposedStart + ImposedEnd TOGETHER")
    print("=" * 80)

    # Create bar
    project.StartTransaction("Create bar")
    b = project.Bars.Add()
    b.Name = "DATE_TOGETHER"
    bar_id = b.ID
    et = b.ExpandedTask

    # Check task type
    print(f"Created bar ID={bar_id}")
    print(f"  type = {et.type}")
    print(f"  Constraint = {et.Constraint}")

    # Try setting type to 0 (Task) in same transaction
    try:
        et.type = 0
        print(f"  type set to 0 => now: {et.type}")
    except Exception as e:
        print(f"  type set error: {e}")

    project.EndTransaction()
    wait(project)

    # Re-fetch
    bar = find_bar(project, bar_id)
    if not bar:
        print("  Can't find bar!")
        return None
    et = bar.ExpandedTask
    print(f"\n  After creation:")
    print(f"    type = {et.type}")
    print(f"    Start = {et.Start}")
    print(f"    End = {et.End}")
    try:
        print(f"    Duration = {et.GetUserDuration().Hours}h")
    except Exception:
        pass
    print(f"    Constraint = {et.Constraint}")

    # Set BOTH dates in ONE transaction
    print(f"\n--- Set ImposedStart + ImposedEnd together ---")
    project.StartTransaction("Set both dates")
    try:
        # Remove existing constraint first
        et.RemoveConstraint()
        print(f"  RemoveConstraint => Constraint={et.Constraint}")

        # Set ImposedStart
        et.ImposedStart = pywintypes.Time(datetime(2026, 6, 1))
        print(f"  ImposedStart = 2026-06-01 => readback: {et.ImposedStart}")

        # Set ImposedEnd
        et.ImposedEnd = pywintypes.Time(datetime(2026, 6, 15))
        print(f"  ImposedEnd = 2026-06-15 => readback: {et.ImposedEnd}")

        project.EndTransaction()
        wait(project)
    except Exception as e:
        print(f"  Error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass

    # Reschedule
    project.Reschedule(pywintypes.Time(datetime.now()))
    wait(project)

    print(f"\n  After both dates + Reschedule:")
    print(f"    Start = {et.Start}")
    print(f"    End = {et.End}")
    try:
        print(f"    Duration = {et.GetUserDuration().Hours}h")
    except Exception:
        pass
    print(f"    Constraint = {et.Constraint}")
    print(f"    ImposedStart = {et.ImposedStart}")
    print(f"    ImposedEnd = {et.ImposedEnd}")

    # Test: Set them in SEPARATE transactions (Start first, then End)
    print(f"\n--- Set separately: Start first, then End ---")
    project.StartTransaction("Clear")
    et.RemoveConstraint()
    project.EndTransaction()
    wait(project)

    project.StartTransaction("Set start")
    et.ImposedStart = pywintypes.Time(datetime(2026, 7, 1))
    print(f"  ImposedStart = 2026-07-01")
    project.EndTransaction()
    wait(project)

    project.StartTransaction("Set end")
    et.ImposedEnd = pywintypes.Time(datetime(2026, 7, 15))
    print(f"  ImposedEnd = 2026-07-15")
    project.EndTransaction()
    wait(project)

    project.Reschedule(pywintypes.Time(datetime.now()))
    wait(project)

    print(f"  After separate dates + Reschedule:")
    print(f"    Start = {et.Start}")
    print(f"    End = {et.End}")
    try:
        print(f"    Duration = {et.GetUserDuration().Hours}h")
    except Exception:
        pass
    print(f"    Constraint = {et.Constraint}")
    print(f"    ImposedStart = {et.ImposedStart}")
    print(f"    ImposedEnd = {et.ImposedEnd}")

    return bar_id


def test_bar_edittoken_dates(project):
    """Test bar-level EditTokenV for dates and duration."""
    print("\n" + "=" * 80)
    print("TEST 2: BAR EditTokenV FOR DATES")
    print("=" * 80)

    project.StartTransaction("Create bar")
    b = project.Bars.Add()
    b.Name = "TOKEN_DATES"
    bar_id = b.ID
    project.EndTransaction()
    wait(project)

    bar = find_bar(project, bar_id)
    if not bar:
        print("  Can't find bar!")
        return None

    print(f"Created bar ID={bar_id}")
    print(f"  Start = {bar.Start}")
    print(f"  End = {bar.End}")
    try:
        print(f"  Duration = {bar.Duration}")
    except Exception:
        pass

    # Try bar-level EditTokenV (not etask-level)
    print(f"\n--- bar.EditTokenV tests ---")
    bar_dyn = win32com.client.Dispatch(bar)

    # First list all available tokens
    for token in ['Start', 'End', 'Finish', 'Duration', 'Name', 'Type',
                  'Constraint', 'Priority', 'Notes', 'WbnCode',
                  'ActualStart', 'ActualEnd', 'PercentComplete',
                  'ImposedStart', 'ImposedEnd', 'Calendar',
                  'EarlyStart', 'EarlyEnd', 'LateStart', 'LateEnd']:
        try:
            val = bar_dyn.GetToken(token)
            print(f"  GetToken('{token}') = {val}")
        except Exception as e:
            err = str(e)[:40]
            if 'Invalid' not in err and 'not found' not in err:
                print(f"  GetToken('{token}') => {err}")

    # Try setting via EditTokenV
    project.StartTransaction("EditTokenV dates")

    for token, value in [
        ('Start', '01/06/2026'),
        ('Start', '2026-06-01'),
        ('Start', '01 Jun 2026'),
        ('End', '15/06/2026'),
        ('End', '2026-06-15'),
        ('Finish', '15/06/2026'),
        ('Finish', '2026-06-15'),
        ('Duration', '10d'),
        ('Duration', '10 days'),
        ('Duration', '80h'),
    ]:
        try:
            bar_dyn.EditTokenV(token, value)
            print(f"  bar.EditTokenV('{token}', '{value}') => OK")
        except Exception as e:
            err = str(e)[:60]
            print(f"  bar.EditTokenV('{token}', '{value}') => {err}")

    project.EndTransaction()
    wait(project)

    project.Reschedule(pywintypes.Time(datetime.now()))
    wait(project)

    print(f"\n  After EditTokenV + Reschedule:")
    print(f"    Start = {bar.Start}")
    print(f"    End = {bar.End}")
    et = bar.ExpandedTask
    try:
        print(f"    Duration = {et.GetUserDuration().Hours}h")
    except Exception:
        pass

    return bar_id


def test_link_creation(project):
    """Test LinkTo between two simple bars."""
    print("\n" + "=" * 80)
    print("TEST 3: LINK CREATION via LinkTo/LinkFrom")
    print("=" * 80)

    # Create 2 bars
    project.StartTransaction("Create bars")
    b1 = project.Bars.Add()
    b1.Name = "LINK_A"
    b2 = project.Bars.Add()
    b2.Name = "LINK_B"
    ids = [b1.ID, b2.ID]
    project.EndTransaction()
    wait(project)

    bar1 = find_bar(project, ids[0])
    bar2 = find_bar(project, ids[1])
    if not (bar1 and bar2):
        print("  Can't find bars!")
        bars = project.Bars
        for i in range(1, bars.Count + 1):
            b = bars.Item(i)
            print(f"  [{i}] ID={b.ID}, Name={b.Name}")
        return ids

    et1 = bar1.ExpandedTask
    et2 = bar2.ExpandedTask

    print(f"  Bar1: ID={bar1.ID}, Name={bar1.Name}")
    print(f"  Bar2: ID={bar2.ID}, Name={bar2.Name}")
    print(f"  et1.LinksOut: {et1.LinksOut.Count}, et1.LinksIn: {et1.LinksIn.Count}")

    # LinkTo
    print(f"\n--- et1.LinkTo(et2) ---")
    project.StartTransaction("LinkTo")
    try:
        link = et1.LinkTo(et2)
        print(f"  Result: {link}")
        if link:
            print(f"  Link type: {type(link)}")
            try:
                link_dyn = win32com.client.Dispatch(link)
                for attr in sorted([a for a in dir(link_dyn) if not a.startswith('_')]):
                    try:
                        val = getattr(link_dyn, attr)
                        if not callable(val):
                            print(f"    {attr} = {str(val)[:50]}")
                        else:
                            print(f"    {attr}() [callable]")
                    except Exception as e:
                        print(f"    {attr} => {str(e)[:40]}")
            except Exception:
                pass
        project.EndTransaction()
        wait(project)
    except Exception as e:
        print(f"  ERROR: {e}")
        traceback.print_exc()
        try:
            project.AbandonTransaction()
        except Exception:
            pass

    # Verify
    print(f"\n  After LinkTo:")
    print(f"    et1.LinksOut: {et1.LinksOut.Count}")
    print(f"    et2.LinksIn: {et2.LinksIn.Count}")
    print(f"    et1.HasSuccessor: {et1.HasSuccessor}")
    print(f"    et2.HasPredecessor: {et2.HasPredecessor}")

    # If link created, dump link properties
    if et1.LinksOut.Count > 0:
        link_obj = et1.LinksOut.Item(1)
        link_dyn = win32com.client.Dispatch(link_obj)
        print(f"\n  Full ILink dump from LinksOut.Item(1):")
        for attr in sorted([a for a in dir(link_dyn) if not a.startswith('_')]):
            try:
                val = getattr(link_dyn, attr)
                if not callable(val):
                    print(f"    {attr} = {str(val)[:60]}")
                else:
                    print(f"    {attr}() [callable]")
            except Exception as e:
                print(f"    {attr} => {str(e)[:40]}")

    # Also test LinkFrom on a third bar
    print(f"\n--- LinkFrom test ---")
    project.StartTransaction("Create C")
    b3 = project.Bars.Add()
    b3.Name = "LINK_C"
    ids.append(b3.ID)
    project.EndTransaction()
    wait(project)

    bar3 = find_bar(project, ids[2])
    if bar3:
        et3 = bar3.ExpandedTask
        project.StartTransaction("LinkFrom")
        try:
            link2 = et3.LinkFrom(et2)
            print(f"  et3.LinkFrom(et2) => {link2}")
            if link2:
                print(f"  Link created! ID={link2.ID}")
            project.EndTransaction()
            wait(project)
        except Exception as e:
            print(f"  LinkFrom error: {e}")
            try:
                project.AbandonTransaction()
            except Exception:
                pass

        print(f"  et2.LinksOut: {et2.LinksOut.Count}")
        print(f"  et3.LinksIn: {et3.LinksIn.Count}")

    # Reschedule
    project.Reschedule(pywintypes.Time(datetime.now()))
    wait(project)

    return ids


def test_task_types(project):
    """Investigate task type enum values."""
    print("\n" + "=" * 80)
    print("TEST 4: TASK TYPE INVESTIGATION")
    print("=" * 80)

    # Check existing bars' types
    bars = project.Bars
    print(f"Existing bars and types:")
    for i in range(1, bars.Count + 1):
        b = bars.Item(i)
        et = b.ExpandedTask
        print(f"  [{i}] ID={b.ID}, Name={b.Name[:30]}, type={et.type}, Constraint={et.Constraint}")

    # Create bar and try different types
    print(f"\n--- Type enum exploration ---")
    project.StartTransaction("Type test")
    b = project.Bars.Add()
    b.Name = "TYPE_TEST"
    bar_id = b.ID
    et = b.ExpandedTask
    print(f"  New bar type = {et.type}")

    for type_val in range(10):
        try:
            et.type = type_val
            result = et.type
            dur_h = et.GetUserDuration().Hours if et.GetUserDuration() else "N/A"
            print(f"    type={type_val} => OK, readback={result}, Duration={dur_h}h")
            if result != type_val:
                print(f"      NOTE: readback differs from set value!")
        except Exception as e:
            print(f"    type={type_val} => {str(e)[:50]}")
    project.EndTransaction()
    wait(project)

    return bar_id


def cleanup(project, bar_ids):
    if not bar_ids:
        return
    print(f"\n--- Cleanup: {len(bar_ids)} bars ---")
    project.StartTransaction("Cleanup")
    bars = project.Bars
    for tid in reversed(bar_ids):
        for i in range(bars.Count, 0, -1):
            try:
                b = bars.Item(i)
                if b.ID == tid:
                    bars.Remove(i)
                    print(f"  Deleted ID={tid}")
                    break
            except Exception:
                pass
    project.EndTransaction()
    wait(project)
    print(f"  Done. bars.Count = {bars.Count}")


if __name__ == "__main__":
    print("COM Explorer v10b")
    print("=" * 80)
    all_ids = []
    try:
        app, project = connect()

        # First clean up any stale test bars
        print("\n--- Pre-cleanup stale bars ---")
        bars = project.Bars
        stale = []
        for i in range(1, bars.Count + 1):
            b = bars.Item(i)
            if b.Name.startswith(("WF_TEST", "LF_", "DUR_", "DATE_", "TOKEN_", "LINK_", "TYPE_", "TEST_")):
                stale.append(b.ID)
                print(f"  Found stale: ID={b.ID}, Name={b.Name}")
        if stale:
            cleanup(project, stale)

        id1 = test_dates_together(project)
        if id1:
            all_ids.append(id1)

        id2 = test_bar_edittoken_dates(project)
        if id2:
            all_ids.append(id2)

        ids3 = test_link_creation(project)
        if ids3:
            all_ids.extend(ids3)

        id4 = test_task_types(project)
        if id4:
            all_ids.append(id4)

        cleanup(project, all_ids)

        print("\n" + "=" * 80)
        print("ALL TESTS COMPLETE")
        print("=" * 80)
    except Exception as e:
        print(f"Fatal: {e}")
        traceback.print_exc()
        if all_ids:
            try:
                cleanup(project, all_ids)
            except Exception:
                pass
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
