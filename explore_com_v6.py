"""
COM Explorer v6 — TARGETED solutions:
1. ChildBars recursive traversal (IBars collection Count/Item)
2. Find bars with links -> test link creation
3. Duration setting (SetUserDuration, MoveToDate, ImposedStart+End)
4. bars.Remove by index
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


def recursive_childbars(etask, depth=0, max_depth=5, max_total=200, all_bars=None):
    """Recursively traverse ChildBars using IBars.Count/Item."""
    if all_bars is None:
        all_bars = []
    if depth > max_depth or len(all_bars) >= max_total:
        return all_bars

    try:
        cb = etask.ChildBars
        count = cb.Count
        for i in range(1, count + 1):
            if len(all_bars) >= max_total:
                break
            try:
                child = cb.Item(i)
                if child:
                    all_bars.append((child, depth + 1))
                    # Recurse into this child
                    child_etask = child.ExpandedTask
                    if child_etask:
                        recursive_childbars(child_etask, depth + 1, max_depth, max_total, all_bars)
            except Exception as e:
                if i <= 3:
                    print(f"{'  '*(depth+1)}Item({i}) error: {str(e)[:60]}")
    except:
        pass  # No children

    return all_bars


def test_recursive_childbars(project):
    """TEST: Full recursive traversal via ChildBars.Count/Item."""
    print("\n" + "="*80)
    print("TEST 1: RECURSIVE CHILDBARS TRAVERSAL")
    print("="*80)

    bars = project.Bars
    bar1 = bars.Item(1)
    etask1 = bar1.ExpandedTask
    print(f"Top bar: ID={bar1.ID}, Name={bar1.Name}")

    # Direct test of ChildBars collection
    print("\n--- Direct ChildBars test ---")
    try:
        cb = etask1.ChildBars
        print(f"  Type: {type(cb)}")
        count = cb.Count
        print(f"  Count: {count}")
        for i in range(1, min(count + 1, 6)):
            try:
                child = cb.Item(i)
                print(f"  Item({i}): ID={child.ID}, Name={child.Name[:50]}")
                # Check if this child has its own children
                try:
                    sub_cb = child.ExpandedTask.ChildBars
                    sub_count = sub_cb.Count
                    if sub_count > 0:
                        print(f"    -> Has {sub_count} children")
                        # Show first child of first child
                        if i == 1 and sub_count > 0:
                            sub_child = sub_cb.Item(1)
                            print(f"    -> First grandchild: ID={sub_child.ID}, Name={sub_child.Name[:50]}")
                except:
                    print(f"    -> No children (leaf)")
            except Exception as e:
                print(f"  Item({i}): {str(e)[:60]}")
    except Exception as e:
        print(f"  ChildBars error: {e}")

    # Full recursive traversal
    print("\n--- Full recursive traversal ---")
    all_bars = recursive_childbars(etask1, max_total=50)
    print(f"Total bars found: {len(all_bars)}")
    for bar_obj, depth in all_bars[:30]:
        try:
            indent = "  " * depth
            name = bar_obj.Name[:45] if bar_obj.Name else "?"
            # Check links
            li_count = lo_count = 0
            try:
                etask = bar_obj.ExpandedTask
                if etask:
                    try:
                        li = etask.LinksIn
                        li_count = li.Count if li else 0
                    except:
                        pass
                    try:
                        lo = etask.LinksOut
                        lo_count = lo.Count if lo else 0
                    except:
                        pass
            except:
                pass
            link_str = f" [In={li_count},Out={lo_count}]" if (li_count + lo_count) > 0 else ""
            print(f"  {indent}ID={bar_obj.ID}: {name}{link_str}")
        except:
            pass

    return all_bars


def test_link_creation(project, all_bars):
    """Find existing links, explore Add method on collections."""
    print("\n" + "="*80)
    print("TEST 2: LINK CREATION")
    print("="*80)

    # Find bars with links
    bars_with_links = []
    for bar_obj, depth in all_bars:
        try:
            etask = bar_obj.ExpandedTask
            if etask:
                li = etask.LinksIn
                lo = etask.LinksOut
                li_count = li.Count if li else 0
                lo_count = lo.Count if lo else 0
                if li_count > 0 or lo_count > 0:
                    bars_with_links.append((bar_obj, etask, li_count, lo_count))
        except:
            pass

    print(f"Bars with links: {len(bars_with_links)}")
    if not bars_with_links:
        print("No bars with links found in first 50 bars!")
        print("Trying deeper traversal...")
        # Try deeper
        bars = project.Bars
        bar1 = bars.Item(1)
        etask1 = bar1.ExpandedTask
        all_deep = recursive_childbars(etask1, max_total=500, max_depth=8)
        print(f"Deep traversal found {len(all_deep)} bars total")
        for bar_obj, depth in all_deep:
            try:
                etask = bar_obj.ExpandedTask
                if etask:
                    li = etask.LinksIn
                    lo = etask.LinksOut
                    li_count = li.Count if li else 0
                    lo_count = lo.Count if lo else 0
                    if li_count > 0 or lo_count > 0:
                        bars_with_links.append((bar_obj, etask, li_count, lo_count))
                        if len(bars_with_links) <= 5:
                            print(f"  Found: ID={bar_obj.ID}, {bar_obj.Name[:40]}, In={li_count}, Out={lo_count}")
            except:
                pass
        print(f"Total bars with links: {len(bars_with_links)}")

    if bars_with_links:
        bar, etask, li_count, lo_count = bars_with_links[0]
        print(f"\nExploring: ID={bar.ID}, {bar.Name[:50]}")

        if lo_count > 0:
            lo = etask.LinksOut
            print(f"\n--- LinksOut ---")
            print(f"  Count: {lo_count}")
            attrs = [a for a in dir(lo) if not a.startswith('_')]
            print(f"  Attrs: {attrs}")

            # Get existing link details
            link1 = lo.Item(1)
            print(f"\n  Link[1] attrs: {[a for a in dir(link1) if not a.startswith('_')]}")
            for prop in ['Task', 'PredecessorTask', 'SuccessorTask', 'PredecessorBar',
                         'SuccessorBar', 'Link_Category', 'LinkCategory', 'Category',
                         'Lag', 'LagDuration', 'StartLag', 'EndLag', 'Type', 'LinkType',
                         'Critical', 'ID', 'Name']:
                try:
                    val = getattr(link1, prop)
                    if callable(val):
                        val = val()
                    if hasattr(val, 'ID'):
                        print(f"    {prop} = ID:{val.ID} Name:{val.Name[:30]}")
                    elif hasattr(val, 'Name'):
                        print(f"    {prop} = {val.Name}")
                    else:
                        print(f"    {prop} = {val}")
                except:
                    pass

            # Try Add()
            print(f"\n--- LinksOut.Add() attempts ---")
            # Find a target bar (one without link to our bar)
            target = None
            for b, d in all_bars[-5:]:
                if b.ID != bar.ID:
                    target = b
                    break

            if target:
                target_etask = target.ExpandedTask
                print(f"  Target: ID={target.ID}, {target.Name[:40]}")

                attempts = [
                    ("lo.Add()", lambda: lo.Add()),
                    ("lo.Add(target)", lambda: lo.Add(target)),
                    ("lo.Add(target_etask)", lambda: lo.Add(target_etask)),
                    ("lo.Add(target.ID)", lambda: lo.Add(target.ID)),
                    ("lo.Add(target, 0)", lambda: lo.Add(target, 0)),
                    ("lo.Add(target_etask, 0)", lambda: lo.Add(target_etask, 0)),
                ]
                for desc, fn in attempts:
                    try:
                        result = fn()
                        print(f"  {desc} => SUCCESS! {result}")
                        # Clean up - don't leave test links
                        break
                    except Exception as e:
                        err = str(e)[:100]
                        print(f"  {desc} => {err}")

        if li_count > 0:
            li = etask.LinksIn
            print(f"\n--- LinksIn ---")
            attrs = [a for a in dir(li) if not a.startswith('_')]
            print(f"  Attrs: {attrs}")

            # Try Add
            print(f"\n--- LinksIn.Add() attempts ---")
            target = None
            for b, d in all_bars[-5:]:
                if b.ID != bar.ID:
                    target = b
                    break
            if target:
                target_etask = target.ExpandedTask
                print(f"  Target: ID={target.ID}, {target.Name[:40]}")
                attempts = [
                    ("li.Add()", lambda: li.Add()),
                    ("li.Add(target)", lambda: li.Add(target)),
                    ("li.Add(target_etask)", lambda: li.Add(target_etask)),
                    ("li.Add(target.ID)", lambda: li.Add(target.ID)),
                ]
                for desc, fn in attempts:
                    try:
                        result = fn()
                        print(f"  {desc} => SUCCESS! {result}")
                        break
                    except Exception as e:
                        err = str(e)[:100]
                        print(f"  {desc} => {err}")

    # Also explore project-level link creation methods
    print(f"\n--- Project attributes with 'link' ---")
    proj_attrs = [a for a in dir(project) if not a.startswith('_') and 'link' in a.lower()]
    for attr in proj_attrs:
        print(f"  {attr}")


def test_duration_and_dates(project):
    """Test all date/duration setting approaches with verification."""
    print("\n" + "="*80)
    print("TEST 3: DATE & DURATION SETTING WITH VERIFICATION")
    print("="*80)

    bars = project.Bars

    project.StartTransaction("Duration Test v6")
    try:
        new_bar = bars.Add()
        new_bar.Name = "TEST_DUR_V6"
        bar_id = new_bar.ID
        etask = new_bar.ExpandedTask
        print(f"Created: ID={bar_id}")

        start = pywintypes.Time(datetime(2026, 6, 1))
        end = pywintypes.Time(datetime(2026, 6, 15))

        # APPROACH 1: ImposedStart + ImposedEnd
        print("\n--- Approach 1: ImposedStart + ImposedEnd ---")
        etask.ImposedStart = start
        etask.ImposedEnd = end
        print(f"  Set ImposedStart=2026-06-01, ImposedEnd=2026-06-15")
        print(f"  Readback: ImposedStart={etask.ImposedStart}")
        print(f"  Readback: ImposedEnd={etask.ImposedEnd}")

        project.EndTransaction()
        try:
            project.WaitForNotificationProcessing()
        except:
            pass

        # Reschedule
        project.Reschedule(pywintypes.Time(datetime(2026, 2, 28)))
        try:
            project.WaitForNotificationProcessing()
        except:
            pass

        print(f"  After reschedule:")
        print(f"    etask.Start = {etask.Start}")
        print(f"    etask.End = {etask.End}")
        print(f"    GetUserStart = {etask.GetUserStart()}")
        print(f"    GetUserEnd = {etask.GetUserEnd()}")
        print(f"    Duration = {etask.Duration()}")
        print(f"    GetUserDuration = {etask.GetUserDuration()}")

        # APPROACH 2: SetUserStart/End/Duration
        print("\n--- Approach 2: SetUserStart, SetUserEnd, SetUserDuration ---")
        project.StartTransaction("SetUser test")

        new_start = pywintypes.Time(datetime(2026, 7, 1))
        new_end = pywintypes.Time(datetime(2026, 7, 20))

        # SetUserStart
        for desc, fn in [
            ("SetUserStart(datetime)", lambda: etask.SetUserStart(new_start)),
            ("SetUserStart(str)", lambda: etask.SetUserStart("2026-07-01")),
            ("SetUserStart(str2)", lambda: etask.SetUserStart("01/07/2026")),
        ]:
            try:
                fn()
                print(f"  {desc} => SUCCESS!")
                print(f"    Start now = {etask.Start}")
                break
            except Exception as e:
                print(f"  {desc} => {str(e)[:80]}")

        # SetUserEnd
        for desc, fn in [
            ("SetUserEnd(datetime)", lambda: etask.SetUserEnd(new_end)),
            ("SetUserEnd(str)", lambda: etask.SetUserEnd("2026-07-20")),
        ]:
            try:
                fn()
                print(f"  {desc} => SUCCESS!")
                print(f"    End now = {etask.End}")
                break
            except Exception as e:
                print(f"  {desc} => {str(e)[:80]}")

        # SetUserDuration - try different formats
        print("\n  --- SetUserDuration ---")
        for desc, val in [
            ("float 80.0 (10 days * 8 hours)", 80.0),
            ("int 80", 80),
            ("float 10.0", 10.0),
            ("pywintypes.Time", pywintypes.Time(datetime(2026, 1, 11))),
            ("str '10d'", "10d"),
            ("str '80'", "80"),
            ("VARIANT float", win32com.client.VARIANT(pythoncom.VT_R8, 80.0)),
            ("VARIANT int", win32com.client.VARIANT(pythoncom.VT_I4, 80)),
        ]:
            try:
                etask.SetUserDuration(val)
                dur = etask.GetUserDuration()
                print(f"  SetUserDuration({desc}) => SUCCESS! Duration={dur}")
                break
            except Exception as e:
                print(f"  SetUserDuration({desc}) => {str(e)[:80]}")

        project.EndTransaction()
        try:
            project.WaitForNotificationProcessing()
        except:
            pass

        # APPROACH 3: MoveToDate and MoveStart
        print("\n--- Approach 3: MoveToDate, MoveStart ---")
        project.StartTransaction("Move test")
        move_date = pywintypes.Time(datetime(2026, 8, 1))

        for desc, fn in [
            ("MoveToDate(date)", lambda: etask.MoveToDate(move_date)),
            ("MoveStart(date)", lambda: etask.MoveStart(move_date)),
            ("MoveToDate(date, True)", lambda: etask.MoveToDate(move_date, True)),
            ("MoveToDate(date, False)", lambda: etask.MoveToDate(move_date, False)),
        ]:
            try:
                fn()
                print(f"  {desc} => SUCCESS!")
                print(f"    Start now = {etask.Start}")
                break
            except Exception as e:
                print(f"  {desc} => {str(e)[:80]}")

        project.EndTransaction()
        try:
            project.WaitForNotificationProcessing()
        except:
            pass

        # APPROACH 4: AddConstraint
        print("\n--- Approach 4: AddConstraint ---")
        project.StartTransaction("Constraint test")
        con_date = pywintypes.Time(datetime(2026, 9, 1))

        for desc, fn in [
            ("AddConstraint(0, date)", lambda: etask.AddConstraint(0, con_date)),
            ("AddConstraint(1, date)", lambda: etask.AddConstraint(1, con_date)),
            ("AddConstraint(2, date)", lambda: etask.AddConstraint(2, con_date)),
            ("AddConstraint(3, date)", lambda: etask.AddConstraint(3, con_date)),
            ("AddConstraint(date)", lambda: etask.AddConstraint(con_date)),
        ]:
            try:
                fn()
                print(f"  {desc} => SUCCESS!")
                print(f"    Constraint = {etask.Constraint}")
                print(f"    StartConstraintDate = {etask.StartConstraintDate}")
                break
            except Exception as e:
                print(f"  {desc} => {str(e)[:80]}")

        project.EndTransaction()
        try:
            project.WaitForNotificationProcessing()
        except:
            pass

        # Cleanup
        print("\n--- Cleanup ---")
        project.StartTransaction("Delete v6")
        for i in range(1, bars.Count + 1):
            try:
                b = bars.Item(i)
                if b.ID == bar_id:
                    bars.Remove(i)
                    print(f"  Removed bar ID={bar_id} at index {i}")
                    break
            except:
                pass
        project.EndTransaction()

    except Exception as e:
        print(f"Error: {e}")
        traceback.print_exc()
        try:
            project.AbandonTransaction()
        except:
            pass


def test_bar_cleanup(project):
    """Clean up leftover test bars."""
    print("\n" + "="*80)
    print("TEST 4: CLEANUP LEFTOVER TEST BARS")
    print("="*80)

    bars = project.Bars
    print(f"Total top-level bars: {bars.Count}")

    test_bars = []
    for i in range(1, bars.Count + 1):
        try:
            b = bars.Item(i)
            if 'TEST' in (b.Name or ''):
                test_bars.append((i, b.ID, b.Name))
                print(f"  Test bar: index={i}, ID={b.ID}, Name={b.Name}")
        except:
            pass

    if test_bars:
        project.StartTransaction("Cleanup test bars")
        # Remove in reverse order (so indices don't shift)
        for idx, bid, name in reversed(test_bars):
            try:
                bars.Remove(idx)
                print(f"  Removed: index={idx}, ID={bid}, {name}")
            except Exception as e:
                print(f"  Remove index={idx} failed: {str(e)[:60]}")
        project.EndTransaction()
        try:
            project.WaitForNotificationProcessing()
        except:
            pass
        print(f"After cleanup: {bars.Count} top-level bars")
    else:
        print("No test bars to clean up")


if __name__ == "__main__":
    print("Asta COM Explorer v6 — Targeted Solutions")
    print("="*80)
    try:
        app, project = connect()
        # First clean up test bars from previous runs
        test_bar_cleanup(project)
        # Then run tests
        all_bars = test_recursive_childbars(project)
        test_link_creation(project, all_bars)
        test_duration_and_dates(project)
        print("\n" + "="*80)
        print("ALL TESTS COMPLETE")
        print("="*80)
    except Exception as e:
        print(f"Fatal error: {e}")
        traceback.print_exc()
    finally:
        try:
            pythoncom.CoUninitialize()
        except:
            pass
