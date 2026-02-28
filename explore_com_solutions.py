"""
Runtime COM exploration script for Asta Powerproject v2.
Tests solutions for link creation, date/duration, summary, ChildBars.
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

def test_childbars(project):
    """Test ChildBars traversal — including with parameters."""
    print("\n" + "="*80)
    print("TEST 1: CHILDBARS TRAVERSAL")
    print("="*80)

    bars = project.Bars
    top_count = bars.Count
    print(f"Top-level bars: {top_count}")

    bar1 = bars.Item(1)
    etask = bar1.ExpandedTask
    print(f"Bar 1: ID={bar1.ID}, Name={bar1.Name}")
    print(f"ExpandedTask: {etask}")

    if not etask:
        print("No ExpandedTask!")
        return

    # Try various ways to call ChildBars
    print("\n--- ChildBars calling variants ---")
    attempts = [
        ("etask.ChildBars", lambda: etask.ChildBars),
        ("etask.ChildBars()", lambda: etask.ChildBars()),
        ("etask.ChildBars(0)", lambda: etask.ChildBars(0)),
        ("etask.ChildBars(1)", lambda: etask.ChildBars(1)),
        ("etask.ChildBars(True)", lambda: etask.ChildBars(True)),
        ("etask.ChildBars(False)", lambda: etask.ChildBars(False)),
        ("etask.ChildBars(None)", lambda: etask.ChildBars(None)),
    ]
    for desc, fn in attempts:
        try:
            result = fn()
            print(f"  {desc} => {result}, type={type(result)}")
            if result is not None:
                try:
                    c = result.Count
                    print(f"    Count = {c}")
                    if c > 0:
                        for i in range(1, min(c + 1, 4)):
                            try:
                                cb = result.Item(i)
                                print(f"    Child {i}: ID={cb.ID}, Name={cb.Name}")
                            except Exception as e2:
                                print(f"    Child {i}: {e2}")
                except Exception as e2:
                    print(f"    No .Count: {e2}")
        except Exception as e:
            err = str(e)
            if len(err) > 100:
                err = err[:100]
            print(f"  {desc} => {err}")

    # Try NextBar traversal (walks display order, may include children)
    print("\n--- NextBar traversal (first 30 bars) ---")
    try:
        bar = bars.Item(1)
        visited = set()
        count = 0
        while bar and count < 30:
            try:
                bid = bar.ID
                if bid in visited:
                    break
                visited.add(bid)
                name = bar.Name[:50] if bar.Name else "?"
                # Check if it has ExpandedTask and if it's a summary
                etask_check = None
                has_children = "?"
                try:
                    etask_check = bar.ExpandedTask
                    if etask_check:
                        try:
                            # Try to get ChildBars count
                            cb = etask_check.ChildBars
                            has_children = f"ChildBars={cb}" if cb else "no"
                        except:
                            has_children = "err"
                except:
                    pass
                print(f"  [{count+1}] ID={bid}, Name={name}, children={has_children}")
                count += 1
                bar = bar.NextBar()
            except Exception as e:
                print(f"  NextBar error: {e}")
                break
        print(f"Total via NextBar: {count}")
    except Exception as e:
        print(f"NextBar traversal error: {e}")


def test_existing_links(project, app):
    """Find and explore existing link objects."""
    print("\n" + "="*80)
    print("TEST 2: EXISTING LINKS EXPLORATION")
    print("="*80)

    # Try AllLinkIds from view
    print("\n--- View.AllLinkIds() ---")
    try:
        view = app.ActiveView
        link_ids = view.AllLinkIds()
        if link_ids is not None:
            id_list = list(link_ids) if hasattr(link_ids, '__iter__') else [link_ids]
            print(f"Found {len(id_list)} link IDs")
            if len(id_list) > 0:
                print(f"First 10: {id_list[:10]}")
        else:
            print("AllLinkIds returned None")
    except Exception as e:
        print(f"AllLinkIds error: {e}")

    # Find bar with links via NextBar
    print("\n--- Finding bar with links ---")
    bars = project.Bars
    try:
        bar = bars.Item(1)
        visited = set()
        found_bar = None
        count = 0
        while bar and count < 200:
            try:
                bid = bar.ID
                if bid in visited:
                    break
                visited.add(bid)
                count += 1
                etask = bar.ExpandedTask
                if etask:
                    try:
                        li = etask.LinksIn
                        lo = etask.LinksOut
                        li_count = li.Count if li else 0
                        lo_count = lo.Count if lo else 0
                        if li_count > 0 or lo_count > 0:
                            print(f"  Bar {bid} ({bar.Name[:40]}): LinksIn={li_count}, LinksOut={lo_count}")
                            if not found_bar:
                                found_bar = (bar, etask, li_count, lo_count)
                    except:
                        pass
                bar = bar.NextBar()
            except:
                break
        print(f"Scanned {count} bars via NextBar")
    except Exception as e:
        print(f"Error: {e}")

    if not found_bar:
        print("No bar with links found!")
        return

    bar, etask, li_count, lo_count = found_bar
    print(f"\nExploring links on: ID={bar.ID}, Name={bar.Name}")

    # Explore a link object in detail
    links_coll = etask.LinksIn if li_count > 0 else etask.LinksOut
    link_name = "LinksIn" if li_count > 0 else "LinksOut"

    print(f"\n--- {link_name} collection ---")
    print(f"  Type: {type(links_coll)}")
    print(f"  Attributes: {[a for a in dir(links_coll) if not a.startswith('_')]}")

    # Check for Add method specifically
    print(f"\n  Has 'Add': {hasattr(links_coll, 'Add')}")

    # Get first link object
    try:
        link = links_coll.Item(1)
        print(f"\n--- Link object ({link_name}[1]) ---")
        print(f"  Type: {type(link)}")
        all_attrs = [a for a in dir(link) if not a.startswith('_')]
        print(f"  Attributes ({len(all_attrs)}):")
        for attr in sorted(all_attrs):
            try:
                val = getattr(link, attr)
                if callable(val):
                    print(f"    {attr}() [callable]")
                else:
                    val_str = str(val)
                    if len(val_str) > 80:
                        val_str = val_str[:80] + "..."
                    print(f"    {attr} = {val_str}")
            except Exception as e:
                err = str(e)[:60]
                print(f"    {attr} => error: {err}")
    except Exception as e:
        print(f"  Link Item(1) error: {e}")

    # Try LinksIn collection Add
    print(f"\n--- Trying {link_name}.Add() variants ---")
    # Find another bar to link to
    try:
        other_bar = bar.NextBar()
        if other_bar:
            other_etask = other_bar.ExpandedTask
            print(f"  Other bar: ID={other_bar.ID}, Name={other_bar.Name}")

            add_attempts = [
                (f"{link_name}.Add()", lambda: links_coll.Add()),
                (f"{link_name}.Add(other_etask)", lambda: links_coll.Add(other_etask)),
                (f"{link_name}.Add(other_bar)", lambda: links_coll.Add(other_bar)),
                (f"{link_name}.Add(other_bar.ID)", lambda: links_coll.Add(other_bar.ID)),
                (f"{link_name}.Add(other_etask, 0)", lambda: links_coll.Add(other_etask, 0)),
            ]
            for desc, fn in add_attempts:
                try:
                    result = fn()
                    print(f"  {desc} => SUCCESS! result={result}")
                except Exception as e:
                    err = str(e)[:100]
                    print(f"  {desc} => {err}")
    except Exception as e:
        print(f"  Other bar error: {e}")


def test_date_duration(project):
    """Test date/duration setting methods."""
    print("\n" + "="*80)
    print("TEST 3: DATE / DURATION SETTING")
    print("="*80)

    bars = project.Bars
    project.StartTransaction("Test Date Duration")
    try:
        new_bar = bars.Add()
        new_bar.Name = "TEST_DATE_DUR"
        bar_id = new_bar.ID
        etask = new_bar.ExpandedTask
        print(f"Created test bar: ID={bar_id}")

        # Test EditToken for dates
        test_date = datetime(2026, 6, 1)
        ole_date = pywintypes.Time(test_date)

        print("\n--- EditToken on ExpandedTask ---")
        token_tests = [
            ("Start", "01/06/2026"),
            ("Start", "2026-06-01"),
            ("Start", ole_date),
            ("End", "15/06/2026"),
            ("Finish", "15/06/2026"),
            ("Duration", "10d"),
            ("Duration", "10"),
            ("Duration", 10),
            ("UserStart", "01/06/2026"),
            ("UserEnd", "15/06/2026"),
        ]
        for name, val in token_tests:
            try:
                etask.EditToken(name, val)
                print(f"  EditToken('{name}', {repr(val)}) => SUCCESS!")
            except Exception as e:
                err = str(e)[:80]
                print(f"  EditToken('{name}', {repr(val)}) => {err}")

        print("\n--- EditTokenV on ExpandedTask ---")
        for name, val in [("Start", "01/06/2026"), ("Start", ole_date), ("Duration", "10d")]:
            try:
                etask.EditTokenV(name, val)
                print(f"  EditTokenV('{name}', {repr(val)}) => SUCCESS!")
            except Exception as e:
                err = str(e)[:80]
                print(f"  EditTokenV('{name}', {repr(val)}) => {err}")

        print("\n--- EditToken on IBar ---")
        for name, val in [("Start", "01/06/2026"), ("Start", ole_date), ("Duration", "10d")]:
            try:
                new_bar.EditToken(name, val)
                print(f"  bar.EditToken('{name}', {repr(val)}) => SUCCESS!")
            except Exception as e:
                err = str(e)[:80]
                print(f"  bar.EditToken('{name}', {repr(val)}) => {err}")

        print("\n--- EditTokenV on IBar ---")
        for name, val in [("Start", "01/06/2026"), ("Start", ole_date), ("Duration", "10d")]:
            try:
                new_bar.EditTokenV(name, val)
                print(f"  bar.EditTokenV('{name}', {repr(val)}) => SUCCESS!")
            except Exception as e:
                err = str(e)[:80]
                print(f"  bar.EditTokenV('{name}', {repr(val)}) => {err}")

        # Test SetUserDuration
        print("\n--- SetUserDuration on ExpandedTask ---")
        for val, desc in [(10, "int 10"), (10.0, "float 10.0"), ("10", "str '10'"), ("10d", "str '10d'"), (4800.0, "float 4800")]:
            try:
                etask.SetUserDuration(val)
                print(f"  SetUserDuration({desc}) => SUCCESS!")
                try:
                    d = etask.GetUserDuration()
                    print(f"    GetUserDuration = {d}")
                except:
                    pass
            except Exception as e:
                err = str(e)[:80]
                print(f"  SetUserDuration({desc}) => {err}")

        # Direct property setting
        print("\n--- Direct property set ---")
        for desc, fn in [
            ("bar.Start = ole_date", lambda: setattr(new_bar, 'Start', ole_date)),
            ("bar.End = ole_date", lambda: setattr(new_bar, 'End', pywintypes.Time(datetime(2026, 6, 15)))),
            ("etask.ImposedStart = ole_date", lambda: setattr(etask, 'ImposedStart', ole_date)),
            ("etask.ImposedEnd = ole_date", lambda: setattr(etask, 'ImposedEnd', pywintypes.Time(datetime(2026, 6, 15)))),
            ("etask.StartConstraintDate = ole_date", lambda: setattr(etask, 'StartConstraintDate', ole_date)),
        ]:
            try:
                fn()
                print(f"  {desc} => SUCCESS!")
            except Exception as e:
                err = str(e)[:80]
                print(f"  {desc} => {err}")

        # Read back dates
        print("\n--- Read back ---")
        try:
            print(f"  bar.Start = {new_bar.Start}")
            print(f"  bar.End = {new_bar.End}")
        except Exception as e:
            print(f"  Read error: {e}")

        # Delete test bar
        new_bar.Delete()
        project.EndTransaction()
        print(f"\nTest bar {bar_id} deleted.")

    except Exception as e:
        print(f"Error: {e}")
        traceback.print_exc()
        try:
            project.AbandonTransaction()
        except:
            pass


def test_summary_type(project):
    """Test task.type for summary conversion."""
    print("\n" + "="*80)
    print("TEST 4: SUMMARY / TYPE CONVERSION")
    print("="*80)

    bars = project.Bars
    project.StartTransaction("Test Summary Type")
    try:
        new_bar = bars.Add()
        new_bar.Name = "TEST_TYPE"
        bar_id = new_bar.ID
        etask = new_bar.ExpandedTask
        print(f"Created test bar: ID={bar_id}")

        # Read current type
        print("\n--- Current type ---")
        try:
            t = etask.type
            print(f"  type = {t} (python type: {type(t)})")
        except Exception as e:
            print(f"  type read error: {e}")

        # Try setting type
        print("\n--- Set type ---")
        for val in range(11):
            try:
                etask.type = val
                new_t = etask.type
                print(f"  type = {val} => SUCCESS! Now type = {new_t}")
            except Exception as e:
                err = str(e)[:80]
                print(f"  type = {val} => {err}")

        # DurationType
        print("\n--- DurationType ---")
        try:
            dt = etask.DurationType
            print(f"  Current: {dt}")
        except Exception as e:
            print(f"  Read error: {e}")
        for val in range(6):
            try:
                etask.DurationType = val
                print(f"  DurationType = {val} => SUCCESS! Now = {etask.DurationType}")
            except Exception as e:
                err = str(e)[:60]
                print(f"  DurationType = {val} => {err}")

        # ConvertToTask
        print("\n--- ConvertToTask ---")
        try:
            result = etask.ConvertToTask()
            print(f"  ConvertToTask() => {result}")
        except Exception as e:
            print(f"  ConvertToTask() => {e}")

        # Hierarchy info
        print("\n--- Hierarchy ---")
        for prop in ['HierarchyLevel', 'SummarisedBy', 'GetActualParentBar', 'Summary']:
            try:
                val = getattr(etask, prop)
                if callable(val):
                    val = val()
                print(f"  {prop} = {val}")
            except Exception as e:
                err = str(e)[:60]
                print(f"  {prop} => {err}")

        # Delete
        new_bar.Delete()
        project.EndTransaction()
        print(f"\nTest bar {bar_id} deleted.")
    except Exception as e:
        print(f"Error: {e}")
        traceback.print_exc()
        try:
            project.AbandonTransaction()
        except:
            pass


if __name__ == "__main__":
    print("Asta Powerproject COM Explorer v2")
    print("="*80)
    try:
        app, project = connect()
        test_childbars(project)
        test_existing_links(project, app)
        test_date_duration(project)
        test_summary_type(project)
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
