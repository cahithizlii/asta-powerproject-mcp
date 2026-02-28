"""
COM Explorer v10c — Fix LinkTo by placing bars inside project summary
1. Use ChangeParentBar() to move bars inside project summary
2. Test ConvertToTask() for link compatibility
3. Try linking existing project bars
4. Find working date+duration approach
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


def test_existing_bars_links(project):
    """Check existing bars with links to understand their properties."""
    print("\n" + "=" * 80)
    print("TEST 1: EXISTING BARS WITH LINKS")
    print("=" * 80)

    view = project.CurrentView
    bcv = win32com.client.Dispatch(view)

    # Show hierarchy to reveal all bars
    try:
        bcv.ShowHierarchy()
        print("ShowHierarchy() => OK")
    except Exception:
        pass

    all_ids = bcv.AllBarIds()
    all_link_ids = bcv.AllLinkIds()
    print(f"Total bars: {len(all_ids)}")
    print(f"Total links: {len(all_link_ids)}")

    # Find bars with links
    bars_with_links = []
    bars = project.Bars

    # We need to iterate through ALL bars, not just top-level
    # Use the expanded task's NextTask to traverse
    print(f"\nTop-level bars.Count: {bars.Count}")
    for i in range(1, bars.Count + 1):
        b = bars.Item(i)
        et = b.ExpandedTask
        li = et.LinksIn.Count
        lo = et.LinksOut.Count
        hl = et.HierarchyLevel
        print(f"  [{i}] ID={b.ID}, Name={b.Name[:40]}, type={et.type}, "
              f"Level={hl}, LinksIn={li}, LinksOut={lo}")
        if li > 0 or lo > 0:
            bars_with_links.append((b.ID, b.Name, li, lo))

        # Also check child bars
        try:
            cb = et.ChildBars
            if cb and cb.Count > 0:
                print(f"       ChildBars.Count = {cb.Count}")
                for ci in range(1, min(cb.Count + 1, 4)):
                    try:
                        child = cb.Item(ci)
                        cet = child.ExpandedTask
                        cli = cet.LinksIn.Count
                        clo = cet.LinksOut.Count
                        print(f"       Child [{ci}] ID={child.ID}, Name={child.Name[:30]}, "
                              f"type={cet.type}, LinksIn={cli}, LinksOut={clo}")
                        if cli > 0 or clo > 0:
                            bars_with_links.append((child.ID, child.Name, cli, clo))
                    except Exception:
                        pass
        except Exception:
            pass

    print(f"\nBars with links: {len(bars_with_links)}")
    for bid, bname, li, lo in bars_with_links[:5]:
        print(f"  ID={bid}, Name={bname[:30]}, LinksIn={li}, LinksOut={lo}")

    return bars_with_links


def test_link_existing_bars(project, bars_with_links):
    """Try to link two existing bars that already have links."""
    print("\n" + "=" * 80)
    print("TEST 2: LINK BETWEEN EXISTING BARS")
    print("=" * 80)

    if len(bars_with_links) < 2:
        print("  Need at least 2 bars with links to test")
        return

    # Find two bars we can link
    # Pick bars that have links (proven linkable)
    bid1, bn1, _, _ = bars_with_links[0]
    bid2, bn2, _, _ = bars_with_links[1]

    # Find them - need to search recursively
    bar1 = find_bar_recursive(project, bid1)
    bar2 = find_bar_recursive(project, bid2)

    if not (bar1 and bar2):
        print(f"  Can't find bars {bid1}, {bid2}")
        return

    et1 = bar1.ExpandedTask
    et2 = bar2.ExpandedTask

    print(f"  Bar1: ID={bar1.ID}, Name={bar1.Name[:30]}, type={et1.type}, Level={et1.HierarchyLevel}")
    print(f"  Bar2: ID={bar2.ID}, Name={bar2.Name[:30]}, type={et2.type}, Level={et2.HierarchyLevel}")

    # Try to link them
    print(f"\n--- LinkTo between existing bars ---")
    project.StartTransaction("Link existing")
    try:
        link = et1.LinkTo(et2)
        print(f"  et1.LinkTo(et2) => {link}")
        if link:
            print(f"  SUCCESS! Link created: ID={link.ID}")
        project.EndTransaction()
        wait(project)
    except Exception as e:
        print(f"  et1.LinkTo(et2) ERROR: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass


def find_bar_recursive(project, target_id):
    """Find a bar by ID, including inside summaries via ChildBars."""
    bars = project.Bars
    for i in range(1, bars.Count + 1):
        try:
            b = bars.Item(i)
            if b.ID == target_id:
                return b
            # Check children
            result = _search_children(b, target_id)
            if result:
                return result
        except Exception:
            pass
    return None


def _search_children(bar, target_id, depth=0):
    if depth > 5:
        return None
    try:
        et = bar.ExpandedTask
        cb = et.ChildBars
        if cb and cb.Count > 0:
            for i in range(1, cb.Count + 1):
                try:
                    child = cb.Item(i)
                    if child.ID == target_id:
                        return child
                    result = _search_children(child, target_id, depth + 1)
                    if result:
                        return result
                except Exception:
                    pass
    except Exception:
        pass
    return None


def test_change_parent_and_link(project):
    """Create bars inside project summary, then link them."""
    print("\n" + "=" * 80)
    print("TEST 3: CREATE BARS INSIDE SUMMARY + LINK")
    print("=" * 80)

    bars = project.Bars

    # Find the project summary (first bar, typically)
    summary = bars.Item(1)
    summary_et = summary.ExpandedTask
    print(f"Summary bar: ID={summary.ID}, Name={summary.Name[:40]}")
    print(f"  type={summary_et.type}, Level={summary_et.HierarchyLevel}")

    created_ids = []

    # Create bars and move them inside the summary
    project.StartTransaction("Create + parent")
    b1 = bars.Add()
    b1.Name = "CHILD_A"
    b2 = bars.Add()
    b2.Name = "CHILD_B"
    created_ids = [b1.ID, b2.ID]
    print(f"  Created: A={b1.ID}, B={b2.ID}")

    # Try ChangeParentBar to move inside summary
    et1 = b1.ExpandedTask
    et2 = b2.ExpandedTask

    try:
        et1.ChangeParentBar(summary, True)
        print(f"  A.ChangeParentBar(summary) => OK")
    except Exception as e:
        print(f"  A.ChangeParentBar error: {e}")
        # Try with bar instead of summary
        try:
            et1.ChangeParentBar(summary, False)
            print(f"  A.ChangeParentBar(summary, False) => OK")
        except Exception as e2:
            print(f"  A.ChangeParentBar(summary, False) error: {e2}")

    try:
        et2.ChangeParentBar(summary, True)
        print(f"  B.ChangeParentBar(summary) => OK")
    except Exception as e:
        print(f"  B.ChangeParentBar error: {e}")

    project.EndTransaction()
    wait(project)

    # Verify parent
    b1_found = find_bar(project, created_ids[0])
    b2_found = find_bar(project, created_ids[1])
    if not b1_found:
        b1_found = find_bar_recursive(project, created_ids[0])
    if not b2_found:
        b2_found = find_bar_recursive(project, created_ids[1])

    if b1_found:
        et1 = b1_found.ExpandedTask
        print(f"\n  A: Level={et1.HierarchyLevel}, Parent={et1.Parentname}")
    if b2_found:
        et2 = b2_found.ExpandedTask
        print(f"  B: Level={et2.HierarchyLevel}, Parent={et2.Parentname}")

    if not (b1_found and b2_found):
        print("  Can't find bars after parent change!")
        return created_ids

    et1 = b1_found.ExpandedTask
    et2 = b2_found.ExpandedTask

    # Try LinkTo
    print(f"\n--- LinkTo after ChangeParentBar ---")
    project.StartTransaction("Link children")
    try:
        link = et1.LinkTo(et2)
        print(f"  A.LinkTo(B) => {link}")
        if link:
            print(f"  SUCCESS! Link ID={link.ID}")
            link_dyn = win32com.client.Dispatch(link)
            for attr in sorted([a for a in dir(link_dyn) if not a.startswith('_')]):
                try:
                    val = getattr(link_dyn, attr)
                    if not callable(val):
                        print(f"    {attr} = {str(val)[:50]}")
                except Exception:
                    pass
        project.EndTransaction()
        wait(project)
    except Exception as e:
        print(f"  LinkTo error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass

    # Check
    print(f"\n  After LinkTo:")
    print(f"    A.LinksOut: {et1.LinksOut.Count}")
    print(f"    B.LinksIn: {et2.LinksIn.Count}")

    # Also try setting duration on the child bars
    print(f"\n--- Set duration on child bars ---")
    project.StartTransaction("Set dur A")
    try:
        dur_obj = et1.GetDurationFromString("10d")
        et1.SetUserDuration(dur_obj)
        print(f"  A.SetUserDuration(10d) => OK")
    except Exception as e:
        print(f"  A.SetUserDuration error: {e}")
    project.EndTransaction()
    wait(project)

    project.Reschedule(pywintypes.Time(datetime.now()))
    wait(project)

    print(f"  A after duration+reschedule:")
    try:
        print(f"    Duration={et1.GetUserDuration().Hours}h")
    except Exception:
        pass
    print(f"    Start={et1.Start}, End={et1.End}")

    return created_ids


def test_add_child_directly(project):
    """Test adding a bar as a child using _com_add_child approach."""
    print("\n" + "=" * 80)
    print("TEST 4: ADD CHILD BAR USING ExpandedTask.ChildBars")
    print("=" * 80)

    bars = project.Bars
    summary = bars.Item(1)
    et_summary = summary.ExpandedTask

    print(f"Summary: ID={summary.ID}, Name={summary.Name[:40]}")
    cb = et_summary.ChildBars
    print(f"  ChildBars.Count = {cb.Count}")

    # Try adding to ChildBars
    project.StartTransaction("Add child")
    try:
        # Method 1: ChildBars.Add()
        child = cb.Add()
        child.Name = "DIRECT_CHILD"
        child_id = child.ID
        print(f"  ChildBars.Add() => ID={child_id}")
        et_child = child.ExpandedTask
        print(f"  Child Level={et_child.HierarchyLevel}, Parent={et_child.Parentname}")
        project.EndTransaction()
        wait(project)

        # Try setting dates on child
        project.StartTransaction("Set child dates")
        et_child.ImposedStart = pywintypes.Time(datetime(2026, 6, 1))
        dur_obj = et_child.GetDurationFromString("10d")
        et_child.SetUserDuration(dur_obj)
        print(f"  Set ImposedStart=2026-06-01 + Duration=10d")
        project.EndTransaction()
        wait(project)

        project.Reschedule(pywintypes.Time(datetime.now()))
        wait(project)

        print(f"  After dates+reschedule:")
        try:
            print(f"    Duration={et_child.GetUserDuration().Hours}h")
        except Exception:
            pass
        print(f"    Start={et_child.Start}")
        print(f"    End={et_child.End}")
        print(f"    Constraint={et_child.Constraint}")

        return child_id

    except Exception as e:
        print(f"  ChildBars.Add error: {e}")
        traceback.print_exc()
        try:
            project.AbandonTransaction()
        except Exception:
            pass
        return None


def test_convert_to_task_and_link(project):
    """Test ConvertToTask and then linking."""
    print("\n" + "=" * 80)
    print("TEST 5: ConvertToTask + LinkTo")
    print("=" * 80)

    bars = project.Bars
    created_ids = []

    project.StartTransaction("Create")
    b1 = bars.Add()
    b1.Name = "CONV_A"
    b2 = bars.Add()
    b2.Name = "CONV_B"
    created_ids = [b1.ID, b2.ID]
    project.EndTransaction()
    wait(project)

    b1 = find_bar(project, created_ids[0])
    b2 = find_bar(project, created_ids[1])
    if not (b1 and b2):
        print("  Can't find bars!")
        return created_ids

    et1 = b1.ExpandedTask
    et2 = b2.ExpandedTask

    # Try ConvertToTask
    print(f"  Bar1 type before: {et1.type}")
    print(f"  Bar2 type before: {et2.type}")

    project.StartTransaction("Convert")
    try:
        t1 = et1.ConvertToTask()
        print(f"  et1.ConvertToTask() => {t1}")
        print(f"    type: {type(t1)}")
    except Exception as e:
        print(f"  ConvertToTask error: {e}")
        t1 = None
    try:
        t2 = et2.ConvertToTask()
        print(f"  et2.ConvertToTask() => {t2}")
    except Exception as e:
        print(f"  ConvertToTask error: {e}")
        t2 = None
    project.EndTransaction()
    wait(project)

    # Try linking with converted tasks
    if t1 and t2:
        project.StartTransaction("Link converted")
        try:
            link = t1.LinkTo(t2)
            print(f"  t1.LinkTo(t2) => {link}")
            project.EndTransaction()
            wait(project)
        except Exception as e:
            print(f"  t1.LinkTo(t2) error: {e}")
            try:
                project.AbandonTransaction()
            except Exception:
                pass

    # Also try linking etask to converted task
    project.StartTransaction("Link mixed")
    try:
        link = et1.LinkTo(t2) if t2 else et1.LinkTo(et2)
        print(f"  et1.LinkTo(t2/et2) => {link}")
        project.EndTransaction()
        wait(project)
    except Exception as e:
        print(f"  et1.LinkTo(t2/et2) error: {e}")
        try:
            project.AbandonTransaction()
        except Exception:
            pass

    return created_ids


def cleanup(project, bar_ids):
    if not bar_ids:
        return
    bar_ids = [x for x in bar_ids if x is not None]
    if not bar_ids:
        return
    print(f"\n--- Cleanup: {len(bar_ids)} bars ---")
    project.StartTransaction("Cleanup")
    bars = project.Bars
    for tid in reversed(bar_ids):
        found = False
        for i in range(bars.Count, 0, -1):
            try:
                b = bars.Item(i)
                if b.ID == tid:
                    bars.Remove(i)
                    print(f"  Deleted ID={tid}")
                    found = True
                    break
            except Exception:
                pass
        if not found:
            # Try recursive search
            b = find_bar_recursive(project, tid)
            if b:
                # Can't remove from children directly, try removing from Bars
                print(f"  ID={tid} found in hierarchy but can't remove from top-level")
    project.EndTransaction()
    wait(project)


if __name__ == "__main__":
    print("COM Explorer v10c — ChangeParentBar + ConvertToTask + LinkTo")
    print("=" * 80)
    all_ids = []
    try:
        app, project = connect()

        # First, clean stale bars
        bars = project.Bars
        stale = []
        for i in range(1, bars.Count + 1):
            b = bars.Item(i)
            if b.Name.startswith(("WF_TEST", "LF_", "DUR_", "DATE_", "TOKEN_",
                                  "LINK_", "TYPE_", "TEST_", "CHILD_", "CONV_",
                                  "WORKFLOW", "DIRECT_")):
                stale.append(b.ID)
        if stale:
            print(f"Pre-cleanup: {len(stale)} stale bars")
            cleanup(project, stale)

        # Test 1: Check existing bars with links
        bars_with_links = test_existing_bars_links(project)

        # Test 2: Link existing bars
        if bars_with_links:
            test_link_existing_bars(project, bars_with_links)

        # Test 3: ChangeParentBar + LinkTo
        ids3 = test_change_parent_and_link(project)
        if ids3:
            all_ids.extend(ids3)

        # Test 4: Add child directly
        id4 = test_add_child_directly(project)
        if id4:
            all_ids.append(id4)

        # Test 5: ConvertToTask + LinkTo
        ids5 = test_convert_to_task_and_link(project)
        if ids5:
            all_ids.extend(ids5)

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
