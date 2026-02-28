"""
COM Explorer v5 — Focus on SOLVING the critical issues:
1. Nested bar access via ChildBars COLLECTION (Count/Item)
2. IView full attribute dump for bar access
3. Date setting: constraint types, EditTokenV token names
4. Link creation: project.Links, project.LinkCategorys
5. bars.Remove with correct index parameter
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


def test_childbars_collection(project):
    """KEY TEST: ChildBars as a collection with Count/Item."""
    print("\n" + "="*80)
    print("TEST 1: CHILDBARS AS COLLECTION (Count/Item)")
    print("="*80)

    bars = project.Bars
    bar1 = bars.Item(1)
    etask = bar1.ExpandedTask
    print(f"Top bar: ID={bar1.ID}, Name={bar1.Name}")

    # Get ChildBars as property (NOT calling it)
    print("\n--- ChildBars property exploration ---")
    try:
        cb = etask.ChildBars
        print(f"  etask.ChildBars = {cb}")
        print(f"  type = {type(cb)}")

        # List ALL attributes on the ChildBars object
        attrs = [a for a in dir(cb) if not a.startswith('_')]
        print(f"  Attributes ({len(attrs)}): {attrs}")

        # Try Count
        try:
            count = cb.Count
            print(f"  cb.Count = {count}")
        except Exception as e:
            print(f"  cb.Count => {str(e)[:80]}")

        # Try Item
        for i in [0, 1, 2, 3]:
            try:
                item = cb.Item(i)
                if item:
                    print(f"  cb.Item({i}) => ID={item.ID}, Name={item.Name[:50]}")
                else:
                    print(f"  cb.Item({i}) => None")
            except Exception as e:
                print(f"  cb.Item({i}) => {str(e)[:80]}")

        # Try direct indexing
        for i in [0, 1, 2]:
            try:
                item = cb(i)
                if item:
                    print(f"  cb({i}) => ID={item.ID}, Name={item.Name[:50]}")
            except Exception as e:
                print(f"  cb({i}) => {str(e)[:60]}")

        # Try iteration
        try:
            print("\n  --- Iterating ChildBars ---")
            count = 0
            for child in cb:
                if count >= 10:
                    print(f"  ... (stopped at 10)")
                    break
                try:
                    print(f"  [{count}] ID={child.ID}, Name={child.Name[:50]}")
                except:
                    print(f"  [{count}] {child}")
                count += 1
            print(f"  Iterated: {count} children")
        except Exception as e:
            print(f"  Iteration error: {str(e)[:80]}")

        # Try _NewEnum
        try:
            enum = cb._NewEnum()
            print(f"  _NewEnum => {enum}")
        except Exception as e:
            print(f"  _NewEnum => {str(e)[:60]}")

    except Exception as e:
        print(f"  ChildBars error: {e}")

    # Also try ExpandedTask's other child-related properties
    print("\n--- Other child properties on ExpandedTask ---")
    child_props = ['Tasks', 'SubTasks', 'Children', 'ChildTasks',
                   'AllBars', 'Bars', 'SubBars', 'NestedBars',
                   'FirstChildBar', 'LastChildBar',
                   'Contents', 'Members', 'Items',
                   'ChildBar', 'GetChildBars', 'GetChildren',
                   'ChildCount', 'NumberOfChildren', 'HasChildren']
    for prop in child_props:
        try:
            val = getattr(etask, prop)
            if callable(val):
                try:
                    result = val()
                    print(f"  etask.{prop}() => {result}")
                except:
                    print(f"  etask.{prop}() => callable but failed")
            else:
                print(f"  etask.{prop} => {val}")
        except:
            pass  # Skip non-existent

    # Try bar-level child access
    print("\n--- Bar-level child access ---")
    bar_child_props = ['ChildBars', 'Children', 'SubBars', 'Contents',
                       'FirstChild', 'LastChild', 'FirstBar', 'LastBar',
                       'NextBar', 'PrevBar', 'PreviousBar',
                       'NextSibling', 'PrevSibling', 'FirstSibling',
                       'Parent', 'ParentBar', 'SummarisedBy']
    for prop in bar_child_props:
        try:
            val = getattr(bar1, prop)
            if callable(val):
                try:
                    result = val()
                    if result:
                        try:
                            print(f"  bar.{prop}() => ID={result.ID}, Name={result.Name[:40]}")
                        except:
                            print(f"  bar.{prop}() => {result}")
                except Exception as e:
                    print(f"  bar.{prop}() => {str(e)[:60]}")
            else:
                if val:
                    try:
                        print(f"  bar.{prop} => ID={val.ID}, Name={val.Name[:40]}")
                    except:
                        print(f"  bar.{prop} => {val}")
        except:
            pass  # Skip non-existent


def test_iview_full(project, app):
    """Full dump of IView attributes to find bar access methods."""
    print("\n" + "="*80)
    print("TEST 2: IVIEW FULL ATTRIBUTE DUMP")
    print("="*80)

    try:
        view = project.CurrentView
        print(f"View: {view.Name[:60] if hasattr(view, 'Name') else view}")

        attrs = [a for a in dir(view) if not a.startswith('_')]
        print(f"\nAll attributes ({len(attrs)}):")
        for attr in sorted(attrs):
            try:
                val = getattr(view, attr)
                if callable(val):
                    print(f"  {attr}() [callable]")
                else:
                    val_str = str(val)[:80]
                    print(f"  {attr} = {val_str}")
            except Exception as e:
                err = str(e)[:60]
                print(f"  {attr} => {err}")
    except Exception as e:
        print(f"View error: {e}")


def test_project_links(project):
    """Test project-level link access and LinkCategorys."""
    print("\n" + "="*80)
    print("TEST 3: PROJECT-LEVEL LINKS & LINK CATEGORYS")
    print("="*80)

    # LinkCategorys
    print("\n--- project.LinkCategorys ---")
    try:
        lc = project.LinkCategorys
        print(f"  LinkCategorys = {lc}")
        if lc:
            print(f"  type = {type(lc)}")
            try:
                count = lc.Count
                print(f"  Count = {count}")
                for i in range(1, min(count + 1, 6)):
                    try:
                        cat = lc.Item(i)
                        print(f"  [{i}] {cat.Name if hasattr(cat, 'Name') else cat}")
                        # Check for Add method
                        if i == 1:
                            attrs = [a for a in dir(cat) if not a.startswith('_')]
                            print(f"    Attributes: {attrs}")
                    except Exception as e:
                        print(f"  [{i}] => {str(e)[:60]}")
            except Exception as e:
                print(f"  Count error: {e}")
    except Exception as e:
        print(f"  Error: {e}")

    # AllLinkCategorys
    print("\n--- project.AllLinkCategorys ---")
    try:
        alc = project.AllLinkCategorys
        print(f"  AllLinkCategorys = {alc}")
        if alc:
            try:
                count = alc.Count
                print(f"  Count = {count}")
            except:
                pass
    except Exception as e:
        print(f"  Error: {e}")

    # Try project.Links
    print("\n--- project.Links ---")
    try:
        links = project.Links
        print(f"  Links = {links}")
    except Exception as e:
        print(f"  Error: {str(e)[:60]}")

    # Try project.AllLinks
    print("\n--- project.AllLinks ---")
    try:
        links = project.AllLinks
        print(f"  AllLinks = {links}")
    except Exception as e:
        print(f"  Error: {str(e)[:60]}")

    # Check project for AddLink, CreateLink methods
    print("\n--- Project link methods ---")
    for m in ['AddLink', 'CreateLink', 'LinkBars', 'Link', 'MakeLink',
              'AddDependency', 'CreateDependency', 'AddRelation']:
        if hasattr(project, m):
            print(f"  project.{m} EXISTS!")


def test_date_constraints(project):
    """Test different date constraint approaches."""
    print("\n" + "="*80)
    print("TEST 4: DATE CONSTRAINTS & SETTING")
    print("="*80)

    bars = project.Bars

    # First, list ALL ExpandedTask date-related properties
    print("\n--- ExpandedTask date-related attributes ---")
    bar1 = bars.Item(1)
    etask1 = bar1.ExpandedTask
    attrs = [a for a in dir(etask1) if not a.startswith('_')]
    date_attrs = [a for a in attrs if any(k in a.lower() for k in
                  ['date', 'start', 'end', 'finish', 'duration', 'constraint',
                   'imposed', 'early', 'late', 'actual', 'planned', 'baseline',
                   'user', 'schedule', 'calendar'])]
    print(f"  Date-related attrs ({len(date_attrs)}):")
    for attr in sorted(date_attrs):
        try:
            val = getattr(etask1, attr)
            if callable(val):
                try:
                    result = val()
                    print(f"    {attr}() = {result}")
                except:
                    print(f"    {attr}() [callable]")
            else:
                print(f"    {attr} = {val}")
        except Exception as e:
            print(f"    {attr} => {str(e)[:60]}")

    # Create a new bar for testing
    print("\n--- Creating test bar for date experiments ---")
    project.StartTransaction("Date Test v5")
    try:
        new_bar = bars.Add()
        new_bar.Name = "TEST_DATES_V5"
        bar_id = new_bar.ID
        etask = new_bar.ExpandedTask
        print(f"Created: ID={bar_id}")

        start = pywintypes.Time(datetime(2026, 6, 1))
        end = pywintypes.Time(datetime(2026, 6, 15))

        # Test all possible EditTokenV names
        print("\n--- EditTokenV token name exploration ---")
        tokens = ['Start', 'End', 'Finish', 'UserStart', 'UserEnd', 'UserFinish',
                  'PlannedStart', 'PlannedEnd', 'PlannedFinish',
                  'EarlyStart', 'EarlyEnd', 'EarlyFinish',
                  'LateStart', 'LateEnd', 'LateFinish',
                  'ActualStart', 'ActualEnd', 'ActualFinish',
                  'ConstraintDate', 'StartConstraintDate', 'EndConstraintDate',
                  'ImposedStart', 'ImposedEnd',
                  'ScheduleStart', 'ScheduleEnd',
                  'BaselineStart', 'BaselineEnd']
        for token in tokens:
            try:
                etask.EditTokenV(token, start)
                print(f"  EditTokenV('{token}') => SUCCESS")
            except Exception as e:
                err = str(e)[:60]
                if 'Unknown token' in err:
                    pass  # Skip
                else:
                    print(f"  EditTokenV('{token}') => {err}")

        # Check constraint type property
        print("\n--- Constraint properties ---")
        constraint_props = ['ConstraintType', 'StartConstraintType', 'EndConstraintType',
                           'ConstraintDate', 'StartConstraintDate', 'EndConstraintDate',
                           'Constraint', 'StartConstraint', 'EndConstraint']
        for prop in constraint_props:
            try:
                val = getattr(etask, prop)
                if callable(val):
                    val = val()
                print(f"  {prop} = {val}")
            except:
                pass

        # Try setting constraint type before date
        print("\n--- Set ConstraintType then date ---")
        # In Asta, constraint types might be: 0=ASAP, 1=ALAP, 2=MustStartOn, etc.
        for ct in range(8):
            try:
                etask.StartConstraintType = ct
                print(f"  StartConstraintType = {ct} => SUCCESS")
            except Exception as e:
                if ct < 3:
                    print(f"  StartConstraintType = {ct} => {str(e)[:60]}")

        # Now set ImposedStart and read back
        print("\n--- ImposedStart then read ---")
        etask.ImposedStart = start
        print(f"  Set ImposedStart = 2026-06-01")
        print(f"  ImposedStart readback = {etask.ImposedStart}")
        print(f"  etask.Start = {etask.Start}")
        print(f"  bar.Start = {new_bar.Start}")

        project.EndTransaction()
        try:
            project.WaitForNotificationProcessing()
        except:
            pass

        # After transaction
        print(f"\n--- After EndTransaction ---")
        print(f"  etask.Start = {etask.Start}")
        print(f"  ImposedStart = {etask.ImposedStart}")

        # Reschedule
        project.Reschedule(pywintypes.Time(datetime(2026, 2, 28)))
        try:
            project.WaitForNotificationProcessing()
        except:
            pass
        print(f"\n--- After Reschedule ---")
        print(f"  etask.Start = {etask.Start}")
        print(f"  etask.End = {etask.End}")
        print(f"  ImposedStart = {etask.ImposedStart}")

        # Clean up
        print("\n--- Cleanup (bars.Remove by index) ---")
        project.StartTransaction("Delete v5")
        # Find the bar by iterating
        for i in range(1, bars.Count + 1):
            try:
                b = bars.Item(i)
                if b.ID == bar_id:
                    bars.Remove(i)
                    print(f"  bars.Remove({i}) for ID={bar_id} => SUCCESS!")
                    break
            except Exception as e:
                print(f"  bars.Remove({i}) => {str(e)[:60]}")
        project.EndTransaction()

    except Exception as e:
        print(f"Error: {e}")
        traceback.print_exc()
        try:
            project.AbandonTransaction()
        except:
            pass


def test_nextbar_deep(project):
    """Test if NextBar on child bars traverses deeply."""
    print("\n" + "="*80)
    print("TEST 5: NEXTBAR DEEP TRAVERSAL")
    print("="*80)

    bars = project.Bars
    bar1 = bars.Item(1)
    etask1 = bar1.ExpandedTask
    print(f"Top bar: ID={bar1.ID}, Name={bar1.Name}")

    # Try getting first child via ChildBars
    print("\n--- ChildBars as collection ---")
    try:
        cb = etask1.ChildBars
        print(f"  ChildBars type: {type(cb)}")

        # Try to get Count property
        for prop in ['Count', 'Length', 'Size']:
            try:
                val = getattr(cb, prop)
                print(f"  {prop} = {val}")
            except:
                pass

        # If ChildBars returns an IBar directly (not a collection)
        # then it might BE the first child bar
        try:
            # Check if cb itself is a bar
            cb_id = cb.ID
            cb_name = cb.Name
            print(f"  ChildBars IS a bar! ID={cb_id}, Name={cb_name[:50]}")

            # Now try NextBar from this child
            print("\n  --- NextBar from first child ---")
            current = cb
            visited = set()
            count = 0
            while current and count < 50:
                try:
                    cid = current.ID
                    if cid in visited:
                        print(f"  Loop detected at ID={cid}")
                        break
                    visited.add(cid)
                    name = current.Name[:45] if current.Name else "?"

                    # Check if this bar has its own children
                    has_children = False
                    try:
                        sub_cb = current.ExpandedTask.ChildBars
                        sub_id = sub_cb.ID
                        if sub_id != cid:
                            has_children = True
                    except:
                        pass

                    ch_str = " [HAS CHILDREN]" if has_children else ""
                    print(f"  [{count+1}] ID={cid}, {name}{ch_str}")
                    count += 1

                    current = current.NextBar()
                except Exception as e:
                    print(f"  NextBar error: {str(e)[:60]}")
                    break
            print(f"  Total via NextBar from ChildBars: {count}")
        except Exception as e:
            print(f"  ChildBars is NOT a bar: {str(e)[:60]}")

    except Exception as e:
        print(f"  ChildBars error: {e}")

    # Also try bar1.NextBar() to see if it walks through children
    print("\n--- bar1.NextBar() walking ---")
    try:
        current = bar1.NextBar()
        if current:
            print(f"  NextBar after top = ID={current.ID}, Name={current.Name[:50]}")
            # Continue walking
            visited = {bar1.ID, current.ID}
            count = 1
            while count < 20:
                try:
                    current = current.NextBar()
                    if not current:
                        break
                    cid = current.ID
                    if cid in visited:
                        print(f"  Loop at ID={cid}")
                        break
                    visited.add(cid)
                    print(f"  [{count+1}] ID={cid}, Name={current.Name[:50]}")
                    count += 1
                except:
                    break
            print(f"  Total via NextBar from bar1: {count}")
        else:
            print("  NextBar returned None")
    except Exception as e:
        print(f"  Error: {e}")


def test_link_creation_from_existing(project):
    """Find an existing link and explore its creation pattern."""
    print("\n" + "="*80)
    print("TEST 6: LINK CREATION FROM EXISTING LINK PATTERNS")
    print("="*80)

    # We know ChildBars returns the first child as an IBar
    # Try to find bars with links by walking via ChildBars + NextBar
    bars = project.Bars
    bar1 = bars.Item(1)
    etask1 = bar1.ExpandedTask

    # Get first child
    try:
        cb = etask1.ChildBars
        current = cb
        visited = set()
        bars_with_links = []
        count = 0

        while current and count < 100:
            try:
                cid = current.ID
                if cid in visited:
                    break
                visited.add(cid)
                count += 1

                etask = current.ExpandedTask
                if etask:
                    try:
                        li = etask.LinksIn
                        lo = etask.LinksOut
                        li_count = li.Count if li else 0
                        lo_count = lo.Count if lo else 0
                        if li_count > 0 or lo_count > 0:
                            bars_with_links.append((current, etask, li_count, lo_count))
                            if len(bars_with_links) <= 3:
                                print(f"  Found: ID={cid}, {current.Name[:40]}, In={li_count}, Out={lo_count}")
                    except:
                        pass

                current = current.NextBar()
            except:
                break

        print(f"Scanned {count} bars, found {len(bars_with_links)} with links")

        if bars_with_links:
            bar, etask, li_count, lo_count = bars_with_links[0]
            print(f"\n--- Exploring links on: {bar.Name[:50]} ---")

            # Get a link object
            if lo_count > 0:
                lo = etask.LinksOut
                link = lo.Item(1)
                print(f"  LinksOut[1] type: {type(link)}")

                # Full attribute dump of link
                attrs = [a for a in dir(link) if not a.startswith('_')]
                print(f"  Link attributes ({len(attrs)}):")
                for attr in sorted(attrs):
                    try:
                        val = getattr(link, attr)
                        if callable(val):
                            print(f"    {attr}() [callable]")
                        else:
                            val_str = str(val)[:60]
                            print(f"    {attr} = {val_str}")
                    except Exception as e:
                        print(f"    {attr} => {str(e)[:50]}")

                # Full attribute dump of LinksOut collection
                print(f"\n  LinksOut collection attributes:")
                lo_attrs = [a for a in dir(lo) if not a.startswith('_')]
                print(f"    {lo_attrs}")

                # Check Add with detailed error
                print(f"\n  --- LinksOut.Add() detailed ---")
                try:
                    new_link = lo.Add()
                    print(f"  Add() => {new_link}")
                except Exception as e:
                    print(f"  Add() => {e}")
                    # Try with various params
                    if len(bars_with_links) > 1:
                        target_bar = bars_with_links[1][0]
                        target_etask = bars_with_links[1][1]
                        for desc, args in [
                            ("Add(target_bar)", (target_bar,)),
                            ("Add(target_etask)", (target_etask,)),
                            ("Add(target_bar.ID)", (target_bar.ID,)),
                            ("Add(target_bar, 0)", (target_bar, 0)),
                            ("Add(target_etask, 0)", (target_etask, 0)),
                            ("Add(target_bar, 0, 0)", (target_bar, 0, 0)),
                        ]:
                            try:
                                new_link = lo.Add(*args)
                                print(f"  {desc} => SUCCESS! {new_link}")
                            except Exception as e2:
                                print(f"  {desc} => {str(e2)[:80]}")

            if li_count > 0:
                li = etask.LinksIn
                print(f"\n  LinksIn collection attributes:")
                li_attrs = [a for a in dir(li) if not a.startswith('_')]
                print(f"    {li_attrs}")
    except Exception as e:
        print(f"Error: {e}")
        traceback.print_exc()


if __name__ == "__main__":
    print("Asta COM Explorer v5 — Solving Critical Issues")
    print("="*80)
    try:
        app, project = connect()
        test_childbars_collection(project)
        test_iview_full(project, app)
        test_nextbar_deep(project)
        test_project_links(project)
        test_date_constraints(project)
        test_link_creation_from_existing(project)
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
