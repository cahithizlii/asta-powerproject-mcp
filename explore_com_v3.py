"""
COM Explorer v3 — Uses ChildBars(index) for deep traversal.
Finds bars with links, tests link creation, tests EditTokenV for dates.
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

def get_child_bars_indexed(etask):
    """Get child bars via indexed ChildBars(i) — since it's not a collection."""
    children = []
    i = 1
    while True:
        try:
            child = etask.ChildBars(i)
            if child is None:
                break
            children.append(child)
            i += 1
            if i > 5000:  # safety limit
                break
        except:
            break
    return children

def deep_traverse(bars, max_depth=10, max_total=200):
    """Recursively traverse all bars using ChildBars(index)."""
    all_bars = []

    def _traverse(parent_bars_list, depth):
        if depth > max_depth or len(all_bars) >= max_total:
            return
        for bar in parent_bars_list:
            if len(all_bars) >= max_total:
                return
            all_bars.append(bar)
            try:
                etask = bar.ExpandedTask
                if etask:
                    children = get_child_bars_indexed(etask)
                    if children:
                        _traverse(children, depth + 1)
            except:
                pass

    # Get top-level bars
    top_bars = []
    try:
        count = bars.Count
        for i in range(1, count + 1):
            try:
                top_bars.append(bars.Item(i))
            except:
                pass
    except:
        pass

    _traverse(top_bars, 0)
    return all_bars


def test_deep_traversal(project):
    """Deep traversal and find bars with links."""
    print("\n" + "="*80)
    print("TEST 1: DEEP TRAVERSAL via ChildBars(index)")
    print("="*80)

    bars = project.Bars
    all_bars = deep_traverse(bars, max_total=100)
    print(f"Total bars found: {len(all_bars)}")

    # Show hierarchy
    print("\nFirst 30 bars:")
    for i, bar in enumerate(all_bars[:30]):
        try:
            name = bar.Name[:50] if bar.Name else "?"
            etask = bar.ExpandedTask
            children = get_child_bars_indexed(etask) if etask else []
            child_info = f" [{len(children)} children]" if children else ""

            # Check links
            li_count = 0
            lo_count = 0
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
            link_info = ""
            if li_count > 0 or lo_count > 0:
                link_info = f" [Links: In={li_count}, Out={lo_count}]"

            print(f"  [{i+1}] ID={bar.ID}, {name}{child_info}{link_info}")
        except Exception as e:
            print(f"  [{i+1}] error: {e}")

    return all_bars


def test_links_deep(project, all_bars):
    """Find bars with links and explore link objects."""
    print("\n" + "="*80)
    print("TEST 2: LINK EXPLORATION (via deep traversal)")
    print("="*80)

    # Find bars with links
    bars_with_links = []
    for bar in all_bars:
        try:
            etask = bar.ExpandedTask
            if etask:
                li = etask.LinksIn
                lo = etask.LinksOut
                li_count = li.Count if li else 0
                lo_count = lo.Count if lo else 0
                if li_count > 0 or lo_count > 0:
                    bars_with_links.append((bar, etask, li_count, lo_count))
        except:
            pass

    print(f"Bars with links: {len(bars_with_links)}")
    if not bars_with_links:
        print("No bars with links found!")
        return

    # Explore first linked bar
    bar, etask, li_count, lo_count = bars_with_links[0]
    print(f"\nExploring: ID={bar.ID}, Name={bar.Name}")
    print(f"  LinksIn={li_count}, LinksOut={lo_count}")

    # Explore LinksIn collection
    if li_count > 0:
        links_in = etask.LinksIn
        print(f"\n--- LinksIn collection ---")
        print(f"  Type: {type(links_in)}")
        print(f"  Non-underscore attributes:")
        for attr in sorted(dir(links_in)):
            if not attr.startswith('_'):
                print(f"    {attr}")

        # Get first link
        try:
            link = links_in.Item(1)
            print(f"\n--- Link object (LinksIn[1]) ---")
            print(f"  Type: {type(link)}")
            for attr in sorted(dir(link)):
                if not attr.startswith('_'):
                    try:
                        val = getattr(link, attr)
                        if callable(val):
                            print(f"    {attr}() [callable]")
                        else:
                            val_str = str(val)[:80]
                            print(f"    {attr} = {val_str}")
                    except Exception as e:
                        err = str(e)[:60]
                        print(f"    {attr} => {err}")

            # Try to read link properties
            print(f"\n--- Link properties ---")
            for prop in ['PredecessorTask', 'SuccessorTask', 'PredecessorBar', 'SuccessorBar',
                         'LinkType', 'Type', 'Lag', 'LagDuration', 'LinkCategory', 'Category',
                         'Critical', 'TotalFloat', 'FreeFloat', 'DrivingLink',
                         'PredecessorID', 'SuccessorID', 'ID', 'Name']:
                try:
                    val = getattr(link, prop)
                    if callable(val):
                        val = val()
                    print(f"    {prop} = {val}")
                except:
                    pass
        except Exception as e:
            print(f"  Link Item(1) error: {e}")

    # Explore LinksOut collection
    if lo_count > 0:
        links_out = etask.LinksOut
        print(f"\n--- LinksOut collection ---")
        print(f"  Type: {type(links_out)}")
        print(f"  Non-underscore attributes:")
        for attr in sorted(dir(links_out)):
            if not attr.startswith('_'):
                print(f"    {attr}")

        # Check for Add method
        print(f"\n  Has Add: {hasattr(links_out, 'Add')}")
        if hasattr(links_out, 'Add'):
            print(f"  Add type: {type(getattr(links_out, 'Add'))}")

    # Now test link creation
    print("\n--- Link creation attempts ---")
    # Find two adjacent bars without links between them
    if len(all_bars) >= 3:
        test_pred = all_bars[-3]
        test_succ = all_bars[-2]
        pred_etask = test_pred.ExpandedTask
        succ_etask = test_succ.ExpandedTask

        if pred_etask and succ_etask:
            print(f"  Pred: ID={test_pred.ID}, Name={test_pred.Name}")
            print(f"  Succ: ID={test_succ.ID}, Name={test_succ.Name}")

            # Try various link creation approaches
            # 1. LinksOut on pred
            try:
                lo = pred_etask.LinksOut
                print(f"\n  pred.LinksOut attributes: {[a for a in dir(lo) if not a.startswith('_')]}")

                # Try Add
                add_attempts = [
                    ("lo.Add()", lambda: lo.Add()),
                    ("lo.Add(succ_etask)", lambda: lo.Add(succ_etask)),
                    ("lo.Add(test_succ)", lambda: lo.Add(test_succ)),
                    ("lo.Add(test_succ.ID)", lambda: lo.Add(test_succ.ID)),
                    ("lo.Add(succ_etask, 0)", lambda: lo.Add(succ_etask, 0)),
                ]
                for desc, fn in add_attempts:
                    try:
                        result = fn()
                        print(f"  {desc} => SUCCESS! result={result}")
                    except Exception as e:
                        err = str(e)[:100]
                        print(f"  {desc} => {err}")
            except Exception as e:
                print(f"  pred.LinksOut error: {e}")

            # 2. LinksIn on succ
            try:
                li = succ_etask.LinksIn
                add_attempts = [
                    ("li.Add()", lambda: li.Add()),
                    ("li.Add(pred_etask)", lambda: li.Add(pred_etask)),
                    ("li.Add(test_pred)", lambda: li.Add(test_pred)),
                    ("li.Add(test_pred.ID)", lambda: li.Add(test_pred.ID)),
                ]
                for desc, fn in add_attempts:
                    try:
                        result = fn()
                        print(f"  {desc} => SUCCESS! result={result}")
                    except Exception as e:
                        err = str(e)[:100]
                        print(f"  {desc} => {err}")
            except Exception as e:
                print(f"  succ.LinksIn error: {e}")

            # 3. Project-level
            for m in ['AddLink', 'CreateLink', 'LinkBars', 'Link']:
                if hasattr(project, m):
                    print(f"  project.{m} EXISTS!")


def test_edittoken_dates(project):
    """Test EditTokenV for date and duration setting — the working approach."""
    print("\n" + "="*80)
    print("TEST 3: EditTokenV DATE/DURATION (working approach)")
    print("="*80)

    bars = project.Bars
    project.StartTransaction("Test EditTokenV Dates")
    try:
        new_bar = bars.Add()
        new_bar.Name = "TEST_EDITTOKEN_V"
        bar_id = new_bar.ID
        etask = new_bar.ExpandedTask
        print(f"Created test bar: ID={bar_id}")

        test_date = datetime(2026, 6, 1)
        ole_date = pywintypes.Time(test_date)
        end_date = pywintypes.Time(datetime(2026, 6, 15))

        # EditTokenV with datetime objects - known working
        print("\n--- EditTokenV (datetime objects) ---")
        for name, val in [
            ("Start", ole_date),
            ("End", end_date),
            ("Finish", end_date),
            ("Duration", ole_date),  # Probably wrong type but let's try
            ("ActualStart", ole_date),
            ("PlannedStart", ole_date),
            ("UserStart", ole_date),
            ("EarlyStart", ole_date),
        ]:
            try:
                etask.EditTokenV(name, val)
                print(f"  EditTokenV('{name}', datetime) => SUCCESS!")
            except Exception as e:
                err = str(e)[:80]
                print(f"  EditTokenV('{name}', datetime) => {err}")

        # Read back
        print("\n--- Read back after EditTokenV ---")
        try:
            print(f"  bar.Start = {new_bar.Start}")
            print(f"  bar.End = {new_bar.End}")
        except Exception as e:
            print(f"  Read error: {e}")
        try:
            print(f"  etask.GetUserStart() = {etask.GetUserStart()}")
            print(f"  etask.GetUserEnd() = {etask.GetUserEnd()}")
        except Exception as e:
            print(f"  GetUser error: {e}")
        try:
            print(f"  etask.Start = {etask.Start}")
            print(f"  etask.End = {etask.End}")
        except Exception as e:
            print(f"  etask date error: {e}")

        # ImposedStart/ImposedEnd - also known working
        print("\n--- ImposedStart/End (known working) ---")
        try:
            etask.ImposedStart = pywintypes.Time(datetime(2026, 7, 1))
            print(f"  ImposedStart = 2026-07-01 => SUCCESS!")
            print(f"  bar.Start after = {new_bar.Start}")
        except Exception as e:
            print(f"  ImposedStart error: {e}")

        try:
            etask.ImposedEnd = pywintypes.Time(datetime(2026, 7, 15))
            print(f"  ImposedEnd = 2026-07-15 => SUCCESS!")
            print(f"  bar.End after = {new_bar.End}")
        except Exception as e:
            print(f"  ImposedEnd error: {e}")

        # Try Duration via EditTokenV with different value types
        print("\n--- Duration via EditTokenV ---")
        dur_vals = [10, 10.0, "10", "10d", "10.0", pywintypes.Time(datetime(2026, 1, 11))]
        dur_descs = ["int 10", "float 10.0", "str '10'", "str '10d'", "str '10.0'", "Time(11 days)"]
        for val, desc in zip(dur_vals, dur_descs):
            try:
                etask.EditTokenV("Duration", val)
                print(f"  EditTokenV('Duration', {desc}) => SUCCESS!")
                try:
                    d = etask.Duration
                    print(f"    Duration now = {d}")
                except:
                    pass
            except Exception as e:
                err = str(e)[:80]
                print(f"  EditTokenV('Duration', {desc}) => {err}")

        # Try to remove the bar (find correct delete method)
        print("\n--- Bar deletion methods ---")
        for m in ['Delete', 'Remove', 'Destroy']:
            if hasattr(new_bar, m):
                print(f"  bar.{m} exists!")
            if hasattr(bars, m):
                print(f"  bars.{m} exists!")
        # Try bars.Remove
        for m in ['Remove', 'Delete']:
            try:
                getattr(bars, m)(new_bar)
                print(f"  bars.{m}(bar) => SUCCESS!")
                break
            except Exception as e:
                print(f"  bars.{m}(bar) => {str(e)[:60]}")
            try:
                getattr(bars, m)(bar_id)
                print(f"  bars.{m}(bar_id) => SUCCESS!")
                break
            except Exception as e:
                print(f"  bars.{m}(bar_id) => {str(e)[:60]}")

        project.EndTransaction()
        try:
            project.WaitForNotificationProcessing()
        except:
            pass
        print(f"\nTransaction ended.")
    except Exception as e:
        print(f"Error: {e}")
        traceback.print_exc()
        try:
            project.AbandonTransaction()
        except:
            pass


if __name__ == "__main__":
    print("Asta COM Explorer v3 — ChildBars(index) deep traversal")
    print("="*80)
    try:
        app, project = connect()
        all_bars = test_deep_traversal(project)
        test_links_deep(project, all_bars)
        test_edittoken_dates(project)
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
