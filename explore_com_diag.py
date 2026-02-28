"""Quick diagnostic: check current project state and Tasks(1) availability."""
import pythoncom
import win32com.client

pythoncom.CoInitialize()
clsid = "{A57A0000-0200-0000-B2C5-00C0DF438041}"
app = win32com.client.GetActiveObject(clsid)
project = app.ActiveProject
print(f"Project: {project.Name}")

# Root bar
root_bar = project.Bars.Item(1)
print(f"\nRoot bar: ID={root_bar.ID}, Name={root_bar.Name[:40]}")

# Try Tasks(1) on root
try:
    root_task = win32com.client.Dispatch(root_bar.Tasks(1))
    print(f"  Tasks(1): ID={root_task.ID}, type={type(root_task).__name__}")
except Exception as e:
    print(f"  Tasks(1) FAILED: {e}")
    # Try ExpandedTask
    try:
        et = root_bar.ExpandedTask
        print(f"  ExpandedTask: ID={et.ID}, type={type(et).__name__}")
    except Exception as e2:
        print(f"  ExpandedTask FAILED: {e2}")

# Try to list children via ExpandedTask.ChildBars
print(f"\n--- ExpandedTask.ChildBars ---")
try:
    et = root_bar.ExpandedTask
    cb = et.ChildBars
    print(f"  Count: {cb.Count}")
    for i in range(1, cb.Count + 1):
        b = win32com.client.Dispatch(cb.Item(i))
        print(f"  [{i}] ID={b.ID}, Name={b.Name[:40]}")
except Exception as e:
    print(f"  Error: {e}")

# Try Tasks(1).ChildBars
print(f"\n--- Tasks(1).ChildBars ---")
try:
    root_task = win32com.client.Dispatch(root_bar.Tasks(1))
    cb = root_task.ChildBars
    print(f"  Count: {cb.Count}")
    for i in range(1, cb.Count + 1):
        b = win32com.client.Dispatch(cb.Item(i))
        print(f"  [{i}] ID={b.ID}, Name={b.Name[:40]}")
        # Test Tasks(1) on each child
        try:
            ct = win32com.client.Dispatch(b.Tasks(1))
            print(f"       Tasks(1): ID={ct.ID}, type={type(ct).__name__}")
        except Exception as e:
            print(f"       Tasks(1) FAILED: {str(e)[:60]}")
        try:
            et2 = b.ExpandedTask
            print(f"       ExpandedTask: ID={et2.ID}")
        except Exception as e:
            print(f"       ExpandedTask FAILED: {str(e)[:60]}")
except Exception as e:
    print(f"  Tasks(1) on root FAILED: {e}")

# Count total bars
print(f"\n--- IBarChartView ---")
try:
    bcv = win32com.client.Dispatch(project.CurrentView)
    all_ids = bcv.AllBarIds()
    all_links = bcv.AllLinkIds()
    print(f"  AllBarIds: {len(all_ids)}")
    print(f"  AllLinkIds: {len(all_links)}")
except Exception as e:
    print(f"  Error: {e}")

# Check Bars collection
print(f"\n--- project.Bars ---")
try:
    bars = project.Bars
    print(f"  Bars.Count: {bars.Count}")
    for i in range(1, min(bars.Count + 1, 5)):
        b = win32com.client.Dispatch(bars.Item(i))
        print(f"  [{i}] ID={b.ID}, Name={b.Name[:40]}")
except Exception as e:
    print(f"  Error: {e}")

pythoncom.CoUninitialize()
print("\nDone.")
