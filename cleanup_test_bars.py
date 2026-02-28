"""Quick cleanup: remove test bars (links first, then bars)."""
import pythoncom
import win32com.client

pythoncom.CoInitialize()
app = win32com.client.GetActiveObject("{A57A0000-0200-0000-B2C5-00C0DF438041}")
project = app.ActiveProject
print(f"Project: {project.Name}")

def wait():
    try: project.WaitForNotificationProcessing()
    except: pass

root_task = win32com.client.Dispatch(project.Bars.Item(1).Tasks(1))
insaat = win32com.client.Dispatch(root_task.ChildBars.Item(3))
insaat_task = win32com.client.Dispatch(insaat.Tasks(1))

# Find test bars
test_ids = []
for j in range(1, insaat_task.ChildBars.Count + 1):
    sb = win32com.client.Dispatch(insaat_task.ChildBars.Item(j))
    if sb.Name.startswith(("V19_", "V20_", "V21_")):
        test_ids.append(sb.ID)
        print(f"Test bar: ID={sb.ID}, Name={sb.Name}")

if not test_ids:
    print("No test bars found!")
else:
    # Remove links first
    print(f"\n--- Removing links ---")
    for bid in test_ids:
        insaat_task = win32com.client.Dispatch(insaat.Tasks(1))
        for j in range(1, insaat_task.ChildBars.Count + 1):
            sb = win32com.client.Dispatch(insaat_task.ChildBars.Item(j))
            if sb.ID == bid:
                try:
                    t = win32com.client.Dispatch(sb.Tasks(1))
                    while t.LinksOut.Count > 0:
                        project.StartTransaction("Del link")
                        t.LinksOut.Remove(1)
                        project.EndTransaction()
                        wait()
                        t = win32com.client.Dispatch(sb.Tasks(1))
                        print(f"  Removed link from {bid}")
                    while t.LinksIn.Count > 0:
                        project.StartTransaction("Del link in")
                        t.LinksIn.Remove(1)
                        project.EndTransaction()
                        wait()
                        t = win32com.client.Dispatch(sb.Tasks(1))
                        print(f"  Removed link to {bid}")
                except Exception as e:
                    print(f"  Link cleanup {bid}: {str(e)[:50]}")
                break

    # Remove bars
    print(f"\n--- Removing bars ---")
    for bid in reversed(test_ids):
        try:
            insaat_task = win32com.client.Dispatch(insaat.Tasks(1))
            project.StartTransaction(f"Del {bid}")
            cb = insaat_task.ChildBars
            for i in range(cb.Count, 0, -1):
                b = win32com.client.Dispatch(cb.Item(i))
                if b.ID == bid:
                    cb.Remove(i)
                    print(f"  Removed bar {bid}")
                    break
            project.EndTransaction()
            wait()
        except Exception as e:
            print(f"  Error {bid}: {str(e)[:60]}")
            try: project.AbandonTransaction()
            except: pass

# Also cleanup top-level bars
bars = project.Bars
for i in range(bars.Count, 0, -1):
    b = win32com.client.Dispatch(bars.Item(i))
    if b.Name.startswith(("V19_", "V20_", "V21_")):
        project.StartTransaction("Del top")
        bars.Remove(i)
        project.EndTransaction()
        wait()
        print(f"  Removed top-level bar {b.ID}")

pythoncom.CoUninitialize()
print("\nDone.")
