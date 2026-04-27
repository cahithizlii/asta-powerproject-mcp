"""Build a 200-task villa project in a NEW MS Project workspace.
Acceptance test for Phase 1: must complete in <5 sec.

SAFETY: Creates a fresh empty project via FileNew. User's original projects untouched.
On completion, leaves the new test project visible in MS Project for verification.
User can close it without saving (Ctrl+W -> No) when done.
"""
import time
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from msproject_mcp_core import _msp_task_bulk_add, _connect_app


def main():
    # Verify connection
    app = _connect_app()
    print(f"Connected to MS Project {app.Version}")

    # Create a fresh isolated test project (does NOT touch user's existing projects)
    app.FileNew()
    proj = app.ActiveProject
    if proj is None:
        print("ERROR: FileNew failed to produce an active project.")
        sys.exit(1)
    print(f"Created test project: {proj.Name} (Tasks.Count: {proj.Tasks.Count})")

    # Build 200 villa tasks
    items = []
    for i in range(200):
        items.append({"name": f"Villa T{i+1:03d}", "duration": "1d"})

    print(f"Building {len(items)} tasks...")
    start = time.time()
    r = _msp_task_bulk_add(items=items)
    elapsed = time.time() - start

    print(f"  Tasks added: {r['count']} via {r['path']} in {elapsed:.2f}s")
    if elapsed >= 5.0:
        print(f"WARNING: Bulk took {elapsed:.2f}s (target: <5s)")
        sys.exit(1)
    else:
        print(f"OK Acceptance: {elapsed:.2f}s < 5s target ({(5.0 - elapsed) / 5.0 * 100:.0f}% headroom)")

    print(f"Final Tasks.Count: {proj.Tasks.Count}")
    print()
    print(f"NOTE: Test project '{proj.Name}' is open for verification.")
    print(f"      Your original projects are untouched.")
    print(f"      Close test project: File -> Close -> Don't Save")


if __name__ == "__main__":
    main()
