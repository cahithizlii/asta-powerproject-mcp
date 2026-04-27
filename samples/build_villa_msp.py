"""Build a 200-task villa project in active MS Project.
Acceptance test for Phase 1: must complete in <5 sec.
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
    proj = app.ActiveProject
    if proj is None:
        print("ERROR: No active project. Open MS Project with an empty project first.")
        sys.exit(1)
    print(f"Active project: {proj.Name} (current Tasks.Count: {proj.Tasks.Count})")

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


if __name__ == "__main__":
    main()
