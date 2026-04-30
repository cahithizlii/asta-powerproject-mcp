"""Phase 4 acceptance: full file MCP lifecycle including HERO 2800 assignments.

SAFETY: Operates on a temp XML file (no MS Project COM merge into user's
active project — file MCP semantics only act on matching open projects,
and our temp XML is never opened).

Scenario (target <30s wall clock):
  1. Build base from sample fixture: 200 villa tasks + 14 CAU resources
  2. HERO: bulk_add_assignments 200x14=2800 via Phase 4 file path <5s
  3. Read demo: read_tasks/links/resources/assignments/calendars/baselines/progress
  4. Query demo: ad-hoc filter expression
  5. Write demo: update_task duration -> verify
  6. save_as demo: project to a new path
  7. Cleanup: remove temp dir

Tool: msproject_file (8th MS Project MCP tool, 14 actions).
"""
import os
import shutil
import sys
import tempfile
import time

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from msproject_mcp_core import (  # noqa: E402
    _msp_file_add_tasks,
    _msp_file_add_resources,
    _msp_file_bulk_add_assignments,
    _msp_file_read_tasks,
    _msp_file_read_links,
    _msp_file_read_resources,
    _msp_file_read_assignments,
    _msp_file_read_calendars,
    _msp_file_read_baselines,
    _msp_file_read_progress,
    _msp_file_query,
    _msp_file_update_task,
    _msp_file_save_as,
)


SAMPLE_XML = os.path.join(REPO_ROOT, "tests", "fixtures", "sample_msp.xml")


def main():
    tmpdir = tempfile.mkdtemp(prefix="msp_phase4_")
    work_xml = os.path.join(tmpdir, "lifecycle.xml")
    shutil.copy(SAMPLE_XML, work_xml)
    print(f"[SAFE] working file: {work_xml}")

    try:
        t0 = time.time()

        # 1. Base build via file MCP
        print("\n1. Building 200 villa tasks + 14 CAU resources via file MCP...")
        task_items = [{"name": f"V{i:03d}", "duration": "2d"} for i in range(200)]
        r = _msp_file_add_tasks(file_path=work_xml, items=task_items)
        assert r["status"] == "ok"
        new_task_ids = r["task_ids"]
        cau = ["COW", "EXT", "STL", "CAR", "MSN", "DRW", "INW",
               "EWI", "CWI", "ACP", "ELW", "DMS", "MTR", "LBR"]
        r = _msp_file_add_resources(
            file_path=work_xml,
            items=[{"name": n, "type": "Work"} for n in cau],
        )
        assert r["status"] == "ok"
        new_res_ids = r["resource_ids"]
        print(f"   OK in {time.time()-t0:.2f}s")

        # 2. HERO
        print("\n2. HERO: bulk_add_assignments 2800 via Phase 4 file path...")
        items = [{"task_id": tid, "resource_id": rid}
                 for tid in new_task_ids for rid in new_res_ids]
        h0 = time.time()
        r = _msp_file_bulk_add_assignments(file_path=work_xml, items=items)
        h_elapsed = time.time() - h0
        assert r["status"] == "ok", f"hero failed: {r}"
        assert r["count"] == 2800
        assert h_elapsed < 5.0, f"HERO took {h_elapsed:.2f}s (target <5s)"
        print(f"   OK [HERO] {h_elapsed:.2f}s "
              f"(reported elapsed_s={r.get('elapsed_s')}, "
              f"auto_imported={r.get('auto_imported')})")

        # 3. Read demo
        print("\n3. Read demo via Phase 4 file MCP...")
        rt = _msp_file_read_tasks(file_path=work_xml)
        rl = _msp_file_read_links(file_path=work_xml)
        rs = _msp_file_read_resources(file_path=work_xml)
        ra = _msp_file_read_assignments(file_path=work_xml)
        rc = _msp_file_read_calendars(file_path=work_xml)
        rb = _msp_file_read_baselines(file_path=work_xml)
        rp = _msp_file_read_progress(file_path=work_xml)
        print(f"   tasks={rt['count']}, links={rl['count']}, "
              f"resources={rs['count']}, assignments={ra['count']}, "
              f"calendars={rc['count']}, baseline.tasks={len(rb.get('tasks', []))}, "
              f"progress.tasks={len(rp.get('tasks', []))}")
        assert ra["count"] >= 2800

        # 4. Query demo
        print("\n4. Query demo: tasks with duration_h >= 16...")
        q = _msp_file_query(file_path=work_xml, expression="duration_h >= 16")
        print(f"   matches={q['count']}")

        # 5. Write demo: update first new task duration
        print("\n5. Write demo: update_task duration of first villa task...")
        u = _msp_file_update_task(file_path=work_xml,
                                  task_id=new_task_ids[0],
                                  fields={"duration": "5d"})
        assert u["status"] == "ok"
        # Verify
        rt2 = _msp_file_read_tasks(file_path=work_xml)
        updated = next((t for t in rt2["tasks"] if t["id"] == new_task_ids[0]), None)
        assert updated is not None
        assert updated["duration_h"] == 40.0, (
            f"expected 40.0h, got {updated['duration_h']}h"
        )
        print(f"   OK auto_imported={u.get('auto_imported')}")

        # 6. save_as
        print("\n6. save_as demo: copy to a new path...")
        dst = os.path.join(tmpdir, "lifecycle_archive.xml")
        s = _msp_file_save_as(file_path=work_xml, output_path=dst)
        assert s["status"] == "ok"
        print(f"   OK output_path={dst}, size_bytes={s['size_bytes']}")

        elapsed = time.time() - t0
        print(f"\n[OK] ACCEPTANCE: {elapsed:.2f}s total (target <30s)")
        assert elapsed < 30.0, f"Too slow: {elapsed}s"

    finally:
        try:
            shutil.rmtree(tmpdir, ignore_errors=True)
        except Exception as e:
            print(f"[WARN] cleanup error: {e}")


if __name__ == "__main__":
    main()
