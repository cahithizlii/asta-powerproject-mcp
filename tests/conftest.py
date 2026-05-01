"""Pytest fixtures shared across MS Project tests."""
import pytest
import os

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


# ---------- Phase 5d: synthetic CAU-style XER fixture ----------

SAMPLE_CAU_XER_CONTENT = """ERMHDR\t18.8\t2026-05-01\tcahit\tProject Management\tUSD
%T\tPROJECT
%F\tproj_id\tproj_short_name\tplan_start_date\tplan_end_date\tlast_recalc_date
%R\t1\tCAU\t2024-07-08 08:00\t2028-06-20 17:00\t2026-05-01 17:00
%T\tCALENDAR
%F\tclndr_id\tclndr_name\tday_hr_cnt\tweek_hr_cnt
%R\t1\tCAU 6x9\t9.0\t54.0
%T\tRSRC
%F\trsrc_id\trsrc_name\trsrc_short_name\trsrc_type\tmax_qty_per_hr
%R\t101\tConcrete Workers\tCOW\tRT_Labor\t10.0
%R\t102\tExtractors\tEXT\tRT_Labor\t5.0
%R\t103\tSteel\tSTL\tRT_Mat\t100.0
%R\t104\tCarpenters\tCAR\tRT_Labor\t8.0
%T\tTASK
%F\ttask_id\twbs_id\tproj_id\tclndr_id\ttask_code\ttask_name\ttask_type\ttarget_drtn_hr_cnt\ttarget_start_date\ttarget_end_date\tact_start_date\tact_end_date\tphys_complete_pct\ttotal_float_hr_cnt\tcstr_type\tstatus_code
%R\t1001\t1\t1\t1\tA1010\tFoundation\tTT_Task\t180.0\t2024-07-08 08:00\t2024-07-29 17:00\t2024-07-08 08:00\t2024-07-29 17:00\t100.0\t0.0\tCS_ASAP\tTK_Complete
%R\t1002\t1\t1\t1\tA1020\tFrame\tTT_Task\t360.0\t2024-07-30 08:00\t2024-09-09 17:00\t2024-07-30 08:00\t\t75.0\t0.0\tCS_ASAP\tTK_Active
%R\t1003\t1\t1\t1\tA1030\tWalls\tTT_Task\t180.0\t2024-09-10 08:00\t2024-10-01 17:00\t\t\t0.0\t72.0\tCS_ASAP\tTK_NotStart
%R\t1004\t1\t1\t1\tA1040\tRoof\tTT_Task\t180.0\t2024-10-02 08:00\t2024-10-23 17:00\t\t\t0.0\t72.0\tCS_ASAP\tTK_NotStart
%R\t1005\t1\t1\t1\tA1050\tInterior\tTT_Task\t360.0\t2024-10-24 08:00\t2024-12-04 17:00\t\t\t0.0\t72.0\tCS_ASAP\tTK_NotStart
%R\t1006\t1\t1\t1\tA1060\tHandover\tTT_FinMile\t0.0\t2024-12-15 17:00\t2024-12-15 17:00\t\t\t0.0\t81.0\tCS_MFO\tTK_NotStart
%T\tTASKPRED
%F\ttask_pred_id\ttask_id\tpred_task_id\tpred_type\tlag_hr_cnt
%R\t1\t1002\t1001\tPR_FS\t0.0
%R\t2\t1003\t1002\tPR_FS\t0.0
%R\t3\t1004\t1003\tPR_FS\t0.0
%R\t4\t1005\t1004\tPR_FS\t0.0
%R\t5\t1006\t1005\tPR_FS\t0.0
%T\tTASKRSRC
%F\ttaskrsrc_id\ttask_id\trsrc_id\ttarget_qty\tact_reg_qty\ttarget_cost\tact_reg_cost
%R\t1\t1001\t101\t180.0\t180.0\t180.0\t180.0
%R\t2\t1001\t103\t1000.0\t1000.0\t1000.0\t1000.0
%R\t3\t1002\t101\t360.0\t270.0\t360.0\t270.0
%R\t4\t1002\t104\t180.0\t135.0\t180.0\t135.0
%R\t5\t1003\t101\t180.0\t0.0\t180.0\t0.0
%R\t6\t1004\t104\t180.0\t0.0\t180.0\t0.0
%R\t7\t1005\t102\t360.0\t0.0\t360.0\t0.0
%E
"""


@pytest.fixture
def sample_cau_xer(tmp_path):
    """Synthetic CAU-style XER (UTF-16-LE BOM) at tmp_path/sample_cau.xer.

    Models a 6-task hospital construction chain with 4 CAU resources, FS
    chain, 1 CAU 6x9 calendar, partial progress (Foundation 100%, Frame 75%).
    """
    path = tmp_path / "sample_cau.xer"
    with open(path, "wb") as f:
        f.write(b"\xff\xfe")  # UTF-16-LE BOM
        f.write(SAMPLE_CAU_XER_CONTENT.encode("utf-16-le"))
    return str(path)


@pytest.fixture(scope="session")
def fixtures_dir() -> str:
    return os.path.join(REPO_ROOT, "tests", "fixtures")


@pytest.fixture(scope="session")
def msproject_app():
    """Session-level MS Project COM connection. Skips if not available."""
    pythoncom = None
    try:
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        try:
            app = win32com.client.GetActiveObject("MSProject.Application")
        except Exception as com_err:
            # COM-specific failure (MS Project not running, etc.)
            pytest.skip(f"MS Project not available: {com_err}")
        if app.ActiveProject is None:
            pytest.skip("No active MS Project project")
        yield app
    except ImportError as e:
        pytest.skip(f"COM bindings unavailable: {e}")
    finally:
        if pythoncom is not None:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


@pytest.fixture
def clean_test_project(msproject_app):
    """Function-scoped fixture providing an ISOLATED empty MS Project for tests.

    SAFETY: Creates a brand new project via FileNew, leaves user's projects untouched.
    On teardown: closes test project without saving, restores user's original active
    project. Tests using this fixture will NEVER modify the user's real work.
    """
    app = msproject_app
    # Remember user's active project so we can restore focus on teardown
    original_name = None
    try:
        if app.ActiveProject is not None:
            original_name = app.ActiveProject.Name
    except Exception:
        pass

    # Create a fresh isolated project (typically named "ProjectN")
    app.FileNew()
    test_proj = app.ActiveProject
    test_name = test_proj.Name

    yield test_proj

    # Teardown: close test project without saving, restore user's original
    try:
        # Find and activate the test project window, then close without save
        for i in range(1, app.Projects.Count + 1):
            try:
                if app.Projects(i).Name == test_name:
                    try:
                        app.WindowActivate(app.Projects(i).Windows(1).Caption)
                    except Exception:
                        pass
                    try:
                        app.FileClose(0)  # 0 = pjDoNotSave
                    except Exception:
                        pass
                    break
            except Exception:
                continue
        # Restore user's original window focus
        if original_name:
            for i in range(1, app.Projects.Count + 1):
                try:
                    if app.Projects(i).Name == original_name:
                        try:
                            app.WindowActivate(app.Projects(i).Windows(1).Caption)
                        except Exception:
                            pass
                        break
                except Exception:
                    continue
    except Exception:
        pass  # Never let teardown fail tests
