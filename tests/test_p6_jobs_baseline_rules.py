"""Offline rules for p6.jobs (dispatchable job types) and p6.baseline.delete.

The DISPATCHABLE whitelist was measured on 26.08.2026: JT_Level and
JT_UpdateBaseline queued fine but the service failed both with
"Invalid Job type"; prmjob.exe's UTF-16LE string table carries the same
seven-type dispatch list immediately before that error string. These tests
pin the refusal so a doomed job is never queued again.

baseline.delete's project guard exists because action='revision' creates a
REAL project (orig_proj_id empty) that the plain baseline delete refused,
leaving revision copies undeletable.
"""
import pytest

from p6 import baseline, jobs


# ---------------------------------------------------------------------------
# jobs: dispatchable whitelist
# ---------------------------------------------------------------------------
class TestDispatchable:
    def test_measured_whitelist(self):
        assert jobs.DISPATCHABLE == {
            "JT_Sched", "JT_ApplyActuals", "JT_XERExport", "JT_Sum",
            "JT_Enterprise_Sum", "JT_Batch", "JT_Report",
        }

    def test_section_names_from_pm_exe(self):
        # PM.exe string table: only JT_Sched uses "Schedule Projects"; every
        # other job type's JOB_DATA section is plain "Projects". Measured:
        # "Apply Actuals" as section -> "No projects to apply actual to.";
        # "Projects" -> JS_Complete.
        assert jobs._SECTION[jobs.JT_SCHEDULE] == "Schedule Projects"
        for jt in (jobs.JT_SUMMARIZE, jobs.JT_APPLY_ACTUALS,
                   jobs.JT_XER_EXPORT, jobs.JT_BATCH_REPORT):
            assert jobs._SECTION[jt] == "Projects", jt

    def test_undispatchable_types_have_no_section(self):
        assert jobs.JT_LEVEL not in jobs._SECTION
        assert jobs.JT_UPDATE_BASELINE not in jobs._SECTION

    @pytest.mark.parametrize("job_type", [jobs.JT_LEVEL, jobs.JT_UPDATE_BASELINE,
                                          jobs.JT_CREATE_BASELINE])
    def test_submit_refuses_undispatchable(self, job_type):
        # The refusal must fire before any cursor use -- cur=None proves it.
        with pytest.raises(jobs.P6JobError, match="CALISTIRAMAZ"):
            jobs.submit(None, job_type, [368], user_id=25)

    def test_schedule_passes_the_gate(self):
        # JT_Sched is dispatchable; with cur=None it must get PAST the gate
        # and fail later, on the cursor -- not with P6JobError.
        with pytest.raises(AttributeError):
            jobs.submit(None, jobs.JT_SCHEDULE, [368], user_id=25)


# ---------------------------------------------------------------------------
# baseline.delete: real-project guard
# ---------------------------------------------------------------------------
class _FakeCursor:
    def __init__(self, row):
        self._row = row

    def fetchone(self):
        return self._row


class _FakeSession:
    """Just enough of write.open_session's session for the guard branches."""

    def __init__(self, project_row, live_baselines=0):
        self._project_row = project_row
        self._live_baselines = live_baselines
        self.stamp = "2026-08-26 00:00:00"

    def execute(self, sql, *args):
        return _FakeCursor(self._project_row)

    def scalar(self, sql, *args):
        return self._live_baselines

    def reserve(self, name, n):  # pragma: no cover - guards fire first
        raise AssertionError("guard must fire before key reservation")

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False


@pytest.fixture
def fake_session(monkeypatch):
    def _install(project_row, live_baselines=0):
        session = _FakeSession(project_row, live_baselines)
        monkeypatch.setattr(baseline.w, "open_session",
                            lambda params: session)
        return session
    return _install


class TestDeleteProjectGuard:
    def test_missing_row_refused(self, fake_session):
        fake_session(project_row=None)
        with pytest.raises(baseline.w.P6WriteError, match="zaten silinmis"):
            baseline.delete({"baseline_proj_id": 999, "confirm": True})

    def test_real_project_needs_delete_project_flag(self, fake_session):
        fake_session(project_row=(None, "sandbox"))
        with pytest.raises(baseline.w.P6WriteError, match="delete_project"):
            baseline.delete({"baseline_proj_id": 378, "confirm": True})

    def test_real_project_needs_matching_name(self, fake_session):
        fake_session(project_row=(None, "sandbox"))
        with pytest.raises(baseline.w.P6WriteError, match="eslesmiyor"):
            baseline.delete({"baseline_proj_id": 378, "confirm": True,
                             "delete_project": True,
                             "expected_short_name": "WRONG"})

    def test_real_project_with_live_baselines_refused(self, fake_session):
        fake_session(project_row=(None, "sandbox"), live_baselines=2)
        with pytest.raises(baseline.w.P6WriteError, match="canli baseline"):
            baseline.delete({"baseline_proj_id": 378, "confirm": True,
                             "delete_project": True,
                             "expected_short_name": "sandbox"})

    def test_confirm_still_required_first(self, fake_session):
        fake_session(project_row=(None, "sandbox"))
        with pytest.raises(Exception):
            baseline.delete({"baseline_proj_id": 378})
