"""Offline rules for p6.tasks -- Phase 6 CRUD validation logic.

Everything here fires BEFORE a database session opens (parameter guards) or
runs on a fake session (cycle detection), so the suite needs no P6/SQL
Server. The DB-facing behaviour -- create-from-nothing, F9 date computation,
assignment ledger carry -- is proven live in acceptance part N.
"""
import re

import pytest

from p6 import tasks
from p6.write import P6WriteError


class TestMaps:
    def test_link_types_bidirectional(self):
        assert tasks.LINK_TYPES == {"FS": "PR_FS", "SS": "PR_SS",
                                    "FF": "PR_FF", "SF": "PR_SF"}
        assert tasks.LINK_TYPES_REV["PR_FS"] == "FS"

    def test_guid_is_p6_shaped(self):
        g = tasks._guid()
        assert re.fullmatch(r"[A-Za-z0-9+/]{22}", g), g
        assert tasks._guid() != g


class TestParamGuards:
    def test_add_task_needs_name(self):
        with pytest.raises((P6WriteError, KeyError)):
            tasks.add_task({"proj_id": 1, "duration_h": 8, "confirm": True})

    def test_add_task_needs_duration(self):
        with pytest.raises(P6WriteError, match="duration_h"):
            tasks.add_task({"proj_id": 1, "name": "x", "confirm": True})

    def test_add_task_negative_duration_refused(self):
        with pytest.raises(P6WriteError, match="negatif"):
            tasks.add_task({"proj_id": 1, "name": "x", "duration_h": -1,
                            "confirm": True})

    def test_milestone_needs_no_duration(self):
        # Must get past the duration guard and fail later (no DB) -- not on
        # the duration check.
        with pytest.raises(Exception) as exc:
            tasks.add_task({"proj_id": 1, "name": "m",
                            "task_type": "milestone", "confirm": True})
        assert "duration_h" not in str(exc.value)

    def test_add_task_requires_confirm(self):
        with pytest.raises(P6WriteError, match="confirm"):
            tasks.add_task({"proj_id": 1, "name": "x", "duration_h": 8})

    def test_add_link_self_link_refused(self):
        with pytest.raises(P6WriteError, match="kendine"):
            tasks.add_link({"proj_id": 1, "predecessor": "A",
                            "successor": "A", "confirm": True})

    def test_add_link_bad_type_refused(self):
        with pytest.raises(P6WriteError, match="FS/SS/FF/SF"):
            tasks.add_link({"proj_id": 1, "predecessor": "A",
                            "successor": "B", "link_type": "XX",
                            "confirm": True})

    def test_add_link_accepts_pr_form(self):
        # 'PR_FS' must pass type validation and fail later on the DB.
        with pytest.raises(Exception) as exc:
            tasks.add_link({"proj_id": 1, "predecessor": "A",
                            "successor": "B", "link_type": "PR_FS",
                            "confirm": True})
        assert "FS/SS/FF/SF" not in str(exc.value)

    def test_update_task_needs_fields(self):
        with pytest.raises(P6WriteError, match="alan yok"):
            tasks.update_task({"proj_id": 1, "task_code": "A",
                               "confirm": True})

    def test_create_project_needs_short_name_and_start(self):
        with pytest.raises(P6WriteError, match="zorunlu"):
            tasks.create_project({"short_name": "X", "confirm": True})
        with pytest.raises(P6WriteError, match="zorunlu"):
            tasks.create_project({"plan_start": "2026-01-01",
                                  "confirm": True})


class _CycleSession:
    """Fake session: only what _creates_cycle touches."""

    def __init__(self, edges):
        self._edges = edges  # list of (pred_task_id, task_id)

    def execute(self, sql, *args):
        class _Cur:
            def __init__(self, rows):
                self._rows = rows

            def fetchall(self):
                return self._rows
        return _Cur(self._edges)


class TestCycleDetection:
    def test_direct_cycle(self):
        s = _CycleSession([(1, 2)])          # 1 -> 2 exists
        assert tasks._creates_cycle(s, 99, 2, 1) is True   # adding 2 -> 1

    def test_transitive_cycle(self):
        s = _CycleSession([(1, 2), (2, 3)])  # 1 -> 2 -> 3
        assert tasks._creates_cycle(s, 99, 3, 1) is True   # adding 3 -> 1

    def test_no_cycle_parallel_chains(self):
        s = _CycleSession([(1, 2), (3, 4)])
        assert tasks._creates_cycle(s, 99, 2, 3) is False  # 2 -> 3 is fine

    def test_diamond_is_not_cycle(self):
        s = _CycleSession([(1, 2), (1, 3), (2, 4), (3, 4)])
        assert tasks._creates_cycle(s, 99, 4, 1) is True   # closing loop
        assert tasks._creates_cycle(s, 99, 2, 3) is False  # cross edge ok


class TestActionTable:
    def test_all_actions_registered(self):
        assert set(tasks.ACTIONS) == {
            "create_project", "add_wbs", "add_task", "update_task",
            "delete_task", "add_link", "delete_link", "assign_resource",
            "remove_assignment"}
