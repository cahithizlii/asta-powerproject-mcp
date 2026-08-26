"""Transactional write layer for the P6 database.

Everything that changes P6 project data goes through here, and it is
deliberately narrow. Three rules the callers rely on:

* **One transaction per operation.** ``autocommit`` is off; a partial progress
  update or a half-copied baseline is worse than no change at all.
* **Keys come from NEXTKEY, reserved in blocks.** P6 hands out ids from that
  table; taking them one row at a time would be slow and would interleave with
  P6's own allocation. ``reserve`` bumps the counter once for the whole block.
* **Audit columns are always stamped.** P6 shows update_user/update_date in
  its own UI, and a row this code wrote should be identifiable as such rather
  than silently inheriting whatever was copied.

There is no COM here and no P6 client involved: P6 Professional exposes no
automation interface (see docs/P6_HANDOFF.md), so the database and the Job
Service queue are the only headless levers that exist. After changing
schedule-relevant data, run p6_job action='schedule' -- P6's own CPM engine
is what recomputes dates, never this module.
"""
from __future__ import annotations

import datetime as _dt
from typing import Any, Iterable, Mapping, Sequence

from . import db as p6db

AUDIT_USER_DEFAULT = "MCP"

# Tables that belong to one project, in the order a copy must insert them
# (parents before children).
PROJECT_TABLES = ("PROJECT", "PROJPROP", "PROJWBS", "TASK", "TASKPRED", "TASKRSRC")


class P6WriteError(RuntimeError):
    """Refused or failed write; tools turn it into a JSON error."""


class Session:
    """A writable connection with an open transaction.

    Use as a context manager: it commits on a clean exit and rolls back on any
    exception, so a caller can never leave half a change behind.
    """

    def __init__(self, alias, user: str = AUDIT_USER_DEFAULT):
        if alias.driver != "SQLServer":
            raise P6WriteError(
                "Yazma yalnizca SQL Server alias'inda desteklenir; bu alias: "
                + alias.driver + ". (SQLite standalone'da Job Service de yok.)")
        self.alias = alias
        self.user = user
        self.conn = p6db.connect_rw(alias)
        self.conn.autocommit = False
        self.cur = self.conn.cursor()
        self.stamp = _dt.datetime.now().replace(microsecond=0)
        self._committed = False

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        try:
            if exc_type is None:
                self.conn.commit()
                self._committed = True
            else:
                self.conn.rollback()
        finally:
            self.conn.close()
        return False

    # -- keys --------------------------------------------------------------
    def reserve(self, key_name: str, count: int) -> list[int]:
        """Take `count` consecutive ids from NEXTKEY, bumping it once.

        key_name is P6's own convention: '<table>_<column>' lower case, e.g.
        'task_task_id'.
        """
        if count <= 0:
            return []
        self.cur.execute("SELECT key_seq_num FROM NEXTKEY WHERE key_name = ?",
                         key_name)
        row = self.cur.fetchone()
        if row is None:
            raise P6WriteError("NEXTKEY kaydi yok: " + key_name)
        start = int(row[0])
        self.cur.execute("UPDATE NEXTKEY SET key_seq_num = ? WHERE key_name = ?",
                         start + count, key_name)
        return list(range(start, start + count))

    # -- reads inside the transaction --------------------------------------
    def columns(self, table: str) -> list[str]:
        self.cur.execute(
            "SELECT c.name FROM sys.columns c WHERE c.object_id = OBJECT_ID(?) "
            "ORDER BY c.column_id", table)
        return [r[0] for r in self.cur.fetchall()]

    def select_rows(self, table: str, where: str, params: Sequence[Any]
                    ) -> tuple[list[str], list[dict[str, Any]]]:
        cols = self.columns(table)
        col_list = ", ".join("[%s]" % c for c in cols)
        self.cur.execute("SELECT %s FROM [%s] WHERE %s" % (col_list, table, where),
                         *params)
        rows = [dict(zip(cols, r)) for r in self.cur.fetchall()]
        return cols, rows

    # -- writes ------------------------------------------------------------
    def has_lob(self, table: str) -> bool:
        """text / varchar(max) / image columns present?

        pyodbc's fast_executemany pre-allocates a buffer per column from the
        declared width, so a single `text` column (TASKRSRC.rsrc_request_data,
        PROJWBS.wbs_memo) makes it try to reserve gigabytes and raise
        MemoryError. Those tables get the ordinary, slower path.
        """
        self.cur.execute(
            "SELECT COUNT(*) FROM sys.columns c "
            "JOIN sys.types ty ON ty.user_type_id = c.user_type_id "
            "WHERE c.object_id = OBJECT_ID(?) "
            "AND (ty.name IN ('text','ntext','image') OR c.max_length = -1)",
            table)
        return int(self.cur.fetchone()[0]) > 0

    def insert_rows(self, table: str, cols: Sequence[str],
                    rows: Iterable[Mapping[str, Any]]) -> int:
        rows = list(rows)
        if not rows:
            return 0
        placeholders = ", ".join("?" for _ in cols)
        sql = "INSERT INTO [%s] (%s) VALUES (%s)" % (
            table, ", ".join("[%s]" % c for c in cols), placeholders)
        try:
            self.cur.fast_executemany = not self.has_lob(table)
        except AttributeError:
            pass
        self.cur.executemany(sql, [[r.get(c) for c in cols] for r in rows])
        return len(rows)

    def stamp_audit(self, row: dict[str, Any], cols: Sequence[str]) -> dict[str, Any]:
        """Mark the row as written by this tool, if the table carries the columns."""
        if "update_date" in cols:
            row["update_date"] = self.stamp
        if "update_user" in cols:
            row["update_user"] = self.user
        if "create_date" in cols and row.get("create_date") is None:
            row["create_date"] = self.stamp
        if "create_user" in cols and row.get("create_user") is None:
            row["create_user"] = self.user
        return row

    def execute(self, sql: str, *params: Any):
        self.cur.execute(sql, *params)
        return self.cur

    def scalar(self, sql: str, *params: Any):
        self.cur.execute(sql, *params)
        row = self.cur.fetchone()
        return row[0] if row else None


def open_session(params: Mapping[str, Any]) -> Session:
    alias = p6db.resolve_alias(params.get("alias"))
    return Session(alias, user=params.get("audit_user") or AUDIT_USER_DEFAULT)


def require_confirm(params: Mapping[str, Any], what: str) -> None:
    """Data-changing actions need an explicit confirm, like p6_job's do."""
    if not params.get("confirm"):
        raise P6WriteError(
            "%s proje verisini DEGISTIRIR. Once dry_run ile ne olacagini "
            "gorun, sonra confirm=true ile tekrar cagirin." % what)


def project_exists(session: Session, proj_id: int) -> str:
    name = session.scalar(
        "SELECT proj_short_name FROM PROJECT WHERE proj_id = ? "
        "AND delete_session_id IS NULL", proj_id)
    if name is None:
        raise P6WriteError("Proje bulunamadi (veya silinmis): %s" % proj_id)
    return name
