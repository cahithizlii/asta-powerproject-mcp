"""Read-only access to a P6 Professional database.

Two backends behind one interface:

* ``SqliteBackend``    -- standalone ``PPMDBSQLite.db``. Opened ``mode=ro`` with
  ``PRAGMA query_only``; callers should read a ``snapshot()`` because the live
  file is WAL-mode and P6 may be writing to it.
* ``SqlServerBackend`` -- a P6 PPM schema on SQL Server, via pyodbc.

Both produce the same *table bag*: an object exposing ``.tables`` shaped exactly
like ``xer_parser.XerFile.tables``. That lets us call the already tested
``xer_parser`` readers unbound, so the XER field mapping -- notably
``forecast_finish = reend_date`` (RULE 16.B) -- lives in exactly one place.

Nothing in this module writes to the database.
"""
from __future__ import annotations

import os
import re
import shutil
import sqlite3
import time
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from typing import Any, Iterable, Sequence

# Tables the readers need. Project-scoped ones carry PROJ_ID; the rest are global.
PROJECT_SCOPED = ("PROJECT", "PROJWBS", "TASK", "TASKPRED", "TASKRSRC", "PROJPROP")
GLOBAL_TABLES = ("RSRC", "CALENDAR")
DEFAULT_TABLES = PROJECT_SCOPED + GLOBAL_TABLES

BOOTSTRAP = os.path.join(
    os.environ.get("APPDATA", ""),
    "Oracle", "Primavera P6", "P6 Professional", "24.12.0", "prmbootstrapV2.xml",
)


class P6DbError(RuntimeError):
    """Recoverable database problem; tools turn it into a JSON error."""


# ---------------------------------------------------------------------------
# Alias resolution -- the ONLY way a database location is determined.
# ---------------------------------------------------------------------------
@dataclass
class Alias:
    name: str
    driver: str                 # SQLite | SQLServer | Oracle | CloudServer
    database: str               # file path (SQLite) or database name
    host: str = ""
    user: str = ""
    public_group_id: str = "1"
    driver_func: str = ""       # IJob.Execute conDriverFunc
    library_name: str = ""      # IJob.Execute conLibraryName
    vendor_lib: str = ""        # IJob.Execute conVendorLib
    is_default: bool = False


def _driver_table(root: ET.Element) -> dict[str, dict[str, str]]:
    out: dict[str, dict[str, str]] = {}
    drivers = root.find(".//Drivers")
    if drivers is None:
        return out
    for drv in drivers:
        out[drv.tag] = {
            "driver_func": (drv.findtext("GetDriverFunc") or "").strip(),
            "library_name": (drv.findtext("LibraryName") or "").strip(),
            "vendor_lib": (drv.findtext("VendorLib") or "").strip(),
        }
    return out


def list_aliases(bootstrap: str = BOOTSTRAP) -> list[Alias]:
    """Parse prmbootstrapV2.xml. Never guesses, never searches the disk."""
    if not os.path.exists(bootstrap):
        raise P6DbError("P6 bootstrap dosyasi bulunamadi: " + bootstrap)
    root = ET.parse(bootstrap).getroot()
    node = root.find(".//DataBaseAliases")
    if node is None:
        raise P6DbError("prmbootstrapV2.xml icinde DataBaseAliases yok")
    default_name = node.get("DefaultPMAlias", "")
    drivers = _driver_table(root)
    out: list[Alias] = []
    for al in node.findall("Alias"):
        conn = al.find("Connection")

        def get(tag: str, _conn=conn) -> str:
            if _conn is None:
                return ""
            return (_conn.findtext(tag) or "").strip()

        name = (al.findtext("Name") or "").strip()
        drv = get("DriverName")
        meta = drivers.get(drv, {})
        out.append(Alias(
            name=name,
            driver=drv,
            database=get("DataBase"),
            host=get("HostName"),
            user=get("User_Name"),
            public_group_id=get("PublicGroupId") or "1",
            driver_func=meta.get("driver_func", ""),
            library_name=meta.get("library_name", ""),
            vendor_lib=meta.get("vendor_lib", ""),
            is_default=(name == default_name),
        ))
    return out


def resolve_alias(name: str | None = None, bootstrap: str = BOOTSTRAP) -> Alias:
    aliases = list_aliases(bootstrap)
    if not aliases:
        raise P6DbError("Kayitli P6 alias'i yok")
    if name:
        for a in aliases:
            if a.name.lower() == name.lower():
                return a
        raise P6DbError(
            "Alias bulunamadi: " + name + ". Kayitli alias'lar: "
            + ", ".join(a.name for a in aliases)
        )
    for a in aliases:
        if a.is_default:
            return a
    return aliases[0]


# ---------------------------------------------------------------------------
# Table bag -- duck-types xer_parser.XerFile for the read_* functions
# ---------------------------------------------------------------------------
@dataclass
class TableBag:
    tables: dict[str, dict[str, Any]] = field(default_factory=dict)
    source: dict[str, Any] = field(default_factory=dict)

    def counts(self) -> dict[str, int]:
        return {k: len(v["rows"]) for k, v in self.tables.items()}


def _normalise(cols: Sequence[str], rows: Iterable[Sequence[Any]]) -> dict[str, Any]:
    """Lower-case column names and stringify every value.

    xer_parser expects XER text semantics: lower-case keys, string values.
    Native SQL types must be stringified or ``_to_int(0)`` returns ``None`` --
    it short-circuits on falsy input. That is the R10 trap.
    """
    lower = [c.lower() for c in cols]
    out_rows = [
        {k: ("" if v is None else str(v)) for k, v in zip(lower, row)}
        for row in rows
    ]
    return {"headers": lower, "rows": out_rows}


# ---------------------------------------------------------------------------
# Backends
# ---------------------------------------------------------------------------
class Backend:
    """Minimal read interface every backend implements."""

    quote_open = '"'
    quote_close = '"'
    param = "?"

    def columns(self, table: str) -> list[str]:
        raise NotImplementedError

    def select(self, sql: str, params: Sequence[Any]) -> list[Sequence[Any]]:
        raise NotImplementedError

    def select_named(self, sql: str, params: Sequence[Any]
                     ) -> tuple[list[str], list[Sequence[Any]]]:
        """Like select() but also returns the column names."""
        raise NotImplementedError

    def close(self) -> None:
        raise NotImplementedError

    def __enter__(self):
        return self

    def __exit__(self, *exc) -> None:
        self.close()


class SqliteBackend(Backend):
    def __init__(self, path: str):
        if not os.path.exists(path):
            raise P6DbError("Veritabani dosyasi yok: " + path)
        # mode=ro + query_only: writing is impossible even by accident.
        # immutable=1 is deliberately NOT used -- it ignores the -wal file and
        # would silently return stale data while P6 is open (RULE 0).
        self.path = path
        self.con = sqlite3.connect("file:" + path + "?mode=ro", uri=True)
        self.con.execute("PRAGMA query_only = 1")

    def columns(self, table: str) -> list[str]:
        cur = self.con.execute("PRAGMA table_info(" + table + ")")
        return [r[1] for r in cur.fetchall()]

    def select(self, sql: str, params: Sequence[Any]) -> list[Sequence[Any]]:
        return self.con.execute(sql, tuple(params)).fetchall()

    def select_named(self, sql: str, params: Sequence[Any]):
        cur = self.con.execute(sql, tuple(params))
        cols = [c[0] for c in (cur.description or [])]
        return cols, cur.fetchall()

    def close(self) -> None:
        self.con.close()


class SqlServerBackend(Backend):
    quote_open = "["
    quote_close = "]"

    def __init__(self, server: str, database: str, user: str = "",
                 password: str = "", driver: str | None = None,
                 trusted: bool | None = None):
        import pyodbc  # lazy: SQLite-only use needs no ODBC stack

        driver = driver or self.pick_driver()
        # The alias stores the public user's password encrypted with P6's own
        # key, which we cannot read. Windows authentication is therefore the
        # default: no credential has to be handled at all. Pass a user and
        # password explicitly when the OS account is not a SQL Server login.
        if trusted is None:
            trusted = not password
        cs = ("DRIVER={" + driver + "};SERVER=" + server + ";DATABASE="
              + database + ";TrustServerCertificate=yes;Encrypt=no;")
        cs += ("Trusted_Connection=yes;" if trusted
               else "UID=" + user + ";PWD=" + password + ";")
        self.con = pyodbc.connect(cs, readonly=True, timeout=30)
        self.database = database
        self.auth = "windows" if trusted else ("sql:" + user)

    @staticmethod
    def pick_driver() -> str:
        import pyodbc

        available = pyodbc.drivers()
        for name in ("ODBC Driver 18 for SQL Server",
                     "ODBC Driver 17 for SQL Server",
                     "SQL Server Native Client 11.0",
                     "SQL Server"):
            if name in available:
                return name
        raise P6DbError("SQL Server ODBC surucusu yok. Kurulu: " + repr(available))

    def columns(self, table: str) -> list[str]:
        """Column names of one object, in ordinal order.

        Resolved through OBJECT_ID rather than INFORMATION_SCHEMA filtered on
        the table name alone: P6's schema installs a `privuser` VIEW over
        every `dbo` base table (164 of them here), so a name-only filter
        returns each column TWICE -- once per schema. Reads survived that
        because the rows land in a dict keyed by column name and the
        duplicates collapse, but every query fetched each column twice, and
        the XER writer turned the duplicate list straight into a malformed
        `%F` header. OBJECT_ID resolves to a single object via the
        connection's default schema.
        """
        cur = self.con.cursor()
        cur.execute(
            "SELECT c.name FROM sys.columns c WHERE c.object_id = OBJECT_ID(?) "
            "ORDER BY c.column_id",
            table,
        )
        return [r[0] for r in cur.fetchall()]

    def select(self, sql: str, params: Sequence[Any]) -> list[Sequence[Any]]:
        cur = self.con.cursor()
        cur.execute(sql, *params)
        return cur.fetchall()

    def select_named(self, sql: str, params: Sequence[Any]):
        cur = self.con.cursor()
        cur.execute(sql, *params)
        cols = [c[0] for c in (cur.description or [])]
        return cols, cur.fetchall()

    def close(self) -> None:
        self.con.close()


def connect_rw(alias: Alias, user: str = "", password: str = "",
               trusted: bool = True, driver: str | None = None):
    """Writable pyodbc connection -- ONLY for the JOBSVC queue (p6/jobs.py).

    Read paths must keep using the read-only backends above. Windows
    authentication is the default so no credential has to be handled at all;
    pass user/password when the OS account is not a SQL Server login.
    """
    import pyodbc  # lazy

    if alias.driver != "SQLServer":
        raise P6DbError(
            "Yazilabilir baglanti yalnizca SQL Server icin: " + alias.driver)
    drv = driver or SqlServerBackend.pick_driver()
    server = alias.host or "localhost"
    cs = ("DRIVER={" + drv + "};SERVER=" + server + ";DATABASE=" + alias.database
          + ";TrustServerCertificate=yes;Encrypt=no;")
    cs += "Trusted_Connection=yes;" if trusted else ("UID=" + user + ";PWD=" + password + ";")
    return pyodbc.connect(cs, timeout=30, autocommit=True)


def open_backend(alias: Alias, password: str = "", use_snapshot: bool = True,
                 snapshot_dir: str | None = None) -> tuple[Backend, dict[str, Any]]:
    """Open the right backend for `alias`. Returns (backend, source_info)."""
    if alias.driver == "SQLite":
        path = alias.database
        info: dict[str, Any] = {"driver": "SQLite", "alias": alias.name}
        if use_snapshot:
            snap = snapshot(path, snapshot_dir or os.path.join(
                os.environ.get("TEMP", "."), "p6_mcp_snapshots"))
            info["snapshot"] = snap
            path = snap["path"]
        else:
            info["snapshot"] = None
        info["path"] = path
        return SqliteBackend(path), info
    if alias.driver == "SQLServer":
        server = alias.host or "localhost"
        backend = SqlServerBackend(server, alias.database, alias.user, password)
        info = {"driver": "SQLServer", "alias": alias.name,
                "server": server, "database": alias.database,
                "auth": backend.auth}
        return backend, info
    raise P6DbError("Desteklenmeyen surucu: " + alias.driver)


# ---------------------------------------------------------------------------
# Snapshot -- consistent read while P6 holds the live SQLite file
# ---------------------------------------------------------------------------
def snapshot(db_path: str, dest_dir: str) -> dict[str, Any]:
    """VACUUM INTO a consistent copy; fall back to copying .db + -wal + -shm.

    Copying only the .db loses everything still sitting in the WAL, so the
    fallback copies all three files together.
    """
    os.makedirs(dest_dir, exist_ok=True)
    stamp = time.strftime("%Y%m%d_%H%M%S")
    dest = os.path.join(dest_dir, "p6_snap_" + stamp + ".db")
    src_stat = os.stat(db_path)
    wal = db_path + "-wal"
    info: dict[str, Any] = {
        "source": db_path,
        "source_mtime": time.strftime("%Y-%m-%d %H:%M:%S",
                                      time.localtime(src_stat.st_mtime)),
        "wal_bytes": os.path.getsize(wal) if os.path.exists(wal) else 0,
        "taken_at": time.strftime("%Y-%m-%d %H:%M:%S"),
    }
    try:
        con = sqlite3.connect("file:" + db_path + "?mode=ro", uri=True)
        try:
            con.execute("VACUUM INTO ?", (dest,))
        finally:
            con.close()
        info.update(path=dest, method="VACUUM INTO", bytes=os.path.getsize(dest))
        return info
    except Exception as exc:  # noqa: BLE001 - fall back, never fail the read
        for suffix in ("", "-wal", "-shm"):
            if os.path.exists(db_path + suffix):
                shutil.copy2(db_path + suffix, dest + suffix)
        info.update(path=dest, method="3-file copy (" + str(exc) + ")",
                    bytes=os.path.getsize(dest))
        return info


# ---------------------------------------------------------------------------
# Loading
# ---------------------------------------------------------------------------
def load_bag(backend: Backend, proj_id: int | None,
             tables: Sequence[str] = DEFAULT_TABLES) -> TableBag:
    """Read `tables` into a bag, filtered to `proj_id` and to live rows.

    Every P6 table carries DELETE_SESSION_ID for soft deletes; those rows are
    absent from an XER export, so they must be excluded here too (R9).
    """
    bag = TableBag()
    qo, qc = backend.quote_open, backend.quote_close
    for table in tables:
        cols = backend.columns(table)
        if not cols:
            continue
        upper = {c.upper() for c in cols}
        where: list[str] = []
        params: list[Any] = []
        if proj_id is not None and "PROJ_ID" in upper and table in PROJECT_SCOPED:
            where.append(qo + "PROJ_ID" + qc + " = " + backend.param)
            params.append(proj_id)
        if "DELETE_SESSION_ID" in upper:
            where.append(qo + "DELETE_SESSION_ID" + qc + " IS NULL")
        col_list = ",".join(qo + c + qc for c in cols)
        sql = "SELECT " + col_list + " FROM " + qo + table + qc
        if where:
            sql += " WHERE " + " AND ".join(where)
        bag.tables[table] = _normalise(cols, backend.select(sql, params))
    return bag


def list_projects(backend: Backend, project_flag: str = "Y") -> list[dict[str, Any]]:
    """Real projects (`Y`) or EPS nodes (`N`)."""
    qo, qc, p = backend.quote_open, backend.quote_close, backend.param
    sql = (
        "SELECT " + qo + "PROJ_ID" + qc + ", " + qo + "PROJ_SHORT_NAME" + qc
        + ", " + qo + "CLNDR_ID" + qc + ", " + qo + "LAST_RECALC_DATE" + qc
        + ", " + qo + "LAST_SCHEDULE_DATE" + qc
        + " FROM " + qo + "PROJECT" + qc
        + " WHERE " + qo + "PROJECT_FLAG" + qc + " = " + p
        + " AND " + qo + "DELETE_SESSION_ID" + qc + " IS NULL"
    )
    rows = backend.select(sql, [project_flag])
    return [
        {"proj_id": r[0], "proj_short_name": r[1], "clndr_id": r[2],
         "data_date": str(r[3])[:10] if r[3] else None,
         "last_schedule_date": str(r[4])[:19] if r[4] else None}
        for r in rows
    ]


def resolve_day_hr_cnt(backend: Backend, proj_id: int,
                       explicit: float | None = None) -> tuple[float, str]:
    """Hours per working day. RULE 0: never silently defaults to 8.

    Order: explicit argument -> PROJECT.CLNDR_ID -> CALENDAR.DAY_HR_CNT -> error.
    """
    if explicit is not None:
        return float(explicit), "parametre"
    qo, qc, p = backend.quote_open, backend.quote_close, backend.param
    rows = backend.select(
        "SELECT " + qo + "CLNDR_ID" + qc + " FROM " + qo + "PROJECT" + qc
        + " WHERE " + qo + "PROJ_ID" + qc + " = " + p,
        [proj_id],
    )
    if not rows or rows[0][0] is None:
        raise P6DbError(
            "Proje " + str(proj_id) + " icin CLNDR_ID okunamadi; gun-saat "
            "degeri varsayilamaz (RULE 0). day_hr_cnt parametresini acikca verin."
        )
    clndr_id = rows[0][0]
    rows = backend.select(
        "SELECT " + qo + "DAY_HR_CNT" + qc + ", " + qo + "CLNDR_NAME" + qc
        + " FROM " + qo + "CALENDAR" + qc
        + " WHERE " + qo + "CLNDR_ID" + qc + " = " + p,
        [clndr_id],
    )
    if not rows or rows[0][0] is None:
        raise P6DbError(
            "Takvim " + str(clndr_id) + " icin DAY_HR_CNT bos; gun-saat degeri "
            "varsayilamaz (RULE 0)."
        )
    return float(rows[0][0]), "CALENDAR:CLNDR_ID=" + str(clndr_id) + " (" + str(rows[0][1]) + ")"


# ---------------------------------------------------------------------------
# Readers -- delegate to xer_parser so the field mapping stays in one place
# ---------------------------------------------------------------------------
def _xer():
    import xer_parser  # local import keeps this module importable standalone
    return xer_parser


def read_tasks(bag: TableBag, day_hr_cnt: float) -> list[dict[str, Any]]:
    return _xer().XerFile.read_tasks(bag, day_hr_cnt=day_hr_cnt)


def read_links(bag: TableBag, day_hr_cnt: float) -> list[dict[str, Any]]:
    return _xer().XerFile.read_links(bag, day_hr_cnt=day_hr_cnt)


def read_resources(bag: TableBag) -> list[dict[str, Any]]:
    return _xer().XerFile.read_resources(bag)


def read_assignments(bag: TableBag) -> list[dict[str, Any]]:
    return _xer().XerFile.read_assignments(bag)


def read_calendars(bag: TableBag) -> list[dict[str, Any]]:
    return _xer().XerFile.read_calendars(bag)


def read_wbs(bag: TableBag) -> list[dict[str, Any]]:
    return _xer().XerFile.read_wbs(bag)


def read_project(bag: TableBag) -> dict[str, Any]:
    return _xer().XerFile.read_project(bag)


# ---------------------------------------------------------------------------
# Schedule options -- PROJPROP blob; the gate before trusting an external CPM
# ---------------------------------------------------------------------------
P6_DEFAULT_SCHED_OPTIONS = {
    "sched_retained_logic": "Y",
    "sched_progress_override": "N",
    "sched_float_type": "FT_FF",
    "sched_calendar_on_relationship_lag": "rcal_Predecessor",
    "sched_outer_depend_type": "SD_Both",
    "sched_open_critical_flag": "N",
}

_OPTION_PREFIXES = ("sched_", "level_", "enable_", "limit_", "max_", "use_", "key_")

# key|value pairs anywhere in the blob, ignoring the surrounding nesting tokens
_OPTION_RE = re.compile(
    r"((?:sched|level|enable|limit|max|use|key)_[a-z0-9_]+)\|([^|()]*)"
)


def parse_schedule_options(bag: TableBag) -> dict[str, Any]:
    """Pull sched_*/level_* keys out of PROJPROP.PROP_NAME='scheduling'.

    The blob is a Primavera nested-list dump; the settings sit in the first
    group as a flat ``key|value|key|value`` run.
    """
    tbl = bag.tables.get("PROJPROP", {"rows": []})
    raw = ""
    for row in tbl["rows"]:
        if (row.get("prop_name") or "").lower() == "scheduling":
            raw = row.get("prop_value") or ""
            break
    if not raw:
        return {"found": False, "options": {}, "deviations": {},
                "cpm_trustworthy": None}

    # The blob is a nested Primavera list dump, e.g.
    #   (0||(sched_float_type|FT_FF|sched_retained_logic|Y|...)((0||LevelList()()))
    # Splitting on "|" and zipping pairs misaligns as soon as a nesting token
    # lands on an even index, so scan for known keys directly instead.
    opts: dict[str, str] = {}
    for match in _OPTION_RE.finditer(raw):
        opts[match.group(1)] = match.group(2).strip()

    if not opts:
        # Parsed nothing -> we know nothing. Never report "trustworthy" here.
        return {
            "found": True,
            "options": {},
            "deviations": {},
            "cpm_trustworthy": None,
            "parse_error": "PROJPROP 'scheduling' blob'undan hicbir sched_* "
                           "anahtari cikarilamadi; blob formati degismis olabilir.",
            "raw_length": len(raw),
        }

    missing = [k for k in P6_DEFAULT_SCHED_OPTIONS if k not in opts]
    deviations = {
        k: {"actual": opts[k], "p6_default": v}
        for k, v in P6_DEFAULT_SCHED_OPTIONS.items()
        if k in opts and opts[k] != v
    }
    return {
        "found": True,
        "options": opts,
        "deviations": deviations,
        "missing_keys": missing,
        # Only claim trustworthy when every default key was actually seen
        # and matched. Missing keys leave the verdict undecided.
        "cpm_trustworthy": (not deviations) if not missing else None,
        "raw_length": len(raw),
    }
