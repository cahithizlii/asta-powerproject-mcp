"""Resolve a `source` parameter into a table bag the readers can consume.

Three source types:

* ``db``    -- a P6 database via an alias (SQLite read-only/snapshot, or SQL Server)
* ``xer``   -- an XER file, parsed by our own xer_parser
* ``pmxml`` -- deliberately refused for now, see below

PMXML is NOT wired up. MPXJ mis-reads P6 data: on the standalone database it
returned 1 project out of 82 and then threw, and on XER it counts PROJWBS rows
as activities (1735 vs the real 950). Until a raw parser exists, returning a
clear error beats returning quietly wrong numbers (RULE 0).
"""
from __future__ import annotations

import os
from typing import Any, Mapping

from . import db as p6db

DEFAULT_LIMIT = 100
MAX_LIMIT = 500


class SourceError(RuntimeError):
    """Bad or unsupported source; tools turn it into a JSON error."""


def clamp_limit(params: Mapping[str, Any]) -> int:
    try:
        n = int(params.get("limit", DEFAULT_LIMIT))
    except (TypeError, ValueError):
        n = DEFAULT_LIMIT
    return max(1, min(n, MAX_LIMIT))


def _resolve_proj_id(backend, params: Mapping[str, Any]) -> int | None:
    """proj_id, or look it up from proj_short_name."""
    if params.get("proj_id") is not None:
        return int(params["proj_id"])
    name = params.get("proj_short_name") or params.get("project")
    if not name:
        return None
    qo, qc, p = backend.quote_open, backend.quote_close, backend.param
    rows = backend.select(
        "SELECT " + qo + "PROJ_ID" + qc + " FROM " + qo + "PROJECT" + qc
        + " WHERE " + qo + "PROJ_SHORT_NAME" + qc + " = " + p
        + " AND " + qo + "DELETE_SESSION_ID" + qc + " IS NULL",
        [name])
    if not rows:
        raise SourceError("Proje bulunamadi: " + str(name))
    if len(rows) > 1:
        raise SourceError(
            "Ayni kisa ada sahip birden fazla proje var: " + str(name)
            + ". proj_id kullanin.")
    return int(rows[0][0])


class OpenSource:
    """Context manager yielding (bag, day_hr_cnt, meta)."""

    def __init__(self, params: Mapping[str, Any]):
        self.params = dict(params or {})
        src = self.params.get("source")
        if isinstance(src, Mapping):
            merged = dict(src)
            # top-level keys still win, so {"source":{...},"proj_id":1} works
            for k, v in self.params.items():
                if k != "source" and v is not None:
                    merged.setdefault(k, v)
            self.params = merged
        self.backend = None
        self.bag = None
        self.day_hr_cnt = None
        self.meta: dict[str, Any] = {}

    # -- lifecycle ---------------------------------------------------------
    def __enter__(self):
        stype = (self.params.get("type") or self.params.get("source_type")
                 or ("xer" if self.params.get("path") else "db")).lower()
        if stype == "db":
            self._open_db()
        elif stype == "xer":
            self._open_xer()
        elif stype in ("pmxml", "xml"):
            raise SourceError(
                "PMXML kaynagi henuz desteklenmiyor. MPXJ P6 verisini yanlis "
                "okuyor (veritabaninda 82 projeden 1'i, XER'de WBS satirlarini "
                "aktivite sayiyor), ham bir parser yazilana kadar sessizce "
                "yanlis sonuc dondurmemek icin kapali. XER kullanin.")
        else:
            raise SourceError("Bilinmeyen kaynak tipi: " + stype)
        return self

    def __exit__(self, *exc) -> None:
        if self.backend is not None:
            try:
                self.backend.close()
            except Exception:  # noqa: BLE001
                pass

    # -- db ----------------------------------------------------------------
    def _open_db(self) -> None:
        alias = p6db.resolve_alias(self.params.get("alias"))
        use_snapshot = self.params.get("snapshot", True) and alias.driver == "SQLite"
        backend, info = p6db.open_backend(
            alias, password=self.params.get("_password", ""),
            use_snapshot=bool(use_snapshot))
        self.backend = backend
        self.meta.update(info)
        self.meta["type"] = "db"

        proj_id = _resolve_proj_id(backend, self.params)
        self.meta["proj_id"] = proj_id
        if proj_id is None:
            # project-less read (list_projects, db_info, raw sql)
            self.bag = p6db.TableBag()
            return

        self.day_hr_cnt, src = p6db.resolve_day_hr_cnt(
            backend, proj_id, self.params.get("day_hr_cnt"))
        self.meta["day_hr_cnt"] = self.day_hr_cnt
        self.meta["day_hr_cnt_source"] = src
        self.bag = p6db.load_bag(backend, proj_id)
        self.meta["row_counts"] = self.bag.counts()

    # -- xer ---------------------------------------------------------------
    def _open_xer(self) -> None:
        path = self.params.get("path") or self.params.get("file_path")
        if not path:
            raise SourceError("XER kaynagi icin 'path' zorunlu.")
        if not os.path.exists(path):
            raise SourceError("Dosya yok: " + path)
        import xer_parser

        self.bag = xer_parser.XerFile(path)
        self.meta.update(type="xer", path=path,
                         bytes=os.path.getsize(path),
                         tables=sorted(self.bag.tables))
        explicit = self.params.get("day_hr_cnt")
        if explicit is not None:
            self.day_hr_cnt = float(explicit)
            self.meta["day_hr_cnt_source"] = "parametre"
        else:
            self.day_hr_cnt, self.meta["day_hr_cnt_source"] = _xer_day_hr_cnt(self.bag)
        self.meta["day_hr_cnt"] = self.day_hr_cnt

    # -- helpers -----------------------------------------------------------
    def require_project(self) -> None:
        if self.meta.get("type") == "db" and self.meta.get("proj_id") is None:
            raise SourceError(
                "Bu islem icin proj_id (veya proj_short_name) zorunlu. "
                "Projeleri p6_query action='list_projects' ile listeleyin.")


def _xer_day_hr_cnt(bag) -> tuple[float, str]:
    """Hours per day from the XER's own CALENDAR/PROJECT rows. Never assumes 8."""
    proj = bag.tables.get("PROJECT", {"rows": []})["rows"]
    cal = bag.tables.get("CALENDAR", {"rows": []})["rows"]
    clndr_id = proj[0].get("clndr_id") if proj else None
    if clndr_id:
        for row in cal:
            if row.get("clndr_id") == clndr_id and row.get("day_hr_cnt"):
                return (float(row["day_hr_cnt"]),
                        "CALENDAR:clndr_id=%s (%s)" % (clndr_id,
                                                       row.get("clndr_name", "")))
    for row in cal:
        if row.get("default_flag") == "Y" and row.get("day_hr_cnt"):
            return (float(row["day_hr_cnt"]),
                    "CALENDAR:default (%s)" % row.get("clndr_name", ""))
    raise SourceError(
        "XER icinde gun-saat (day_hr_cnt) bulunamadi; varsayilamaz (RULE 0). "
        "day_hr_cnt parametresini acikca verin.")
