"""Local SQLite data layer. Everything stays on this machine; the file is replaced atomically.

Each refresh writes a brand-new database file and swaps it in, so readers never see a half-written state.
Two tables are *carried over* from the database being replaced, so they survive refreshes:
  refresh_history  one row of headline KPIs per successful refresh (trend over time)
  bom_changes      model-level differences between consecutive refreshes (what moved, per refresh)
"""
from __future__ import annotations

import datetime as dt
import json
import os
import sqlite3
import tempfile
import time
from typing import Optional

from . import transform as T

BOM_COLS = ["model", "project", "inch", "bomHQ", "bomHQDate", "bomLocal", "bomLocalDate", "version", "mp", "mpDate",
            "hqTarget", "hqTargetDate", "localTarget", "localTargetDate", "hqStatus", "localStatus",
            "hqDelta", "localDelta", "firstSop", "firstAppearMpGap",
            # schema 3
            "bomHQIso", "bomLocalIso", "mpIso", "hqTargetIso", "localTargetIso", "gapWeeks", "sopIssue", "mpFrom",
            "nmMpIso", "mpMismatch", "hqActual", "hqActualIso", "hqConfirmed", "localActual", "localActualIso",
            "localConfirmed", "hqAfterLocal"]
MASTER_COLS = ["model", "project", "inch", "projectName", "version", "mp", "mpDate", "hqTarget", "hqTargetDate",
               "localTarget", "localTargetDate", "bomHQ", "bomHQDate", "bomLocal", "bomLocalDate", "hqStatus",
               "localStatus", "hqDelta", "localDelta", "coverage", "firstSop", "firstAppearMpGap",
               # schema 3
               "bomHQIso", "bomLocalIso", "mpIso", "hqTargetIso", "localTargetIso", "gapWeeks", "sopIssue", "mpFrom",
               "nmMpIso", "mpMismatch", "hqActual", "hqActualIso", "hqConfirmed", "localActual", "localActualIso",
               "localConfirmed", "hqAfterLocal"]
BOOL_COLS = {"sopIssue", "hqConfirmed", "localConfirmed", "hqAfterLocal"}
HISTORY_KEEP = 200            # refreshes kept in refresh_history
CHANGES_KEEP = 20             # refreshes whose change lists are kept in bom_changes

# Dataset columns are declared without a type on purpose: SQLite then keeps ints as ints and '' as text
# (e.g. `version` is 202636 for SOP models and '' otherwise, exactly like the original dataset).
SCHEMA = f"""
CREATE TABLE meta (key TEXT PRIMARY KEY, value TEXT);
CREATE TABLE source_files (role TEXT PRIMARY KEY, name TEXT, sheet TEXT, size INTEGER, sha256 TEXT);
CREATE TABLE dash (model TEXT PRIMARY KEY, project TEXT, project_name TEXT, raw_json TEXT);
-- raw New Model cells are untyped so Excel date serials stay numbers (TEXT affinity would turn 46300.0 into text)
CREATE TABLE newmodel_plan (model TEXT PRIMARY KEY, project TEXT, inch TEXT, bom_hq, bom_local, mp, actual_hq, actual_local);
CREATE TABLE sop_item (item TEXT PRIMARY KEY, version INTEGER, first_sop_version INTEGER, mp_week TEXT,
                       mp_from TEXT, has_qty INTEGER, project TEXT, inch TEXT);
CREATE TABLE sop_qty (item TEXT, cate TEXT, version INTEGER, total_qty REAL, first_pos_week TEXT,
                      PRIMARY KEY (item, cate, version));
CREATE TABLE bom_data ({", ".join(BOM_COLS)}, ord INTEGER PRIMARY KEY);
CREATE TABLE master_data ({", ".join(MASTER_COLS)}, ord INTEGER PRIMARY KEY);
CREATE UNIQUE INDEX ix_bom_model ON bom_data(model);
CREATE UNIQUE INDEX ix_master_model ON master_data(model);
CREATE INDEX ix_master_project ON master_data(project);
CREATE TABLE refresh_history (id INTEGER PRIMARY KEY, updated_at TEXT NOT NULL, sop_version INTEGER,
                              models INTEGER, master_models INTEGER, on_time INTEGER, late INTEGER, no_plan INTEGER,
                              sop_issue INTEGER, hq_delay INTEGER, local_delay INTEGER, any_issue INTEGER,
                              local_confirmed INTEGER, added INTEGER, removed INTEGER, changed INTEGER);
CREATE TABLE bom_changes (refresh_id INTEGER NOT NULL, model TEXT NOT NULL, project TEXT, kind TEXT NOT NULL,
                          field TEXT, old TEXT, new TEXT);
CREATE INDEX ix_changes_refresh ON bom_changes(refresh_id);
"""

SCHEMA_VERSION = 3
_HIST_COLS = ["id", "updated_at", "sop_version", "models", "master_models", "on_time", "late", "no_plan", "sop_issue",
              "hq_delay", "local_delay", "any_issue", "local_confirmed", "added", "removed", "changed"]
_CHANGE_COLS = ["refresh_id", "model", "project", "kind", "field", "old", "new"]


def _atomic_replace(tmp: str, dest: str, attempts: int = 40) -> None:
    """os.replace fails on Windows while another process has the target open; retry for a few seconds."""
    for i in range(attempts):
        try:
            os.replace(tmp, dest)
            return
        except PermissionError:
            if i == attempts - 1:
                raise
            time.sleep(0.15)


def cleanup_stale_temp(db_path: str) -> int:
    """Remove half-written temp databases left behind by a crashed run."""
    d = os.path.dirname(os.path.abspath(db_path))
    n = 0
    for f in os.listdir(d) if os.path.isdir(d) else []:
        if f.startswith("tmp") and f.endswith(".db"):
            try:
                os.remove(os.path.join(d, f))
                n += 1
            except OSError:
                pass
    return n


def _cell(v):
    if v is None or isinstance(v, (int, float, str)):
        return v
    return str(v)


def _tables(con) -> set[str]:
    return {r[0] for r in con.execute("SELECT name FROM sqlite_master WHERE type='table'")}


def _cols(con, table: str) -> list[str]:
    return [r[1] for r in con.execute(f"PRAGMA table_info({table})")]


def read_previous(db_path: str) -> dict:
    """Tolerant read of the database about to be replaced - any schema version, missing or corrupt file.
    Returns {"bom": [...] | None, "history": [...], "changes": [...]} (only the columns that exist)."""
    out = {"bom": None, "history": [], "changes": []}
    if not os.path.isfile(db_path):
        return out
    try:
        con = sqlite3.connect(f"file:{db_path}?mode=ro", uri=True)
    except sqlite3.Error:
        return out
    try:
        tables = _tables(con)
        if "bom_data" in tables:
            cols = [c for c in _cols(con, "bom_data") if c != "ord"]
            out["bom"] = [dict(zip(cols, r)) for r in con.execute(f"SELECT {','.join(cols)} FROM bom_data ORDER BY ord")]
        if "refresh_history" in tables:
            out["history"] = [dict(zip(_HIST_COLS, r)) for r in
                              con.execute(f"SELECT {','.join(_HIST_COLS)} FROM refresh_history ORDER BY id")]
        if "bom_changes" in tables:
            out["changes"] = [dict(zip(_CHANGE_COLS, r)) for r in
                              con.execute(f"SELECT {','.join(_CHANGE_COLS)} FROM bom_changes")]
    except sqlite3.Error:
        pass
    finally:
        con.close()
    return out


def write_db(db_path: str, ext: dict, bom: list[dict], master: list[dict], report: dict,
             previous: Optional[dict] = None, changes: Optional[list[dict]] = None,
             sop_tables: Optional[tuple[list, list]] = None, updated_at: Optional[str] = None) -> dict:
    """sop_tables: ready-made (sop_item rows, sop_qty rows) - used when migrating a database, where the full
    weekly SOP vectors are no longer available; normally they are computed from ext["sop"]."""
    os.makedirs(os.path.dirname(os.path.abspath(db_path)), exist_ok=True)
    previous = previous or {"history": [], "changes": []}
    changes = changes or []
    fd, tmp = tempfile.mkstemp(suffix=".db", dir=os.path.dirname(os.path.abspath(db_path)))
    os.close(fd)
    con = None
    try:
        con = sqlite3.connect(tmp)
        con.executescript(SCHEMA)
        cur = con.cursor()
        cur.executemany("INSERT INTO source_files VALUES (?,?,?,?,?)",
                        [(f["role"], f["name"], f["sheet"], f["size"], f["sha256"]) for f in ext["files"]])
        cur.executemany("INSERT OR REPLACE INTO dash VALUES (?,?,?,?)",
                        [(r["model"], r["project"], r["projectName"], json.dumps(r["raw"], default=str)) for r in ext["dash"]["rows"]])
        act = ext["nm"].get("actual", {})
        cur.executemany("INSERT OR IGNORE INTO newmodel_plan VALUES (?,?,?,?,?,?,?,?)",
                        [(r["model"], r["project"], r["inch"], _cell(r["bom_hq"]), _cell(r["bom_local"]), _cell(r["mp"]),
                          _cell(act.get(r["model"], {}).get("bom_hq")), _cell(act.get(r["model"], {}).get("bom_local")))
                         for r in ext["nm"]["plan"]])
        if sop_tables is None:
            sop_tables = _sop_tables(ext["sop"])
        cur.executemany("INSERT INTO sop_item VALUES (?,?,?,?,?,?,?,?)", sop_tables[0])
        cur.executemany("INSERT INTO sop_qty VALUES (?,?,?,?,?)", sop_tables[1])
        cur.executemany(f"INSERT INTO bom_data VALUES ({','.join('?' * (len(BOM_COLS) + 1))})",
                        [tuple(_store(c, r.get(c)) for c in BOM_COLS) + (i,) for i, r in enumerate(bom)])
        cur.executemany(f"INSERT INTO master_data VALUES ({','.join('?' * (len(MASTER_COLS) + 1))})",
                        [tuple(_store(c, r.get(c)) for c in MASTER_COLS) + (i,) for i, r in enumerate(master)])

        updated_at = updated_at or dt.datetime.now().isoformat(timespec="seconds")
        # ---- carried-over history + this refresh
        hist = previous["history"][-(HISTORY_KEEP - 1):]
        rid = (max((h["id"] for h in hist), default=0) or 0) + 1
        k = T.kpis(bom)
        count = lambda kind: sum(1 for c in changes if c["kind"] == kind)
        this = {"id": rid, "updated_at": updated_at, "sop_version": (report.get("sop_versions") or [None, None])[-1],
                "models": k["models"], "master_models": len(master), "on_time": k["on_time"], "late": k["late"],
                "no_plan": k["no_plan"], "sop_issue": k["sop_issue"], "hq_delay": k["hq_delay"],
                "local_delay": k["local_delay"], "any_issue": k["any_issue"], "local_confirmed": k["local_confirmed"],
                "added": count("added"), "removed": count("removed"), "changed": len({c["model"] for c in changes if c["kind"] == "changed"})}
        cur.executemany(f"INSERT INTO refresh_history VALUES ({','.join('?' * len(_HIST_COLS))})",
                        [tuple(h.get(c) for c in _HIST_COLS) for h in hist + [this]])
        keep_ids = {h["id"] for h in (hist + [this])[-CHANGES_KEEP:]}
        old_changes = [c for c in previous["changes"] if c["refresh_id"] in keep_ids]
        cur.executemany(f"INSERT INTO bom_changes VALUES ({','.join('?' * len(_CHANGE_COLS))})",
                        [tuple(c.get(k2) for k2 in _CHANGE_COLS) for c in old_changes] +
                        [(rid, c["model"], c["project"], c["kind"], c["field"], c["old"], c["new"]) for c in changes])

        meta = {
            "schema_version": SCHEMA_VERSION,
            "updated_at": updated_at,
            "refresh_id": rid,
            "files": [{k2: f[k2] for k2 in ("role", "name", "sheet", "size")} for f in ext["files"]],
            "warnings": ext["warnings"],
            "report": report,
            "rules": T.RULES,
            "extract_seconds": ext["seconds"],
            "first_refresh": previous.get("bom") is None,
        }
        cur.executemany("INSERT INTO meta VALUES (?,?)", [(k2, json.dumps(v)) for k2, v in meta.items()])
        con.commit()
        con.close()
        con = None
        _atomic_replace(tmp, db_path)     # atomic swap: readers never see a half-written database
        return meta
    except BaseException:
        if con is not None:
            try:
                con.close()
            except Exception:
                pass
        if os.path.exists(tmp):
            os.remove(tmp)
        raise


def _sop_tables(sop: dict) -> tuple[list, list]:
    summ = T.summarise_sop(sop)
    items = [(i, s["version"], s["first_sop_version"], s["mp_code"], s["mp_from"], int(s["has_qty"]),
              sop["dims"].get(i, {}).get("project"), sop["dims"].get(i, {}).get("inch")) for i, s in summ.items()]
    wk = sop["week_codes"]
    qty = []
    for (item, cate, ver), vec in sop["agg"].items():
        first = next((wk[j] for j, x in enumerate(vec) if x > 0), None)
        qty.append((item, cate, ver, float(sum(vec)), first))
    return items, qty


def _store(col: str, v):
    if col in BOOL_COLS:
        return 1 if v else 0
    return _cell(v)


def read_datasets(db_path: str) -> dict:
    """Return {"bom": [...], "master": [...], "meta": {...}, "history": [...], "changes": [...]}.

    Raises ValueError if the file is missing/corrupt or was written by an incompatible version."""
    if not os.path.isfile(db_path):
        raise ValueError("no database")
    try:
        con = sqlite3.connect(f"file:{db_path}?mode=ro", uri=True)
    except sqlite3.Error as e:
        raise ValueError(f"database unreadable: {e}") from e
    try:
        def rows(table, cols):
            cur = con.execute(f"SELECT {','.join(cols)} FROM {table} ORDER BY ord")
            out = [dict(zip(cols, r)) for r in cur]
            for r in out:
                for c in BOOL_COLS:
                    if c in r:
                        r[c] = bool(r[c])
            return out
        meta = {k: json.loads(v) for k, v in con.execute("SELECT key, value FROM meta")}
        if meta.get("schema_version") != SCHEMA_VERSION:
            raise ValueError("database was written by a different version; refresh the data")
        history = [dict(zip(_HIST_COLS, r)) for r in con.execute(f"SELECT {','.join(_HIST_COLS)} FROM refresh_history ORDER BY id")]
        changes = [dict(zip(_CHANGE_COLS, r)) for r in
                   con.execute(f"SELECT {','.join(_CHANGE_COLS)} FROM bom_changes WHERE refresh_id=? ORDER BY rowid",
                               (meta.get("refresh_id"),))]
        return {"bom": rows("bom_data", BOM_COLS), "master": rows("master_data", MASTER_COLS), "meta": meta,
                "history": history, "changes": changes}
    except sqlite3.Error as e:
        raise ValueError(f"database unreadable: {e}") from e
    finally:
        con.close()
