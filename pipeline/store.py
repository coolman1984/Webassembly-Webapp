"""Local SQLite data layer. Everything stays on this machine; the file is replaced atomically."""
from __future__ import annotations

import datetime as dt
import json
import os
import sqlite3
import tempfile
import time

from . import transform as T

BOM_COLS = ["model", "project", "inch", "bomHQ", "bomHQDate", "bomLocal", "bomLocalDate", "version", "mp", "mpDate",
            "hqTarget", "hqTargetDate", "localTarget", "localTargetDate", "hqStatus", "localStatus",
            "hqDelta", "localDelta", "firstSop", "firstAppearMpGap"]
MASTER_COLS = ["model", "project", "inch", "projectName", "version", "mp", "mpDate", "hqTarget", "hqTargetDate",
               "localTarget", "localTargetDate", "bomHQ", "bomHQDate", "bomLocal", "bomLocalDate", "hqStatus",
               "localStatus", "hqDelta", "localDelta", "coverage", "firstSop", "firstAppearMpGap"]

# Columns are declared without a type on purpose: SQLite then keeps ints as ints and '' as text
# (e.g. `version` is 202636 for SOP models and '' otherwise, exactly like the original dataset).
SCHEMA = f"""
CREATE TABLE meta (key TEXT PRIMARY KEY, value TEXT);
CREATE TABLE source_files (role TEXT PRIMARY KEY, name TEXT, sheet TEXT, size INTEGER, sha256 TEXT);
CREATE TABLE dash (model TEXT PRIMARY KEY, project TEXT, project_name TEXT, raw_json TEXT);
CREATE TABLE newmodel_plan (model TEXT PRIMARY KEY, project TEXT, inch TEXT, bom_hq TEXT, bom_local TEXT, mp TEXT);
CREATE TABLE sop_item (item TEXT PRIMARY KEY, version INTEGER, first_sop_version INTEGER, mp_week TEXT,
                       mp_from TEXT, has_qty INTEGER, project TEXT, inch TEXT);
CREATE TABLE sop_qty (item TEXT, cate TEXT, version INTEGER, total_qty REAL, first_pos_week TEXT,
                      PRIMARY KEY (item, cate, version));
CREATE TABLE bom_data ({", ".join(BOM_COLS)}, ord INTEGER PRIMARY KEY);
CREATE TABLE master_data ({", ".join(MASTER_COLS)}, ord INTEGER PRIMARY KEY);
"""


SCHEMA_VERSION = 2


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


def write_db(db_path: str, ext: dict, bom: list[dict], master: list[dict], report: dict) -> dict:
    os.makedirs(os.path.dirname(os.path.abspath(db_path)), exist_ok=True)
    fd, tmp = tempfile.mkstemp(suffix=".db", dir=os.path.dirname(os.path.abspath(db_path)))
    os.close(fd)
    try:
        con = sqlite3.connect(tmp)
        con.executescript(SCHEMA)
        cur = con.cursor()
        cur.executemany("INSERT INTO source_files VALUES (?,?,?,?,?)",
                        [(f["role"], f["name"], f["sheet"], f["size"], f["sha256"]) for f in ext["files"]])
        cur.executemany("INSERT OR REPLACE INTO dash VALUES (?,?,?,?)",
                        [(r["model"], r["project"], r["projectName"], json.dumps(r["raw"], default=str)) for r in ext["dash"]["rows"]])
        cur.executemany("INSERT OR IGNORE INTO newmodel_plan VALUES (?,?,?,?,?,?)",
                        [(r["model"], r["project"], r["inch"], _cell(r["bom_hq"]), _cell(r["bom_local"]), _cell(r["mp"]))
                         for r in ext["nm"]["plan"]])
        sop = ext["sop"]
        summ = T.summarise_sop(sop)
        cur.executemany("INSERT INTO sop_item VALUES (?,?,?,?,?,?,?,?)",
                        [(i, s["version"], s["first_sop_version"], s["mp_code"], s["mp_from"], int(s["has_qty"]),
                          sop["dims"].get(i, {}).get("project"), sop["dims"].get(i, {}).get("inch"))
                         for i, s in summ.items()])
        wk = sop["week_codes"]
        rows = []
        for (item, cate, ver), vec in sop["agg"].items():
            first = next((wk[j] for j, x in enumerate(vec) if x > 0), None)
            rows.append((item, cate, ver, float(sum(vec)), first))
        cur.executemany("INSERT INTO sop_qty VALUES (?,?,?,?,?)", rows)
        cur.executemany(f"INSERT INTO bom_data VALUES ({','.join('?' * (len(BOM_COLS) + 1))})",
                        [tuple(r[c] for c in BOM_COLS) + (i,) for i, r in enumerate(bom)])
        cur.executemany(f"INSERT INTO master_data VALUES ({','.join('?' * (len(MASTER_COLS) + 1))})",
                        [tuple(r[c] for c in MASTER_COLS) + (i,) for i, r in enumerate(master)])
        meta = {
            "schema_version": SCHEMA_VERSION,
            "updated_at": dt.datetime.now().isoformat(timespec="seconds"),
            "files": [{k: f[k] for k in ("role", "name", "sheet", "size")} for f in ext["files"]],
            "warnings": ext["warnings"],
            "report": report,
            "extract_seconds": ext["seconds"],
        }
        cur.executemany("INSERT INTO meta VALUES (?,?)", [(k, json.dumps(v)) for k, v in meta.items()])
        con.commit()
        con.close()
        _atomic_replace(tmp, db_path)     # atomic swap: readers never see a half-written database
        return meta
    except BaseException:
        try:
            con.close()
        except Exception:
            pass
        if os.path.exists(tmp):
            os.remove(tmp)
        raise


def read_datasets(db_path: str) -> dict:
    """Return {"bom": [...], "master": [...], "meta": {...}} in the exact shape the dashboard uses.

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
            return [dict(zip(cols, r)) for r in cur]
        meta = {k: json.loads(v) for k, v in con.execute("SELECT key, value FROM meta")}
        if meta.get("schema_version") != SCHEMA_VERSION:
            raise ValueError("database was written by a different version; refresh the data")
        return {"bom": rows("bom_data", BOM_COLS), "master": rows("master_data", MASTER_COLS), "meta": meta}
    except sqlite3.Error as e:
        raise ValueError(f"database unreadable: {e}") from e
    finally:
        con.close()
