"""Weekly history: every successful refresh is kept as a snapshot in data/history.db.

dashboard.db always holds the data currently shown; history.db keeps every earlier version so the dashboard can
show trends, "what changed since last week", how a model's dates moved over time, and can go back to (or restore)
any earlier refresh.

  snapshot        one row per refresh: when, which files (with SHA-256), KPIs, warnings
  snapshot_data   the complete BOM / All Models datasets of that refresh, gzip-compressed JSON
  snapshot_model  one narrow row per BOM model per refresh (fast per-model timelines)

Writes are a single transaction (WAL journal), so a crash can never leave half a snapshot behind.
"""
from __future__ import annotations

import datetime as dt
import gzip
import json
import os
import re
import sqlite3
from contextlib import closing
from typing import Optional

SCHEMA = """
CREATE TABLE IF NOT EXISTS snapshot (
    id INTEGER PRIMARY KEY AUTOINCREMENT, taken_at TEXT NOT NULL, fingerprint TEXT NOT NULL,
    files_json TEXT NOT NULL, kpi_json TEXT NOT NULL, warnings_json TEXT NOT NULL, sop_max_version INTEGER);
CREATE TABLE IF NOT EXISTS snapshot_data (snapshot_id INTEGER PRIMARY KEY REFERENCES snapshot(id) ON DELETE CASCADE,
    bom_gz BLOB NOT NULL, master_gz BLOB NOT NULL);
CREATE TABLE IF NOT EXISTS snapshot_model (snapshot_id INTEGER REFERENCES snapshot(id) ON DELETE CASCADE, model TEXT,
    project TEXT, mp TEXT, bom_hq TEXT, bom_local TEXT, hq_target TEXT, local_target TEXT, hq_status TEXT,
    local_status TEXT, gap TEXT, PRIMARY KEY (snapshot_id, model));
CREATE INDEX IF NOT EXISTS snapshot_model_by_model ON snapshot_model(model, snapshot_id);
"""
# Fields compared between two refreshes ("what changed"); week fields also give a slip in weeks.
TRACKED = [("mp", "1st MP"), ("bomHQ", "BOM HQ"), ("bomLocal", "BOM LOCAL"), ("hqStatus", "HQ status"),
           ("localStatus", "LOCAL status"), ("firstAppearMpGap", "1st Appear → 1st MP"), ("project", "Project")]
WEEK_FIELDS = {"mp", "bomHQ", "bomLocal"}
_WEEK = re.compile(r"W(\d{1,2})\s*/\s*(\d{4})")
_GAP = re.compile(r"(-?\d+)\s*weeks?")
KEEP = int(os.environ.get("BOM_HISTORY_KEEP", "520"))      # ~10 years of weekly refreshes


def _connect(path: str) -> sqlite3.Connection:
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    con = sqlite3.connect(path, timeout=30)
    con.execute("PRAGMA journal_mode=WAL")
    con.execute("PRAGMA foreign_keys=ON")
    con.executescript(SCHEMA)
    return con


def week_monday(label) -> Optional[dt.date]:
    m = _WEEK.search(str(label or ""))
    if not m:
        return None
    try:
        return dt.date.fromisocalendar(int(m.group(2)), int(m.group(1)), 1)
    except ValueError:
        return None


def gap_weeks(v) -> Optional[int]:
    m = _GAP.search(str(v or ""))
    return int(m.group(1)) if m else None


def kpis(bom: list[dict], master: list[dict]) -> dict:
    """The dashboard's headline numbers (same rules as the page)."""
    return {
        "bom": len(bom), "master": len(master),
        "onTime": sum(1 for r in bom if r.get("localStatus") == "MATCH"),
        "sopIssue": sum(1 for r in bom if (g := gap_weeks(r.get("firstAppearMpGap"))) is not None and g < 12),
        "hqDelay": sum(1 for r in bom if r.get("hqStatus") == "LATER"),
        "seegDelay": sum(1 for r in bom if r.get("localStatus") == "LATER"),
    }


def fingerprint(files: list[dict]) -> str:
    return "|".join(sorted(f"{f['role']}:{f.get('sha256', '')}" for f in files))


def _pack(rows) -> bytes:
    return gzip.compress(json.dumps(rows, ensure_ascii=False, default=str).encode("utf-8"), 6)


def _summary(row) -> dict:
    sid, taken, files, k, w = row
    return {"id": sid, "taken_at": taken, "files": json.loads(files), "kpi": json.loads(k), "warnings": json.loads(w)}


def latest(path: str) -> Optional[dict]:
    if not os.path.isfile(path):
        return None
    with closing(_connect(path)) as con:
        r = con.execute("SELECT id, taken_at, files_json, kpi_json, warnings_json, fingerprint, sop_max_version "
                        "FROM snapshot ORDER BY id DESC LIMIT 1").fetchone()
    if not r:
        return None
    return _summary(r[:5]) | {"fingerprint": r[5], "sop_max_version": r[6]}


def compare_with_latest(path: str, files: list[dict], sop_max_version: Optional[int]) -> list[str]:
    """Weekly-update checks against the previous refresh: files that did not change and an older SOP file."""
    prev = latest(path)
    if not prev:
        return []
    when = prev["taken_at"].replace("T", " ")[:16]
    notes = []
    if fingerprint(files) == prev["fingerprint"]:
        return [f"These are exactly the same 3 files as the refresh of {when} - nothing new was added to the history."]
    old = {f["role"]: f for f in prev["files"]}
    for f in files:
        o = old.get(f["role"])
        if o and o.get("sha256") and o["sha256"] == f.get("sha256"):
            notes.append(f"'{f['name']}' is unchanged since the refresh of {when} - was the new {f['role']} file forgotten?")
    pv = prev.get("sop_max_version")
    if pv and sop_max_version and sop_max_version < pv:
        notes.append(f"The SOP file looks OLDER than last time: its latest version is {sop_max_version}, "
                     f"the refresh of {when} already had {pv}. Check that the newest SOP file was used.")
    return notes


def record(path: str, meta: dict, bom: list[dict], master: list[dict], files: list[dict],
           sop_max_version: Optional[int] = None) -> Optional[int]:
    """Store a refresh as a snapshot. The same 3 files as the latest snapshot are not stored twice."""
    fp = fingerprint(files)
    with closing(_connect(path)) as con, con:
        last = con.execute("SELECT fingerprint FROM snapshot ORDER BY id DESC LIMIT 1").fetchone()
        if last and last[0] == fp:
            return None
        cur = con.execute("INSERT INTO snapshot (taken_at, fingerprint, files_json, kpi_json, warnings_json, sop_max_version) "
                          "VALUES (?,?,?,?,?,?)",
                          (meta.get("updated_at") or dt.datetime.now().isoformat(timespec="seconds"), fp,
                           json.dumps([{k: f.get(k) for k in ("role", "name", "sheet", "size", "sha256")} for f in files]),
                           json.dumps(kpis(bom, master)), json.dumps(meta.get("warnings", [])), sop_max_version))
        sid = cur.lastrowid
        con.execute("INSERT INTO snapshot_data VALUES (?,?,?)", (sid, _pack(bom), _pack(master)))
        con.executemany("INSERT OR IGNORE INTO snapshot_model VALUES (?,?,?,?,?,?,?,?,?,?,?)",
                        [(sid, r["model"], r.get("project"), r.get("mp"), r.get("bomHQ"), r.get("bomLocal"),
                          r.get("hqTarget"), r.get("localTarget"), r.get("hqStatus"), r.get("localStatus"),
                          r.get("firstAppearMpGap")) for r in bom])
        old = [r[0] for r in con.execute("SELECT id FROM snapshot ORDER BY id DESC LIMIT -1 OFFSET ?", (KEEP,))]
        con.executemany("DELETE FROM snapshot WHERE id=?", [(i,) for i in old])
    return sid


def list_snapshots(path: str) -> list[dict]:
    if not os.path.isfile(path):
        return []
    with closing(_connect(path)) as con:
        rows = con.execute("SELECT id, taken_at, files_json, kpi_json, warnings_json FROM snapshot ORDER BY id").fetchall()
    out = []
    for r in rows:
        s = _summary(r)
        s["files"] = [{k: f.get(k) for k in ("role", "name")} for f in s["files"]]
        s["warnings"] = len(s["warnings"])
        out.append(s)
    return out


def load(path: str, sid: int) -> Optional[dict]:
    """A snapshot's complete datasets, in the /api/data shape."""
    if not os.path.isfile(path):
        return None
    with closing(_connect(path)) as con:
        r = con.execute("SELECT s.id, s.taken_at, s.files_json, s.kpi_json, s.warnings_json, d.bom_gz, d.master_gz "
                        "FROM snapshot s JOIN snapshot_data d ON d.snapshot_id = s.id WHERE s.id=?", (sid,)).fetchone()
    if not r:
        return None
    s = _summary(r[:5])
    return {"bom": json.loads(gzip.decompress(r[5])), "master": json.loads(gzip.decompress(r[6])),
            "meta": {"updated_at": s["taken_at"], "files": s["files"], "warnings": s["warnings"], "snapshot_id": s["id"]}}


def delete(path: str, sid: int) -> bool:
    with closing(_connect(path)) as con, con:
        return con.execute("DELETE FROM snapshot WHERE id=?", (sid,)).rowcount > 0


def _ids(con, a: Optional[int], b: Optional[int]) -> tuple[Optional[int], Optional[int]]:
    if a and b:
        return a, b
    ids = [r[0] for r in con.execute("SELECT id FROM snapshot ORDER BY id DESC LIMIT 2")]
    if b and not a:
        prev = con.execute("SELECT id FROM snapshot WHERE id < ? ORDER BY id DESC LIMIT 1", (b,)).fetchone()
        return (prev[0] if prev else None), b
    return (ids[1] if len(ids) > 1 else None), (ids[0] if ids else None)


def changes(path: str, a: Optional[int] = None, b: Optional[int] = None) -> dict:
    """BOM models added / removed / changed between snapshot a (older) and b (newer); default: the last two."""
    if not os.path.isfile(path):
        return {"from": None, "to": None, "added": [], "removed": [], "changed": []}
    with closing(_connect(path)) as con:
        a, b = _ids(con, a, b)
    old, new = (load(path, a) if a else None), (load(path, b) if b else None)
    if not new:
        return {"from": None, "to": None, "added": [], "removed": [], "changed": []}
    ob = {r["model"]: r for r in (old["bom"] if old else [])}
    nb = {r["model"]: r for r in new["bom"]}
    brief = lambda r: {k: r.get(k) for k in ("model", "project", "mp", "bomLocal", "localStatus")}
    out = {"from": old and {"id": a, "taken_at": old["meta"]["updated_at"]},
           "to": {"id": b, "taken_at": new["meta"]["updated_at"]},
           "added": [brief(nb[m]) for m in sorted(nb.keys() - ob.keys())] if old else [],
           "removed": [brief(ob[m]) for m in sorted(ob.keys() - nb.keys())], "changed": []}
    for m in sorted(nb.keys() & ob.keys()):
        for key, label in TRACKED:
            ov, nv = ob[m].get(key), nb[m].get(key)
            if (ov or "") == (nv or ""):
                continue
            slip = None
            if key in WEEK_FIELDS:
                od, nd = week_monday(ov), week_monday(nv)
                slip = (nd - od).days // 7 if od and nd else None
            out["changed"].append({"model": m, "project": nb[m].get("project"), "field": key, "label": label,
                                   "old": ov, "new": nv, "slip": slip})
    return out


def model_history(path: str, model: str) -> list[dict]:
    if not os.path.isfile(path):
        return []
    with closing(_connect(path)) as con:
        rows = con.execute("SELECT s.id, s.taken_at, m.mp, m.bom_hq, m.bom_local, m.hq_target, m.local_target, "
                           "m.hq_status, m.local_status, m.gap, m.model FROM snapshot s LEFT JOIN snapshot_model m "
                           "ON m.snapshot_id = s.id AND m.model = ? ORDER BY s.id", (model,)).fetchall()
    keys = ("id", "taken_at", "mp", "bomHQ", "bomLocal", "hqTarget", "localTarget", "hqStatus", "localStatus", "firstAppearMpGap")
    out = [dict(zip(keys, r)) | {"inBom": r[10] is not None} for r in rows]
    while out and not out[0]["inBom"]:            # start at the model's first appearance
        out.pop(0)
    return out
