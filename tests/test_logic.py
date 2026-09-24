"""Fast tests for the business rules and the database layer - no Excel needed.

    runtime\\python.exe -I -m unittest tests.test_logic -v
"""
from __future__ import annotations

import json
import os
import shutil
import sqlite3
import sys
import tempfile
import unittest
from array import array

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
from pipeline import run                     # noqa: E402
from pipeline import store                   # noqa: E402
from pipeline import transform as T          # noqa: E402

WEEKS = ["202630", "202631", "202640", "202641", "202642"]


def ext(plan_mp=46300.0, local="2026/07/27", actual_local=None, extra_model=False):
    """A minimal, hand-checkable extraction result (the shape sources.extract_* produce)."""
    agg = {("TV-A1", "sop", 202630): array("d", [0, 0, 0, 5, 0]),
           ("TV-A1", "ship", 202638): array("d", [0, 0, 0, 0, 9]),        # MP = 202642 (W42 = 12 Oct 2026)
           ("TV-D4", "ship", 202638): array("d", [0, 0, 3, 0, 0])}
    plan = [{"model": "TV-A1", "project": "PRJ-A", "inch": "50", "bom_hq": "2026/07/13", "bom_local": local, "mp": plan_mp}]
    actual = {"TV-A1": {"bom_hq": "OK", "bom_local": actual_local}}
    if extra_model:
        plan.append({"model": "TV-Z9", "project": "PRJ-Z", "inch": "43.0", "bom_hq": None, "bom_local": None, "mp": None})
    return {
        "files": [{"role": r, "name": f"{r}.xlsx", "sheet": "S", "size": 1, "sha256": r} for r in ("DASH", "SEEG", "NEWMODEL")],
        "dash": {"rows": [{"model": "TV-A1", "project": "MKT-A", "projectName": "Name A", "raw": {}},
                          {"model": "TV-C3", "project": "", "projectName": "", "raw": {}}], "rows_read": 2},
        "sop": {"week_codes": WEEKS, "agg": agg, "rows_read": 3,
                "dims": {"TV-A1": {"project": "PRJ-A", "inch": "50.0", "any_project": "PRJ-A", "any_inch": "50.0"},
                         "TV-D4": {"project": "", "inch": "65.0", "any_project": "-", "any_inch": "65.0"}}},
        "nm": {"plan": plan, "actual": actual, "actual_rows": len(plan)},
        "warnings": [], "seconds": 0.1,
    }


class TestRules(unittest.TestCase):
    def test_bom_row(self):
        bom, master, rep = run.build_datasets(ext())
        a = bom[0]
        self.assertEqual((a["mp"], a["localTarget"], a["bomLocal"], a["localStatus"], a["localDelta"]),
                         ("W42 / 2026", "W30 / 2026", "W31 / 2026", "LATER", 1))
        self.assertEqual((a["hqTarget"], a["bomHQ"], a["hqStatus"]), ("W29 / 2026", "W29 / 2026", "MATCH"))
        self.assertEqual((a["firstSop"], a["gapWeeks"], a["firstAppearMpGap"], a["sopIssue"]), ("Version 202630", 12, "12 weeks", False))
        # New Model MP 46300 = 05 Oct 2026 (W41): SOP is one week LATER -> not a pull-in
        self.assertEqual((a["mpFrom"], a["nmMpIso"], a["mpMismatch"]), ("ship", "2026-10-05", 1))
        self.assertEqual((a["hqActual"], a["hqConfirmed"], a["localActual"], a["localConfirmed"]), ("OK", True, "—", False))
        self.assertEqual(rep["kpis"]["late"], 1)

    def test_on_time_includes_earlier(self):
        bom, _, _ = run.build_datasets(ext(local="2026/07/06"))             # W28: 2 weeks before the W30 target
        self.assertEqual((bom[0]["localStatus"], bom[0]["localDelta"]), ("EARLIER", -2))
        k = T.kpis(bom)
        self.assertEqual((k["on_time"], k["late"], k["no_plan"]), (1, 0, 0))

    def test_actual_cells(self):
        self.assertEqual(T.actual_state("2026/07/20"), ("W30 / 2026", "2026-07-20", True))
        self.assertEqual(T.actual_state(46223.0)[2], True)                  # an Excel date serial
        self.assertEqual(T.actual_state("Done"), ("Done", None, True))
        self.assertEqual(T.actual_state("waiting HQ"), ("waiting HQ", None, False))
        for blank in (None, "", "-", "TBD"):
            self.assertEqual(T.actual_state(blank), ("—", None, False))

    def test_normalisation(self):
        self.assertEqual([T.norm_inch(v) for v in ("43", 43.0, "43.0", "-", None, "32.5", "abc")], ["43", "43", "43", "", "", "32.5", "abc"])
        _, master, _ = run.build_datasets(ext(extra_model=True))
        m = {r["model"]: r for r in master}
        self.assertEqual((m["TV-D4"]["project"], m["TV-D4"]["inch"]), ("", "65"))       # '-' is not a project
        self.assertEqual(m["TV-Z9"]["inch"], "43")
        self.assertEqual(m["TV-A1"]["project"], "MKT-A")                                # DASH project wins

    def test_hq_after_local_and_logic_notes(self):
        e = ext(local="2026/07/06")
        e["nm"]["plan"][0]["bom_hq"] = "2026/07/20"
        bom, _, _ = run.build_datasets(e)
        self.assertTrue(bom[0]["hqAfterLocal"])
        self.assertTrue(any("HQ BOM planned after the LOCAL BOM" in w for w in e["warnings"]), e["warnings"])

    def test_pull_in(self):
        # New Model MP 46342 = 16 Nov 2026 (W47); SOP says W42 -> MP pulled in by 5 weeks
        bom, _, _ = run.build_datasets(ext(plan_mp=46342.0))
        self.assertEqual(bom[0]["mpMismatch"], -5)

    def test_diff(self):
        old, _, _ = run.build_datasets(ext())
        new, _, _ = run.build_datasets(ext(local="2026/07/20", extra_model=True))
        ch = T.diff_bom(old, new)
        kinds = sorted((c["kind"], c["field"]) for c in ch)
        self.assertEqual(kinds, [("added", ""), ("changed", "BOM LOCAL plan"), ("changed", "LOCAL status")])
        self.assertEqual(T.diff_bom(None, new), [])
        self.assertEqual(T.diff_bom(new, new), [])


class TestStore(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp(prefix="bomlogic_")
        self.db = os.path.join(self.dir, "dashboard.db")
        self.addCleanup(shutil.rmtree, self.dir, True)

    def refresh(self, e):
        bom, master, rep = run.build_datasets(e)
        prev = store.read_previous(self.db)
        return store.write_db(self.db, e, bom, master, rep, prev, T.diff_bom(prev["bom"], bom))

    def test_roundtrip_history_and_changes(self):
        m1 = self.refresh(ext())
        d = store.read_datasets(self.db)
        self.assertTrue(m1["first_refresh"])
        self.assertEqual((len(d["history"]), d["changes"]), (1, []))
        self.assertIs(d["bom"][0]["localConfirmed"], False)                  # booleans survive SQLite
        self.assertIs(d["bom"][0]["hqConfirmed"], True)
        self.assertEqual(d["meta"]["rules"], T.RULES)
        self.refresh(ext(local="2026/07/20", extra_model=True))
        d = store.read_datasets(self.db)
        self.assertEqual([h["id"] for h in d["history"]], [1, 2])
        self.assertEqual((d["history"][1]["added"], d["history"][1]["changed"]), (1, 1))
        self.assertEqual({c["kind"] for c in d["changes"]}, {"added", "changed"})
        self.assertTrue(all(c["refresh_id"] == 2 for c in d["changes"]))     # only the latest refresh's changes
        self.refresh(ext(local="2026/07/20", extra_model=True))               # identical data: no changes
        d = store.read_datasets(self.db)
        self.assertEqual((len(d["history"]), d["changes"]), (3, []))

    def test_history_is_capped(self):
        old = store.HISTORY_KEEP
        store.HISTORY_KEEP = 3
        try:
            for _ in range(5):
                self.refresh(ext())
        finally:
            store.HISTORY_KEEP = old
        self.assertEqual([h["id"] for h in store.read_datasets(self.db)["history"]], [3, 4, 5])

    def test_corrupt_previous_database_is_ignored(self):
        with open(self.db, "wb") as f:
            f.write(b"not a database")
        self.assertEqual(store.read_previous(self.db), {"bom": None, "history": [], "changes": []})
        self.refresh(ext())
        self.assertEqual(len(store.read_datasets(self.db)["bom"]), 1)

    def test_migrate_schema_2(self):
        """A database written by the previous version is upgraded in place without Excel."""
        con = sqlite3.connect(self.db)
        con.executescript("""
            CREATE TABLE meta (key TEXT PRIMARY KEY, value TEXT);
            CREATE TABLE source_files (role TEXT PRIMARY KEY, name TEXT, sheet TEXT, size INTEGER, sha256 TEXT);
            CREATE TABLE dash (model TEXT PRIMARY KEY, project TEXT, project_name TEXT, raw_json TEXT);
            CREATE TABLE newmodel_plan (model TEXT PRIMARY KEY, project TEXT, inch TEXT, bom_hq TEXT, bom_local TEXT, mp TEXT);
            CREATE TABLE sop_item (item TEXT PRIMARY KEY, version INTEGER, first_sop_version INTEGER, mp_week TEXT,
                                   mp_from TEXT, has_qty INTEGER, project TEXT, inch TEXT);
            CREATE TABLE sop_qty (item TEXT, cate TEXT, version INTEGER, total_qty REAL, first_pos_week TEXT, PRIMARY KEY (item, cate, version));
            CREATE TABLE bom_data (model, ord INTEGER PRIMARY KEY);
            CREATE TABLE master_data (model, ord INTEGER PRIMARY KEY);
        """)
        con.executemany("INSERT INTO meta VALUES (?,?)", [(k, json.dumps(v)) for k, v in
                        {"schema_version": 2, "updated_at": "2026-09-01T10:00:00", "warnings": ["old note"], "report": {"sop_versions": [202601, 202638]}}.items()])
        con.executemany("INSERT INTO source_files VALUES (?,?,?,?,?)", [("DASH", "d.xlsx", "S", 1, "x")])
        con.executemany("INSERT INTO dash VALUES (?,?,?,?)", [("TV-A1", "MKT-A", "Name A", "{}")])
        # a date serial stored in a TEXT column comes back as '46300.0' and must still be understood as a date
        con.execute("INSERT INTO newmodel_plan VALUES ('TV-A1','PRJ-A','50','2026/07/13','2026/07/27','46300.0')")
        con.execute("INSERT INTO sop_item VALUES ('TV-A1',202638,202630,'202642','ship',1,'PRJ-A','50.0')")
        con.execute("INSERT INTO sop_qty VALUES ('TV-A1','ship',202638,9,'202642')")
        con.commit()
        con.close()
        self.assertTrue(run.migrate(self.db, log=lambda m: None))
        d = store.read_datasets(self.db)
        a = d["bom"][0]
        self.assertEqual((a["localStatus"], a["localDelta"], a["gapWeeks"], a["nmMpIso"], a["mpMismatch"]), ("LATER", 1, 12, "2026-10-05", 1))
        self.assertEqual(d["meta"]["updated_at"], "2026-09-01T10:00:00")        # the data's own timestamp is kept
        self.assertIn("old note", d["meta"]["warnings"])
        self.assertEqual(len(d["history"]), 1)
        self.assertFalse(run.migrate(self.db, log=lambda m: None))              # already current: no-op


if __name__ == "__main__":
    unittest.main()
