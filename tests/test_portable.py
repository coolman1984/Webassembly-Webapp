"""Portable tests: business rules, the SQLite layer and the HTTP service, without Windows or Excel.

The Windows-only COM modules are replaced by stubs, so this runs on any OS (and in CI):

    python -m unittest tests.test_portable -v

The Excel-dependent behaviour is covered by test_edge_cases / test_service / test_stress on a Windows PC.
"""
from __future__ import annotations

import datetime as dt
import gzip
import http.client
import json
import os
import shutil
import sys
import tempfile
import threading
import unittest
from unittest import mock

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
for _m in ("pythoncom", "pywintypes", "win32api", "win32com", "win32com.client", "win32con", "win32event", "win32process"):
    sys.modules.setdefault(_m, mock.MagicMock(name=_m))

from pipeline import store                    # noqa: E402
from pipeline import transform as T          # noqa: E402


def sop(rows: dict, weeks=("202610", "202611", "202612", "202613"), dims=None) -> dict:
    """rows: {(item, cate, version): [qty per week]} in the shape sources.extract_sop produces."""
    return {"week_codes": list(weeks), "agg": rows, "dims": dims or {}}


def plan(model, hq="", local="", mp="", project="P1", inch="55") -> dict:
    return {"model": model, "project": project, "inch": inch, "bom_hq": hq, "bom_local": local, "mp": mp}


class TestDates(unittest.TestCase):
    def test_excel_serials_and_unambiguous_text(self):
        self.assertEqual(T.parse_date(46023), dt.date(2026, 1, 1))
        self.assertEqual(T.parse_date("2026-03-09"), dt.date(2026, 3, 9))
        self.assertEqual(T.parse_date("2026/3/9 00:00:00"), dt.date(2026, 3, 9))
        self.assertEqual(T.parse_date("20260309"), dt.date(2026, 3, 9))
        self.assertEqual(T.parse_date("9-Mar-2026"), dt.date(2026, 3, 9))
        self.assertEqual(T.parse_date(dt.datetime(2026, 3, 9, 15, 0)), dt.date(2026, 3, 9))

    def test_ambiguous_placeholder_and_out_of_range_values_are_never_guessed(self):
        for v in ("06/07/2026", "-", "TBD", "", None, 12, 999999, "2026-02-30", True):
            self.assertIsNone(T.parse_date(v), repr(v))

    def test_iso_week_helpers(self):
        self.assertEqual(T.week_label(dt.date(2026, 1, 1)), "W01 / 2026")
        self.assertEqual(T.week_label(dt.date(2027, 1, 1)), "W53 / 2026")     # ISO year differs from calendar year
        self.assertEqual(T.week_label(None), T.DASH_CH)
        self.assertEqual(T.week_code_to_monday("202610"), dt.date(2026, 3, 2))
        self.assertIsNone(T.week_code_to_monday("202753"))                        # 2027 has only 52 ISO weeks
        self.assertEqual(T.weeks_between(dt.date(2026, 3, 15), dt.date(2026, 3, 2)), 1)   # Sunday vs Monday
        self.assertEqual(T.fmt_date(dt.date(2026, 3, 2)), "02 Mar 2026")
        self.assertEqual(T.invalid_week_codes(sop({}, weeks=("202610", "202753"))), ["202753"])


class TestRules(unittest.TestCase):
    def test_mp_prefers_latest_ship_then_latest_sop_then_new_model_date(self):
        s = sop({
            ("A", "ship", 202601): [0, 5, 0, 0], ("A", "ship", 202605): [0, 0, 7, 0], ("A", "sop", 202608): [9, 0, 0, 0],
            ("B", "ship", 202605): [0, 0, 0, 0], ("B", "sop", 202603): [0, 0, 0, 4],
            ("C", "sop", 202604): [0, 0, 0, 0],
        })
        summ = T.summarise_sop(s)
        self.assertEqual((summ["A"]["mp_code"], summ["A"]["mp_from"], summ["A"]["version"]), ("202612", "ship", 202608))
        self.assertEqual((summ["B"]["mp_code"], summ["B"]["mp_from"]), ("202613", "sop"))
        self.assertEqual(summ["A"]["first_sop_version"], 202608)
        self.assertFalse(summ["C"]["has_qty"])
        bom, stats = T.build_bom({"plan": [plan("C", mp="2026-06-01"), plan("A"), plan("Z")]}, summ)
        by = {r["model"]: r for r in bom}
        self.assertEqual(by["A"]["mp"], "W12 / 2026")
        self.assertEqual(by["C"]["mp"], "W23 / 2026")                 # nothing in SOP -> New Model MP date
        self.assertEqual(by["Z"]["mp"], T.DASH_CH)
        self.assertEqual((stats["mp_from_sop"], stats["mp_from_newmodel"], stats["mp_missing"]), (1, 1, 1))
        self.assertEqual([r["model"] for r in bom], ["A", "C", "Z"])  # ordered by model code

    def test_targets_status_and_gap(self):
        summ = T.summarise_sop(sop({("M", "ship", 202601): [0, 0, 0, 3], ("M", "sop", 202601): [1, 0, 0, 0]}))
        # MP = W13 2026 (Mon 23 Mar) -> HQ target W00 = W52 2025, LOCAL target W01 2026
        rows = {r["model"]: r for r in T.build_bom({"plan": [
            plan("M", hq="2025-12-22", local="2025-12-29"),
        ]}, summ)[0]}
        m = rows["M"]
        self.assertEqual((m["hqTarget"], m["localTarget"]), ("W52 / 2025", "W01 / 2026"))
        self.assertEqual((m["hqStatus"], m["hqDelta"]), ("MATCH", 0))
        self.assertEqual((m["localStatus"], m["localDelta"]), ("MATCH", 0))
        self.assertEqual((m["firstSop"], m["firstAppearMpGap"]), ("Version 202601", "12 weeks"))
        for hq, status, delta in (("2026-01-12", "LATER", 3), ("2025-12-08", "EARLIER", -2), ("TBD", "N/A", None)):
            r = T.build_bom({"plan": [plan("M", hq=hq)]}, summ)[0][0]
            self.assertEqual((r["hqStatus"], r["hqDelta"]), (status, delta), hq)

    def test_duplicate_plan_rows_keep_the_first(self):
        bom, stats = T.build_bom({"plan": [plan("D", project="first"), plan("D", project="second")]}, {})
        self.assertEqual((len(bom), bom[0]["project"], stats["duplicates"]), (1, "first", 1))

    def test_master_scope_and_coverage(self):
        s = sop({("S1", "sop", 202601): [0, 1, 0, 0], ("S0", "sop", 202601): [0, 0, 0, 0], ("B1", "ship", 202601): [0, 0, 2, 0]},
                dims={"S1": {"project": "", "any_project": "PX", "inch": "65", "any_inch": "65"}})
        summ = T.summarise_sop(s)
        bom, _ = T.build_bom({"plan": [plan("B1")]}, summ)
        dash = {"rows": [{"model": "D1", "project": "PD", "projectName": ""}, {"model": "D1", "project": "", "projectName": "Name"}]}
        master = {r["model"]: r for r in T.build_master(dash, s, summ, bom)}
        self.assertEqual(sorted(master), ["B1", "D1", "S1"])          # S0 has no quantity anywhere -> out of scope
        self.assertEqual(master["D1"]["coverage"], "DASH")
        self.assertEqual((master["D1"]["project"], master["D1"]["projectName"]), ("PD", "Name"))
        self.assertEqual((master["S1"]["coverage"], master["S1"]["project"]), ("SOP", "PX"))
        self.assertEqual(master["B1"]["coverage"], "SOP + BOM")


def _datasets():
    s = sop({("A", "ship", 202601): [0, 0, 0, 3]})
    summ = T.summarise_sop(s)
    bom, bstats = T.build_bom({"plan": [plan("A", hq="2025-12-22", local="2026-01-05"), plan("B")]}, summ)
    master = T.build_master({"rows": [{"model": "A", "project": "P1", "projectName": "Proj عربي"}]}, s, summ, bom)
    ext = {"files": [{"role": "DASH", "name": "d.xlsx", "sheet": "Sheet1", "size": 10, "sha256": "x"}],
           "dash": {"rows": [{"model": "A", "project": "P1", "projectName": "", "raw": {"k": 1}}]},
           "nm": {"plan": [plan("A"), plan("B")]}, "sop": s, "warnings": ["note one"], "seconds": 1.0}
    return ext, bom, master, {"bom_models": len(bom), "bom_stats": bstats}


class TestStore(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp(prefix="bomstore_")
        self.addCleanup(shutil.rmtree, self.dir, True)
        self.db = os.path.join(self.dir, "dashboard.db")

    def test_round_trip_keeps_order_types_and_meta(self):
        ext, bom, master, report = _datasets()
        store.write_db(self.db, ext, bom, master, report)
        d = store.read_datasets(self.db)
        self.assertEqual(d["bom"], bom)
        self.assertEqual(d["master"], master)
        self.assertEqual(d["meta"]["warnings"], ["note one"])
        self.assertEqual(d["meta"]["files"][0]["name"], "d.xlsx")

    def test_missing_corrupt_and_stale_databases_are_rejected_cleanly(self):
        with self.assertRaises(ValueError):
            store.read_datasets(self.db)
        with open(self.db, "wb") as f:
            f.write(b"not a database")
        with self.assertRaises(ValueError):
            store.read_datasets(self.db)
        os.remove(self.db)
        ext, bom, master, report = _datasets()
        with mock.patch.object(store, "SCHEMA_VERSION", store.SCHEMA_VERSION - 1):
            store.write_db(self.db, ext, bom, master, report)
        with self.assertRaises(ValueError):
            store.read_datasets(self.db)

    def test_stale_temp_files_are_cleaned(self):
        open(os.path.join(self.dir, "tmpabc.db"), "w").close()
        self.assertEqual(store.cleanup_stale_temp(self.db), 1)


class TestService(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.dir = tempfile.mkdtemp(prefix="bomsvc_")
        import server
        cls.S = server
        cls.patch = mock.patch.multiple(server, DB_PATH=os.path.join(cls.dir, "dashboard.db"), DATA_DIR=cls.dir,
                                        LOG_PATH=os.path.join(cls.dir, "pipeline.log"), INBOX_DIR=os.path.join(cls.dir, "inbox"),
                                        UPLOAD_DIR=os.path.join(cls.dir, "uploads"))
        cls.patch.start()
        cls.no_excel = mock.patch.object(server.X, "excel_installed", return_value=False)
        cls.no_excel.start()
        ext, bom, master, report = _datasets()
        store.write_db(server.DB_PATH, ext, bom, master, report)
        server._load_cache()
        cls.srv = server.ThreadingHTTPServer(("127.0.0.1", 0), server.Handler)
        server.PORT = cls.port = cls.srv.server_address[1]
        threading.Thread(target=cls.srv.serve_forever, daemon=True).start()

    @classmethod
    def tearDownClass(cls):
        cls.srv.shutdown()
        cls.srv.server_close()
        cls.patch.stop()
        cls.no_excel.stop()
        shutil.rmtree(cls.dir, ignore_errors=True)

    def req(self, method, path, headers=None, body=None, host=None):
        c = http.client.HTTPConnection("127.0.0.1", self.port, timeout=10)
        c.putrequest(method, path, skip_host=True, skip_accept_encoding=True)
        c.putheader("Host", host or f"127.0.0.1:{self.port}")
        for k, v in (headers or {}).items():
            c.putheader(k, v)
        if body is not None:
            c.putheader("Content-Length", str(len(body)))
        c.endheaders(body)
        r = c.getresponse()
        data = r.read()
        c.close()
        return r.status, {k.lower(): v for k, v in r.getheaders()}, data

    def test_page_and_data(self):
        st, h, body = self.req("GET", "/")
        self.assertEqual(st, 200)
        self.assertIn(b"BOM Confirmation", body)
        st, h, body = self.req("GET", "/api/data")
        self.assertEqual(st, 200)
        self.assertNotIn("content-encoding", h)
        d = json.loads(body)
        self.assertEqual([r["model"] for r in d["bom"]], ["A", "B"])
        self.assertEqual(self.req("GET", "/api/data", {"If-None-Match": h["etag"]})[0], 304)
        st, hz, gz = self.req("GET", "/api/data", {"Accept-Encoding": "gzip"})
        self.assertEqual((st, hz.get("content-encoding"), hz.get("vary")), (200, "gzip", "Accept-Encoding"))
        self.assertEqual(gzip.decompress(gz), body)

    def test_security_guards(self):
        self.assertEqual(self.req("GET", "/api/data", host="evil.example")[0], 403)
        self.assertEqual(self.req("POST", "/api/process", body=b"")[0], 403)
        self.assertEqual(self.req("PUT", "/api/upload/1", {"X-Requested-With": "bom-dashboard", "Origin": "http://evil.example"}, b"x")[0], 403)
        st, h, _ = self.req("GET", "/api/status", {"Origin": "http://evil.example"})  # readable, but no CORS
        self.assertEqual(st, 200)
        self.assertNotIn("access-control-allow-origin", h)
        self.assertEqual(h.get("x-content-type-options"), "nosniff")
        self.assertFalse(json.loads(_)["excelAvailable"])

    def test_upload_name_is_sanitised(self):
        st, _, body = self.req("PUT", "/api/upload/2", {"X-Requested-With": "bom-dashboard",
                                                        "X-Filename": "..%2F..%2Fevil%3C%3E.xlsx"}, b"data")
        self.assertEqual(st, 200)
        name = json.loads(body)["name"]
        self.assertNotIn("/", name)
        self.assertTrue(os.path.isfile(os.path.join(self.S.UPLOAD_DIR, "slot2__" + name)))


if __name__ == "__main__":
    unittest.main()
