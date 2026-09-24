"""Edge-case tests for the extraction/validation/calculation pipeline (real Excel, real files).

    runtime\\python.exe -m unittest tests.test_edge_cases -v
"""
from __future__ import annotations

import os
import random
import shutil
import sys
import tempfile
import time
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
from pipeline import excel_com as X          # noqa: E402
from pipeline import run                     # noqa: E402
from pipeline import sources as S            # noqa: E402
from tests import fixtures as F              # noqa: E402

_xl = None
_tmp = None
_n = 0


def setUpModule():
    global _xl, _tmp
    _xl = F.Excel().__enter__()
    _tmp = tempfile.mkdtemp(prefix="bomtest_")


def tearDownModule():
    _xl.__exit__(None, None, None)
    shutil.rmtree(_tmp, ignore_errors=True)


def make(**o) -> dict:
    global _n
    _n += 1
    return F.build(os.path.join(_tmp, f"case{_n}"), _xl, **o)


def process(files: dict, order=("nm", "sop", "dash")):
    ext = run.extract_all([files[k] for k in order], log=lambda m: None)
    bom, master, rep = run.build_datasets(ext)
    return ext, bom, master


class Base(unittest.TestCase):
    def assertBaseline(self, bom, master):
        self.assertEqual([r["model"] for r in bom], F.BOM_ORDER, "BOM model list/order")
        for r in bom:
            exp = F.EXPECTED_BOM[r["model"]]
            for k, v in exp.items():
                self.assertEqual(r[k], v, f"{r['model']}.{k}")
        cov = {r["model"]: r["coverage"] for r in master}
        self.assertEqual(cov, F.EXPECTED_COVERAGE)                 # also proves TV-E5 (all-zero SOP) is out of scope
        self.assertEqual([r["model"] for r in master], sorted(cov), "master ordered by model")


class TestBaseline(Base):
    def test_hand_checked_oracle(self):
        ext, bom, master = process(make())
        self.assertBaseline(bom, master)
        m = {r["model"]: r for r in master}
        self.assertEqual((m["TV-A1"]["project"], m["TV-A1"]["projectName"], m["TV-A1"]["inch"], m["TV-A1"]["version"]),
                         ("MKT-A", "Name A", "50", 202638))
        d4 = m["TV-D4"]                                           # SOP-only: SOP attributes, '-' project = unknown, 65.0 -> '65'
        self.assertEqual((d4["project"], d4["inch"], d4["mp"], d4["hqTarget"], d4["bomHQ"], d4["hqStatus"], d4["firstSop"]),
                         ("", "65", "W50 / 2026", "W37 / 2026", "—", "N/A", "—"))
        c3 = m["TV-C3"]
        self.assertEqual((c3["version"], c3["mp"], c3["inch"], c3["projectName"]), ("", "—", "", ""))

    def test_no_false_alarms(self):
        """Clean data must produce only the two legitimate notes: TV-F6 has no SOP rows, and 4 models have a
        New Model MP date (W41) that differs from their SOP MP week."""
        ext, _, _ = process(make())
        self.assertEqual(len(ext["warnings"]), 2, ext["warnings"])
        self.assertIn("no rows in the SOP sheet", ext["warnings"][0])
        self.assertIn("4 BOM models have a different 1st MP week", ext["warnings"][1])

    def test_logic_fields(self):
        """Machine-readable fields, New Model 'Actual' rows and cross-file checks."""
        _, bom, master = process(make())
        b = {r["model"]: r for r in bom}
        a1, b2, f6, h8 = b["TV-A1"], b["TV-B2"], b["TV-F6"], b["TV-H8"]
        self.assertEqual((a1["mpIso"], a1["localTargetIso"], a1["bomLocalIso"], a1["gapWeeks"], a1["sopIssue"]),
                         ("2026-10-12", "2026-07-20", "2026-07-27", 12, False))
        self.assertEqual((b2["gapWeeks"], b2["sopIssue"], b2["bomLocalIso"]), (10, True, None))
        self.assertEqual((a1["mpFrom"], a1["nmMpIso"], a1["mpMismatch"]), ("ship", "2026-10-05", 1))
        self.assertEqual((f6["mpFrom"], f6["mpMismatch"]), ("newmodel", None))     # MP from New Model: nothing to compare
        self.assertEqual(h8["mpMismatch"], 13)                                     # W01/2027 vs W41/2026 across the year end
        self.assertEqual((a1["localActual"], a1["localConfirmed"], a1["hqConfirmed"]), ("OK", True, True))
        self.assertFalse(any(r["hqAfterLocal"] for r in bom))
        m = {r["model"]: r for r in master}
        self.assertEqual((m["TV-A1"]["localConfirmed"], m["TV-D4"]["localConfirmed"], m["TV-D4"]["mpFrom"]), (True, False, "ship"))

    def test_any_slot_order(self):
        for order in (("dash", "sop", "nm"), ("sop", "nm", "dash")):
            _, bom, master = process(make(), order)
            self.assertBaseline(bom, master)


class TestStorageVariants(Base):
    """Same meaning, different storage -> identical dashboard data."""
    CASES = {
        "shuffled columns": dict(perm=True, seed=3),
        "shuffled columns (other seed)": dict(perm=True, seed=11),
        "extra junk columns": dict(extra=6, seed=5),
        "noisy headers (case/NBSP/newline/zero-width)": dict(hdr_noise=True, seed=2),
        "title rows above headers": dict(title_rows=3),
        "SOP sheet renamed": dict(sop_name="Weekly plan"),
        "extra + hidden sheets": dict(extra_sheets=True, hidden_first_sheet=True),
        "1904 date system": dict(date1904=True),
        "messy model codes (case/NBSP/zero-width)": dict(model_noise=True, seed=4),
        "item code stored as text": dict(numeric_item=False),
        "SOP saved as .xlsb": dict(sop_fmt="xlsb"),
        "explicit group labels": dict(nm_explicit_groups=True),
        "everything at once": dict(perm=True, extra=4, hdr_noise=True, title_rows=2, sop_name="X", extra_sheets=True,
                                   model_noise=True, date1904=True, sop_fmt="xlsb", seed=9),
    }

    def test_variants(self):
        for name, o in self.CASES.items():
            with self.subTest(name):
                _, bom, master = process(make(**o))
                self.assertBaseline(bom, master)

    def test_wide_sheets_beyond_200_columns(self):
        """Regression: header probe used to stop at column 200 and silently lose weekly columns."""
        _, bom, master = process(make(extra=260, seed=6))
        self.assertBaseline(bom, master)


class TestFuzz(Base):
    """Randomised combinations of the perturbations must never change the result."""

    def test_random_combinations(self):
        rng = random.Random(20260921)
        keys = ["perm", "hdr_noise", "model_noise", "date1904", "nm_explicit_groups", "extra_sheets"]
        for i in range(10):
            o = {k: True for k in keys if rng.random() < 0.5}
            o.update(seed=rng.randrange(1, 10**6), extra=rng.choice([0, 0, 3, 8]), title_rows=rng.choice([0, 0, 1, 4]),
                     sop_fmt=rng.choice(["xlsx", "xlsb"]), numeric_item=rng.random() < 0.5)
            with self.subTest(o=o):
                _, bom, master = process(make(**o), tuple(rng.sample(["nm", "sop", "dash"], 3)))
                self.assertBaseline(bom, master)


class TestValidation(Base):
    def rejected(self, files, order=("nm", "sop", "dash")):
        with self.assertRaises((S.ValidationError, X.ExcelError)) as cm:
            process(files, order)
        return str(cm.exception)

    def test_missing_column_named_in_message(self):
        for rename, must in ((dict(Cate="Category"), "Cate"), (dict(MODEL_CODE="MODEL"), "MODEL_CODE"),
                             (dict(MP="Manufacturing"), "SET PLANT > MP"), (dict(HQ="Head Office"), "BOM > HQ"),
                             (dict(Version="Ver"), "Version")):
            with self.subTest(rename):
                msg = self.rejected(make(rename=rename))
                self.assertIn(must, msg)
                self.assertIn("missing column", msg)

    def test_same_role_twice_and_missing_role(self):
        a, b = make(), make(seed=2, hdr_noise=True)       # b differs in content, so it is not a byte-identical duplicate
        b_dash = os.path.join(os.path.dirname(b["dash"]), "Another DASH file.xlsx")
        shutil.copy(b["dash"], b_dash)
        self.assertIn("Two files look like the DASH", self.rejected({"nm": a["nm"], "sop": a["dash"], "dash": b_dash}))
        # two New Model files + one SOP file: the DASH source is absent
        b_nm = os.path.join(os.path.dirname(b["nm"]), "Another New Model.xlsx")
        shutil.copy(b["nm"], b_nm)
        self.assertIn("Missing source: no DASH", self.rejected({"nm": a["nm"], "sop": a["sop"], "dash": b_nm}))

    def test_same_file_name_in_two_folders_gives_a_clear_message(self):
        a, b = make(), make(seed=2)
        self.assertIn("same file name", self.rejected({"nm": a["nm"], "sop": a["sop"], "dash": b["nm"]}))

    def test_identical_file_twice(self):
        a = make()
        self.assertIn("same file", self.rejected({"nm": a["nm"], "sop": a["nm"], "dash": a["dash"]}))

    def test_not_an_excel_file(self):
        a = make()
        junk = os.path.join(_tmp, "notexcel.xlsx")
        open(junk, "w").write("this is plain text, not a workbook")
        t0 = time.time()
        msg = self.rejected({"nm": a["nm"], "sop": a["sop"], "dash": junk})
        self.assertIn("could not open", msg)
        self.assertLess(time.time() - t0, 60)

    def test_zero_byte_and_wrong_extension_and_lock_file(self):
        a = make()
        z = os.path.join(_tmp, "empty.xlsx"); open(z, "wb").close()
        self.assertIn("empty", self.rejected({"nm": a["nm"], "sop": a["sop"], "dash": z}))
        c = os.path.join(_tmp, "data.csv"); open(c, "w").write("a,b")
        self.assertIn("unsupported type", self.rejected({"nm": a["nm"], "sop": a["sop"], "dash": c}))
        l = os.path.join(_tmp, "~$lock.xlsx"); open(l, "wb").write(b"x")
        self.assertIn("lock", self.rejected({"nm": a["nm"], "sop": a["sop"], "dash": l}))
        self.assertIn("Exactly 3", self.rejected({"nm": a["nm"], "sop": a["sop"]}, order=("nm", "sop")))

    def test_password_protected_fails_fast_instead_of_hanging(self):
        a = make()
        pw = os.path.join(_tmp, "locked.xlsx")
        _xl.save(pw, [("Sheet1", [["a", "b"], [1.0, 2.0]])], F.XLSX, password="secret")
        t0 = time.time()
        msg = self.rejected({"nm": a["nm"], "sop": a["sop"], "dash": pw})
        self.assertIn("password-protected", msg)
        self.assertLess(time.time() - t0, 60, "must not wait for an invisible password prompt")

    def test_empty_sop_data_rejected(self):
        a = make()
        hdr_only = os.path.join(_tmp, "sop_header_only.xlsx")
        weeks = ["202640", "202641"]
        _xl.save(hdr_only, [("SOP", [["Item", "Project", "Inch", "Cate", "Version"] + weeks])], F.XLSX)
        self.assertIn("no data rows", self.rejected({"nm": a["nm"], "sop": hdr_only, "dash": a["dash"]}))

    def test_file_open_in_another_excel_is_still_readable(self):
        a = make()
        other = F.Excel().__enter__()                # a second Excel (like the user's own) holding the file open
        try:
            wb = other.app.Workbooks.Open(a["nm"])
            _, bom, master = process(a)
            self.assertBaseline(bom, master)
            wb.Close(SaveChanges=False)
        finally:
            other.__exit__(None, None, None)

    def test_stale_external_link_does_not_prompt(self):
        """A workbook whose formulas point at a file that no longer exists must open without any dialog."""
        a = make()
        src = os.path.join(_tmp, "linksrc.xlsx"); lnk = os.path.join(_tmp, "linked.xlsx")
        _xl.save(src, [("Sheet1", [[42.0]])], F.XLSX)
        wb = _xl.app.Workbooks.Add()
        wb.Worksheets(1).Range("A1").Formula = f"='{_tmp}\\[linksrc.xlsx]Sheet1'!A1"
        wb.SaveAs(lnk, F.XLSX); wb.Close(SaveChanges=False)
        os.remove(src)
        t0 = time.time()
        msg = self.rejected({"nm": a["nm"], "sop": a["sop"], "dash": lnk})    # opens fine, then not recognised as a source
        self.assertIn("not one of the 3 expected files", msg)
        self.assertLess(time.time() - t0, 60)


class TestDataQuality(Base):
    def test_unknown_cate_reported_and_ignored(self):
        ext, bom, master = process(make(unknown_cate=True))
        self.assertBaseline(bom, master)                            # the extra 'Forecast' row must not change anything
        self.assertTrue(any("unknown Cate" in w and "forecast" in w for w in ext["warnings"]), ext["warnings"])

    def test_invalid_iso_week_header_reported(self):
        ext, bom, master = process(make(bad_week=True))
        self.assertTrue(any("not real ISO weeks" in w and "202753" in w for w in ext["warnings"]), ext["warnings"])

    def test_duplicate_dash_row_first_value_wins(self):
        ext, bom, master = process(make(dup_dash=True))
        self.assertTrue(any("duplicate model rows" in w for w in ext["warnings"]))
        a1 = next(r for r in master if r["model"] == "TV-A1")
        self.assertEqual((a1["project"], a1["projectName"]), ("MKT-A", "Name A"))

    def test_placeholder_dates_are_not_reported_as_errors(self):
        ext, bom, _ = process(make())                                # TV-B2 has '-' as BOM LOCAL date
        self.assertFalse(any("unreadable" in w for w in ext["warnings"]), ext["warnings"])
        b2 = next(r for r in bom if r["model"] == "TV-B2")
        self.assertEqual((b2["bomLocal"], b2["localStatus"], b2["localDelta"]), ("—", "N/A", None))


class TestExcelLayer(unittest.TestCase):
    def _open(self, path):
        return X.ExcelSession()

    def test_last_cell_ignores_formatting_and_named_args_regression(self):
        """Formatting at row 1,000,000 inflates UsedRange; Find(xlPrevious) must still give the real end.
        (Regression: the named-argument form of Find silently returned the *first* cell.)"""
        p = os.path.join(_tmp, "inflated.xlsx")
        wb = _xl.app.Workbooks.Add()
        ws = wb.Worksheets(1)
        ws.Range("A1:C3").Value2 = ((1.0, 2.0, 3.0), (4.0, 5.0, 6.0), (7.0, 8.0, 9.0))
        ws.Range("A1000000").Interior.Color = 255                   # formatting only
        wb.SaveAs(p, F.XLSX); wb.Close(SaveChanges=False)
        with X.ExcelSession() as xl:
            w = xl.open(p).Worksheets(1)
            self.assertGreater(w.UsedRange.Rows.Count, 900000, "fixture must really inflate UsedRange")
            t0 = time.time()
            self.assertEqual(X.last_cell(w), (3, 3))
            self.assertEqual(sum(len(c) for _, c in X.iter_used_range(w)), 3)
            self.assertLess(time.time() - t0, 10)

    def test_hidden_rows_and_filters_are_still_read(self):
        p = os.path.join(_tmp, "hidden.xlsx")
        wb = _xl.app.Workbooks.Add()
        ws = wb.Worksheets(1)
        ws.Range("A1:B4").Value2 = (("h1", "h2"), (1.0, 2.0), (3.0, 4.0), (5.0, 6.0))
        ws.Rows(2).Hidden = True
        ws.Range("A1:B4").AutoFilter(1, "<>1")
        wb.SaveAs(p, F.XLSX); wb.Close(SaveChanges=False)
        with X.ExcelSession() as xl:
            rows = [r for _, c in X.iter_used_range(xl.open(p).Worksheets(1)) for r in c]
        self.assertEqual(len(rows), 4)
        self.assertEqual(rows[1][:2], (1.0, 2.0), "hidden/filtered-out rows carry data too and must be included")

    def test_error_cells_and_blank_normalisation(self):
        self.assertIsNone(X.clean_cell(-2146826246))                # #N/A
        self.assertIsNone(X.clean_cell(-2146826281))                # #DIV/0!
        self.assertIsNone(X.clean_cell(""))
        self.assertEqual(X.clean_cell(0.0), 0.0)
        self.assertEqual(X.clean_cell(-5), -5)                      # ordinary negative numbers untouched

    def test_macros_force_disabled_and_no_startup_workbook_left_open(self):
        """Defence-in-depth for VBA/PERSONAL.XLSB safety: does not create or touch any global Excel startup
        file (that would itself be exactly the kind of environment change this pipeline must never make) -
        it only checks the session's own state, which is safe to inspect while the session is open."""
        with X.ExcelSession() as xl:
            self.assertEqual(xl.app.AutomationSecurity, X.MSO_AUTOMATION_SECURITY_FORCE_DISABLE,
                             "macros must be force-disabled for the whole automated session")
            self.assertFalse(xl.app.EnableEvents, "Application events (incl. Workbook_Open) must be off")
            self.assertEqual(xl.startup_workbooks_closed, [],
                             "on this machine nothing should have auto-loaded; the attribute must still exist and be empty")
            # our own blank workbook is the only one open, and it is never one we deliberately opened as a source
            self.assertEqual(xl.app.Workbooks.Count, 1)

    def test_no_excel_process_left_behind_even_after_an_exception(self):
        p = os.path.join(_tmp, "small.xlsx")
        _xl.save(p, [("Sheet1", [["a"], [1.0]])], F.XLSX)
        xl = None
        try:
            with X.ExcelSession() as xl:
                self.assertTrue(xl.own_process_alive())
                wb = xl.open(p)
                held = wb.Worksheets(1)                              # live COM refs during the exception
                raise RuntimeError("boom")
        except RuntimeError:
            pass
        self.assertFalse(xl.own_process_alive(), "our own Excel must be gone (other programs' Excels are not our business)")

    def test_watchdog_stops_a_run_that_takes_too_long(self):
        real = r"D:\WORK\Software Development\GitHub\PE BOM TEAM\Final (SEEG) W38 SOP .xlsb"
        if not os.path.exists(real):
            self.skipTest("large real file not available")
        t0 = time.time()
        with self.assertRaises(X.ExcelError) as cm:
            with X.ExcelSession(timeout_s=1) as xl:
                xl.open(real)                                        # takes several seconds -> killed after 1 s
        self.assertIn("did not finish", str(cm.exception))
        self.assertLess(time.time() - t0, 60)


if __name__ == "__main__":
    unittest.main(verbosity=2)
