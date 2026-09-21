"""Stress tests: data volumes well beyond the real files, up to Excel's hard row limit.

    runtime\\python.exe -I tests\\test_stress.py -v        (takes ~15 min the first time; workbooks are cached)

Expectations are computed here independently (plain datetime maths), never by calling the pipeline's own helpers.
"""
from __future__ import annotations

import ctypes
import datetime as dt
import os
import random
import sys
import tempfile
import time
import unittest
from ctypes import wintypes

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
from pipeline import excel_com as X          # noqa: E402
from pipeline import run, store              # noqa: E402
from pipeline import sources as S            # noqa: E402
from tests import fixtures as F              # noqa: E402

CACHE = os.path.join(tempfile.gettempdir(), "bom_stress_cache_v1")
N_ITEMS, N_NM, N_DASH_EXTRA = 25000, 20000, 35000
VERSIONS = list(range(202629, 202639))                       # 10 versions per item, x2 categories = 500k SOP rows
WEEKS = [f"2026{w:02d}" for w in range(1, 53)] + [f"2027{w:02d}" for w in range(1, 9)]     # 60 week columns


def peak_mb() -> float:
    class PMC(ctypes.Structure):
        _fields_ = [("cb", wintypes.DWORD), ("PageFaultCount", wintypes.DWORD), ("PeakWorkingSetSize", ctypes.c_size_t),
                    ("WorkingSetSize", ctypes.c_size_t), ("QuotaPeakPagedPoolUsage", ctypes.c_size_t),
                    ("QuotaPagedPoolUsage", ctypes.c_size_t), ("QuotaPeakNonPagedPoolUsage", ctypes.c_size_t),
                    ("QuotaNonPagedPoolUsage", ctypes.c_size_t), ("PagefileUsage", ctypes.c_size_t),
                    ("PeakPagefileUsage", ctypes.c_size_t)]
    k32, ps = ctypes.WinDLL("kernel32"), ctypes.WinDLL("psapi")
    k32.GetCurrentProcess.restype = wintypes.HANDLE                     # 64-bit pseudo-handle: default int would truncate it
    ps.GetProcessMemoryInfo.argtypes = [wintypes.HANDLE, ctypes.POINTER(PMC), wintypes.DWORD]
    ps.GetProcessMemoryInfo.restype = wintypes.BOOL
    c = PMC(); c.cb = ctypes.sizeof(c)
    ok = ps.GetProcessMemoryInfo(k32.GetCurrentProcess(), ctypes.byref(c), c.cb)
    assert ok and c.PeakWorkingSetSize > 0, "memory probe failed"
    return c.PeakWorkingSetSize / 1048576


def first_week_index(i: int) -> int:
    return (i * 7) % 52                                       # the first positive week of item i (0..51, inside 2026)


def item(i: int) -> str:
    return f"S{i:06d}"


# ------------------------------------------------------------------------------------------------ generators
def _write_chunked(ws, rows_iter, ncols, chunk=40000, start_row=1):
    r = start_row
    buf = []

    def flush():
        ws.Range(ws.Cells(r, 1), ws.Cells(r + len(buf) - 1, ncols)).Value2 = tuple(tuple(x) for x in buf)

    for row in rows_iter:
        buf.append(row)
        if len(buf) >= chunk:
            flush()
            r += len(buf); buf = []
    if buf:
        flush()
        r += len(buf)
    return r - 1


def _sop_rows():
    zero = [0.0] * len(WEEKS)
    for i in range(N_ITEMS):
        k = first_week_index(i)
        for v in VERSIONS:
            for cate in ("Ship", "SOP"):
                q = list(zero)
                kk = k if v == VERSIONS[-1] else min(k + 1, len(WEEKS) - 1)      # older versions: later first week
                q[kk] = 10.0
                yield [item(i), "S1", "P%d" % (i % 50), 50.0, cate, float(v)] + q


def build_big(xl: F.Excel, d: str) -> dict:
    os.makedirs(d, exist_ok=True)
    files = {"dash": os.path.join(d, "(DASH)big.xlsx"), "sop": os.path.join(d, "Big SOP.xlsb"), "nm": os.path.join(d, "New Model big.xlsx")}
    if all(os.path.exists(p) for p in files.values()):
        return files
    app = xl.app
    # ---- SOP (xlsb): 500,000 rows x 66 columns
    wb = app.Workbooks.Add(); ws = wb.Worksheets(1); ws.Name = "SOP"
    hdr = ["Item", "Site(To)", "Project", "Inch", "Cate", "Version"] + WEEKS
    ws.Range(ws.Cells(1, 1), ws.Cells(1, len(hdr))).Value2 = (tuple(hdr),)
    _write_chunked(ws, _sop_rows(), len(hdr), start_row=2)
    wb.SaveAs(files["sop"], F.XLSB); wb.Close(SaveChanges=False)
    # ---- DASH: 60,000 rows
    wb = app.Workbooks.Add(); ws = wb.Worksheets(1); ws.Name = "Sheet1"
    h = ("PRODUCT_TYPE", "PRODUCT", "MKT_PROJECT", "YEAR", "PROJECT_NAME", "MODEL_CODE", "NATION")
    ws.Range("A1:G1").Value2 = (h,)
    rows = ((("TV", "LED", "MKT%d" % (i % 90), "2026", "Name %d" % i, item(i) if i < N_ITEMS else f"D{i:06d}", "TESTLANDIA")) for i in range(N_ITEMS + N_DASH_EXTRA))
    _write_chunked(ws, rows, 7, start_row=2)
    wb.SaveAs(files["dash"], F.XLSX); wb.Close(SaveChanges=False)
    # ---- NEW MODEL: 20,000 models, Plan/Actual pairs, two header rows
    wb = app.Workbooks.Add(); ws = wb.Worksheets(1); ws.Name = "Sheet1"
    ws.Range("A1:H2").Value2 = (("Project", "Model", "Type", "Inch", "Plant", "SET PLANT", "BOM", None), (None, None, None, None, None, "MP", "HQ", "LOCAL"))
    def nm_rows():
        for i in range(N_NM):
            k = first_week_index(i)
            mp = dt.date.fromisocalendar(2026, k + 1, 1)
            hq = mp - dt.timedelta(days=91) + dt.timedelta(days=7 * (i % 2))
            lo = mp - dt.timedelta(days=84) + dt.timedelta(days=7 * ((i % 3) - 1))
            yield ("P%d" % (i % 50), item(i), "Plan", "50", "SEEG", 46000.0, hq.strftime("%Y/%m/%d"), lo.strftime("%Y/%m/%d"))
            yield ("P%d" % (i % 50), item(i), "Actual", "50", "SEEG", None, "OK", "OK")
    _write_chunked(ws, nm_rows(), 8, start_row=3)
    wb.SaveAs(files["nm"], F.XLSX); wb.Close(SaveChanges=False)
    return files


class TestBigData(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        t0 = time.time()
        with F.Excel() as xl:
            cls.files = build_big(xl, CACHE)
        print(f"\n[fixtures ready in {time.time() - t0:.0f}s: SOP {os.path.getsize(cls.files['sop']) / 1e6:.0f} MB]", flush=True)

    def test_500k_row_sop_60k_models(self):
        t0 = time.time()
        ext = run.extract_all([self.files["nm"], self.files["sop"], self.files["dash"]], log=lambda m: None)
        t_extract = time.time() - t0
        t1 = time.time()
        bom, master, rep = run.build_datasets(ext)
        t_build = time.time() - t1
        print(f"\n  extract {t_extract:.0f}s | transform {t_build:.0f}s | peak memory {peak_mb():.0f} MB | "
              f"{ext['sop']['rows_read']:,} SOP rows -> {len(bom):,} BOM / {len(master):,} models", flush=True)

        self.assertEqual(ext["sop"]["rows_read"], N_ITEMS * len(VERSIONS) * 2)
        self.assertEqual(len(bom), N_NM)
        self.assertEqual(len(master), N_ITEMS + N_DASH_EXTRA)               # DASH already lists every SOP item
        self.assertEqual(sum(1 for r in master if "SOP" in r["coverage"]), N_ITEMS)
        self.assertEqual(sum(1 for r in master if "BOM" in r["coverage"]), N_NM)
        by = {r["model"]: r for r in bom}
        rng = random.Random(7)
        for i in rng.sample(range(N_NM), 400) + [0, 1, 2, N_NM - 1]:
            r = by[item(i)]
            mp = dt.date.fromisocalendar(2026, first_week_index(i) + 1, 1)
            self.assertEqual(r["mp"], "W%02d / 2026" % (first_week_index(i) + 1), item(i))
            self.assertEqual(r["mpDate"], mp.strftime("%d %b %Y"))
            self.assertEqual(r["hqStatus"], "MATCH" if i % 2 == 0 else "LATER", item(i))
            self.assertEqual(r["localStatus"], {0: "EARLIER", 1: "MATCH", 2: "LATER"}[i % 3], item(i))
            self.assertEqual(r["version"], VERSIONS[-1])
            first = dt.date.fromisocalendar(2026, VERSIONS[0] % 100, 1)
            self.assertEqual(r["firstSop"], f"Version {VERSIONS[0]}")
            self.assertEqual(r["firstAppearMpGap"], f"{(mp - first).days // 7} weeks")
        self.assertLess(t_extract, 600, "extraction of 500k rows should stay within 10 minutes")
        self.assertLess(peak_mb(), 3500)

        # the database layer and the served JSON at this size
        d = tempfile.mkdtemp(prefix="bigdb_")
        t2 = time.time()
        store.write_db(os.path.join(d, "x.db"), ext, bom, master, rep)
        t_db = time.time() - t2
        t3 = time.time()
        data = store.read_datasets(os.path.join(d, "x.db"))
        import json
        body = json.dumps(data, ensure_ascii=False).encode()
        print(f"  db write {t_db:.0f}s | read+json {time.time() - t3:.1f}s | payload {len(body) / 1e6:.0f} MB", flush=True)
        self.assertEqual(len(data["master"]), len(master))
        self.assertEqual(data["bom"][0]["model"], bom[0]["model"])

    def test_excel_row_limit_1048576_rows(self):
        """A sheet filled to the last row Excel allows; header + 1,048,575 data rows."""
        p = os.path.join(CACHE, "rowlimit.xlsb")
        with F.Excel() as xl:
            if not os.path.exists(p):
                wb = xl.app.Workbooks.Add(); ws = wb.Worksheets(1); ws.Name = "SOP"
                ws.Range("A1:I1").Value2 = (("Item", "Project", "Inch", "Cate", "Version", "202640", "202641", "202642", "202643"),)
                def rows():
                    for n in range(1048575):
                        yield (f"M{n:07d}", "PRJ", 50.0, "Ship" if n % 2 else "SOP", 202638.0, 0.0, 5.0 if n % 3 == 0 else 0.0, 0.0, 1.0)
                last = _write_chunked(ws, rows(), 9, chunk=100000, start_row=2)
                self.assertEqual(last, 1048576)
                wb.SaveAs(p, F.XLSB); wb.Close(SaveChanges=False)
        t0 = time.time()
        with X.ExcelSession() as xl:
            wb = xl.open(p); ws = wb.Worksheets("SOP")
            self.assertEqual(X.last_cell(ws), (1048576, 9))
            det, _, _ = S.detect_workbook(wb)
            sop = S.extract_sop(ws, det)
        print(f"\n  1,048,575 rows read in {time.time() - t0:.0f}s | peak memory {peak_mb():.0f} MB", flush=True)
        self.assertEqual(sop["rows_read"], 1048575)
        self.assertEqual(len(sop["agg"]), 1048575)
        self.assertEqual(sop["cates"], {"ship": 524287, "sop": 524288})


if __name__ == "__main__":
    unittest.main(verbosity=2)
