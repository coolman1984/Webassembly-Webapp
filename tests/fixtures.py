"""Synthetic workbook generator for the test-suite. Everything is written through real Excel (COM), so the
pipeline is exercised against genuine .xlsx/.xlsb files, not mocks.

The base data set is tiny and hand-checkable (see EXPECTED_BOM). Options perturb *how it is stored*
(column order, header noise, extra columns, 1904 dates, odd model codes, ...) without changing what it means,
so every perturbation must yield exactly the same dashboard data.
"""
from __future__ import annotations

import os
import random
import sys

import pythoncom
import win32com.client

XLSX, XLSB = 51, 50
NBSP, ZWSP = " ", "​"
SERIAL = {"2026-11-09": 46335, "2027-01-04": 46391, "2026-10-12": 46307}     # 1900 system serials
WEEKS = ["202640", "202641", "(+)Oct", "202642", "202643", "202644", "202645", "202646",
         "202650", "202651", "202652", "202653", "202701"]
WEEKCOL = {w: i for i, w in enumerate(WEEKS)}

# ------------------------------------------------------------------------------------------ expectations
_BASE = ("project", "inch", "bomHQ", "bomHQDate", "bomLocal", "bomLocalDate", "version", "mp", "mpDate", "hqTarget",
         "hqTargetDate", "localTarget", "localTargetDate", "hqStatus", "localStatus", "hqDelta", "localDelta",
         "firstSop", "firstAppearMpGap")


def _row(project, inch, hq, hqd, lo, lod, ver, mp, mpd, ht, htd, lt, ltd, hs, ls, dh, dl, fs, gap):
    return dict(zip(_BASE, (project, inch, hq, hqd, lo, lod, ver, mp, mpd, ht, htd, lt, ltd, hs, ls, dh, dl, fs, gap)))


EXPECTED_BOM = {
    "TV-A1": _row("PRJ-A", "50", "W29 / 2026", "13 Jul 2026", "W31 / 2026", "27 Jul 2026", 202638, "W42 / 2026", "12 Oct 2026",
                  "W29 / 2026", "13 Jul 2026", "W30 / 2026", "20 Jul 2026", "MATCH", "LATER", 0, 1, "Version 202630", "12 weeks"),
    "TV-B2": _row("PRJ-B", "55", "W31 / 2026", "27 Jul 2026", "—", "—", 202636, "W45 / 2026", "02 Nov 2026",
                  "W32 / 2026", "03 Aug 2026", "W33 / 2026", "10 Aug 2026", "EARLIER", "N/A", -1, None, "Version 202635", "10 weeks"),
    "TV-F6": _row("PRJ-F", "32", "W33 / 2026", "10 Aug 2026", "W33 / 2026", "10 Aug 2026", "", "W46 / 2026", "09 Nov 2026",
                  "W33 / 2026", "10 Aug 2026", "W34 / 2026", "17 Aug 2026", "MATCH", "EARLIER", 0, -1, "—", "—"),
    "12345": _row("PRJ-G", "43", "W33 / 2026", "10 Aug 2026", "W34 / 2026", "17 Aug 2026", 202638, "W46 / 2026", "09 Nov 2026",
                  "W33 / 2026", "10 Aug 2026", "W34 / 2026", "17 Aug 2026", "MATCH", "MATCH", 0, 0, "Version 202638", "8 weeks"),
    "TV-H8": _row("PRJ-H", "75", "W42 / 2026", "12 Oct 2026", "W42 / 2026", "12 Oct 2026", 202638, "W01 / 2027", "04 Jan 2027",
                  "W41 / 2026", "05 Oct 2026", "W42 / 2026", "12 Oct 2026", "LATER", "MATCH", 1, 0, "Version 202638", "16 weeks"),
}
# added by build(new_model=True): a model that first appears in a later refresh
EXPECTED_NEW = _row("PRJ-N", "65", "W30 / 2026", "20 Jul 2026", "W31 / 2026", "27 Jul 2026", 202638, "W43 / 2026", "19 Oct 2026",
                    "W30 / 2026", "20 Jul 2026", "W31 / 2026", "27 Jul 2026", "MATCH", "MATCH", 0, 0, "Version 202638", "5 weeks")
EXPECTED_COVERAGE = {"TV-A1": "DASH + SOP + BOM", "TV-B2": "DASH + SOP + BOM", "TV-C3": "DASH", "TV-D4": "SOP",
                     "TV-F6": "DASH + BOM", "12345": "DASH + SOP + BOM", "TV-H8": "DASH + SOP + BOM"}
BOM_ORDER = sorted(EXPECTED_BOM)


# ------------------------------------------------------------------------------------------ helpers
def T(s: str) -> str:
    """Force a cell to be stored as text (e.g. '12345' must not become a number)."""
    return "'" + s


def noisy_header(s: str, rng: random.Random) -> str:
    s = rng.choice([s, s.upper(), s.lower(), s.title()])
    if " " in s:
        s = s.replace(" ", rng.choice([" ", "  ", NBSP, "\n", " \n"]))
    return rng.choice(["", " ", NBSP]) + s + rng.choice(["", " ", ZWSP, NBSP])


def noisy_model(s: str, rng: random.Random) -> str:
    """Same model, different cosmetics: case, NBSP/space padding, zero-width chars."""
    if s.isdigit():
        return s
    s = rng.choice([s, s.lower(), s.title()])
    return rng.choice(["", " ", NBSP]) + s + rng.choice(["", " ", NBSP, ZWSP])


class Excel:
    """Fixture-side Excel instance (separate from the pipeline's ExcelSession)."""

    def __enter__(self):
        pythoncom.CoInitialize()
        self.app = win32com.client.DispatchEx("Excel.Application")
        self.app.Visible = False
        self.app.DisplayAlerts = False
        self.app.ScreenUpdating = False
        self.app.EnableEvents = False
        return self

    def __exit__(self, *a):
        # no CoUninitialize: it would tear down the whole thread apartment and disconnect every other live Excel
        try:
            self.app.Quit()
        finally:
            self.app = None

    def save(self, path, sheets, fmt=XLSX, date1904=False, password=None, hidden=(), formulas=None):
        """sheets: list of (name, rows). rows are lists (None = empty cell, str, float)."""
        wb = self.app.Workbooks.Add()
        try:
            if date1904:
                wb.Date1904 = True
            while wb.Worksheets.Count < len(sheets):
                wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
            for i in range(1, len(sheets) + 1):                      # two-phase rename avoids name clashes
                wb.Worksheets(i).Name = f"_tmp{i}"
            for i, (name, rows) in enumerate(sheets, start=1):
                ws = wb.Worksheets(i)
                ws.Name = name
                if rows:
                    width = max(len(r) for r in rows)
                    data = tuple(tuple((list(r) + [None] * (width - len(r)))) for r in rows)
                    ws.Range(ws.Cells(1, 1), ws.Cells(len(rows), width)).Value2 = data
                for (r, c, f) in (formulas or {}).get(name, []):
                    ws.Cells(r, c).Formula = f
            for h in hidden:
                wb.Worksheets(h).Visible = 0
            if os.path.exists(path):
                os.remove(path)
            if password:
                wb.SaveAs(path, fmt, password)
            else:
                wb.SaveAs(path, fmt)
        finally:
            wb.Close(SaveChanges=False)


# ------------------------------------------------------------------------------------------ base tables
def _tables(o: dict):
    """Column-oriented base data: {'dash': [(header, [values])], 'sop': [...], 'nm': [(group, sub, [values])]}."""
    m = (lambda s, rng: noisy_model(s, rng)) if o.get("model_noise") else (lambda s, rng: s)
    rng = random.Random(o.get("seed", 1) * 7919)
    num_item = o.get("numeric_item", True)

    # ---- DASH
    dm = ["TV-A1", "TV-B2", "TV-C3", "12345", "TV-H8", "TV-F6"]
    proj = ["MKT-A", "MKT-B", "MKT-C", "MKT-G", "MKT-H", "MKT-F"]
    pname = ["Name A", "Name B", None, "Name G", "Name H", "Name F"]
    if o.get("dup_dash"):                                   # same model twice: first non-empty value must win
        dm, proj, pname = dm + ["TV-A1"], proj + ["OTHER"], pname + ["Other name"]
    if o.get("new_model"):
        dm, proj, pname = dm + ["TV-N9"], proj + ["MKT-N"], pname + ["Name N"]
    nd = len(dm)
    dash = [
        ("PRODUCT_TYPE", ["TV"] * nd), ("PRODUCT", ["LED"] * nd), ("MKT_PROJECT", proj),
        ("YEAR", ["2026"] * nd), ("PROJECT_NAME", pname),
        ("MODEL_CODE", [T(x) if x.isdigit() else m(x, rng) for x in dm]),
        ("NATION", ["TESTLANDIA"] * nd), ("HQ_CONFIRM_DATE", ["2026-06-05"] * nd),
    ]

    # ---- SOP: (item, cate, version, project, inch, {week: qty})
    def rec(item, cate, ver, proj, inch, **q):
        return (item, cate, ver, proj, inch, {k[1:]: v for k, v in q.items()})
    sop = [
        rec("TV-A1", "Ship", 202638, "PRJ-A", 50.0, w202642=50.0),
        rec("TV-A1", "Ship", 202638, "PRJ-A", 50.0, w202644=7.0),                 # 2nd destination row -> summed
        rec("TV-A1", "SOP", 202630, "PRJ-A", 50.0, w202643=10.0, w202646="=NA()"),   # earliest SOP version + error cell
        rec("TV-A1", "SOP", 202638, "PRJ-A", 50.0, w202643=10.0),
        rec("tv-b2 ", "SOP", 202635, "PRJ-B", 55.0, w202644=4.0),                    # lower-case + trailing space in SOP
        rec("tv-b2 ", "SOP", 202636, "PRJ-B", 55.0, w202645=4.0),
        rec("TV-D4", "Ship", 202638, "-", 65.0, w202650=9.0),
        rec("TV-E5", "SOP", 202638, "PRJ-E", 40.0, w202642=0.0),                     # all zero -> out of scope
        rec("TV-H8", "Ship", 202638, "PRJ-H", 75.0, w202701=30.0),
        rec("TV-H8", "SOP", 202638, "PRJ-H", 75.0, w202701=30.0),
        rec(12345.0 if num_item else "12345", "Ship", 202638, "PRJ-G", 43.0, w202646=20.0),   # numeric item code
        rec(12345.0 if num_item else "12345", "SOP", 202638, "PRJ-G", 43.0, w202646=20.0),
    ]
    if o.get("unknown_cate"):
        sop.append(rec("TV-A1", "Forecast", 202638, "PRJ-A", 50.0, w202642=1.0))
    if o.get("new_model"):
        sop += [rec("TV-N9", "Ship", 202638, "PRJ-N", 65.0, w202643=15.0), rec("TV-N9", "SOP", 202638, "PRJ-N", 65.0, w202643=15.0)]
    if o.get("model_noise"):
        sop = [(m(i, rng) if isinstance(i, str) and i.strip().lower() != "tv-b2" else i,) + r[1:] for r in sop for i in [r[0]]]
    cols = ["Item", "Site(To)", "Project", "Inch", "Cate", "Version"]
    sopcols = [
        ("Item", [r[0] for r in sop]), ("Site(To)", ["S1"] * len(sop)), ("Project", [r[3] for r in sop]),
        ("Inch", [r[4] for r in sop]), ("Cate", [r[1] for r in sop]), ("Version", [float(r[2]) for r in sop]),
    ]
    for w in WEEKS:
        vals = []
        for r in sop:
            v = r[5].get(w)
            vals.append(v if v is not None else (0.0 if not w.startswith("(") else None))
        # 'bad_week': 2027 has only 52 ISO weeks, so 202753 is not a real week
        sopcols.append(("202753" if (o.get("bad_week") and w == "202652") else w, vals))

    # ---- NEW MODEL  (group, sub, values) - two header rows, Plan/Actual pairs
    nmm = ["TV-A1", "TV-B2", "TV-F6", "12345", "TV-H8"]
    plan = {
        "TV-A1": ("PRJ-A", "50", 46300.0, "2026/07/13", "2026/07/27"),
        "TV-B2": ("PRJ-B", "55", 46300.0, "2026-07-27", "-"),
        "TV-F6": ("PRJ-F", "32", float(SERIAL["2026-11-09"]), "2026/08/10", "2026/08/10"),
        "12345": ("PRJ-G", "43", 46300.0, "2026/08/10", "2026/08/17"),
        "TV-H8": ("PRJ-H", "75", 46300.0, "2026/10/12", "2026/10/12"),
    }
    if o.get("new_model"):
        nmm = nmm + ["TV-N9"]
        plan["TV-N9"] = ("PRJ-N", "65", 46300.0, "2026/07/20", "2026/07/27")
    n = len(nmm) * 2
    def col(fn):
        out = []
        for mm in nmm:
            out.append(fn(mm, "Plan")); out.append(fn(mm, "Actual"))
        return out
    nm = [
        ("Project", "", col(lambda mm, t: plan[mm][0])), ("Model", "", col(lambda mm, t: T(mm) if mm.isdigit() else m(mm, rng))),
        ("Type", "", col(lambda mm, t: t)), ("Inch", "", col(lambda mm, t: T(plan[mm][1]))),
        ("Plant", "", ["TESTPLANT"] * n),
        ("SET PLANT", "MP", col(lambda mm, t: plan[mm][2] if t == "Plan" else None)),
        ("BOM", "HQ", col(lambda mm, t: plan[mm][3] if t == "Plan" else "OK")),
        ("BOM", "LOCAL", col(lambda mm, t: plan[mm][4] if t == "Plan" else "OK")),
    ]
    return dash, sopcols, nm


def _mp_shift(nm, date1904):
    """1904 workbooks store the same date as a serial that is 1462 smaller."""
    if not date1904:
        return nm
    out = []
    for g, s, v in nm:
        if (g, s) == ("SET PLANT", "MP"):
            v = [x - 1462 if isinstance(x, float) else x for x in v]
        out.append((g, s, v))
    return out


def build(dirpath: str, xl: Excel, **o) -> dict:
    """Write DASH/SOP/NEWMODEL workbooks into dirpath. Options (all optional):
       seed, perm (shuffle columns), extra (insert junk columns), hdr_noise, title_rows, sop_name, date1904,
       model_noise, numeric_item, sop_fmt, nm_explicit_groups, extra_sheets, hidden_first_sheet."""
    os.makedirs(dirpath, exist_ok=True)
    rng = random.Random(o.get("seed", 1))
    dash, sop, nm = _tables(o)
    nm = _mp_shift(nm, o.get("date1904", False))

    def shape(cols, two_row=False):
        cols = list(cols)
        if o.get("extra"):
            for k in range(o["extra"]):
                junk = ("Junk%d" % k, "", ["x%d" % k] * len(cols[0][-1])) if two_row else ("Junk%d" % k, ["j"] * len(cols[0][1]))
                cols.insert(rng.randrange(len(cols) + 1), junk)
        if o.get("perm"):
            rng.shuffle(cols)
        return cols

    def hdr(s):
        s = o.get("rename", {}).get(s, s) if s else s
        return noisy_header(s, rng) if o.get("hdr_noise") and s else s

    def title(rows_head, width):
        return [["Confidential report %d" % i] + [None] * (width - 1) for i in range(o.get("title_rows", 0))] + rows_head

    # DASH
    dc = shape(dash)
    nrows = len(dc[0][1])
    drows = title([[hdr(h) for h, _ in dc]], len(dc)) + [[v[i] for _, v in dc] for i in range(nrows)]
    # SOP
    sc = shape(sop)
    srows = title([[hdr(h) if not h.isdigit() and not h.startswith("(") else h for h, _ in sc]], len(sc)) + \
        [[v[i] for _, v in sc] for i in range(len(sc[0][1]))]
    # New Model: two header rows
    nc = shape(nm, two_row=True)
    if o.get("nm_explicit_groups") or o.get("perm") or o.get("extra"):
        groups = [hdr(g) for g, _, _ in nc]
    else:                                                    # like the real file: a group label only on its first column
        groups, prev = [], None
        for g, _, _ in nc:
            groups.append(hdr(g) if g != prev else None); prev = g
    subs = [hdr(s) if s else None for _, s, _ in nc]
    nrows_nm = len(nc[0][2])
    nrows_all = title([groups, subs], len(nc)) + [[v[i] for _, _, v in nc] for i in range(nrows_nm)]

    files = {
        "dash": os.path.join(dirpath, "(DASH)PLANT_BOM_STATUS_test.xlsx"),
        "sop": os.path.join(dirpath, "Final SEEG SOP test." + ("xlsb" if o.get("sop_fmt") == "xlsb" else "xlsx")),
        "nm": os.path.join(dirpath, "New Model_test.xlsx"),
    }
    extra = [("Notes", [["nothing to see"]])] if o.get("extra_sheets") else []
    xl.save(files["dash"], [("Sheet1", drows)] + extra, XLSX, formulas=o.get("dash_formulas"))
    formulas = {}
    sop_name = o.get("sop_name", "SOP")
    # error cell: put =NA() where the fixture asked for it
    for r_i, row in enumerate(srows):
        for c_i, v in enumerate(row):
            if v == "=NA()":
                srows[r_i][c_i] = None
                formulas.setdefault(sop_name, []).append((r_i + 1, c_i + 1, "=NA()"))
    xl.save(files["sop"], [(sop_name, srows)] + extra, XLSB if o.get("sop_fmt") == "xlsb" else XLSX,
            hidden=[sop_name] if o.get("hidden_first_sheet") and extra else (), formulas=formulas)
    xl.save(files["nm"], [("Sheet1", nrows_all)], XLSX, date1904=o.get("date1904", False))
    return files
