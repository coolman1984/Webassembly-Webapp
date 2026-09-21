"""Source-file recognition, validation and extraction for the three dashboard workbooks.

Roles (identified by *content*, never by file or sheet name):
  DASH     flat plant-status table            (MODEL_CODE, MKT_PROJECT, PROJECT_NAME, ...)
  NEWMODEL two-row-header BOM plan (Plan/Actual row pairs; BOM HQ/LOCAL, SET PLANT MP)
  SEEG     workbook with a sheet holding Item / Project / Inch / Cate / Version + weekly quantity columns

Columns are located by header text (case, spacing, line breaks and invisible characters are ignored), so
inserting, moving or adding columns is harmless. Header rows are searched for in the first rows of each sheet.
"""
from __future__ import annotations

import re
import unicodedata
from array import array
from collections import Counter
from dataclasses import dataclass, field
from operator import add, itemgetter
from typing import Optional

from . import excel_com as X

ROLES = ("DASH", "NEWMODEL", "SEEG")
WEEK_RE = re.compile(r"^\d{6}$")
HEADER_SCAN_ROWS = 15
PLACEHOLDERS = {"-", "—", "–", "n/a", "na", "tbd", "none", "null", "#n/a", "?", "x"}
_ZERO_WIDTH = dict.fromkeys(map(ord, "​‌‍‎‏⁠﻿­"), None)
_1904_OFFSET = 1462

REQUIRED = {
    "DASH": ["MODEL_CODE", "MKT_PROJECT", "PROJECT_NAME"],
    "SEEG": ["Item", "Project", "Inch", "Cate", "Version", "weekly columns like 202601"],
    "NEWMODEL": ["Project", "Model", "Type", "Inch", "SET PLANT > MP", "BOM > HQ", "BOM > LOCAL"],
}


class ValidationError(ValueError):
    """Raised with a user-readable message when a source file is not what the dashboard expects."""


# ----------------------------------------------------------------------------- text helpers
def clean_str(s: str) -> str:
    s = unicodedata.normalize("NFKC", s).translate(_ZERO_WIDTH)
    return re.sub(r"\s+", " ", s).strip()


def as_text(v) -> str:
    if v is None:
        return ""
    if isinstance(v, bool):
        return str(v).upper()
    if isinstance(v, float):
        return str(int(v)) if v.is_integer() else repr(v)
    return clean_str(str(v))


def norm(v) -> str:
    """Header normaliser: NFKC, invisible chars removed, whitespace collapsed, lower-case."""
    return as_text(v).lower()


def key_model(v) -> str:
    """Join key: numbers stored as float become plain ints, invisible characters/NBSP removed, upper-case."""
    return as_text(v).upper()


def is_blank(v) -> bool:
    return v is None or as_text(v).lower() in PLACEHOLDERS or as_text(v) == ""


# ----------------------------------------------------------------------------- detection
@dataclass
class Detected:
    role: str
    sheet: str
    header_row: int                     # 1-based row of the (first) header line
    columns: dict = field(default_factory=dict)  # logical name -> 0-based column index


@dataclass
class Diagnosis:
    role: str
    sheet: str
    missing: list
    matched: int


def _hdr(row) -> list[str]:
    return [norm(c) for c in row]


def _find_dash(top):
    best = None
    need = {"model": "model_code", "project": "mkt_project", "project_name": "project_name"}
    for i, row in enumerate(top[:HEADER_SCAN_ROWS]):
        h = _hdr(row)
        miss = [v.upper() for v in need.values() if v not in h]
        matched = len(need) - len(miss)
        if not miss:
            return Detected("DASH", "", i + 1, {k: h.index(v) for k, v in need.items()}), None
        if best is None or matched > best.matched:
            best = Diagnosis("DASH", "", miss, matched)
    return None, best


def _week_columns(row) -> list[tuple[str, int]]:
    """(week code, column index) for every header that looks like a week (YYYYWW), in CHRONOLOGICAL order,
    wherever the columns sit - the sheet's left-to-right order must never decide which week is 'first'."""
    out = []
    for j, c in enumerate(row):
        t = as_text(c)
        if WEEK_RE.match(t) and 2000 <= int(t[:4]) <= 2100 and 1 <= int(t[4:]) <= 53:
            out.append((t, j))
    out.sort(key=lambda x: x[0])                 # stable: equal codes keep their sheet order
    return out


def _find_sop(top):
    best = None
    base = ("item", "project", "inch", "cate", "version")
    for i, row in enumerate(top[:HEADER_SCAN_ROWS]):
        h = _hdr(row)
        miss = [n.capitalize() for n in base if n not in h]
        cols = {n: h.index(n) for n in base if n in h}
        wk = _week_columns(row)
        if not wk:
            miss.append("weekly columns like 202601")
        matched = len(base) + 1 - len(miss)
        if not miss:
            cols["weeks"] = [j for _, j in wk]
            cols["week_codes"] = [t for t, _ in wk]
            return Detected("SEEG", "", i + 1, cols), None
        if best is None or matched > best.matched:
            best = Diagnosis("SEEG", "", miss, matched)
    return None, best


def _find_newmodel(top):
    best = None
    base = ("project", "model", "type", "inch")
    for i in range(min(len(top) - 1, HEADER_SCAN_ROWS)):
        a = _hdr(top[i])
        b = _hdr(top[i + 1])
        group, cur = [], ""
        for x in a:
            cur = x or cur
            group.append(cur)

        def find(g, s):
            for j in range(len(a)):
                if group[j] == g and b[j] == s:
                    return j
            return None
        cols = {n: a.index(n) for n in base if n in a}
        combos = {"mp": ("set plant", "mp", "SET PLANT > MP"), "bom_hq": ("bom", "hq", "BOM > HQ"),
                  "bom_local": ("bom", "local", "BOM > LOCAL")}
        miss = [n.capitalize() for n in base if n not in a]
        for k, (g, s, label) in combos.items():
            j = find(g, s)
            if j is None:
                miss.append(label)
            else:
                cols[k] = j
        matched = len(base) + len(combos) - len(miss)
        if not miss:
            return Detected("NEWMODEL", "", i + 1, cols), None
        if best is None or matched > best.matched:
            best = Diagnosis("NEWMODEL", "", miss, matched)
    return None, best


def detect_workbook(wb) -> tuple[Optional[Detected], list[str], Optional[Diagnosis]]:
    """Recognise a workbook from header rows only. Returns (detected, notes, best partial diagnosis)."""
    notes: list[str] = []
    found: list[Detected] = []
    best: Optional[Diagnosis] = None
    for name in X.sheet_names(wb):
        try:
            top = X.read_top(wb.Worksheets(name), nrows=HEADER_SCAN_ROWS + 2)
        except Exception as e:
            notes.append(f"sheet '{name}' could not be read: {e}")
            continue
        if not top:
            continue
        for finder in (_find_sop, _find_dash, _find_newmodel):
            d, diag = finder(top)
            if d:
                d.sheet = name
                found.append(d)
            elif diag and diag.matched >= 2 and (best is None or diag.matched > best.matched):
                diag.sheet = name
                best = diag
    if not found:
        return None, notes, best
    roles = {d.role for d in found}
    if len(roles) > 1:
        raise ValidationError("one workbook contains sheets that look like different sources "
                              f"({', '.join(f'{d.role}: {d.sheet}' for d in found)}); supply each source as its own file.")
    pool = found
    if found[0].role == "SEEG" and len(found) > 1:                      # several SOP-like sheets: prefer one named SOP
        named = [d for d in found if d.sheet.strip().lower() == "sop"]
        if named:
            pool = named
        else:
            raise ValidationError("several sheets look like the SOP data "
                                  f"({', '.join(d.sheet for d in found)}); rename the real one to 'SOP'.")
    elif len(found) > 1:
        notes.append(f"several sheets match {found[0].role} ({', '.join(d.sheet for d in found)}); using '{pool[0].sheet}'")
    return pool[0], notes, best


def explain_unrecognised(name: str, diag: Optional[Diagnosis]) -> str:
    if diag:
        label = {"DASH": "DASH status", "SEEG": "SEEG SOP", "NEWMODEL": "New Model BOM"}[diag.role]
        return (f"'{name}' looks like the {label} file (sheet '{diag.sheet}') but is missing column(s): "
                f"{', '.join(diag.missing)}. Check the header names in that file.")
    return (f"'{name}' is not one of the 3 expected files. Expected: the DASH status file "
            f"({', '.join(REQUIRED['DASH'])}), the SEEG SOP file ({', '.join(REQUIRED['SEEG'])}) "
            f"or the New Model file ({', '.join(REQUIRED['NEWMODEL'])}).")


# ----------------------------------------------------------------------------- extraction
def extract_dash(ws, det: Detected, date1904: bool = False) -> dict:
    header = None
    rows: list[dict] = []
    for start, chunk in X.iter_used_range(ws):
        for k, r in enumerate(chunk):
            srow = start + k
            if srow < det.header_row:
                continue
            if srow == det.header_row:
                header = [as_text(c) or f"Column {i + 1}" for i, c in enumerate(r)]
                continue
            if all(X.clean_cell(c) is None for c in r):
                continue
            rows.append({header[i]: X.clean_cell(r[i]) for i in range(min(len(header), len(r)))})
    m, p, n = det.columns["model"], det.columns["project"], det.columns["project_name"]
    hm, hp, hn = header[m], header[p], header[n]
    out, skipped = [], 0
    for r in rows:
        model = key_model(r.get(hm))
        if not model:
            skipped += 1
            continue
        out.append({"model": model, "project": as_text(r.get(hp)), "projectName": as_text(r.get(hn)), "raw": r})
    return {"header": header, "rows": out, "rows_read": len(rows), "skipped_no_model": skipped}


def _nm_date(v, date1904: bool):
    if is_blank(v):
        return None
    if isinstance(v, float) and date1904 and v > 0:
        return v + _1904_OFFSET
    return v


def extract_newmodel(ws, det: Detected, date1904: bool = False) -> dict:
    c = det.columns
    plan, actual_rows, other_types = [], 0, Counter()
    for start, chunk in X.iter_used_range(ws):
        for k, r in enumerate(chunk):
            if start + k <= det.header_row + 1:
                continue
            model = key_model(r[c["model"]])
            if not model:
                continue
            t = norm(r[c["type"]])
            if t == "plan":
                plan.append({
                    "model": model,
                    "project": as_text(r[c["project"]]),
                    "inch": as_text(r[c["inch"]]),
                    "bom_hq": _nm_date(X.clean_cell(r[c["bom_hq"]]), date1904),
                    "bom_local": _nm_date(X.clean_cell(r[c["bom_local"]]), date1904),
                    "mp": _nm_date(X.clean_cell(r[c["mp"]]), date1904),
                })
            elif t == "actual":
                actual_rows += 1
            else:
                other_types[t or "(blank)"] += 1
    return {"plan": plan, "actual_rows": actual_rows, "other_types": dict(other_types), "date1904": date1904}


def extract_sop(ws, det: Detected, date1904: bool = False) -> dict:
    """Aggregate the SOP sheet to (item, cate, version) -> weekly quantity vector, chunk by chunk."""
    c = det.columns
    weeks = c["weeks"]
    week_codes = c["week_codes"]
    getw = itemgetter(*weeks) if len(weeks) > 1 else (lambda r, j=weeks[0]: (r[j],))
    agg: dict[tuple, array] = {}
    dims: dict[str, dict] = {}
    cates: Counter = Counter()
    rows_read = non_numeric = bad_version = no_item = 0
    ci, cc, cv, cp, cinch = c["item"], c["cate"], c["version"], c["project"], c["inch"]
    for start, chunk in X.iter_used_range(ws):
        for i, r in enumerate(chunk):
            if start + i <= det.header_row:
                continue
            item = key_model(r[ci])
            if not item:
                if any(v is not None for v in r):
                    no_item += 1
                continue
            try:
                ver = int(float(as_text(r[cv])))
            except (TypeError, ValueError):
                bad_version += 1
                continue
            cate = as_text(r[cc]).lower()
            cates[cate] += 1
            raw = getw(r)
            try:
                vec = array("d", [v or 0.0 for v in raw])
                if min(vec) < -1e9:                              # Excel error codes (#N/A ...) -> 0
                    raise TypeError
            except TypeError:
                vals = []
                for v in raw:
                    if isinstance(v, (int, float)) and not isinstance(v, bool) and v > -1e9:
                        vals.append(float(v))
                    else:
                        vals.append(0.0)
                        if v is not None and v != "" and not (isinstance(v, int) and v <= -1e9):
                            non_numeric += 1
                vec = array("d", vals)
            rows_read += 1
            k = (item, cate, ver)
            prev = agg.get(k)
            agg[k] = vec if prev is None else array("d", map(add, prev, vec))
            d = dims.get(item)
            if d is None:
                d = dims[item] = {"ver": -1, "project": "", "inch": "", "any_project": "", "any_inch": ""}
            # SOP inch is numeric in the workbook and the dashboard shows it as '43.0' (BOM inch is text '43')
            proj = as_text(r[cp])
            iv = r[cinch]
            inch = str(iv) if isinstance(iv, float) else as_text(iv)
            if ver >= d["ver"]:                                  # latest version wins; '-' = "no value"
                if proj and proj != "-":
                    d["project"] = proj
                if inch and inch != "-":
                    d["inch"] = inch
                d["ver"] = ver
            if proj and not d["any_project"]:
                d["any_project"] = proj
            if inch and not d["any_inch"]:
                d["any_inch"] = inch
    if not agg:
        raise ValidationError("the SOP sheet has a valid header but no data rows")
    dup_weeks = [k for k, n in Counter(week_codes).items() if n > 1]
    return {"week_codes": week_codes, "agg": agg, "dims": dims, "rows_read": rows_read, "non_numeric": non_numeric,
            "bad_version": bad_version, "cates": dict(cates), "rows_without_item": no_item, "dup_weeks": dup_weeks}
