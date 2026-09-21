"""Business calculations: turn the three extracted sources into the dashboard datasets.

Every rule here was reverse-engineered from - and verified against - the dataset embedded in the original
dashboard (see verify.py):

  version         latest SOP-sheet version the model appears in (any category)
  firstSop        earliest version in which the model appears with category 'SOP'  -> "Version 2026NN"
  mp              first week with qty > 0 in the model's latest 'Ship' version; if that gives nothing,
                  the same on its latest 'SOP' version; if still nothing, the New Model 'MP' date
  hqTarget        MP - 13 weeks          localTarget   MP - 12 weeks   (ISO weeks, Monday dates)
  hq/localDelta   whole weeks between the BOM HQ/LOCAL plan week and the target week
  hq/localStatus  MATCH (0) / LATER (>0) / EARLIER (<0); N/A when there is no BOM row
  firstAppearMpGap  weeks between the firstSop version week and the MP week -> "N weeks"
  coverage        DASH / SOP / BOM parts joined with ' + '
"""
from __future__ import annotations

import datetime as dt
import re
from collections import OrderedDict, defaultdict
from typing import Optional

from .sources import as_text, is_blank

HQ_WEEKS = 13
LOCAL_WEEKS = 12
DASH_CH = "—"
_MON = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
_EPOCH = dt.date(1899, 12, 30)


# ----------------------------------------------------------------------------- date helpers
_ISO_LIKE = re.compile(r"^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})(?:[T\s].*)?$")
_COMPACT = re.compile(r"^(\d{4})(\d{2})(\d{2})$")
_DMY_MON = re.compile(r"^(\d{1,2})[-\s]([A-Za-z]{3})[a-z]*[-\s,]*(\d{4})$")


def parse_date(v) -> Optional[dt.date]:
    """Excel serial or unambiguous text date -> date. Placeholders ('-', 'TBD', ...) and ambiguous formats
    such as 06/07/2026 give None (never guessed)."""
    if is_blank(v):
        return None
    try:
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            return _EPOCH + dt.timedelta(days=int(v)) if 30000 < v < 80000 else None   # ~1982 - 2119
        if isinstance(v, dt.datetime):
            return v.date()
        if isinstance(v, dt.date):
            return v
        s = as_text(v)
        m = _ISO_LIKE.match(s) or _COMPACT.match(s)
        if m:
            return dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        m = _DMY_MON.match(s)
        if m and m.group(2).title() in _MON:
            return dt.date(int(m.group(3)), _MON.index(m.group(2).title()) + 1, int(m.group(1)))
    except (ValueError, OverflowError):
        return None
    return None


def monday(d: dt.date) -> dt.date:
    return d - dt.timedelta(days=d.weekday())


def week_label(d: Optional[dt.date]) -> str:
    if not d:
        return DASH_CH
    y, w, _ = d.isocalendar()
    return f"W{w:02d} / {y}"


def fmt_date(d: Optional[dt.date]) -> str:
    return f"{d.day:02d} {_MON[d.month - 1]} {d.year}" if d else DASH_CH


def week_code_to_monday(code: str) -> Optional[dt.date]:
    try:
        return dt.date.fromisocalendar(int(code[:4]), int(code[4:]), 1)
    except (ValueError, TypeError):
        return None


def version_to_monday(ver: int) -> Optional[dt.date]:
    return week_code_to_monday(str(int(ver)))


def weeks_between(later: dt.date, earlier: dt.date) -> int:
    return (monday(later) - monday(earlier)).days // 7


# ----------------------------------------------------------------------------- SOP summary
def summarise_sop(sop: dict, max_version: Optional[int] = None) -> dict[str, dict]:
    """Per item: latest version, first SOP version, MP week code, whether it ever has qty>0."""
    week_codes = sop["week_codes"]
    by_item: dict[str, list] = defaultdict(list)
    for (item, cate, ver), vec in sop["agg"].items():
        if max_version is not None and ver > max_version:
            continue
        by_item[item].append((cate, ver, vec))

    def first_pos(vec) -> Optional[str]:
        for j, x in enumerate(vec):
            if x > 0:
                return week_codes[j]
        return None

    out = {}
    for item, rows in by_item.items():
        latest: dict[str, tuple] = {}
        for cate, ver, vec in rows:
            if cate not in latest or ver > latest[cate][0]:
                latest[cate] = (ver, vec)
        mp_code, mp_from = None, None
        for cate in ("ship", "sop"):
            if cate in latest:
                mp_code = first_pos(latest[cate][1])
                if mp_code:
                    mp_from = cate
                    break
        sop_versions = [ver for cate, ver, _ in rows if cate == "sop"]
        out[item] = {
            "version": max(ver for _, ver, _ in rows),
            "first_sop_version": min(sop_versions) if sop_versions else None,
            "mp_code": mp_code,
            "mp_from": mp_from,
            "has_qty": any(sum(vec) > 0 for _, _, vec in rows),
        }
    return out


def invalid_week_codes(sop: dict) -> list[str]:
    """Week headers that are not a real ISO week (e.g. 202753 - 2027 has only 52 weeks)."""
    return [c for c in sop["week_codes"] if week_code_to_monday(c) is None]


# ----------------------------------------------------------------------------- row builders
def _status(delta: Optional[int]) -> str:
    if delta is None:
        return "N/A"
    return "MATCH" if delta == 0 else ("LATER" if delta > 0 else "EARLIER")


def _mp_fields(mp_day: Optional[dt.date]) -> dict:
    hq_t = mp_day - dt.timedelta(weeks=HQ_WEEKS) if mp_day else None
    lo_t = mp_day - dt.timedelta(weeks=LOCAL_WEEKS) if mp_day else None
    return {
        "mp": week_label(mp_day), "mpDate": fmt_date(mp_day),
        "hqTarget": week_label(hq_t), "hqTargetDate": fmt_date(hq_t),
        "localTarget": week_label(lo_t), "localTargetDate": fmt_date(lo_t),
    }, hq_t, lo_t


def build_bom(nm: dict, sop_sum: dict[str, dict]) -> tuple[list[dict], dict]:
    rows: "OrderedDict[str, dict]" = OrderedDict()
    stats = {"mp_from_sop": 0, "mp_from_newmodel": 0, "mp_missing": 0, "bad_dates": 0, "duplicates": 0, "mp_bad_week": 0}
    for p in nm["plan"]:
        model = p["model"]
        if model in rows:
            stats["duplicates"] += 1
            continue
        s = sop_sum.get(model)
        mp_day = week_code_to_monday(s["mp_code"]) if s and s["mp_code"] else None
        if s and s["mp_code"] and not mp_day:
            stats["mp_bad_week"] += 1
        if mp_day:
            stats["mp_from_sop"] += 1
        else:
            mp_day = parse_date(p["mp"])
            stats["mp_from_newmodel" if mp_day else "mp_missing"] += 1
        hq, lo = parse_date(p["bom_hq"]), parse_date(p["bom_local"])
        if (not is_blank(p["bom_hq"]) and not hq) or (not is_blank(p["bom_local"]) and not lo):
            stats["bad_dates"] += 1
        mpf, hq_t, lo_t = _mp_fields(mp_day)
        dh = weeks_between(hq, hq_t) if hq and hq_t else None
        dl = weeks_between(lo, lo_t) if lo and lo_t else None
        fs = s["first_sop_version"] if s else None
        gap = DASH_CH
        fs_day = version_to_monday(fs) if fs else None
        if fs_day and mp_day:
            gap = f"{weeks_between(mp_day, fs_day)} weeks"
        rows[model] = {
            "model": model, "project": p["project"], "inch": p["inch"],
            "bomHQ": week_label(hq), "bomHQDate": fmt_date(hq),
            "bomLocal": week_label(lo), "bomLocalDate": fmt_date(lo),
            "version": s["version"] if s else "",
            **mpf,
            "hqStatus": _status(dh), "localStatus": _status(dl), "hqDelta": dh, "localDelta": dl,
            "firstSop": f"Version {fs}" if fs else DASH_CH,
            "firstAppearMpGap": gap,
        }
    return sorted(rows.values(), key=lambda r: r["model"]), stats


def build_master(dash: Optional[dict], sop: dict, sop_sum: dict[str, dict], bom: list[dict]) -> list[dict]:
    bom_by = {b["model"]: b for b in bom}
    dash_by: "OrderedDict[str, dict]" = OrderedDict()
    for r in dash["rows"]:
        d = dash_by.setdefault(r["model"], {"project": "", "projectName": ""})
        for k in ("project", "projectName"):
            if not d[k] and r[k]:
                d[k] = r[k]
    sop_set = {m for m, s in sop_sum.items() if s["has_qty"]}   # scope: any qty > 0 in any version
    universe = OrderedDict()
    for m in dash_by:
        universe[m] = True
    for m in sorted(sop_set):
        universe[m] = True
    for m in bom_by:
        universe[m] = True

    rows = []
    for m in universe:
        d, b = dash_by.get(m), bom_by.get(m)
        s = sop_sum.get(m)
        in_sop = s is not None and (m in sop_set or b is not None)
        # Descriptive SOP attributes are taken from the SOP sheet whenever the item appears there at all
        # (also when its quantities are all zero and it is therefore outside the SOP scope).
        dims = sop["dims"].get(m)
        if d is not None:
            project = d["project"]
        elif dims:
            project = dims["project"] or dims["any_project"]
        else:
            project = ""
        if b is not None:
            inch = b["inch"]
        elif dims:
            inch = dims["inch"] or dims["any_inch"]
        else:
            inch = ""
        row = {
            "model": m, "project": project, "inch": inch,
            "projectName": d["projectName"] if d else "",
            "version": s["version"] if in_sop else "",
        }
        if b is not None:
            for k in ("mp", "mpDate", "hqTarget", "hqTargetDate", "localTarget", "localTargetDate",
                      "bomHQ", "bomHQDate", "bomLocal", "bomLocalDate", "hqStatus", "localStatus",
                      "hqDelta", "localDelta"):
                row[k] = b[k]
        else:
            mp_day = week_code_to_monday(s["mp_code"]) if in_sop and s["mp_code"] else None
            mpf, _, _ = _mp_fields(mp_day)
            row.update(mpf)
            row.update({"bomHQ": DASH_CH, "bomHQDate": DASH_CH, "bomLocal": DASH_CH, "bomLocalDate": DASH_CH,
                        "hqStatus": "N/A", "localStatus": "N/A", "hqDelta": None, "localDelta": None})
        parts = [n for n, ok in (("DASH", d is not None), ("SOP", in_sop), ("BOM", b is not None)) if ok]
        row["coverage"] = " + ".join(parts)
        row["firstSop"] = b["firstSop"] if b else DASH_CH
        row["firstAppearMpGap"] = b["firstAppearMpGap"] if b else DASH_CH
        rows.append(row)
    rows.sort(key=lambda r: r["model"])   # the original dashboard data is ordered by model code
    return rows
