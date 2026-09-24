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

Additional fields (machine-readable, used for sorting, current-week logic and data-quality checks):

  *Iso            ISO date (YYYY-MM-DD, Monday of the week) of every week field, None when unknown
  gapWeeks        firstAppearMpGap as an int;  sopIssue = gapWeeks < SOP_LEAD_WEEKS
  mpFrom          where the MP week came from: 'ship' / 'sop' (SOP sheet) or 'newmodel' (New Model file)
  nmMpIso         the New Model file's own SET PLANT MP date;  mpMismatch = weeks(SOP MP - New Model MP)
  hq/localActual  the New Model 'Actual' row: a date (confirmed on that week), a mark such as 'OK'
                  (confirmed, date unknown) or '—' (not confirmed yet);  hq/localConfirmed booleans
  hqAfterLocal    the HQ BOM is planned later than the LOCAL BOM (HQ is expected first: MP-13W vs MP-12W)
"""
from __future__ import annotations

import datetime as dt
import re
from collections import OrderedDict, defaultdict
from typing import Optional

from .sources import as_text, is_blank

HQ_WEEKS = 13
LOCAL_WEEKS = 12
SOP_LEAD_WEEKS = 12      # a model must appear in the SOP at least this many weeks before its first MP
RULES = {"hqWeeks": HQ_WEEKS, "localWeeks": LOCAL_WEEKS, "sopLeadWeeks": SOP_LEAD_WEEKS}
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


def iso(d: Optional[dt.date]) -> Optional[str]:
    """Monday of the date's week as YYYY-MM-DD (sortable, and what the page uses for current-week logic)."""
    return monday(d).isoformat() if d else None


def norm_inch(v) -> str:
    """'43', 43.0 and '43.0' are the same screen size; placeholders such as '-' mean 'unknown'."""
    if is_blank(v):
        return ""
    s = as_text(v)
    try:
        f = float(s)
        return str(int(f)) if f.is_integer() else str(f)
    except ValueError:
        return s


def norm_project(v) -> str:
    return "" if is_blank(v) else as_text(v)


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
_CONFIRMED_MARKS = {"ok", "o", "y", "yes", "done", "confirmed", "complete", "completed", "v", "✓", "✔", "√", "●"}


def _status(delta: Optional[int]) -> str:
    if delta is None:
        return "N/A"
    return "MATCH" if delta == 0 else ("LATER" if delta > 0 else "EARLIER")


def _mp_fields(mp_day: Optional[dt.date]) -> tuple[dict, Optional[dt.date], Optional[dt.date]]:
    hq_t = mp_day - dt.timedelta(weeks=HQ_WEEKS) if mp_day else None
    lo_t = mp_day - dt.timedelta(weeks=LOCAL_WEEKS) if mp_day else None
    return {
        "mp": week_label(mp_day), "mpDate": fmt_date(mp_day), "mpIso": iso(mp_day),
        "hqTarget": week_label(hq_t), "hqTargetDate": fmt_date(hq_t), "hqTargetIso": iso(hq_t),
        "localTarget": week_label(lo_t), "localTargetDate": fmt_date(lo_t), "localTargetIso": iso(lo_t),
    }, hq_t, lo_t


def actual_state(v) -> tuple[str, Optional[str], bool]:
    """New Model 'Actual' cell -> (display, iso date, confirmed).
    A date means confirmed in that week; a positive mark (OK/Done/Y/...) means confirmed without a date;
    blank/placeholder means not confirmed; any other text is shown as-is but not counted as confirmed."""
    if is_blank(v):
        return DASH_CH, None, False
    d = parse_date(v)
    if d:
        return week_label(d), iso(d), True
    s = as_text(v)
    return s, None, s.lower() in _CONFIRMED_MARKS


def build_bom(nm: dict, sop_sum: dict[str, dict]) -> tuple[list[dict], dict]:
    rows: "OrderedDict[str, dict]" = OrderedDict()
    stats = {"mp_from_sop": 0, "mp_from_newmodel": 0, "mp_missing": 0, "bad_dates": 0, "duplicates": 0,
             "mp_bad_week": 0, "mp_mismatch": 0, "hq_after_local": 0, "hq_confirmed": 0, "local_confirmed": 0,
             "actual_other_text": 0}
    actual = nm.get("actual", {})
    for p in nm["plan"]:
        model = p["model"]
        if model in rows:
            stats["duplicates"] += 1
            continue
        s = sop_sum.get(model)
        nm_mp = parse_date(p["mp"])
        mp_day = week_code_to_monday(s["mp_code"]) if s and s["mp_code"] else None
        mp_from = s["mp_from"] if mp_day else None
        if s and s["mp_code"] and not mp_day:
            stats["mp_bad_week"] += 1
        if mp_day:
            stats["mp_from_sop"] += 1
        else:
            mp_day = nm_mp
            mp_from = "newmodel" if mp_day else None
            stats["mp_from_newmodel" if mp_day else "mp_missing"] += 1
        mismatch = weeks_between(mp_day, nm_mp) if mp_day and nm_mp and mp_from != "newmodel" else None
        if mismatch:
            stats["mp_mismatch"] += 1
        hq, lo = parse_date(p["bom_hq"]), parse_date(p["bom_local"])
        if (not is_blank(p["bom_hq"]) and not hq) or (not is_blank(p["bom_local"]) and not lo):
            stats["bad_dates"] += 1
        hq_after_local = bool(hq and lo and monday(hq) > monday(lo))
        stats["hq_after_local"] += hq_after_local
        mpf, hq_t, lo_t = _mp_fields(mp_day)
        dh = weeks_between(hq, hq_t) if hq and hq_t else None
        dl = weeks_between(lo, lo_t) if lo and lo_t else None
        fs = s["first_sop_version"] if s else None
        gap, gap_n = DASH_CH, None
        fs_day = version_to_monday(fs) if fs else None
        if fs_day and mp_day:
            gap_n = weeks_between(mp_day, fs_day)
            gap = f"{gap_n} weeks"
        a = actual.get(model, {})
        ha, ha_iso, ha_ok = actual_state(a.get("bom_hq"))
        la, la_iso, la_ok = actual_state(a.get("bom_local"))
        stats["hq_confirmed"] += ha_ok
        stats["local_confirmed"] += la_ok
        stats["actual_other_text"] += (ha != DASH_CH and not ha_ok) + (la != DASH_CH and not la_ok)
        rows[model] = {
            "model": model, "project": p["project"], "inch": norm_inch(p["inch"]),
            "bomHQ": week_label(hq), "bomHQDate": fmt_date(hq), "bomHQIso": iso(hq),
            "bomLocal": week_label(lo), "bomLocalDate": fmt_date(lo), "bomLocalIso": iso(lo),
            "version": s["version"] if s else "",
            **mpf,
            "hqStatus": _status(dh), "localStatus": _status(dl), "hqDelta": dh, "localDelta": dl,
            "firstSop": f"Version {fs}" if fs else DASH_CH,
            "firstAppearMpGap": gap, "gapWeeks": gap_n,
            "sopIssue": gap_n is not None and gap_n < SOP_LEAD_WEEKS,
            "mpFrom": mp_from, "nmMpIso": iso(nm_mp), "mpMismatch": mismatch,
            "hqActual": ha, "hqActualIso": ha_iso, "hqConfirmed": ha_ok,
            "localActual": la, "localActualIso": la_iso, "localConfirmed": la_ok,
            "hqAfterLocal": hq_after_local,
        }
    return sorted(rows.values(), key=lambda r: r["model"]), stats


# fields a master row copies from its BOM row (a non-BOM row gets the "no BOM" defaults below)
BOM_ONLY = {"bomHQ": DASH_CH, "bomHQDate": DASH_CH, "bomHQIso": None, "bomLocal": DASH_CH, "bomLocalDate": DASH_CH,
            "bomLocalIso": None, "hqStatus": "N/A", "localStatus": "N/A", "hqDelta": None, "localDelta": None,
            "firstSop": DASH_CH, "firstAppearMpGap": DASH_CH, "gapWeeks": None, "sopIssue": False,
            "nmMpIso": None, "mpMismatch": None, "hqActual": DASH_CH, "hqActualIso": None, "hqConfirmed": False,
            "localActual": DASH_CH, "localActualIso": None, "localConfirmed": False, "hqAfterLocal": False}
MP_KEYS = ("mp", "mpDate", "mpIso", "hqTarget", "hqTargetDate", "hqTargetIso", "localTarget", "localTargetDate",
           "localTargetIso", "mpFrom")


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
        if d is not None and d["project"]:
            project = d["project"]
        elif b is not None and b["project"]:
            project = b["project"]
        elif dims:
            project = norm_project(dims["project"]) or norm_project(dims["any_project"])
        else:
            project = ""
        if b is not None and b["inch"]:
            inch = b["inch"]
        elif dims:
            inch = norm_inch(dims["inch"]) or norm_inch(dims["any_inch"])
        else:
            inch = ""
        row = {
            "model": m, "project": project, "inch": inch,
            "projectName": d["projectName"] if d else "",
            "version": s["version"] if in_sop else "",
        }
        if b is not None:
            for k in MP_KEYS + tuple(BOM_ONLY):
                row[k] = b[k]
        else:
            mp_day = week_code_to_monday(s["mp_code"]) if in_sop and s["mp_code"] else None
            mpf, _, _ = _mp_fields(mp_day)
            row.update(mpf)
            row["mpFrom"] = s["mp_from"] if mp_day else None
            row.update(BOM_ONLY)
        parts = [n for n, ok in (("DASH", d is not None), ("SOP", in_sop), ("BOM", b is not None)) if ok]
        row["coverage"] = " + ".join(parts)
        rows.append(row)
    rows.sort(key=lambda r: r["model"])   # the original dashboard data is ordered by model code
    return rows


# ----------------------------------------------------------------------------- KPIs / refresh comparison
def kpis(bom: list[dict]) -> dict:
    """Headline numbers of the BOM scope. 'On time' = LOCAL BOM planned no later than MP-12W (MATCH or EARLIER)."""
    n = len(bom)
    on = sum(r["localStatus"] in ("MATCH", "EARLIER") for r in bom)
    late = sum(r["localStatus"] == "LATER" for r in bom)
    return {"models": n, "on_time": on, "late": late, "no_plan": n - on - late,
            "sop_issue": sum(bool(r.get("sopIssue")) for r in bom),
            "hq_delay": sum(r["hqStatus"] == "LATER" for r in bom),
            "local_delay": late,
            "any_issue": sum(bool(r.get("sopIssue")) or r["hqStatus"] == "LATER" or r["localStatus"] == "LATER" for r in bom),
            "local_confirmed": sum(bool(r.get("localConfirmed")) for r in bom)}


# fields whose change between two refreshes is worth reporting, with a readable label
TRACKED = {"mp": "1st MP", "bomHQ": "BOM HQ plan", "bomLocal": "BOM LOCAL plan", "hqStatus": "HQ status",
           "localStatus": "LOCAL status", "version": "Latest SOP", "localActual": "LOCAL actual", "hqActual": "HQ actual"}


def diff_bom(old: Optional[list[dict]], new: list[dict]) -> list[dict]:
    """Model-level changes between the previous refresh and this one (added / removed / field changed)."""
    if old is None:
        return []
    ob = {r["model"]: r for r in old}
    nb = {r["model"]: r for r in new}
    out = []
    for m in sorted(nb.keys() - ob.keys()):
        out.append({"model": m, "project": nb[m]["project"], "kind": "added", "field": "", "old": "", "new": ""})
    for m in sorted(ob.keys() - nb.keys()):
        out.append({"model": m, "project": ob[m].get("project", ""), "kind": "removed", "field": "", "old": "", "new": ""})
    for m in sorted(nb.keys() & ob.keys()):
        for k, label in TRACKED.items():
            if k not in ob[m]:                               # previous database did not have this field yet
                continue
            a, b = ob[m][k], nb[m][k]
            if str(a if a is not None else "") != str(b if b is not None else ""):
                out.append({"model": m, "project": nb[m]["project"], "kind": "changed", "field": label,
                            "old": "" if a is None else str(a), "new": "" if b is None else str(b)})
    return out
