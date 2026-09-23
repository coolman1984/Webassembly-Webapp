"""Pipeline orchestrator:  3 Excel files -> validate -> Excel COM extract -> transform -> SQLite.

CLI:
    python -m pipeline.run FILE FILE FILE          # any order; roles are detected from content
    python -m pipeline.run --dir "D:\\some\\folder" # folder that contains exactly the 3 workbooks
"""
from __future__ import annotations

import argparse
import hashlib
import os
import re
import sys
import time
from typing import Callable, Optional

from . import excel_com as X
from . import history as H
from . import sources as S
from . import transform as T

ALLOWED_EXT = (".xlsx", ".xlsm", ".xlsb")
KNOWN_CATES = {"ship", "sop"}
DROP_WARN_RATIO = 0.30          # warn when a refresh has >30% fewer models than the data it replaces


def _disp(path: str) -> str:
    """File name as the user knows it (the service stores uploads as slotN__<name>)."""
    return re.sub(r"^slot\d__", "", os.path.basename(path))


def _sha256(path: str) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def precheck(paths: list[str]) -> list[str]:
    """Cheap validation before Excel is started."""
    if len(paths) != 3:
        raise S.ValidationError(f"Exactly 3 files are required (DASH status, SEEG SOP, New Model BOM); got {len(paths)}.")
    seen = {}
    names = [os.path.basename(p).lower() for p in paths]
    if len(set(names)) != len(names):
        raise S.ValidationError("Two of the files have the same file name - Excel cannot open two workbooks with "
                                "the same name at once. Rename one of them.")
    for p in paths:
        name = _disp(p)
        if name.startswith("~$"):
            raise S.ValidationError(f"'{name}' is an Excel lock/temp file, not a workbook.")
        if not p.lower().endswith(ALLOWED_EXT):
            raise S.ValidationError(f"'{name}': unsupported type (expected {', '.join(ALLOWED_EXT)}).")
        if not os.path.isfile(p):
            raise S.ValidationError(f"'{name}': file not found.")
        if os.path.getsize(p) == 0:
            raise S.ValidationError(f"'{name}': file is empty.")
        digest = _sha256(p)
        if digest in seen:
            raise S.ValidationError(f"'{name}' is identical to '{seen[digest]}' - the same file was supplied twice.")
        seen[digest] = name
    return list(seen.keys())


def extract_all(paths: list[str], log: Callable[[str], None] = print, pidfile: Optional[str] = None,
                timeout_s: int = 1200) -> dict:
    """Validate + extract the three sources. Raises ValidationError with a user-readable message."""
    hashes = precheck(paths)
    t0 = time.time()
    warnings: list[str] = []
    with X.ExcelSession(timeout_s=timeout_s, pidfile=pidfile) as xl:
        if xl.startup_workbooks_closed:
            log(f"Closed Excel startup file(s) unopened/unread: {xl.startup_workbooks_closed}")
            warnings.append(f"Excel auto-loaded a startup workbook ({', '.join(xl.startup_workbooks_closed)}) - "
                            "it was closed immediately without being read or run; no macros from it were executed.")
        opened = []
        for p, digest in zip(paths, hashes):
            name = _disp(p)
            log(f"Opening {name} ...")
            wb = xl.open(p)                       # each workbook is opened exactly once
            det, notes, diag = S.detect_workbook(wb)
            if det is None:
                raise S.ValidationError(S.explain_unrecognised(name, diag))
            warnings += [f"{name}: {n}" for n in notes]
            log(f"  recognised as {det.role} (sheet '{det.sheet}')")
            opened.append((p, digest, wb, det))

        roles = [d.role for _, _, _, d in opened]
        for r in S.ROLES:
            if roles.count(r) == 0:
                raise S.ValidationError(f"Missing source: no {r} file among the 3 uploads (detected: {roles}).")
            if roles.count(r) > 1:
                dup = [_disp(p) for p, _, _, d in opened if d.role == r]
                raise S.ValidationError(f"Two files look like the {r} source: {dup}. Supply one of each.")

        out: dict = {"files": [], "warnings": warnings}
        for p, digest, wb, det in opened:
            ws = wb.Worksheets(det.sheet)
            log(f"Extracting {det.role} ({_disp(p)}) ...")
            t1 = time.time()
            fn = {"DASH": S.extract_dash, "NEWMODEL": S.extract_newmodel, "SEEG": S.extract_sop}[det.role]
            data = fn(ws, det, X.is_1904(wb))
            key = {"DASH": "dash", "NEWMODEL": "nm", "SEEG": "sop"}[det.role]
            out[key] = data
            out["files"].append({"role": det.role, "name": _disp(p), "sha256": digest,
                                 "size": os.path.getsize(p), "sheet": det.sheet})
            log(f"  done in {time.time() - t1:.1f}s")
        del opened, ws, wb                        # release COM references before Excel is closed
    if xl.forced_kill:
        log("Excel did not close by itself and was stopped.")
        warnings.append("Excel did not close normally after reading the files and was stopped by the program.")
    out["seconds"] = round(time.time() - t0, 1)
    _sanity(out)
    return out


def _sanity(ext: dict) -> None:
    """Post-extraction content validation: hard errors for unusable data, warnings for suspicious data."""
    w = ext["warnings"]
    if not ext["dash"]["rows"]:
        raise S.ValidationError("DASH file has a valid header but no usable MODEL_CODE rows.")
    if not ext["nm"]["plan"]:
        raise S.ValidationError("New Model file has no 'Plan' rows with a Model code.")
    dm = [r["model"] for r in ext["dash"]["rows"]]
    if len(dm) != len(set(dm)):
        w.append(f"DASH: {len(dm) - len(set(dm))} duplicate model rows (merged, first non-empty value wins).")
    if ext["dash"]["skipped_no_model"]:
        w.append(f"DASH: {ext['dash']['skipped_no_model']} rows without a model code were skipped.")
    blank_proj = sum(1 for r in ext["dash"]["rows"] if not r["project"])
    if blank_proj > 0.5 * len(dm):
        w.append(f"DASH: {blank_proj} of {len(dm)} models have no MKT_PROJECT - check the file.")
    nm = [r["model"] for r in ext["nm"]["plan"]]
    if len(nm) != len(set(nm)):
        w.append(f"New Model: {len(nm) - len(set(nm))} duplicate Plan rows (first one kept).")
    if ext["nm"]["actual_rows"] != len(ext["nm"]["plan"]):
        w.append(f"New Model: {len(ext['nm']['plan'])} Plan rows vs {ext['nm']['actual_rows']} Actual rows (expected pairs).")
    if ext["nm"]["other_types"]:
        w.append(f"New Model: rows with unexpected Type values were ignored: {ext['nm']['other_types']}.")
    if ext["nm"].get("date1904"):
        w.append("New Model: workbook uses the 1904 date system; dates were converted.")
    sop = ext["sop"]
    if sop["bad_version"]:
        w.append(f"SOP: {sop['bad_version']} rows skipped because Version is not numeric.")
    if sop["non_numeric"]:
        w.append(f"SOP: {sop['non_numeric']} non-numeric quantity cells treated as 0.")
    if sop["rows_without_item"]:
        w.append(f"SOP: {sop['rows_without_item']} rows without an Item were skipped.")
    unknown = {c: n for c, n in sop["cates"].items() if c not in KNOWN_CATES}
    if unknown:
        w.append(f"SOP: unknown Cate values ignored by the calculations (only Ship/SOP are used): {unknown}.")
    if sop["dup_weeks"]:
        w.append(f"SOP: the same week appears in more than one column: {sop['dup_weeks']} - check the header row.")
    bad_weeks = T.invalid_week_codes(sop)
    if bad_weeks:
        w.append(f"SOP: week columns that are not real ISO weeks: {bad_weeks} (their dates cannot be computed).")
    sop_items = {k[0] for k in sop["agg"]}
    missing = [m for m in nm if m not in sop_items]
    if missing:
        w.append(f"New Model: {len(missing)} BOM models have no rows in the SOP sheet (MP falls back to the New Model MP date).")
    bad = sum(1 for r in ext["nm"]["plan"] if (not S.is_blank(r["bom_hq"]) and not T.parse_date(r["bom_hq"]))
              or (not S.is_blank(r["bom_local"]) and not T.parse_date(r["bom_local"])))
    if bad:
        w.append(f"New Model: {bad} BOM rows have unreadable HQ/LOCAL dates (shown as unknown).")


def build_datasets(ext: dict, max_version: Optional[int] = None) -> tuple[list[dict], list[dict], dict]:
    sop_sum = T.summarise_sop(ext["sop"], max_version)
    bom, bstats = T.build_bom(ext["nm"], sop_sum)
    master = T.build_master(ext["dash"], ext["sop"], sop_sum, bom)
    versions = sorted({k[2] for k in ext["sop"]["agg"]})
    report = {
        "bom_models": len(bom), "master_models": len(master), "bom_stats": bstats,
        "sop_versions": [versions[0], versions[-1]], "sop_items": len(sop_sum),
        "sop_rows": ext["sop"]["rows_read"], "dash_rows": ext["dash"]["rows_read"],
    }
    if bstats["mp_bad_week"]:
        ext["warnings"].append(f"{bstats['mp_bad_week']} BOM models have an MP week that is not a real ISO week; "
                               "their MP was taken from the New Model file instead.")
    return bom, master, report


def _previous_counts(db_path: str) -> Optional[tuple[int, int]]:
    from . import store
    try:
        d = store.read_datasets(db_path)
        return len(d["bom"]), len(d["master"])
    except Exception:
        return None


def history_path(db_path: str) -> str:
    return os.path.join(os.path.dirname(os.path.abspath(db_path)), "history.db")


def process(paths: list[str], db_path: str, log: Callable[[str], None] = print, pidfile: Optional[str] = None,
            timeout_s: int = 1200, hist_path: Optional[str] = None) -> dict:
    from . import store
    hist_path = hist_path or history_path(db_path)
    ext = extract_all(paths, log, pidfile, timeout_s)
    bom, master, report = build_datasets(ext)
    sop_max = report["sop_versions"][1]
    try:
        ext["warnings"] += H.compare_with_latest(hist_path, ext["files"], sop_max)
    except Exception as e:                    # history is an extra: it must never block a refresh
        log(f"History check skipped: {e}")
    prev = _previous_counts(db_path)
    if prev:
        for label, old, new in (("BOM", prev[0], len(bom)), ("All Models", prev[1], len(master))):
            if old and new < old * (1 - DROP_WARN_RATIO):
                ext["warnings"].append(f"{label} dropped from {old} to {new} models - please check the files are the right ones.")
    log("Writing local database ...")
    meta = store.write_db(db_path, ext, bom, master, report)
    try:
        sid = H.record(hist_path, meta, bom, master, ext["files"], sop_max)
        meta["snapshot_id"] = sid
        log("Saved to history" if sid else "Same files as the last refresh - history unchanged")
    except Exception as e:
        log(f"WARNING: the refresh worked but could not be saved to the history ({e})")
    log(f"Done: {len(bom)} BOM models, {len(master)} total models")
    return meta


def _pick_from_dir(d: str) -> list[str]:
    files = [os.path.join(d, f) for f in os.listdir(d) if f.lower().endswith(ALLOWED_EXT) and not f.startswith("~$")]
    files.sort(key=os.path.getmtime, reverse=True)
    return files


def main(argv=None) -> int:
    try:
        sys.stdout.reconfigure(errors="replace")
        sys.stderr.reconfigure(errors="replace")
    except Exception:
        pass
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("files", nargs="*")
    ap.add_argument("--dir", help="folder that contains exactly the 3 source workbooks")
    ap.add_argument("--db", default=os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data", "dashboard.db"))
    a = ap.parse_args(argv)
    files = a.files or (_pick_from_dir(a.dir) if a.dir else [])
    try:
        process(files, a.db)
    except S.ValidationError as e:
        print(f"VALIDATION FAILED: {e}", file=sys.stderr)
        return 2
    except X.ExcelError as e:
        print(f"EXCEL ERROR: {e}", file=sys.stderr)
        return 3
    return 0


if __name__ == "__main__":
    sys.exit(main())
