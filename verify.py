"""Regression check: rebuild the dashboard data from the Excel files and compare with the data that was
embedded in the ORIGINAL dashboard (backup/*.original.html).

The original snapshot was built from an older SOP export (latest version 202636), so the SOP versions are
capped at 202636 for the comparison. Every remaining difference must be explainable; anything else fails.

    python verify.py FILE FILE FILE [--cap 202636]
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)
from pipeline import run  # noqa: E402

ORIGINAL = os.path.join(ROOT, "backup", "BOM_Confirmation_Plan_SYSTEM_STATUS_FONT_MATCH.original.html")


def load_embedded() -> tuple[list[dict], list[dict]]:
    lines = open(ORIGINAL, encoding="utf-8").read().split("\n")
    pick = lambda name: json.loads(re.match(rf"let {name}=(.*);\s*$", next(l for l in lines if l.startswith(f"let {name}="))).group(1))
    return pick("BOM_DATA"), pick("MASTER_DATA")


def gap(s):
    m = re.search(r"(-?\d+(?:\.\d+)?)\s*weeks?", str(s or ""), re.I)
    return float(m.group(1)) if m else None


def kpis(bom):
    return {"models": len(bom), "on_time": sum(x["localStatus"] == "MATCH" for x in bom),
            "sop_issue": sum(1 for x in bom if gap(x["firstAppearMpGap"]) is not None and gap(x["firstAppearMpGap"]) < 12),
            "hq_delay": sum(x["hqStatus"] == "LATER" for x in bom), "seeg_delay": sum(x["localStatus"] == "LATER" for x in bom)}


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("files", nargs=3)
    ap.add_argument("--cap", type=int, default=202636)
    a = ap.parse_args()
    emb_bom, emb_master = load_embedded()
    ext = run.extract_all(a.files, print)
    bom, master, _ = run.build_datasets(ext, max_version=a.cap)
    sop_min = {}
    for (item, cate, ver) in ext["sop"]["agg"]:
        sop_min[item] = min(ver, sop_min.get(item, 10**9))

    bad = 0
    eb = {r["model"]: r for r in emb_bom}
    mb = {r["model"]: r for r in bom}
    d = [m for m in eb if m not in mb or eb[m] != mb[m]]
    print(f"\nBOM_DATA     : {len(mb)} rebuilt vs {len(eb)} original -> {len(d)} differing rows, {len(set(mb) ^ set(eb))} membership differences")
    bad += len(d) + len(set(mb) ^ set(eb))

    em = {r["model"]: r for r in emb_master}
    mm = {r["model"]: r for r in master}
    lost = set(em) - set(mm)
    new_scope = [m for m in set(mm) - set(em)]
    explained = {"new SOP-scope rows": 0, "DASH row now in SOP scope": 0, "SOP rows only after the snapshot": 0}
    unexplained = []
    for m, e in em.items():
        r = mm.get(m)
        if r is None or r == e:
            continue
        keys = [k for k in e if e[k] != r.get(k)]
        if "SOP" not in e["coverage"] and "SOP" in r["coverage"]:
            explained["DASH row now in SOP scope"] += 1            # scope decision: any qty > 0 in any version
        elif keys == ["inch"] and sop_min.get(m, 0) > a.cap:
            explained["SOP rows only after the snapshot"] += 1     # item first appears in a newer SOP version
        else:
            unexplained.append((m, keys))
    explained["new SOP-scope rows"] = len(new_scope)
    print(f"MASTER_DATA  : {len(mm)} rebuilt vs {len(em)} original | missing from rebuild: {len(lost)}")
    for k, v in explained.items():
        print(f"   explained  {v:4d}  {k}")
    print(f"   UNEXPLAINED {len(unexplained)} {unexplained[:5]}")
    bad += len(lost) + len(unexplained)

    print("\nKPIs (BOM scope)   original:", kpis(emb_bom))
    print("                   rebuilt :", kpis(bom), "(capped at version %d)" % a.cap)
    bad += kpis(emb_bom) != kpis(bom)
    full_bom, _, rep = run.build_datasets(ext)
    print("                   CURRENT DATA (no cap):", kpis(full_bom), "| SOP versions", rep["sop_versions"])
    print("\nRESULT:", "PASS - rebuilt data reproduces the original dashboard data" if not bad else f"FAIL ({bad} problems)")
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
