"""Browser stress/edge tests for the dashboard page (developer tool: needs Playwright, not part of the runtime).

    python tests\\test_browser.py [http://127.0.0.1:8765/]     # optional: a running service with big data loaded
"""
from __future__ import annotations

import json
import os
import sys
import time

from playwright.sync_api import sync_playwright

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PAGE = "file:///" + os.path.join(ROOT, "BOM_Confirmation_Plan_SYSTEM_STATUS_FONT_MATCH.html").replace("\\", "/")
RESULTS = []


def check(name, ok, extra=""):
    RESULTS.append(bool(ok))
    print(("PASS " if ok else "FAIL ") + name + (" | " + str(extra) if extra else ""), flush=True)


GEN = """
(n)=>{
  const st=['MATCH','LATER','EARLIER'], bom=[], master=[];
  const pad=(i,w)=>String(i).padStart(w,'0');
  for(let i=0;i<n;i++){
    const wk=1+(i%52), lab='W'+pad(wk,2)+' / 2026';
    bom.push({model:'M'+pad(i,7),project:'P'+(i%300),inch:String(32+(i%8)*11),bomHQ:lab,bomHQDate:'01 Jan 2026',bomLocal:lab,bomLocalDate:'08 Jan 2026',
      version:202638,mp:lab,mpDate:'05 Oct 2026',hqTarget:lab,hqTargetDate:'',localTarget:lab,localTargetDate:'',hqStatus:st[i%3],localStatus:st[(i+1)%3],
      hqDelta:i%3,localDelta:(i+1)%3,firstSop:'Version 202601',firstAppearMpGap:(i%30)+' weeks'});
  }
  for(let i=0;i<n*2;i++){
    const b=bom[i%n];
    master.push(Object.assign({},b,{model:'X'+pad(i,7),projectName:'Project name '+i,coverage:i%2?'DASH + SOP':'DASH + SOP + BOM'}));
  }
  return {bom,master};
}
"""


def main():
    served = sys.argv[1] if len(sys.argv) > 1 else None
    with sync_playwright() as p:
        b = p.chromium.launch(executable_path=os.environ.get("PW_CHROMIUM") or None)   # optional: a pre-installed Chromium
        page = b.new_context(viewport={"width": 1500, "height": 950}, accept_downloads=True).new_page()
        errs = []
        page.on("pageerror", lambda e: errs.append(str(e)))
        page.on("dialog", lambda d: d.accept())
        page.goto(PAGE)
        page.wait_for_timeout(500)

        # ---- 1. large data: 100,000 BOM models + 200,000 master models -------------------------------------
        # the data is generated and kept INSIDE the page: shipping 300k rows over the automation protocol would be the bottleneck
        page.evaluate("(n)=>{window.__d=(" + GEN.strip() + ")(n);return 0}", 100000)
        t = page.evaluate("""()=>{const d=window.__d;const t0=performance.now();BOM_DATA=d.bom;
            MASTER_DATA.splice(0,MASTER_DATA.length);for(const r of d.master)MASTER_DATA.push(r);
            state.issueFilter='';state.bomPage=1;state.masterPage=1;renderAll();return performance.now()-t0}""")
        check("1a. render 100k BOM + 200k master models", t < 3000, f"{t:.0f} ms")
        kp = page.evaluate("[document.getElementById('stOnTotal').textContent, document.getElementById('bomInfo').textContent]")
        check("1b. KPIs/table reflect the big data", kp[0] == "100000" and "of 100000" in kp[1], kp)
        lat = []
        for q in ("M00", "M000012", "P17", "zzz-no-match", ""):
            t = page.evaluate("""(q)=>{const e=document.getElementById('globalSearch');const t0=performance.now();e.value=q;e.dispatchEvent(new Event('input'));return performance.now()-t0}""", q)
            lat.append(round(t))
        check("1c. search filter latency on 300k rows", max(lat) < 2500, f"{lat} ms")
        t = page.evaluate("""()=>{const t0=performance.now();for(let i=0;i<20;i++)document.getElementById('bomNext').click();return performance.now()-t0}""")
        check("1d. 20 page clicks", t < 3000, f"{t:.0f} ms")
        t = page.evaluate("""()=>{const t0=performance.now();document.getElementById('viewHqBtn').click();return performance.now()-t0}""")
        check("1e. 'View HQ delays' drill-down", t < 3000, f"{t:.0f} ms")
        page.evaluate("document.getElementById('resetBtn').click()")
        with page.expect_download(timeout=60000) as dl:
            page.evaluate("document.getElementById('exportMaster').click()")
        sz = os.path.getsize(dl.value.path())
        check("1f. CSV export of 200k rows", sz > 10_000_000, f"{sz / 1e6:.0f} MB")

        # ---- 2. hostile / odd strings -------------------------------------------------------------------------
        page.reload(); page.wait_for_timeout(400)
        odd = ['<img src=x onerror="window.__pwn=1">', '"><script>window.__pwn=2</script>', "'; alert(1); //", "Zero​Width", "موديل عربي",
               "emoji \U0001F600", "x" * 5000, "=HYPERLINK(\"http://evil\",\"x\")", "+1+1", "@SUM(A1)", "line\nbreak", "tab\tchar", "&amp; &lt;b&gt;"]
        rows = [dict(model=s, project=s, inch=s, bomHQ=s, bomHQDate="", bomLocal=s, bomLocalDate="", version=s, mp=s, mpDate="", hqTarget=s, hqTargetDate="", localTarget=s,
                     localTargetDate="", hqStatus="LATER", localStatus="MATCH", hqDelta=0, localDelta=0, firstSop=s, firstAppearMpGap="5 weeks") for s in odd]
        page.evaluate("""(r)=>{BOM_DATA=r;MASTER_DATA.splice(0,MASTER_DATA.length,...r.map(x=>Object.assign({},x,{projectName:x.model,coverage:x.model})));
            const sel=document.getElementById('globalProject');[...new Set(MASTER_DATA.map(x=>x.project))].forEach(p=>{const o=document.createElement('option');o.value=p;o.textContent=p;sel.appendChild(o)});renderAll()}""", rows)
        page.click("button[data-tab='bom']"); page.wait_for_timeout(300)
        pwn = page.evaluate("window.__pwn")
        injected = page.evaluate("document.querySelectorAll('#bomBody img, #bomBody script, #masterBody img, #masterBody script').length")
        check("2a. markup in data is escaped (no script/img injected, no code ran)", pwn is None and injected == 0, f"pwn={pwn} injected={injected}")
        with page.expect_download() as dl:
            page.evaluate("document.getElementById('exportBom').click()")
        csv = open(dl.value.path(), encoding="utf-8").read()
        cells = [c for line in csv.split("\n")[1:] for c in line.split('","')]
        bad = [c for c in cells if c.lstrip('"').startswith(("=", "+", "@"))]
        check("2b. CSV export neutralises spreadsheet formulas (= + @)", not bad, bad[:2])
        check("2c. no JS errors with hostile strings", not errs, errs[:2])

        # ---- 3. empty / partial data ------------------------------------------------------------------------------
        errs.clear()
        page.click("button[data-tab='dashboard']")
        page.evaluate("BOM_DATA=[];MASTER_DATA.splice(0,MASTER_DATA.length);renderAll()")
        k = page.evaluate("[document.getElementById('stOn').textContent, document.getElementById('issGap').textContent, document.getElementById('v28StatusCenter').textContent, document.getElementById('bomInfo').textContent]")
        check("3a. empty datasets render as zeros without errors", not errs and k[0] == "0" and k[1] == "0" and k[3] == "0 models", f"{k} errs={errs[:1]}")
        page.evaluate("BOM_DATA=[{model:'ONLY-MODEL'},{model:'B',hqStatus:null,localStatus:undefined,firstAppearMpGap:null,project:null}];MASTER_DATA.splice(0,MASTER_DATA.length,{model:'ONLY-MODEL'});renderAll()")
        page.click("button[data-tab='bom']"); page.click("button[data-tab='master']"); page.fill("#globalSearch", "only")
        check("3b. records with missing/null fields render without errors", not errs, errs[:2])

        # ---- 4. thousands of projects in the filter -------------------------------------------------------------
        page.fill("#globalSearch", "")                       # 3b left a search term in the box
        page.evaluate("""()=>{const r=[];for(let i=0;i<5000;i++)r.push({model:'Z'+i,project:'Project '+i,inch:'',hqStatus:'MATCH',localStatus:'MATCH'});
            BOM_DATA=r;MASTER_DATA.splice(0,MASTER_DATA.length,...r);const sel=document.getElementById('globalProject');while(sel.options.length>1)sel.remove(1);
            [...new Set(r.map(x=>x.project))].sort().forEach(p=>{const o=document.createElement('option');o.value=p;o.textContent=p;sel.appendChild(o)});renderAll()}""")
        page.select_option("#globalProject", "Project 4999")
        n = page.evaluate("document.getElementById('masterInfo').textContent")
        check("4. 5,000-project filter works", "of 1 models" in n, n)

        # ---- 4b. usability: sorting, issue-filter chip, empty state, live planning week, Excel-friendly CSV --------
        page.reload(); page.wait_for_timeout(400)
        errs.clear()
        check("4b-1. empty state is shown before any data", page.is_visible("#emptyState"))
        wk = page.evaluate("document.getElementById('planningWeek').textContent")
        exp = time.strftime("W%V / %G")
        check("4b-2. planning week follows today's date", wk == exp, f"{wk} vs {exp}")
        page.evaluate("""()=>{BOM_DATA=[
            {model:'B',project:'P',mp:'W02 / 2027',firstAppearMpGap:'20 weeks',hqStatus:'LATER',localStatus:'MATCH'},
            {model:'A',project:'P',mp:'W50 / 2026',firstAppearMpGap:'3 weeks',hqStatus:'MATCH',localStatus:'MATCH'},
            {model:'C',project:'P',mp:'—',firstAppearMpGap:'—',hqStatus:'LATER',localStatus:'LATER'},
            {model:'D',project:'P',mp:'W01 / 2026',firstAppearMpGap:'100 weeks',hqStatus:'N/A',localStatus:'N/A'}];
            MASTER_DATA.splice(0,MASTER_DATA.length,...BOM_DATA);renderAll()}""")
        check("4b-3. empty state hides once data exists", not page.is_visible("#emptyState"))
        page.click("button[data-tab='bom']")
        col = "() => [...document.querySelectorAll('#bomBody tr')].map(r => r.cells[0].textContent)"
        page.click("#bomBody >> xpath=ancestor::table//th[9]")                 # 1st MP: chronological, blanks last
        check("4b-4. week column sorts chronologically, blanks last", page.evaluate(col) == ["D", "A", "B", "C"], page.evaluate(col))
        page.click("#bomBody >> xpath=ancestor::table//th[9]")
        check("4b-5. second click sorts descending", page.evaluate(col) == ["B", "A", "D", "C"], page.evaluate(col))
        page.click("#bomBody >> xpath=ancestor::table//th[8]")
        check("4b-6. 'N weeks' column sorts numerically", page.evaluate(col) == ["A", "B", "D", "C"], page.evaluate(col))
        page.click("button[data-tab='dashboard']"); page.click("#viewHqBtn")
        chip = page.evaluate("[document.getElementById('bomFilterChip').classList.contains('show'), document.getElementById('bomFilterText').textContent]")
        check("4b-7. drill-down shows which issue filter is active", chip == [True, "Only: HQ delay"] and page.evaluate(col) == ["B", "C"],
              [chip, page.evaluate(col)])
        page.click("#bomFilterClear")
        check("4b-8. clearing the chip shows all BOM models again", len(page.evaluate(col)) == 4 and not page.is_visible("#bomFilterChip"))
        with page.expect_download() as dl:
            page.evaluate("document.getElementById('exportBom').click()")
        raw = open(dl.value.path(), "rb").read()
        check("4b-9. CSV starts with a UTF-8 BOM and follows the visible sort", raw.startswith(b"\xef\xbb\xbf") and
              [l.split(",")[0] for l in raw.decode("utf-8-sig").splitlines()[1:]] == ['"A"', '"B"', '"D"', '"C"'])
        check("4b-10. no JS errors in the usability features", not errs, errs[:2])

        # ---- 5. optional: a live service with a big database ------------------------------------------------
        if served:
            page2 = b.new_context().new_page()
            e2 = []
            page2.on("pageerror", lambda e: e2.append(str(e)))
            t0 = time.time()
            page2.goto(served)
            page2.wait_for_function("document.getElementById('updateStatus').textContent.includes('local database')", timeout=120000)
            dt = time.time() - t0
            n = page2.evaluate("[BOM_DATA.length, MASTER_DATA.length]")
            check("5. served page loads the big database", not e2 and n[0] > 1000, f"{dt:.1f}s for {n[0]:,} BOM / {n[1]:,} master models")
        check("6. no unexpected JS errors overall", not errs, errs[:2])
        b.close()
    print(f"\nRESULT: {sum(RESULTS)}/{len(RESULTS)} passed")
    return 0 if all(RESULTS) else 1


if __name__ == "__main__":
    sys.exit(main())
