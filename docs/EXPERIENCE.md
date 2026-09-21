# Project experience log: BOM Confirmation Plan dashboard data pipeline

This is a narrative record of how this pipeline was built and hardened, why the harder decisions were made
the way they were, and what was actually found and fixed during testing - written for a future maintainer (human
or AI) picking this project back up. For reusable, project-independent Excel-COM/packaging/security patterns
extracted from this work, see `.claude/skills/excel-com-automation/` - that skill is the generalised version of
everything below.

## 1. Starting point

The dashboard (`BOM_Confirmation_Plan_SYSTEM_STATUS_FONT_MATCH.html`) already existed as a self-contained,
single-file HTML page with a working "Update Data" panel that let a user attach 3 Excel-like files and have the
page re-derive its data client-side in JavaScript. The task was to replace that JS-only ingestion with a real
Excel-COM-based pipeline (per an explicit architecture requirement: Excel COM extraction → clean/normalise →
local data layer → browser), without changing the dashboard's design, and to reverse-engineer the exact
calculation rules from the data already embedded in the page, since no separate spec existed.

**The untouched original was preserved first** (`backup/BOM_Confirmation_Plan_SYSTEM_STATUS_FONT_MATCH.original.html`)
before any edit, and `verify.py` was written specifically to diff a freshly rebuilt dataset against that
original's embedded `BOM_DATA`/`MASTER_DATA` arrays - this became the project's ground-truth regression test and
stayed useful through every later refactor.

## 2. Reverse-engineering the calculation rules

The three source files (a flat DASH status export, a ~246,000-row SEEG SOP workbook with 36 weekly-quantity
snapshots, and a two-header-row New Model BOM plan) had to be inspected, then every field in the embedded
dashboard data (`mp`, `firstSop`, `hqTarget`/`localTarget`, `hqStatus`/`localStatus`, `firstAppearMpGap`,
`coverage`) had to be matched to a specific rule by hypothesis-and-verification against the real data, not
assumed. The one rule that could *not* be recovered exactly from the files - which SOP-only models belong in the
"All Models" list - was resolved by asking the user directly rather than guessing (recorded as "any item with
qty>0 in any SOP version, plus all BOM models" - documented in the README as a deliberate reconstruction, not a
rediscovered original rule).

Every calculation rule that *was* recoverable was verified to reproduce the original dashboard's numbers exactly
(259 on-time / 65 SOP issues / 111 HQ delays / 111 SEEG delays across 371 BOM models) once the SOP version set
was capped to match the snapshot the original data was built from; with the current (newer) source data the
numbers legitimately differ (257/65/112/112) because the SOP file itself has moved on - `verify.py` reports both
so this is never mistaken for a bug.

## 3. Building the pipeline

`pipeline/` was built as: `excel_com.py` (the Excel automation layer), `sources.py` (content-based file
recognition + per-source extraction), `transform.py` (the calculation rules), `store.py` (SQLite with atomic
writes and a schema version), `run.py` (orchestration, validation, warnings). `server.py` wraps it as a local
HTTP service so the existing dashboard's upload UI could drive it with minimal changes (a new "engine v2" script
replacing the old client-side-only one, keeping every existing DOM id/class/handler intact so the visual design
and all other JS - filters, pagination, charts - were untouched).

## 4. The stress-testing pass: what was actually found

The user later asked for exhaustive stress-testing, self-containment (no install required), and specific
hardening against VBA/`PERSONAL.XLSB` ever running. This phase is where most of the real bugs surfaced - all of
them are written up in detail, generalised, in `.claude/skills/excel-com-automation/references/gotchas.md`.
Summary, in the order found:

1. **`Find()` with keyword arguments silently returned the wrong cell** - the very first regression test written
   against the "real last cell, not `UsedRange`" helper caught this immediately, because the fixture deliberately
   had a different first-match and last-match. Fixed by switching to positional arguments.
2. **A shared COM apartment bug broke a *second*, unrelated Excel instance** - a test fixture's own Excel
   instance started failing with "Object is not connected to server" right after the code-under-test's session
   exited cleanly. Root cause: `CoUninitialize()` was tearing down the whole thread's COM apartment, not just
   releasing the one object it was meant to. Fixed by only uninitialising apartments the session's own code
   created for itself.
3. **A password-protected test file made the whole suite hang** - twice, at the exact same point, in two
   independent runs, which was the signal this wasn't a coincidental slow machine. Diagnosed by checking CPU time
   vs. wall-clock time on the stuck process (near-zero CPU = blocked, not computing) and finding the stuck Excel
   process had a real (if invisible) window handle and reported itself "Responding" - the signature of an
   unanswerable modal dialog, not a crash. Fixed with a dedicated ~30-second watchdog around just the
   `Workbooks.Open()` call, separate from (and much shorter than) the overall session timeout, verified safe
   against the real 38MB source file (which opens in ~7 seconds).
4. **PERSONAL.XLSB/VBA hardening**, added defensively per explicit user instruction, without ever creating or
   touching a real macro-enabled startup file on the build machine (that would have been exactly the kind of
   irreversible environment change being guarded against) - verified instead by inspecting the session's own
   `AutomationSecurity`/`EnableEvents` state and confirming no workbook is ever left open at session start on a
   clean machine.
5. A genuine self-contained Windows Python runtime was built and verified to work with the machine's own Python
   and `PATH` fully scrubbed away - confirming the tool truly needs nothing beyond Windows + Excel.
6. Local-service security (`Host` header validation against DNS rebinding, a custom-header + `Origin` check on
   all writes, sanitised uploads, CSV formula-injection neutralisation) was added and tested with an actual
   hostile test suite (path traversal attempts, oversized/negative `Content-Length`, cross-site-style requests
   without the required header, a dropped mid-upload connection), not just happy-path checks.
7. An "inbox" folder fallback was added after discovering, empirically on the actual development machine, that a
   corporate security/DRM agent intercepted the browser's file-picker reads on the real source files (rewriting
   them to opaque temp names) while leaving direct filesystem access from the local service completely
   unaffected - a real-world failure mode that would otherwise have been very hard to diagnose remotely.

## 5. Final verified state

- `verify.py` against the real files: exact match with the original dashboard's data (all differences fully
  explained: newer SOP data, and the user-confirmed SOP-scope reconstruction).
- `tests/test_edge_cases.py`: 26/26 passing, run through the bundled runtime, covering storage-fuzzed variants,
  randomised combinations, hostile/edge-case files, and Excel-layer behaviour (last-cell correctness, hidden/
  filtered rows, error-cell handling, zero-leak-on-exception, the open-watchdog, and the PERSONAL.XLSB defence).
- `tests/test_service.py`: 19/19 passing - security (DNS rebinding, CSRF), hostile uploads, concurrency, crash
  recovery, port fallback, the inbox flow.
- `tests/test_stress.py`: passing at 500,000 SOP rows / 60,000 models, and at Excel's literal hard limit of
  1,048,575 data rows in one sheet.
- `tests/test_browser.py`: passing at 100,000 BOM / 200,000 master rows rendered in the actual dashboard page,
  plus hostile-string/injection and CSV-export-injection checks.
- No leftover `EXCEL.EXE` or Python processes after any of the above, verified explicitly, not assumed.

## 6. Known limitations / things a future maintainer should know

- Static labels not driven by the source data (the sidebar date, "Current planning week") were left as in the
  original dashboard - not derived from the SOP data automatically. Worth asking the user whether the planning
  week should follow the latest SOP version if this is revisited.
- The "All Models" SOP-inclusion rule (section 2 above) is a documented reconstruction, not a rediscovered
  original rule - if the true original rule ever surfaces (e.g. from someone who built the original dashboard),
  it should replace the current one in `pipeline/transform.py`.
- The skill-creator's full eval/benchmark/subagent-comparison workflow was deliberately *not* run when building
  `.claude/skills/excel-com-automation/` - it was written directly from this session's verified findings, given
  the project's time constraints. If the skill's triggering or content quality ever needs tuning, that fuller
  process (trigger-phrase evals, blind comparison) is documented in the skill-creator skill itself and can be
  run at that point.
