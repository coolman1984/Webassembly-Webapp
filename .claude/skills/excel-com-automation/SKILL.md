---
name: excel-com-automation
description: >
  Hard-won, verified guidance for automating Microsoft Excel from Python via win32com (COM automation on
  Windows), and for packaging such tools as self-contained, no-install Windows applications for non-technical
  users. Covers hardening an Excel COM session against hangs, zombie EXCEL.EXE processes, unwanted macro/VBA
  execution and PERSONAL.XLSB auto-load, the Find()/UsedRange traps that silently give wrong row/column counts,
  and the password-dialog deadlock that a generic session timeout does not catch; bundling a private Python
  runtime so end users never install anything; securing a localhost HTTP service against DNS rebinding and CSRF;
  and testing an Excel pipeline against real Excel and real file-size limits instead of mocks. Use this whenever
  the user wants to read/write/automate .xlsx, .xlsm or .xlsb files from Python or any COM-capable language on
  Windows, build a tool that ingests Excel uploads into a dashboard/database, package a Python tool so it runs on
  a colleague's PC with zero setup, or run into problems like a hung/zombie EXCEL.EXE process, Workbooks.Open
  never returning, wrong last-row/last-column results, or macros firing when they should not - even if the user
  does not use the word "COM" or "automation" themselves.
---

# Excel COM automation & self-contained Windows tooling

This skill distils lessons from building a production data pipeline that reads three real, messy Excel
workbooks (one nearly 1,050,000 rows) through Excel COM, validates them, and republishes a dashboard - designed
to run on a non-technical colleague's PC with nothing pre-installed but Windows and Excel itself. Every claim
below was hit and fixed in that project, not theorised. Read the reference file named in each section before
writing code in that area; they hold the full detail, code patterns and the reasoning behind each rule.

## Mental model: Excel COM is a hostile, stateful GUI app you are borrowing

`win32com.client.Dispatch`/`DispatchEx` does not give you a clean library - it starts (or reuses) an actual
Excel.exe with its own windows, its own idea of what your users last opened, its own startup files, and its own
opinions about showing dialogs. Every default is tuned for an interactive human sitting in front of it, not for
an unattended script. Treat every property you don't explicitly set as "will surprise you eventually," and treat
every long-running COM call as capable of never returning. That posture is *why* the rules below exist - see
`references/excel-com-hardening.md` for the full, ordered session-setup checklist and the reasoning per step.

## The gotchas that will bite you (verified, not theoretical)

| Symptom | Root cause | Fix |
|---|---|---|
| Wrong "last row/column" from a sheet that looks empty past row 50 but reports rows: 1,048,576 | `UsedRange` includes cells that only ever had *formatting* applied, even with no data or formulas left | Find the real last cell with `Cells.Find("*", ..., SearchDirection=xlPrevious)`, never trust `UsedRange.Rows.Count` |
| `Find()` silently returns the *first* match instead of the last, even with `SearchDirection=xlPrevious` passed | Through **named** COM arguments, late-bound `win32com` can silently drop/ignore some parameters (notably `SearchDirection`) | Call `Find()` with **positional** arguments in the documented VBA order: `Find(What, After, LookIn, LookAt, SearchOrder, SearchDirection)` |
| `EXCEL.EXE` processes pile up in Task Manager after your script exits, one per run | `Application.Quit()` asks Excel to close - it does not guarantee the process actually exits, especially with live COM references still held | Record the PID at launch; after `Quit()`, wait on the process handle with a timeout, then `TerminateProcess` as a fallback - and only ever target the exact PID you started |
| A helper/test script's *own* separate Excel instance suddenly throws "Object is not connected to server" | `pythoncom.CoUninitialize()` tears down the **whole COM apartment of the calling thread**, disconnecting every other live COM object in it, not just the one you meant to release | Only call `CoUninitialize()` on an apartment you created yourself (e.g. a dedicated worker thread); leave the main/shared thread's apartment alone |
| Opening a password-protected (or wrong-password) file hangs for the *entire* session timeout - minutes, not seconds | On some Excel builds/security postures, a wrong `Password:=` does not raise a clean COM error - it pops a modal "password not correct" dialog that `DisplayAlerts`/`AutomationSecurity`/`Interactive` do not suppress, and nothing can click | Never rely on the overall session timeout for this. Wrap **each** `Workbooks.Open()` call in its own short, dedicated watchdog (~30s is generous - even a 38MB real file opens in ~7s) that kills just that stuck attempt |
| A macro runs when you never intended any macro to execute at all | Excel auto-loads whatever sits in its XLSTART folder(s) (typically `PERSONAL.XLSB`) as part of its own process startup, before your script has necessarily had a chance to disable macros | Set `AutomationSecurity = msoAutomationSecurityForceDisable` and `EnableEvents = False` as the *very first* thing on the freshly created `Application` object; then close (unsaved, unread) any workbook already open - in a fresh instance, that can only be an auto-loaded startup file |
| A script that reads Excel files touches/renames the user's real `PERSONAL.XLSB` or XLSTART folder "just to be safe" | Overreach: modifying someone's real, global Excel configuration to test or "guarantee" macro safety is itself the kind of irreversible environment change you're trying to avoid | Never touch global Excel config/files. Verify macro-safety defensively (session-state checks) instead of by engineering a real trigger on a real machine |
| Numbers-as-text, error cells (`#N/A`), or a 1904-dated workbook quietly corrupt your data | `Value2` returns Excel's raw representation: error cells come back as specific negative integers, not blanks; 1904-system dates are offset by 1,462 days from the default 1900 system | Detect the known error-code integer range and treat as blank; check `Workbook.Date1904` and shift accordingly; never assume `Value2`'s type without checking |

Full detail, code, and the "why" behind each row: `references/gotchas.md`.

## Reading data: bulk, chunked, by name

- Read with `Range.Value2` in large rectangular blocks - never cell by cell, and prefer `Value2` over `.Value`
  (skips COM's currency/date auto-formatting so you get raw numbers/serials you convert yourself).
- Chunk very large sheets by rows (tune chunk size so `rows × cols` stays in the low millions of cells) so no
  single COM call is huge and memory stays bounded; this project chunked a 246k-row sheet fine and a synthetic
  1,048,575-row sheet (Excel's hard limit) fine too.
- Locate columns by **normalised header text** (case/whitespace/invisible-Unicode-insensitive), never by fixed
  position - real workbooks get columns inserted, reordered, or renamed slightly, and position-based reads break
  silently while header-based reads keep working.
- When a required header truly is missing, say exactly which one, don't just say "wrong file" - this is the
  single biggest lever for making a non-technical user able to self-diagnose a bad upload.
- Never guess an ambiguous date (`06/07/2026` could be 6 June or 7 June) - accept only unambiguous formats and
  report anything else as unreadable, never silently interpreted one way.

Full patterns: `references/data-extraction.md`.

## Packaging for a non-technical end user: truly no-install

The bar is: someone with no dev tools, no admin rights, and no idea what Python is, double-clicks one `.bat` and
it works, on a PC that has never seen this project before, forever, as long as Windows + desktop Excel exist.
That means bundling a **private, trimmed Python runtime inside the project folder** (interpreter + only the
stdlib/third-party packages actually used, e.g. `pywin32`; strip test suites/Tk/idle; precompile `.pyc`), and a
launcher that always calls `runtime\python.exe -I` (isolated mode: ignores the user's own Python installs,
`PYTHONPATH`, and site-packages entirely) - never a bare `python` that might resolve to something else, or
nothing, on the target machine. Verify this for real: run the bundled runtime with the system's own Python
*and* its `PATH` scrubbed down to bare Windows system directories, and confirm it still works.

Full build recipe and verification method: `references/self-contained-packaging.md`.

## Securing a "double-click, opens in your browser" local service

A local HTTP service is not automatically safe just because it's `127.0.0.1` - any webpage open in the same
browser can still reach it via DNS rebinding or a simple cross-site form POST. Bind to loopback only, validate
the inbound `Host` header against the exact host:port you're listening on for **every** request, and require a
custom header (e.g. `X-Requested-With`) plus an `Origin` check on every state-changing call (a cross-site form
can't set a custom header, so this alone stops the simplest attack class). Also: some corporate endpoint-security
tools intercept the browser's file picker and make the browser unable to read the very file the user just
selected (this is real, not a bug in your code) - offer an "inbox" folder the service reads directly from
disk as a fallback path that sidesteps the browser entirely.

Full checklist: `references/local-service-security.md`.

## Testing: against real Excel, real limits, the real bundled runtime

Mocks cannot reproduce the actual COM quirks above (they're specific to the real automation surface, not to
"what a spreadsheet reader should do"). Generate test fixtures by writing them out through the **same** Excel COM
layer under test, deliberately fuzzing storage details (shuffled columns, noisy headers, extra junk columns,
different date systems, model-code casing/whitespace noise, `.xlsb` vs `.xlsx`) while keeping the underlying data
meaning fixed and independently hand-computed - so a test failure means "this doesn't survive realistic
messiness," not "the fixture generator disagrees with the code that was copy-pasted from it." Run the suite
against the **bundled** `runtime\python.exe -I`, not just a normal dev interpreter - packaging bugs only show up
there. Stress-test to the platform's real hard limit (Excel's 1,048,576-row cap), not an arbitrary "big enough"
number, and measure wall-clock time and peak memory, not just correctness.

When a background test run stalls, don't assume "just slow": a process near-zero CPU over several minutes of
elapsed wall time is *blocked*, not computing - check for that before adding more patience, and check whether the
process still has a live (even if invisible) window handle, which is the signature of a stuck, unanswerable
dialog rather than a crash.

Full strategy and fixture-generation pattern: `references/testing-strategy.md`.

## Ready-to-use code

`scripts/excel_session.py` is a generic, project-agnostic version of the hardened `ExcelSession` context manager
built and proven in this project: single-instance lifecycle, all the settings above in the right order, the
per-open watchdog, the real-last-cell/bulk-read helpers, and guaranteed process cleanup even after an exception.
Copy it as a starting point rather than re-deriving these rules from scratch.
