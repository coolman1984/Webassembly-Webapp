# Excel COM gotchas: symptom, cause, fix, evidence

Every entry here was actually hit while building the reference pipeline, diagnosed to its real cause (not
patched around blindly), and verified fixed. Where a fix trades off against something (e.g. a shorter timeout
risking a false positive on a legitimately huge file), the verification method is given too.

## 1. `UsedRange` lies about how big a sheet really is

**Symptom:** `ws.UsedRange.Rows.Count` reports over a million rows on a sheet that visibly has three rows of
data.

**Cause:** Excel's `UsedRange` includes any cell that was ever *formatted* - a fill colour, a border, a
conditional format - even if its value and formula were later cleared. A single cell coloured at row 1,000,000
inflates the whole sheet's reported extent, and this survives saving and reloading the file.

**Fix:** Never use `UsedRange` to find the real extent of data. Use `Find` searching backwards for any non-blank
cell:

```python
XL_FORMULAS, XL_BY_ROWS, XL_BY_COLUMNS, XL_PREVIOUS, XL_PART = -4123, 1, 2, 2, 2

def last_cell(ws):
    def find(order):
        # POSITIONAL args - see gotcha #2 below for why this matters
        return ws.Cells.Find("*", ws.Cells(1, 1), XL_FORMULAS, XL_PART, order, XL_PREVIOUS)
    r = find(XL_BY_ROWS)
    if r is None:
        return 0, 0          # truly empty sheet
    c = find(XL_BY_COLUMNS)
    return int(r.Row), int(c.Column)
```

**Verification:** a fixture sheet with real data in `A1:C3` and only a fill colour applied at `A1000000` -
`UsedRange.Rows.Count` reports 900,000+, `last_cell()` correctly reports `(3, 3)`.

## 2. `Find()`'s named arguments can silently do the wrong thing

**Symptom:** Code that explicitly passes `SearchDirection=xlPrevious` (searching backwards, for the *last*
match) instead returns the *first* match in the sheet - the opposite of what was asked, with no error.

**Cause:** Through `win32com`'s late (dynamic) binding, calling a COM method with Python keyword arguments
relies on the method's named-parameter metadata being resolved correctly at runtime. For `Range.Find`, this
resolution can silently drop or misapply `SearchDirection` when called with keywords, while the call still
"succeeds" and returns *a* match - just the wrong one. This is worse than a crash: it produces a plausible,
wrong answer.

**Fix:** Call `Find` with **positional** arguments, in the exact order VBA documents them:
`Find(What, After, LookIn, LookAt, SearchOrder, SearchDirection)`. Positional binding does not go through the
same named-parameter resolution path and behaves correctly.

**Verification:** on the same fixture, the keyword-argument call returned `(row=1, col=2)`; the positional call
returned `(row=3, col=3)` - the actual last cell. Always sanity-check a `Find`-based helper against a fixture
where "first" and "last" are known to differ.

## 3. Zombie `EXCEL.EXE` processes after the script exits

**Symptom:** Task Manager accumulates one `EXCEL.EXE` process per pipeline run, none of them ever closing on
their own, silently eating memory until the machine is rebooted.

**Cause:** `Application.Quit()` is a *request* to close, dispatched through COM - it does not synchronously wait
for (or guarantee) the actual process exit, particularly when Python-side COM references to the Application or
its child objects (Workbooks, Worksheets, Ranges) are still alive and hold the process open, or when an
exception skips the cleanup path.

**Fix:**
- Record the exact PID at launch (`win32process.GetWindowThreadProcessId(app.Hwnd)`).
- In a `finally`/context-manager `__exit__`, close every workbook (`SaveChanges=False`), release Python
  references (`del`, `gc.collect()`), then call `Quit()`.
- Then wait on that specific process handle with a timeout (a few seconds is plenty); if it hasn't exited,
  `TerminateProcess` it.
- **Always target the exact PID you started, via a process handle, never a name-based "kill all EXCEL.EXE"** -
  the user may have their own, unrelated Excel windows open, and killing those is a much worse bug than a
  zombie process.
- Store that PID (with its process start-time, to disambiguate PID reuse) on disk; on the next run, check for
  and clean up any orphan left by a crash of the *previous* run - matched by both PID and start-time, never PID
  alone.

**Verification:** loop-launched 15 sessions back to back, checking after each one whether *that session's own*
PID was still alive; forcing an exception mid-session (with live COM references held at the moment of the
exception) still resulted in zero leaked processes.

## 4. `CoUninitialize()` breaks *other* live COM objects on the same thread

**Symptom:** A test helper that creates its own separate Excel instance (e.g. to build fixture files) starts
throwing `pywintypes.com_error: Object is not connected to server` right after the code under test finishes a
session - even though the two Excel instances are otherwise unrelated.

**Cause:** `pythoncom.CoUninitialize()` tears down the calling thread's entire COM apartment - every COM
interface pointer held by that thread, for every object, becomes invalid at once. It is not scoped to "the
object I'm done with."

**Fix:** Only call `CoUninitialize()` on an apartment your own code created for itself (e.g. inside a dedicated
worker thread that owns nothing else). If your code might run on a thread shared with other callers (the main
thread, in particular), skip `CoUninitialize()` there and let the process's normal shutdown handle it.

**Verification:** started a "fixture" Excel instance, then ran the code-under-test's own `ExcelSession` on the
same thread and let it exit normally; before the fix, the fixture instance became unusable immediately
afterward; after the fix, it kept working for the rest of the process's lifetime.

## 5. A wrong/dummy password can hang instead of failing

**Symptom:** Opening a password-protected file (deliberately, to test rejection) with a wrong password doesn't
raise an error - the whole session hangs for however long the overall timeout is (in this project, that
defaulted to 20 minutes before the cause was found).

**Cause:** Passing `Password:=` explicitly to `Workbooks.Open` is documented to make Excel raise an error on a
wrong password rather than prompting - and it usually does. But on some Excel builds/security postures it
instead shows a real, modal "The password you supplied is not correct" dialog. `DisplayAlerts=False`,
`Interactive=False`, and `AutomationSecurity` do **not** reliably suppress this specific dialog, because it
originates from the file-format/security validation layer, not from the "Excel automation alert" layer those
settings govern. With `Visible=False`, the dialog exists but nothing can see or click it - the COM call blocks
forever.

**Diagnostic technique that found this:** two independent full test-suite runs both stalled at the exact same
point; `Get-Process` showed the Python process alive with near-zero CPU time over several minutes (blocked, not
computing) and an Excel process with `Responding: True` and a real (if invisible) window handle - the signature
of "waiting on an unanswerable dialog," as opposed to `Responding: False` (crashed/deadlocked message pump).

**Fix:** Never rely on the overall session timeout for this - it's much too coarse. Wrap *each*
`Workbooks.Open()` call in its own short, dedicated watchdog (a `threading.Timer`, same safe pattern as the
zombie-process watchdog: it only calls a plain Win32 `TerminateProcess`, never touches COM from the timer
thread) with a bound like 30 seconds - generous, since real files (even a genuine 38MB workbook) open in single
digits of seconds; only a stuck dialog would ever hit it. On timeout, kill just this Excel process, let the
blocked COM call fail with an RPC error, and translate that into a specific, accurate message rather than the
generic "could not open" one.

**Verification:** the real 38MB workbook's `Workbooks.Open()` alone measured ~7 seconds; the induced-hang test
went from "never returns" to failing cleanly in ~38 seconds with the fix, safely under a 60-second acceptance
bound and nowhere near the old 20-minute session timeout.

## 6. Auto-loaded startup files (PERSONAL.XLSB) can run macros you never asked for

**Symptom:** Not directly observed as a failure in this project (no real `PERSONAL.XLSB` with macros existed on
the build machine) - documented here as defence-in-depth because the failure mode, if it did occur, would be
silent and serious: unrequested VBA execution.

**Cause:** When Excel.exe starts (including when started via COM automation), it loads whatever workbooks sit
in its XLSTART folder(s) - conventionally `PERSONAL.XLSB`, the user's personal macro workbook - as part of its
own startup sequence, and can fire that workbook's `Workbook_Open`/`Auto_Open` macro. If your script sets
`AutomationSecurity`/`EnableEvents` *after* other setup work, there is a race: Excel's own startup file-loading
might complete, and fire a macro, before your script's disabling code runs.

**Fix:** On the freshly created `Application` object, set `EnableEvents = False` and
`AutomationSecurity = msoAutomationSecurityForceDisable` as the very first two lines - before `Visible`,
`DisplayAlerts`, anything else. Then, before doing anything else, enumerate `Workbooks`: in a brand-new isolated
instance you haven't opened anything in yet, the *only* way something can already be open is exactly this
auto-load. Close it immediately, unsaved, without reading it.

**Do not "verify" this by creating a real macro-enabled `PERSONAL.XLSB` on a real machine and checking whether
its macro ran.** That is itself an irreversible change to someone's actual Excel environment and exactly the
kind of thing this defence exists to prevent doing accidentally elsewhere. Verify defensively instead: inspect
the session's own `AutomationSecurity`/`EnableEvents` property values, and that the "any workbook already open"
list is empty on a clean machine.

## 7. `Value2` error cells and date-system differences

**Symptom:** A cell showing `#N/A` in Excel comes back through `Value2` as a large negative integer (not a
string, not `None`); dates in some workbooks are consistently ~4 years off.

**Cause:** `Value2` gives you Excel's internal representation, unfiltered. Error values are encoded as specific
negative integers in a known range (roughly -2146826300 to -2146826200). Dates are day-count serials from an
epoch that depends on the workbook's `Date1904` setting (the default "1900 system" epoch is 1899-12-30; the
rarer "1904 system," mainly from old Mac-authored files, is offset by exactly 1,462 days).

**Fix:** Treat any integer in the known error-code range as blank/`None`, not as a literal number. Check
`Workbook.Date1904` once per workbook and add the 1,462-day offset to any date serial before converting, when
set.

## 8. Position-based column reads break on real-world files

**Symptom:** A pipeline that reads "column F" works on the sample file and then silently reads the wrong data
(or crashes) on the next month's export.

**Cause:** Real recurring exports get columns inserted, reordered, or renamed slightly release to release -
this is normal, not user error.

**Fix:** Locate every column by matching its header text after normalising away case, extra whitespace,
newlines, and invisible Unicode characters (zero-width spaces, NBSP, BOM) - never by fixed index. When a
required header genuinely can't be found, say exactly which one is missing in the rejection message, not a
generic "this isn't the right file" - it lets a non-technical user fix their own export instead of escalating
to support.

## 9. Ambiguous dates must not be guessed

**Symptom:** A date like `06/07/2026` gets silently parsed as one of June-7 or July-6, and it's a coin flip
which - wrong exactly as often as the format assumption doesn't match the source file's locale.

**Fix:** Only parse unambiguous formats (ISO `YYYY-MM-DD`/`YYYY/MM/DD`, or a format with a named month like
`06-Jul-2026`) plus real Excel date serials/`datetime` objects. Anything else - including any bare
`DD/MM/YYYY`-shaped string - is reported as unreadable rather than guessed, and surfaced as a warning so a human
can check the source file.
