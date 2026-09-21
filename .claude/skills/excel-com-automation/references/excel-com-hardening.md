# The hardened Excel COM session, step by step

This is the exact order that matters, and why. See `scripts/excel_session.py` for the runnable version.

## 1. Start one Excel instance per run, with `DispatchEx`

Use `win32com.client.DispatchEx("Excel.Application")`, not `Dispatch` - `DispatchEx` always creates a *new*
Excel process, rather than possibly attaching to (and disrupting) an Excel window the user already has open with
their own unsaved work. Open every workbook for the run through this one instance; don't start a fresh instance
per file.

## 2. Disable macros and events *first*, before anything else touches the Application object

```python
app = win32com.client.DispatchEx("Excel.Application")
app.EnableEvents = False
app.AutomationSecurity = 3   # msoAutomationSecurityForceDisable
```

This must be the very first thing done, before `Visible`, `DisplayAlerts`, or anything else - see gotcha #6 in
`gotchas.md` for why the ordering itself matters (Excel's own startup file-loading is a race you want to win as
early as possible).

## 3. Close anything Excel opened on its own

```python
for wb in list(app.Workbooks):
    name = wb.Name
    wb.Close(SaveChanges=False)   # never read, never save - just close it
```

In a freshly created, isolated instance, the only way a workbook can already be open at this point is an
XLSTART auto-load. Log/report if this ever fires (it should be rare/never) rather than silencing it - it's a
signal worth a human seeing.

## 4. Record the PID before doing anything else

```python
_, pid = win32process.GetWindowThreadProcessId(app.Hwnd)
```

You'll need this exact PID for cleanup (section 7) and for a crash-recovery pidfile if the whole process
(yours, not just Excel) might itself be killed mid-run - write `{"pid": pid, "started": <process start time>}`
to disk so a *later* run can find and clean up an orphan, matched by both PID and start-time (PIDs get reused;
start-time disambiguates).

## 5. The rest of the session-level settings

```python
app.Visible = False
app.DisplayAlerts = False
app.ScreenUpdating = False
app.Interactive = False
app.AskToUpdateLinks = False
```

Then add one blank workbook and switch calculation to manual - `Application.Calculation` can only be changed
once a workbook exists:

```python
blank = app.Workbooks.Add()
app.Calculation = -4135   # xlCalculationManual - you're reading cached values, not recalculating formulas
```

## 6. Open workbooks defensively

```python
DUMMY_PASSWORD = "\u0001no-password-supplied\u0001"

wb = app.Workbooks.Open(
    path, UpdateLinks=0, ReadOnly=True,
    Password=DUMMY_PASSWORD, WriteResPassword=DUMMY_PASSWORD,
    IgnoreReadOnlyRecommended=True, AddToMru=False,
)
```

- `UpdateLinks=0` - never prompt about (or silently trigger) updating external links to other files.
- `ReadOnly=True` - you're extracting, not editing; also avoids some lock-file prompts.
- `Password=`/`WriteResPassword=` a harmless dummy value even for files you don't expect to be protected - it's
  a no-op for a normal file, but for a genuinely protected one it steers Excel toward "wrong password, fail"
  instead of "no password given, prompt the (invisible) user forever."
- `IgnoreReadOnlyRecommended=True` - suppresses the "this file is recommended to be opened read-only, proceed?"
  prompt for files saved with that flag.
- `AddToMru=False` - don't pollute the real user's Recent Files list with files your automation touched.
- Wrap this call in its own short watchdog timer (~30s) - see gotcha #5. This is separate from, and much
  shorter than, the overall session timeout.
- Check for `wb is None` - Excel can silently refuse to return a workbook object (e.g. a file with the same
  base name is already open elsewhere) rather than raising.

## 7. Guaranteed cleanup, even after an exception

In a context manager's `__exit__` (or an equivalent `finally`):

```python
for wb in books:
    try: wb.Close(SaveChanges=False)
    except Exception: pass
try: blank.Close(SaveChanges=False)
except Exception: pass
try: app.Quit()
except Exception: pass
del app  # drop the last Python reference
gc.collect()
# then: wait on the recorded PID's process handle with a short timeout; TerminateProcess if it's still alive
```

Every step is independently wrapped - a failure closing one workbook must not skip closing the others or
stop `Quit()` from being attempted. The final PID-based wait-then-terminate is what actually guarantees no
zombie process, regardless of how gracefully (or not) `Quit()` behaved.

## 8. One thread, one COM apartment, unless you know why not

Call `pythoncom.CoInitialize()`/`CoUninitialize()` only around a dedicated worker thread your own code fully
owns. Never call `CoUninitialize()` on a thread another part of the program (or a caller) might also be using
COM on - see gotcha #4. If in doubt, only uninitialise apartments you created for exactly this purpose and
nothing else runs on.
