"""Excel COM access layer.

One Excel instance (DispatchEx) is started per pipeline run, every workbook is opened once with
ReadOnly=True / UpdateLinks=0, and sheets are read with bulk Range.Value2 calls (row-chunked) - never cell by cell.

Hardening (from Excel-automation best practice):
  * Macros are force-disabled (AutomationSecurity) and Application events are turned off as the very first thing
    done to the Application object - before Visible/DisplayAlerts, before anything else. If Excel auto-loaded a
    startup workbook (PERSONAL.XLSB or anything else in an XLSTART folder), it is closed immediately, unsaved,
    without being read or otherwise touched. This pipeline never runs any VBA and never opens PERSONAL.XLSB on
    purpose; this block is defence-in-depth against Excel opening it on its own during process start.
  * DisplayAlerts/Interactive/AskToUpdateLinks off, calculation switched to manual *before* any real workbook is
    opened (a blank workbook is added first so the calculation setting is allowed to be changed).
  * A dummy Password is always passed on open, and a short dedicated watchdog (OPEN_TIMEOUT_S) guards the open
    call itself: on most machines a wrong password makes Excel raise an error immediately, but on some
    Excel builds/security postures it instead shows a modal "password not correct" dialog that DisplayAlerts
    does not suppress and that nothing can click - the watchdog kills our own Excel in that case so the
    call fails fast with a clear message instead of hanging.
  * The real last row/column comes from Find(xlPrevious), not UsedRange (formatting can inflate UsedRange to
    row 1,048,576).
  * A watchdog terminates *our own* Excel process if a run exceeds the time limit (hung dialog, huge file).
  * The PID is recorded on disk so a crashed run's orphan can be cleaned up on the next start.
"""
from __future__ import annotations

import gc
import json
import os
import threading
import time
from typing import Iterator, Optional

import pythoncom
import pywintypes
import win32api
import win32com.client
import win32con
import win32event
import win32process

XL_CALC_MANUAL = -4135
XL_FORMULAS, XL_BY_ROWS, XL_BY_COLUMNS, XL_PREVIOUS, XL_PART = -4123, 1, 2, 2, 2
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3
DUMMY_PASSWORD = "\u0001no-password-supplied\u0001"
OPEN_TIMEOUT_S = 30   # a single Workbooks.Open() call must never legitimately take this long; see ExcelSession.open()
# Excel error cells (#N/A, #REF!, ...) come back through Value2 as these negative ints.
_XL_ERROR_MIN, _XL_ERROR_MAX = -2146826300, -2146826200
_CO_E_CLASSNOTREG = -2147221005
MAX_CELLS_PER_CHUNK = 2_000_000


class ExcelError(RuntimeError):
    pass


def excel_installed() -> bool:
    """True if the Excel COM class is registered (cheap registry check, no Excel start)."""
    if os.environ.get("BOM_FORCE_NO_EXCEL"):          # test hook: simulate a PC without Excel
        return False
    try:
        import winreg
        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Excel.Application\CLSID"))
        return True
    except OSError:
        return False


def clean_cell(v):
    """Normalise a Value2 cell: Excel error codes -> None, '' -> None."""
    if v is None:
        return None
    if isinstance(v, int) and not isinstance(v, bool) and _XL_ERROR_MIN <= v <= _XL_ERROR_MAX:
        return None
    if isinstance(v, str) and v == "":
        return None
    return v


def _com_msg(e: Exception) -> str:
    try:
        ex = getattr(e, "excepinfo", None)
        if ex and ex[2]:
            return str(ex[2]).strip()
    except Exception:
        pass
    return str(e)


def _proc_start_time(pid: int) -> Optional[int]:
    try:
        h = win32api.OpenProcess(win32con.PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
    except Exception:
        return None
    try:
        t = win32process.GetProcessTimes(h)["CreationTime"]
        return int(t.timestamp() * 1000) if hasattr(t, "timestamp") else int(t)
    except Exception:
        return None
    finally:
        win32api.CloseHandle(h)


def _terminate(pid: int) -> None:
    try:
        h = win32api.OpenProcess(win32con.PROCESS_TERMINATE, False, pid)
    except Exception:
        return
    try:
        win32api.TerminateProcess(h, 1)
    except Exception:
        pass
    finally:
        win32api.CloseHandle(h)


def kill_orphan(pidfile: str) -> Optional[int]:
    """Terminate an Excel process left behind by a crashed run (matched by PID *and* start time)."""
    try:
        rec = json.load(open(pidfile, encoding="utf-8"))
        pid, started = int(rec["pid"]), rec["started"]
    except Exception:
        return None
    finally:
        try:
            os.remove(pidfile)
        except OSError:
            pass
    if _proc_start_time(pid) == started:
        _terminate(pid)
        return pid
    return None


class ExcelSession:
    """Context manager owning a single hidden Excel instance."""

    def __init__(self, timeout_s: int = 1200, pidfile: Optional[str] = None):
        self.app = None
        self._pid: Optional[int] = None
        self._books: list = []
        self._blank = None
        self.timeout_s = timeout_s
        self.pidfile = pidfile
        self._timer: Optional[threading.Timer] = None
        self.timed_out = False
        self._open_stuck = False
        self.startup_workbooks_closed: list[str] = []   # names of any workbook Excel auto-loaded (e.g. PERSONAL.XLSB)

    # ------------------------------------------------------------------ lifecycle
    def __enter__(self) -> "ExcelSession":
        if not excel_installed():
            raise ExcelError("Microsoft Excel (desktop version) is required on this computer but was not found.")
        # CoUninitialize() tears down the *whole* COM apartment of the thread, which disconnects every other live
        # COM object in it (e.g. another Excel the caller holds). So only uninitialise an apartment we created
        # ourselves (a worker thread); on the main thread we leave it alone.
        self._own_apartment = threading.current_thread() is not threading.main_thread()
        pythoncom.CoInitialize()
        try:
            self.app = win32com.client.DispatchEx("Excel.Application")
        except pywintypes.com_error as e:
            self._release_apartment()
            if e.hresult == _CO_E_CLASSNOTREG:
                raise ExcelError("Microsoft Excel (desktop version) is required on this computer but was not found.") from e
            raise ExcelError(f"Excel could not be started: {_com_msg(e)}") from e
        # --- macro/PERSONAL.XLSB safety: this block runs FIRST, before anything else touches self.app, to close
        # the race window as tightly as possible. Excel.Application, on process start, auto-loads whatever sits in
        # its XLSTART folder(s) - normally that is exactly PERSONAL.XLSB, the user's personal macro workbook, and
        # doing so can fire a Workbook_Open/Auto_Open VBA macro. We never want that to run as part of this pipeline.
        for attr, val in (("EnableEvents", False), ("AutomationSecurity", MSO_AUTOMATION_SECURITY_FORCE_DISABLE)):
            try:
                setattr(self.app, attr, val)
            except Exception:
                pass
        try:
            # In a freshly created DispatchEx instance the ONLY way a workbook can already be open here is if
            # Excel auto-loaded it from an XLSTART folder - we never opened anything yet ourselves. Close it
            # immediately, unsaved, without reading or otherwise interacting with it any further.
            for wb in list(self.app.Workbooks):
                name = wb.Name
                wb.Close(SaveChanges=False)
                self.startup_workbooks_closed.append(name)
        except Exception:
            pass
        try:
            _, self._pid = win32process.GetWindowThreadProcessId(self.app.Hwnd)
        except Exception:
            self._pid = None
        if self.pidfile and self._pid:
            try:
                json.dump({"pid": self._pid, "started": _proc_start_time(self._pid)}, open(self.pidfile, "w", encoding="utf-8"))
            except OSError:
                pass
        self._timer = threading.Timer(self.timeout_s, self._on_timeout)
        self._timer.daemon = True
        self._timer.start()
        for attr, val in (("Visible", False), ("DisplayAlerts", False), ("ScreenUpdating", False),
                          ("Interactive", False), ("AskToUpdateLinks", False)):
            try:
                setattr(self.app, attr, val)
            except Exception:
                pass
        try:  # calculation can only be changed while a workbook exists
            self._blank = self.app.Workbooks.Add()
            self.app.Calculation = XL_CALC_MANUAL
        except Exception:
            pass
        return self

    def _release_apartment(self) -> None:
        if getattr(self, "_own_apartment", False):
            pythoncom.CoUninitialize()

    def _on_timeout(self) -> None:
        self.timed_out = True
        if self._pid:
            _terminate(self._pid)

    def __exit__(self, *exc):
        if self._timer:
            self._timer.cancel()
        for wb in self._books:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        self._books.clear()
        try:
            if self._blank is not None:
                self._blank.Close(SaveChanges=False)
        except Exception:
            pass
        self._blank = None
        try:
            if self.app is not None:
                self.app.Quit()
        except Exception:
            pass
        self.app = None
        gc.collect()                      # drop lingering COM references so Excel can exit
        self._reap_own_process()
        if self.pidfile:
            try:
                os.remove(self.pidfile)
            except OSError:
                pass
        self._release_apartment()
        if self.timed_out and exc[0] is not None:
            raise ExcelError(f"Excel did not finish within {self.timeout_s // 60} minutes and was stopped "
                             "(a file may be too large or stuck on a dialog).") from exc[1]
        return False

    def _reap_own_process(self):
        """If the Excel process we started outlives Quit(), terminate exactly that PID (never others)."""
        self.forced_kill = False
        if not self._pid:
            return
        try:
            h = win32api.OpenProcess(win32con.SYNCHRONIZE | win32con.PROCESS_TERMINATE, False, self._pid)
        except Exception:
            return  # already gone
        try:
            if win32event.WaitForSingleObject(h, 8000) == win32event.WAIT_TIMEOUT:
                win32api.TerminateProcess(h, 1)
                self.forced_kill = True
                win32event.WaitForSingleObject(h, 5000)      # termination is asynchronous: wait until it is really gone
        finally:
            win32api.CloseHandle(h)

    def own_process_alive(self) -> bool:
        if not self._pid:
            return False
        try:
            h = win32api.OpenProcess(win32con.SYNCHRONIZE, False, self._pid)
        except Exception:
            return False
        try:
            return win32event.WaitForSingleObject(h, 0) == win32event.WAIT_TIMEOUT
        finally:
            win32api.CloseHandle(h)

    # ------------------------------------------------------------------ workbooks
    def open(self, path: str):
        if not os.path.isfile(path):
            raise ExcelError(f"File not found: {path}")
        # A single Workbooks.Open() call should never legitimately take anywhere near this long - even the
        # largest real workbooks in this project open in a few seconds (reading their contents is the slow
        # part, and happens afterwards, outside this call). If it does not return in time, the most likely
        # cause is Excel showing a modal dialog that our automation settings cannot suppress or answer (for
        # example a wrong/dummy-password prompt on some Excel builds/security postures - DisplayAlerts and
        # AutomationSecurity do not reliably suppress that specific dialog). We do not wait for the full
        # session timeout in that case: a short, dedicated watchdog kills our own Excel process so the blocked
        # COM call fails fast instead of hanging for minutes with nobody able to click the invisible dialog.
        self._open_stuck = False
        open_timer = threading.Timer(min(OPEN_TIMEOUT_S, self.timeout_s), self._on_open_timeout)
        open_timer.daemon = True
        open_timer.start()
        try:
            wb = self.app.Workbooks.Open(
                os.path.abspath(path), UpdateLinks=0, ReadOnly=True, Password=DUMMY_PASSWORD,
                WriteResPassword=DUMMY_PASSWORD, IgnoreReadOnlyRecommended=True, AddToMru=False)
        except pywintypes.com_error as e:
            if self._open_stuck:
                raise ExcelError(f"Excel could not open '{_disp(path)}' - it needs a dialog answered that this "
                                 "program cannot click for you (for example a wrong-password prompt); "
                                 "password-protected files are not supported here. Remove the password and try again.") from e
            if self.timed_out:
                raise ExcelError(f"Excel did not finish within {self.timeout_s // 60} minutes and was stopped.") from e
            raise ExcelError(f"Excel could not open '{_disp(path)}' - it may be password-protected, corrupt, "
                             f"not a real Excel file, or locked ({_com_msg(e)}).") from e
        finally:
            open_timer.cancel()
        if wb is None:   # Excel refuses (returns nothing) e.g. when a workbook with the same file name is already open
            raise ExcelError(f"Excel could not open '{_disp(path)}' - a file with the same name may already be open.")
        self._books.append(wb)
        return wb

    def _on_open_timeout(self) -> None:
        self._open_stuck = True
        if self._pid:
            _terminate(self._pid)


def _disp(path: str) -> str:
    import re
    return re.sub(r"^slot\d__", "", os.path.basename(path))


# ---------------------------------------------------------------------- sheet helpers
def sheet_names(wb) -> list[str]:
    return [ws.Name for ws in wb.Worksheets]


def is_1904(wb) -> bool:
    try:
        return bool(wb.Date1904)
    except Exception:
        return False


def last_cell(ws) -> tuple[int, int]:
    """(last_row, last_col) that really contain data; (0, 0) for an empty sheet. Ignores formatting-only cells."""
    def find(order):
        # POSITIONAL arguments on purpose: through late-bound COM the named form silently ignored
        # SearchDirection and returned the first cell instead of the last.
        # Find(What, After, LookIn, LookAt, SearchOrder, SearchDirection)
        try:
            return ws.Cells.Find("*", ws.Cells(1, 1), XL_FORMULAS, XL_PART, order, XL_PREVIOUS)
        except pywintypes.com_error:
            return None
    r = find(XL_BY_ROWS)
    if r is None:
        return 0, 0
    c = find(XL_BY_COLUMNS)
    return int(r.Row), int(c.Column)


def read_block(ws, row: int, col: int, nrows: int, ncols: int) -> list[tuple]:
    """Bulk-read a rectangle via one Range.Value2 call. Always returns a list of row tuples."""
    if nrows <= 0 or ncols <= 0:
        return []
    v = ws.Range(ws.Cells(row, col), ws.Cells(row + nrows - 1, col + ncols - 1)).Value2
    if not isinstance(v, tuple):  # single cell
        return [(v,)]
    return list(v)


def read_top(ws, nrows: int = 12, max_cols: int = 3000) -> list[tuple]:
    """Header probe: the first rows of the sheet, A1-anchored, as wide as the real data (bounded)."""
    lr, lc = last_cell(ws)
    return read_block(ws, 1, 1, min(nrows, lr), min(max_cols, lc))


def iter_used_range(ws, chunk_rows: int = 15000) -> Iterator[tuple[int, list[tuple]]]:
    """Yield (first_sheet_row, rows) chunks covering all real data, each a bulk Value2 read (A1-anchored)."""
    lr, lc = last_cell(ws)
    if lr == 0:
        return
    step = max(500, min(chunk_rows, MAX_CELLS_PER_CHUNK // max(lc, 1)))
    for s in range(1, lr + 1, step):
        n = min(step, lr - s + 1)
        yield s, read_block(ws, s, 1, n, lc)
