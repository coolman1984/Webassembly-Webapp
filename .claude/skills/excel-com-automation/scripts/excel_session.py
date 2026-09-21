"""Generic, hardened Excel COM session for Windows (win32com/pywin32).

Copy this into a new project as a starting point rather than re-deriving these rules from scratch. Every
behaviour here is explained in ../references/excel-com-hardening.md and ../references/gotchas.md - read those
for the *why*, not just the *what*, before changing anything.

Usage:

    with ExcelSession() as xl:
        wb = xl.open(r"C:\\path\\to\\file.xlsx")
        for ws_name in sheet_names(wb):
            ws = wb.Worksheets(ws_name)
            lr, lc = last_cell(ws)                 # NOT ws.UsedRange - see gotchas.md #1
            for start_row, rows in iter_used_range(ws):
                ...                                 # rows: list of tuples, bulk Value2 reads, chunked

No workbook is ever saved. Every workbook opened through `.open()` and the session's own blank workbook are
closed on exit, and the Excel process this session started is guaranteed dead when the `with` block ends, even
after an exception - see ExcelSession.__exit__.

Requires: pywin32 (`pip install pywin32`).
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
OPEN_TIMEOUT_S = 30          # a single Workbooks.Open() call must never legitimately take this long
MAX_CELLS_PER_CHUNK = 2_000_000
# Excel error cells (#N/A, #REF!, ...) come back through Value2 as ints in this range.
_XL_ERROR_MIN, _XL_ERROR_MAX = -2146826300, -2146826200
_CO_E_CLASSNOTREG = -2147221005


class ExcelError(RuntimeError):
    """Raised for anything Excel-side that should be shown to a user, not a stack trace."""


def excel_installed() -> bool:
    """Cheap registry check - no Excel process is started."""
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


class ExcelSession:
    """Context manager owning exactly one hidden Excel instance, guaranteed cleaned up."""

    def __init__(self, timeout_s: int = 1200):
        self.app = None
        self._pid: Optional[int] = None
        self._books: list = []
        self._blank = None
        self.timeout_s = timeout_s
        self._timer: Optional[threading.Timer] = None
        self.timed_out = False
        self._open_stuck = False
        self.forced_kill = False
        self.startup_workbooks_closed: list[str] = []   # e.g. an auto-loaded PERSONAL.XLSB - see gotchas.md #6

    # ------------------------------------------------------------------ lifecycle
    def __enter__(self) -> "ExcelSession":
        if not excel_installed():
            raise ExcelError("Microsoft Excel (desktop version) is required but was not found on this computer.")
        self._own_apartment = threading.current_thread() is not threading.main_thread()
        pythoncom.CoInitialize()
        try:
            self.app = win32com.client.DispatchEx("Excel.Application")
        except pywintypes.com_error as e:
            self._release_apartment()
            if e.hresult == _CO_E_CLASSNOTREG:
                raise ExcelError("Microsoft Excel (desktop version) is required but was not found.") from e
            raise ExcelError(f"Excel could not be started: {_com_msg(e)}") from e

        # Macro/PERSONAL.XLSB safety FIRST - see gotchas.md #6 for why the ordering matters.
        for attr, val in (("EnableEvents", False), ("AutomationSecurity", MSO_AUTOMATION_SECURITY_FORCE_DISABLE)):
            try:
                setattr(self.app, attr, val)
            except Exception:
                pass
        try:
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
        self._timer = threading.Timer(self.timeout_s, self._on_timeout)
        self._timer.daemon = True
        self._timer.start()
        for attr, val in (("Visible", False), ("DisplayAlerts", False), ("ScreenUpdating", False),
                          ("Interactive", False), ("AskToUpdateLinks", False)):
            try:
                setattr(self.app, attr, val)
            except Exception:
                pass
        try:  # Calculation can only be changed once a workbook exists.
            self._blank = self.app.Workbooks.Add()
            self.app.Calculation = XL_CALC_MANUAL
        except Exception:
            pass
        return self

    def _release_apartment(self) -> None:
        if getattr(self, "_own_apartment", False):
            pythoncom.CoUninitialize()   # NEVER on a thread other COM callers might share - see gotchas.md #4

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
        gc.collect()                     # drop lingering COM references so Excel can actually exit
        self._reap_own_process()
        self._release_apartment()
        if self.timed_out and exc[0] is not None:
            raise ExcelError(f"Excel did not finish within {self.timeout_s // 60} minutes and was stopped "
                             "(a file may be too large or stuck on a dialog).") from exc[1]
        return False

    def _reap_own_process(self) -> None:
        """If the Excel process we started outlives Quit(), terminate exactly that PID - never any other."""
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
                win32event.WaitForSingleObject(h, 5000)   # termination is asynchronous - wait until it's really gone
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
        """Open a workbook read-only, defensively. See gotchas.md #5 for why the per-open watchdog exists."""
        if not os.path.isfile(path):
            raise ExcelError(f"File not found: {path}")
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
                raise ExcelError(f"Excel could not open '{os.path.basename(path)}' - it needs a dialog answered "
                                 "that nothing can click (for example a wrong-password prompt). "
                                 "Password-protected files are not supported here.") from e
            if self.timed_out:
                raise ExcelError(f"Excel did not finish within {self.timeout_s // 60} minutes and was stopped.") from e
            raise ExcelError(f"Excel could not open '{os.path.basename(path)}' - it may be password-protected, "
                             f"corrupt, not a real Excel file, or locked ({_com_msg(e)}).") from e
        finally:
            open_timer.cancel()
        if wb is None:   # e.g. a file with the same base name is already open elsewhere
            raise ExcelError(f"Excel could not open '{os.path.basename(path)}' - a file with the same name "
                             "may already be open.")
        self._books.append(wb)
        return wb

    def _on_open_timeout(self) -> None:
        self._open_stuck = True
        if self._pid:
            _terminate(self._pid)


# ---------------------------------------------------------------------- sheet helpers
def sheet_names(wb) -> list[str]:
    return [ws.Name for ws in wb.Worksheets]


def is_1904(wb) -> bool:
    try:
        return bool(wb.Date1904)
    except Exception:
        return False


def last_cell(ws) -> tuple[int, int]:
    """(last_row, last_col) that really contain data; (0, 0) for an empty sheet. See gotchas.md #1 and #2:
    NOT UsedRange, and Find() called with POSITIONAL arguments only."""
    def find(order):
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
    return [(v,)] if not isinstance(v, tuple) else list(v)


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
