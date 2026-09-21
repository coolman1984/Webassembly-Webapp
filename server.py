"""Local dashboard service (offline, 127.0.0.1 only).

    runtime\\python.exe server.py        (start_dashboard.bat does this and opens the browser)

  GET  /                       the dashboard
  GET  /api/data               current datasets (from memory; persisted in the local SQLite database)
  GET  /api/meta               {updated_at} - cheap change check
  GET  /api/status             pipeline job state (idle | running | done | error) + log + excel availability
  PUT  /api/upload/<1|2|3>     raw file body, X-Filename header  (one call per attached file)
  POST /api/process            validate + extract (Excel COM) + rebuild database

Security: bound to loopback; every request must carry a loopback Host header (DNS-rebinding defence);
state-changing calls also need the same-origin Origin and an X-Requested-With header.
"""
from __future__ import annotations

import json
import os
import re
import http.client
import shutil
import sys
import threading
import time
import traceback
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, unquote, urlparse

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)
from pipeline import excel_com as X          # noqa: E402
from pipeline import run as P                # noqa: E402
from pipeline import sources as S            # noqa: E402
from pipeline import store                   # noqa: E402

APP_ID = "bom-confirmation-dashboard"
BASE_PORT, PORT_TRIES = 8765, 20
HOST = "127.0.0.1"
HTML = os.path.join(ROOT, "BOM_Confirmation_Plan_SYSTEM_STATUS_FONT_MATCH.html")
DATA_DIR = os.environ.get("BOM_DATA_DIR") or os.path.join(ROOT, "data")
UPLOAD_DIR = os.path.join(DATA_DIR, "uploads")
RUN_DIR = os.path.join(DATA_DIR, "runs")
DB_PATH = os.path.join(DATA_DIR, "dashboard.db")
LOG_PATH = os.path.join(DATA_DIR, "pipeline.log")
PID_PATH = os.path.join(DATA_DIR, "excel.pid")
INBOX_DIR = os.environ.get("BOM_INBOX_DIR") or os.path.join(ROOT, "inbox")
MAX_UPLOAD = 1 << 30      # 1 GB per file
JOB_TIMEOUT_S = int(os.environ.get("BOM_JOB_TIMEOUT", "1200"))
CSRF_HEADER = "bom-dashboard"

PORT = BASE_PORT
_lock = threading.Lock()
_job = {"state": "idle", "log": [], "error": None, "started": None, "finished": None, "meta": None}
_cache = {"body": None, "etag": None, "updated": None}


def _log(msg: str) -> None:
    line = f"{time.strftime('%H:%M:%S')}  {msg}"
    _job["log"].append(line)
    try:
        if os.path.isfile(LOG_PATH) and os.path.getsize(LOG_PATH) > 5_000_000:
            os.replace(LOG_PATH, LOG_PATH + ".old")
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except OSError:
        pass


def _load_cache() -> None:
    """(Re)load the datasets from the database into memory; a missing/corrupt database just means 'no data'."""
    try:
        d = store.read_datasets(DB_PATH)
    except ValueError as e:
        _cache.update(body=None, etag=None, updated=None)
        if os.path.isfile(DB_PATH):
            _log(f"Existing database ignored: {e}")
        return
    body = json.dumps(d, ensure_ascii=False, default=str).encode("utf-8")
    _cache.update(body=body, etag='"%s"' % d["meta"].get("updated_at", ""), updated=d["meta"].get("updated_at"))


def _inbox_files() -> list[str]:
    """Workbooks the user dropped into the inbox folder (read in place, never modified or deleted)."""
    if not os.path.isdir(INBOX_DIR):
        return []
    return sorted(os.path.join(INBOX_DIR, f) for f in os.listdir(INBOX_DIR)
                  if f.lower().endswith(P.ALLOWED_EXT) and not f.startswith("~$") and os.path.isfile(os.path.join(INBOX_DIR, f)))


def _run_job(paths: list[str], run_dir: str | None) -> None:
    try:
        if not X.excel_installed():
            raise X.ExcelError("Microsoft Excel (desktop version) is required on this computer but was not found.")
        meta = P.process(paths, DB_PATH, _log, pidfile=PID_PATH, timeout_s=JOB_TIMEOUT_S)
        _load_cache()
        _job.update(state="done", meta=meta, error=None)
    except (S.ValidationError, X.ExcelError) as e:
        _log(f"REJECTED: {e}")
        _job.update(state="error", error=str(e))
    except Exception as e:                      # unexpected: keep the trace in the log, show a short message
        _log("FAILED:\n" + traceback.format_exc())
        _job.update(state="error", error=f"Unexpected error ({type(e).__name__}: {e}). Details are in data\\pipeline.log")
    finally:
        _job["finished"] = time.time()
        if run_dir:                                  # uploaded copies are not needed once processed (inbox files stay)
            shutil.rmtree(run_dir, ignore_errors=True)


def _staged_uploads() -> list[str]:
    if not os.path.isdir(UPLOAD_DIR):
        return []
    out = []
    for slot in (1, 2, 3):
        hits = [f for f in os.listdir(UPLOAD_DIR) if f.startswith(f"slot{slot}__")]
        if hits:
            out.append(os.path.join(UPLOAD_DIR, hits[0]))
    return out


class Handler(BaseHTTPRequestHandler):
    server_version = "BOMDashboard/2.0"
    timeout = 300                                  # a stalled client cannot hold a thread forever

    def log_message(self, fmt, *args):             # quiet console
        pass

    # ------------------------------------------------------------------ helpers
    def _send(self, code: int, body: bytes, ctype: str, extra: dict | None = None) -> None:
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(body)

    def _json(self, code: int, obj) -> None:
        self._send(code, json.dumps(obj, ensure_ascii=False, default=str).encode("utf-8"), "application/json; charset=utf-8")

    def _host_ok(self) -> bool:
        return (self.headers.get("Host") or "").lower() in (f"{HOST}:{PORT}", f"localhost:{PORT}")

    def _write_ok(self) -> bool:
        o = self.headers.get("Origin")
        return (o is None or o in (f"http://{HOST}:{PORT}", f"http://localhost:{PORT}")) \
            and self.headers.get("X-Requested-With") == CSRF_HEADER

    def _guard(self, write: bool) -> bool:
        if not self._host_ok():
            self._json(403, {"error": "forbidden host"})
            return False
        if write and not self._write_ok():
            self._json(403, {"error": "forbidden origin"})
            return False
        return True

    def _safe(self, fn) -> None:
        try:
            fn()
        except (ConnectionError, TimeoutError):
            pass                                     # client went away
        except Exception as e:                       # never let a handler crash the service
            _log("HANDLER ERROR:\n" + traceback.format_exc())
            try:
                self._json(500, {"error": f"Internal error: {type(e).__name__}"})
            except Exception:
                pass

    # ------------------------------------------------------------------ routes
    def do_GET(self):
        self._safe(self._get)

    def do_PUT(self):
        self._safe(self._put)

    def do_POST(self):
        self._safe(self._post)

    def _get(self):
        if not self._guard(False):
            return
        path = urlparse(self.path).path
        if path in ("/", "/index.html"):
            with open(HTML, "rb") as f:
                self._send(200, f.read(), "text/html; charset=utf-8")
        elif path == "/api/ping":
            self._json(200, {"app": APP_ID, "data": os.path.abspath(DATA_DIR)})
        elif path == "/api/data":
            if _cache["body"] is None:
                return self._json(404, {"error": "No data has been loaded yet."})
            if self.headers.get("If-None-Match") == _cache["etag"]:
                return self._send(304, b"", "application/json", {"ETag": _cache["etag"]})
            self._send(200, _cache["body"], "application/json; charset=utf-8", {"ETag": _cache["etag"]})
        elif path == "/api/meta":
            self._json(200, {"updated_at": _cache["updated"]})
        elif path == "/api/status":
            inbox = _inbox_files()
            self._json(200, {k: _job[k] for k in ("state", "log", "error", "started", "finished")}
                       | {"hasData": _cache["body"] is not None, "excelAvailable": X.excel_installed(),
                          "inbox": [os.path.basename(f) for f in inbox], "inboxReady": len(inbox) == 3})
        else:
            self._json(404, {"error": "not found"})

    def _put(self):
        if not self._guard(True):
            return
        m = re.fullmatch(r"/api/upload/([123])", urlparse(self.path).path)
        if not m:
            return self._json(404, {"error": "not found"})
        if _job["state"] == "running":
            return self._json(409, {"error": "A refresh is already running."})
        slot = int(m.group(1))
        name = os.path.basename(unquote(self.headers.get("X-Filename", "upload.xlsx")).replace("\\", "/"))
        name = re.sub(r"[^\w.\- ()]", "_", name).strip(" .") or "upload.xlsx"
        stem, ext = os.path.splitext(name)
        name = stem[:80] + ext[:8]                          # keep Windows paths short
        try:
            n = int(self.headers.get("Content-Length", ""))
        except ValueError:
            return self._json(411, {"error": "Content-Length required."})
        if n <= 0:
            return self._json(400, {"error": "Empty upload."})
        if n > MAX_UPLOAD:
            return self._json(413, {"error": "File too large."})
        os.makedirs(UPLOAD_DIR, exist_ok=True)
        free = shutil.disk_usage(UPLOAD_DIR).free
        if free < n * 2 + (200 << 20):
            return self._json(507, {"error": "Not enough free disk space for this upload."})
        for f in os.listdir(UPLOAD_DIR):                     # replace whatever this slot held before
            if f.startswith(f"slot{slot}__"):
                os.remove(os.path.join(UPLOAD_DIR, f))
        dest = os.path.join(UPLOAD_DIR, f"slot{slot}__{name}")
        left = n
        with open(dest, "wb") as f:
            while left:
                chunk = self.rfile.read(min(1 << 20, left))
                if not chunk:
                    break
                f.write(chunk)
                left -= len(chunk)
        if left:
            os.remove(dest)
            return self._json(400, {"error": "Upload interrupted."})
        self._json(200, {"slot": slot, "name": name, "size": n})

    def _post(self):
        if not self._guard(True):
            return
        u = urlparse(self.path)
        if u.path != "/api/process":
            return self._json(404, {"error": "not found"})
        if not X.excel_installed():
            return self._json(400, {"error": "Microsoft Excel (desktop version) is required on this computer but was not found."})
        from_inbox = parse_qs(u.query).get("source") == ["inbox"]
        with _lock:
            if _job["state"] == "running":
                return self._json(409, {"error": "A refresh is already running."})
            if from_inbox:
                staged, run_dir = _inbox_files(), None          # read in place; the user's files are never moved
                if len(staged) != 3:
                    return self._json(400, {"error": f"The inbox folder must contain exactly 3 Excel files (found {len(staged)})."})
            else:
                paths = _staged_uploads()
                if len(paths) != 3:
                    return self._json(400, {"error": "Attach all 3 source files first."})
                # Move the files into a private run folder so later uploads (e.g. from a second browser tab)
                # can never change the files a running job is reading.
                run_dir = os.path.join(RUN_DIR, time.strftime("run_%Y%m%d_%H%M%S"))
                os.makedirs(run_dir, exist_ok=True)
                staged = []
                for p in paths:
                    dest = os.path.join(run_dir, os.path.basename(p))
                    shutil.move(p, dest)
                    staged.append(dest)
            _job.update(state="running", log=[], error=None, started=time.time(), finished=None)
        threading.Thread(target=_run_job, args=(staged, run_dir), daemon=True).start()
        self._json(202, {"state": "running"})


def _ping(port: int) -> bool:
    """Is a dashboard service for THIS data folder already listening on that port?
    Uses a direct socket connection: urllib would route 127.0.0.1 through a corporate proxy and time out."""
    try:
        c = http.client.HTTPConnection(HOST, port, timeout=1.0)
        c.request("GET", "/api/ping", headers={"Host": f"{HOST}:{port}"})
        info = json.loads(c.getresponse().read())
        c.close()
        return info.get("app") == APP_ID and os.path.normcase(info.get("data", "")) == os.path.normcase(os.path.abspath(DATA_DIR))
    except Exception:
        return False


def main() -> int:
    global PORT
    try:
        sys.stdout.reconfigure(errors="replace")
    except Exception:
        pass
    open_browser = "--no-browser" not in sys.argv
    os.makedirs(DATA_DIR, exist_ok=True)
    # Already running? just open it. Only the port used last time (and the default) are probed: on machines with
    # endpoint-security software a connect to a closed loopback port can hang for a second, so never scan a range.
    candidates = []
    try:
        candidates.append(int(open(os.path.join(DATA_DIR, "port.txt")).read().strip()))
    except (OSError, ValueError):
        pass
    if BASE_PORT not in candidates:
        candidates.append(BASE_PORT)
    for p in candidates:
        if _ping(p):
            print(f"The dashboard is already running at http://{HOST}:{p}/")
            if open_browser:
                webbrowser.open(f"http://{HOST}:{p}/")
            return 0
    srv = None
    for p in range(BASE_PORT, BASE_PORT + PORT_TRIES):
        try:
            srv = ThreadingHTTPServer((HOST, p), Handler)
            PORT = p
            break
        except OSError:
            continue
    if srv is None:
        print(f"ERROR: no free port between {BASE_PORT} and {BASE_PORT + PORT_TRIES - 1}.")
        return 1
    srv.daemon_threads = True
    # housekeeping after a crash / forced close
    orphan = X.kill_orphan(PID_PATH)
    if orphan:
        _log(f"Closed a stray Excel process (pid {orphan}) left by an interrupted run.")
    store.cleanup_stale_temp(DB_PATH)
    os.makedirs(INBOX_DIR, exist_ok=True)
    shutil.rmtree(UPLOAD_DIR, ignore_errors=True)
    shutil.rmtree(RUN_DIR, ignore_errors=True)
    _load_cache()
    with open(os.path.join(DATA_DIR, "port.txt"), "w") as f:
        f.write(str(PORT))
    url = f"http://{HOST}:{PORT}/"
    print(f"Dashboard is running at {url}")
    print("Keep this window open while you use the dashboard. Close it (or press Ctrl+C) to stop.")
    if open_browser:
        threading.Timer(0.5, lambda: webbrowser.open(url)).start()
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass
    return 0


if __name__ == "__main__":
    sys.exit(main())
