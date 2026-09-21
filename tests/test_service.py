"""Service-level tests: HTTP security, hostile input, concurrency, crash recovery, ports, timeouts, leaks.
The service is started exactly as an end user gets it: the bundled runtime\\python.exe in isolated mode.

    runtime\\python.exe -I tests\\test_service.py -v
"""
from __future__ import annotations

import http.client
import json
import os
import shutil
import socket
import subprocess
import sys
import tempfile
import threading
import time
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
from pipeline import excel_com as X          # noqa: E402
from tests import fixtures as F              # noqa: E402

PY = os.path.join(ROOT, "runtime", "python.exe")
_xl = None
_tmp = None
_n = 0


def setUpModule():
    global _xl, _tmp
    _xl = F.Excel().__enter__()
    _tmp = tempfile.mkdtemp(prefix="bomsvc_")


def tearDownModule():
    _xl.__exit__(None, None, None)
    shutil.rmtree(_tmp, ignore_errors=True)


def make(**o) -> dict:
    global _n
    _n += 1
    return F.build(os.path.join(_tmp, f"fx{_n}"), _xl, **o)


def excel_pids() -> set[int]:
    out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/FO", "CSV", "/NH"], capture_output=True, text=True).stdout
    return {int(l.split(",")[1].strip('"')) for l in out.splitlines() if l.startswith('"EXCEL.EXE"')}


def rss_mb(pid: int) -> float:
    out = subprocess.run(["tasklist", "/FI", f"PID eq {pid}", "/FO", "CSV", "/NH"], capture_output=True, text=True).stdout
    return float(out.strip().split(",")[-1].strip('"').replace(" K", "").replace(",", "").replace(".", "")) / 1024


class Srv:
    def __init__(self, env=None):
        self.dir = tempfile.mkdtemp(prefix="data_", dir=_tmp)
        self.env = dict(os.environ, BOM_DATA_DIR=self.dir, **(env or {}))
        self.proc = None
        self.port = None

    def start(self, timeout=20):
        self.proc = subprocess.Popen([PY, "-I", "-X", "utf8", os.path.join(ROOT, "server.py"), "--no-browser"],
                                     cwd=ROOT, env=self.env, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
        t0 = time.time()
        pf = os.path.join(self.dir, "port.txt")
        while time.time() - t0 < timeout:
            if self.proc.poll() is not None:
                raise RuntimeError("server exited: " + self.proc.stdout.read().decode(errors="replace"))
            if os.path.isfile(pf) and open(pf).read().strip():
                self.port = int(open(pf).read())
                try:
                    if self.req("GET", "/api/ping")[0] == 200:
                        return self
                except OSError:
                    pass
            time.sleep(0.1)
        raise RuntimeError("server did not start")

    def stop(self):
        if self.proc and self.proc.poll() is None:
            self.proc.terminate()
            try:
                self.proc.wait(10)
            except subprocess.TimeoutExpired:
                self.proc.kill()

    def req(self, method, path, body=None, headers=None, host=None, write=False, raw=False):
        h = {"X-Requested-With": "bom-dashboard"} if write else {}
        h.update(headers or {})
        c = http.client.HTTPConnection("127.0.0.1", self.port, timeout=60)
        c.putrequest(method, path, skip_host=True)
        c.putheader("Host", host or f"127.0.0.1:{self.port}")
        for k, v in h.items():
            c.putheader(k, v)
        if body is not None and "Content-Length" not in h:
            c.putheader("Content-Length", str(len(body)))
        c.endheaders(body if body is not None else None)
        r = c.getresponse()
        data = r.read()
        c.close()
        if raw:
            return r.status, dict(r.getheaders()), data
        try:
            return r.status, json.loads(data) if data else None
        except ValueError:
            return r.status, data

    def upload(self, slot, path, name=None):
        with open(path, "rb") as f:
            body = f.read()
        import urllib.parse
        return self.req("PUT", f"/api/upload/{slot}", body, {"X-Filename": urllib.parse.quote(name or os.path.basename(path))}, write=True)

    def refresh(self, files, order=("nm", "sop", "dash"), names=None, wait=True, timeout=240):
        for i, k in enumerate(order, start=1):
            st, r = self.upload(i, files[k], (names or {}).get(k))
            assert st == 200, (st, r)
        st, r = self.req("POST", "/api/process", write=True)
        if not wait or st != 202:
            return st, r
        t0 = time.time()
        while time.time() - t0 < timeout:
            st, s = self.req("GET", "/api/status")
            if s["state"] != "running":
                return st, s
            time.sleep(0.5)
        raise TimeoutError("job did not finish")


class SvcCase(unittest.TestCase):
    def setUp(self):
        self.srv = Srv().start()
        self.addCleanup(self.srv.stop)


class TestSecurity(SvcCase):
    def test_dns_rebinding_host_header_rejected_everywhere(self):
        for path in ("/", "/api/data", "/api/status", "/api/ping"):
            st, _ = self.srv.req("GET", path, host="evil.example.com")
            self.assertEqual(st, 403, path)
        st, _ = self.srv.req("GET", "/api/ping", host=f"localhost:{self.srv.port}")
        self.assertEqual(st, 200)

    def test_state_changing_calls_need_csrf_header_and_same_origin(self):
        st, _ = self.srv.req("POST", "/api/process")                                   # no custom header
        self.assertEqual(st, 403)
        st, _ = self.srv.req("POST", "/api/process", headers={"X-Requested-With": "bom-dashboard", "Origin": "https://evil.example"})
        self.assertEqual(st, 403)
        st, _ = self.srv.req("PUT", "/api/upload/1", b"x", headers={"Origin": "https://evil.example", "X-Requested-With": "bom-dashboard"})
        self.assertEqual(st, 403)
        st, _ = self.srv.req("PUT", "/api/upload/1", b"x")                             # simple cross-site form style request
        self.assertEqual(st, 403)

    def test_no_cors_headers_leak_data_cross_origin(self):
        st, headers, _ = self.srv.req("GET", "/api/status", raw=True, headers={"Origin": "https://evil.example"})
        self.assertNotIn("access-control-allow-origin", {k.lower() for k in headers})

    def test_unknown_routes_and_methods(self):
        self.assertEqual(self.srv.req("GET", "/../../etc/passwd")[0], 404)
        self.assertEqual(self.srv.req("GET", "/api/upload/9")[0], 404)
        self.assertEqual(self.srv.req("PUT", "/api/upload/9", b"x", write=True)[0], 404)
        self.assertEqual(self.srv.req("DELETE", "/api/data", write=True)[0], 501)


class TestHostileUploads(SvcCase):
    def test_bad_content_length_variants(self):
        self.assertEqual(self.srv.req("PUT", "/api/upload/1", None, {"Content-Length": "0"}, write=True)[0], 400)
        self.assertEqual(self.srv.req("PUT", "/api/upload/1", None, {"Content-Length": "abc"}, write=True)[0], 411)
        self.assertEqual(self.srv.req("PUT", "/api/upload/1", None, {"Content-Length": str(5 << 30)}, write=True)[0], 413)
        self.assertEqual(self.srv.req("PUT", "/api/upload/1", None, {"Content-Length": "-5"}, write=True)[0], 400)

    def test_truncated_upload_leaves_nothing_behind(self):
        s = socket.create_connection(("127.0.0.1", self.srv.port), timeout=10)
        s.sendall((f"PUT /api/upload/2 HTTP/1.1\r\nHost: 127.0.0.1:{self.srv.port}\r\nX-Requested-With: bom-dashboard\r\n"
                   "X-Filename: half.xlsx\r\nContent-Length: 100000\r\n\r\n").encode() + b"x" * 500)
        s.close()                                                # client vanishes mid-upload
        time.sleep(1.5)
        up = os.path.join(self.srv.dir, "uploads")
        self.assertEqual([f for f in os.listdir(up) if f.startswith("slot2")] if os.path.isdir(up) else [], [])
        self.assertEqual(self.srv.req("GET", "/api/ping")[0], 200, "service must survive a dropped client")

    def test_path_traversal_and_odd_filenames_are_neutralised(self):
        up = os.path.join(self.srv.dir, "uploads")
        for name in ("..\\..\\evil.xlsx", "../../evil.xlsx", "C:\\Windows\\evil.xlsx", "a" * 400 + ".xlsx", "\u0645\u0644\u0641 \u0639\u0631\u0628\u064a.xlsx",
                     "na<>me|?*.xlsx", "..."):
            st, r = self.srv.upload(1, __file__, name)
            self.assertEqual(st, 200, name)
            files = os.listdir(up)
            self.assertEqual(len(files), 1)
            self.assertTrue(files[0].startswith("slot1__"), files)
            self.assertLess(len(files[0]), 110)
            self.assertFalse(os.path.exists(os.path.join(self.srv.dir, "evil.xlsx")))
            self.assertFalse(os.path.exists(os.path.join(ROOT, "evil.xlsx")))

    def test_process_without_three_files_is_a_clean_400(self):
        self.srv.upload(1, __file__, "one.xlsx")
        st, r = self.srv.req("POST", "/api/process", write=True)
        self.assertEqual(st, 400)
        self.assertIn("all 3", r["error"])


class TestRefreshFlow(SvcCase):
    def test_full_refresh_new_model_and_failed_refresh_keeps_old_data(self):
        srv = self.srv
        self.assertEqual(srv.req("GET", "/api/data")[0], 404)
        base = make()
        st, s = srv.refresh(base, names={"dash": "\u0645\u0644\u0641 DASH \u0639\u0631\u0628\u064a.xlsx"})     # Arabic file name
        self.assertEqual(s["state"], "done", s)
        st, d = srv.req("GET", "/api/data")
        self.assertEqual([r["model"] for r in d["bom"]], F.BOM_ORDER)
        for r in d["bom"]:
            for k, v in F.EXPECTED_BOM[r["model"]].items():
                self.assertEqual(r[k], v, f"{r['model']}.{k}")
        upd1 = d["meta"]["updated_at"]
        # ETag / conditional GET
        st, h, body = srv.req("GET", "/api/data", raw=True)
        etag = {k.lower(): v for k, v in h.items()}["etag"]
        self.assertEqual(srv.req("GET", "/api/data", headers={"If-None-Match": etag}, raw=True)[0], 304)
        # uploads are cleaned up after processing
        time.sleep(0.5)
        self.assertFalse(os.path.isdir(os.path.join(srv.dir, "runs")) and os.listdir(os.path.join(srv.dir, "runs")))

        # ---- a NEW MODEL appears in all three files
        time.sleep(1.1)
        st, s = srv.refresh(make(new_model=True))
        self.assertEqual(s["state"], "done", s)
        st, d = srv.req("GET", "/api/data")
        self.assertEqual(len(d["bom"]), len(F.BOM_ORDER) + 1)
        new = next(r for r in d["bom"] if r["model"] == "TV-N9")
        for k, v in F.EXPECTED_NEW.items():
            self.assertEqual(new[k], v, f"TV-N9.{k}")
        self.assertIn("TV-N9", {r["model"] for r in d["master"]})
        upd2 = d["meta"]["updated_at"]
        self.assertNotEqual(upd1, upd2)
        self.assertEqual(srv.req("GET", "/api/meta")[1]["updated_at"], upd2)

        # ---- a broken refresh must leave the current data untouched
        bad = make(rename={"Cate": "Category"})
        st, s = srv.refresh(bad)
        self.assertEqual(s["state"], "error")
        self.assertIn("Cate", s["error"])
        st, d2 = srv.req("GET", "/api/data")
        self.assertEqual(d2["meta"]["updated_at"], upd2)
        self.assertEqual(len(d2["bom"]), len(F.BOM_ORDER) + 1)

    def test_concurrent_processing_requests_run_exactly_one_job(self):
        srv = self.srv
        files = make()
        for i, k in enumerate(("nm", "sop", "dash"), start=1):
            srv.upload(i, files[k])
        codes = []
        lock = threading.Lock()
        def go():
            st, _ = srv.req("POST", "/api/process", write=True)
            with lock:
                codes.append(st)
        ts = [threading.Thread(target=go) for _ in range(6)]
        [t.start() for t in ts]; [t.join() for t in ts]
        self.assertEqual(codes.count(202), 1, codes)                 # exactly one job was started
        self.assertTrue(all(c in (202, 409, 400) for c in codes), codes)
        # uploads while the job runs are refused, and the running job is unaffected
        st, r = srv.upload(1, files["nm"])
        self.assertIn(st, (409, 200))
        while srv.req("GET", "/api/status")[1]["state"] == "running":
            time.sleep(0.5)
        self.assertEqual(srv.req("GET", "/api/status")[1]["state"], "done")

    def test_repeated_refreshes_do_not_leak_excel_or_memory(self):
        srv = self.srv
        files = make()
        # NB: other programs on a PC may start/stop their own Excel at any time, so a global before/after diff of
        # Excel processes proves nothing. What matters: our own Excel is recorded, closed gracefully, and forgotten.
        sizes = []
        for i in range(5):
            st, s = srv.refresh(files)
            self.assertEqual(s["state"], "done", s)
            self.assertFalse(os.path.exists(os.path.join(srv.dir, "excel.pid")), f"own Excel still registered after run {i + 1}")
            self.assertFalse(any("stopped" in l for l in s["log"]), f"Excel needed a forced stop in run {i + 1}: {s['log']}")
            sizes.append(rss_mb(srv.proc.pid))
        self.assertLess(sizes[-1] - sizes[0], 60, f"service memory grew: {sizes}")


class TestExcelProblems(unittest.TestCase):
    def test_no_excel_installed_gives_a_clear_message(self):
        srv = Srv({"BOM_FORCE_NO_EXCEL": "1"}).start()
        self.addCleanup(srv.stop)
        self.assertFalse(srv.req("GET", "/api/status")[1]["excelAvailable"])
        st, r = srv.req("POST", "/api/process", write=True)
        self.assertEqual(st, 400)
        self.assertIn("Microsoft Excel", r["error"])

    def test_job_timeout_stops_excel_and_service_survives(self):
        srv = Srv({"BOM_JOB_TIMEOUT": "1"}).start()
        self.addCleanup(srv.stop)
        st, s = srv.refresh(make())
        self.assertEqual(s["state"], "error", s)
        self.assertIn("did not finish", s["error"])
        self.assertFalse(os.path.exists(os.path.join(srv.dir, "excel.pid")), "killed Excel must be deregistered")
        self.assertEqual(srv.req("GET", "/api/ping")[0], 200)


class TestCrashRecovery(unittest.TestCase):
    def test_corrupt_database_at_startup_is_ignored_and_recovered(self):
        d = tempfile.mkdtemp(dir=_tmp)
        open(os.path.join(d, "dashboard.db"), "wb").write(b"this is not sqlite" * 50)
        open(os.path.join(d, "tmpabc123.db"), "wb").write(b"half written")           # stale temp from a crash
        srv = Srv(); srv.dir = d; srv.env["BOM_DATA_DIR"] = d
        srv.start(); self.addCleanup(srv.stop)
        self.assertEqual(srv.req("GET", "/api/data")[0], 404)
        self.assertFalse(os.path.exists(os.path.join(d, "tmpabc123.db")), "stale temp db should be removed")
        st, s = srv.refresh(make())
        self.assertEqual(s["state"], "done", s)
        self.assertEqual(srv.req("GET", "/api/data")[0], 200)

    def test_orphan_excel_from_crashed_run_is_closed_but_unrelated_process_is_not(self):
        other = F.Excel().__enter__()                                   # stands in for an orphan of a killed run
        import win32process
        pid = win32process.GetWindowThreadProcessId(other.app.Hwnd)[1]
        d = tempfile.mkdtemp(dir=_tmp)
        json.dump({"pid": pid, "started": X._proc_start_time(pid)}, open(os.path.join(d, "excel.pid"), "w"))
        # decoy: a live, unrelated process referenced with a WRONG start time - must never be killed
        decoy = subprocess.Popen([PY, "-I", "-c", "import time; time.sleep(120)"])
        d2 = tempfile.mkdtemp(dir=_tmp)
        json.dump({"pid": decoy.pid, "started": 12345}, open(os.path.join(d2, "excel.pid"), "w"))
        try:
            s1 = Srv(); s1.dir = d; s1.env["BOM_DATA_DIR"] = d; s1.start(); self.addCleanup(s1.stop)
            time.sleep(1)
            self.assertNotIn(pid, excel_pids(), "orphan Excel should have been closed at startup")
            s2 = Srv(); s2.dir = d2; s2.env["BOM_DATA_DIR"] = d2; s2.start(); self.addCleanup(s2.stop)
            time.sleep(1)
            self.assertIsNone(decoy.poll(), "an unrelated process must never be terminated")
        finally:
            decoy.kill()


class TestInbox(unittest.TestCase):
    def test_inbox_files_are_read_in_place_and_never_touched(self):
        import hashlib
        inbox = tempfile.mkdtemp(dir=_tmp)
        fx = make()
        for k in ("nm", "sop", "dash"):
            shutil.copy(fx[k], inbox)
        digest = lambda: sorted((f, hashlib.sha256(open(os.path.join(inbox, f), "rb").read()).hexdigest()) for f in os.listdir(inbox))
        before = digest()
        srv = Srv({"BOM_INBOX_DIR": inbox}).start(); self.addCleanup(srv.stop)
        s = srv.req("GET", "/api/status")[1]
        self.assertTrue(s["inboxReady"], s["inbox"])
        st, r = srv.req("POST", "/api/process?source=inbox", write=True)
        self.assertEqual(st, 202, r)
        while srv.req("GET", "/api/status")[1]["state"] == "running":
            time.sleep(0.5)
        s = srv.req("GET", "/api/status")[1]
        self.assertEqual(s["state"], "done", s)
        d = srv.req("GET", "/api/data")[1]
        self.assertEqual([r["model"] for r in d["bom"]], F.BOM_ORDER)
        self.assertEqual(digest(), before, "inbox files must be left exactly as they were")

    def test_inbox_needs_exactly_three_workbooks_and_ignores_lock_and_other_files(self):
        inbox = tempfile.mkdtemp(dir=_tmp)
        fx = make()
        shutil.copy(fx["nm"], inbox); shutil.copy(fx["dash"], inbox)
        open(os.path.join(inbox, "~$lock.xlsx"), "wb").write(b"x")       # Excel lock file: not a workbook
        open(os.path.join(inbox, "notes.txt"), "w").write("hello")       # not a workbook
        srv = Srv({"BOM_INBOX_DIR": inbox}).start(); self.addCleanup(srv.stop)
        s = srv.req("GET", "/api/status")[1]
        self.assertFalse(s["inboxReady"])
        self.assertEqual(len(s["inbox"]), 2)
        st, r = srv.req("POST", "/api/process?source=inbox", write=True)
        self.assertEqual(st, 400)
        self.assertIn("exactly 3", r["error"])
        shutil.copy(fx["sop"], inbox)
        self.assertTrue(srv.req("GET", "/api/status")[1]["inboxReady"])


class TestPorts(unittest.TestCase):
    def test_second_start_reuses_the_running_instance(self):
        a = Srv().start(); self.addCleanup(a.stop)
        p = subprocess.Popen([PY, "-I", "-X", "utf8", os.path.join(ROOT, "server.py"), "--no-browser"], cwd=ROOT, env=a.env,
                             stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
        out = p.communicate(timeout=30)[0].decode()
        self.assertEqual(p.returncode, 0)
        self.assertIn("already running", out)

    def test_falls_back_to_another_port_when_busy(self):
        blocker = socket.socket(); blocker.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 0)
        try:
            blocker.bind(("127.0.0.1", 8765)); blocker.listen(5)
        except OSError:
            self.skipTest("port 8765 already in use by something else")
        self.addCleanup(blocker.close)
        s = Srv().start(); self.addCleanup(s.stop)
        self.assertNotEqual(s.port, 8765)
        self.assertEqual(s.req("GET", "/api/ping")[0], 200)


if __name__ == "__main__":
    unittest.main(verbosity=2)
