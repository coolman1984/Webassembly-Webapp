"""The dashboard as served by server.py, in a real browser, without Excel (developer tool: needs Playwright).

    python tests/test_served_page.py            # PW_CHROMIUM=<path> to use a pre-installed Chromium

Covers what only the served page does: loading /api/data (gzip), the refresh details / notes panel, and loading a
database far bigger than a function call can take as spread arguments (a >120k-row regression).
"""
from __future__ import annotations

import os
import shutil
import sys
import tempfile
import threading
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
from tests import test_portable as TP            # noqa: E402  (installs the COM stubs when pywin32 is absent)
import server                                     # noqa: E402
from pipeline import store                        # noqa: E402
from playwright.sync_api import sync_playwright   # noqa: E402

BIG = 150_000


class TestServedPage(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.dir = tempfile.mkdtemp(prefix="bomserved_")
        cls.patch = TP.mock.patch.multiple(server, DB_PATH=os.path.join(cls.dir, "dashboard.db"), DATA_DIR=cls.dir,
                                           LOG_PATH=os.path.join(cls.dir, "pipeline.log"), INBOX_DIR=os.path.join(cls.dir, "inbox"))
        cls.patch.start()
        cls.no_excel = TP.mock.patch.object(server.X, "excel_installed", return_value=True)
        cls.no_excel.start()
        cls.srv = server.ThreadingHTTPServer(("127.0.0.1", 0), server.Handler)
        server.PORT = cls.srv.server_address[1]
        cls.url = f"http://127.0.0.1:{server.PORT}/"
        threading.Thread(target=cls.srv.serve_forever, daemon=True).start()
        cls.pw = sync_playwright().start()
        cls.browser = cls.pw.chromium.launch(executable_path=os.environ.get("PW_CHROMIUM") or None)

    @classmethod
    def tearDownClass(cls):
        cls.browser.close()
        cls.pw.stop()
        cls.srv.shutdown()
        cls.srv.server_close()
        cls.no_excel.stop()
        cls.patch.stop()
        shutil.rmtree(cls.dir, ignore_errors=True)

    def load(self, ext, bom, master, report):
        store.write_db(server.DB_PATH, ext, bom, master, report)
        server._load_cache()
        page = self.browser.new_page()
        errs, enc = [], []
        page.on("pageerror", lambda e: errs.append(str(e)))
        page.on("response", lambda r: enc.append(r.headers.get("content-encoding")) if r.url.endswith("/api/data") else None)
        page.goto(self.url)
        page.wait_for_function("document.getElementById('updateStatus').textContent.includes('local database')", timeout=60000)
        self.addCleanup(page.close)
        return page, errs, enc

    def test_refresh_details_and_notes_are_shown(self):
        ext, bom, master, report = TP._datasets()
        page, errs, enc = self.load(ext, bom, master, report)
        self.assertEqual(enc, ["gzip"])
        self.assertIn("1 note from the last refresh", page.inner_text("#dataStrip"))
        page.click("#notesLink")
        info = page.inner_text("#dataInfo")
        self.assertIn("d.xlsx", info)
        self.assertIn("note one", info)
        self.assertFalse(page.is_visible("#emptyState"))
        self.assertEqual(errs, [])

    def test_very_large_database_loads(self):
        ext, bom, master, report = TP._datasets()
        row = dict(master[0])
        big = [dict(row, model=f"M{i:07d}") for i in range(BIG)]
        page, errs, _ = self.load(ext, bom, big, report)
        self.assertEqual(errs, [])
        self.assertEqual(page.evaluate("MASTER_DATA.length"), BIG)
        page.click("button[data-tab='master']")
        self.assertIn(f"of {BIG} models", page.inner_text("#masterInfo"))


if __name__ == "__main__":
    unittest.main()
