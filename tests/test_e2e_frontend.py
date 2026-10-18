import os
import sys
import time
import subprocess
import unittest
from playwright.sync_api import sync_playwright

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ARTIFACTS_DIR = r"C:\Users\ntc27\.gemini\antigravity\brain\9166a9f1-2c38-41c5-967b-79b51c7e5cf6"
PORT = 8092
SERVER_URL = f"http://127.0.0.1:{PORT}"

class TestE2EFrontend(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        # Clean previous state
        for fname in ["license.json", "trials_db.json"]:
            fpath = os.path.join(BASE_DIR, fname)
            if os.path.exists(fpath):
                try:
                    os.remove(fpath)
                except Exception:
                    pass

        # Start LUV-DOC-TRAN server on port 8092
        env = os.environ.copy()
        env["PORT"] = str(PORT)
        cls.server_proc = subprocess.Popen(
            [sys.executable, "-m", "uvicorn", "server:app", "--host", "127.0.0.1", "--port", str(PORT)],
            cwd=BASE_DIR,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            env=env
        )
        time.sleep(3)

    @classmethod
    def tearDownClass(cls):
        if cls.server_proc:
            cls.server_proc.terminate()
            cls.server_proc.wait(timeout=5)

    def test_complete_frontend_flow(self):
        """Test full frontend UX: TitanPDF minimalist branding, multi-file queue, per-file instant download, Split-view, and PDF modal."""
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            context = browser.new_context(viewport={"width": 1440, "height": 950})
            page = context.new_page()

            # 1. Visit homepage
            page.goto(SERVER_URL, timeout=15000)
            page.wait_for_selector(".brand-name")

            brand_text = page.inner_text(".brand-name").replace("\n", "")
            self.assertIn("LUV-DOC-TRAN", brand_text)

            # 2. Check 48-Hour Free Trial Pill
            trial_text = page.inner_text("#trial-timer-text")
            self.assertTrue("Dùng thử" in trial_text or "Bản Quyền" in trial_text)

            # 3. Check Tech Intro Bar
            tech_bar = page.inner_text(".tech-bar")
            self.assertIn("#002060", tech_bar)
            self.assertIn("Bảng biểu", tech_bar)

            # Take screenshot of Initial State
            ss_path1 = os.path.join(ARTIFACTS_DIR, "luv_doctran_e2e_home.png")
            page.screenshot(path=ss_path1)
            print(f"[E2E] Saved homepage screenshot: {ss_path1}")

            # 4. Open License Activation Modal
            page.click("#btn-open-license")
            page.wait_for_selector("#modal-license", state="visible")

            # Fill in License Key
            page.fill("#act-reg-name", "Doanh Nghiệp Kỹ Thuật Hóa Chất Alpha")
            page.fill("#act-key", "LUV-PDF-2026-PRO-VN99")

            # Click Activate
            page.click("#btn-submit-activate")
            time.sleep(2)

            # Verify badge updated
            header_badge = page.inner_text("#trial-pill")
            self.assertIn("Bản Quyền", header_badge)

            # 5. Create 2 sample files for Multi-file Upload testing
            sample_docx1 = os.path.join(BASE_DIR, "data", "temp", "e2e_sample_msds.docx")
            sample_docx2 = os.path.join(BASE_DIR, "data", "temp", "e2e_sample_tds.docx")
            import docx
            for path, title, chemical in [
                (sample_docx1, "Material Safety Data Sheet (MSDS)", "Acetone (CAS: 67-64-1)"),
                (sample_docx2, "Technical Data Sheet (TDS)", "Polyethylene Glycol (CAS: 25322-68-3)")
            ]:
                d = docx.Document()
                d.add_heading(title, level=1)
                d.add_paragraph(f"Active Chemical Component: {chemical}")
                d.save(path)

            # Multi-file Upload via file input
            page.set_input_files("#file-input", [sample_docx1, sample_docx2])
            time.sleep(0.5)

            # Check Queue Table is visible
            page.wait_for_selector("#queue-container", state="visible")
            badge_count = page.inner_text("#queue-count-badge")
            self.assertEqual(badge_count, "2 tệp")

            # Click Batch Processing Button
            page.click("#btn-start-batch")

            # Wait for at least first row to complete
            page.wait_for_selector(".btn-row-word", state="visible", timeout=40000)

            # Take screenshot of Queue with completed files
            ss_path_queue = os.path.join(ARTIFACTS_DIR, "luv_doctran_e2e_queue.png")
            page.screenshot(path=ss_path_queue)
            print(f"[E2E] Saved Queue screenshot: {ss_path_queue}")

            # 6. Test Split-View Studio
            page.wait_for_selector("#split-view-studio", state="visible", timeout=30000)

            # Check HUD doc name
            hud_name = page.inner_text("#hud-doc-name")
            self.assertTrue(len(hud_name) > 0)

            # Verify Word Priority button is prominent in Split-View
            word_btn_text = page.inner_text("#btn-studio-dl-word")
            self.assertIn("TẢI FILE WORD (.DOCX)", word_btn_text)
            self.assertIn("Khuyên dùng", word_btn_text)

            # 7. Test PDF Download Button & Reminder Modal
            page.click("#btn-studio-dl-pdf")
            page.wait_for_selector("#modal-pdf-progress", state="visible")

            # Verify 2-3 minute notice in modal
            modal_notice = page.inner_text(".modal-alert-box")
            self.assertIn("2 - 3 phút", modal_notice)
            self.assertIn("Khuyên dùng", modal_notice)

            # Take screenshot of PDF Loading Modal
            ss_path2 = os.path.join(ARTIFACTS_DIR, "luv_doctran_e2e_pdf_modal.png")
            page.screenshot(path=ss_path2)
            print(f"[E2E] Saved PDF progress modal screenshot: {ss_path2}")

            # Close modal via close button
            page.click(".btn-modal-cancel")
            time.sleep(0.5)

            # Take final full page screenshot of Split-View Studio
            ss_path3 = os.path.join(ARTIFACTS_DIR, "luv_doctran_e2e_split_view.png")
            page.screenshot(path=ss_path3)
            print(f"[E2E] Saved Split-View Studio screenshot: {ss_path3}")

            browser.close()


if __name__ == "__main__":
    unittest.main()
