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
        # Wait for server to boot
        time.sleep(3)

    @classmethod
    def tearDownClass(cls):
        if cls.server_proc:
            cls.server_proc.terminate()
            cls.server_proc.wait(timeout=5)

    def test_complete_frontend_flow(self):
        """Test full frontend UX: Kawaii branding, 48h trial, license modal, translation, Split-view, Word priority, and PDF modal."""
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            context = browser.new_context(viewport={"width": 1400, "height": 900})
            page = context.new_page()

            # 1. Visit homepage
            page.goto(SERVER_URL, timeout=15000)
            page.wait_for_selector(".brand-title")

            title_text = page.inner_text(".brand-title").replace("\n", "")
            self.assertIn("LUV-DOC-TRAN", title_text)

            # 2. Check 48-Hour Free Trial Pill
            trial_text = page.inner_text("#trial-timer-text")
            self.assertTrue("Dùng thử" in trial_text or "Bản Quyền" in trial_text)

            # 3. Check Mascot Greeting
            speech_title = page.inner_text(".speech-title")
            self.assertIn("Docy", speech_title)

            # Take screenshot of Initial State
            ss_path1 = os.path.join(ARTIFACTS_DIR, "luv_doctran_e2e_home.png")
            page.screenshot(path=ss_path1)
            print(f"[E2E] Saved homepage screenshot: {ss_path1}")

            # 4. Open License Activation Modal
            page.click("#btn-open-license")
            page.wait_for_selector("#modal-license", state="visible")

            # Fill in License Key
            page.fill("#act-reg-name", "Doanh Nghiệp Hóa Chất Alpha")
            page.fill("#act-key", "LUV-PDF-2026-PRO-VN99")

            # Click Activate
            page.click("#btn-submit-activate")
            time.sleep(2)

            # Verify badge updated
            header_badge = page.inner_text("#trial-pill")
            self.assertIn("Bản Quyền", header_badge)

            # 5. Create a sample Word file to test translation
            sample_docx = os.path.join(BASE_DIR, "data", "temp", "e2e_sample.docx")
            import docx
            d = docx.Document()
            d.add_heading("Technical Specification Sheet", level=1)
            d.add_paragraph("Product: Polyethylene Glycol (CAS: 25322-68-3).")
            d.save(sample_docx)

            # Upload file via file input
            page.set_input_files("#file-input", sample_docx)
            time.sleep(0.5)

            selected_name = page.inner_text("#selected-file-name")
            self.assertEqual(selected_name, "e2e_sample.docx")

            # Click Translate
            page.click("#btn-start-translate")

            # Wait for Split-View Studio to appear
            page.wait_for_selector("#split-view-studio", state="visible", timeout=30000)

            # Check HUD doc name
            hud_name = page.inner_text("#hud-doc-name")
            self.assertEqual(hud_name, "e2e_sample.docx")

            # Verify Word Priority button is prominent
            word_btn_text = page.inner_text("#btn-dl-word")
            self.assertIn("TẢI FILE WORD (.DOCX)", word_btn_text)
            self.assertIn("KHUYÊN DÙNG", word_btn_text)

            # 6. Test PDF Download Button & Reminder Modal
            page.click("#btn-dl-pdf")
            page.wait_for_selector("#modal-pdf-progress", state="visible")

            # Verify 2-3 minute notice in modal
            modal_notice = page.inner_text(".modal-notice-banner")
            self.assertIn("2 - 3 phút", modal_notice)
            self.assertIn("Khuyên dùng", modal_notice)

            # Take screenshot of PDF Loading Modal & Split-View
            ss_path2 = os.path.join(ARTIFACTS_DIR, "luv_doctran_e2e_pdf_modal.png")
            page.screenshot(path=ss_path2)
            print(f"[E2E] Saved PDF progress modal screenshot: {ss_path2}")

            # Close modal via close button
            page.click(".btn-modal-close")
            time.sleep(0.5)

            # Take final full page screenshot of Split-View Studio
            ss_path3 = os.path.join(ARTIFACTS_DIR, "luv_doctran_e2e_split_view.png")
            page.screenshot(path=ss_path3)
            print(f"[E2E] Saved Split-View Studio screenshot: {ss_path3}")

            browser.close()


if __name__ == "__main__":
    unittest.main()
