import os
import sys
import unittest
from unittest.mock import patch
import tempfile
from pathlib import Path
from fastapi.testclient import TestClient
import docx
import openpyxl

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE_DIR)

from server import app
from core.license_client import LUVLicenseClient

class TestTranslationPipeline(unittest.TestCase):

    def setUp(self):
        self.client = TestClient(app)
        self.test_device = "TEST-E2E-DEV-001"
        self.patcher = patch.object(
            LUVLicenseClient,
            "check_access",
            return_value={"has_access": True, "access_type": "licensed", "tier": "Key Kích Hoạt (ENTERPRISE)", "revoked": False}
        )
        self.patcher.start()

    def tearDown(self):
        self.patcher.stop()

    def test_unactivated_translation_blocked(self):
        """Verify translation is strictly blocked with HTTP 403 when unactivated."""
        with patch.object(
            LUVLicenseClient,
            "check_access",
            return_value={"has_access": False, "access_type": "unactivated", "license_required": True, "revoked": False}
        ):
            resp = self.client.post("/api/translate", data={"source_lang": "en", "target_lang": "vi"}, files={"file": ("test.docx", b"PK\x03\x04", "application/docx")})
            self.assertEqual(resp.status_code, 403)
            self.assertTrue(resp.json().get("license_required"))

    def test_status_endpoint(self):
        """Verify system status endpoint."""
        resp = self.client.get("/api/status")
        self.assertEqual(resp.status_code, 200)
        data = resp.json()
        self.assertEqual(data["app_name"], "LUV-DOC-TRAN")
        self.assertEqual(data["status"], "online")

    def test_word_translation_and_word_priority(self):
        """Verify Word (.docx) translation executes and returns prioritized Word download."""
        # Create a sample docx
        doc = docx.Document()
        doc.add_heading("Safety Data Sheet (MSDS) Sample", level=1)
        p = doc.add_paragraph("Chemical Name: Sulfuric Acid (CAS: 7664-93-9). Highly corrosive liquid.")
        table = doc.add_table(rows=2, cols=2)
        table.cell(0, 0).text = "Property"
        table.cell(0, 1).text = "Value"
        table.cell(1, 0).text = "Flash Point"
        table.cell(1, 1).text = "Non-flammable"

        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tf:
            doc.save(tf.name)
            docx_path = tf.name

        try:
            with open(docx_path, "rb") as f:
                resp = self.client.post(
                    "/api/translate",
                    files={"file": ("sample_msds.docx", f, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")},
                    data={
                        "source_lang": "en",
                        "target_lang": "vi",
                        "output_mode": "bilingual",
                        "provider": "google",
                        "engine_mode": "auto",
                        "device_id": self.test_device
                    }
                )

            self.assertEqual(resp.status_code, 200)
            res = resp.json()
            self.assertTrue(res.get("success"))

            # Critical user requirement: Word Priority check
            self.assertIn("download_docx_url", res)
            self.assertEqual(res.get("primary_format"), "docx")
            self.assertIn("Word (.docx) đã sẵn sàng", res.get("word_priority_note", ""))

            # Verify file can be downloaded
            dl_url = res["download_docx_url"]
            dl_resp = self.client.get(dl_url)
            self.assertEqual(dl_resp.status_code, 200)
            self.assertGreater(len(dl_resp.content), 0)

            # Test on-demand PDF export endpoint with Word file
            docx_filename = os.path.basename(res["docx_path"])
            pdf_resp = self.client.post(
                "/api/export-pdf",
                data={"docx_filename": docx_filename}
            )
            self.assertEqual(pdf_resp.status_code, 200)
            pdf_res = pdf_resp.json()
            self.assertTrue(pdf_res.get("success"))
            self.assertIn("download_pdf_url", pdf_res)

        finally:
            if os.path.exists(docx_path):
                os.remove(docx_path)

    def test_excel_translation(self):
        """Verify Excel (.xlsx) translation works and preserves formula cells."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Inventory"
        ws["A1"] = "Chemical Item"
        ws["B1"] = "CAS Number"
        ws["C1"] = "Quantity"
        ws["A2"] = "Hydrochloric Acid 37%"
        ws["B2"] = "7647-01-0"
        ws["C2"] = 150
        ws["A3"] = "Total"
        ws["B3"] = ""
        ws["C3"] = "=SUM(C2:C2)"

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tf:
            wb.save(tf.name)
            xlsx_path = tf.name

        try:
            with open(xlsx_path, "rb") as f:
                resp = self.client.post(
                    "/api/translate",
                    files={"file": ("inventory.xlsx", f, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
                    data={
                        "source_lang": "en",
                        "target_lang": "vi",
                        "provider": "google",
                        "device_id": self.test_device
                    }
                )

            self.assertEqual(resp.status_code, 200)
            res = resp.json()
            self.assertTrue(res.get("success"))
            self.assertEqual(res.get("type"), "excel")
            self.assertIn("download_url", res)

            # Download translated workbook
            dl_resp = self.client.get(res["download_url"])
            self.assertEqual(dl_resp.status_code, 200)

            # Verify formula is preserved
            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as out_tf:
                out_tf.write(dl_resp.content)
                out_xlsx = out_tf.name

            try:
                out_wb = openpyxl.load_workbook(out_xlsx, data_only=False)
                out_ws = out_wb.active
                self.assertEqual(str(out_ws["C3"].value).upper(), "=SUM(C2:C2)")
                out_wb.close()
            finally:
                if os.path.exists(out_xlsx):
                    os.remove(out_xlsx)

        finally:
            if os.path.exists(xlsx_path):
                os.remove(xlsx_path)


if __name__ == "__main__":
    unittest.main()
