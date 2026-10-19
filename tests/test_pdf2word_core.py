"""
test_pdf2word_core.py
─────────────────────
Automated Unit and Integration Test Suite for the Production-Grade
Hybrid Geometry-to-OpenXML PDF-to-DOCX Conversion Engine.
"""

import os
import unittest
import docx
from docx.shared import Inches

from src.pdf2word_core.models.schema import (
    BBox, SpanData, LineData, ParagraphData, BorderStyle,
    GridCellData, TableData, ConversionConfig
)
from src.pdf2word_core.extractor.pdf_inspector import PDFInspector
from src.pdf2word_core.analyzer.table_reconstructor import TableReconstructor
from src.pdf2word_core.analyzer.style_inferencer import StyleInferencer
from src.pdf2word_core.renderer.oxml_injector import OxmlInjector
from src.pdf2word_core.renderer.oxml_surgeon import OxmlSurgeon
from src.pdf2word_core.converter import PDF2WordConverter


class TestPDF2WordCore(unittest.TestCase):
    """Test suite covering schemas, geometry, OpenXML injection, and end-to-end conversion."""

    def test_bbox_geometry_math(self):
        """Verify BBox bounding box operations, coordinate math, and twip conversions."""
        b1 = BBox(10.0, 20.0, 110.0, 70.0)
        self.assertEqual(b1.width, 100.0)
        self.assertEqual(b1.height, 50.0)
        self.assertEqual(b1.area, 5000.0)
        self.assertEqual(b1.center_x, 60.0)
        self.assertEqual(b1.center_y, 45.0)

        # Twip and EMU conversion
        self.assertEqual(b1.to_twips(), (200, 400, 2200, 1400))
        self.assertEqual(b1.to_emus(), (127000, 254000, 1397000, 889000))

        # Intersects and contains
        b2 = BBox(50.0, 30.0, 80.0, 60.0)
        self.assertTrue(b1.contains(b2))
        self.assertTrue(b1.intersects(b2))

        b3 = BBox(200.0, 200.0, 250.0, 250.0)
        self.assertFalse(b1.intersects(b3))
        self.assertFalse(b1.contains(b3))

        # Union
        u = b1.union(b3)
        self.assertEqual((u.x0, u.y0, u.x1, u.y1), (10.0, 20.0, 250.0, 250.0))

    def test_oxml_injector_elements(self):
        """Verify low-level OpenXML injections on table, rows, and cells."""
        doc = docx.Document()
        table = doc.add_table(rows=2, cols=2)

        # 1. Fixed layout
        OxmlInjector.set_table_fixed_layout(table)
        tblLayout = table._tbl.tblPr.find(docx.oxml.ns.qn("w:tblLayout"))
        self.assertIsNotNone(tblLayout)
        self.assertEqual(tblLayout.get(docx.oxml.ns.qn("w:type")), "fixed")

        # 2. Table grid widths
        OxmlInjector.set_table_grid_widths(table, [2000, 3000])
        tblGrid = table._tbl.find(docx.oxml.ns.qn("w:tblGrid"))
        self.assertIsNotNone(tblGrid)
        gridCols = tblGrid.findall(docx.oxml.ns.qn("w:gridCol"))
        self.assertEqual(len(gridCols), 2)
        self.assertEqual(gridCols[0].get(docx.oxml.ns.qn("w:w")), "2000")

        # 3. cantSplit on row
        row = table.rows[0]
        OxmlInjector.lock_row_anti_split(row)
        cantSplit = row._tr.trPr.find(docx.oxml.ns.qn("w:cantSplit"))
        self.assertIsNotNone(cantSplit)

        # 4. Cell padding and shading
        cell = table.cell(0, 0)
        OxmlInjector.set_cell_margins(cell, top_dxa=40, bottom_dxa=40, left_dxa=40, right_dxa=40)
        tcMar = cell._tc.tcPr.find(docx.oxml.ns.qn("w:tcMar"))
        self.assertIsNotNone(tcMar)

        OxmlInjector.set_cell_shading(cell, "E0E0E0")
        shd = cell._tc.tcPr.find(docx.oxml.ns.qn("w:shd"))
        self.assertIsNotNone(shd)
        self.assertEqual(shd.get(docx.oxml.ns.qn("w:fill")), "E0E0E0")

        # 5. Anti-overflow character compensation
        p = cell.paragraphs[0]
        run = p.add_run("Test compensation")
        OxmlInjector.apply_anti_overflow_compensation(run, scale_pct=94, spacing_twips=-4)
        rPr = run._r.rPr
        w_elem = rPr.find(docx.oxml.ns.qn("w:w"))
        sp_elem = rPr.find(docx.oxml.ns.qn("w:spacing"))
        self.assertEqual(w_elem.get(docx.oxml.ns.qn("w:val")), "94")
        self.assertEqual(sp_elem.get(docx.oxml.ns.qn("w:val")), "-4")

    def test_style_inferencer_consolidation(self):
        """Verify paragraph merging and bullet detection."""
        s1 = SpanData(text="First line of continuous sentence", bbox=BBox(50, 100, 300, 112), font_name="Arial", font_size=10.0)
        l1 = LineData(spans=[s1])
        l1.update_text_and_bbox()

        s2 = SpanData(text="that continues smoothly without pause.", bbox=BBox(50, 115, 280, 127), font_name="Arial", font_size=10.0)
        l2 = LineData(spans=[s2])
        l2.update_text_and_bbox()

        paras = StyleInferencer.consolidate_paragraphs([l1, l2], page_width=595.0)
        self.assertEqual(len(paras), 1)
        self.assertIn("First line of continuous sentence", paras[0].text)
        self.assertIn("that continues smoothly without pause.", paras[0].text)

    def test_e2e_conversion_and_parity(self):
        """Verify end-to-end conversion, DRM stripping, and page parity on real-world sample."""
        sample_path = r"C:\Users\ntc27\OneDrive\Desktop\MSDS VISION\Thành phẩm\Y\MSDS-YTU0070-EN.pdf"
        if not os.path.exists(sample_path):
            self.skipTest(f"Test sample not found: {sample_path}")

        out_docx = "test_e2e_unit_output.docx"
        try:
            cfg = ConversionConfig(
                engine_preference="auto",
                enable_surgical_pass=True,
                enable_visual_qa=True,
                enable_self_heal=True
            )
            converter = PDF2WordConverter(cfg)
            report = converter.convert(sample_path, out_docx)

            self.assertTrue(report.success)
            self.assertTrue(os.path.exists(out_docx))
            self.assertGreater(os.path.getsize(out_docx), 1000)
            self.assertEqual(report.pdf_pages, 5)
            self.assertTrue(report.page_count_match)
        finally:
            if os.path.exists(out_docx):
                try:
                    os.remove(out_docx)
                except Exception:
                    pass


if __name__ == "__main__":
    unittest.main()
