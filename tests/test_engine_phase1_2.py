"""
test_engine_phase1_2.py
───────────────────────
Comprehensive test suite for:
- Phase 1 Extractor (models, pdf_inspector, vector_extractor, table_grid_graph, image_clipper)
- Phase 2 Sanitizer (spatial_deinterleaver, paragraph_defragmenter, token_sanitizer, structure_normalizer)
- End-to-End integration from ExtractedDocument to SanitizedDocument
"""

import os
import unittest
import tempfile
import fitz  # PyMuPDF

from src.engine.phase1_extractor.models import (
    BBox, TextSpan, TextBlock, GridCellData, TableGridData,
    DrawingData, ImageData, ExtractedPage, ExtractedDocument
)
from src.engine.phase1_extractor.pdf_inspector import (
    PDFInspector, int_to_hex_color
)
from src.engine.phase1_extractor.vector_extractor import (
    VectorExtractor, rgb_tuple_to_hex
)
from src.engine.phase1_extractor.table_grid_graph import TableGridGraph
from src.engine.phase1_extractor.image_clipper import ImageClipper

from src.engine.phase2_sanitizer.spatial_deinterleaver import SpatialDeinterleaver
from src.engine.phase2_sanitizer.paragraph_defragmenter import ParagraphDefragmenter
from src.engine.phase2_sanitizer.token_sanitizer import TokenSanitizer
from src.engine.phase2_sanitizer.structure_normalizer import (
    StructureNormalizer, ROMAN_NUMERALS, GHS_SECTIONS
)
from src.engine.phase2_sanitizer import (
    SanitizedPage, SanitizedDocument
)


class TestPhase1Models(unittest.TestCase):
    """Test data models and coordinate math in Phase 1."""

    def test_bbox_operations(self):
        b1 = BBox(10.0, 20.0, 110.0, 70.0)
        self.assertEqual(b1.width, 100.0)
        self.assertEqual(b1.height, 50.0)
        self.assertEqual(b1.area, 5000.0)
        self.assertEqual(b1.center_x, 60.0)
        self.assertEqual(b1.center_y, 45.0)

        # Twips and EMUs
        self.assertEqual(b1.to_twips(), (200, 400, 2200, 1400))
        self.assertEqual(b1.to_emus(), (127000, 254000, 1397000, 889000))
        self.assertEqual(b1.to_tuple(), (10.0, 20.0, 110.0, 70.0))

        # Inversion normalization
        b_inv = BBox(100.0, 200.0, 50.0, 150.0)
        self.assertEqual(b_inv.x0, 50.0)
        self.assertEqual(b_inv.x1, 100.0)
        self.assertEqual(b_inv.y0, 150.0)
        self.assertEqual(b_inv.y1, 200.0)

        # Intersects & Contains
        b2 = BBox(50.0, 30.0, 80.0, 60.0)
        self.assertTrue(b1.contains(b2))
        self.assertTrue(b1.intersects(b2))

        b3 = BBox(200.0, 200.0, 250.0, 250.0)
        self.assertFalse(b1.intersects(b3))
        self.assertFalse(b1.contains(b3))

        # Union & Intersection
        union_box = b1.union(b3)
        self.assertEqual(union_box.to_tuple(), (10.0, 20.0, 250.0, 250.0))

        inter_box = b1.intersection(b2)
        self.assertIsNotNone(inter_box)
        self.assertEqual(inter_box.to_tuple(), (50.0, 30.0, 80.0, 60.0))
        self.assertIsNone(b1.intersection(b3))

        # IoU
        self.assertGreater(b1.iou(b2), 0.0)
        self.assertEqual(b1.iou(b3), 0.0)

    def test_text_span_and_block(self):
        s1 = TextSpan(
            text="• Danger: Highly flammable liquid",
            bbox=(50.0, 100.0, 250.0, 112.0),
            font_name="Arial-BoldMT",
            font_size=11.0,
            flags=16
        )
        self.assertTrue(s1.is_bold)
        self.assertFalse(s1.is_italic)
        self.assertEqual(s1.to_bbox().width, 200.0)

        s2 = TextSpan(
            text=" and vapor.",
            bbox=(250.0, 100.0, 320.0, 112.0),
            font_name="Arial-ItalicMT",
            font_size=11.0,
            flags=2
        )
        self.assertTrue(s2.is_italic)

        block = TextBlock(spans=[s1, s2])
        block.update_from_spans()

        self.assertEqual(block.text, "• Danger: Highly flammable liquid and vapor.")
        self.assertEqual(block.bbox, (50.0, 100.0, 320.0, 112.0))
        self.assertTrue(block.is_bullet)
        self.assertEqual(block.bullet_char, "•")
        self.assertTrue(block.is_bold)

    def test_table_grid_and_cell_models(self):
        c00 = GridCellData(row_idx=0, col_idx=0, bbox=(50.0, 100.0, 150.0, 120.0), text="Header 1")
        c01 = GridCellData(row_idx=0, col_idx=1, bbox=(150.0, 100.0, 350.0, 120.0), text="Header 2")
        c10 = GridCellData(row_idx=1, col_idx=0, bbox=(50.0, 120.0, 150.0, 140.0), text="Val 1")
        c11 = GridCellData(row_idx=1, col_idx=1, bbox=(150.0, 120.0, 350.0, 140.0), text="Val 2")

        table = TableGridData(
            rows=2,
            cols=2,
            col_widths=[100.0, 200.0],
            cells=[c00, c01, c10, c11],
            bbox=(50.0, 100.0, 350.0, 140.0)
        )
        self.assertEqual(table.get_cell(0, 1).text, "Header 2")
        self.assertEqual(table.get_cell(1, 0).text, "Val 1")
        self.assertIsNone(table.get_cell(5, 5))
        self.assertEqual(table.to_bbox().height, 40.0)

    def test_extracted_document_and_page(self):
        p1 = ExtractedPage(
            page_num=1,
            width=595.0,
            height=842.0,
            blocks=[TextBlock(text="Page 1 header", bbox=(50, 50, 200, 65))]
        )
        p2 = ExtractedPage(
            page_num=2,
            width=595.0,
            height=842.0,
            blocks=[TextBlock(text="Page 2 header", bbox=(50, 50, 200, 65))]
        )
        doc = ExtractedDocument(pages=[p1, p2], pdf_path="test.pdf")

        self.assertEqual(doc.page_count, 2)
        self.assertEqual(doc.get_page(1).text, "Page 1 header")
        self.assertEqual(doc.get_page(2).text, "Page 2 header")
        self.assertIsNone(doc.get_page(3))


class TestPhase1ExtractorComponents(unittest.TestCase):
    """Test PDFInspector, VectorExtractor, TableGridGraph, and ImageClipper."""

    @classmethod
    def setUpClass(cls):
        # Create an in-memory synthetic PDF with text, vector lines, and an image
        cls.doc = fitz.open()
        page = cls.doc.new_page(width=600, height=800)

        # 1. Text blocks
        page.insert_text(fitz.Point(50, 70), "SECTION 1: Identification", fontname="helv", fontsize=14)
        page.insert_text(fitz.Point(50, 100), "Product Name: Ultra Poly Resin", fontname="helv", fontsize=10)
        page.insert_text(fitz.Point(50, 115), "Manufacturer: Tech Chemicals Inc.", fontname="helv", fontsize=10)

        # 2. Table Vector lines (2x2 grid)
        # Horizontal lines: y=150, y=180, y=210
        page.draw_line(fitz.Point(50, 150), fitz.Point(350, 150), color=(0, 0, 0), width=1.0)
        page.draw_line(fitz.Point(50, 180), fitz.Point(350, 180), color=(0, 0, 0), width=1.0)
        page.draw_line(fitz.Point(50, 210), fitz.Point(350, 210), color=(0, 0, 0), width=1.0)
        # Vertical lines: x=50, x=150, x=350
        page.draw_line(fitz.Point(50, 150), fitz.Point(50, 210), color=(0, 0, 0), width=1.0)
        page.draw_line(fitz.Point(150, 150), fitz.Point(150, 210), color=(0, 0, 0), width=1.0)
        page.draw_line(fitz.Point(350, 150), fitz.Point(350, 210), color=(0, 0, 0), width=1.0)

        # Text inside table cells
        page.insert_text(fitz.Point(60, 170), "Property", fontname="helv", fontsize=10)
        page.insert_text(fitz.Point(160, 170), "Specification", fontname="helv", fontsize=10)
        page.insert_text(fitz.Point(60, 200), "Viscosity", fontname="helv", fontsize=10)
        page.insert_text(fitz.Point(160, 200), "500 - 800 mPa.s", fontname="helv", fontsize=10)

        # 3. Dense vector cluster (mock GHS diamond: multiple path strokes close together)
        for i in range(16):
            offset = i * 2
            page.draw_rect(fitz.Rect(400 + offset, 150 + offset, 450, 200), color=(1, 0, 0), fill=(1, 1, 0))

        # 4. Tiny embedded raster pixmap
        pix = fitz.Pixmap(fitz.csRGB, fitz.IRect(0, 0, 16, 16), False)
        pix.clear_with(255)
        page.insert_image(fitz.Rect(50, 250, 100, 300), pixmap=pix)

    @classmethod
    def tearDownClass(cls):
        cls.doc.close()

    def test_vector_extractor(self):
        page = self.doc[0]
        extractor = VectorExtractor(snap_tolerance=2.0)
        all_drawings, dense_clusters = extractor.extract_drawings(page)

        self.assertGreater(len(all_drawings), 5)
        h_lines = [d for d in all_drawings if d.is_horizontal]
        v_lines = [d for d in all_drawings if d.is_vertical]
        self.assertGreaterEqual(len(h_lines), 3)
        self.assertGreaterEqual(len(v_lines), 3)

        # Dense cluster should be detected
        self.assertGreaterEqual(len(dense_clusters), 1)

        # Stroke graph
        graph = extractor.build_stroke_graph(all_drawings)
        self.assertIn("intersections", graph)
        self.assertGreater(len(graph["intersections"]), 0)

    def test_table_grid_graph(self):
        page = self.doc[0]
        inspector = PDFInspector(doc=self.doc)
        raw_blocks = inspector._extract_text_blocks(page)

        extractor = VectorExtractor(snap_tolerance=2.0)
        all_drawings, _ = extractor.extract_drawings(page)

        t_graph = TableGridGraph(snap_tolerance=2.0)
        tables, remaining = t_graph.reconstruct_tables(raw_blocks, all_drawings, 600, 800)

        self.assertEqual(len(tables), 1)
        tbl = tables[0]
        self.assertEqual(tbl.rows, 2)
        self.assertEqual(tbl.cols, 2)
        self.assertTrue(tbl.is_bordered)

        # Check cell contents
        cell_00 = tbl.get_cell(0, 0)
        self.assertIsNotNone(cell_00)
        self.assertIn("Property", cell_00.text)
        self.assertEqual(cell_00.borders["top"]["val"], "single")

        cell_11 = tbl.get_cell(1, 1)
        self.assertIsNotNone(cell_11)
        self.assertIn("500 - 800", cell_11.text)

    def test_image_clipper(self):
        page = self.doc[0]
        clipper = ImageClipper(target_dpi=150, high_dpi=300)
        images = clipper.extract_page_images(page, self.doc)

        self.assertGreaterEqual(len(images), 1)
        img = images[0]
        self.assertGreater(len(img.image_bytes), 0)
        self.assertEqual(img.ext, "png")

        # Region rasterization with alpha
        clip_bytes = clipper.clip_region_raster(page, (50, 70, 200, 100), dpi=150)
        self.assertGreater(len(clip_bytes), 0)
        self.assertEqual(clip_bytes[:4], b"\x89PNG")

    def test_pdf_inspector_inspect_page(self):
        inspector = PDFInspector(doc=self.doc)
        page_data = inspector.inspect_page(0)

        self.assertEqual(page_data.page_num, 1)
        self.assertEqual(page_data.width, 600.0)
        self.assertEqual(page_data.height, 800.0)
        self.assertGreater(len(page_data.blocks), 0)
        self.assertEqual(len(page_data.tables), 1)
        self.assertGreater(len(page_data.drawings), 0)
        self.assertGreater(len(page_data.images), 0)

        # Margins should be estimated
        top, bottom, left, right = page_data.margins
        self.assertGreaterEqual(top, 18.0)
        self.assertGreaterEqual(left, 18.0)

    def test_strip_pdf_restrictions(self):
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tf_in:
            in_path = tf_in.name
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tf_out:
            out_path = tf_out.name

        try:
            self.doc.save(in_path)
            success = PDFInspector.strip_pdf_restrictions(in_path, out_path)
            self.assertTrue(success)
            self.assertTrue(os.path.exists(out_path))
            self.assertGreater(os.path.getsize(out_path), 0)
        finally:
            for p in (in_path, out_path):
                if os.path.exists(p):
                    try:
                        os.remove(p)
                    except Exception:
                        pass


class TestPhase2SanitizerComponents(unittest.TestCase):
    """Test SpatialDeinterleaver, ParagraphDefragmenter, TokenSanitizer, and StructureNormalizer."""

    def test_token_sanitizer(self):
        # 1. OCR typos
        self.assertEqual(TokenSanitizer.sanitize("Avoid hest and flame."), "Avoid heat and flame.")
        self.assertEqual(TokenSanitizer.sanitize("Keep away form ignition sources"), "Keep away from ignition sources")
        self.assertEqual(TokenSanitizer.sanitize("In teh container"), "In the container")
        self.assertEqual(TokenSanitizer.sanitize("Payment recieved"), "Payment received")
        self.assertEqual(TokenSanitizer.sanitize("BEN CAT CITY,BINH DUONG PROVINCE"), "BEN CAT CITY, BINH DUONG PROVINCE")

        # 2. Units spacing
        self.assertEqual(TokenSanitizer.sanitize("Flush eyes for 15minutes."), "Flush eyes for 15 minutes.")
        self.assertEqual(TokenSanitizer.sanitize("Take 500mg daily"), "Take 500 mg daily")
        self.assertEqual(TokenSanitizer.sanitize("Wait 24hours before use"), "Wait 24 hours before use")
        self.assertEqual(TokenSanitizer.sanitize("Concentration: 50ppm"), "Concentration: 50 ppm")

        # 3. Sentence boundary periods
        self.assertEqual(
            TokenSanitizer.sanitize("Flush for 15 minutes.Assure adequate flushing."),
            "Flush for 15 minutes. Assure adequate flushing."
        )
        self.assertEqual(
            TokenSanitizer.sanitize("Call physician.Consult safety data."),
            "Call physician. Consult safety data."
        )

        # 4. Comma spacing
        self.assertEqual(TokenSanitizer.sanitize("City,Country"), "City, Country")

        # 5. CamelCase brands & websites
        self.assertEqual(TokenSanitizer.sanitize("Visit ScienceLab.com"), "Visit Science Lab.com")
        self.assertEqual(TokenSanitizer.sanitize("Registered GmbH"), "Registered GMBH")
        self.assertEqual(TokenSanitizer.sanitize("Not vPvB"), "Not VPVB")

    def test_spatial_deinterleaver(self):
        # 1. Label Value recognition
        self.assertTrue(SpatialDeinterleaver.is_label_value_line("Appearance: Clear colorless liquid"))
        self.assertTrue(SpatialDeinterleaver.is_label_value_line("pH Value \t : 7.2"))
        self.assertFalse(SpatialDeinterleaver.is_label_value_line("SECTION 1: Identification"))
        self.assertFalse(SpatialDeinterleaver.is_label_value_line("Just a plain body sentence."))

        lbl, val = SpatialDeinterleaver.parse_label_value("Specific Gravity: 1.05 (±0.02)")
        self.assertEqual(lbl, "Specific Gravity")
        self.assertEqual(val, "1.05 (±0.02)")

        # 2. Multi-field line splitting
        multi_line = "Telephone: 0274 3555 208  Fax: 0274 3555 209  Email: contact@chem.com"
        split_fields = SpatialDeinterleaver.split_multi_field_line(multi_line)
        self.assertEqual(len(split_fields), 3)
        self.assertIn("Telephone: 0274 3555 208", split_fields[0])
        self.assertIn("Fax: 0274 3555 209", split_fields[1])
        self.assertIn("Email: contact@chem.com", split_fields[2])

        # 3. Double-tab stripping
        self.assertEqual(SpatialDeinterleaver.strip_double_tabs("Label\t\tValue"), "Label Value")

        # 4. Reconstruct borderless 2-column table
        b1 = TextBlock(text="Boiling Point: 180 °C", bbox=(50, 100, 300, 115))
        b2 = TextBlock(text="Flash Point: 65 °C", bbox=(50, 118, 300, 133))
        b3 = TextBlock(text="Solubility in Water: Complete", bbox=(50, 136, 300, 151))
        b_other = TextBlock(text="This product is strictly for industrial use.", bbox=(50, 170, 500, 185))

        tables, remaining = SpatialDeinterleaver.reconstruct_grid_data([b1, b2, b3, b_other], page_width=600.0)
        self.assertEqual(len(tables), 1)
        grid = tables[0]
        self.assertEqual(grid.rows, 3)
        self.assertEqual(grid.cols, 2)
        self.assertFalse(grid.is_bordered)
        self.assertEqual(grid.get_cell(0, 0).text, "Boiling Point")
        self.assertEqual(grid.get_cell(0, 1).text, ": 180 °C")
        self.assertEqual(len(remaining), 1)
        self.assertEqual(remaining[0].text, b_other.text)

    def test_paragraph_defragmenter(self):
        # 1. Continuous wrapped lines (should merge)
        upper = TextBlock(
            text="In case of accidental eye contact, thoroughly wash with plentiful water",
            bbox=(50, 100, 500, 112),
            font_size=10.0,
            is_bold=False
        )
        lower = TextBlock(
            text="for at least fifteen minutes and consult an ophthalmologist immediately.",
            bbox=(50, 115, 480, 127),
            font_size=10.0,
            is_bold=False
        )
        self.assertTrue(ParagraphDefragmenter.should_merge(upper, lower))

        defragged = ParagraphDefragmenter.defragment_blocks([upper, lower])
        self.assertEqual(len(defragged), 1)
        self.assertIn("wash with plentiful water for at least fifteen minutes", defragged[0].text)

        # 2. Punctuated line (should NOT merge)
        punctuated_upper = TextBlock(
            text="Store in a cool, well-ventilated location.",
            bbox=(50, 100, 350, 112),
            font_size=10.0
        )
        next_line = TextBlock(
            text="Keep containers tightly closed when not in use.",
            bbox=(50, 115, 380, 127),
            font_size=10.0
        )
        self.assertFalse(ParagraphDefragmenter.should_merge(punctuated_upper, next_line))
        self.assertEqual(len(ParagraphDefragmenter.defragment_blocks([punctuated_upper, next_line])), 2)

        # 3. Bullet list (should NOT merge)
        bullet_line = TextBlock(
            text="• Wear protective nitrile gloves.",
            bbox=(50, 115, 300, 127),
            font_size=10.0,
            is_bullet=True
        )
        self.assertFalse(ParagraphDefragmenter.should_merge(upper, bullet_line))

        # 4. Section header (should NOT merge)
        sec_header = TextBlock(
            text="SECTION 2: Hazards Identification",
            bbox=(50, 150, 300, 165),
            font_size=12.0
        )
        self.assertFalse(ParagraphDefragmenter.should_merge(upper, sec_header))

    def test_structure_normalizer(self):
        # 1. Section matching
        m1 = StructureNormalizer.match_section_header("SECTION 1: Chemical Product and Company Identification")
        self.assertIsNotNone(m1)
        self.assertEqual(m1[0], 1)

        m2 = StructureNormalizer.match_section_header("Section 4. First-Aid Measures")
        self.assertIsNotNone(m2)
        self.assertEqual(m2[0], 4)

        m16 = StructureNormalizer.match_section_header("XVI. Other Information")
        self.assertIsNotNone(m16)
        self.assertEqual(m16[0], 16)

        m_invalid = StructureNormalizer.match_section_header("General precautions to take")
        self.assertIsNone(m_invalid)

        # 2. Header normalization to Roman numerals
        h1 = StructureNormalizer.normalize_section_header("SECTION 1: Product Identification")
        self.assertTrue(h1.startswith("I."))

        h9 = StructureNormalizer.normalize_section_header("Section 9: Physical & Chemical Properties")
        self.assertTrue(h9.startswith("IX."))

        # 3. Inverted section header repair
        b_pre = TextBlock(text="MATERIAL SAFETY DATA SHEET", bbox=(50, 20, 300, 35))
        b_sec4 = TextBlock(text="SECTION 4: First Aid Measures", bbox=(50, 50, 250, 65))
        b_sec4_body = TextBlock(text="Inhalation: Move to fresh air.", bbox=(50, 70, 300, 85))
        b_sec1 = TextBlock(text="SECTION 1: Identification", bbox=(50, 100, 250, 115))
        b_sec1_body = TextBlock(text="Product: Resin X", bbox=(50, 120, 200, 135))
        b_sec2 = TextBlock(text="SECTION 2: Hazards Identification", bbox=(50, 150, 250, 165))
        b_sec2_body = TextBlock(text="Classification: Flammable", bbox=(50, 170, 250, 185))

        stream = [b_pre, b_sec4, b_sec4_body, b_sec1, b_sec1_body, b_sec2, b_sec2_body]
        reordered = StructureNormalizer.repair_inverted_headers(stream)

        self.assertEqual(reordered[0].text, b_pre.text)
        self.assertEqual(reordered[1].text, b_sec1.text)
        self.assertEqual(reordered[2].text, b_sec1_body.text)
        self.assertEqual(reordered[3].text, b_sec2.text)
        self.assertEqual(reordered[4].text, b_sec2_body.text)
        self.assertEqual(reordered[5].text, b_sec4.text)
        self.assertEqual(reordered[6].text, b_sec4_body.text)


class TestPhase1AndPhase2PipelineIntegration(unittest.TestCase):
    """End-to-End integration test from raw ExtractedDocument to SanitizedDocument."""

    def test_full_sanitization_pipeline(self):
        # Construct an extracted document with all real-world phenomena:
        # - unpunctuated wrapped lines
        # - OCR typo
        # - missing unit spacing and period spacing
        # - multi-field contact row
        # - inverted section headers
        b_pre = TextBlock(text="COMPANY LOGO & TITLE", bbox=(50, 30, 250, 45))

        # Inverted Section 2 before Section 1
        b_sec2 = TextBlock(text="SECTION 2: Hazards Identification", bbox=(50, 60, 300, 75))
        b_sec2_line1 = TextBlock(
            text="Avoid hest and flame. Flush with water for 15minutes.Assure",
            bbox=(50, 80, 450, 92),
            font_size=10.0
        )
        b_sec2_line2 = TextBlock(
            text="adequate ventilation during handling in teh plant.",
            bbox=(50, 95, 400, 107),
            font_size=10.0
        )

        b_sec1 = TextBlock(text="SECTION 1: Identification", bbox=(50, 130, 250, 145))
        b_contact = TextBlock(
            text="Telephone: 0274 3555 208  Fax: 0274 3555 209",
            bbox=(50, 150, 450, 165)
        )
        b_prop1 = TextBlock(text="Appearance: Clear liquid", bbox=(50, 180, 300, 195))
        b_prop2 = TextBlock(text="Odor: Characteristic", bbox=(50, 198, 300, 213))

        page = ExtractedPage(
            page_num=1,
            width=595.0,
            height=842.0,
            blocks=[b_pre, b_sec2, b_sec2_line1, b_sec2_line2, b_sec1, b_contact, b_prop1, b_prop2]
        )
        doc = ExtractedDocument(pages=[page], pdf_path="sample.pdf")

        sanitized_doc = SanitizedDocument.sanitize_extracted_document(doc)

        self.assertEqual(sanitized_doc.page_count, 1)
        sp = sanitized_doc.pages[0]

        # 1. Inverted section header repaired: Section 1 (I.) should appear before Section 2 (II.)
        header_blocks = [b.text for b in sp.blocks if b.text.startswith("I.") or b.text.startswith("II.")]
        self.assertEqual(len(header_blocks), 2)
        self.assertTrue(header_blocks[0].startswith("I."))
        self.assertTrue(header_blocks[1].startswith("II."))

        # 2. Token sanitizer verified: typos fixed, unit spacing fixed, sentence boundary fixed
        all_text = sp.text
        self.assertNotIn("hest and flame", all_text)
        self.assertIn("heat and flame", all_text)
        self.assertNotIn("15minutes.Assure", all_text)
        self.assertIn("15 minutes. Assure", all_text)
        self.assertNotIn("in teh plant", all_text)
        self.assertIn("in the plant", all_text)

        # 3. Paragraph defragmenter verified: wrapped lines consolidated into a single block
        merged_para = next(b for b in sp.blocks if "heat and flame" in b.text)
        self.assertIn("adequate ventilation during handling", merged_para.text)

        # 4. Spatial deinterleaver verified: Appearance & Odor clustered into borderless 2-column table
        self.assertGreaterEqual(len(sp.tables), 1)
        prop_table = next(t for t in sp.tables if not t.is_bordered)
        self.assertEqual(prop_table.cols, 2)
        self.assertEqual(prop_table.get_cell(2, 0).text, "Appearance")
        self.assertEqual(prop_table.get_cell(2, 1).text, ": Clear liquid")


if __name__ == "__main__":
    unittest.main()
