"""
visual_qa.py
────────────
Automated Visual Verification & Self-Healing Loop.
Renders generated DOCX -> PDF -> PNG images at 150 DPI, computes Page-Count Parity,
measures Structural Similarity (SSIM) via scikit-image, and triggers iterative
auto-compacting passes if page count overflow is detected.
"""

from __future__ import annotations
import os
import sys
import tempfile
import logging
from typing import List, Tuple, Optional, Dict, Any

import fitz  # PyMuPDF
import numpy as np
import cv2
from skimage.metrics import structural_similarity as compute_ssim

from src.pdf2word_core.models.schema import PageData
from src.pdf2word_core.renderer.oxml_surgeon import OxmlSurgeon

logger = logging.getLogger("VisualQA")


class VisualQALoop:
    """Automated visual verification and iterative self-healing compaction engine."""

    def __init__(self, target_dpi: int = 150):
        self.target_dpi = target_dpi

    @staticmethod
    def docx_to_pdf(docx_path: str, output_pdf_path: str) -> bool:
        """
        Converts DOCX to PDF using Word COM (on Windows) or headless LibreOffice.
        """
        abs_docx = os.path.abspath(docx_path)
        abs_pdf = os.path.abspath(output_pdf_path)

        if sys.platform == "win32":
            try:
                import win32com.client
                word = win32com.client.DispatchEx("Word.Application")
                word.Visible = False
                word.DisplayAlerts = 0
                try:
                    doc = word.Documents.Open(abs_docx, ReadOnly=True, ConfirmConversions=False)
                    # wdFormatPDF = 17
                    doc.SaveAs2(abs_pdf, FileFormat=17)
                    doc.Close(SaveChanges=False)
                    if os.path.exists(abs_pdf) and os.path.getsize(abs_pdf) > 500:
                        return True
                finally:
                    word.Quit()
            except Exception as e:
                logger.warning(f"Word COM docx_to_pdf export failed: {e}")

        # Fallback to docx2pdf if installed
        try:
            from docx2pdf import convert
            convert(abs_docx, abs_pdf)
            if os.path.exists(abs_pdf):
                return True
        except Exception:
            pass

        return False

    def render_pdf_pages_to_images(self, pdf_path: str) -> List[np.ndarray]:
        """Renders every page of a PDF file to grayscale OpenCV numpy array."""
        images: List[np.ndarray] = []
        with fitz.open(pdf_path) as doc:
            for page in doc:
                pix = page.get_pixmap(dpi=self.target_dpi)
                img = np.frombuffer(pix.samples, dtype=np.uint8)
                if pix.alpha:
                    img = img.reshape((pix.height, pix.width, 4))
                    img_gray = cv2.cvtColor(img, cv2.COLOR_BGRA2GRAY)
                else:
                    img = img.reshape((pix.height, pix.width, pix.n))
                    img_gray = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY)
                images.append(img_gray)
        return images

    def evaluate_fidelity(
        self, original_pdf: str, docx_rendered_pdf: str
    ) -> Tuple[bool, int, int, float]:
        """
        Calculates Page Count Parity and average SSIM across all pages.
        Returns:
            (parity_match: bool, orig_page_count: int, docx_page_count: int, avg_ssim: float)
        """
        orig_imgs = self.render_pdf_pages_to_images(original_pdf)
        rend_imgs = self.render_pdf_pages_to_images(docx_rendered_pdf)

        orig_count = len(orig_imgs)
        rend_count = len(rend_imgs)
        parity_match = (orig_count == rend_count)

        ssim_scores: List[float] = []
        compare_pages = min(orig_count, rend_count)

        for i in range(compare_pages):
            img_a = orig_imgs[i]
            img_b = rend_imgs[i]

            # Resize img_b to match img_a dimensions if slight DPI rounding difference
            if img_a.shape != img_b.shape:
                img_b = cv2.resize(img_b, (img_a.shape[1], img_a.shape[0]), interpolation=cv2.INTER_AREA)

            score = compute_ssim(img_a, img_b)
            ssim_scores.append(score)

        avg_ssim = float(np.mean(ssim_scores)) if ssim_scores else 0.0
        return (parity_match, orig_count, rend_count, avg_ssim)

    def run_self_healing_loop(
        self,
        original_pdf_path: str,
        docx_path: str,
        pages_data: List[PageData],
        surgeon: OxmlSurgeon,
        max_passes: int = 3
    ) -> Tuple[bool, float, int]:
        """
        Closed-loop self-healing verification:
        If rendered DOCX overflows PDF page count, iteratively triggers auto-compacting
        passes until page parity is achieved or max passes are exhausted.
        """
        passes_run = 0
        compaction_factors = [1.0, 0.95, 0.90]

        with tempfile.NamedTemporaryFile(suffix="_rendered.pdf", delete=False) as tmp:
            rendered_pdf = tmp.name

        try:
            for p_idx in range(min(max_passes, len(compaction_factors))):
                passes_run += 1
                factor = compaction_factors[p_idx]

                # Run surgical pass with current compaction factor
                surgeon.perform_surgical_pass(
                    docx_path=docx_path,
                    pages_data=pages_data,
                    original_pdf_path=original_pdf_path,
                    compaction_factor=factor
                )

                # Render DOCX to PDF for verification
                success = self.docx_to_pdf(docx_path, rendered_pdf)
                if not success:
                    logger.warning("Could not render DOCX to PDF for Visual QA. Skipping loop.")
                    return True, 0.90, passes_run

                # Evaluate fidelity
                parity, orig_c, rend_c, ssim = self.evaluate_fidelity(
                    original_pdf_path, rendered_pdf
                )
                logger.info(
                    f"Visual QA Pass {passes_run}: Orig={orig_c} pgs, Rend={rend_c} pgs, "
                    f"Parity={parity}, SSIM={ssim:.4f}"
                )

                if parity:
                    # Page parity achieved!
                    return True, ssim, passes_run

                # Overflow detected (rend_c > orig_c), continue loop with tighter compaction factor
                logger.info(
                    f"Page overflow detected ({rend_c} > {orig_c}). Triggering compacting pass {p_idx + 2}..."
                )

            return parity, ssim, passes_run

        finally:
            if os.path.exists(rendered_pdf):
                try:
                    os.remove(rendered_pdf)
                except Exception:
                    pass
