"""
self_healing.py
───────────────
Phase 5 Automated Self-Healing Compaction Loop.
Iteratively compacts DOCX typography and table padding (1.0 -> 0.95 -> 0.90)
until page-count parity and zero-defect 7-point audit are achieved.
"""

from __future__ import annotations
import os
import tempfile
import logging
from typing import Optional, List

from src.engine.phase4_surgeon.surgeon import OpenXmlSurgeon
from src.engine.phase5_qc_qa.audit_gate import MasterQualityGate, GateResult
from src.engine.phase5_qc_qa.visual_qa import VisualQA

logger = logging.getLogger("SelfHealingLoop")


class SelfHealingLoop:
    """Orchestrates closed-loop iterative self-healing for DOCX generation."""

    def __init__(self, target_dpi: int = 150):
        self.surgeon = OpenXmlSurgeon(target_dpi=target_dpi * 2)
        self.visual_qa = VisualQA(target_dpi=target_dpi)

    def run_healing_loop(
        self,
        docx_path: str,
        pdf_path: Optional[str] = None,
        enable_visual_qa: bool = True,
        compaction_steps: Optional[List[float]] = None
    ) -> GateResult:
        """
        Executes self-healing compaction loop:
        1. Evaluates 7-point audit gate on docx_path.
        2. If visual QA enabled and PDF exists, checks page-count parity & SSIM.
        3. If page-count overflow is detected, progressively applies compaction factors (1.0 -> 0.95 -> 0.90).
        """
        if compaction_steps is None:
            compaction_steps = [1.0, 0.95, 0.90]

        best_result: Optional[GateResult] = None

        for step_idx, factor in enumerate(compaction_steps):
            # Apply surgery pass with current compaction factor
            if factor < 1.0 or step_idx > 0:
                logger.info(f"Self-healing pass {step_idx + 1}: Applying compaction factor {factor}")
                self.surgeon.process_document(docx_path, compaction_factor=factor)

            # Evaluate 7-Point Audit Gate
            gate_result = MasterQualityGate.evaluate(docx_path, pdf_path=pdf_path)

            # If visual QA enabled and original PDF is available, check parity
            if enable_visual_qa and pdf_path and os.path.exists(pdf_path):
                with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp_pdf:
                    tmp_pdf_path = tmp_pdf.name

                try:
                    exported = self.visual_qa.docx_to_pdf(docx_path, tmp_pdf_path)
                    if exported:
                        parity, orig_pgs, docx_pgs, ssim = self.visual_qa.evaluate_fidelity(
                            pdf_path, tmp_pdf_path
                        )
                        gate_result.page_parity = parity
                        gate_result.ssim_score = ssim

                        # If parity matches and audit passed, we are golden
                        if parity and gate_result.passed:
                            logger.info(f"Self-healing succeeded at factor {factor} (Parity={parity}, SSIM={ssim:.4f})")
                            best_result = gate_result
                            break
                        elif docx_pgs > orig_pgs and step_idx < len(compaction_steps) - 1:
                            logger.warning(f"Page overflow detected (orig={orig_pgs}, docx={docx_pgs}). Tightening compaction...")
                            best_result = gate_result
                            continue
                except Exception as e:
                    logger.warning(f"Visual QA check encountered non-fatal error: {e}")
                finally:
                    if os.path.exists(tmp_pdf_path):
                        try:
                            os.remove(tmp_pdf_path)
                        except Exception:
                            pass

            best_result = gate_result
            if gate_result.passed:
                break

        return best_result or MasterQualityGate.evaluate(docx_path, pdf_path=pdf_path)
