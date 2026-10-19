"""
audit_gate.py
─────────────
Master Quality Gate wrapping the authoritative 7-Point OpenXML Audit Verifier.
Validates Zero Concatenation, Zero Double-Tab, 100% Pair Completeness, Zero Leaked English,
Ascending Header Order, Zero Clipped Rows, and Visual Parity.
"""

from __future__ import annotations
import os
import json
from dataclasses import dataclass, field, asdict
from typing import Dict, Any, Optional, List

from tests.audit_verifier import verify_bilingual_docx


@dataclass
class GateResult:
    """Consolidated outcome of the Master Quality Gate."""
    passed: bool
    concatenation_violations: int
    double_tab_violations: int
    pair_completeness: float
    leaked_english_count: int
    header_order_violations: int
    clipped_row_violations: int
    visual_parity_passed: bool
    mock_artifact_violations: int = 0
    details: Dict[str, Any] = field(default_factory=dict)
    ssim_score: Optional[float] = None
    page_parity: Optional[bool] = None
    json_report: Dict[str, Any] = field(default_factory=dict)

    def summary_table(self) -> str:
        """Renders an ASCII diagnostic summary table."""
        status_str = "PASS" if self.passed else "FAIL"
        lines = [
            "=" * 60,
            f"MASTER QUALITY GATE AUDIT RESULT: [{status_str}]",
            "=" * 60,
            f"1. Zero Concatenation Violations : {self.concatenation_violations} (Goal: 0)",
            f"2. Zero Internal Double-Tabs     : {self.double_tab_violations} (Goal: 0)",
            f"3. Pair Completeness Ratio       : {self.pair_completeness * 100:.1f}% (Goal: 100.0%)",
            f"4. Leaked English Sentences      : {self.leaked_english_count} (Goal: 0)",
            f"5. Header Order Violations       : {self.header_order_violations} (Goal: 0)",
            f"6. Clipped Row Height Violations : {self.clipped_row_violations} (Goal: 0)",
            f"7. Visual Parity Compliance      : {self.visual_parity_passed} (Goal: True)",
            f"8. Zero Mock Artifacts Violations: {self.mock_artifact_violations} (Goal: 0)",
        ]
        if self.ssim_score is not None:
            lines.append(f"   Visual SSIM Fidelity Score    : {self.ssim_score:.4f}")
        if self.page_parity is not None:
            lines.append(f"   Page Count Parity             : {self.page_parity}")
        lines.append("=" * 60)
        return "\n".join(lines)


class MasterQualityGate:
    """Evaluates generated DOCX against the authoritative 8-Point OpenXML Audit."""

    @classmethod
    def evaluate(cls, docx_path: str, pdf_path: Optional[str] = None) -> GateResult:
        if not os.path.exists(docx_path):
            raise FileNotFoundError(f"DOCX file not found: {docx_path}")

        raw_audit = verify_bilingual_docx(docx_path)

        gate_res = GateResult(
            passed=bool(raw_audit.get("passed", False)),
            concatenation_violations=int(raw_audit.get("concatenation_violations", 0)),
            double_tab_violations=int(raw_audit.get("double_tab_violations", 0)),
            pair_completeness=float(raw_audit.get("pair_completeness", 0.0)),
            leaked_english_count=int(raw_audit.get("leaked_english_count", 0)),
            header_order_violations=int(raw_audit.get("header_order_violations", 0)),
            clipped_row_violations=int(raw_audit.get("clipped_row_violations", 0)),
            visual_parity_passed=bool(raw_audit.get("visual_parity_passed", False)),
            mock_artifact_violations=int(raw_audit.get("mock_artifact_violations", 0)),
            details=raw_audit.get("details", {})
        )

        # Assemble diagnostic json report
        gate_res.json_report = {
            "docx_path": docx_path,
            "pdf_path": pdf_path,
            "overall_status": "PASS" if gate_res.passed else "FAIL",
            "audit_metrics": asdict(gate_res)
        }

        return gate_res
