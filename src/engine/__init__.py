"""
src/engine/__init__.py
──────────────────────
End-to-End PDF-to-Bilingual-Word Conversion & QC Pipeline Engine.
Exports the public top-level interface for multi-phase document processing.
"""

from typing import Optional, Tuple
from src.engine.orchestrator import PipelineOrchestrator, EngineConfig
from src.engine.phase4_surgeon import OpenXmlSurgeon
from src.engine.phase5_qc_qa import MasterQualityGate, GateResult, VisualQA, SelfHealingLoop


def run_pipeline(
    input_path: str,
    output_path: Optional[str] = None,
    config: Optional[EngineConfig] = None
) -> Tuple[str, GateResult]:
    """Convenience helper to run the full 5-phase pipeline on a document."""
    orchestrator = PipelineOrchestrator(config=config)
    return orchestrator.run(input_path=input_path, output_path=output_path)


__all__ = [
    "run_pipeline",
    "PipelineOrchestrator",
    "EngineConfig",
    "GateResult",
    "MasterQualityGate",
    "OpenXmlSurgeon",
    "VisualQA",
    "SelfHealingLoop",
]
