#!/usr/bin/env python3
"""
pdf2bilingual_docx_engine.py
────────────────────────────
Top-Level CLI Executable for the End-to-End PDF-to-Bilingual-Word Conversion & QC Pipeline.

Transforms technical PDFs/DOCs into pristine, line-by-line bilingual Word (.docx) documents
with 100% structural fidelity, zero concatenation, zero clipped cells, and native headers/footers.

Usage:
    python pdf2bilingual_docx_engine.py --input <file_or_dir> [--output <path>]
                                       [--engine {offline,deepseek,openai}]
                                       [--no-qa] [--report]

Exit Codes:
    0: Success (100% 7-Point Audit Pass)
    1: General Error (Input not found, fatal exception)
    2: Audit Gate Failure (One or more of the 7 checks failed)
    3: Cost Violation Error (Attempted unauthorized external API call)
"""

import os
import sys
import json
import argparse
import logging
from typing import List, Optional

if sys.stdout.encoding != "utf-8":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

# Add project root to sys.path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.engine import PipelineOrchestrator, EngineConfig, GateResult
from src.engine.phase3_bilingual.cost_guard import ApiCostViolationError

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s"
)
logger = logging.getLogger("pdf2bilingual_docx_engine")


def process_single_file(
    input_file: str,
    output_file: Optional[str],
    config: EngineConfig,
    emit_report: bool
) -> int:
    """Processes a single document and returns the corresponding exit code."""
    orchestrator = PipelineOrchestrator(config=config)

    try:
        docx_path, gate_result = orchestrator.run(input_file, output_path=output_file)
    except FileNotFoundError as e:
        logger.error(f"Input file error: {e}")
        return 1
    except ApiCostViolationError as e:
        logger.error(f"API Cost Violation: {e}")
        return 3
    except Exception as e:
        logger.exception(f"Unhandled pipeline failure on {input_file}: {e}")
        return 1

    # Print summary table
    print("\n" + gate_result.summary_table() + "\n")
    print(f"Generated Output Document: {docx_path}")

    # Emit JSON diagnostic report if requested
    if emit_report:
        report_path = f"{os.path.splitext(docx_path)[0]}.diagnostic.json"
        with open(report_path, "w", encoding="utf-8") as f:
            json.dump(gate_result.json_report, f, indent=2, ensure_ascii=False)
        print(f"Diagnostic Report Saved : {report_path}")

    if not gate_result.passed:
        logger.warning(f"Audit gate failed for {input_file}")
        return 2

    return 0


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Production-grade End-to-End PDF-to-Bilingual-Word Conversion & QC Pipeline."
    )
    parser.add_argument(
        "--input", "-i",
        required=True,
        help="Path to input PDF, DOC, or DOCX document (or directory of documents)."
    )
    parser.add_argument(
        "--output", "-o",
        default=None,
        help="Path to output .docx file (or output directory if input is directory)."
    )
    parser.add_argument(
        "--engine", "-e",
        choices=["offline", "deepseek", "openai"],
        default="offline",
        help="Translation engine to utilize. Default is offline zero-cost glossary & synthesizer."
    )
    parser.add_argument(
        "--api-key",
        default=None,
        help="Optional API key for online engine (DeepSeek or OpenAI)."
    )
    parser.add_argument(
        "--no-qa",
        action="store_true",
        help="Disable SSIM Visual QA loop and page parity verification (faster)."
    )
    parser.add_argument(
        "--report", "-r",
        action="store_true",
        help="Generate comprehensive JSON diagnostic report alongside .docx."
    )
    parser.add_argument(
        "--title",
        default=None,
        help="Optional document title to inject into native Word section headers."
    )

    args = parser.parse_args()

    config = EngineConfig(
        engine=args.engine,
        api_key=args.api_key,
        enable_visual_qa=not args.no_qa,
        title_text=args.title
    )

    input_path = os.path.abspath(args.input)

    # Batch directory processing
    if os.path.isdir(input_path):
        supported_exts = {".pdf", ".doc", ".docx"}
        files = [
            os.path.join(input_path, f) for f in os.listdir(input_path)
            if os.path.splitext(f)[1].lower() in supported_exts
        ]

        if not files:
            logger.warning(f"No supported documents (.pdf, .doc, .docx) found in {input_path}")
            return 1

        print(f"Batch processing {len(files)} documents in {input_path}...")
        overall_exit = 0
        for f in files:
            out_file = None
            if args.output:
                out_dir = os.path.abspath(args.output)
                os.makedirs(out_dir, exist_ok=True)
                out_file = os.path.join(out_dir, f"{os.path.splitext(os.path.basename(f))[0]}_bilingual.docx")

            code = process_single_file(f, out_file, config, args.report)
            if code != 0 and overall_exit == 0:
                overall_exit = code

        return overall_exit
    else:
        # Single file processing
        return process_single_file(input_path, args.output, config, args.report)


if __name__ == "__main__":
    sys.exit(main())
