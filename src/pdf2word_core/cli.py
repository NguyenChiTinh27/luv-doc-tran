"""
cli.py
──────
Command Line Interface (CLI) for the Production-Grade Hybrid PDF-to-DOCX Conversion Engine.
Supports batch conversion, engine selection, visual self-healing toggle, and JSON diagnostic reporting.
"""

from __future__ import annotations
import os
import sys
import json
import argparse
import logging
from dataclasses import asdict

from src.pdf2word_core.models.schema import ConversionConfig
from src.pdf2word_core.converter import PDF2WordConverter


def setup_logging(verbose: bool = False):
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format="[%(asctime)s] [%(levelname)s] %(name)s: %(message)s",
        datefmt="%H:%M:%S"
    )


def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")

    parser = argparse.ArgumentParser(
        description="Production-Grade Hybrid Geometry-to-OpenXML PDF to DOCX Converter."
    )
    parser.add_argument("input_pdf", help="Path to input PDF file.")
    parser.add_argument("-o", "--output", help="Path to output .docx file (default: same name with .docx).")
    parser.add_argument(
        "--engine",
        choices=["auto", "engine_a", "engine_b"],
        default="auto",
        help="Select base engine: 'auto' (recommended), 'engine_b' (Word COM), or 'engine_a' (PyMuPDF)."
    )
    parser.add_argument("--no-qa", action="store_true", help="Disable Visual QA loop.")
    parser.add_argument("--no-heal", action="store_true", help="Disable auto-compacting self-healing pass.")
    parser.add_argument("--no-surgery", action="store_true", help="Disable OpenXML surgical pass.")
    parser.add_argument("--json", action="store_true", help="Output diagnostic report in JSON format.")
    parser.add_argument("-v", "--verbose", action="store_true", help="Enable verbose debug logging.")

    args = parser.parse_args()
    setup_logging(args.verbose)

    input_path = os.path.abspath(args.input_pdf)
    if not os.path.exists(input_path):
        print(f"Error: Input file does not exist: {input_path}", file=sys.stderr)
        sys.exit(1)

    if args.output:
        output_path = os.path.abspath(args.output)
    else:
        output_path = os.path.splitext(input_path)[0] + ".docx"

    config = ConversionConfig(
        engine_preference=args.engine,
        enable_surgical_pass=not args.no_surgery,
        enable_visual_qa=not args.no_qa,
        enable_self_heal=not args.no_heal
    )

    converter = PDF2WordConverter(config)
    report = converter.convert(input_path, output_path)

    if args.json:
        print(json.dumps(asdict(report), indent=2, ensure_ascii=False))
    else:
        print("\n" + "=" * 60)
        print("🎉 PDF-TO-DOCX CONVERSION COMPLETED")
        print("=" * 60)
        print(f"• Input:        {report.input_file}")
        print(f"• Output:       {report.output_file}")
        print(f"• Status:       {'SUCCESS ✅' if report.success else 'FAILED ❌'}")
        print(f"• Engine Used:  {report.engine_used}")
        print(f"• PDF Pages:    {report.pdf_pages}")
        print(f"• Page Parity:  {'100% MATCH ✅' if report.page_count_match else 'EXPANDED ⚠️'}")
        print(f"• SSIM Score:   {report.ssim_score:.4f}")
        print(f"• QA Passes:    {report.passes_executed}")
        print(f"• Duration:     {report.execution_time_seconds:.2f} seconds")
        if report.warnings:
            print("• Warnings:")
            for w in report.warnings:
                print(f"   - {w}")
        print("=" * 60 + "\n")

    sys.exit(0 if report.success else 1)


if __name__ == "__main__":
    main()
