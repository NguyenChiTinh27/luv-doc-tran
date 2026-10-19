# E2E Test Infra: PDF-to-Bilingual-Word Conversion & QC Pipeline

## Test Philosophy
- **Opaque-Box & Requirement-Driven**: Tests are derived strictly from `ORIGINAL_REQUEST.md` (sections 2026-10-19T08:38:46Z and 2026-10-19T10:30:02Z). Tests treat the conversion engine as an opaque black box via CLI and high-level Python entry points.
- **Methodology**: Systematic 4-tier testing:
  - Tier 1: Feature Coverage (Category-Partition)
  - Tier 2: Boundary Value Analysis & Corner Cases
  - Tier 3: Pairwise Cross-Feature Interaction Testing
  - Tier 4: Real-World Workload Testing (MSDS/SDS torture samples)
  - Tier 5: Adversarial Coverage Hardening (white-box stress testing)

## Feature Inventory & Test Mapping
| # | Feature | Source | Tier 1 | Tier 2 | Tier 3 | Tier 4 |
|---|---------|--------|:------:|:------:|:------:|:------:|
| F1 | Phase 1 PDF Extractor (geometry, text, vectors, images) | §R1 | 5 | 5 | ✓ | ✓ |
| F2 | Phase 2 Sanitizer (de-interleave, de-fragment, tokens) | §R1 | 5 | 5 | ✓ | ✓ |
| F3 | Phase 3 Offline Glossary & Zero-Cost Synthesizer | §R2 | 5 | 5 | ✓ | ✓ |
| F4 | Phase 3 Pairing Typography (Pattern A & Pattern B) | §R2 | 5 | 5 | ✓ | ✓ |
| F5 | Phase 3 External LLM Adapter (--online toggle) | §R2 | 5 | 3 | ✓ | - |
| F6 | Phase 4 Dynamic Row Height Unlocker (`atLeast`) | §R3 | 5 | 5 | ✓ | ✓ |
| F7 | Phase 4 Native Header/Footer Promotion (PAGE/NUMPAGES) | §R3 | 5 | 5 | ✓ | ✓ |
| F8 | Phase 4 Tab Normalizer & 2-Column Grid (strip `\t\t`) | §R3 | 5 | 5 | ✓ | ✓ |
| F9 | Phase 4 Anti-Overflow Line Compaction (`w:w=94`, `w:spacing=-4`) | §R3 | 5 | 5 | ✓ | ✓ |
| F10 | Phase 5 Authoritative 7-Point OpenXML Audit Gate | §R4 | 7 | 7 | ✓ | ✓ |
| F11 | Phase 5 Visual QA SSIM Loop & Page Parity | §R4 | 5 | 4 | ✓ | ✓ |
| F12 | Phase 5 Self-Healing Compaction Loop (1.0 -> 0.95 -> 0.90) | §R4 | 5 | 4 | ✓ | ✓ |
| F13 | Pipeline Orchestration & State Management | §R5 | 5 | 5 | ✓ | ✓ |
| F14 | Unified CLI Runner (`pdf2bilingual_docx_engine.py`) | §R5 | 5 | 5 | ✓ | ✓ |
| F15 | JSON Diagnostic Reporting Schema & Generation | §R5 | 5 | 5 | ✓ | ✓ |
| F16 | Zero Regression on 169 Existing Tests | Acceptance | All | - | - | ✓ |
| F17 | Tier 5 Adversarial Coverage Hardening | Acceptance | - | - | - | ✓ (26 tests) |
| F18 | Real-World Torture Sample Verification | Acceptance | - | - | - | ✓ (3 samples) |

## Test Architecture
- **Location**: `tests/e2e_engine/`
  - `test_tier1_features.py`: Feature-by-feature verification (Pattern A, Pattern B, Offline Glossary, Row height unlock, Native headers/footers, Tab stripper, 7-point audit checks, CLI flags).
  - `test_tier2_boundaries.py`: Boundary and corner cases (Empty PDFs, 100-page limits, unicode symbols, complex IUPAC formulas, extreme token concats, corrupted XML).
  - `test_tier3_combinations.py`: Cross-feature pairwise interactions (Dense vector drawings + multi-column table; unpunctuated wrapped lines + chemical formula in cell; CLI batch + offline mode).
  - `test_tier4_scenarios.py`: Real-world document pipelines (`MSDS-YTU0035-EN.pdf`, `MSDS-YTU0070-EN.pdf`, `MSDS-FU43288-EN.doc`).
- **Runner**: Python `unittest` / `pytest` under `py -3.13`
  - Command: `py -3.13 -m unittest discover -s tests/e2e_engine -p "test_*.py"`
- **Pass/Fail Semantics**: Exit code 0 on all tests passing. 7-Point Audit must report 0 violations. Zero paid API requests allowed in offline mode.

## Real-World Application Scenarios (Tier 4)
| # | Scenario | Features Exercised | Complexity |
|---|----------|--------------------|------------|
| 1 | REACH standard European 12-page chemical MSDS (`MSDS-YTU0035-EN.pdf`) | F1, F2, F3, F4, F6, F7, F8, F10, F11, F13, F14, F15 | High (12 pages, dense tables) |
| 2 | ScienceLab standard 5-page chemical MSDS with unpunctuated lines (`MSDS-YTU0070-EN.pdf`) | F1, F2, F3, F4, F6, F7, F8, F10, F13, F14 | Medium (5 pages, line wraps) |
| 3 | Zhanchen baseline 17-table document (`MSDS-FU43288-EN.doc`) | F1, F2, F3, F4, F6, F8, F9, F10, F13, F14 | High (17 tables, nested cells) |
| 4 | Batch folder execution over mixed PDFs with JSON report generation | F13, F14, F15 | Medium (Multi-file CLI) |
| 5 | Negative detection stress test (deliberately defective documents fail audit gate with code 2) | F10, F14, F15 | Medium (Robustness verification) |

## Coverage Thresholds
- Tier 1: ≥ 5 test cases per feature (~75 tests)
- Tier 2: ≥ 5 boundary test cases per feature (~75 tests)
- Tier 3: ≥ 15 pairwise cross-feature tests
- Tier 4: ≥ 5 realistic application-level scenarios
- Total E2E test target: ~170+ E2E tests + baseline 169 existing tests passing 100%.
