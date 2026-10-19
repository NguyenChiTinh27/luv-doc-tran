# Project: End-to-End PDF-to-Bilingual-Word Conversion & QC Pipeline

## Architecture
The system converts and translates complex technical PDFs (MSDS/SDS, TDS, contracts, multi-column reports) into line-by-line bilingual Word (`.docx`) documents with 100% structural fidelity and zero defects.
Architecture is refactored into a unified 5-phase modular engine in `src/engine/` and a top-level CLI runner:

```text
src/engine/
├── __init__.py                 # Exports: run_pipeline, PipelineOrchestrator, EngineConfig
├── phase1_extractor/           # PyMuPDF rawdict, vector graphics, table grid graph, image clipper
├── phase2_sanitizer/           # X-coord de-interleaver, paragraph defragmenter, token spacing & punctuation fixer
├── phase3_bilingual/           # Offline zero-cost glossary & synthesizer (default) + optional DeepSeek/OpenAI client
├── phase4_surgeon/             # OpenXML row height unlocker, native header/footer promoter, tab/grid normalizer
├── phase5_qc_qa/               # Visual QA SSIM loop, page parity check, authoritative 7-point structural audit gate
└── orchestrator.py             # Pipeline coordination, state management, self-healing loop
CLI: pdf2bilingual_docx_engine.py  # Single-file & batch folder processing, JSON diagnostic report
```

## Feature Inventory
| # | Feature | Description | Milestone | Source | Status |
|---|---------|-------------|:---------:|--------|:------:|
| 1 | F1: Phase 1 Extractor | Extract text spans, bounding boxes, vector lines, table grid topology, and raster images from PDF via PyMuPDF rawdict. | M1 | ORIGINAL_REQUEST §R1 | PLANNED |
| 2 | F2: Phase 2 Sanitizer | Spatial X-coordinate de-interleaver, paragraph defragmenter for unpunctuated lines, token spacing (`15minutes.Assure`), OCR typo repair. | M1 | ORIGINAL_REQUEST §R1 | PLANNED |
| 3 | F3: Offline Glossary & Synthesizer | 100% offline zero-cost domain dictionary (16 GHS headers, 60+ chemical substances, HSE phrases, MSDS property labels) with ZeroCostGuard blocking paid API calls. | M2 | ORIGINAL_REQUEST §R2 | PLANNED |
| 4 | F4: Pairing Typography Engine | Pattern A (Body/Lists: `<w:keepNext/>`, `space_after=0pt` source + `#1B365D` italic translation `space_after=4pt`) and Pattern B (Table Cells: `<w:vAlign w:val="center"/>`, `<w:r><w:br/></w:r>` separation). | M2 | ORIGINAL_REQUEST §R2 | PLANNED |
| 5 | F5: External LLM Client Adapter | Optional DeepSeek/OpenAI translation client invoked only when user credentials are provided via `--online`. | M2 | ORIGINAL_REQUEST §R2 | PLANNED |
| 6 | F6: Dynamic Reflow Row Height Unlocker | Deep OpenXML post-processing replacing `<w:trHeight w:hRule="exact"/>` with `w:hRule="atLeast"` to prevent row clipping. | M3 | ORIGINAL_REQUEST §R3 | PLANNED |
| 7 | F7: Native Header & Footer Promoter | Promotes recurring body pagination/headers into native Word Section `<w:hdr>` / `<w:ftr>` with dynamic `PAGE` / `NUMPAGES` fields. | M3 | ORIGINAL_REQUEST §R3 | PLANNED |
| 8 | F8: Tab Normalizer & 2-Column Grid | Converts consecutive `Label : Value` lines into 2-column borderless fixed tables (28%/72% width) and strips 100% of internal double-tabs (`\t\t`). | M3 | ORIGINAL_REQUEST §R3 | PLANNED |
| 9 | F9: Anti-Overflow Line Compaction | Injects font character scaling (`<w:w w:val="94"/>`) and negative spacing (`<w:spacing w:val="-4"/>`) when cell overflow is detected. | M3 | ORIGINAL_REQUEST §R3 | PLANNED |
| 10 | F10: 7-Point OpenXML Audit Verifier | Authoritative XML gate verifying Zero Concatenation, Zero Double-Tab, 100% Pair Completeness, Zero Leaked English, Strict Header Order (I–XVI), Zero Clipped Rows, and Visual Parity. | M4 | ORIGINAL_REQUEST §R4 | PLANNED |
| 11 | F11: Visual QA SSIM Loop | Renders DOCX to PDF/images at 150 DPI and measures structural similarity (SSIM) against original PDF. | M4 | ORIGINAL_REQUEST §R4 | PLANNED |
| 12 | F12: Self-Healing Compaction Loop | Closed-loop iterative compaction (factors 1.0 -> 0.95 -> 0.90) triggered automatically if page count parity or visual alignment fails. | M4 | ORIGINAL_REQUEST §R4 | PLANNED |
| 13 | F13: Pipeline Orchestrator | Pipeline coordinator coordinating Phase 1 -> Phase 2 -> Phase 3 -> Phase 4 -> Phase 5 with state management and error handling. | M5 | ORIGINAL_REQUEST §R5 | PLANNED |
| 14 | F14: Unified CLI Runner | Top-level executable `pdf2bilingual_docx_engine.py` supporting single-file and batch folder processing, CLI flags, and standardized exit codes. | M5 | ORIGINAL_REQUEST §R5 | PLANNED |
| 15 | F15: JSON Diagnostic Reporting | Emits comprehensive JSON report alongside generated `.docx` containing pair completeness, audit status, execution metrics, and SSIM scores. | M5 | ORIGINAL_REQUEST §R5 | PLANNED |
| 16 | F16: E2E Test Suite (Tiers 1-4) & Zero Regression | 100% pass rate on requirement-driven opaque-box test suite; zero regressions on all 169 existing tests. | M6 | ORIGINAL_REQUEST §Acceptance | PLANNED |
| 17 | F17: Adversarial Coverage Hardening (Tier 5) | White-box stress-testing of extreme concatenations, malformed XML, 100-row tables, and inverted headers. | M6 | ORIGINAL_REQUEST §Acceptance | PLANNED |
| 18 | F18: Real-World Torture Sample Verification | 100% pass on 7-point audit for `MSDS-YTU0035-EN.pdf` (12 pages), `MSDS-YTU0070-EN.pdf` (5 pages), and `MSDS-FU43288-EN.doc` (17 tables). | M6 | ORIGINAL_REQUEST §Acceptance | PLANNED |

## Milestones
| # | Name | Scope | Dependencies | Status |
|---|------|-------|:------------:|:------:|
| M1 | Extraction & Sanitization Foundations | `src/engine/phase1_extractor/`, `src/engine/phase2_sanitizer/`, unit tests | None | IN_PROGRESS |
| M2 | Offline Bilingual Engine & Typography Pairing | `src/engine/phase3_bilingual/`, typography constants, zero-cost guard, unit tests | None | IN_PROGRESS |
| M3 | OpenXML Deep Surgeon & Layout Compactor | `src/engine/phase4_surgeon/`, row unlocker, header/footer promoter, tab/grid normalizer, unit tests | M1, M2 | PLANNED |
| M4 | 7-Point OpenXML Audit Verifier & Visual QA | `src/engine/phase5_qc_qa/`, authoritative 7-point audit, SSIM loop, self-healing, unit tests | M3 | PLANNED |
| M5 | Pipeline Orchestrator & Unified CLI | `src/engine/orchestrator.py`, `pdf2bilingual_docx_engine.py`, JSON diagnostic report, CLI tests | M1, M2, M3, M4 | PLANNED |
| M6 | E2E Test Suite Pass (Tiers 1-4) & Adversarial Hardening (Tier 5) | Pass 100% E2E tests, zero regression on 169 existing tests, torture test verification, Tier 5 hardening | M5, E2E Testing Track | PLANNED |

## Interface Contracts

### 1. Extractor ↔ Sanitizer (`phase1_extractor` ↔ `phase2_sanitizer`)
- **Data Model**: `ExtractedDocument` containing `List[ExtractedPage]`
- **ExtractedPage**:
  - `page_num: int`, `width: float`, `height: float`
  - `blocks: List[TextBlock]` (with `spans: List[TextSpan]`, `bbox: Tuple[float, float, float, float]`, `font_name`, `font_size`, `is_bold`, `is_italic`)
  - `tables: List[TableGridData]` (with `cells: List[GridCellData]`, `rows: int`, `cols: int`, `bbox`)
  - `drawings: List[DrawingData]`, `images: List[ImageData]`
- **Sanitizer Output**: `SanitizedDocument` with defragmented paragraphs, clustered `Label : Value` structures, and cleaned tokens.

### 2. Sanitizer ↔ Bilingual Engine (`phase2_sanitizer` ↔ `phase3_bilingual`)
- **Function**: `BilingualTranslator.translate_document(doc: SanitizedDocument, config: TranslationConfig) -> PairedDocument`
- **Typography Output**:
  - `BodyPair`: Source `ParagraphData` + Translated `ParagraphData` (Pattern A: `<w:keepNext/>`, `space_after=0pt`, translated `space_after=4pt`, color `#1B365D`, italic `<w:i/>`).
  - `CellPair`: Source `RunData` + explicit `<w:br/>` + Translated `RunData` (Pattern B: `#1B365D`, italic).
- **Cost Invariant**: `ZeroCostGuard` active by default; throws `ApiCostViolationError` on paid endpoint calls.

### 3. Bilingual Engine ↔ OpenXML Surgeon (`phase3_bilingual` ↔ `phase4_surgeon`)
- **Function**: `OpenXmlSurgeon.process_document(docx_path: str, options: SurgeonOptions) -> str`
- **Mutations**:
  - Unlock row heights: `//w:trHeight[@w:hRule='exact']` $\to$ `w:hRule="atLeast"`.
  - Strip double tabs: `\t\t` $\to$ separate table cells or clean space.
  - Native header/footer promotion: promote recurring headers to `<w:hdr>` / `<w:ftr>` with `PAGE` / `NUMPAGES` fields.
  - Compaction: inject `<w:w w:val="94"/>` and `<w:spacing w:val="-4"/>` if text overflow detected.

### 4. Surgeon ↔ QC/QA Gate (`phase4_surgeon` ↔ `phase5_qc_qa`)
- **Function**: `MasterQualityGate.evaluate(docx_path: str, pdf_path: Optional[str] = None) -> GateResult`
- **GateResult**:
  - `passed: bool` (Strict AND across all 7 checks)
  - `checks: Dict[str, CheckResult]` (check1 through check7)
  - `ssim_score: Optional[float]`, `page_parity: Optional[bool]`
  - `json_report: Dict[str, Any]`

### 5. CLI Interface Contract (`pdf2bilingual_docx_engine.py`)
- **Invocation**:
  ```powershell
  py -3.13 pdf2bilingual_docx_engine.py --input <path> [--output <path>] [--engine {offline,deepseek,openai}] [--no-qa] [--report]
  ```
- **Exit Codes**:
  - `0`: Success (DOCX generated, 100% 7-point audit pass)
  - `1`: General Error (Input file not found, unhandled exception)
  - `2`: Audit Gate Failure (One or more of the 7 checks failed)
  - `3`: Cost Violation Error (Attempted unauthorized paid API request)
- **Standard Output**: Formatted console progress, final audit status table, generated DOCX path.
- **Diagnostic JSON**: Emitted to `<output_base>.diagnostic.json` when `--report` is set.

## Code Layout
- `src/engine/`:
  - `__init__.py`: Package export interface
  - `orchestrator.py`: `PipelineOrchestrator`, `EngineConfig`
  - `phase1_extractor/`: `pdf_inspector.py`, `vector_extractor.py`, `table_grid_graph.py`, `image_clipper.py`
  - `phase2_sanitizer/`: `spatial_deinterleaver.py`, `paragraph_defragmenter.py`, `token_sanitizer.py`, `structure_normalizer.py`
  - `phase3_bilingual/`: `offline_glossary.py`, `rule_synthesizer.py`, `pairing_engine.py`, `leak_hunter.py`, `cost_guard.py`, `api_client.py`, `typography.py`
  - `phase4_surgeon/`: `row_height_unlocker.py`, `header_footer_promoter.py`, `tab_normalizer.py`, `oxml_injector.py`, `compactor.py`
  - `phase5_qc_qa/`: `structural_audit.py`, `visual_qa.py`, `page_parity.py`, `quality_gate.py`, `self_heal.py`, `checks/`
- Root CLI:
  - `pdf2bilingual_docx_engine.py`: CLI driver
- Tests:
  - `tests/e2e_engine/`: 4-Tier requirement-driven opaque-box test suite (`test_tier1_features.py`, `test_tier2_boundaries.py`, `test_tier3_combinations.py`, `test_tier4_scenarios.py`)
  - `tests/test_tier5_adversarial_bilingual_qc.py`: Tier 5 white-box adversarial suite
  - Existing 169 tests in `tests/test_*.py`
