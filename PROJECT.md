# Project: End-to-End PDF-to-Bilingual-Word Conversion & QC Pipeline (MSDS/SDS Zero-Defect Overhaul)

## Architecture
The system converts and translates complex technical PDFs (MSDS/SDS, TDS, contracts, multi-column reports) into line-by-line bilingual Word (`.docx`) documents with 100% structural fidelity and zero defects.
Architecture is modularized across `src/engine/` and unified in `core/bilingual_qc_v2.py`:

```text
src/engine/
├── __init__.py                 # Exports: run_pipeline, PipelineOrchestrator, EngineConfig
├── phase1_extractor/           # PyMuPDF rawdict, vector graphics, table grid graph, image clipper
├── phase2_sanitizer/           # Multi-field splitter, merge_continuation_lines, paragraph defragmenter
├── phase3_bilingual/           # FullSentenceTranslator (Google GTX + translation_cache.json), glossary, leak hunter, hard assertion
├── phase4_surgeon/             # Body header purger, sectPr-safe promoter, 2-column table restructurer, row height unlocker
├── phase5_qc_qa/               # Visual QA, 7-point audit gate + mock artifact gate
└── orchestrator.py             # Pipeline coordination, state management, self-healing loop
core/bilingual_qc_v2.py         # Unified production pipeline engine
CLI: pdf2bilingual_docx_engine.py  # Single-file & batch folder processing, JSON diagnostic report
```

## Feature Inventory
| # | Feature | Description | Milestone | Source | Status |
|---|---------|-------------|:---------:|--------|:------:|
| 1 | F1: True Full-Sentence AI Translation Engine | Eliminate `_fallback_word_replace` and `(Thông tin)` injection. Translation must use EXCLUSIVELY AI LLM (DeepSeek Chat / OpenAI API) with expert MSDS/SDS system prompt + master `translation_cache.json`. Zero Google GTX or dictionary substitution rules. | M1 | ORIGINAL_REQUEST §R1 | PLANNED |
| 2 | F2: Regulatory Abbreviation Shield | Shield abbreviations pre-translation (`UN No.` -> `Số UN`, `EMS No.` -> `Số EMS`, `CAS No.` -> `Số CAS`) and post-translation regex repair in LeakHunter. | M1 | ORIGINAL_REQUEST §R1 | PLANNED |
| 3 | F3: Hard Output Quality Assertion | Raise `TranslationQualityError` if any translation contains `(Thông tin)`, `UN không có`, `EMS không có`, `can Dạng`, or >2 common untranslated English words. | M1 | ORIGINAL_REQUEST §R1 | PLANNED |
| 4 | F4: Section Header Splitting & Promotion | Split `MATERIAL SAFETY DATA SHEETAccording...` into 2 distinct lines (Title and Regulation) and place strictly in `doc.sections[0].header`. | M2 | ORIGINAL_REQUEST §R2 | PLANNED |
| 5 | F5: Purge Recurring Body Headers | Remove 100% of recurring header tables/paragraphs containing `MATERIAL SAFETY DATA SHEET` or `GB/T16483` from `<w:body>` (with strict Section 15 regulatory exemption and `<w:sectPr>` preservation). | M2 | ORIGINAL_REQUEST §R2 | PLANNED |
| 6 | F6: Cross-Page Sentence Stitching | Pre-translation cross-page sentence stitching when body headers/pagination are dropped: stitch unclosed clauses to lowercase continuation lines across table cells and body paragraphs. | M2 | ORIGINAL_REQUEST §R2 | PLANNED |
| 7 | F7: Multi-Field Row Splitting (I, II, XIV) | Section I multi-field line splitting with compound lookahead (`Emergency telephone`, `E-mail address`); Section II split into 3 rows; Section XIV split into 8 1-to-1 rows; repair Table 38 `dangerous Measures goods`. | M3 | ORIGINAL_REQUEST §R3 | PLANNED |
| 8 | F8: Continuation Line Merger (`merge_continuation_lines`) | Pre-translation merger of hanging lines starting with lowercase or following prepositions/conjunctions/commas into parent Label or Value across paragraphs and cells without colon blocking. | M3 | ORIGINAL_REQUEST §R4 | PLANNED |
| 9 | F9: Clean 2-Column Tables & Colspan Normalization | Convert 3-column `[Label, ':', Value]` tables into 2 columns (28%/72% width) with zero spurious colspans; strictly preserve Section III at 3 columns and Section VIII at 4 columns. | M3 | ORIGINAL_REQUEST §R5 | PLANNED |
| 10 | F10: E2E Integration & Automated Verification Gate | Automated test harness running torture samples (`test_zp76011`, `MSDS-YTU0035-EN.pdf`, `MSDS-YTU0070-EN.pdf`, `MSDS-FU43288-EN.doc`) with 100% pass on 7-point audit and 0 mock artifacts. | M4 | ORIGINAL_REQUEST §Acceptance | PLANNED |

## Milestones
| # | Name | Scope | Dependencies | Status |
|---|------|-------|:------------:|:------:|
| M1 | Full-Sentence Neural Translation & Cache Engine | `src/engine/phase3_bilingual/`, `core/bilingual_qc_v2.py`, `translation_cache.json`, abbreviation shielding, hard output assertions | None | DONE |
| M2 | Body Header Purge & Cross-Page Sentence Stitching | `src/engine/phase4_surgeon/header_footer_promoter.py`, `core/bilingual_qc_v2.py`, Section 15 exemption guard, `<w:sectPr>` preservation, pre-translation stitching | None | IN_PROGRESS |
| M3 | Multi-Field Splitting, Continuation Merger & 2-Col Grid | `src/engine/phase2_sanitizer/`, `core/bilingual_qc_v2.py`, `merge_continuation_lines`, Sections I, II, XIV splitting, Table 38 repair, 28%/72% 2-column tables | None | PLANNED |
| M4 | E2E Integration, Torture Sample Verification & Audit Gates | Full pipeline run on torture samples, `tests/audit_verifier.py`, `tests/test_fix_5_errors.py`, CLI runner, forensic integrity verification | M1, M2, M3 | PLANNED |

## Interface Contracts

### 1. Translation Engine Contract (`phase3_bilingual` / `core/bilingual_qc_v2.py`)
- **Class**: `FullSentenceTranslator` / `BilingualTranslator`
- **Provider**: EXCLUSIVELY AI LLM (DeepSeek Chat API / OpenAI API). No Google GTX or dictionary rule substitutions.
- **System Prompt**: Professional chemical MSDS/SDS & technical document translation expert with domain accuracy and line-by-line pairing.
- **Cache**: Master `translation_cache.json` stores verified AI translations for fast recall and budget conservation.
- **Assertion**:
  - `"(Thông tin)" in vi` -> raise `TranslationQualityError`
  - `"UN không có"` or `"EMS không có"` or `"CAS không có"` or `"can Dạng"` in vi -> raise `TranslationQualityError`
  - Count of untranslated English words > 2 in Vietnamese string -> raise `TranslationQualityError`

### 2. Header Purge & Stitching Contract (`phase4_surgeon` / `core/bilingual_qc_v2.py`)
- **Class**: `BodyHeaderPurgeAndStitcher`
- **Signature**: `purge_and_stitch(doc: docx.Document) -> Dict[str, int]`
- **Invariants**:
  - `doc.sections[0].header.paragraphs[0]` contains 2 distinct lines:
    Line 1: `MATERIAL SAFETY DATA SHEET\nBẢNG DỮ LIỆU AN TOÀN HÓA CHẤT`
    Line 2: `According to regulation GB/T16483, GB/T17519\nTheo quy định GB/T16483, GB/T17519`
  - Count of `MATERIAL SAFETY DATA SHEET` in `<w:body>` paragraphs and tables == 0.
  - Section 15 legal citations (`General rules for preparation... (GB/T 16483-2008)`) are 100% preserved.
  - Paragraphs containing `<w:sectPr>` are protected from deletion.

### 3. DOM Sanitizer & Table Contract (`phase2_sanitizer` / `core/bilingual_qc_v2.py`)
- **Function**: `merge_continuation_lines(doc: docx.Document) -> None`
- **Function**: `split_multifield_blocks(doc: docx.Document) -> None`
- **Function**: `restructure_tables(doc: docx.Document) -> None`
- **Invariants**:
  - Section I multi-field lines are split into independent rows without severing `Emergency telephone`.
  - Section II has 3 distinct rows: `- Hazard pictogram`, `- Signal word`, `- Hazard statements`.
  - Section XIV has 8 distinct 1-to-1 rows, and Table 38 `Measures` is merged into Row 0, cleaning `dangerous Measures goods`.
  - Borderless `Label : Value` tables have exactly 2 columns (`len(row.cells) == 2`) with widths 28% and 72%.
  - Section III remains 3 columns (`['Ingredients', 'CAS No.', 'Weight(%)']`).
  - Section VIII remains 4 columns (`['Component name', 'Reference standard', 'PC-TWA (mg/m3)', 'PC-STEL (mg/m3)']`).

## Code Layout
- `src/engine/`:
  - `phase2_sanitizer/`: `spatial_deinterleaver.py`, `paragraph_defragmenter.py`, `token_sanitizer.py`, `structure_normalizer.py`
  - `phase3_bilingual/`: `full_sentence_translator.py`, `offline_glossary.py`, `rule_synthesizer.py`, `api_client.py`, `leak_hunter.py`, `cost_guard.py`
  - `phase4_surgeon/`: `header_footer_promoter.py`, `row_height_unlocker.py`, `tab_normalizer.py`, `oxml_injector.py`
  - `phase5_qc_qa/`: `audit_gate.py`, `visual_qa.py`, `checks/`
  - `orchestrator.py`: Pipeline coordination
- `core/bilingual_qc_v2.py`: Production engine combining all phases
- `translation_cache.json`: Master persistent disk cache
- Root CLI: `pdf2bilingual_docx_engine.py`
- Tests: `tests/audit_verifier.py`, `tests/test_fix_5_errors.py`, `tests/test_bilingual_qc_v2.py`, `tests/test_engine_*.py`
