# Project: LUV-DOC-TRAN Studio Pro — Nền Tảng Dịch Thuật Đa Ngữ & Đa Định Dạng Doanh Nghiệp

## Architecture
- **Architecture Overview**:
  - **Client Core Backend (`LUV-DOC-TRAN`)**: FastAPI backend with document extraction and interlinear translation.
    - Translation Engines: DeepSeek Enterprise Engine (default) and Google Fast Engine.
    - Format Handlers: Word (`.docx`, `.doc`), Excel (`.xlsx`, `.xls`), and PDF (`.pdf`) inputs. Outputs produce **strictly Word (.docx)** for Word/PDF inputs and **strictly Excel (.xlsx)** for Excel inputs.
    - Security & Licensing Client (`core/license_client.py`): Real-time online validation against Cloudflare Worker `/v1/validate` with asymmetric Ed25519 verification (`hsHHVWWmquFmME9wW4d2u1Qh5NdVtu3tktRDMl437Fo=`). Immediate auto-purge of local tokens (`data/license.sec`, `license.json`) on revocation and permanent translation access lock. Offline tolerance window bounded by `offline_until`.
  - **Studio Pro UI (`static/index.html`, `static/style.css`, `static/app.js`)**:
    - Theme: 100% Dark Obsidian (`#080b11`) with hairline borders (`#1e293b`), subtle glow, and glassmorphic cards (`backdrop-filter: blur(16px)`).
    - Multi-file Dropzone: Dynamic drag-and-drop animation, format badges for Word, PDF, Excel.
    - Smart Studio Config: Segmented Toggle for DeepSeek Enterprise vs Google Fast Engine, language selector.
    - Real-time Batch Queue: Per-file instantaneous download actions (`⚡ Tải Word (.docx)` or `📊 Tải Excel (.xlsx)`), and `🔍 Split-View Studio`. 100% elimination of PDF export buttons and progress modal/spinners.
    - Split-View Studio: Side-by-side comparison, synchronized scrolling, zoom, quick inline text edits.
    - Enterprise License Modal: Clean enterprise B2B activation, zero mention of `.exe`, live revoked state indication.
    - Toast Notifications: Replaces all browser native `alert()`.
  - **Cloudflare Worker & D1 Database**: Authoritative cloud licensing authority (`https://luv-license-api.luv-sds.workers.dev`). Issues and revokes Ed25519 signed tokens with claims.

## Feature Inventory
| # | Feature | Description | Milestone | Source | Status |
|---|---------|-------------|-----------|--------|:------:|
| 1 | F1: Real-time Online License Validation & Immediate Revocation | Call `POST /v1/validate` during license status checks and prior to translation. On HTTP 400/403/404 or revoked status, immediately purge local token files (`data/license.sec`, `license.json`), lock translation with HTTP 403, and update status to revoked. Support offline tolerance bounded strictly by `offline_until`. | M1 | ORIGINAL_REQUEST §R3 | DONE |
| 2 | F2: 100% Elimination of Chemical/CAS/ISO/GHS/MSDS Keywords | Cleanse all chemical/CAS/ISO/GHS/MSDS keywords across frontend (`index.html`, `app.js`), backend system prompts (`interlinear_translator.py`, `excel_interlinear_translator.py`, `server.py`), and metadata. Rebrand as "LUV-DOC-TRAN Studio - Nền Tảng Dịch Thuật Đa Ngữ & Đa Định Dạng Doanh Nghiệp". Maintain technical token protection regex as neutral numeric/technical identifiers. | M2 | ORIGINAL_REQUEST §R1 | PLANNED |
| 3 | F3: Instantaneous Word (.docx) & Excel (.xlsx) Output Only | Eliminate backward PDF export entirely. Accept PDF, Word, and Excel inputs. PDF and Word inputs produce strictly Word (`.docx`) output instantly. Excel inputs produce strictly Excel (`.xlsx`) output preserving formulas. Eliminate `/api/export-pdf`, remove all PDF download buttons, remove `#modal-pdf-progress` and radar spinners. | M3 | ORIGINAL_REQUEST §R2 | PLANNED |
| 4 | F4: Dark Obsidian & Glassmorphism Studio UI | 100% brand-new UI with `#080b11` background, `#1e293b` hairline borders, glassmorphic cards, segmented toggle for AI engines, sleek multi-file dropzone, per-file instant download buttons, side-by-side split-view comparison studio, clean enterprise license modal without any mention of exe, and toast notifications. | M4 | ORIGINAL_REQUEST §R4 | PLANNED |
| 5 | F5: 100% Automated Test Pass & Git Commit/Push | Pass 100% of automated test suites: Unit tests, Online Revocation suite, Keyword Cleanse audit, Translation Pipeline tests, and E2E Playwright tests. Stage, commit, and push cleanly to GitHub repository triggering automated Render deployment. | M5 | ORIGINAL_REQUEST §R5 | PLANNED |

## Milestones
| # | Name | Scope | Dependencies | Status |
|---|------|-------|-------------|:------:|
| M1 | Real-time License Validation & Revocation | `core/license_client.py`, `server.py`, `tests/test_online_revocation.py` | None | DONE |
| M2 | Keyword Cleanse & Rebranding | `core/interlinear_translator.py`, `core/excel_interlinear_translator.py`, `server.py`, `static/index.html`, `tests/test_keyword_cleanse.py` | None | PLANNED |
| M3 | Eliminate PDF Export (Word & Excel Only) | `core/interlinear_translator.py`, `core/excel_interlinear_translator.py`, `server.py`, `static/index.html`, `static/app.js`, `tests/test_translation_pipeline.py` | M2 | PLANNED |
| M4 | Dark Obsidian & Glassmorphism Studio UI | `static/index.html`, `static/style.css`, `static/app.js` | M2, M3 | PLANNED |
| M5 | Test Suite Verification & Remote Push | Full test execution (Unit, Revocation, Cleanse, E2E Playwright), Git commit & push | M1, M2, M3, M4 | PLANNED |

## Interface Contracts
### Real-time Online License Validation (`core/license_client.py` ↔ Cloudflare Worker)
- **Endpoint**: `POST https://luv-license-api.luv-sds.workers.dev/v1/validate`
- **Headers**:
  - `Content-Type: application/json`
  - `User-Agent: LUV-DOC/1.1.0`
- **Payload**:
  ```json
  {
    "token": "<base64url_claims>.<base64url_sig>",
    "device_id": "<64_char_lowercase_hex>",
    "app_version": "1.1.0"
  }
  ```
- **Success (HTTP 200)**:
  `{"ok": true, "valid": true, "token": "<renewed_token>", "expires_at": "...", "offline_until": "...", "plan": "enterprise"}`
- **Revoked / Expired (HTTP 400 / 401 / 403 / 404)**:
  `{"ok": false, "error": "license_revoked" | "license_denied", "message": "..."}`
  Trigger: `LUVLicenseClient.purge_license()`, `is_active = False`, `is_revoked = True`.

### Translation API Contract (`POST /api/translate`)
- **Inputs**: Multipart Form: `file` (Word/PDF/Excel), `source_lang`, `target_lang`, `engine` (default: deepseek), `device_id`.
- **Response**:
  ```json
  {
    "success": true,
    "filename": "<base>_SongNgu_<session_id>.<docx|xlsx>",
    "download_url": "/api/download/<filename>",
    "download_docx_url": "/api/download/<filename>", // If Word/PDF input
    "download_xlsx_url": "/api/download/<filename>", // If Excel input
    "output_format": "Word (.docx)" | "Excel (.xlsx)",
    "total_items": 12,
    "details": "..."
  }
  ```
  *(Zero PDF download URL, zero PDF conversion overhead).*

### Code Layout
- `core/`:
  - `license_client.py`: Ed25519 cryptography, DPAPI persistence, online validation `/v1/validate`, auto-purge, access control
  - `interlinear_translator.py`: Word & PDF interlinear translation, technical token protection, AI prompts
  - `excel_interlinear_translator.py`: Excel formula-preserving translation
  - `pdf_to_word_converter.py`: PDF input extraction to Word
- `static/`:
  - `index.html`: Studio Pro layout, segmented toggle, dropzone, batch queue, split-view HUD, enterprise modal
  - `style.css`: Dark Obsidian (`#080b11`), hairline borders (`#1e293b`), glassmorphic cards, responsive styles
  - `app.js`: Dropzone events, batch upload, queue manager, split-view studio, toast notifications, license events
- `server.py`: FastAPI server routes, license checks, instant download endpoints
- `tests/`:
  - `test_online_revocation.py`: Unit and integration tests for `/v1/validate`, revocation, and offline grace
  - `test_keyword_cleanse.py`: Zero-tolerance audit for chemical/CAS/ISO/GHS/MSDS terms
  - `test_translation_pipeline.py`: Word, PDF, Excel instant translation tests
  - `test_license_and_trial.py`: Core license activation and trial tests
  - `test_e2e_licensing.py`: Asymmetric crypto and end-to-end licensing tests
  - `test_e2e_frontend.py`: Playwright browser E2E test for Studio Pro UI
