# E2E Test Infra: LUV-DOC-TRAN Studio

## Test Philosophy
- Opaque-box, requirement-driven derived from `ORIGINAL_REQUEST.md` (R1-R5).
- Methodology: Category-Partition + Boundary Value Analysis + Pairwise Combinations + Real-World Workloads.
- Strict pass criteria: 100% pass across all unit, integration, online revocation, and Playwright browser tests.

## Feature Inventory
| # | Feature | Source (Requirement) | Tier 1 | Tier 2 | Tier 3 |
|---|---------|---------------------|:------:|:------:|:------:|
| 1 | Real-time Cloudflare Worker Ed25519 Validate | ORIGINAL_REQUEST §R3 | 5 | 5 | ✓ |
| 2 | Auto-Purge Local Token on Revocation | ORIGINAL_REQUEST §R3 | 5 | 5 | ✓ |
| 3 | Translation Lock on Revoked License | ORIGINAL_REQUEST §R3 | 5 | 5 | ✓ |
| 4 | Offline Outage Grace Period Handling | ORIGINAL_REQUEST §R3 | 5 | 5 | ✓ |
| 5 | Cleanse Chemical/CAS/ISO/GHS/MSDS Keywords | ORIGINAL_REQUEST §R1 | 5 | 5 | ✓ |
| 6 | Word & PDF Input to Instant Word (.docx) Only | ORIGINAL_REQUEST §R2 | 5 | 5 | ✓ |
| 7 | Excel Input to Instant Excel (.xlsx) Only | ORIGINAL_REQUEST §R2 | 5 | 5 | ✓ |
| 8 | Total Elimination of PDF Export & Modals | ORIGINAL_REQUEST §R2 | 5 | 5 | ✓ |
| 9 | Dark Obsidian & Glassmorphism Studio UI | ORIGINAL_REQUEST §R4 | 5 | 5 | ✓ |
| 10 | Real-time Batch Queue & Instant Download Buttons | ORIGINAL_REQUEST §R4 | 5 | 5 | ✓ |
| 11 | Split-View Document Studio | ORIGINAL_REQUEST §R4 | 5 | 5 | ✓ |
| 12 | Enterprise License Modal (Zero .exe Mentions) | ORIGINAL_REQUEST §R4 | 5 | 5 | ✓ |

## Test Architecture
- **Unit & Licensing Runner**:
  - `python -m unittest tests/test_license_and_trial.py`
  - `python -m unittest tests/test_e2e_licensing.py`
- **Translation Pipeline Runner**:
  - `python -m unittest tests/test_translation_pipeline.py`
- **Frontend E2E Browser Runner**:
  - `python -m unittest tests/test_e2e_frontend.py` (Playwright Chromium)
- **Keyword Audit Command**:
  - `git grep -i -E "(hóa chất|hoa chat|mã cas|iso & ghs|msds)" -- "static/*" "templates/*"` (must return 0)

## Real-World Application Scenarios (Tier 4)
| # | Scenario | Features Exercised | Complexity |
|---|----------|--------------------|------------|
| 1 | Batch upload of mixed Word, Excel, and PDF corporate documents | F6, F7, F8, F9, F10 | High |
| 2 | Admin revokes license on Cloudflare D1 while client is running | F1, F2, F3, F12 | Critical |
| 3 | Client experiences network outage during translation | F1, F4 | Medium |
| 4 | End-to-end translation of 50-row multi-sheet Excel with formulas | F7, F10, F11 | High |
| 5 | Playwright E2E UI verification of Dark Obsidian Studio & Split-View | F8, F9, F10, F11, F12 | High |
