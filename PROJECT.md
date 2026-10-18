# Project: LUV-DOC-TRAN & TitanPDF Authentic Cloudflare Worker & Ed25519 Licensing Integration

## Architecture
- **Architecture Overview**:
  - **Cloudflare Worker & D1 Database**: Authoritative cloud licensing authority (`https://luv-license-api.luv-sds.workers.dev`). Issues Ed25519 signed tokens with asymmetric cryptography (Public Key: `hsHHVWWmquFmME9wW4d2u1Qh5NdVtu3tktRDMl437Fo=`).
  - **LUV-DOC-TRAN Client**: Production document translation service with FastAPI backend. Validates license keys directly against Cloudflare Worker `/v1/activate`, verifies Ed25519 signatures locally via `cryptography`, and securely persists signed claims in `data/license.sec`.
  - **TitanPDF Admin Dashboard**: Web management interface. Connects via Admin API Token (`rHHfK9eC_8lTH02SP0CQPnyXAqMMDUXxwV2MiPothM7Avc4TxNZgnkKp5CFsIBP_`) to Cloudflare Worker Admin endpoints to create genuine keys, list licenses, revoke keys, and reset device bindings.
  - **E2E Testing Track**: Multi-tiered opaque-box and white-box test suite executing unit tests, API boundary tests, and integration workflows.

## Feature Inventory
| # | Feature | Description | Milestone | Source | Status |
|---|---------|-------------|-----------|--------|:------:|
| 1 | F1: Ed25519 Cryptographic Verification | Authentic asymmetric verification using `Ed25519PublicKey.verify()` with public key `hsHHVWWmquFmME9wW4d2u1Qh5NdVtu3tktRDMl437Fo=`. 100% elimination of mock/pseudo-hash code. | M1 | Survey / ORIGINAL_REQUEST §R1 | DONE |
| 2 | F2: Cloudflare Worker Client Activation & Validation | Connect to `/v1/activate` and `/v1/validate` with `User-Agent: LUV-DOC/1.1.0`. 2-step fallback (`DOC` -> `SDS`) for owner key compatibility. Catch HTTP 400/403/404 and return Vietnamese user-friendly errors. | M1 | Survey / ORIGINAL_REQUEST §R1 | DONE |
| 3 | F3: Secure License Persistence (`data/license.sec`) | Store verified tokens in `data/license.sec` with platform-aware security (DPAPI on Windows, safe file envelope on Linux/Docker/Render). | M1 | Survey / ORIGINAL_REQUEST §R1 | DONE |
| 4 | F4: LUV-DOC-TRAN Dependency & Deployment Hygiene | Add `cryptography>=42.0.0` to `requirements.txt`. Clean up dead `TITAN_LICENSE_VERIFY_URL` tunnel in `render.yaml`. Update UI placeholder to `LUV-XXXXX-XXXXX-XXXXX-XXXXX`. | M1 | Survey / Explorer 1 | DONE |
| 5 | F5: TitanPDF Live License Admin Backend | Remove `clean_key.startswith("LUV-")` bypass in `server.py` across all routes. Fix field mapping (`license_key_prefix`, `active_devices`/`max_devices`). Implement `POST /api/admin/licenses/reset-devices` and update `revoke_key` with UUID `license_id`. Eliminate mock key fallback. Add alias `GET /api/admin/licenses/list`. | M2 | Survey / Explorer 2 & ORIGINAL_REQUEST §R2 | DONE |
| 6 | F6: TitanPDF Live License Admin Frontend | Display live D1 licenses with accurate prefixes and device counts in `static/license_manager.js`. Add "Reset Máy" action button. Pass `item.id` for revoke and reset-devices. | M2 | Survey / Explorer 2 & ORIGINAL_REQUEST §R2 | DONE |
| 7 | F7: E2E Automated Test Suite & Adversarial Testing | Automated test suite verifying 100% rejection of bogus/fake keys, 100% success of authentic keys, token refresh, admin CRUD operations, and cross-feature workflows (41/41 tests pass). | E2E / M3 | Survey / ORIGINAL_REQUEST §R3 | DONE |
| 8 | F8: Git Commit & Remote Push | Commit and push verified code to GitHub remotes (`luv-doc-tran` and `titanpdf`) triggering automatic Render deployment. | M4 | Survey / ORIGINAL_REQUEST §R3 | IN_PROGRESS |

## Milestones
| # | Name | Scope | Dependencies | Status |
|---|------|-------|-------------|:------:|
| M1 | LUV-DOC-TRAN Authentic Licensing Integration | `core/license_client.py`, `server.py`, `requirements.txt`, `render.yaml`, `static/index.html`, `data/license.sec` | None | DONE |
| M2 | TitanPDF Live License Admin Integration | `c:\Users\ntc27\OneDrive\文件\TitanPDF` (`server.py`, `core/license_manager.py`, `static/license_manager.js`, `static/license_manager.css`) | None | DONE |
| E2E | E2E Testing Suite Track | Design and implement automated test suite across Tiers 1-4 (`tests/test_e2e_licensing.py`) | M1 & M2 contracts | DONE |
| M3 | Final Verification & Adversarial Coverage Hardening | Run 100% test suite pass (Tiers 1-4) and Tier 5 adversarial stress testing | M1, M2, E2E | DONE |
| M4 | Git Commit & Remote Push | Git commit and push to `luv-doc-tran` and `titanpdf` GitHub repositories | M3 | IN_PROGRESS |

## Interface Contracts
### Client Activation Contract (`core/license_client.py` ↔ Cloudflare Worker)
- **Endpoint**: `POST https://luv-license-api.luv-sds.workers.dev/v1/activate`
- **Headers**:
  - `Content-Type: application/json`
  - `User-Agent: LUV-DOC/1.1.0` (or `LUV-SDS/1.1.0` during fallback)
- **Payload**:
  ```json
  {
    "license_key": "LUV-XXXXX-XXXXX-XXXXX-XXXXX",
    "device_id": "<64_char_lowercase_hex_sha256>",
    "device_label": "<Machine_Name>",
    "app_version": "1.1.0",
    "product_code": "DOC"
  }
  ```
- **Success Response (HTTP 200)**:
  ```json
  {
    "ok": true,
    "token": "<base64url_payload>.<base64url_ed25519_signature>",
    "expires_at": "<ISO8601>",
    "offline_until": "<ISO8601>",
    "features": ["..."],
    "plan": "enterprise"
  }
  ```
- **Error Response (HTTP 400 / 403 / 404 / 409)**:
  ```json
  {
    "error": "invalid_request | license_denied | product_denied | device_limit",
    "message": "<Vietnamese description>"
  }
  ```

### Admin Management Contract (TitanPDF ↔ Cloudflare Worker)
- **Base URL**: `https://luv-license-api.luv-sds.workers.dev/v1/admin`
- **Headers**:
  - `Authorization: Bearer rHHfK9eC_8lTH02SP0CQPnyXAqMMDUXxwV2MiPothM7Avc4TxNZgnkKp5CFsIBP_`
  - `User-Agent: LUV-License-Admin/1.0`
- **Endpoints**:
  - `GET /licenses?q=...&status=...`: Returns `{"ok": true, "licenses": [{"id": "<UUID>", "license_key_prefix": "LUV-XXXXX", "customer_name": "...", "plan": "...", "status": "active", "active_devices": N, "max_devices": M, ...}]}`
  - `POST /licenses`: Creates new key in D1.
  - `POST /licenses/:id/revoke`: Revokes license by UUID `:id`.
  - `POST /licenses/:id/reset-devices`: Unbinds all active devices by UUID `:id`.

### Code Layout
- `c:\Users\ntc27\OneDrive\文件\LUV-DOC-TRAN`:
  - `core/license_client.py`: Authentic Ed25519 licensing client
  - `server.py`: FastAPI server routes and error handling
  - `requirements.txt`: Python package requirements (must include `cryptography>=42.0.0`)
  - `render.yaml`: Render deployment specification
  - `static/index.html`: Web interface and license modal
  - `data/license.sec`: Secure license token storage
  - `tests/test_e2e_licensing.py`: Comprehensive test suite (35 tests)
- `c:\Users\ntc27\OneDrive\文件\TitanPDF`:
  - `server.py`: FastAPI server routes for TitanPDF
  - `core/license_manager.py`: Cloudflare Admin API client
  - `static/license_manager.js`: Admin frontend logic
  - `static/license_manager.css`: Admin frontend styles
