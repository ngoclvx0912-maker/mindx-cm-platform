# Access Control 3-tier Refactoring

## Date: 2026-03-29

## Summary
Replaced single password modal (MindX@123) with 3-tier access control system.

## Changes

### index.html
- Replaced `#pw-modal` with `#auth-modal` — supports both text (MSNV) and password input modes
- Added `auth-modal-footer` for future use
- Updated cache busting: `?v=fix4` → `?v=ac3`

### style.css
- Added `.auth-modal-footer` styles
- Added `.header-user-badge` — shows logged-in user in header
- Added `.topbar-select.bu-locked` — visual lock state for BU selectors

### app.js (main changes)
- **Removed**: `ACCESS_PROTECTED_VIEWS`, `accessGranted`, `ENCODED_PW`, `checkPassword()`, `showPasswordModal()`, `closePasswordModal()`, `submitPassword()`
- **Added access tiers**:
  - `ACCESS_CM_VIEWS = ['cm', 'daily', 'discussion']` — requires MSNV login
  - `ACCESS_FM_VIEWS = ['dashboard']` — requires FM password (MindX@2026)
  - `ACCESS_ADMIN_VIEWS = ['config']` — requires Admin password (Admin@123)
- **Added CM_STAFF_MAP**: Static MSNV→BU mapping for 43 CMs + STLs
- **Added dynamic staff loading**: `loadStaffMap()` fetches from `get_staff_list` API
- **Added `lookupStaff()`**: Checks dynamic map first, then static fallback
- **Added session management**: `restoreSession()` + `sessionStorage` persistence
- **Added `applyCMLogin()`**: Auto-selects BU, locks selector, updates header badge
- **Added `handleLogout()`**: Clears all sessions, unlocks selectors
- **Modified `navigate()`**: 3-tier check instead of single password
- **Modified `init()`**: Calls `restoreSession()` first, applies CM login, loads staff map
- **Legacy compatibility**: `accessGranted`, `showPasswordModal()`, `closePasswordModal()`, `submitPassword()` kept as no-ops

### google-apps-script.js
- Added `get_staff_list` action in `doGet()`
- Added `STAFF_BU_MAPPING` object (raw BU + region → app BU name)
- Added `resolveStaffBU()` function with exact + fuzzy matching
- Added `handleGetStaffList()` — reads staff sheet (gid=1896503199), filters Active, returns JSON

## Access Flow
1. **CM**: Click CM/Daily/Discussion → MSNV modal → validate → auto-select BU → lock selector
2. **FM**: Click FM Dashboard → password modal → MindX@2026 → full dashboard access
3. **Admin**: Click Cấu hình → password modal → Admin@123 → config access
4. All sessions stored in sessionStorage (cleared on tab close or manual logout)
