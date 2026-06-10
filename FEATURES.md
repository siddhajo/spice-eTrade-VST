# Spice e-Trade — Feature Inventory & Status

A living checklist of everything the app does and how finished it is.
Pair this with [ARCHITECTURE.md](ARCHITECTURE.md).

**Last reviewed:** 2026-06-09

**Status legend:** ✅ Complete & in use · 🟡 Partial / optional · ⚪ Deprecated/legacy · 🔧 Needs config to activate

---

## Core auction & trading

| Feature | Status | Notes |
|---------|:---:|-------|
| Trades (auctions) — create, edit, delete, import | ✅ | Multi-branch; cascade-delete of child lots/docs |
| Lot entry & management | ✅ | CRUD, bulk grade/seller/buyer, lock/unlock |
| Lot price auto-calculation | ✅ | Per-lot and whole-trade batch |
| Lot-number allocations per branch | ✅ | Reserves number ranges; prevents collisions between field staff |
| Sellers (traders) management | ✅ | CRUD, Excel import, multi-bank accounts, default bank |
| Buyers management | ✅ | CRUD, Excel import, GSTIN-based state detection |

## Documents & billing

| Feature | Status | Notes |
|---------|:---:|-------|
| Sales invoices (GST) | ✅ | Single, bulk, batch PDF; grouped by buyer; revert/undo |
| Purchase invoices | ✅ | Registered sellers; grouped by seller |
| Agri bills / bills of supply | ✅ | Unregistered agriculturists |
| Debit notes | ✅ | `generate-bulk` per trade; PDF; Tally-exportable |
| Invoice preview (dry run) | ✅ | See output before saving |
| Eligibility checks + debug | ✅ | Explains why a buyer/seller isn't eligible |
| Amount-in-words (Indian: lakh/crore) | ✅ | On invoices & receipts |
| Cross-trade `debit-notes/generate-all` | ⚪ | Replaced by per-trade `generate-bulk` |

## Payments

| Feature | Status | Notes |
|---------|:---:|-------|
| Seller payment summary | ✅ | Generated from purchases |
| Bank / NEFT payment data | ✅ | For bank upload files |
| TDS return summary | ✅ | CSV for filing |
| Payment PDFs (summary + lot receipts) | ✅ | Single, bulk |

## Reports & analytics

| Feature | Status | Notes |
|---------|:---:|-------|
| Dashboard stats & insights | ✅ | Counts, totals, min/max/avg price |
| Revenue trend over time | ✅ | |
| Per-trade summary (incl. PDF) | ✅ | |
| Branch comparison (TN vs KL) | ✅ | |
| Lorry / transport reports | ✅ | Lot slips, truck list, buyer-lot summary |
| Sales & purchase journals | ✅ | |

## Exports

| Feature | Status | Notes |
|---------|:---:|-------|
| Tally ERP XML | ✅ | Sales, RD/URD purchases, debit notes, ledgers; intra/inter-state aware |
| Excel reports (18+ types) | ✅ | Standardised branding, Indian number formatting |
| DBF (FoxPro) export | ✅ | CPA1/INV/PURCHASE/BILL/DEBIT/NAM/SBL for legacy systems |
| PDF exports | ✅ | Invoices, bills, debit notes, receipts, lorry slips |
| E-way bill distance tracking | ✅ | PIN→PIN cache + manual per-invoice overrides |

## Data integrity & migration

| Feature | Status | Notes |
|---------|:---:|-------|
| Price verification (price-check) | ✅ | Upload a price sheet, diff vs live, download annotated |
| Import old data (legacy XLS/DBF) | ✅ | Preview, dry-run, commit, **undo within 30 min** |
| Auto-backup scheduler | ✅ | Configurable interval; prunes old snapshots |
| Manual backup / restore | ✅ | Download / upload full DB |
| Bulk delete with undo window | ✅ | ~30-second reversible; logged for forensics |
| Audit log | ✅ | User action trail |

## Auth, users & config

| Feature | Status | Notes |
|---------|:---:|-------|
| Login / sessions (bcrypt + token) | ✅ | Legacy SHA-256 auto-upgraded on login |
| Role-based permissions (5 roles) | ✅ | viewer → lot_entry → operator → manager → admin |
| Forced password change | ✅ | Online mode |
| Login rate-limiting | ✅ | Online mode (10 / 15 min / IP) |
| User management | ✅ | CRUD, role change, admin password reset |
| Company settings (~109 keys) | ✅ | Identity, addresses, rates, feature flags |
| Logo upload (stored in DB) | ✅ | With bundled fallback |
| Settings import / export (JSON) | ✅ | Backup & clone a configuration |

## Platforms & deployment

| Feature | Status | Notes |
|---------|:---:|-------|
| Desktop app (Electron) | ✅ | macOS Intel + ARM, Windows x64; auto-update via GitHub |
| Cloud deploy (Railway / Docker) | ✅ | `Dockerfile`, `nixpacks.toml`, `Procfile` |
| Plain Node server | ✅ | `node server.js`, port 3001 |
| Mobile PWA (field lot entry) | ✅ | `public-mobile/app.html` via `mobile-bridge.js` |
| Client profiles (desktop vs online) | ✅ | Security posture switch via `CLIENT` env var |

## Integrations (optional)

| Feature | Status | Notes |
|---------|:---:|-------|
| Licensing (sign / renew tokens) | ✅ | HMAC-SHA256; per-install expiry |
| GST lookup from GSTIN | 🔧 | Needs an API key in settings |
| WhatsApp (Meta Cloud API) | 🔧 | Needs credentials; falls back to manual copy-paste if unconfigured |

---

## Known limitations / notes for maintainers

- **Single-company build.** Old dual-company (ISP/ASP) plumbing remains only for
  backward-compatible reports; it is not a multi-tenant SaaS.
- **Licensing is honest-user DRM.** It deters forgetting to renew, not a determined attacker
  with source/shell access. (By design.)
- **HTTPS is the host's job.** The app does bcrypt + rate-limiting; TLS must be terminated by
  the platform (Railway, reverse proxy, etc.).
- **Stale comment in `mobile-bridge.js`** mentions some routes are "stubbed (Pass 2)"; those
  routes (receipt PDF, batch print) are in fact implemented. Clean up the comment when next in
  that file.
