# Spice e-Trade — Architecture Guide

> **What this app is, in one line:** a complete admin system for running spice/cardamom
> auctions — entering lots, generating GST invoices and bills, paying sellers, and
> exporting everything to Tally and Excel.

This document explains how the app is put together so a developer (or future you) can
understand, maintain, and extend it. For *setting up a new customer*, see
[ONBOARDING.md](ONBOARDING.md).

**Last reviewed:** 2026-06-09

---

## 1. The Big Picture

Spice e-Trade is a single **Node.js + Express** backend that serves a browser UI. The
same backend runs in three different "skins":

```
                       ┌─────────────────────────────────────┐
                       │          server.js  (Express)        │
                       │   ~190 API routes · all the logic    │
                       └──────────────────┬───────────────────┘
                                          │
              ┌───────────────────────────┼───────────────────────────┐
              │                           │                           │
     ┌────────▼────────┐         ┌────────▼────────┐         ┌────────▼────────┐
     │  Desktop App     │         │  Cloud / Server  │         │  Mobile PWA      │
     │  (Electron)      │         │  (Railway/Docker)│         │  (phone/tablet)  │
     │                  │         │                  │         │                  │
     │ runs server.js   │         │ public 0.0.0.0   │         │ /mobile + bridge │
     │ in-process,      │         │ rate-limited,    │         │ lot entry in the │
     │ localhost only   │         │ 30-day sessions  │         │ field            │
     └──────────────────┘         └──────────────────┘         └──────────────────┘
                                          │
                                ┌─────────▼─────────┐
                                │   SQLite database  │
                                │   data/config.db   │
                                └────────────────────┘
```

**Key idea:** one codebase, one database engine, and a small **client profile** flag
decides how strict/locked-down the deployment is. There is no separate "desktop version"
of the code — it's the same `server.js`.

---

## 2. Technology Stack

| Layer | Technology | Notes |
|-------|-----------|-------|
| Backend | Node.js + Express | All logic lives in `server.js` (~11k lines) plus helper modules |
| Database | SQLite via **sql.js** | Pure-JavaScript SQLite (no native compile). Optional `better-sqlite3` for speed |
| Frontend | Single-page app | `public/index.html` (admin UI), `public-mobile/app.html` (mobile PWA) |
| PDF generation | PDFKit | Invoices, bills, receipts, reports |
| Excel/CSV | ExcelJS + SheetJS (`xlsx`) | 18+ export formats |
| Desktop wrapper | Electron + electron-builder | macOS (Intel/ARM) and Windows installers |
| Auth | bcryptjs | Password hashing; token sessions (not JWT) |
| Deployment | Docker / Nixpacks / Railway | Plus the Electron installer for desktop |

### Why sql.js instead of a "real" database?

sql.js holds the whole database in memory and writes the full file on changes (debounced
~200 ms). This avoids native-module compilation headaches on cloud platforms and inside
Electron on Windows. For higher write throughput you can install the optional
`better-sqlite3` — both expose the **same `db.js` API**, so no other code changes.
See [MIGRATION.md](MIGRATION.md).

---

## 3. Code Map — What Lives Where

| File | Size | Responsibility |
|------|------|----------------|
| [server.js](server.js) | ~540 KB | The Express app — every API route, auth, permissions, bootstrap |
| [db.js](db.js) | ~45 KB | SQLite init, **all table schemas**, migrations, save debouncing |
| [company-config.js](company-config.js) | ~47 KB | ~109 company settings, defaults, feature flags, getters |
| [calculations.js](calculations.js) | ~45 KB | Lot pricing, deductions, GST/tax math (ISP vs ASP modes) |
| [invoice-pdf.js](invoice-pdf.js) | ~108 KB | PDF for invoices, bills, debit notes, receipts |
| [tally-xml.js](tally-xml.js) | ~153 KB | Tally ERP XML export (sales, purchases, debit notes, ledgers) |
| [exports.js](exports.js) / [exports-pdf.js](exports-pdf.js) | ~57 / 42 KB | Excel + PDF report builders |
| [auction-reports.js](auction-reports.js) | ~68 KB | Summary reports, lot listings |
| [lorry-reports.js](lorry-reports.js) | ~42 KB | Transport/freight reports |
| [dbf-exports.js](dbf-exports.js) | ~17 KB | Legacy FoxPro DBF import/export |
| [mobile-bridge.js](mobile-bridge.js) | ~65 KB | Translates the mobile PWA's API calls to backend routes |
| [price-check.js](price-check.js) | ~25 KB | Verify uploaded price sheets against live data |
| [license.js](license.js) | ~9 KB | Per-install licensing (sign/verify tokens, expiry) |
| [client-profile.js](client-profile.js) | ~4 KB | Loads the active client profile (desktop vs online) |
| [distance.js](distance.js) | ~6 KB | PIN-to-PIN distance cache for e-way bills |
| [amount-words.js](amount-words.js) / [date-format.js](date-format.js) | small | Rupees-in-words, configurable date formatting |
| [electron/main.js](electron/main.js) | ~8 KB | Desktop wrapper — boots server, window, menus, auto-update |

---

## 4. The Data Model

The database (`data/config.db`) is plain SQLite. Tables fall into a few groups.

### Business entities (the heart of the app)

| Table | Holds | Legacy source |
|-------|-------|---------------|
| `auctions` | A **trade** — one auction event (date, crop, branch/state) | — |
| `lots` | Individual lots in a trade (seller, buyer, weight, grade, price) | CPA1.DBF |
| `traders` | **Sellers / poolers** (name, PAN, bank details) | NAM.DBF |
| `trader_banks` | Multiple bank accounts per seller, with a default | — |
| `buyers` | **Buyers / dealers** (GSTIN, address, state) | SBL.DBF |

### Documents (generated outputs)

| Table | Holds |
|-------|-------|
| `invoices` | Sales invoices (buyer-facing, with GST) |
| `purchases` | Purchase invoices (registered seller-facing) |
| `bills` | Agri bills / bills of supply (unregistered agriculturists) |
| `debit_notes` | Adjustment / discount notes against purchases |

### Configuration & operations

| Table | Holds |
|-------|-------|
| `company_settings` | Key-value config — the company's whole identity & rules |
| `company_logos` | Uploaded logos (stored as BLOBs) |
| `users` / `sessions` | Login accounts and active login tokens |
| `license_state` | Install ID and license expiry (one row) |
| `lot_allocations` | Reserved lot-number ranges per branch (prevents collisions) |
| `route_distances` | Cached PIN→PIN road distances for e-way bills |
| `whatsapp_config` / `whatsapp_messages` | Optional WhatsApp integration |
| `audit_log` / `delete_log` / `import_log` | Audit trail, wipe forensics, import history (with undo) |

**Relationships:** `lots → auctions`, `lots → traders`, `trader_banks → traders`,
`sessions → users`.

> 💡 The schema is the single source of truth in [db.js](db.js). Tables are created with
> `CREATE TABLE IF NOT EXISTS`, and schema upgrades for existing databases live in the
> migrations block further down the same file.

---

## 5. How a Trade Flows Through the System

This is the day-to-day workflow the app is built around:

```
1. Create a Trade        →  POST /api/auctions          (date + crop + branch)
2. Enter Lots            →  POST /api/lots              (seller, buyer, weight, grade)
        │                    (or via the Mobile PWA in the field)
        ▼
3. Calculate prices      →  POST /api/lots/calculate-all
        ▼
4. (optional) Price-check →  POST /api/price-check/verify
        ▼
5. Generate documents:
     • Sales invoices     →  POST /api/invoices/generate-all/:auctionId   (grouped by buyer)
     • Purchase invoices  →  POST /api/purchases/generate-all/:auctionId  (grouped by seller)
     • Agri bills         →  POST /api/bills/generate-all/:auctionId
     • Debit notes        →  POST /api/debit-notes/generate-bulk
        ▼
6. Pay sellers           →  GET  /api/payments/:auctionId  (+ bank/NEFT summary)
        ▼
7. Export everything:
     • Tally XML          →  GET /api/tally/export/:type/:auctionId
     • Excel reports      →  GET /api/exports/:type/:auctionId
     • PDFs               →  per-document /pdf endpoints
```

Lots can be **locked** after invoices are generated so they can't be accidentally edited,
and invoices can be **reverted** (which unlocks the lots again) by a manager/admin.

### ISP vs ASP / TN vs Kerala

The `business_state` setting (`TAMIL NADU` or `KERALA`) switches the calculation engine
between two modes (historically "ISP" and "ASP"). This affects how seller-side quantities,
rates, and taxes are computed. This is a **single-company** build — the dual-company logic
remains only for backward-compatible reports.

---

## 6. API Surface (by area)

There are ~190 routes. They group like this:

| Area | Representative endpoints |
|------|--------------------------|
| **Auth & users** | `POST /api/login`, `GET /api/me`, `POST /api/users`, `PUT /api/users/:id/role` |
| **Company config** | `GET/PUT /api/company-settings`, `GET /api/branding`, logo upload |
| **Sellers (traders)** | `GET/POST/PUT/DELETE /api/traders`, `/import`, `/template`, bank defaults |
| **Buyers** | `GET/POST/PUT/DELETE /api/buyers`, `/import`, GST lookup |
| **Trades (auctions)** | `GET/POST/PUT/DELETE /api/auctions`, allocations, lot validation |
| **Lots** | CRUD, `calculate-all`, `lock`/`unlock`, bulk grade/seller/buyer |
| **Invoices** | `generate(-all)`, `eligible-buyers`, `/pdf`, `revert`, `preview` |
| **Purchases / Bills** | parallel generate / eligible / pdf / edit routes |
| **Debit notes** | `generate-bulk`, `next-note-no`, `/pdf` |
| **Payments** | `/payments/:auctionId`, `/bank`, payment PDFs, TDS return |
| **Exports** | `/exports/:type`, `/dbf-exports/:type`, `/tally/...` |
| **Reports** | `/stats`, `/insights`, `/reports/trade-summary`, lorry reports |
| **System** | backups, restore, import-old-data (with undo), license |
| **Integrations** | WhatsApp (Meta Cloud API), GST lookup |

For the exhaustive list, search `app.get(` / `app.post(` in [server.js](server.js).

---

## 7. Authentication & Permissions

**Sessions, not JWT.** Login verifies the password (bcrypt, cost 12) and creates a row in
`sessions` with a random 64-char token. The client sends it as `Authorization: Bearer <token>`.
Every protected route runs `requireAuth` then a `requirePermission('...')` check.

**Roles** (each inherits the powers of the one above it):

| Role | Can do |
|------|--------|
| `viewer` | Read-only |
| `lot_entry` | Create/edit lots and sellers inline (field worker / mobile) |
| `operator` | + invoices, traders, buyers — daily auction work |
| `manager` | + create trades, revert invoices, edit settings, toggle state |
| `admin` | + delete records, wipe tables, manage users, licensing |

**Safety features baked in:**

- **Forced password change** — new/reset users must change the default password before they
  can do anything (online mode).
- **Login rate-limit** — 10 attempts per 15 min per IP (online mode only).
- **Soft deletes + undo windows** — bulk "delete all" operations are logged and reversible
  for ~30 seconds; imports are reversible too.
- **Auto-backup** — a scheduler snapshots the DB on a configurable interval and prunes old
  copies.

---

## 8. Deployment Modes (Client Profiles)

The `CLIENT` environment variable picks a profile from `clients/<name>/profile.json`,
which flips security defaults:

| Behaviour | `spice-etrade-desktop` | `spice-online` |
|-----------|:---:|:---:|
| Network bind | `127.0.0.1` (local only) | `0.0.0.0` (public) |
| Login rate-limit | off | **on** |
| Session expiry | never | **30 days** |
| Forced password change | off | **on** |
| Upload body limit | 50 MB | 2 MB |

```bash
CLIENT=spice-etrade-desktop npm start   # desktop build
CLIENT=spice-online npm start           # cloud / server
npm start                               # defaults to the safe "online" profile
```

A profile may also load `clients/<name>/overrides/index.js` to add extra routes without
touching the shared code. See [clients/README.md](clients/README.md).

### Three ways to run it

1. **Desktop (Electron)** — `electron/main.js` boots `server.js` *in the same process* and
   opens a window at `localhost`. Data lives in the OS app-data folder. Auto-updates from
   GitHub Releases. Build with `npm run build:win` / `npm run build:mac`.
2. **Cloud (Railway/Docker)** — `Dockerfile` / `nixpacks.toml` / `Procfile`. Set
   `SPICE_DATA_DIR` to a persistent volume so `config.db` survives restarts.
3. **Plain server** — `node server.js` on any Node host (port 3001 by default).

---

## 9. Licensing (Honest-User DRM)

Each install generates a UUID **install ID** and starts a 30-day trial. To extend it, the
developer signs a token (bound to that install ID) and the customer pastes it into
`/renew.html`.

- **Crypto:** HMAC-SHA256 signed tokens (`license.js`, signed with `tools/license-sign.js`).
- **What it controls:** expiry date only — there are no per-feature or per-seat limits.
- **Enforcement:** the login route returns HTTP **451** once expired and the UI redirects to
  the renewal page.
- **Threat model:** this stops honest customers from forgetting to renew. It does **not**
  defend against someone with source/shell access — and that's an accepted trade-off.

Full operational detail (commands, troubleshooting, the issue/renew flow) lives in
[LICENSING.md](LICENSING.md) and in the onboarding guide.

---

## 10. Integrations & Legacy Bridges

| Integration | What it does | Status |
|-------------|--------------|--------|
| **Tally ERP** | Exports sales/purchase/debit-note/ledger vouchers as Tally XML | ✅ Working |
| **DBF (FoxPro)** | Reads/writes the old CPA1/NAM/SBL/INV/etc. DBF files for migration | ✅ Working |
| **GST lookup** | Auto-fills buyer trade name/address from a GSTIN | ✅ Working (needs API key) |
| **WhatsApp** | Sends invoices/receipts via Meta Cloud API (template messages) | ✅ Working, optional |
| **Import old data** | Migrates legacy XLS/DBF into the DB, with dry-run and undo | ✅ Working |

---

## 11. Where to Start When Making a Change

- **Adding/altering a table or column?** → [db.js](db.js) (schema + migrations block).
- **New API endpoint?** → [server.js](server.js); copy an existing route's
  `requireAuth` + `requirePermission` pattern.
- **Changing a calculation (price, GST, deductions)?** → [calculations.js](calculations.js).
- **Changing a company setting / feature flag?** → [company-config.js](company-config.js).
- **Tweaking a PDF?** → [invoice-pdf.js](invoice-pdf.js) or [exports-pdf.js](exports-pdf.js).
- **Tweaking a Tally export?** → [tally-xml.js](tally-xml.js).
- **Adding a mobile feature?** → wire it in [mobile-bridge.js](mobile-bridge.js).

> ⚠️ **Windows desktop gotcha:** new `<input>` fields with the `.fg` class can render with no
> caret/border on Windows even though they're focusable. Add scoped CSS protection for any
> new input. (Recurring perception bug.)

---

## 12. Feature Status — see the companion table

The full, up-to-date feature inventory with status flags lives in
[FEATURES.md](FEATURES.md). Keep that file in sync whenever a feature is added, completed,
or deprecated.
