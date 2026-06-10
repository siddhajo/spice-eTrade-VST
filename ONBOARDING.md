# Onboarding a New Customer

> A step-by-step playbook for setting up Spice e-Trade for a brand-new company.
> Read [ARCHITECTURE.md](ARCHITECTURE.md) first if you want the "why" behind any step.

**Audience:** you (the developer/installer). **Time:** ~30–60 min for a clean setup.

---

## Overview — the 5 stages

```
  1. DEPLOY         →  2. FIRST LOGIN    →  3. CONFIGURE      →  4. LOAD DATA     →  5. LICENSE
  pick a platform      change the          company identity,     sellers, buyers,    issue a
  & start the app      default password    rates, branches       past records        renewal token
```

Each stage is below. You can do stage 4 (data) and stage 5 (license) in either order.

---

## Stage 1 — Deploy the app

Pick the platform that fits the customer:

### Option A — Desktop app (most common for a single shop)

1. Build the installer (or reuse a release):
   ```bash
   npm run build:win      # Windows x64 installer
   npm run build:mac      # macOS DMG (Intel + ARM)
   ```
2. Send the installer; the customer runs it like any normal app.
3. Data is stored automatically per-user:
   - **Windows:** `%APPDATA%\Spice e-Trade`
   - **macOS:** `~/Library/Application Support/Spice e-Trade`

   Nothing else to configure — it binds to `localhost` only and runs offline.

### Option B — Cloud / server (multi-user or remote access)

1. Deploy to Railway (or any Node host) using the included `Dockerfile` / `nixpacks.toml`.
2. Set these environment variables (see [.env.example](.env.example)):

   | Variable | Why | Example |
   |----------|-----|---------|
   | `CLIENT` | Use the hardened profile | `spice-online` |
   | `SPICE_DATA_DIR` | **Persistent** path so `config.db` survives restarts | `/data` |
   | `PORT` | Usually injected by the platform | `3001` |
   | `LICENSE_SECRET` | 64-char hex — **must match your signing laptop** (see Stage 5) | `a1b2…` |

   Generate the license secret once:
   ```bash
   node -e "console.log(require('crypto').randomBytes(32).toString('hex'))"
   ```
3. ⚠️ **Make sure `SPICE_DATA_DIR` points at a mounted volume.** Without it, the database
   resets on every redeploy.

> On first start the app automatically creates the database, seeds default settings, and
> starts a 30-day trial license. No manual DB setup needed.

---

## Stage 2 — First login & password change

1. Open the app. The seeded admin account is:

   | Username | Password |
   |----------|----------|
   | `admin` | `admin123` |

2. The app **forces a password change** before anything else works (online mode). Set a real
   password. Done.

> 🔒 Always change `admin123` immediately. On a public/cloud deploy this is non-negotiable.

---

## Stage 3 — Configure the company

All of this is in the **Settings** screen. Work top to bottom:

| Step | Setting area | What to fill in |
|------|--------------|-----------------|
| 3.1 | **Company details** | Trade name, short name, legal name, PAN, CIN / partnership name |
| 3.2 | **Logo** | Upload the company logo (stored in the DB; appears on every PDF) |
| 3.3 | **Business state** | `TAMIL NADU` (ISP mode) or `KERALA` (ASP mode) — ⚠️ this drives the whole calculation engine, set it correctly |
| 3.4 | **Addresses** | Address lines, phone, email, **GSTIN** for the relevant state |
| 3.5 | **Branches** | Branch names (br1–br9) and their phone numbers — these populate lot-entry dropdowns |
| 3.6 | **Rates & charges** | Commission %, handling %, deductions, refund kg, gunny rate, GST goods (5%), GST service (18%) |
| 3.7 | **Feature flags** | Turn on only what this customer uses — dispatch, shipping, bank details on invoices, pooling, price-check, debit notes, WhatsApp, etc. |
| 3.8 | **Tally export** | Ledger names, item names, HSN codes, state codes — **must match the customer's Tally company exactly** or imports into Tally will fail |

> 💡 **Cloning a setup?** If you've configured a similar customer before, use
> **Settings → Export** on the old install and **Settings → Import** on the new one, then just
> change the identity fields (name, GSTIN, PAN, addresses). Big time-saver.

### Optional integrations (only if the customer wants them)

- **GST lookup** — paste an API key in settings to auto-fill buyer details from a GSTIN.
- **WhatsApp** — enter Meta Cloud API credentials (token, phone ID, templates) to send
  invoices/receipts. If left blank, the app cleanly falls back to manual sharing.

---

## Stage 4 — Load the customer's data

You have three ways to get sellers, buyers, and history in:

### 4a. Excel import (clean start)

1. Download the templates: **Sellers → Download template**, **Buyers → Download template**.
2. Fill them in (column headers are auto-detected, so order is flexible).
3. Upload via **Sellers → Import** and **Buyers → Import**.

### 4b. Import old data (migrating from a previous system)

Use **Import Old Data** for past trades/invoices/payments from legacy XLS/DBF files:

1. **Preview** — uploads the file and shows detected columns + first rows.
2. **Dry run / verify** — checks for problems without writing.
3. **Run** — commits the import.
4. **Undo** is available for ~30 minutes if something looks wrong.

Supported: sellers, buyers, sales, purchases, bills, debit notes, payments.

### 4c. DBF bridge (coming straight from the old FoxPro app)

The app reads/writes the legacy DBF tables (CPA1, NAM, SBL, INV, PURCHASE, BILL, DEBIT) so an
existing FoxPro installation can hand its data over. See [dbf-exports.js](dbf-exports.js).

> ✅ After importing, spot-check a few records and run a **backup** (System → Backup now) so you
> have a clean restore point.

---

## Stage 5 — Issue the license

Every install starts with a **30-day trial**. To extend it you sign a token bound to that
install's ID.

### One-time setup (per deployment)

Make sure your signing laptop has the **same** `LICENSE_SECRET` you set on the server:
```bash
export LICENSE_SECRET="<the exact 64-char hex from Stage 1>"
```

### Each time you issue/renew

1. **Get the install ID** from the customer. They can find it on the `/renew.html` page (a
   **Copy** button is provided), or you can fetch it:
   ```bash
   curl https://<deployment>/api/license/status
   ```
2. **Mint a token** on your laptop:
   ```bash
   node tools/license-sign.js --install-id <THEIR-INSTALL-ID> --days 365 \
     --note "Acme Co — initial setup"
   ```
   (Use `--days 30` for a monthly cadence, or `--until 2027-03-31` for a fixed date.)
3. **Send the token** (one line, no spaces) via WhatsApp/email.
4. **Customer applies it** at `https://<deployment>/renew.html` → paste → **Apply renewal
   token** → sees a green "New expiry" confirmation. Takes effect immediately, no restart.

**Good to know:**
- A token only works for the install ID it was minted for.
- A token that would *shorten* the current license is rejected (so stale tokens are harmless).
- When expired, login returns HTTP 451 and the UI auto-redirects to the renewal page.

Full licensing reference, troubleshooting, and the emergency `admin/set-expiry` endpoint are in
[LICENSING.md](LICENSING.md).

---

## Stage 6 — Create the customer's users

Add the people who'll actually use the app, each with the least powerful role they need:

| Role | Give it to… |
|------|-------------|
| `lot_entry` | Field staff entering lots (often via the **mobile PWA**) |
| `operator` | Daily desk work — lots, invoices, traders, buyers |
| `manager` | Branch supervisor — also trades, reverting invoices, settings |
| `admin` | The owner / you — full control incl. user management & deletes |

Each new user is forced to set their own password on first login.

### Mobile field entry (optional)

Field staff can open `https://<deployment>/mobile/` on a phone/tablet to enter lots and print
receipts on the spot. Give them a `lot_entry` account. (Desktop installs are localhost-only, so
mobile is mainly for the cloud deployment.)

---

## Quick checklist

Copy this into your handover notes:

```
□ App deployed (desktop installer / cloud env vars set)
□ SPICE_DATA_DIR points at a persistent volume   (cloud only)
□ LICENSE_SECRET set on server AND signing laptop (cloud only)
□ admin123 password changed
□ Company identity: name, PAN, CIN/partnership, logo
□ Business state set correctly (TN/ISP or KL/ASP)
□ Addresses + GSTIN entered
□ Branches + phone numbers
□ Rates & charges (commission, GST, gunny, deductions)
□ Feature flags set to what the customer uses
□ Tally export ledger/item names match their Tally company
□ Sellers & buyers imported
□ Old data migrated (if applicable) + verified
□ Backup taken
□ License token issued (>30 days)
□ User accounts created with correct roles
□ Mobile lot entry tested (if used)
```

---

## Troubleshooting first-setup issues

| Symptom | Likely cause | Fix |
|---------|-------------|-----|
| Data disappears after cloud redeploy | `SPICE_DATA_DIR` not on a persistent volume | Point it at a mounted volume and restore from backup |
| "Token was issued for a different install" | Token minted for the wrong install ID | Re-mint with the correct ID from `/api/license/status` |
| License token rejected as "bad signature" | `LICENSE_SECRET` differs between laptop and server | Make both identical, re-mint |
| Tally import fails on the customer's side | Ledger/item/HSN names don't match their Tally company | Fix the names in **Settings → Tally export** |
| Wrong tax/price calculations | `business_state` set to the wrong mode | Set TN (ISP) or KL (ASP) correctly; recalculate lots |
| New input box shows no cursor on Windows | Known focus-paint quirk with `.fg` inputs | Add scoped CSS for that input (see ARCHITECTURE §11) |
```
