# Feature Flag Audit — Spice e-Trade

A wiring check of every `flag_*` setting: is it actually connected to working
functionality, or just a dead toggle in the Settings → Flags screen?

Pair with [FEATURES.md](FEATURES.md) and [ARCHITECTURE.md](ARCHITECTURE.md).

**Last audited:** 2026-06-22

## How flags work here

There are **two distinct flag systems**, both keyed `flag_*` in company settings:

1. **UI feature flags** — gated declaratively. `applyFeatureFlags()`
   (`public/index.html`) reads the settings and flips `body[data-feat-*]`
   attributes; CSS rules (`public/index.html`, the `body:not([data-feat-...])`
   block) show/hide whole surfaces from there. No per-call-site `if (flag)`.
2. **Business / calculation flags** — read directly in backend logic
   (`calculations.js`, `invoice-pdf.js`, `tally-xml.js`, `exports*.js`,
   `server.js`).

---

## ✅ UI feature flags — all 8 fully wired

flag → gated control (`.feat-*`) → frontend handler → backend route

| Flag | Gates | Handler | Backend |
|---|---|---|---|
| `flag_whatsapp` | WhatsApp buttons (purchase/payment bulk + per-row, mobile app) | `whatsappSelectedPurchases()`, `whatsappSelectedPayments()` | `/api/whatsapp/*` |
| `flag_bills` | "Bills of Supply" sidebar tab + delete-all | `go('bills')`, `deleteAll('bills')` | bills routes |
| `flag_debit_notes` | "Debit Notes" sidebar tab + delete-all | `go('debit')`, `deleteAll('debit-notes')` | debit-note routes |
| `flag_price_list_mapping` | sidebar tab + Lots toolbar button | `go('plmap')` | plmap routes |
| `flag_discount_in_prate` | Lots → "🏷 Set Grade 1" bulk button | `lotBulkSetGrade1()` | recalc route |
| `flag_print_selected_purchase` | Sales → "🖨 Print Selected Purchase" (also requires Kerala/ASP) | `printSelectedPurchaseFromInvoices()` | purchase-mirror PDF |
| `flag_lot_set_buyer` | Lots → "👤 Set Buyer" bulk button | `lotBulkSetBuyer()` | lot-update route |
| `flag_price_check` | Price Check tab + content + `pcRefreshGate` button-locking | `pcRefreshGate()` | `/api/price-check/verify`, `/download`, `price-check-status` |

All handler functions and backend routes verified to exist.

---

## ✅ Business / calculation flags — wired to working logic

| Flag | Label | Wired in |
|---|---|---|
| `flag_round` | Round Invoice Amounts | `calculations.js`, `exports.js`, `exports-pdf.js`, `server.js` |
| `flag_disc_gst` | Discount includes GST | `calculations.js`, `server.js` (invoice + purchase), `db.js` |
| `flag_hsn` | Show HSN Codes | `invoice-pdf.js` |
| `flag_invoice_stripe` | Alternate Row Stripe in Invoice | `invoice-pdf.js` |
| `flag_dispatch` | Show Dispatch Address | `invoice-pdf.js` |
| `flag_ship` | Show Ship To Address | `invoice-pdf.js` |
| `flag_bank` | Bank Details in Invoice | `invoice-pdf.js` |
| `flag_sample` | Discount In Payments | `calculations.js`, `server.js` |
| `flag_wgst` | TDS on Full Invoice Amount | `calculations.js` |
| `flag_tds_purchase` | TDS on Purchase Invoice | `calculations.js`, `tally-xml.js` |
| `flag_inter_transport` | Inter-State Transport rate | `calculations.js` |
| `flag_inter_insurance` | Inter-State Insurance rate | `calculations.js` |
| `flag_local_transport` | Local Transport rate | `calculations.js` |
| `flag_local_insurance` | Local Insurance rate | `calculations.js` |

---

## ⚠️ NOT wired — cleanup targets

### 1. Five dead toggles

Defined in `company-config.js` and shown as switches in Settings → Flags, but
**never read by any code** (no static refs, no dynamic `'flag_'+x` access).
Flipping them does nothing — misleading to the end user.

| Flag | Label shown to user | Note |
|---|---|---|
| `flag_pooling` | "Pooling (Single State)" | — |
| `flag_tds_sales` | "TDS on Sales Invoice" | — |
| `flag_debit_note` | "Debit Note for Discount" | **singular** — distinct from the working `flag_debit_notes` |
| `flag_dummy` | "Allow Dummy Invoices" | dummy-invoice output exists (driven by `duplicate_text`), but this on/off flag gates nothing |
| `flag_export` | "Export Invoices" | — |

**Action:** either remove these from the settings schema, or wire the intended
logic where the meaning is clear (`flag_tds_sales`, `flag_dummy`).

### 2. Orphan read with no setting — `flag_tally_round`

Read in `tally-xml.js` (`cfgBool(cfg, 'flag_tally_round', true)`, controls the
Round On/Off ledger in ISP purchase XML) but **no `flag_tally_round` key is
defined** in `company-config.js` (only the `tally_round` text ledger name and
`tally_round_enabled` boolean exist). So it always falls back to `true` and is
effectively non-configurable.

**Action:** either expose `flag_tally_round` as a proper setting, or point the
read at the existing `tally_round_enabled`.

---

## ✅ Legacy flags — correctly retired (no action)

`flag_sister`, `flag_tnpa`, `flag_rtds_inv`, `flag_eway` are in `REMOVED_KEYS`
(`company-config.js`) and wiped at every boot. Not wired, by design.
