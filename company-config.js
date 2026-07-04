/**
 * company-config.js — Replaces TOOL.DBF + DEPOTS.PRG + COMPANY.PRG
 * All company configuration stored as key-value pairs in SQLite.
 */

const DEFAULTS = [
  // ── COMPANY ────────────────────────────────────────────────
  { key: 'logo',            value: '',           category: 'company',   label: 'Logo Code',                type: 'text' },
  { key: 'trade_name',      value: '',  category: 'company',   label: 'Trade Name',               type: 'text' },
  { key: 'legal_name',      value: '', category: 'company', label: 'Legal Name Suffix',        type: 'text' },
  { key: 'short_name',      value: '', category: 'company', label: 'Short Name', type: 'text' },
  // Nickname — short identifier used as the company code on receipts and
  // in the Praman e-Trade CSV column 2. Falls back to Short Name when
  // blank (handled by exports.js / receipt builders). Was previously
  // `praman_company` under Integrations / To Tally; renamed because it
  // now drives more than one surface.
  { key: 'praman_company',  value: '', category: 'company', label: 'Nickname (fallback: Short Name)', type: 'text' },
  { key: 'pan',             value: '',    category: 'company',   label: 'PAN',                      type: 'text' },
  // Partnership toggle: when true, every PDF that previously rendered
  // a "CIN" line switches to "Partnership" with the value from
  // `partnership_name` (typically the firm's registered partnership
  // deed number or partnership name). When false (default), the
  // existing CIN field continues to render as "CIN: <value>".
  // Stored as text "true"/"false" — same convention used by every
  // other boolean setting in the system.
  { key: 'is_partnership',  value: 'false', category: 'company', label: 'Partnership Firm',         type: 'boolean' },
  { key: 'partnership_name', value: '',    category: 'company', label: 'Partnership Name / No.',   type: 'text' },
  { key: 'cin',             value: '', category: 'company', label: 'CIN',                type: 'text' },
  { key: 'fssai',           value: '',               category: 'company',   label: 'FSSAI No.',               type: 'text' },
  { key: 'sbl',             value: '',               category: 'company',   label: 'SBL No.',                 type: 'text' },
  // MSME / Udyam registration number — single company-wide value.
  // Printed next to the company GSTIN on the Sales Invoice.
  { key: 'msme',            value: '',               category: 'company',   label: 'MSME / Udyam No.',        type: 'text' },

  // ── ADDRESS (Kerala) ───────────────────────────────────────
  { key: 'kl_address1',     value: '', category: 'address_kl', label: 'Address Line 1', type: 'text' },
  { key: 'kl_address2',     value: '', category: 'address_kl', label: 'Address Line 2', type: 'text' },
  { key: 'kl_phone',        value: '',    category: 'address_kl', label: 'Phone',                   type: 'text' },
  { key: 'kl_email',        value: '', category: 'address_kl', label: 'Email',       type: 'text' },
  { key: 'kl_gstin',        value: '', category: 'address_kl', label: 'GSTIN',                 type: 'text' },
  { key: 'kl_branch',       value: '', category: 'address_kl', label: 'Office Branch',           type: 'text' },

  // ── ADDRESS (Tamil Nadu) ───────────────────────────────────
  { key: 'tn_address1',     value: '', category: 'address_tn', label: 'Address Line 1', type: 'text' },
  { key: 'tn_address2',     value: '', category: 'address_tn', label: 'Address Line 2', type: 'text' },
  { key: 'tn_dispatch',     value: '', category: 'address_tn', label: 'Dispatch Address', type: 'text' },
  { key: 'tn_phone',        value: '',    category: 'address_tn', label: 'Phone',                   type: 'text' },
  { key: 'tn_email',        value: '', category: 'address_tn', label: 'Email',       type: 'text' },
  { key: 'tn_gstin',        value: '', category: 'address_tn', label: 'GSTIN',                 type: 'text' },
  { key: 'tn_branch',       value: '', category: 'address_tn', label: 'Office Branch',           type: 'text' },

  // ── BRANCHES ───────────────────────────────────────────────
  { key: 'br1',             value: '',               category: 'branches',  label: 'Branch 1',                type: 'text' },
  { key: 'br2',             value: '',               category: 'branches',  label: 'Branch 2',                type: 'text' },
  { key: 'br3',             value: '',               category: 'branches',  label: 'Branch 3',                type: 'text' },
  { key: 'br4',             value: '',               category: 'branches',  label: 'Branch 4',                type: 'text' },
  { key: 'br5',             value: '',               category: 'branches',  label: 'Branch 5',                type: 'text' },
  { key: 'br6',             value: '',               category: 'branches',  label: 'Branch 6',                type: 'text' },
  { key: 'br7',             value: '',               category: 'branches',  label: 'Branch 7',                type: 'text' },
  { key: 'br8',             value: '',               category: 'branches',  label: 'Branch 8',                type: 'text' },
  { key: 'br9',             value: '',               category: 'branches',  label: 'Branch 9',                type: 'text' },
  { key: 'br1_tel',         value: '',               category: 'branches',  label: 'Branch 1 Mobile',         type: 'text' },
  { key: 'br2_tel',         value: '',               category: 'branches',  label: 'Branch 2 Mobile',         type: 'text' },
  { key: 'br3_tel',         value: '',               category: 'branches',  label: 'Branch 3 Mobile',         type: 'text' },
  { key: 'br4_tel',         value: '',               category: 'branches',  label: 'Branch 4 Mobile',         type: 'text' },
  { key: 'br5_tel',         value: '',               category: 'branches',  label: 'Branch 5 Mobile',         type: 'text' },
  { key: 'br6_tel',         value: '',               category: 'branches',  label: 'Branch 6 Mobile',         type: 'text' },
  { key: 'br7_tel',         value: '',               category: 'branches',  label: 'Branch 7 Mobile',         type: 'text' },
  { key: 'br8_tel',         value: '',               category: 'branches',  label: 'Branch 8 Mobile',         type: 'text' },

  // ── RATES ──────────────────────────────────────────────────
  { key: 'commission',      value: '1',              category: 'rates',     label: 'Commission %',             type: 'number' },
  { key: 'hpc',             value: '10',             category: 'rates',     label: 'Handling %',               type: 'number' },
  { key: 'deduction1',      value: '1.25',           category: 'rates',     label: 'Deduction (Pooler)',       type: 'number' },
  { key: 'deduction2',      value: '1.25',           category: 'rates',     label: 'Deduction (Dealer)',       type: 'number' },
  // Used by flag_discount_in_prate — applies ONLY to Grade 1 lots.
  // When the flag is ON, Grade 1 P_Rate is computed against this
  // "discount-inclusive" deduction value instead of plain deduction1
  // above, and the per-lot Discount (refund) is forced to 0 because
  // it's already baked into the rate. Grade 2 and ungraded lots keep
  // the original deduction1/2 + separate discount behaviour. Stored
  // separately so toggling the flag back never loses the original
  // deduction1 percentage.
  { key: 'deduction1_inclusive', value: '1.25',      category: 'rates',     label: 'Deduction (Pooler) — discount-inclusive (Grade 1 only)', type: 'number' },
  { key: 'refund',          value: '1.9',            category: 'rates',     label: 'Sample Refund (Kgs)',      type: 'number' },
  { key: 'sb_refund',       value: '2.85',           category: 'rates',     label: 'SB Sample Refund (Kgs)',   type: 'number' },
  { key: 'gst_goods',       value: '5',              category: 'rates',     label: 'GST Goods Rate %',         type: 'number' },
  { key: 'gst_service',     value: '18',             category: 'rates',     label: 'GST Service Rate %',       type: 'number' },
  { key: 'discount_gst',    value: '5',              category: 'rates',     label: 'Discount GST %',           type: 'number' },
  { key: 'tcs_tds',         value: '0.1',            category: 'rates',     label: 'TCS / TDS Rate %',         type: 'number' },
  { key: 'tds_purchase_rate', value: '0.1',          category: 'rates',     label: 'TDS on Purchase Rate % (Section 194Q)',  type: 'number' },
  { key: 'tds_threshold',   value: '5000000',        category: 'rates',     label: 'TDS / TCS Annual Threshold (₹) — default ₹50 lakh per Section 194Q/206C(1H)',  type: 'number' },
  { key: 'gunny_rate',      value: '165',            category: 'rates',     label: 'Gunny Rate (₹)',           type: 'number' },
  // Inter-state transport / insurance — visible & applied only when
  // flag_inter_transport / flag_inter_insurance is ON.
  { key: 'flag_inter_transport', value: 'true',      category: 'rates',     label: 'Inter-State Transport (use inter-state transport rate)', type: 'boolean' },
  { key: 'transport',       value: '2.5',            category: 'rates',     label: 'Transport (₹/kg)',         type: 'number' },
  { key: 'flag_inter_insurance', value: 'true',      category: 'rates',     label: 'Inter-State Insurance (use inter-state insurance rate)', type: 'boolean' },
  { key: 'insurance',       value: '0.75',           category: 'rates',     label: 'Insurance (₹/kg)',         type: 'number' },
  // Local transport / insurance — visible & applied only when their
  // flag is ON. Migrated from the legacy `tally_local_transport` /
  // `tally_local_insurance` toggles that lived under To Tally and were
  // never wired to anything.
  { key: 'flag_local_transport', value: 'true',      category: 'rates',     label: 'Local Transport (use local transport rate)', type: 'boolean' },
  { key: 'local_transport', value: '2.5',            category: 'rates',     label: 'Local Transport (₹/kg)',   type: 'number' },
  { key: 'flag_local_insurance', value: 'true',      category: 'rates',     label: 'Local Insurance (use local insurance rate)', type: 'boolean' },
  { key: 'local_insurance', value: '0.75',           category: 'rates',     label: 'Local Insurance (₹/kg)',   type: 'number' },
  { key: 'discount_pct',    value: '0',              category: 'rates',     label: 'Discount %',               type: 'number' },
  { key: 'discount_days',   value: '0',              category: 'rates',     label: 'No. of Days for Discount', type: 'number' },
  { key: 'dealer_days',     value: '0',              category: 'rates',     label: 'No. of Days for Dealer',   type: 'number' },
  { key: 'addl_charge_name',  value: '',             category: 'rates',     label: 'Additional Charge — Name', type: 'text' },
  { key: 'addl_charge_value', value: '0',            category: 'rates',     label: 'Additional Charge — % of cardamom amount (0 to disable)', type: 'number' },

  // ── HSN / SAC CODES ────────────────────────────────────────
  { key: 'hsn_cardamom',    value: '09083120',       category: 'hsn',       label: 'Cardamom HSN',             type: 'text' },
  { key: 'hsn_gunny',       value: '63051040',       category: 'hsn',       label: 'Gunny HSN',                type: 'text' },
  { key: 'sac_transport',   value: '996791',         category: 'hsn',       label: 'Transport SAC',            type: 'text' },
  { key: 'sac_insurance',   value: '997136',         category: 'hsn',       label: 'Insurance SAC',            type: 'text' },
  { key: 'sac_service',     value: '996111',         category: 'hsn',       label: 'Service SAC',              type: 'text' },

  // ── BANK DETAILS ───────────────────────────────────────────
  { key: 'bank_kl_name',    value: '',               category: 'bank',      label: 'Kerala Bank Name',         type: 'text' },
  { key: 'bank_kl_acct',    value: '',               category: 'bank',      label: 'Kerala Account No.',       type: 'text' },
  { key: 'bank_kl_ifsc',    value: '',               category: 'bank',      label: 'Kerala IFSC Code',         type: 'text' },
  { key: 'bank_tn_name',    value: '',               category: 'bank',      label: 'TN Bank Name',             type: 'text' },
  { key: 'bank_tn_acct',    value: '',               category: 'bank',      label: 'TN Account No.',           type: 'text' },
  { key: 'bank_tn_ifsc',    value: '',               category: 'bank',      label: 'TN IFSC Code',             type: 'text' },

  // ── SEASON ─────────────────────────────────────────────────
  { key: 'season',          value: '',      category: 'season',    label: 'Season Name',              type: 'text' },
  { key: 'season_short',    value: '26-27',          category: 'season',    label: 'Season Short',             type: 'text' },
  { key: 'season_start',    value: '2026-04-01',     category: 'season',    label: 'FY Start Date',            type: 'date' },
  { key: 'season_end',      value: '2027-03-31',     category: 'season',    label: 'FY End Date',              type: 'date' },

  // ── INVOICE SETTINGS ───────────────────────────────────────
  { key: 'inv_prefix',      value: '',               category: 'invoice',   label: 'Invoice Prefix',           type: 'text' },
  { key: 'separator',       value: '-',              category: 'invoice',   label: 'Separator Symbol',         type: 'text' },
  // hsn_cardamom and hsn_gunny live in the HSN/SAC category — see above.
  // Earlier they were duplicated here; removed so editing one source of
  // truth (Settings → HSN/SAC) updates every consumer.
  { key: 'dispatched_through_isp', value: '',         category: 'invoice',   label: 'Dispatched Through', type: 'text' },
  { key: 'dispatch_destination', value: '',           category: 'invoice', label: 'Dispatch Destination',     type: 'text' },
  { key: 'duplicate_text',  value: 'DUMMY INVOICE',  category: 'invoice',   label: 'Dummy Invoice Text',       type: 'text' },
  { key: 'signature_text',  value: 'Signature of the Authorised Buyer', category: 'invoice', label: 'Signature Label', type: 'text' },

  // ── FEATURE FLAGS ──────────────────────────────────────────
  { key: 'flag_pooling',    value: 'false',          category: 'flags',     label: 'Pooling (Single State)',    type: 'boolean' },
  // Master switch for discount calculation. ON ⇒ the per-lot Discount
  // (lots.refund) and the auto-computed discount Debit Notes are calculated
  // and flow into the Payments tab. OFF ⇒ no discount is calculated anywhere
  // (full PurAmt flows to Payable). Read by calculateLot() and the DN
  // generators. Label shown in Settings → Feature Flags.
  { key: 'flag_sample',     value: 'false',          category: 'flags',     label: 'Discount In Payments',     type: 'boolean' },
  { key: 'flag_dispatch',   value: 'true',           category: 'flags',     label: 'Show Dispatch Address',    type: 'boolean' },
  { key: 'flag_ship',       value: 'true',           category: 'flags',     label: 'Show Ship To Address',     type: 'boolean' },
  { key: 'flag_hsn',        value: 'true',           category: 'flags',     label: 'Show HSN Codes',           type: 'boolean' },
  { key: 'flag_bank',       value: 'true',           category: 'flags',     label: 'Bank Details in Invoice',  type: 'boolean' },
  { key: 'flag_tds_purchase', value: 'true',         category: 'flags',     label: 'TDS on Purchase Invoice',  type: 'boolean' },
  { key: 'flag_tds_sales',  value: 'false',          category: 'flags',     label: 'TDS on Sales Invoice',     type: 'boolean' },
  { key: 'flag_wgst',       value: 'false',          category: 'flags',     label: 'TDS on Full Invoice Amount', type: 'boolean' },
  { key: 'flag_disc_gst',   value: 'false',          category: 'flags',     label: 'Discount includes GST',    type: 'boolean' },
  // When ON, the per-lot Discount (refund) is rolled into P_Rate via
  // the deduction1_inclusive percentage above — but ONLY for Grade 1
  // lots. Grade 2 and ungraded lots keep the original behaviour
  // (deduction1/2 + separate Discount) regardless of this flag.
  // When OFF, every lot follows the original behaviour.
  { key: 'flag_discount_in_prate', value: 'false',   category: 'flags',     label: 'Roll Discount into P_Rate (Grade 1 only)', type: 'boolean' },
  { key: 'flag_debit_note', value: 'false',          category: 'flags',     label: 'Debit Note for Discount',  type: 'boolean' },
  { key: 'flag_invoice_stripe', value: 'true',       category: 'flags',     label: 'Alternate Row Stripe in Invoice', type: 'boolean' },
  { key: 'flag_dummy',      value: 'true',           category: 'flags',     label: 'Allow Dummy Invoices',     type: 'boolean' },
  { key: 'flag_round',      value: 'true',           category: 'flags',     label: 'Round Invoice Amounts',    type: 'boolean' },
  { key: 'flag_export',     value: 'false',          category: 'flags',     label: 'Export Invoices',          type: 'boolean' },
  // WhatsApp share/send buttons across the app (Purchases, Payments,
  // Debit Notes, single-row icons). When false, the buttons are hidden
  // entirely so users on installs without WhatsApp Business / Web
  // access don't see dead controls.
  { key: 'flag_whatsapp',   value: 'false',          category: 'flags',     label: 'WhatsApp Share Buttons',   type: 'boolean' },
  // Bills of Supply (no-GST seller invoices) — installs that only deal
  // with GST-registered dealers can hide the entire Bills tab + Generate
  // Bill flow with this flag.
  { key: 'flag_bills',         value: 'true',         category: 'flags',     label: 'Bills of Supply',          type: 'boolean' },
  // Debit Notes — adjustment notes against purchases. Installs that
  // don't use this workflow can hide the Debit Notes tab + Generate
  // Debit Note flow with this flag.
  { key: 'flag_debit_notes',   value: 'true',         category: 'flags',     label: 'Debit Notes',              type: 'boolean' },
  // Price List Mapping — sister tool of the Lots → Price Import
  // flow. Installs that don't use the mapping workflow can hide the
  // sidebar entry AND the Lots-toolbar button with this flag.
  // Also gates the "Set Grade 1" bulk button on the Lots screen,
  // since that's part of the same mapping workflow.
  { key: 'flag_price_list_mapping', value: 'true',    category: 'flags',     label: 'Price List Mapping',       type: 'boolean' },
  // Print Selected Purchase — ASP-only "purchase-side mirror" PDF
  // for ticked invoices. Niche; most e-Trade installs don't use it.
  // Defaults to OFF so the button stays hidden until explicitly
  // enabled. Even with this ON the button still requires Kerala
  // business state (the ASP-context check that already gates it).
  { key: 'flag_print_selected_purchase', value: 'false', category: 'flags', label: 'Print Selected Purchase (ASP / Kerala only)', type: 'boolean' },
  // Lots → "Set Buyer" bulk button. Defaults OFF — installs that
  // don't routinely retag buyer codes after lot entry can keep the
  // toolbar simpler. When ON, the cardamom-green "Set Buyer" button
  // appears next to Delete Selected / Set Grade 1 whenever lots are
  // ticked. Server endpoint /api/lots/bulk-buyer is always present
  // regardless of this flag (it's permission-gated by lot_write).
  { key: 'flag_lot_set_buyer',     value: 'false',          category: 'flags',     label: 'Lots → Set Buyer bulk action', type: 'boolean' },
  // Price Check tab + transaction gate. When ON the operator gets the
  // Reports → Price Check tab, the gate banner, and a hard server-side
  // block on Calculate / Invoice / Purchase / Bill / Debit-Note
  // generation until verify clears for the active auction. When OFF
  // the tab is hidden, buttons are never disabled, and the gate is a
  // no-op (writes still happen). Default OFF so existing installs are
  // unaffected on upgrade.
  { key: 'flag_price_check',       value: 'false',          category: 'flags',     label: 'Price Check + transaction gate',  type: 'boolean' },
  // Lot-validation gate: when ON, an operator must run "Validate Entered
  // Lots" on a trade (resolve duplicate lots / lots with no seller, and
  // acknowledge sellers missing GSTIN/bank/PAN/phone) BEFORE price import
  // is allowed. Catches the "lots don't tally between Lots screen and
  // Dealer List" class of error (a no-GSTIN seller drops out of the
  // Dealer List). Default ON for this install. When OFF the validate
  // button still works as a report, but import is never blocked.
  { key: 'flag_lot_validation',    value: 'true',           category: 'flags',     label: 'Validate lots before price import', type: 'boolean' },
  // Date format used for display across UI tables, PDFs and Excel
  // exports. Three options:
  //   DD/MM/YYYY  → 18/05/2026 (Indian/UK default, current behaviour)
  //   DD-MM-YYYY  → 18-05-2026
  //   YYYY-MM-DD  → 2026-05-18 (ISO-style)
  // Storage in the DB is always ISO (YYYY-MM-DD); this setting only
  // affects how dates are rendered. Tally XML keeps its own YYYYMMDD
  // format because Tally itself requires it (machine-to-machine).
  { key: 'date_format',     value: 'DD/MM/YYYY',     category: 'display',   label: 'Date format',              type: 'select', options: ['DD/MM/YYYY','DD-MM-YYYY','YYYY-MM-DD'] },
  // ── SENSITIVE-FIELD MASKING (Display) ──────────────────────
  // Mask bank account no., IFSC code and phone number wherever they're
  // shown to people other than the operator — customer-facing receipt
  // PDFs and on-screen lists/detail panels. Each field has its own
  // policy. Modes: none / last4 / last6 / first4 / first6
  // (e.g. last4 → ********1234, first4 → 1234********). Functional
  // outputs — the bank payment file, payment advice PDF, DBF, Tally XML
  // and WhatsApp send targets — are NEVER masked (they need real values).
  //
  // mask_acct defaults to 'last4' because both lot-receipt renderers
  // (desktop HTML + the WhatsApp PDF) ALREADY masked the seller account
  // to last-4 unconditionally before this setting existed; defaulting to
  // 'none' would have silently exposed full account numbers on upgrade.
  // IFSC and phone were never masked, so they default to 'none' (no
  // change on upgrade) — the operator opts in.
  { key: 'mask_acct',       value: 'last4',          category: 'display',   label: 'Mask Bank Account No.',    type: 'select', options: ['none','last4','last6','first4','first6','full'] },
  { key: 'mask_ifsc',       value: 'none',           category: 'display',   label: 'Mask IFSC Code',           type: 'select', options: ['none','last4','last6','first4','first6','full'] },
  { key: 'mask_phone',      value: 'none',           category: 'display',   label: 'Mask Phone Number',        type: 'select', options: ['none','last4','last6','first4','first6','first2last2','full'] },

  // ── LOT ENTRY DEFAULTS (ported from PWA app.html) ──────────
  // These values pre-populate the Lot Entry form so field staff don't
  // re-type the same numbers every lot. Sample weight is the most
  // commonly-used one — it's the cardamom sample taken from each lot
  // for grading, typically a constant per season. Moisture, default
  // litre, and edit timeout match the PWA's config keys 1:1.
  { key: 'sample_weight',   value: '0.000',          category: 'lot_entry', label: 'Default Sample Weight (kg)', type: 'number' },
  { key: 'gunny_weight',    value: '0.000',          category: 'lot_entry', label: 'Default Gunny Weight (kg)', type: 'number' },
  { key: 'show_moisture',   value: 'false',          category: 'lot_entry', label: 'Show Moisture Column',     type: 'boolean' },
  // Master toggle for the extra lot-entry fields (Crop Receipt no. and
  // Reserved Price). When ON, both desktop and mobile lot-entry forms
  // show the two inputs and Recent Entries adds the matching columns.
  // When OFF, the fields are hidden but still stored on each lot (so
  // toggling back ON later doesn't lose data).
  { key: 'show_extra_lot_fields', value: 'false',    category: 'lot_entry', label: 'Show Extra Lot Fields (Crop Receipt, Reserved Price)', type: 'boolean' },
  { key: 'default_litre',   value: '',               category: 'lot_entry', label: 'Default Litre Weight',     type: 'text' },
  { key: 'default_crop_type', value: '',             category: 'lot_entry', label: 'Default Crop Type',        type: 'text' },
  { key: 'edit_enabled',    value: 'true',           category: 'lot_entry', label: 'Allow Lot Edits (non-admin)', type: 'boolean' },
  { key: 'edit_timeout_sec', value: '0',             category: 'lot_entry', label: 'Edit Timeout (sec; 0 = no limit)', type: 'number' },
  // Default lot receipt format. The Lot Entry print modal lets the
  // user override this per-print, but this setting decides which
  // option is pre-selected. "compact" matches the legacy thermal-
  // printer slip (~80mm wide, monospace, just lot/bags/qty/gross);
  // "detailed" is the modern A4-style ASPPL header with seller bank
  // details. Field staff on thermal hardware should set this to
  // "compact" once and forget it.
  { key: 'lot_receipt_format', value: 'detailed',     category: 'lot_entry', label: 'Lot Receipt Format (compact|detailed)', type: 'text' },
  // Physical paper width of the lot-receipt slip, in millimetres. Thermal
  // receipt printers come in fixed roll widths (e.g. the HOP-HL58 is a
  // 58mm roll, common alternatives are 80mm and 76mm). When this is blank
  // / 0 the slip uses its built-in default page size (80mm for compact,
  // 2.5in for detailed) — which on a narrower 58mm printer overflows the
  // paper and the driver silently falls back to scaling onto an A4 sheet.
  // Set this to the printer's roll width (58 for the HOP-HL58) so the
  // print @page size and the WhatsApp/PDF slip both match the paper.
  // Height is always automatic — receipts grow down the continuous roll.
  { key: 'lot_receipt_width_mm', value: '',           category: 'lot_entry', label: 'Lot Receipt Paper Width (mm; blank = default. e.g. 58 for HOP-HL58 thermal)', type: 'number' },
  // Optional columns on the DETAILED lot receipt. The compact slip never
  // shows these (it prints Lot / Bags / Weight + a Total row). Both default
  // OFF, so the detailed slip is Lot / Bag / Qty unless the user opts in.
  // Gross Wt = Net + Sample, computed live (same as the lot-entry form).
  { key: 'lot_receipt_show_sample', value: 'false',   category: 'lot_entry', label: 'Show Sample Wt column on Detailed Receipt', type: 'boolean' },
  { key: 'lot_receipt_show_gross',  value: 'false',   category: 'lot_entry', label: 'Show Gross Wt column on Detailed Receipt', type: 'boolean' },

  // ── BUSINESS MODE ──────────────────────────────────────────
  // This build is e-Trade only. business_mode is kept in the DB so calc
  // and exports continue to read it, but rendered as read-only in UI.
  { key: 'business_mode',   value: 'e-Trade',        category: 'mode',      label: 'Business Mode',            type: 'readonly' },
  { key: 'business_state',  value: 'TAMIL NADU',     category: 'mode',      label: 'Business State',           type: 'select' },

  // ── INTEGRATIONS ───────────────────────────────────────────
  { key: 'gst_api_key',     value: '',               category: 'integrations', label: 'GST Lookup API Key (gstincheck.co.in)', type: 'text' },
  // Extra-Lot Requests — when a mobile operator asks for more lots, the
  // admin gets a WhatsApp push (in addition to the in-app queue). Leave
  // the number blank to disable the push (the in-app queue still works).
  // Template blank = reuse the generic text template configured below.
  { key: 'extra_lot_alert_whatsapp', value: '',   category: 'integrations', label: 'Extra-Lot Requests — Admin WhatsApp number(s), comma-separated (with country code)', type: 'text' },
  { key: 'extra_lot_alert_tpl',      value: '',   category: 'integrations', label: 'Extra-Lot Requests — WhatsApp template name (blank = default text template)', type: 'text' },
  { key: 'extra_lot_alert_tpl_lang', value: 'en', category: 'integrations', label: 'Extra-Lot Requests — WhatsApp template language code', type: 'text' },

  // ── BACKUP (auto schedule) ────────────────────────────────
  // Server runs a periodic snapshot of the SQLite file into
  // `<DB_DIR>/backups/`. Frequency + retention configurable from
  // Settings. Default OFF — operator opts in.
  { key: 'backup_auto_enabled',   value: 'false', category: 'backup', label: 'Auto Backup',                       type: 'boolean' },
  { key: 'backup_interval_hours', value: '24',    category: 'backup', label: 'Backup Interval (hours)',           type: 'number' },
  { key: 'backup_keep_count',     value: '14',    category: 'backup', label: 'Keep last N backups',               type: 'number' },
  // ── TALLY EXPORT ──────────────────────────────────────────
  // Settings here mirror the macro's Configration form (UserForm1) field-for-field.
  // Identity & defaults
  { key: 'tally_company_name',     value: '',      category: 'tally', label: 'Tally Company Name', type: 'text' },
  { key: 'tally_season',          value: '2026-27',        category: 'tally', label: 'Season Suffix',                  type: 'text' },
  { key: 'tally_separator',       value: '/',              category: 'tally', label: 'Voucher Separator',              type: 'text' },
  { key: 'tally_inv_prefix',      value: '',               category: 'tally', label: 'Voucher Prefix', type: 'text' },
  { key: 'tally_state_code',      value: '33',             category: 'tally', label: 'Home GSTIN State Code (intra)',  type: 'text' },
  { key: 'tally_home_state',      value: 'Tamil Nadu',     category: 'tally', label: 'Home Place of Supply',           type: 'text' },
  { key: 'tally_urd_state',       value: 'Kerala',         category: 'tally', label: 'URD Purchase State (agriculturist)', type: 'text' },

  // Mode toggles (mirror the macro checkboxes)
  { key: 'tally_detailed',        value: 'true',           category: 'tally', label: 'Detailed Inv (one inventory entry per lot)',type: 'boolean' },
  // Purchase XML bill allocations only — independent of the global
  // `tally_detailed` flag so users can mix detailed inventory entries
  // with consolidated bills (or vice versa). When ON, one BILLALLOC-
  // ATIONS.LIST per lot is emitted. When OFF, a single consolidated
  // BILLALLOCATIONS.LIST is emitted with NAME = <ano>/<invoiceNo>/<season>.
  { key: 'tally_purchase_detailed', value: 'true',          category: 'tally', label: 'Purchase XML — Detailed bill allocations (one per lot)', type: 'boolean' },
  { key: 'tally_round_enabled',   value: 'true',           category: 'tally', label: 'Round (Round On/Off ledger)',               type: 'boolean' },
  { key: 'tally_tcs_enabled',     value: 'false',          category: 'tally', label: 'TCS (apply on Sales when applicable)',      type: 'boolean' },
  { key: 'tally_tds_enabled',     value: 'false',          category: 'tally', label: 'TDS (apply 194Q on RD Purchases)',          type: 'boolean' },
  { key: 'tally_optional',        value: 'false',          category: 'tally', label: 'Optional (mark vouchers as Optional)',      type: 'boolean' },
  { key: 'tally_dn_exempt',       value: 'false',          category: 'tally', label: 'Exempted (Debit Note: skip GST tax ledgers)', type: 'boolean' },
  { key: 'tally_ship_to',         value: 'false',          category: 'tally', label: 'Ship To (override consignee with separate Ship-To party)', type: 'boolean' },
  { key: 'tally_dispatch_from',   value: 'true',           category: 'tally', label: 'Dispatch From Address (emit DISPATCHFROMADDRESS in Sales XML)', type: 'boolean' },
  // E-way bill block in Sales XML — independent of the dispatch-from
  // address. Default ON; user can disable per install if they don't
  // want EWAYBILLDETAILS.LIST emitted (e.g. small intra-state supplies
  // below the e-way bill threshold where the buyer doesn't want it).
  { key: 'tally_eway_enabled',    value: 'true',           category: 'tally', label: 'E-way Bill (emit EWAYBILLDETAILS in Sales XML)', type: 'boolean' },
  // Dispatch (origin) PIN used by the e-way bill distance resolver.
  // Read by server.js#getDispatchPin() and by the per-invoice distance
  // hydration; route_distances rows are looked up keyed by this PIN
  // paired with the buyer's PIN. Must be a 6-digit Indian PIN.
  { key: 'tally_dispatch_pin',    value: '',               category: 'tally', label: 'Dispatch PIN (origin for e-way bill distance)', type: 'text' },

  // Sales Account Ledgers (Cardamom)
  { key: 'tally_sales_inter',     value: 'Cardamom Sales 5%',          category: 'tally', label: 'Cardamom Inter-State Sales',  type: 'text' },
  { key: 'tally_sales_intra',     value: 'Cardamom Sales 5% - Local',  category: 'tally', label: 'Cardamom Local Sales',        type: 'text' },
  { key: 'tally_sales_export',    value: 'Cardamom Sales - Export',    category: 'tally', label: 'Cardamom Export Sales (Deemed)', type: 'text' },

  // Sales Account Ledgers (Gunny)
  { key: 'tally_gunny_inter',     value: 'Gunny Sales 5%',             category: 'tally', label: 'Gunny Interstate Sales',      type: 'text' },
  { key: 'tally_gunny_intra',     value: 'Gunny Sales 5% - Local',     category: 'tally', label: 'Gunny Local Sales',           type: 'text' },
  { key: 'tally_gunny_export',    value: 'Gunny Sales - Export',       category: 'tally', label: 'Gunny Export Sales',          type: 'text' },

  // Dealer-Side Sales (when the company sells to a dealer)
  { key: 'tally_dealer_sale_inter', value: 'Interstate Dealer-Purchase', category: 'tally', label: 'Interstate Dealer (sales-side)', type: 'text' },
  { key: 'tally_dealer_sale_intra', value: 'Local Dealer-Purchase',      category: 'tally', label: 'Local Dealer (sales-side)',      type: 'text' },

  // RD Purchase ledgers (when the company buys from a dealer)
  { key: 'tally_purchase_dealer',     value: 'Trade Purchase From Dealer',category: 'tally', label: 'Trade Purchase From Dealer (base; gets -Local / -Inter_State suffix)', type: 'text' },
  { key: 'tally_purchase_dealer_inter', value: 'Interstate Dealer',      category: 'tally', label: 'Interstate Dealer-Pur (purchase-side)',     type: 'text' },
  { key: 'tally_purchase_dealer_intra', value: 'Local Dealer',           category: 'tally', label: 'Local Dealer-Pur (purchase-side)',          type: 'text' },

  // Agriculturist & TDS-on-sales
  { key: 'tally_purchase_auction',value: 'Purchase From Agriculturist', category: 'tally', label: 'Purchase From Agriculturist (URD ledger)', type: 'text' },
  { key: 'tally_tds_paid_sales',  value: 'TDS Paid on Sales',           category: 'tally', label: 'TDS Paid on Sales',           type: 'text' },

  // Tax Ledger Names — Sales 5% (output) and Purchase (input)
  { key: 'tally_cgst',            value: 'OUTPUT CGST 2.5%',           category: 'tally', label: 'CGST 2.5% (output)',          type: 'text' },
  { key: 'tally_sgst',            value: 'OUTPUT SGST 2.5%',           category: 'tally', label: 'SGST 2.5% (output)',          type: 'text' },
  { key: 'tally_igst',            value: 'OUTPUT IGST 5%',             category: 'tally', label: 'IGST 5% (output)',            type: 'text' },
  { key: 'tally_cgst_input',      value: 'INPUT CGST 2.5%',            category: 'tally', label: 'INPUT CGST 2.5%',             type: 'text' },
  { key: 'tally_sgst_input',      value: 'INPUT SGST 2.5%',            category: 'tally', label: 'INPUT SGST 2.5%',             type: 'text' },
  { key: 'tally_igst_input',      value: 'INPUT IGST 5%',              category: 'tally', label: 'INPUT IGST 5%',               type: 'text' },
  { key: 'tally_tcs',             value: 'TCS on Sale of Goods',       category: 'tally', label: 'TCS on Sale of Goods',        type: 'text' },
  { key: 'tally_tds_ledger',      value: 'TDS on Purchase of Goods',   category: 'tally', label: 'TDS on Purchase of Goods', type: 'text' },

  // Tax Ledger Names — Debit Note 18%
  { key: 'tally_dn_discount',     value: 'Discount on Purchase',       category: 'tally', label: 'Discount on Purch (Debit Note ledger)', type: 'text' },
  { key: 'tally_dn_cgst',         value: 'OUTPUT CGST 2.5%',           category: 'tally', label: 'CGST 2.5% (Debit Note)',      type: 'text' },
  { key: 'tally_dn_sgst',         value: 'OUTPUT SGST 2.5%',           category: 'tally', label: 'SGST 2.5% (Debit Note)',      type: 'text' },
  { key: 'tally_dn_igst',         value: 'OUTPUT IGST 5%',             category: 'tally', label: 'IGST 5% (Debit Note)',        type: 'text' },
  { key: 'tally_dn_gst_rate',     value: '5',                          category: 'tally', label: 'Debit Note GST Rate %',       type: 'number' },

  // Other operational ledgers
  { key: 'tally_sample_planter',  value: 'Sample Refund to Planter',   category: 'tally', label: 'Sample Refund to Planter',    type: 'text' },
  { key: 'tally_sample_dealer',   value: 'Sample Refund to Dealer',    category: 'tally', label: 'Sample Refund to Dealer',     type: 'text' },
  { key: 'tally_sample_stock',    value: 'false',                      category: 'tally', label: 'Stock (track sample refund as inventory)', type: 'boolean' },
  { key: 'tally_round',           value: 'Round On/Off',               category: 'tally', label: 'Round On/Off Ledger',         type: 'text' },
  { key: 'tally_transport',       value: 'Transport Charges',          category: 'tally', label: 'Transport Charges Ledger',    type: 'text' },
  { key: 'tally_insurance',       value: 'Insurance Charges',          category: 'tally', label: 'Insurance Charges Ledger',    type: 'text' },

  // Tax / commercial rates (the right-hand "Tax Rate" / "Item Rates" block)
  { key: 'tally_gst_rate',        value: '5',                          category: 'tally', label: 'GST Goods Rate %',            type: 'number' },
  { key: 'tally_service_rate',    value: '18',                         category: 'tally', label: 'Service Rate % (DN/Discount)', type: 'number' },
  { key: 'tally_tcs_rate',        value: '0.1',                        category: 'tally', label: 'TCS / TDS Rate %',            type: 'number' },
  { key: 'tally_export_rate',     value: '0',                          category: 'tally', label: 'Export GST Rate %',           type: 'number' },
  { key: 'tally_sample_kgs',      value: '1.900',                      category: 'tally', label: 'Sample Refund (Kgs)',         type: 'number' },
  { key: 'tally_gunny_rate',      value: '165',                        category: 'tally', label: 'Gunny Rate (₹ per bag)',      type: 'number' },
  { key: 'tally_transport_rate',  value: '2.50',                       category: 'tally', label: 'Transport Rate (₹/Kg, inter-state)', type: 'number' },
  { key: 'tally_local_trans_rate',value: '2.50',                       category: 'tally', label: 'Local Transport Rate (₹/Kg)', type: 'number' },
  { key: 'tally_insurance_rate',  value: '0.75',                       category: 'tally', label: 'Insurance Rate (₹/₹1000)',    type: 'number' },
  { key: 'tally_local_ins_rate',  value: '0.75',                       category: 'tally', label: 'Local Insurance Rate (₹/₹1000)', type: 'number' },

  // Stock Item Names + HSN
  { key: 'tally_item_cardamom',   value: 'Cardamom',                   category: 'tally', label: 'Stock Item — Cardamom',       type: 'text' },
  { key: 'tally_item_gunny',      value: 'Gunny Bag',                  category: 'tally', label: 'Stock Item — Gunny',          type: 'text' },
  { key: 'tally_hsn_cardamom',    value: '09083120',                   category: 'tally', label: 'HSN — Cardamom',              type: 'text' },
  { key: 'tally_hsn_gunny',       value: '63051040',                   category: 'tally', label: 'HSN — Gunny',                 type: 'text' },
  { key: 'tally_hsn_service',     value: '996111',                     category: 'tally', label: 'SAC — Service / Discount',    type: 'text' },
  { key: 'tally_hsn_transport',   value: '996791',                     category: 'tally', label: 'SAC — Transport',             type: 'text' },
  { key: 'tally_hsn_insurance',   value: '997136',                     category: 'tally', label: 'SAC — Insurance',             type: 'text' },

];

const CATEGORIES = {
  mode:       { order: 0, title: 'Business Mode',        icon: '⚙' },
  company:    { order: 1, title: 'Company Details',       icon: '🏢' },
  address_kl: { order: 2, title: 'Address (Kerala)',      icon: '📍' },
  address_tn: { order: 3, title: 'Address (Tamil Nadu)',  icon: '📍' },
  branches:   { order: 5, title: 'Branches & Contacts',  icon: '🏪' },
  rates:      { order: 6, title: 'Rates & Charges',       icon: '💰' },
  hsn:        { order: 7, title: 'HSN / SAC Codes',       icon: '🏷' },
  bank:       { order: 8, title: 'Bank Details',          icon: '🏦' },
  season:     { order: 9, title: 'Season / Financial Year', icon: '📅' },
  invoice:    { order: 10, title: 'Invoice Settings',     icon: '📄' },
  display:    { order: 10.5, title: 'Display',            icon: '🖼', description: 'Visual / formatting preferences. The date format is applied across UI tables, PDFs and Excel exports — storage in the database is always ISO (YYYY-MM-DD). Masking hides bank account no., IFSC and phone on receipts and on-screen lists/panels (pick how many digits stay visible); the bank payment file, DBF, Tally and WhatsApp targets always keep the real numbers.' },
  flags:      { order: 11, title: 'Feature Flags',        icon: '🔧' },
  lot_entry:  { order: 11.5, title: 'Lot Entry Defaults',   icon: '📝', description: 'Defaults applied when field staff enter lots from the Lot Entry tab. Sample weight is auto-filled into each new lot; moisture column shows when enabled; edit timeout limits how long after creation a non-admin user can edit their own lots.' },
  integrations: { order: 12, title: 'Integrations',       icon: '🔌', description: 'Optional third-party services. The GST API key enables auto-fetching trade name and address when you enter a GSTIN. Get a free key at gstincheck.co.in — sign up, copy the key from your dashboard, paste here.' },
  backup:     { order: 12.5, title: 'Auto Backup',        icon: '💾', description: 'Periodic snapshot of the database file into a local backups folder. Set the interval (hours) and how many backups to keep — older ones are pruned automatically.' },
  tally:      { order: 13, title: 'To Tally',             icon: '📤', description: 'Configure all settings for the Tally XML export — laid out exactly like the original Configration form. Ledger names here MUST match what exists in your Tally company; if a ledger is missing or misspelled, Tally will reject the import.' },
};

function initCompanySettings(db) {
  db.exec(`
    CREATE TABLE IF NOT EXISTS company_settings (
      key TEXT PRIMARY KEY,
      value TEXT NOT NULL DEFAULT '',
      category TEXT NOT NULL DEFAULT 'company',
      label TEXT NOT NULL DEFAULT '',
      field_type TEXT NOT NULL DEFAULT 'text'
    );
  `);

  const insert = db.prepare(
    'INSERT OR IGNORE INTO company_settings (key, value, category, label, field_type) VALUES (?, ?, ?, ?, ?)'
  );
  // Sync label / category / field_type for existing rows so older DBs
  // pick up renames done in the source (e.g. "ISP Tally Company Name"
  // → "Tally Company Name"). The user-edited `value` column is left
  // untouched — only metadata refreshes.
  const refresh = db.prepare(
    'UPDATE company_settings SET label = ?, category = ?, field_type = ? WHERE key = ?'
  );
  const seed = db.transaction(() => {
    for (const d of DEFAULTS) {
      insert.run(d.key, d.value, d.category, d.label, d.type);
      refresh.run(d.label, d.category, d.type, d.key);
    }
  });
  seed();

  // Clean up legacy keys that the original dual-company build wrote.
  // They are no longer used in this single-company e-Trade build. Safe
  // to drop on every startup — never referenced by any code path here.
  const REMOVED_KEYS = [
    'asp_profit', 'isp_profit',
    'asp_profit_pooler', 'asp_profit_dealer',
    'isp_profit_pooler', 'isp_profit_dealer',
    'inv_prefix_sister', 'dispatched_through', 'dispatched_through_asp',
    'commission_bill', 'memorandum_text',
    'flag_sister', 'flag_tnpa', 'flag_rtds_inv', 'flag_eway',
    's_logo', 's_company', 's_short_name', 's_address1', 's_address2',
    's_phone', 's_email', 's_gstin', 's_cin', 's_pan', 's_fssai', 's_sbl',
    'tally_asp_company_name', 'tally_ainv_prefix', 'tally_state_code_amazing',
    'tally_amazing_mode',
    'tally_commission', 'tally_cash_handling', 'tally_cash_handling_planter',
    'tally_chc_planter', 'tally_unit_rate',
    'tally_dispatch_company', 'tally_dispatch_address', 'tally_dispatch_place',
    // NOTE: tally_dispatch_pin must NOT be in this list — server.js's
    // getDispatchPin() and the e-way bill distance resolver still read
    // it. Including it here silently wipes the value on every boot,
    // so even after a user saves a route distance the resolver sees
    // an empty dispatch PIN and resolved_distance_km stays null.
    'tally_dispatch_state', 'tally_dispatch_gstin',
    // Distance auto-fill removed — workflow is manual per-invoice now.
    'distance_auto_enabled', 'distance_road_multiplier',
    // Local transport / insurance toggles moved out of Tally to Rates.
    'tally_local_transport', 'tally_local_insurance',
    // Roll-Discount-into-P_Rate is Grade-1-only — the Grade-2 "Dealer
    // discount-inclusive" variant is unused. Drop the orphan row so
    // it stops appearing in Settings → Rates.
    'deduction2_inclusive',
    // Superseded by per-field mask_acct / mask_ifsc / mask_phone (Display).
    // Value migrated to mask_acct just below the REMOVED_KEYS sweep runs.
    'acct_mask',
  ];
  // Sensitive-field masking superseded the single legacy `acct_mask`
  // switch (which only masked the seller account on mobile receipts and
  // used a different token set). Carry any user value over to the new
  // per-field `mask_acct` BEFORE the REMOVED_KEYS sweep drops acct_mask,
  // mapping the old tokens onto the new ones. Only migrate when the new
  // key is still at its 'none' default so we never stomp a fresh choice.
  try {
    const oldRow = db.prepare("SELECT value FROM company_settings WHERE key = 'acct_mask'").get();
    const newRow = db.prepare("SELECT value FROM company_settings WHERE key = 'mask_acct'").get();
    if (oldRow && newRow && (newRow.value === 'none' || !newRow.value)) {
      const map = { show_last4: 'last4', show_last4_star: 'last4', show_first4_last4: 'last4', none: 'none' };
      const mapped = map[oldRow.value];
      if (mapped && mapped !== 'none') {
        db.prepare("UPDATE company_settings SET value = ? WHERE key = 'mask_acct'").run(mapped);
      }
    }
  } catch (e) {}

  const drop = db.prepare('DELETE FROM company_settings WHERE key = ?');
  for (const k of REMOVED_KEYS) drop.run(k);

  // Drop preset tables — Logo Code is a plain textbox in this build,
  // there is no dual-company preset switching.
  try { db.exec('DROP TABLE IF EXISTS company_presets'); } catch (e) {}
  try { db.exec('DROP TABLE IF EXISTS company_preset_meta'); } catch (e) {}

  // Force business_mode to 'e-Trade'. A stale DB row from the previous
  // mode-switching build must not leak through to calculation paths.
  db.prepare("UPDATE company_settings SET value = 'e-Trade' WHERE key = 'business_mode'").run();

  // Clear any stale 'ASP' Logo Code left behind by the legacy dual-
  // company app. Set to blank so the user explicitly fills in their
  // identifier — no fabricated code appearing where they never typed
  // one.
  db.prepare("UPDATE company_settings SET value = '' WHERE key = 'logo' AND UPPER(COALESCE(value,'')) = 'ASP'").run();

  // Drop the short-lived `default_page_size` row if a previous build
  // seeded it. Per-list pagination is now managed entirely via the
  // pager footer dropdown (localStorage); the company-wide default was
  // removed because changes to it weren't propagating reliably and the
  // footer-level control covers the use case on its own.
  db.prepare("DELETE FROM company_settings WHERE key = 'default_page_size'").run();

  // Clear the legacy AMAZING SPICE PARK default that was seeded into
  // `tn_dispatch` by earlier builds. We only clear when the value
  // matches the legacy literal exactly — any user-edited dispatch
  // address is preserved.
  db.prepare(
    "UPDATE company_settings SET value = '' WHERE key = 'tn_dispatch' AND value = ?"
  ).run('AMAZING SPICE PARK PVT LTD WARD No.6 ELLIKKANAM DOOR No.650 NEDUMKANDAM IDUKKI KERALA CODE:32');

  // ── DN GST rate / ledger migration ──
  // Earlier builds seeded the Debit Note ledgers at 18% (`OUTPUT IGST 18%`,
  // `OUTPUT CGST 9%`, `OUTPUT SGST 9%`, rate=18). The correct values for
  // the cardamom-export business are 5% (IGST) / 2.5% each (CGST/SGST).
  // Existing installs would still hold the 18% strings in their DB until
  // a manual edit. We auto-correct ONLY when the values match the legacy
  // un-customized defaults exactly — any user-edited string is preserved.
  const fixIfLegacy = (key, legacyVal, newVal) => {
    db.prepare(
      `UPDATE company_settings SET value = ? WHERE key = ? AND value = ?`
    ).run(newVal, key, legacyVal);
  };
  fixIfLegacy('tally_dn_cgst',     'OUTPUT CGST 9%',  'OUTPUT CGST 2.5%');
  fixIfLegacy('tally_dn_sgst',     'OUTPUT SGST 9%',  'OUTPUT SGST 2.5%');
  fixIfLegacy('tally_dn_igst',     'OUTPUT IGST 18%', 'OUTPUT IGST 5%');
  fixIfLegacy('tally_dn_gst_rate', '18',              '5');

  // praman_company is now the user's Nickname under Company Details —
  // relocate any pre-existing row and refresh its label.
  db.prepare(
    "UPDATE company_settings SET category = 'company', label = ? WHERE key = 'praman_company'"
  ).run('Nickname (fallback: Short Name)');

  // Local transport / insurance toggles moved from To Tally → Rates &
  // Charges and renamed flag_local_*. Carry over any existing user value
  // before the legacy keys get dropped via REMOVED_KEYS.
  const copyOldFlag = (oldKey, newKey) => {
    const r = db.prepare('SELECT value FROM company_settings WHERE key = ?').get(oldKey);
    if (r && (r.value === 'true' || r.value === 'false')) {
      db.prepare('UPDATE company_settings SET value = ? WHERE key = ?').run(r.value, newKey);
    }
  };
  copyOldFlag('tally_local_transport', 'flag_local_transport');
  copyOldFlag('tally_local_insurance', 'flag_local_insurance');

  console.log('Company settings ready (%d defaults)', DEFAULTS.length);
}

function getSetting(db, key) {
  const r = db.prepare('SELECT value FROM company_settings WHERE key = ?').get(key);
  return r ? r.value : null;
}

function getSettingBool(db, key) {
  const v = getSetting(db, key);
  return v === 'true' || v === '1';
}

function getSettingNum(db, key) {
  return parseFloat(getSetting(db, key)) || 0;
}

function getAllSettings(db) {
  const rows = db.prepare('SELECT key, value, category, label, field_type FROM company_settings ORDER BY rowid').all();
  const grouped = {};
  for (const r of rows) {
    if (!grouped[r.category]) grouped[r.category] = [];
    grouped[r.category].push(r);
  }
  return grouped;
}

function updateSettings(db, settings) {
  const upd = db.prepare('UPDATE company_settings SET value = ? WHERE key = ?');
  const batch = db.transaction((items) => {
    let n = 0;
    for (const [k, v] of Object.entries(items)) { upd.run(String(v), k); n++; }
    return n;
  });
  return batch(settings);
}

function getSettingsFlat(db) {
  const rows = db.prepare('SELECT key, value, field_type FROM company_settings').all();
  const flat = {};
  for (const r of rows) {
    if (r.field_type === 'boolean') flat[r.key] = r.value === 'true';
    else if (r.field_type === 'number') flat[r.key] = parseFloat(r.value) || 0;
    else flat[r.key] = r.value;
  }
  return flat;
}

function getGSTRates(db) {
  const g = getSettingNum(db, 'gst_goods');
  return { cgst: g / 2, sgst: g / 2, igst: g, service: getSettingNum(db, 'gst_service'), tcs: getSettingNum(db, 'tcs_tds') };
}

module.exports = { DEFAULTS, CATEGORIES, initCompanySettings, getSetting, getSettingBool, getSettingNum, getAllSettings, updateSettings, getSettingsFlat, getGSTRates };
