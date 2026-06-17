/**
 * calculations.js — Core business logic
 * Replaces: GENERATE.PRG, parts of GSTKBILT/GSTKBILP/GSTBILP/PAYCHECK
 */

const { getSettingsFlat, getGSTRates } = require('./company-config');

// ── Excel-compatible rounding ───────────────────────────────────────────
// Plain JS `Math.round(x * 100) / 100` drifts on values whose binary
// representation falls just short of the half mark — e.g.
//     1.005 * 100  → 100.49999999999999  → rounds to 100 → 1.00 (not 1.01)
//     2.675 * 100  → 267.49999999999994  → rounds to 267 → 2.67 (not 2.68)
// Excel / Google Sheets ROUND() avoid this because their decimal pipeline
// doesn't go through IEEE-754 binary multiplication. We use the same
// trick: serialize the number to a decimal string, append "e2" / "e-2"
// to shift the decimal point via the parser (which is decimal-exact for
// finite doubles per ECMA-262), then Math.round the integerized value.
// This makes round2(1.005) === 1.01, round2(2.675) === 2.68 — i.e.,
// matches Excel ROUND() byte-for-byte on every realistic finance input.
//
// `round2` = round to 2 decimal places (paisa-precision money values).
// `round0` = round to integer (whole rupee — used by the Round on/off line).
// Both are sign-aware (Excel's "round half away from zero").
const round2 = (n) => {
  const x = Number(n);
  if (!isFinite(x)) return 0;
  if (x === 0) return 0;
  const sign = x < 0 ? -1 : 1;
  // Shift two decimal places right via string ('1.005' → '1.005e2' → 100.5),
  // round, then shift back left.
  const shifted = Number(Math.abs(x) + 'e2');
  return sign * Number(Math.round(shifted) + 'e-2');
};
const round0 = (n) => {
  const x = Number(n);
  if (!isFinite(x)) return 0;
  if (x === 0) return 0;
  return (x < 0 ? -1 : 1) * Math.round(Math.abs(x));
};

/**
 * Extract the 2-digit state code from a seller's `cr` field.
 *
 * The `cr` column historically stored "GSTIN.<15-char-gstin>" (where the
 * "GSTIN." prefix was added by the UI for sellers). Some import paths
 * (Excel seller import, GST portal lookup, edge-case manual entry) end up
 * with a bare GSTIN without the prefix. Both forms should be supported,
 * because the state code only depends on the first 2 chars of the GSTIN
 * itself, not on the prefix.
 *
 * Logic:
 *   "GSTIN.32AAACG1234F1Z2" → strip "GSTIN." prefix → first 2 chars → "32"
 *   "32AAACG1234F1Z2"       → already bare         → first 2 chars → "32"
 *   ""  / null / undefined  → ""                                    (no GSTIN)
 *   "CR.001" or other       → ""                                    (not a GSTIN)
 *
 * Only a value that looks like a valid 15-char GSTIN starting with 2 digits
 * yields a state code; anything else returns "" (which means no intra-match,
 * so IGST applies — safe default for non-GSTIN sellers).
 */
function gstinStateCode(cr) {
  if (!cr) return '';
  let s = String(cr).trim().toUpperCase();
  if (s.startsWith('GSTIN.')) s = s.substring(6);
  else if (s.startsWith('GSTIN')) s = s.substring(5);  // e.g. "GSTIN<no-dot>33AAA..."
  // GSTIN format: 2 digits + 5 letters + 4 digits + 1 letter + 1 digit + Z + 1 alphanumeric
  // We only need the first 2 chars, but verify they're digits.
  if (s.length < 2) return '';
  const head = s.substring(0, 2);
  if (!/^\d{2}$/.test(head)) return '';
  return head;
}

/**
 * Calculate purchase amounts for a lot (after trade)
 * This is what GENERATE.PRG does — fills pqty, prate, puramt, com, gst, etc.
 *
 * e-Trade only build: single-company calculation. The legacy dual-company planter
 * views from the original Spice Config app have been removed — there is
 * just one P_Qty / P_Rate / PurAmt path here. The dual-storage columns
 * (isp_/asp_) are kept and mirrored to the single value so existing
 * SELECT statements in legacy reports keep working.
 */
function calculateLot(lot, cfg) {
  const result = { ...lot };
  const gstGoods = cfg.gst_goods || 5;

  // Purchase qty = qty + Sample Refund (from Rates & Charges, global)
  // Falls back to per-lot `refud` for back-compat with older data.
  const sampleRefundKg = (cfg.refund != null && cfg.refund !== '') ? Number(cfg.refund) : (lot.refud || 0);
  result.pqty = (lot.qty || 0) + (Number(sampleRefundKg) || 0);

  // Single-company P_Rate / PurAmt:
  //   Grade 2 → deduction2 (Dealer)
  //   else    → deduction1 (Pooler)
  // The flag_discount_in_prate roll-in applies ONLY to Grade 1 lots —
  // Grade 2 and ungraded lots keep the original deduction/discount
  // behaviour regardless of the flag. When the flag is ON AND the lot
  // is Grade 1, deduction1_inclusive is used and the per-lot Discount
  // is forced to 0 further down because it's baked into the rate.
  const gradeStr = String(lot.grade || '').trim();
  const discInPrateFlag = cfg.flag_discount_in_prate === true
    || String(cfg.flag_discount_in_prate || '').toLowerCase() === 'true';
  const applyRollIn = discInPrateFlag && gradeStr === '1';
  let deduction;
  if (gradeStr === '2') {
    deduction = Number(cfg.deduction2 || 0);
  } else if (applyRollIn) {
    deduction = Number(cfg.deduction1_inclusive || 0);
  } else {
    deduction = Number(cfg.deduction1 || 0);
  }
  const rawRate = (lot.price || 0) * (1 - deduction / 100);
  result.prate = round0(rawRate);
  // PurAmt = P_Qty × P_Rate (sample refund INCLUDED — direct purchase)
  result.puramt = round2(result.pqty * result.prate);
  // Sale amount (raw, pre-deduction) — kept in sync with qty × price so a
  // price edit via Price Check Apply / Set price doesn't leave `amount`
  // stale. The validate endpoint flags `ROUND(qty*price,2) <> ROUND(amount,2)`
  // as a data-integrity bug; Calculate All is the natural place to heal it.
  result.amount = round2((lot.qty || 0) * (lot.price || 0));
  result.com = 0;
  result.sertax = 0;

  // Mirror dual-storage columns for legacy SELECTs that still read isp_*/asp_*.
  result.isp_pqty = result.pqty;
  result.isp_prate = result.prate;
  result.isp_puramt = result.puramt;
  result.asp_pqty = result.pqty;
  result.asp_prate = result.prate;
  result.asp_puramt = result.puramt;

  // Intra/inter-state detection: seller's GSTIN state code (first 2 digits)
  // vs company's state code (cfg.business_state).
  const sellerGstState = gstinStateCode(lot.cr);
  const companyGstState = cfg.business_state === 'KERALA' ? '32' : '33';
  const isIntra = (sellerGstState === companyGstState);

  // GST on the Discount. Single dedicated `discount_gst` setting drives
  // this so the user can configure it independently from gst_service
  // (which historically defaulted to 18%, but the actual discount-GST
  // rate in the trade can be different — typically 5%). Falls back to
  // gst_service for older DBs that don't have the new key yet.
  const gstDiscRate = Number(cfg.discount_gst);
  const gstServiceFallback = Number(cfg.gst_service) || 18;
  const discRate = Number.isFinite(gstDiscRate) && gstDiscRate >= 0 ? gstDiscRate : gstServiceFallback;
  const halfRate = discRate / 2;

  result.cgst = 0;
  result.sgst = 0;
  result.igst = 0;

  // e-Trade discount: round(PurAmt / 1000 × days × discount%) — nearest
  // rupee, half-up. result.refund holds the Discount amount.
  // Days source depends on seller type:
  //   GSTIN present (registered dealer) → cfg.dealer_days
  //   no GSTIN (CR / agriculturist)     → cfg.discount_days
  // SKIPPED when the roll-in is in effect for THIS lot (flag ON AND
  // Grade 1) — the discount is already baked into P_Rate via
  // deduction1_inclusive, so a separate refund value would double-
  // count it. GST on Discount is skipped for the same reason.
  // Grade 2 / ungraded lots fall through to the normal discount calc
  // even when the flag is ON.
  const sellerHasGstin = sellerGstState !== '';
  // "Discount In Payments" master switch (flag_sample). When OFF, no per-lot
  // discount is calculated — refund and GST-on-discount stay 0 and the full
  // PurAmt flows through to Payable. When ON, the e-Trade discount formula
  // below runs (still skipped for Grade-1 roll-in, which bakes it into P_Rate).
  const discountEnabled = cfg.flag_sample === true
    || String(cfg.flag_sample || '').toLowerCase() === 'true';
  if (applyRollIn || !discountEnabled) {
    result.refund = 0;
  } else {
    const days    = sellerHasGstin
      ? (Number(cfg.dealer_days) || 0)
      : (Number(cfg.discount_days) || 0);
    const discPct = Number(cfg.discount_pct)  || 0;
    result.refund = round0((result.puramt / 1000) * days * discPct);

    // GST on the Discount only when flag_disc_gst is ON.
    if (cfg.flag_disc_gst && result.refund > 0) {
      if (isIntra) {
        result.cgst = round2(result.refund * halfRate / 100);
        result.sgst = round2(result.refund * halfRate / 100);
      } else {
        result.igst = round2(result.refund * discRate / 100);
      }
    }
  }

  // advance = GST only (informational; e-Trade has no commission/handling)
  result.advance = result.cgst + result.sgst + result.igst;

  // Payable = PurAmt − Discount − GST-on-Discount
  const totalDeductions = result.refund + result.cgst + result.sgst + result.igst;
  result.balance = round2(result.puramt - totalDeductions);

  // Bill amount (for agriculturist bills) — always equals PurAmt
  result.bilamt = result.puramt;

  return result;
}

/**
 * Calculate TDS under Section 194Q (TDS on Purchase of Goods).
 *
 * Per Section 194Q, a buyer whose turnover exceeds ₹10 cr in the prior FY
 * must deduct TDS at 0.1% of the purchase amount in EXCESS of ₹50 lakh
 * paid to a single seller in the current FY. Once the threshold is crossed,
 * TDS applies to every subsequent rupee bought from that seller for the
 * rest of the year.
 *
 * Inputs (must all be on the SAME basis — either all incl-GST or all
 *   excl-GST; the caller is responsible for keeping units consistent):
 *   purchaseAmount  — current trade's purchase amount (this voucher)
 *   priorPurchases  — sum of all prior purchases from the same seller in
 *                      the current FY (excluding this trade)
 *   cfg.tds_threshold     — usually 5000000 (₹50 L); configurable
 *   cfg.tds_purchase_rate — usually 0.1 (%); configurable, fall-back to
 *                            tcs_tds for back-compat with older configs
 *
 * Returns: TDS amount in rupees (rounded up to nearest paisa).
 */
function calculateTDS(purchaseAmount, priorPurchases, cfg) {
  const threshold = Number(cfg.tds_threshold) || 5000000;
  // Prefer the dedicated TDS-on-purchase rate; fall back to the legacy
  // shared tcs_tds setting if the new key isn't configured yet.
  const tdsRate = Number(cfg.tds_purchase_rate) || Number(cfg.tcs_tds) || 0.1;

  if (priorPurchases > threshold) {
    // Already crossed threshold this FY — TDS on the full new purchase
    return Math.ceil(purchaseAmount * tdsRate / 100);
  } else if ((priorPurchases + purchaseAmount) > threshold) {
    // This trade crosses the threshold — TDS only on the portion above
    const excess = priorPurchases + purchaseAmount - threshold;
    return Math.ceil(excess * tdsRate / 100);
  }
  return 0;
}

/**
 * Lot-wise TDS for the PAYMENTS surfaces (summary / lots modal / statement /
 * bank export). Deliberately FLAT: `rate × purchase amount`, with NO ₹50-lakh
 * cumulative threshold — so the figure is just the rate on each lot's amount
 * (per the user's request to show TDS lot-wise rather than the cumulative
 * Section 194Q value). This intentionally differs from the purchase INVOICE's
 * TDS (`calculateTDS`, threshold-based) — the invoice stays on 194Q; only the
 * payment display/transfer uses this.
 *
 * Scoped like 194Q to registered (GSTIN-bearing) dealers only — URD /
 * agriculturist sellers never attract purchase TDS — and gated by the
 * `flag_tds_purchase` feature flag. Returns rupees, 2-dp.
 */
function lotwisePurchaseTds(purchaseAmount, cr, cfg) {
  const enabled = cfg.flag_tds_purchase === true
    || String(cfg.flag_tds_purchase || '').toLowerCase() === 'true';
  if (!enabled) return 0;
  if (gstinStateCode(cr) === '') return 0; // GSTIN-registered dealers only
  const rate = Number(cfg.tds_purchase_rate) || Number(cfg.tcs_tds) || 0.1;
  return round2((Number(purchaseAmount) || 0) * rate / 100);
}

/**
 * Calculate TCS under Section 206C(1H) — TCS on Sale of Goods.
 * Threshold logic mirrors TDS-on-purchase: TCS applies to amounts in
 * EXCESS of ₹50 lakh per buyer per FY, then to every subsequent rupee.
 */
function calculateTCS(invoiceAmount, priorSales, cfg) {
  const threshold = Number(cfg.tds_threshold) || 5000000;
  const tcsRate = Number(cfg.tcs_tds) || 0.1;

  if (priorSales > threshold) {
    return Math.ceil(invoiceAmount * tcsRate / 100);
  } else if ((priorSales + invoiceAmount) > threshold) {
    const excess = priorSales + invoiceAmount - threshold;
    return Math.ceil(excess * tcsRate / 100);
  }
  return 0;
}

/**
 * Build sales invoice data for a buyer
 * Aggregates lots by buyer for a given auction
 * Sale type filter is optional — if lots don't have sale set yet, filter by buyer only
 */
function buildSalesInvoice(db, auctionId, buyerCode, saleType, cfg) {
  // Get all lots for this buyer in this auction that have amounts
  // Don't filter by sale — we're ASSIGNING the sale type now.
  // Locked lots are excluded — they're finalized records and must not
  // be picked up into a new invoice. The corresponding stamping UPDATE
  // in server.js also carries `AND locked_at IS NULL` as a belt-and-
  // braces guard against any caller that bypasses this helper.
  const lots = db.all(
    `SELECT * FROM lots WHERE auction_id = ? AND buyer = ? AND amount > 0
     AND locked_at IS NULL
     AND (sale IS NULL OR sale = '' OR sale = ?) ORDER BY lot_no`,
    [auctionId, buyerCode, saleType]
  );
  
  if (!lots.length) return null;

  const gstGoods = cfg.gst_goods || 5;
  const companyState = cfg.business_state === 'KERALA' ? '32' : '33';
  
  // Get buyer details
  const buyer = db.get('SELECT * FROM buyers WHERE buyer = ?', [buyerCode]);
  const buyerState = buyer ? buyer.gstin.substring(0, 2) : companyState;
  const isInterState = buyerState !== companyState;

  let totalQty = 0, totalBags = 0, totalAmount = 0;
  const lineItems = [];

  for (const lot of lots) {
    totalBags += lot.bags;
    const calc = calculateLot(lot, cfg);
    const prate = calc.prate;
    const puramt = calc.puramt;

    totalQty += lot.qty;
    totalAmount += lot.amount;

    lineItems.push({
      lot: lot.lot_no, grade: lot.grade, bags: lot.bags, qty: lot.qty,
      price: lot.price, amount: lot.amount,
      prate: prate, puramt: puramt,
    });
  }

  // Gunny cost (HSN: jute bags)
  const gunnyCost = totalBags * (cfg.gunny_rate || 165);

  // Transport & Insurance rates depend on sale type:
  //   L (Local)        → local_transport / local_insurance
  //   I (Inter-state)  → buyer covers freight; T/I hidden from invoice
  //   E (Export)       → use inter-state rates (same interstate logistics)
  const pickRate = (...vals) => {
    for (const v of vals) {
      if (v === undefined || v === null || v === '') continue;
      const n = typeof v === 'number' ? v : parseFloat(v);
      if (!Number.isNaN(n)) return n;
    }
    return 0;
  };
  // Sale-type-driven rate selection:
  //   L (Local)       → local_transport / local_insurance, gated by
  //                     flag_local_transport / flag_local_insurance
  //   I (Inter-state) → transport / insurance, gated by
  //                     flag_inter_transport / flag_inter_insurance
  //   E (Export)      → zero (buyer covers freight; matches the
  //                      hideTransportInsurance render rule in invoice-pdf.js)
  // When the matching flag is OFF the rate is forced to 0, so the
  // component drops out of the taxable value, GST, ledger emission, and
  // PDF entirely. Anything else (legacy / blank) is treated as 'L' for
  // safety.
  const st = String(saleType || '').toUpperCase();
  const isExport = (st === 'E');
  const isInter  = (st === 'I');
  const flagOn = (k, defaultOn) => {
    const v = cfg[k];
    if (v === undefined || v === null || v === '') return defaultOn;
    return v === true || String(v).toLowerCase() === 'true';
  };
  const useLocalTransport = flagOn('flag_local_transport', true);
  const useLocalInsurance = flagOn('flag_local_insurance', true);
  const useInterTransport = flagOn('flag_inter_transport', true);
  const useInterInsurance = flagOn('flag_inter_insurance', true);
  const transportRate = isExport
    ? 0
    : (isInter
        ? (useInterTransport ? pickRate(cfg.transport, 2.5) : 0)
        : (useLocalTransport ? pickRate(cfg.local_transport, cfg.transport, 2.5) : 0));
  const insuranceRate = isExport
    ? 0
    : (isInter
        ? (useInterInsurance ? pickRate(cfg.insurance, 0.75) : 0)
        : (useLocalInsurance ? pickRate(cfg.local_insurance, cfg.insurance, 0.75) : 0));

  // Transport: ₹/kg (qty × rate)
  const transportCost = round2(totalQty * transportRate);

  // Insurance: per ₹1000 of (cardamom + gunny + GST on those)
  //   insurance = ((cardamom_amount + gunny_cost) × (1 + gstGoods/100)) / 1000 × rate
  const subtotalGoods = totalAmount + gunnyCost;
  const gstOnGoods = subtotalGoods * gstGoods / 100;
  const insuranceCost = round2((subtotalGoods + gstOnGoods) / 1000 * insuranceRate);

  // Taxable value = cardamom + gunny + transport + insurance
  const taxableValue = subtotalGoods + transportCost + insuranceCost;

  // Per-line GST: tax each component independently and round at the line
  // level, then sum into the invoice totals. This matches the GSTN /
  // E-invoice convention (each HSN line carries its own tax) and how the
  // dad's worksheet runs the math — a single bundled round on the whole
  // taxable can drift by 1 paisa vs the line-by-line sum at certain
  // values, and Tally rejects vouchers when the line tax sum != invoice
  // tax total. All components share the SAME gstGoods rate (per the
  // existing rule); split is intra (CGST + SGST) vs inter (IGST).
  const halfRate = gstGoods / 2;
  const taxLine = (base) => isInterState
    ? { cgst: 0, sgst: 0, igst: round2(base * gstGoods / 100) }
    : { cgst: round2(base * halfRate / 100), sgst: round2(base * halfRate / 100), igst: 0 };

  let cgst = 0, sgst = 0, igst = 0;
  // 1) Cardamom — one line per lot. Tax stored on the lineItem AND
  //    accumulated into a per-HSN bucket (`cardamomTax`) so the HSN
  //    summary table on the PDF reads the SAME numbers we used for the
  //    OUTPUT GST line — no re-rounding on the bundled total (which
  //    drifts by 1 paisa at .x05 boundaries).
  const cardamomTax = { cgst: 0, sgst: 0, igst: 0 };
  for (const li of lineItems) {
    const t = taxLine(li.amount || 0);
    li.cgst = t.cgst; li.sgst = t.sgst; li.igst = t.igst;
    cardamomTax.cgst += t.cgst; cardamomTax.sgst += t.sgst; cardamomTax.igst += t.igst;
  }
  cardamomTax.cgst = round2(cardamomTax.cgst);
  cardamomTax.sgst = round2(cardamomTax.sgst);
  cardamomTax.igst = round2(cardamomTax.igst);
  // 2) Gunny / 3) Transport / 4) Insurance — one row each. Same rule:
  //    PDF / XML must consume these objects directly, never re-derive.
  const gunnyTax     = taxLine(gunnyCost);
  const transportTax = taxLine(transportCost);
  const insuranceTax = taxLine(insuranceCost);

  cgst = round2(cardamomTax.cgst + gunnyTax.cgst + transportTax.cgst + insuranceTax.cgst);
  sgst = round2(cardamomTax.sgst + gunnyTax.sgst + transportTax.sgst + insuranceTax.sgst);
  igst = round2(cardamomTax.igst + gunnyTax.igst + transportTax.igst + insuranceTax.igst);

  const totalBeforeRound = taxableValue + cgst + sgst + igst;
  const subtotalRounded = round0(totalBeforeRound);
  const roundDiff = subtotalRounded - totalBeforeRound;

  // Additional Charge — sum(cardamom) × cfg.addl_charge_value % .
  // The configured value is a PERCENTAGE (e.g. 2 means 2%). Sits BELOW the
  // Round on/off line and adds straight onto the grand total — does not
  // feed into GST or round-off math. When the percentage is 0 the charge
  // is fully skipped (no row, no XML ledger, no effect on grand total).
  const addlChargePct = Number(cfg.addl_charge_value) || 0;
  const addlCharge = addlChargePct > 0
    ? round2(totalAmount * addlChargePct / 100)
    : 0;
  const addlChargeName = addlCharge > 0 ? String(cfg.addl_charge_name || '').trim() : '';
  const grandTotal = addlCharge > 0
    ? round2(subtotalRounded + addlCharge)
    : subtotalRounded;

  return {
    buyer: buyer || {},
    saleType,
    lineItems,
    summary: {
      totalQty, totalBags, totalAmount,
      gunnyCost, transportCost, insuranceCost,
      taxableValue, cgst, sgst, igst,
      // Per-component pre-rounded GST — consumed by the HSN summary
      // block on the PDF so it can render the exact same numbers the
      // OUTPUT line shows (no re-derivation on a bundled total).
      cardamomTax, gunnyTax, transportTax, insuranceTax,
      roundDiff, subtotalRounded,
      addlCharge, addlChargeName,
      grandTotal,
      isInterState
    }
  };
}

/**
 * Build purchase invoice data for a seller
 * Aggregates lots by seller for a given auction (registered dealers only)
 */
function buildPurchaseInvoice(db, auctionId, sellerName, cfg) {
  // A lot qualifies for a Purchase Invoice if it has a GSTIN-bearing seller —
  // i.e. cr is either "GSTIN.<15-char>" (legacy UI format) or a bare 15-char
  // GSTIN starting with 2 digits (Excel import format). We accept both.
  const lots = db.all(
    `SELECT * FROM lots
     WHERE auction_id = ? AND name = ? AND amount > 0
       AND (UPPER(cr) LIKE 'GSTIN%' OR cr GLOB '[0-9][0-9]*')
     ORDER BY lot_no`,
    [auctionId, sellerName]
  );
  
  if (!lots.length) return null;

  const gstGoods = cfg.gst_goods || 5;
  const companyState = cfg.business_state === 'KERALA' ? '32' : '33';

  let totalQty = 0, totalPuramt = 0, totalBags = 0;
  const lineItems = [];

  for (const lot of lots) {
    const sellerState = gstinStateCode(lot.cr);
    const isInter = sellerState !== companyState;
    const puramt = lot.puramt || 0;

    const rcgst = isInter ? 0 : round2(puramt * (gstGoods / 2) / 100);
    const rsgst = isInter ? 0 : round2(puramt * (gstGoods / 2) / 100);
    const rigst = isInter ? round2(puramt * gstGoods / 100) : 0;

    totalQty += lot.pqty || lot.qty;
    totalPuramt += puramt;
    totalBags += lot.bags || 0;

    lineItems.push({
      lot: lot.lot_no, bags: lot.bags, grade: lot.grade,
      qty: lot.qty, pqty: lot.pqty,
      price: lot.price, prate: lot.prate,
      amount: lot.amount, puramt, 
      com: lot.com, sertax: lot.sertax,
      cgst: rcgst, sgst: rsgst, igst: rigst
    });
  }

  const firstLot = lots[0];
  const sellerState = gstinStateCode(firstLot.cr);
  const isInter = sellerState !== companyState;

  let totalCgst = 0, totalSgst = 0, totalIgst = 0;
  lineItems.forEach(li => { totalCgst += li.cgst; totalSgst += li.sgst; totalIgst += li.igst; });

  const totalBeforeRound = totalPuramt + totalCgst + totalSgst + totalIgst;
  const grandTotal = round0(totalBeforeRound);
  const roundDiff = grandTotal - totalBeforeRound;

  // ── TDS calculation (Section 194Q) ──
  //
  // 1) GSTIN format compatibility: the purchases table may have rows with
  //    gstin in either form ("GSTIN.32AAA..." or bare "32AAA..."). We
  //    derive both candidates from the current lot's cr and match either.
  //
  // 2) Amount basis must match: this trade's amount and the running prior
  //    total must be on the SAME basis (both with-GST or both excl-GST),
  //    otherwise the threshold check is inconsistent. The `purchases.total`
  //    column = puramt + GST = grand total (with GST). So:
  //      • flag_wgst=true  → prior=SUM(total), current=grandTotal       ✓
  //      • flag_wgst=false → prior=SUM(amount), current=totalPuramt    ✓
  //    (`purchases.amount` is stored as the pre-GST puramt subtotal.)
  const cr = String(firstLot.cr || '').trim();
  const gstinPrefixed = cr.toUpperCase().startsWith('GSTIN.') ? cr : ('GSTIN.' + cr);
  const gstinBare     = cr.toUpperCase().startsWith('GSTIN.') ? cr.substring(6) : cr;
  const priorAmountCol = cfg.flag_wgst ? 'total' : 'amount';
  const priorPurchases = db.get(
    `SELECT COALESCE(SUM(${priorAmountCol}),0) as total
       FROM purchases
      WHERE (gstin = ? OR gstin = ?) AND date >= ?`,
    [gstinPrefixed, gstinBare, cfg.season_start || '2026-04-01']
  );
  const tdsAmount = cfg.flag_tds_purchase 
    ? calculateTDS(cfg.flag_wgst ? grandTotal : totalPuramt, priorPurchases ? priorPurchases.total : 0, cfg)
    : 0;
  const invoiceAmount = grandTotal - tdsAmount;

  return {
    seller: { name: firstLot.name, address: firstLot.padd, place: firstLot.ppla, 
              cr: firstLot.cr, pan: firstLot.pan, state: firstLot.pstate },
    lineItems,
    summary: {
      totalQty, totalBags, totalPuramt, totalCgst, totalSgst, totalIgst,
      roundDiff, grandTotal, tdsAmount, invoiceAmount, isInter
    }
  };
}

/**
 * Generate payment summary for sellers (PAYCHECK.PRG equivalent)
 */
function getPaymentSummary(db, auctionId, state, cfg) {
  // The "discount" column is the sum of two parts per seller per auction:
  //   1. Per-lot computed discount (lots.refund in e-Trade, lots.advance in
  //      auction mode) — based on discount_pct × days × puramt
  //   2. Per-seller debit notes for this auction — manual adjustments
  //      (e.g., quality complaints, settlement deductions). Joined by
  //      seller name + auction ano so we sum all debit_notes that apply.
  // Total payable already accounts for these via balance recalc, but the
  // displayed "Discount" column needs the COMBINED figure so the user
  // sees both the policy discount and any manual adjustments.
  // First fetch the per-lot summary (e-Trade only — discount column is
  // always lots.refund). Also pulls per-seller GST sums so the Payments
  // tab can show a "GST 5% (CGST+SGST+IGST)" column next to the discount.
  // GROUP BY l.name (only). Earlier this was `GROUP BY l.name, l.cr`
  // which split a seller across multiple payment rows whenever their
  // lots carried inconsistent `cr` values (a real data state when a
  // dealer's GSTIN was edited mid-trade or imported in different
  // formats). The DN map is keyed by name alone, so the same DN total
  // got applied to each row → doubled (or N-tupled) discount on the
  // Payments tab. Grouping by name only collapses the rows back into
  // one per seller. MAX(l.cr) keeps the most-populated `cr` for the
  // returned payload.
  let query = `SELECT l.name, MAX(l.cr) AS cr,
    SUM(l.qty) as total_qty, SUM(l.amount) as total_amount,
    SUM(l.pqty) as total_pqty, SUM(l.prate) as avg_prate,
    SUM(l.puramt) as total_puramt,
    SUM(l.refund) as lot_discount,
    SUM(COALESCE(l.cgst,0)) as total_cgst,
    SUM(COALESCE(l.sgst,0)) as total_sgst,
    SUM(COALESCE(l.igst,0)) as total_igst,
    SUM(l.balance) as total_payable,
    COUNT(*) as lot_count,
    GROUP_CONCAT(DISTINCT l.bank_id) AS bank_ids,
    COUNT(l.bank_id) AS bank_lot_count,
    MAX(l.trader_id) AS trader_id,
    MAX(l.state) AS state
    FROM lots l WHERE l.auction_id = ? AND l.amount > 0`;
  const params = [auctionId];
  if (state) { query += ' AND l.state = ?'; params.push(state); }
  query += ' GROUP BY l.name ORDER BY MAX(l.state), l.name';
  const sellers = db.all(query, params);

  // How many bank accounts each seller (trader) maintains. Drives the
  // "multiple banks" badge: a mix of tagged + untagged lots is only
  // ambiguous when the seller actually has >1 account on file — with a
  // single account the untagged lots can only route to that same account.
  const bankCountByTraderId = {};
  try {
    const counts = db.all(
      'SELECT trader_id, COUNT(*) AS n FROM trader_banks GROUP BY trader_id'
    );
    for (const c of counts) bankCountByTraderId[c.trader_id] = Number(c.n) || 0;
  } catch (_) { /* trader_banks may not exist on partial migrations */ }

  // Fetch this auction's identifier (ano) so we can match debit_notes.
  // Debit notes are keyed by ano + seller name (no FK to auctions.id),
  // mirroring the legacy FoxPro flow.
  const auction = db.get('SELECT ano FROM auctions WHERE id = ?', [auctionId]);
  const ano = auction ? auction.ano : null;
  // Build a name → debit_note total map for fast lookup
  const debitMap = {};
  if (ano) {
    const debits = db.all(
      'SELECT name, SUM(amount) as total FROM debit_notes WHERE ano = ? GROUP BY name',
      [ano]
    );
    for (const d of debits) debitMap[d.name] = Number(d.total) || 0;
  }
  // Per-seller TDS — LOT-WISE: a flat rate on the seller's purchase amount
  // (no ₹50-lakh cumulative threshold), so it reads as "rate × amount" per
  // lot rather than the threshold-reduced Section 194Q figure. Computed below
  // per seller via lotwisePurchaseTds(total_puramt, cr) so it doesn't depend
  // on whether a purchase invoice was generated. Payable is netted by it so
  // it reads PurAmt − Discount − GST 5% − TDS. (The purchase INVOICE keeps the
  // 194Q threshold figure — only the Payments surfaces use this lot-wise one.)
  // Merge: total_discount per seller = ONE of two sources (never both):
  //
  //   - When debit notes exist for this seller in this trade → DN total
  //     IS the authoritative discount. The DN was generated using the
  //     same `discount_pct × days × puramt` formula as `lots.refund`,
  //     so summing both was double-counting the same money. Furthermore,
  //     the user may have manually edited the DN amount after generation
  //     (via the Edit Debit Note flow), in which case the DN value is
  //     the current source of truth — `lots.refund` is stale.
  //
  //   - When no DN exists yet → fall back to the per-lot computed
  //     `lots.refund` so the Payments tab shows what the seller WILL
  //     be discounted once DNs are generated.
  //
  // Earlier code did `lotDisc + manualDisc` unconditionally — every
  // trade with DNs generated showed double the actual discount, and
  // payable was off by that amount.
  // GST rate to apply when "Discount includes GST" is on. Falls back to
  // gst_service for older DBs that didn't have the dedicated key. We
  // compute the Payments-tab tax live from the authoritative discount
  // instead of summing whatever's stamped in lots.cgst/sgst/igst —
  // those stamps go stale if the user edited discount_gst /
  // flag_disc_gst after lots were calculated, which manifested on
  // screen as a doubled (or otherwise wrong) GST column even though
  // the Discount value was current. Live derivation keeps tax in
  // lockstep with discount and the DN totals.
  const flagDiscGst = String(cfg.flag_disc_gst || '').toLowerCase() === 'true' || cfg.flag_disc_gst === true;
  const discGstRate = Number(cfg.discount_gst);
  const discRate = Number.isFinite(discGstRate) && discGstRate >= 0
    ? discGstRate
    : (Number(cfg.gst_service) || 18);
  const halfDiscRate = discRate / 2;
  const companyStateCode = String(cfg.business_state || '').toUpperCase() === 'KERALA' ? '32' : '33';
  // r2 helper — same Excel-compatible rounding the rest of the calc uses.
  const r2live = (n) => round2(n);

  return sellers.map(s => {
    const lotDisc = Number(s.lot_discount) || 0;
    const manualDisc = Number(debitMap[s.name]) || 0;
    // Authoritative discount: DN total when present, otherwise lot total.
    const totalDiscount = manualDisc > 0 ? manualDisc : lotDisc;
    // Live GST on the discount. Intra/inter classification mirrors the
    // seller's own GSTIN state (cr) — same rule the DN generator uses.
    let cgst = 0, sgst = 0, igst = 0;
    if (flagDiscGst && totalDiscount > 0) {
      let stateCode = '';
      let cr = String(s.cr || '').trim().toUpperCase();
      if (cr.startsWith('GSTIN.')) cr = cr.slice(6);
      else if (cr.startsWith('GSTIN')) cr = cr.slice(5);
      if (/^\d{2}/.test(cr)) stateCode = cr.slice(0, 2);
      const isInter = !!stateCode && stateCode !== companyStateCode;
      if (isInter) {
        igst = r2live(totalDiscount * discRate / 100);
      } else {
        cgst = r2live(totalDiscount * halfDiscRate / 100);
        sgst = r2live(totalDiscount * halfDiscRate / 100);
      }
    }
    return {
      ...s,
      total_discount: totalDiscount,
      total_cgst: cgst,
      total_sgst: sgst,
      total_igst: igst,
      total_tax: r2live(cgst + sgst + igst),
      // Lot-wise TDS = rate × this seller's total purchase amount (flat, no
      // threshold). Spreading it ∝ puramt in the modal/statement makes each
      // lot show rate × its own amount.
      total_tds: lotwisePurchaseTds(s.total_puramt, s.cr, cfg),
      // Payable: lots.balance was computed BEFORE DNs existed, using
      // lots.refund as the discount. So:
      //   - When DNs exist and equal lot refunds → balance is correct
      //   - When DNs exist and DIFFER from lot refunds (user edited) →
      //     adjust by the delta so payable reflects the current DN
      //   - When no DNs → balance is already correct
      // Then subtract TDS so Payable = PurAmt − Discount − GST 5% − TDS.
      total_payable: r2live((manualDisc > 0
        ? (Number(s.total_payable) || 0) - (manualDisc - lotDisc)
        : (Number(s.total_payable) || 0)) - lotwisePurchaseTds(s.total_puramt, s.cr, cfg)),
      // True when this seller's lots point at more than one bank account
      // (or a mix of tagged + untagged AND the seller has >1 account on
      // file). Drives the "multiple banks" badge on the Payments table so
      // the user knows to export each account's lots separately via the lot
      // picker. A mix of tagged + untagged lots is NOT ambiguous when the
      // seller has a single bank account — the untagged lots can only route
      // to that one account — so it's suppressed there (fix A).
      multipleBanks: (() => {
        const ids = String(s.bank_ids || '').split(',')
          .map(x => x.trim()).filter(x => x !== '' && x !== 'null');
        const untagged = Number(s.lot_count || 0) > Number(s.bank_lot_count || 0);
        const distinct = new Set(ids).size;
        const bankCount = Number(bankCountByTraderId[s.trader_id]) || 0;
        return distinct > 1 || (distinct >= 1 && untagged && bankCount > 1);
      })(),
    };
  });
}

/**
 * Format a raw GROUP_CONCAT(lot_no) string into a clean, de-duped,
 * numerically-sorted comma list for the bank payment REMARKS column
 * (e.g. "12,13,14"). Lots that look numeric sort ascending; mixed/text
 * lot numbers fall back to lexical order. Returns '' when there are no
 * lots so callers can omit the suffix entirely.
 */
function formatLotList(raw) {
  if (!raw) return '';
  const uniq = [...new Set(
    String(raw).split(',').map(s => s.trim()).filter(Boolean)
  )];
  const allNumeric = uniq.every(x => /^\d+$/.test(x));
  uniq.sort(allNumeric ? (a, b) => Number(a) - Number(b) : undefined);
  return uniq.join(',');
}

/**
 * Generate bank payment data (BANKPAY.PRG — RTGS/NEFT format).
 * Used by both the "after discount" Bank Payment export (default) and
 * the "Bank Payment (Before)" export when `opts.before === true`.
 */
function getBankPaymentData(db, auctionId, cfg, opts) {
  opts = opts || {};
  const useBefore = !!opts.before;
  // Bank Payment lists every seller in the trade with a non-zero
  // payable (or non-zero pre-discount amount in 'before' mode) — both
  // registered dealers AND unregistered (URD/agriculturist) farmers.
  // The earlier WHERE clause filtered to URD-only by excluding rows
  // whose `cr` looked like a GSTIN. That came from the legacy FoxPro
  // BANKPAY.PRG which only handled farmers — but the e-Trade flow pays
  // every seller via RTGS/NEFT, so all sellers must be included.
  // Result was: registered dealers had IFSC + acctnum on file, but the
  // SQL excluded them and returned empty rows, so the export was blank.
  //
  // Bank details come from `traders` (single-bank legacy) or
  // `trader_banks` (multi-bank). The LEFT JOIN to traders pulls
  // address/IFSC; we then COALESCE with trader_banks default for
  // sellers who maintain multiple bank accounts.
  const payments = db.all(
    // GROUP BY l.name (only) — same fix as getPaymentSummary. Splitting
    // by `cr` produced duplicate bank-payment rows whenever a seller's
    // lots held inconsistent GSTIN values, leading to NEFT files with
    // the dealer listed twice for partial amounts.
    // JOIN trader by lots.trader_id (FK), not by name. Joining by name
    // multiplied each lot row by the number of traders sharing that
    // name (multi-branch sellers / accidental dupes), then SUM(puramt)
    // etc. summed those duplicates → inflated payable. GROUP BY name
    // alone wasn't enough; the fan-out happened BEFORE the aggregate.
    `SELECT MAX(l.state) AS state, l.name, MAX(l.cr) AS cr,
      SUM(l.puramt) as puramt, SUM(l.refund) as advance, SUM(l.balance) as payable,
      GROUP_CONCAT(l.lot_no) as lot_nos,
      MAX(t.id) AS trader_id,
      MAX(t.ifsc) AS t_ifsc, MAX(t.acctnum) AS t_acctnum, MAX(t.holder_name) AS t_holder,
      MAX(t.padd) AS padd, MAX(t.ppla) AS ppla, MAX(t.pin) AS pin,
      -- Per-lot bank routing: distinct non-null bank_ids across this
      -- seller's payable lots, plus counts so we can tell whether they
      -- ALL share one account (single → use it) or differ (mixed → keep
      -- the default account and flag multipleBanks for the UI).
      GROUP_CONCAT(DISTINCT l.bank_id) AS bank_ids,
      COUNT(*) AS lot_count,
      COUNT(l.bank_id) AS bank_lot_count
    FROM lots l
    LEFT JOIN traders t ON t.id = l.trader_id
    WHERE l.auction_id = ? AND l.amount > 0
      AND (l.paid IS NULL OR l.paid = '')
    GROUP BY l.name
    ORDER BY MAX(l.state), l.name`,
    [auctionId]
  );

  // Per-seller bank-details fallback chain:
  //   1. trader_banks default (is_default=1) — picks the explicitly
  //      flagged primary account when the seller has multiple banks
  //   2. trader_banks first row — when no default flagged
  //   3. traders.ifsc/acctnum — legacy single-bank
  // Pre-fetch all default banks once (cheaper than per-seller query).
  const bankByTraderId = {};
  // Also index every bank row by its own id so per-lot bank_id routing can
  // resolve the exact account a seller's lots were tagged with.
  const bankById = {};
  // Count of accounts per trader — a mix of tagged + untagged lots only
  // flags multipleBanks when the seller has >1 account on file (fix A).
  const bankCountByTraderId = {};
  try {
    const banks = db.all(`
      SELECT trader_id, ifsc, acctnum, holder_name, bank_name, is_default, id
        FROM trader_banks
       ORDER BY trader_id, is_default DESC, id ASC
    `);
    for (const b of banks) {
      // First row per trader_id wins (already sorted by is_default DESC).
      if (bankByTraderId[b.trader_id] == null) bankByTraderId[b.trader_id] = b;
      bankById[b.id] = b;
      bankCountByTraderId[b.trader_id] = (bankCountByTraderId[b.trader_id] || 0) + 1;
    }
  } catch (_) { /* trader_banks may not exist on partial migrations */ }

  const auction = db.get('SELECT * FROM auctions WHERE id = ?', [auctionId]);
  const roundAmounts = cfg.flag_round;

  return payments.map(p => {
    // 'before' uses puramt — pre-discount, useful when paying suppliers
    // before the deduction policy is applied. 'after' (default) uses
    // payable = puramt − discount − GST − TDS, netted by the SAME lot-wise
    // TDS the Payments tab shows so the NEFT amount matches the screen.
    const tds = useBefore ? 0 : lotwisePurchaseTds(p.puramt, p.cr, cfg);
    const rawAmount = (useBefore ? (p.puramt || 0) : (p.payable || 0)) - tds;
    const amount = roundAmounts ? round0(rawAmount) : rawAmount;
    const tb = p.trader_id != null ? bankByTraderId[p.trader_id] : null;
    // Per-lot bank routing. Distinct non-null bank_ids tagged on this
    // seller's payable lots:
    const distinctBankIds = String(p.bank_ids || '')
      .split(',').map(s => s.trim()).filter(s => s !== '' && s !== 'null')
      .map(Number).filter(Number.isFinite);
    const hasUntagged = Number(p.lot_count || 0) > Number(p.bank_lot_count || 0);
    // Use the lot-tagged account ONLY when every payable lot points at the
    // same single account (no untagged lots). Otherwise keep the seller's
    // default account and flag `multipleBanks` so the UI can warn the user
    // to export each account's lots separately via the lot picker.
    const lotBank = (distinctBankIds.length === 1 && !hasUntagged)
      ? bankById[distinctBankIds[0]] : null;
    const bankCount = (p.trader_id != null && bankCountByTraderId[p.trader_id]) || 0;
    const multipleBanks = distinctBankIds.length > 1
      || (distinctBankIds.length >= 1 && hasUntagged && bankCount > 1);
    const ifsc      = (lotBank && lotBank.ifsc)        || (tb && tb.ifsc)        || p.t_ifsc    || '';
    const acctnum   = (lotBank && lotBank.acctnum)     || (tb && tb.acctnum)     || p.t_acctnum || '';
    const holderNm  = (lotBank && lotBank.holder_name) || (tb && tb.holder_name) || p.t_holder  || p.name;
    // Lots this seller's payment covers — surfaced in REMARKS so the
    // beneficiary can reconcile the credit against the specific lots.
    const lots = formatLotList(p.lot_nos);
    return {
      // Seller name preserved on the row so callers can filter by the
      // same key the Payments tab UI uses (the ticked checkbox value).
      // beneficiaryName below can diverge from this (it tracks the bank
      // account holder, which may be a different person/entity) so it's
      // not safe to filter against beneficiaryName.
      name: p.name,
      transactionType: rawAmount >= 200000 ? 'RTGS' : 'NEFT',
      ifsc,
      accountNo: acctnum,
      beneficiaryName: holderNm,
      address1: p.padd || '',
      address2: p.ppla || '',
      pin: p.pin || '',
      amount,
      lots,
      remarks: `${auction ? auction.ano : ''} ${p.name} PAYMENT ${rawAmount.toFixed(2)} Credited${lots ? ` for lot${lots.includes(',') ? 's' : ''} ${lots}` : ''}`,
      holderName: holderNm,
      // True when this seller's lots point at more than one bank account
      // (or a mix of tagged + untagged). The row still pays a single
      // account (the default); the Payments UI shows a badge prompting the
      // user to export each account's lots separately via the lot picker.
      multipleBanks,
    };
  });
}

/**
 * TDS return data (TDSRETU.PRG equivalent)
 */
function getTDSReturnData(db, fromDate, toDate, orderBy) {
  const order = orderBy === 'party' ? 'name' : 'date, invo';
  // PAN extraction. The gstin column holds either:
  //   "GSTIN.32AAHCE4551A1Z8" (21 chars, with prefix — most common)
  //   "32AAHCE4551A1Z8"       (15 chars, bare GSTIN)
  // Strip the optional "GSTIN." prefix first, then take chars 3-12 of
  // the bare GSTIN to get the 10-char PAN ("AAHCE4551A").
  return db.all(
    `SELECT invo as invoice, date, name,
      SUBSTR(
        CASE WHEN UPPER(SUBSTR(COALESCE(gstin,''), 1, 6)) = 'GSTIN.'
             THEN SUBSTR(gstin, 7)
             ELSE COALESCE(gstin,'') END,
        3, 10
      ) as pan,
      amount as assess_value, tds
    FROM purchases
    WHERE date BETWEEN ? AND ? AND tds > 0
    ORDER BY ${order}`,
    [fromDate, toDate]
  );
}

/**
 * Build Agriculturist Bill of Supply (GSTKBILP/GSTBILP equivalent)
 * For sellers WITHOUT GSTIN — agricultural produce from farmers.
 * No GST charged (exempt/reverse charge).
 * 
 * Returns: { seller, lineItems, summary } if successful
 *          { error, detail } object if no data (to help debug)
 */
function buildAgriBill(db, auctionId, sellerName, cfg) {
  const trimmedName = String(sellerName || '').trim();
  if (!trimmedName) return { error: 'Seller name is empty' };

  // First check: any lots at all for this seller (case-insensitive)?
  const allLots = db.all(
    `SELECT * FROM lots WHERE auction_id = ? AND UPPER(TRIM(name)) = UPPER(?) ORDER BY lot_no`,
    [auctionId, trimmedName]
  );
  
  if (!allLots.length) {
    return { error: `No lots found for seller "${trimmedName}" in this auction. Check the exact spelling.` };
  }

  // Check if any have GSTIN — those aren't eligible for Bills of Supply
  const withGstin = allLots.filter(l => l.cr && l.cr.toUpperCase().startsWith('GSTIN'));
  const withoutGstin = allLots.filter(l => !l.cr || !l.cr.toUpperCase().startsWith('GSTIN'));
  
  if (withGstin.length && !withoutGstin.length) {
    return { error: `Seller "${trimmedName}" has GSTIN (${withGstin[0].cr}). Use Generate Purchase Invoice instead — Bills of Supply are only for agriculturists without GSTIN.` };
  }

  // Filter to agri-eligible lots with amount > 0
  const lots = withoutGstin.filter(l => (l.amount || 0) > 0);
  
  if (!lots.length) {
    if (withoutGstin.length) {
      return { error: `Seller "${trimmedName}" has ${withoutGstin.length} lot(s) but none have amount > 0. Set prices on the lots first (or click Calculate All).` };
    }
    return { error: `No eligible lots for "${trimmedName}"` };
  }

  let totalQty = 0, totalPuramt = 0;
  const lineItems = [];

  for (const lot of lots) {
    totalQty += lot.pqty || lot.qty;
    totalPuramt += lot.puramt || 0;
    lineItems.push({
      lot: lot.lot_no, qty: lot.qty, pqty: lot.pqty,
      price: lot.price, prate: lot.prate,
      amount: lot.amount, puramt: lot.puramt,
      com: lot.com, sertax: lot.sertax
    });
  }

  const firstLot = lots[0];
  const netAmount = round0(totalPuramt);
  const roundDiff = cfg.flag_round ? netAmount - totalPuramt : 0;

  return {
    seller: {
      name: firstLot.name,
      address: firstLot.padd,
      place: firstLot.ppla,
      pin: firstLot.ppin,
      state: firstLot.pstate,
      st_code: firstLot.pst_code,
      cr: firstLot.cr,
      pan: firstLot.pan,
      aadhar: firstLot.aadhar,
      tel: firstLot.tel,
    },
    lineItems,
    summary: {
      totalQty, totalPuramt, 
      roundDiff, netAmount,
      cgst: 0, sgst: 0, igst: 0,
      tax: 0
    }
  };
}

/**
 * List agri-eligible sellers for an auction
 * (sellers without GSTIN who have lots with amount > 0)
 */
function listAgriSellers(db, auctionId) {
  // An "agri seller" is one without a GSTIN. Reject both prefixed
  // ("GSTIN.<gstin>") and bare ("<gstin>") forms — anything else (empty,
  // CR codes, plain text) qualifies.
  return db.all(
    `SELECT name, COUNT(*) as lot_count, SUM(qty) as total_qty, SUM(amount) as total_amount
     FROM lots 
     WHERE auction_id = ? 
       AND (cr IS NULL OR cr = ''
            OR (UPPER(cr) NOT LIKE 'GSTIN%' AND cr NOT GLOB '[0-9][0-9]*'))
       AND amount > 0
     GROUP BY name
     ORDER BY name`,
    [auctionId]
  );
}

// Display date formatter for journal exports — honours the user's
// Settings → Display → Date format choice via the shared module.
const { fmtDate: _ddmmyyyy } = require('./date-format');

/**
 * Sales Journal (JOUR.PRG)
 * Trade-wise sales invoice register. Filters invoices by auction id
 * (resolved via auctions.ano so old invoices with a NULL auction_id
 * still match by ano). Dates rendered dd/mm/yyyy.
 */
function getSalesJournal(db, auctionId, saleType) {
  const auction = db.get('SELECT id, ano FROM auctions WHERE id = ?', [auctionId]);
  if (!auction) return [];
  let query = `SELECT date, sale, invo, buyer, buyer1, gstin, place,
      bag, qty, amount as cardamom, gunny, pava_hc as transport, ins as insurance,
      cgst, sgst, igst, tcs, rund, tot as total
    FROM invoices WHERE (auction_id = ? OR ano = ?)`;
  const params = [auction.id, auction.ano];
  if (saleType) { query += ' AND sale = ?'; params.push(saleType); }
  query += ' ORDER BY date, sale, invo';
  const rows = db.all(query, params);
  return rows.map(r => ({ ...r, date: _ddmmyyyy(r.date) }));
}

/**
 * Purchase Journal (PUJOUR.PRG / PPUJOUR.PRG)
 * Trade-wise purchase invoice register. Dates rendered dd/mm/yyyy.
 * type: 'dealer' (registered) or 'agri' (agriculturist bills)
 */
function getPurchaseJournal(db, auctionId, type) {
  const auction = db.get('SELECT id, ano FROM auctions WHERE id = ?', [auctionId]);
  if (!auction) return [];
  if (type === 'agri') {
    // bills table only has `ano`, not auction_id, so match by ano alone.
    const rows = db.all(
      `SELECT date, bil as bill_no, name, add_line as address, pla as place, pstate as state,
        crr as cr, pan, qty, cost, igst, net
      FROM bills WHERE ano = ? ORDER BY date, bil`,
      [auction.ano]
    );
    return rows.map(r => ({ ...r, date: _ddmmyyyy(r.date) }));
  }
  // Dealer purchases — match either by auction_id (newer rows) or ano (legacy).
  const rows = db.all(
    `SELECT date, invo as invoice_no, name, add_line as address, place, state,
      gstin, qty, amount, cgst, sgst, igst, rund, total, tds
    FROM purchases WHERE (auction_id = ? OR ano = ?) ORDER BY date, invo`,
    [auction.id, auction.ano]
  );
  return rows.map(r => ({ ...r, date: _ddmmyyyy(r.date) }));
}

/**
 * Purchase Register (lot-wise)
 * One row PER LOT — the seller-side purchase detail. Unlike the Purchase
 * Journal (one row per dealer invoice / agri bill), this is the raw lot
 * ledger: STATE, TNO, DATE, LOT, BRANCH, NAME, PLACE, GSTIN, BAG, QTY,
 * PRICE, AMOUNT, PQTY, PRATE, PURAMT, DISCOUNT, GST5, PAYABLE.
 *
 * DISCOUNT = refund, GST5 = stored GST-on-discount (`advance`), PAYABLE =
 * balance (GST already netted) — see [[payment-field-semantics]]. In
 * auction mode `advance` is the discount, so GST5 → 0.
 *
 * Scope: a specific trade (opts.auctionId) OR a date range across trades
 * (opts.from/opts.to over the auction date). Trade wins when both given.
 */
function getPurchaseRegister(db, opts = {}) {
  const mode = String(opts.mode || 'e-Trade').toLowerCase();
  const discountCol = (mode === 'auction') ? 'advance' : 'refund';
  const gstCol = (mode === 'auction') ? '0' : 'advance';
  let q = `SELECT l.state AS state, a.ano AS tno, a.date AS date, l.lot_no AS lot,
      l.branch AS branch, l.name AS name, l.ppla AS place, l.cr AS gstin,
      l.bags AS bag, l.qty AS qty, l.price AS price, l.amount AS amount,
      l.pqty AS pqty, l.prate AS prate, l.puramt AS puramt,
      l.${discountCol} AS discount, l.${gstCol} AS gst5, l.balance AS payable
    FROM lots l JOIN auctions a ON a.id = l.auction_id
    WHERE l.amount > 0`;
  const params = [];
  if (opts.auctionId) { q += ' AND l.auction_id = ?'; params.push(opts.auctionId); }
  else if (opts.from && opts.to) { q += ' AND a.date BETWEEN ? AND ?'; params.push(opts.from, opts.to); }
  q += ' ORDER BY l.state, a.ano, CAST(l.lot_no AS INTEGER), l.lot_no';
  const rows = db.all(q, params);
  return rows.map(r => ({ ...r, date: _ddmmyyyy(r.date) }));
}

/**
 * Sales Register (invoice-wise)
 * One row PER INVOICE: STATE, TNO, DATE, SALE, INVO, TRADERNAME, BIDDER,
 * BAG, QTY, AMOUNT, LORRY, GUNNY, IGST, CGST, SGST, INS, INVAMT.
 * LORRY = freight charge (pava_hc); INVAMT = invoice grand total (tot).
 *
 * Scope: a specific trade (matched by auction_id OR ano for legacy rows)
 * OR a date range across trades. Optional saleType filter.
 */
function getSalesRegister(db, opts = {}) {
  let q = `SELECT i.state AS state, i.ano AS tno, i.date AS date, i.sale AS sale,
      i.invo AS invo, i.buyer1 AS tradername, i.buyer AS bidder,
      i.bag AS bag, i.qty AS qty, i.amount AS amount,
      i.pava_hc AS lorry, i.gunny AS gunny, i.igst AS igst, i.cgst AS cgst,
      i.sgst AS sgst, i.ins AS ins, i.tot AS invamt
    FROM invoices i`;
  const params = [];
  const where = [];
  if (opts.auctionId) {
    const a = db.get('SELECT id, ano FROM auctions WHERE id = ?', [opts.auctionId]);
    if (a) { where.push('(i.auction_id = ? OR i.ano = ?)'); params.push(a.id, a.ano); }
    else { where.push('i.auction_id = ?'); params.push(opts.auctionId); }
  } else if (opts.from && opts.to) {
    where.push('i.date BETWEEN ? AND ?'); params.push(opts.from, opts.to);
  }
  if (opts.saleType) { where.push('i.sale = ?'); params.push(opts.saleType); }
  if (where.length) q += ' WHERE ' + where.join(' AND ');
  q += ' ORDER BY i.state, i.ano, i.date, i.sale, i.invo';
  const rows = db.all(q, params);
  return rows.map(r => ({ ...r, date: _ddmmyyyy(r.date) }));
}

module.exports = {
  calculateLot,
  calculateTDS,
  calculateTCS,
  buildSalesInvoice,
  buildPurchaseInvoice,
  buildAgriBill,
  listAgriSellers,
  getPaymentSummary,
  getBankPaymentData,
  formatLotList,
  getTDSReturnData,
  getSalesJournal,
  getPurchaseJournal,
  getPurchaseRegister,
  getSalesRegister,
  gstinStateCode,
  round2,
  round0,
};
