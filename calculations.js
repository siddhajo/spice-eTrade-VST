/**
 * calculations.js — Core business logic
 * Replaces: GENERATE.PRG, parts of GSTKBILT/GSTKBILP/GSTBILP/PAYCHECK
 */

const { getSettingsFlat, getGSTRates } = require('./company-config');

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
 * e-Trade only build: single-company calculation. The dual ISP/ASP planter
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
  const gradeStr = String(lot.grade || '').trim();
  const deduction = (gradeStr === '2')
    ? Number(cfg.deduction2 || 0)
    : Number(cfg.deduction1 || 0);
  const rawRate = (lot.price || 0) * (1 - deduction / 100);
  result.prate = Math.round(rawRate);
  // PurAmt = P_Qty × P_Rate (sample refund INCLUDED — direct purchase)
  result.puramt = Math.round(result.pqty * result.prate * 100) / 100;
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

  // GST on the Discount uses Service Rate (the discount is treated as a
  // credit/finance service when flag_disc_gst is ON).
  const gstServiceRate = Number(cfg.gst_service) || 18;
  const halfRate = gstServiceRate / 2;

  result.cgst = 0;
  result.sgst = 0;
  result.igst = 0;

  // e-Trade discount: round(PurAmt / 1000 × days × discount%) — nearest
  // rupee, half-up. result.refund holds the Discount amount.
  const days    = Number(cfg.discount_days) || 0;
  const discPct = Number(cfg.discount_pct)  || 0;
  result.refund = Math.round((result.puramt / 1000) * days * discPct);

  // GST on the Discount only when flag_disc_gst is ON.
  if (cfg.flag_disc_gst && result.refund > 0) {
    if (isIntra) {
      result.cgst = Math.round(result.refund * halfRate / 100 * 100) / 100;
      result.sgst = Math.round(result.refund * halfRate / 100 * 100) / 100;
    } else {
      result.igst = Math.round(result.refund * gstServiceRate / 100 * 100) / 100;
    }
  }

  // advance = GST only (informational; e-Trade has no commission/handling)
  result.advance = result.cgst + result.sgst + result.igst;

  // Payable = PurAmt − Discount − GST-on-Discount
  const totalDeductions = result.refund + result.cgst + result.sgst + result.igst;
  result.balance = Math.round((result.puramt - totalDeductions) * 100) / 100;

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
  // Don't filter by sale — we're ASSIGNING the sale type now
  const lots = db.all(
    `SELECT * FROM lots WHERE auction_id = ? AND buyer = ? AND amount > 0 
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
  //   L (Local)       → local_transport / local_insurance config keys
  //   I (Inter-state) → transport / insurance config keys
  //   E (Export)      → zero (buyer covers freight; matches the
  //                      hideTransportInsurance render rule in invoice-pdf.js)
  // Anything else (legacy / blank) is treated as 'L' for safety.
  const st = String(saleType || '').toUpperCase();
  const isExport = (st === 'E');
  const isInter  = (st === 'I');
  const transportRate = isExport ? 0 : (isInter
    ? pickRate(cfg.transport, 2.5)
    : pickRate(cfg.local_transport, cfg.transport, 2.5));
  const insuranceRate = isExport ? 0 : (isInter
    ? pickRate(cfg.insurance, 0.75)
    : pickRate(cfg.local_insurance, cfg.insurance, 0.75));

  // Transport: ₹/kg (qty × rate)
  const transportCost = Math.round(totalQty * transportRate * 100) / 100;

  // Insurance: per ₹1000 of (cardamom + gunny + GST on those)
  //   insurance = ((cardamom_amount + gunny_cost) × (1 + gstGoods/100)) / 1000 × rate
  const subtotalGoods = totalAmount + gunnyCost;
  const gstOnGoods = subtotalGoods * gstGoods / 100;
  const insuranceCost = Math.round((subtotalGoods + gstOnGoods) / 1000 * insuranceRate * 100) / 100;

  // Taxable value = cardamom + gunny + transport + insurance
  const taxableValue = subtotalGoods + transportCost + insuranceCost;

  // All four components get the SAME gstGoods rate (per user confirmation).
  let cgst = 0, sgst = 0, igst = 0;
  if (isInterState) {
    igst = Math.round(taxableValue * gstGoods / 100 * 100) / 100;
  } else {
    cgst = Math.round(taxableValue * (gstGoods / 2) / 100 * 100) / 100;
    sgst = Math.round(taxableValue * (gstGoods / 2) / 100 * 100) / 100;
  }

  const totalBeforeRound = taxableValue + cgst + sgst + igst;
  const roundDiff = Math.round(totalBeforeRound) - totalBeforeRound;
  const grandTotal = Math.round(totalBeforeRound);

  return {
    buyer: buyer || {},
    saleType,
    lineItems,
    summary: {
      totalQty, totalBags, totalAmount,
      gunnyCost, transportCost, insuranceCost,
      taxableValue, cgst, sgst, igst,
      roundDiff, grandTotal,
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

    const rcgst = isInter ? 0 : Math.round(puramt * (gstGoods / 2) / 100 * 100) / 100;
    const rsgst = isInter ? 0 : Math.round(puramt * (gstGoods / 2) / 100 * 100) / 100;
    const rigst = isInter ? Math.round(puramt * gstGoods / 100 * 100) / 100 : 0;

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
  const roundDiff = Math.round(totalBeforeRound) - totalBeforeRound;
  const grandTotal = Math.round(totalBeforeRound);

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
  let query = `SELECT l.name, l.cr,
    SUM(l.qty) as total_qty, SUM(l.amount) as total_amount,
    SUM(l.pqty) as total_pqty, SUM(l.prate) as avg_prate,
    SUM(l.puramt) as total_puramt,
    SUM(l.refund) as lot_discount,
    SUM(COALESCE(l.cgst,0)) as total_cgst,
    SUM(COALESCE(l.sgst,0)) as total_sgst,
    SUM(COALESCE(l.igst,0)) as total_igst,
    SUM(l.balance) as total_payable,
    COUNT(*) as lot_count
    FROM lots l WHERE l.auction_id = ? AND l.amount > 0`;
  const params = [auctionId];
  if (state) { query += ' AND l.state = ?'; params.push(state); }
  query += ' GROUP BY l.name, l.cr ORDER BY l.state, l.name';
  const sellers = db.all(query, params);

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
  // Merge: total_discount = lot-policy discount + any manual debit notes;
  // total_tax = sum of CGST + SGST + IGST per seller (Payments tab column).
  return sellers.map(s => {
    const lotDisc = Number(s.lot_discount) || 0;
    const manualDisc = Number(debitMap[s.name]) || 0;
    const cgst = Number(s.total_cgst) || 0;
    const sgst = Number(s.total_sgst) || 0;
    const igst = Number(s.total_igst) || 0;
    return {
      ...s,
      total_discount: lotDisc + manualDisc,
      total_tax: cgst + sgst + igst,
      // Subtract the manual debit_notes from payable since balance was
      // computed before debit_notes were added. Lot-policy discount is
      // already factored into balance via puramt - cgst - sgst - igst.
      total_payable: (Number(s.total_payable) || 0) - manualDisc,
    };
  });
}

/**
 * Generate bank payment data (BANKPAY.PRG — RTGS/NEFT format).
 * Used by both the "after discount" Bank Payment export (default) and
 * the "Bank Payment (Before)" export when `opts.before === true`.
 */
function getBankPaymentData(db, auctionId, cfg, opts) {
  opts = opts || {};
  const useBefore = !!opts.before;
  const payments = db.all(
    `SELECT l.state, l.name, l.cr,
      SUM(l.puramt) as puramt, SUM(l.refund) as advance, SUM(l.balance) as payable,
      t.ifsc, t.acctnum, t.padd, t.ppla, t.pin, t.holder_name
    FROM lots l
    LEFT JOIN traders t ON t.name = l.name AND t.cr = l.cr
    WHERE l.auction_id = ? AND l.amount > 0
      AND UPPER(COALESCE(l.cr,'')) NOT LIKE 'GSTIN%'
      AND l.cr NOT GLOB '[0-9][0-9]*'
      AND (l.paid IS NULL OR l.paid = '')
    GROUP BY l.name, l.cr
    ORDER BY l.state, l.name`,
    [auctionId]
  );

  const auction = db.get('SELECT * FROM auctions WHERE id = ?', [auctionId]);
  const roundAmounts = cfg.flag_round;

  return payments.map(p => {
    // 'before' uses puramt — pre-discount, useful when paying suppliers
    // before the deduction policy is applied. 'after' (default) uses
    // payable = puramt − discount − GST.
    const rawAmount = useBefore ? (p.puramt || 0) : (p.payable || 0);
    const amount = roundAmounts ? Math.round(rawAmount) : rawAmount;
    return {
      transactionType: rawAmount >= 200000 ? 'RTGS' : 'NEFT',
      ifsc: p.ifsc || '',
      accountNo: p.acctnum || '',
      beneficiaryName: p.name,
      address1: p.padd || '',
      address2: p.ppla || '',
      pin: p.pin || '',
      amount,
      remarks: `${auction ? auction.ano : ''} ${p.name} PAYMENT ${rawAmount.toFixed(2)} Credited`,
      holderName: p.holder_name || p.name
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
  const roundDiff = cfg.flag_round ? Math.round(totalPuramt) - totalPuramt : 0;
  const netAmount = Math.round(totalPuramt);

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

// Convert YYYY-MM-DD (SQLite date) → dd/mm/yyyy for display in
// journal exports. Falls through unchanged if the input doesn't match.
function _ddmmyyyy(d) {
  if (!d) return '';
  const s = String(d);
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return `${m[3]}/${m[2]}/${m[1]}`;
  return s;
}

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
 * Debit Note calculation
 * For discounts or adjustments against invoices
 */
function buildDebitNote(db, invoiceNo, saleType, discount, cfg) {
  const inv = db.get('SELECT * FROM invoices WHERE invo = ? AND sale = ?', [String(invoiceNo), saleType]);
  if (!inv) return null;

  const gstGoods = cfg.gst_goods || 5;
  const isInter = inv.igst > 0;

  const amount = Math.round(discount * 100) / 100;
  let cgst = 0, sgst = 0, igst = 0;
  
  if (cfg.flag_disc_gst) {
    // Discount amount includes GST — extract it
    const factor = 100 / (100 + gstGoods);
    const taxable = amount * factor;
    if (isInter) igst = Math.round((amount - taxable) * 100) / 100;
    else { 
      const tax = (amount - taxable) / 2;
      cgst = Math.round(tax * 100) / 100;
      sgst = Math.round(tax * 100) / 100;
    }
  } else {
    // Discount is pre-tax — add GST on top
    if (isInter) igst = Math.round(amount * gstGoods / 100 * 100) / 100;
    else {
      cgst = Math.round(amount * (gstGoods / 2) / 100 * 100) / 100;
      sgst = Math.round(amount * (gstGoods / 2) / 100 * 100) / 100;
    }
  }
  
  const total = amount + cgst + sgst + igst;
  return { invoice: inv, amount, cgst, sgst, igst, total };
}

module.exports = {
  calculateLot,
  calculateTDS,
  calculateTCS,
  buildSalesInvoice,
  buildPurchaseInvoice,
  buildAgriBill,
  buildDebitNote,
  listAgriSellers,
  getPaymentSummary,
  getBankPaymentData,
  getTDSReturnData,
  getSalesJournal,
  getPurchaseJournal,
  gstinStateCode,
};
