/**
 * exports.js — All export formats
 * Replaces: EXP.PRG (11 types), TALY.PRG, KOTALLY.PRG, BANKPAY export
 */

const ExcelJS = require('exceljs');
const { collectionXlsx: newCollectionXlsx, tradeReportXlsx } = require('./auction-reports');
const {
  getCompanyHeader,
  writeXlsxCompanyHeader, xlsxNumFmtForHeader,
} = require('./report-formatters');
// Defensive identity resolver — see _company-identity-fallback.js.
// Avoids "getCompanyIdentity is not a function" on partial deploys.
const getCompanyIdentity = require('./_company-identity-fallback').resolve();
// Date formatter honours the user's Settings → Display → Date format
// choice via the shared module. Replaces the ad-hoc
// `String(d).slice(0,10).split('-').reverse().join('/')` pattern that
// was hardcoded to DD/MM/YYYY in several places.
const { fmtDate: fmtUserDate } = require('./date-format');

// Build an XLSX buffer with a unified brand band on top and Indian-format
// numeric columns. `opts.title` is the report title shown in the middle of
// the band; `opts.metaLines` is an array of right-aligned meta strings
// (e.g. ["Trade #3", "15/04/2026"]).
// Reusable XLSX export builder. ALL Excel exports in this app should
// route through this function so they share:
//   - The same three-zone brand band (logo + name | title | meta)
//   - The same column-header look (#E8E4DD fill, thin top/bottom borders,
//     bold 10pt, centered text)
//   - The same Indian-format numFmts via xlsxNumFmtForHeader
//   - The same per-column alignment defaults (right for numeric, center
//     for short-id columns like SL/LOT, left for everything else)
//
// columns[i] shape:
//   { key, header,
//     width:   number,         // optional, default 15
//     align:   'left'|'center'|'right',  // optional, derived from numFmt
//     numFmt:  string,         // optional, overrides xlsxNumFmtForHeader
//   }
//
// opts shape:
//   { db, companyHeader, title, metaLines,    // existing
//     grandTotal: { label, values, fillArgb }, // optional footer row
//     sections:   [{ title, rows }],           // optional grouped layout
//     spacerBetween: true,                      // blank row between groups
//   }
//
// "Grand total" row mirrors the Lorry export's footer: bold 11pt, yellow
// (`#FFF3CD`) fill, double top + bottom borders. Pass `values` keyed by
// column key — only the listed columns get numbers, the rest are blank.
// Set `label` to put a string in any one column (defaults to 'GRAND TOTAL'
// in the first non-numeric column).
async function createExcelBuffer(sheetName, columns, rows, opts) {
  opts = opts || {};
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet(sheetName);

  // Apply column widths up front (the brand band uses these widths too).
  ws.columns = columns.map(c => ({ key: c.key, width: c.width || 15 }));

  // Resolve per-column numFmt + alignment ONCE so we can apply them to
  // data rows and the grand-total row uniformly.
  //
  // Default alignment policy:
  //   - explicit `align` wins
  //   - numeric columns (have a numFmt) → right
  //   - everything else → left
  const colMeta = columns.map(c => {
    const fmt = c.numFmt || xlsxNumFmtForHeader(c.header);
    const align = c.align || (fmt ? 'right' : 'left');
    return { fmt, align };
  });

  // Apply column-level numFmt + alignment FIRST, so any cells we write
  // afterwards (brand band, header row, data rows) can override it via
  // explicit per-cell alignment without being clobbered by a later
  // column.alignment cascade.
  colMeta.forEach((m, i) => {
    const colObj = ws.getColumn(i + 1);
    if (m.fmt) colObj.numFmt = m.fmt;
    colObj.alignment = { horizontal: m.align, vertical: 'middle' };
  });

  // Brand band: company name (row 1) + meta (row 2) + spacer (row 3).
  // `opts.noBrandBand` suppresses it entirely so the sheet starts straight
  // at the column-header row — used for fillable forms (e.g. Price List
  // Before) that buyers print blank and don't want branding/meta on.
  let startRow;
  if (opts.noBrandBand) {
    startRow = 1;
  } else {
    const header = opts.companyHeader || getCompanyHeader(opts.db);
    startRow = writeXlsxCompanyHeader(wb, ws, header, {
      colCount: columns.length,
      metaLines: opts.metaLines || [],
    });
  }

  // Column-header row — explicit per-cell alignment 'center' overrides
  // the column-level left/right cascade.
  const headerRow = ws.getRow(startRow);
  columns.forEach((c, i) => {
    headerRow.getCell(i + 1).value = c.header;
  });
  headerRow.font = { bold: true, size: 10 };
  headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8E4DD' } };
  headerRow.height = 20;
  headerRow.eachCell((cell) => {
    cell.border = { top: { style: 'thin' }, bottom: { style: 'thin' } };
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  });

  // Helper to emit a single data row honouring numeric coercion + per-col align.
  // Rows flagged with `_isSubtotal: true` get distinct styling (bold + light
  // yellow fill + thin top border) so callers can interleave per-group
  // subtotal rows directly in `rows` and have them styled automatically.
  function emitDataRow(rowObj) {
    const dataRow = ws.addRow({});
    columns.forEach((c, i) => {
      let v = rowObj[c.key];
      // Coerce string-numbers to numbers so Excel applies the numFmt.
      if (typeof v === 'string' && v !== '' && !isNaN(Number(v))) {
        const n = Number(v);
        if (!Number.isNaN(n) && colMeta[i].fmt) v = n;
      }
      const cell = dataRow.getCell(i + 1);
      cell.value = v == null ? '' : v;
      // Per-cell alignment guard — vertical:'middle' centers text vertically
      // so rows align consistently regardless of font size differences.
      cell.alignment = { horizontal: colMeta[i].align, vertical: 'middle' };
    });
    if (rowObj && rowObj._isSubtotal) {
      dataRow.font = { bold: true };
      dataRow.eachCell((cell, ci) => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF8E1' } };
        cell.border = { top: { style: 'thin' } };
        const m = colMeta[ci - 1];
        if (m && m.fmt) cell.numFmt = m.fmt;
      });
    }
    return dataRow;
  }

  // ── Section-grouped mode (optional) ──
  // When `opts.sections` is provided, we ignore `rows` and emit each
  // section as: section header (merged, light-green) → its rows. This
  // mirrors the Lorry export's "INTER-STATE SALES" / "INTRA-STATE SALES"
  // structure but is reusable for any grouped data.
  if (Array.isArray(opts.sections) && opts.sections.length) {
    opts.sections.forEach((sec, sIdx) => {
      const titleRow = ws.addRow([sec.title || '']);
      ws.mergeCells(`A${titleRow.number}:${colLetter(columns.length)}${titleRow.number}`);
      titleRow.font = { bold: true, size: 10 };
      titleRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF3CD' } };
      titleRow.alignment = { horizontal: 'left', vertical: 'middle' };
      (sec.rows || []).forEach(emitDataRow);
      if (opts.spacerBetween && sIdx < opts.sections.length - 1) ws.addRow([]);
    });
  } else {
    // Flat mode — original behaviour.
    rows.forEach(emitDataRow);
  }

  // ── Grand total footer (optional) ──
  // Lorry-export style: bold 11pt, yellow `#FFF3CD` fill, double borders.
  // Pass values keyed by column key. Numeric columns get the same numFmt
  // as the data rows for consistent rendering.
  if (opts.grandTotal) {
    const gt = opts.grandTotal;
    const cells = columns.map(c => (gt.values && gt.values[c.key] != null) ? gt.values[c.key] : '');
    // Place label in the first non-numeric column (or column 1 if all
    // columns are numeric). Caller can override by including a label
    // value directly in `gt.values`.
    if (gt.label) {
      const labelIdx = columns.findIndex(c => !colMeta[columns.indexOf(c)].fmt);
      const idx = labelIdx >= 0 ? labelIdx : 0;
      if (cells[idx] === '') cells[idx] = gt.label;
    }
    const gRow = ws.addRow(cells);
    gRow.font = { bold: true, size: 11 };
    gRow.height = 22;
    const fill = gt.fillArgb || 'FFFFF3CD';
    gRow.eachCell((cell, ci) => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fill } };
      cell.border = { top: { style: 'thin' }, bottom: { style: 'double' } };
      const m = colMeta[ci - 1];
      if (m && m.fmt)   cell.numFmt = m.fmt;
      cell.alignment = { horizontal: (m && m.align) || 'left', vertical: 'middle' };
    });
  }

  return wb.xlsx.writeBuffer();
}

// Local helper — A1 column letter. Mirrors the one in writeXlsxCompanyHeader
// but kept private here so we don't widen that module's exports.
function colLetter(n) {
  let s = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

// Build the common XLSX header meta lines for a given auction. Returns
// an array like ["e-TRADE No: 3", "Date: 15/04/2026"]. The crop type
// (ISP/ASP) is omitted — the active preset is already shown via the logo
// and company name in the brand block.
function auctionMeta(db, auctionId) {
  if (!auctionId) return [];
  try {
    const a = db.get(
      'SELECT ano, date, crop_type FROM auctions WHERE id = ?', [auctionId]
    );
    if (!a) return [];
    const dt = fmtUserDate(String(a.date || '').slice(0, 10));
    const meta = [];
    if (a.ano) meta.push(`e-TRADE No: ${a.ano}`);
    if (dt) meta.push(`Date: ${dt}`);
    return meta;
  } catch (_) { return []; }
}

// ── Export Type 1: Lot Slip (before trade) ───────────────────
async function exportLotSlip(db, auctionId, state) {
  const rows = db.all(
    `SELECT state, lot_no as lot, name, grade, bags as bag, qty, litre
     FROM lots WHERE auction_id = ? ${state ? 'AND state = ?' : ''}
     ORDER BY lot_no`, state ? [auctionId, state] : [auctionId]
  );
  const cols = [
    { header: 'STATE', key: 'state', width: 12 },
    { header: 'LOT', key: 'lot', width: 8 },
    { header: 'NAME', key: 'name', width: 30 },
    { header: 'GRADE', key: 'grade', width: 8 },
    { header: 'BAG', key: 'bag', width: 6 },
    { header: 'QTY', key: 'qty', width: 12 },
    { header: 'LITRE', key: 'litre', width: 10 },
  ];
  return createExcelBuffer('LotSlip', cols, rows, {
    db, title: 'Lot Slip', metaLines: auctionMeta(db, auctionId),
  });
}

// ── Export Type 2: Lot Slip After Trade (with price/buyer) ───
async function exportLotSlipAfter(db, auctionId, state) {
  const rows = db.all(
    `SELECT state, lot_no as lot, name, bags as bag, qty, price, amount, code
     FROM lots WHERE auction_id = ? ${state ? 'AND state = ?' : ''}
     ORDER BY lot_no`, state ? [auctionId, state] : [auctionId]
  );
  const cols = [
    { header: 'STATE', key: 'state', width: 12 },
    { header: 'LOT', key: 'lot', width: 8 },
    { header: 'NAME', key: 'name', width: 30 },
    { header: 'BAG', key: 'bag', width: 6 },
    { header: 'QTY', key: 'qty', width: 12 },
    { header: 'PRICE', key: 'price', width: 10 },
    { header: 'AMOUNT', key: 'amount', width: 14 },
    { header: 'CODE', key: 'code', width: 8 },
  ];
  return createExcelBuffer('LotSlipAfter', cols, rows, {
    db, title: 'Lot Slip (After Trade)', metaLines: auctionMeta(db, auctionId),
  });
}

// ── Export Type 3: Price List ─────────────────────────────────
async function exportPriceList(db, auctionId) {
  const rows = db.all(
    `SELECT lot_no as lot, bags as bag, qty, price, code, buyer as bidder
     FROM lots WHERE auction_id = ? ORDER BY lot_no`, [auctionId]
  );
  const cols = [
    { header: 'LOT', key: 'lot', width: 8 },
    { header: 'BAG', key: 'bag', width: 6 },
    { header: 'QTY', key: 'qty', width: 12 },
    { header: 'PRICE', key: 'price', width: 10 },
    { header: 'CODE', key: 'code', width: 8 },
    { header: 'BIDDER', key: 'bidder', width: 20 },
  ];
  return createExcelBuffer('PriceList', cols, rows, {
    db, title: 'Price List', metaLines: auctionMeta(db, auctionId),
  });
}

// ── Export Type 3b: Price List (Before Trade) ─────────────────
// Same shape as Price List but trade-result columns (PRICE, CODE,
// BIDDER) are dropped — meaningful only AFTER buyers bid on lots.
// Useful pre-trade for handing buyers a printable lot inventory.
async function exportPriceListBefore(db, auctionId) {
  const a = db.get('SELECT ano, date FROM auctions WHERE id = ?', [auctionId]) || {};
  const tradeNo = a.ano || '';
  const tradeDate = fmtUserDate(String(a.date || '').slice(0, 10));
  // PRICE + CODE are read straight from the lots table — both will be
  // blank pre-auction (auction not held yet) and populated post-
  // auction. The "Before" form is typically printed empty so buyers
  // can hand-fill PRICE / CODE during the auction.
  // PRICE is blanked when 0 (lot not yet priced) so the column reads
  // empty rather than "0.00" — matching how CODE shows blank when unset.
  const rawRows = db.all(
    `SELECT lot_no as lot, bags as bag, qty,
            CASE WHEN COALESCE(price,0) = 0 THEN '' ELSE price END AS price,
            COALESCE(code,'') AS code
     FROM lots WHERE auction_id = ? ORDER BY lot_no`, [auctionId]
  );
  const rows = rawRows.map(r => ({ trade_no: tradeNo, date: tradeDate, ...r }));
  const cols = [
    { header: 'TNO',   key: 'trade_no', width: 10 },
    { header: 'DATE',  key: 'date',     width: 12 },
    { header: 'LOT',   key: 'lot',      width: 10 },
    { header: 'BAG',   key: 'bag',      width: 8  },
    { header: 'QTY',   key: 'qty',      width: 14 },
    { header: 'PRICE', key: 'price',    width: 10 },
    { header: 'CODE',  key: 'code',     width: 10 },
  ];
  // Printed blank as a fillable form during the auction, so it carries no
  // brand band (company name + trade/date rows) and no TOTAL footer — just
  // the column headers and lot rows.
  return createExcelBuffer('PriceListBefore', cols, rows, {
    db, title: 'Price List (Before)', noBrandBand: true,
  });
}

// ── Export Type 4: Bank Payment (IFT/NEFT/RTGS — bank-import format) ─
// Bare 12-column sheet (no brand band, no totals) matching the bank's
// upload template:
//   Transaction Type | Debit Account Number | Transaction Amount |
//   Value Date | Beneficiary Account Number | Beneficiary Name |
//   IFSC Code | Beneficiary Email ID | Beneficiary ID | Credit Remarks |
//   Debit Remarks | Unique Customer Reference Number
// Header on row 1, data from row 2. Bank software auto-ingests this
// shape — adding a brand band would break the import.
async function exportBankPayment(db, auctionId, cfg, _state, opts) {
  const payments = resolveBankPayments(db, auctionId, cfg, opts);
  return buildBankPaymentSheet(db, auctionId, cfg, payments);
}

// Shared payment-row resolver for the Bank Payment AND Voucher Payment
// exports. Pulls getBankPaymentData(), then applies the Payments-tab
// "Export Selected" filters (seller-name filter + per-seller lot picks +
// already-exported exclusions) and returns the final `payments` array.
// Both exporters feed the result into their own sheet builder.
function resolveBankPayments(db, auctionId, cfg, opts) {
  const { getBankPaymentData, formatLotList, paymentTdsContext } = require('./calculations');
  let payments = getBankPaymentData(db, auctionId, cfg);
  // Same stamped-purchase-TDS source the Payments tab + base bank rows use, so
  // a lot-picked subset nets the proportionate share of the invoice TDS.
  const tdsCtx = paymentTdsContext(db, auctionId);
  // Optional seller-name filter — when the user clicks "Export Bank
  // Payment (Selected)" in the Payments tab, only the ticked sellers'
  // rows should appear in the bank upload file. Match against `p.name`
  // (the seller name) because that's the value the UI checkbox holds;
  // beneficiaryName tracks the bank account holder which can be a
  // different entity. When opts.names is absent or empty the full
  // payment set is exported, preserving the original behaviour.
  if (opts && Array.isArray(opts.names) && opts.names.length) {
    const wanted = new Set(opts.names.map(n => String(n || '').trim().toUpperCase()));
    payments = payments.filter(p =>
      wanted.has(String(p.name || '').trim().toUpperCase())
    );
  }
  // Optional per-seller lot-picks AND already-exported exclusions.
  //
  //   opts.lots[seller]        — operator-picked subset; only these
  //                              lots' balances count for the bank row.
  //   opts.excludeLots[seller] — lots that have already gone out in a
  //                              previous export and must be skipped
  //                              automatically so re-exporting the
  //                              seller doesn't double-pay them.
  //
  // For each seller in either map we re-query SUM(balance) with the
  // appropriate WHERE clauses and override the payments[] row in
  // place. RTGS/NEFT is re-picked from the new amount (₹2L threshold).
  // REMARKS is rebuilt later at row.map time using p.amount, so just
  // updating p.amount + p.transactionType here is enough.
  const lotPicks = (opts && opts.lots && typeof opts.lots === 'object') ? opts.lots : null;
  const excludeLots = (opts && opts.excludeLots && typeof opts.excludeLots === 'object') ? opts.excludeLots : null;
  if (lotPicks || excludeLots) {
    // Index every trader bank by id so a picked-lot subset that all points
    // at one account can route this export row to THAT account (this is how
    // "select Account-1's lots → Export Selected" credits Account 1, then a
    // second export of Account-2's lots credits Account 2).
    const bankById = {};
    try {
      for (const b of db.all('SELECT id, ifsc, acctnum, holder_name FROM trader_banks')) bankById[b.id] = b;
    } catch (_) { /* table may not exist on partial migrations */ }
    const sellersToRecompute = new Set();
    if (lotPicks)     for (const k of Object.keys(lotPicks))     sellersToRecompute.add(k);
    if (excludeLots)  for (const k of Object.keys(excludeLots))  sellersToRecompute.add(k);
    for (const sellerName of sellersToRecompute) {
      const picksArr   = lotPicks    && Array.isArray(lotPicks[sellerName])    ? lotPicks[sellerName]    : null;
      const excludeArr = excludeLots && Array.isArray(excludeLots[sellerName]) ? excludeLots[sellerName] : null;
      if ((!picksArr || !picksArr.length) && (!excludeArr || !excludeArr.length)) continue;
      const wantedUpper = String(sellerName || '').trim().toUpperCase();
      const idx = payments.findIndex(p =>
        String(p.name || '').trim().toUpperCase() === wantedUpper
      );
      if (idx < 0) continue;   // seller not in the current payments set (e.g. fully paid)
      const params = [auctionId, wantedUpper];
      let extraWhere = '';
      if (picksArr && picksArr.length) {
        extraWhere += ` AND l.lot_no IN (${picksArr.map(() => '?').join(',')})`;
        for (const lot of picksArr) params.push(String(lot));
      }
      if (excludeArr && excludeArr.length) {
        extraWhere += ` AND l.lot_no NOT IN (${excludeArr.map(() => '?').join(',')})`;
        for (const lot of excludeArr) params.push(String(lot));
      }
      const sub = db.get(
        `SELECT COALESCE(SUM(l.balance),0) AS payable,
                COALESCE(SUM(l.puramt), 0) AS puramt,
                MAX(l.cr) AS cr,
                GROUP_CONCAT(l.lot_no) AS lot_nos,
                GROUP_CONCAT(DISTINCT l.bank_id) AS bank_ids,
                COUNT(*) AS lot_count,
                COUNT(l.bank_id) AS bank_lot_count
           FROM lots l
          WHERE l.auction_id = ? AND l.amount > 0
            AND (l.paid IS NULL OR l.paid = '')
            AND UPPER(TRIM(l.name)) = ?${extraWhere}`,
        params
      ) || { payable: 0, puramt: 0, cr: '', lot_nos: '' };
      // Net the same purchase-invoice TDS (proportionate to the picked subset's
      // puramt) so this override keeps the bank amount = Payable (PurAmt −
      // Discount − GST − TDS), matching the base getBankPaymentData rows.
      const tdsShare = tdsCtx.share(sellerName, sub.puramt);
      const rawAmount = (Number(sub.payable) || 0) - tdsShare;
      const roundedAmount = cfg.flag_round ? Math.round(rawAmount) : rawAmount;
      const isRTGS = roundedAmount >= 200000;
      // If every picked lot points at the same single bank account, route
      // this row to that account (overrides the seller-default account that
      // getBankPaymentData put on the base row). Mixed/untagged → leave the
      // default account as-is.
      const subBankIds = String(sub.bank_ids || '')
        .split(',').map(s => s.trim()).filter(s => s !== '' && s !== 'null')
        .map(Number).filter(Number.isFinite);
      const subUntagged = Number(sub.lot_count || 0) > Number(sub.bank_lot_count || 0);
      const subBank = (subBankIds.length === 1 && !subUntagged) ? bankById[subBankIds[0]] : null;
      payments[idx] = {
        ...payments[idx],
        amount: roundedAmount,
        // Keep the per-row TDS proportionate to the picked subset so the
        // Voucher Payment TDS column matches this row's recomputed amount.
        tds: tdsShare,
        transactionType: isRTGS ? 'RTGS' : 'NEFT',
        // Re-derive the covered-lots list from the same picked/excluded
        // subset so REMARKS lists exactly the lots this row pays for.
        lots: formatLotList(sub.lot_nos),
        ...(subBank ? {
          ifsc: subBank.ifsc || payments[idx].ifsc,
          accountNo: subBank.acctnum || payments[idx].accountNo,
          beneficiaryName: subBank.holder_name || payments[idx].beneficiaryName,
        } : {}),
      };
    }
    // Drop any zero-amount rows produced by the recompute — banks reject
    // zero-value RTGS rows, and a seller whose remaining lots all net to
    // zero (everything already exported, or only zero-balance lots
    // picked) shouldn't appear at all.
    payments = payments.filter(p => Number(p.amount) > 0);
  }

  return payments;
}

// Shared sheet builder for both Bank Payment variants (after-discount and
// before-discount). Takes the already-computed `payments` rows and emits the
// bank's 12-column upload template. Transaction Type / amount are derived
// per-row, so the caller only has to decide which amount (balance vs puramt)
// getBankPaymentData put on each `p.amount`.
function buildBankPaymentSheet(db, auctionId, cfg, payments) {
  // Sender-side context (state-aware): debit account + IFSC. The IFSC
  // bank prefix decides IFT (same bank) vs NEFT/RTGS (other banks).
  const isKL = String(cfg.business_state || cfg.state || '').toUpperCase().includes('KERALA');
  const senderAcct  = (isKL ? cfg.bank_kl_acct  : cfg.bank_tn_acct)  || cfg.bank_tn_acct  || cfg.bank_kl_acct  || '';
  const senderIfsc  = (isKL ? cfg.bank_kl_ifsc  : cfg.bank_tn_ifsc)  || cfg.bank_tn_ifsc  || cfg.bank_kl_ifsc  || '';
  const senderBankPrefix = String(senderIfsc).slice(0, 4).toUpperCase();

  // Auction context: ano (REMARKS prefix) + value date. Value Date is the
  // date the file is generated (the export date), NOT the trade date — the
  // bank posts each payment as of today, and this matches the Payments-tab
  // "🚫 EXPORTED on …" badge. The trade date still travels in Credit Remarks
  // via `ano`. Server TZ is pinned to IST so local date parts are correct.
  const a = db.get('SELECT ano, date FROM auctions WHERE id = ?', [auctionId]) || {};
  const ano = a.ano || '';
  const _now = new Date();
  const _todayISO = `${_now.getFullYear()}-${String(_now.getMonth() + 1).padStart(2, '0')}-${String(_now.getDate()).padStart(2, '0')}`;
  const valueDate = fmtUserDate(_todayISO);

  // Buyer = the trading company itself, shown as the 3rd segment of Credit
  // Remarks (e.g. "VSTL"). Prefer the Praman nickname (the e-Trade company
  // code), then Short Name, then Trade Name.
  const buyerLabel = String(
    cfg.praman_company || cfg.short_name || cfg.trade_name || ''
  ).toUpperCase();

  const rows = payments.map(p => {
    const amount = Number(p.amount) || 0;
    const beneIfsc = String(p.ifsc || '').toUpperCase();
    const benePrefix = beneIfsc.slice(0, 4);
    // Transaction Type — single column collapsing the old BT/LBT + NEFT/RTGS pair:
    //   IFT  = internal funds transfer (beneficiary is in the SAME bank as the
    //          debit account, matched on the 4-char IFSC bank prefix)
    //   RTGS = different bank, amount >= ₹2L
    //   NEFT = different bank, amount  < ₹2L
    const sameBank = senderBankPrefix && benePrefix === senderBankPrefix;
    const transactionType = sameBank
      ? 'IFT'
      : (amount >= 200000 ? 'RTGS' : 'NEFT');
    return {
      TRANS_TYPE:  transactionType,
      DEBIT_ACCT:  senderAcct,
      AMOUNT:      amount,
      VALUE_DATE:  valueDate,
      BENE_ACCT:   p.accountNo || '',
      BENE_NAME:   String(p.beneficiaryName || '').toUpperCase(),
      BENE_IFSC:   beneIfsc,
      BENE_EMAIL:  '',
      BENE_ID:     '',
      // Pipe-delimited Credit Remarks: ano | seller | buyer | payment | lots
      //   12 | ANN MARIA SPICES | VSTL | PAYMENT 1764625.00 Credited | For lots 002,065,103
      // 2nd segment is the seller being paid (p.name); 3rd is the buyer — the
      // trading company itself (buyerLabel), NOT a repeat of the seller.
      // The lots segment is dropped when the row covers no lots.
      CREDIT_REM:  [
        ano,
        String(p.name || '').toUpperCase(),
        buyerLabel,
        `PAYMENT ${amount.toFixed(2)} Credited`,
        p.lots ? `For lot${p.lots.includes(',') ? 's' : ''} ${p.lots}` : '',
      ].filter(Boolean).join(' | '),
      DEBIT_REM:   '',
      UCRN:        '',
    };
  });

  // Build the sheet directly (bypass createExcelBuffer's brand-band).
  // Column order/headers mirror the bank's upload template exactly.
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('BANK_PAYMENT');
  const cols = [
    { key: 'TRANS_TYPE',  header: 'Transaction Type',                 width: 14 },
    { key: 'DEBIT_ACCT',  header: 'Debit Account Number',             width: 20 },
    { key: 'AMOUNT',      header: 'Transaction Amount',               width: 16, numFmt: '#,##0.00' },
    { key: 'VALUE_DATE',  header: 'Value Date',                       width: 12 },
    { key: 'BENE_ACCT',   header: 'Beneficiary Account Number',       width: 22 },
    { key: 'BENE_NAME',   header: 'Beneficiary Name',                 width: 34 },
    { key: 'BENE_IFSC',   header: 'IFSC Code',                        width: 14 },
    { key: 'BENE_EMAIL',  header: 'Beneficiary Email ID',             width: 24 },
    { key: 'BENE_ID',     header: 'Beneficiary ID',                   width: 14 },
    { key: 'CREDIT_REM',  header: 'Credit Remarks',                   width: 70 },
    { key: 'DEBIT_REM',   header: 'Debit Remarks',                    width: 16 },
    { key: 'UCRN',        header: 'Unique Customer Reference Number', width: 28 },
  ];
  ws.columns = cols.map(c => ({ key: c.key, width: c.width }));
  cols.forEach((c, i) => {
    if (c.numFmt) ws.getColumn(i + 1).numFmt = c.numFmt;
  });
  // Header row 1 — plain bold, no fill (so bank importers don't choke).
  const head = ws.getRow(1);
  cols.forEach((c, i) => { head.getCell(i + 1).value = c.header; });
  head.font = { bold: true };
  // Data rows from row 2.
  rows.forEach(r => ws.addRow(r));
  return wb.xlsx.writeBuffer();
}

// ── Export Type 4c: Voucher Payment ──────────────────────────
// A compact 6-column payment voucher (a different bank's bulk-transfer
// template than the 12-column RTGS sheet). Shares the same row source as
// Bank Payment — so the Payments-tab "Export Selected" seller/lot filters
// apply identically — but renders the simpler voucher layout:
//
//   Row 1 (band): <Bank name>  <Bank account no>            Dt: DD-MM-YYYY
//   Row 2       : (spacer)
//   Row 3 (green): Particulars | Customer Name | Amount | Account Number | IFSC | TDS
//   Row 4+      : one row per seller
//
// Amount is the net Payable already wired (puramt − discount − GST − TDS);
// the TDS column shows the deducted TDS for reference.
async function exportVoucherPayment(db, auctionId, cfg, _state, opts) {
  const payments = resolveBankPayments(db, auctionId, cfg, opts);
  return buildVoucherSheet(db, auctionId, cfg, payments);
}

function buildVoucherSheet(db, auctionId, cfg, payments) {
  // Sender-side context (state-aware), same source the 12-column bank sheet
  // uses for the debit account — Kerala vs TN picks the company's bank.
  const isKL = String(cfg.business_state || cfg.state || '').toUpperCase().includes('KERALA');
  const senderName = (isKL ? cfg.bank_kl_name : cfg.bank_tn_name) || cfg.bank_tn_name || cfg.bank_kl_name || '';
  const senderAcct = (isKL ? cfg.bank_kl_acct : cfg.bank_tn_acct) || cfg.bank_tn_acct || cfg.bank_kl_acct || '';
  // Sender IFSC's 4-char bank prefix — used to split each row into
  // IFT (same bank) vs NEFT/RTGS (other bank), exactly like the 12-column
  // bank sheet's Transaction Type column.
  const senderIfsc = (isKL ? cfg.bank_kl_ifsc : cfg.bank_tn_ifsc) || cfg.bank_tn_ifsc || cfg.bank_kl_ifsc || '';
  const senderBankPrefix = String(senderIfsc).slice(0, 4).toUpperCase();

  // Auction context: ano (Particulars prefix) + value date. Value Date (the
  // "Dt:" band) is the date the file is generated (the export date), NOT the
  // trade date — matches the 12-column bank sheet and the Payments-tab
  // "🚫 EXPORTED on …" badge. Server TZ is pinned to IST so local date parts
  // are correct.
  const a = db.get('SELECT ano, date FROM auctions WHERE id = ?', [auctionId]) || {};
  const ano = a.ano || '';
  const _now = new Date();
  const _todayISO = `${_now.getFullYear()}-${String(_now.getMonth() + 1).padStart(2, '0')}-${String(_now.getDate()).padStart(2, '0')}`;
  const valueDate = fmtUserDate(_todayISO);

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('VOUCHER_PAYMENT');
  ws.columns = [
    { key: 'particulars', width: 38 },
    { key: 'customer',    width: 34 },
    { key: 'amount',      width: 16 },
    { key: 'acct',        width: 22 },
    { key: 'ifsc',        width: 16 },
    { key: 'tds',         width: 14 },
  ];
  ws.getColumn(3).numFmt = '#,##0.00';
  ws.getColumn(6).numFmt = '#,##0.00';

  const thin = { style: 'thin', color: { argb: 'FF000000' } };
  const boxAll = (row) => { for (let c = 1; c <= 6; c++) row.getCell(c).border = { top: thin, bottom: thin, left: thin, right: thin }; };

  // Row 1 — header band: company bank name, debit account, value date.
  const band = ws.getRow(1);
  band.getCell(1).value = senderName;
  band.getCell(2).value = senderAcct;
  band.getCell(5).value = valueDate ? `Dt: ${valueDate}` : '';
  band.font = { bold: true };
  boxAll(band);

  // Row 2 — spacer (left blank, matching the template).

  // Row 3 — green column header.
  const head = ws.getRow(3);
  const headers = ['Particulars', 'Customer Name', 'Amount', 'Account Number', 'IFSC', 'TDS'];
  headers.forEach((h, i) => {
    const cell = head.getCell(i + 1);
    cell.value = h;
    cell.font = { bold: true };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF93C47D' } };
    cell.border = { top: thin, bottom: thin, left: thin, right: thin };
  });

  // Data rows from row 4. Tally running totals + per-transaction-type
  // splits (IFT/NEFT/RTGS) for the summary blocks below the table.
  let rowIdx = 4;
  let totAmount = 0, totTds = 0;
  const byType = { IFT: { amt: 0, cnt: 0 }, NEFT: { amt: 0, cnt: 0 }, RTGS: { amt: 0, cnt: 0 } };
  for (const p of payments) {
    const lots = p.lots || '';
    // Particulars = auction/trade ref + the lots this payment covers.
    const particulars = `${ano}${lots ? ` / lot${lots.includes(',') ? 's' : ''} ${lots}` : ''}`.trim();
    const amount = Number(p.amount) || 0;
    const rawTds = Number(p.tds) || 0;
    const tdsVal = cfg.flag_round ? Math.round(rawTds) : rawTds;
    // Transaction type, same rule as the bank sheet: same bank → IFT,
    // else ≥₹2L → RTGS, else NEFT.
    const benePrefix = String(p.ifsc || '').toUpperCase().slice(0, 4);
    const sameBank = senderBankPrefix && benePrefix === senderBankPrefix;
    const txType = sameBank ? 'IFT' : (amount >= 200000 ? 'RTGS' : 'NEFT');
    const row = ws.getRow(rowIdx++);
    row.getCell(1).value = particulars;
    row.getCell(2).value = String(p.beneficiaryName || p.name || '').toUpperCase();
    row.getCell(3).value = amount;
    // Account number kept as text so long numbers don't render in scientific
    // notation / lose leading zeros on import.
    row.getCell(4).value = String(p.accountNo || '');
    row.getCell(5).value = String(p.ifsc || '').toUpperCase();
    row.getCell(6).value = tdsVal;
    boxAll(row);
    totAmount += amount;
    totTds += tdsVal;
    byType[txType].amt += amount;
    byType[txType].cnt += 1;
  }

  // Total row — sum of Amount (col 3) and TDS (col 6), boxed like the table.
  const totalRow = ws.getRow(rowIdx++);
  totalRow.getCell(1).value = 'Total';
  totalRow.getCell(3).value = totAmount;
  totalRow.getCell(6).value = totTds;
  totalRow.font = { bold: true };
  boxAll(totalRow);

  // Spacer row, then the signature row (Prepared / Checked / Approved), each
  // label paired with the blank cell to its right for the actual sign-off.
  rowIdx++;
  const sigRow = ws.getRow(rowIdx++);
  sigRow.getCell(1).value = 'Prepared By';
  sigRow.getCell(3).value = 'Checked By';
  sigRow.getCell(5).value = 'Approved By';
  sigRow.font = { bold: true };
  boxAll(sigRow);

  // Two spacer rows, then the IFT / NEFT / RTGS breakdown: label (col 2,
  // right-aligned) | amount (col 3) | count (col 4), capped with a TOTAL.
  rowIdx += 2;
  const totalCount = byType.IFT.cnt + byType.NEFT.cnt + byType.RTGS.cnt;
  const sumRow = (label, amt, cnt) => {
    const r = ws.getRow(rowIdx++);
    const lblCell = r.getCell(2);
    lblCell.value = label;
    lblCell.font = { bold: true };
    lblCell.alignment = { horizontal: 'right' };
    const amtCell = r.getCell(3);
    amtCell.value = amt;
    amtCell.font = { bold: true };
    const cntCell = r.getCell(4);
    cntCell.value = cnt;
    cntCell.font = { bold: true };
    cntCell.alignment = { horizontal: 'right' };
  };
  sumRow('IFT',  byType.IFT.amt,  byType.IFT.cnt);
  sumRow('NEFT', byType.NEFT.amt, byType.NEFT.cnt);
  sumRow('RTGS', byType.RTGS.amt, byType.RTGS.cnt);
  rowIdx++;
  sumRow('TOTAL', totAmount, totalCount);

  return wb.xlsx.writeBuffer();
}

// ── Export Type 4b: Bank Payment (Before discount) ───────────
// Identical 12-column bank-upload layout as bank_payment, except each row's
// amount is the pre-discount puramt (raw purchase amount before refund/GST)
// — useful when paying suppliers before the deduction policy is applied.
// Transaction Type is still re-derived from that amount + IFSC.
async function exportBankPaymentBefore(db, auctionId, cfg, _state, opts) {
  const { getBankPaymentData } = require('./calculations');
  let payments = getBankPaymentData(db, auctionId, cfg, { before: true });
  // Optional seller-name filter (same semantics as exportBankPayment) — only
  // the ticked sellers' rows are written when the UI passes opts.names.
  if (opts && Array.isArray(opts.names) && opts.names.length) {
    const wanted = new Set(opts.names.map(n => String(n || '').trim().toUpperCase()));
    payments = payments.filter(p =>
      wanted.has(String(p.name || '').trim().toUpperCase())
    );
  }
  return buildBankPaymentSheet(db, auctionId, cfg, payments);
}

// ── Export Type 5: Pooler-wise Register ───────────────────────
async function exportPoolerRegister(db, auctionId) {
  // Filter on NOT-withdrawn rather than `amount > 0`. Post-auction the
  // only 0-amount lots are withdrawn ones, so this matches the old
  // output exactly; but PRE-auction (lots imported, not yet priced) the
  // `amount > 0` filter excluded everything and the register came out
  // empty. Money columns are blanked when 0 so the pre-priced register
  // reads clean instead of showing "0.00" everywhere (matches the
  // Price List / Code blank-when-unset convention).
  const rows = db.all(
    `SELECT state, lot_no as lot, name as poolername, branch as br, qty,
            CASE WHEN COALESCE(price,0)  = 0 THEN '' ELSE price  END AS price,
            CASE WHEN COALESCE(amount,0) = 0 THEN '' ELSE amount END AS amount,
            CASE WHEN COALESCE(pqty,0)   = 0 THEN '' ELSE pqty   END AS pqty,
            CASE WHEN COALESCE(prate,0)  = 0 THEN '' ELSE prate  END AS prate,
            CASE WHEN COALESCE(puramt,0) = 0 THEN '' ELSE puramt END AS puramt
     FROM lots WHERE auction_id = ? AND UPPER(TRIM(COALESCE(code,''))) != 'WD'
     ORDER BY name`, [auctionId]
  );
  const cols = [
    { header: 'STATE', key: 'state', width: 12 },
    { header: 'NAME', key: 'poolername', width: 30 },
    { header: 'BRANCH', key: 'br', width: 15 },
    { header: 'LOT', key: 'lot', width: 8 },
    { header: 'QTY', key: 'qty', width: 12 },
    { header: 'PRICE', key: 'price', width: 10 },
    { header: 'AMOUNT', key: 'amount', width: 14 },
    { header: 'PQTY', key: 'pqty', width: 12 },
    { header: 'PRATE', key: 'prate', width: 10 },
    { header: 'PURAMT', key: 'puramt', width: 14 },
  ];
  return createExcelBuffer('PoolerRegister', cols, rows, {
    db, title: 'Pooler Register', metaLines: auctionMeta(db, auctionId),
  });
}

// ── Export Type 6: Full File ─────────────────────────────────
async function exportFullFile(db, auctionId) {
  const rows = db.all(`SELECT * FROM lots WHERE auction_id = ? ORDER BY lot_no`, [auctionId]);
  const cols = [
    { header: 'STATE', key: 'state' }, { header: 'LOT', key: 'lot_no', width: 8 },
    { header: 'CROP', key: 'crop' }, { header: 'GRADE', key: 'grade' },
    { header: 'CRPT', key: 'crpt' }, { header: 'BRANCH', key: 'branch', width: 15 },
    { header: 'NAME', key: 'name', width: 30 }, { header: 'CR', key: 'cr', width: 25 },
    { header: 'PAN', key: 'pan' }, { header: 'TEL', key: 'tel' },
    { header: 'BAG', key: 'bags', width: 6 }, { header: 'QTY', key: 'qty', width: 12 },
    { header: 'PRICE', key: 'price', width: 10 }, { header: 'AMOUNT', key: 'amount', width: 14 },
    { header: 'CODE', key: 'code' }, { header: 'BUYER', key: 'buyer', width: 15 },
    { header: 'BUYER1', key: 'buyer1', width: 20 }, { header: 'SALE', key: 'sale' },
    { header: 'INVO', key: 'invo' }, { header: 'PQTY', key: 'pqty', width: 12 },
    { header: 'PRATE', key: 'prate', width: 10 }, { header: 'PURAMT', key: 'puramt', width: 14 },
    { header: 'COM', key: 'com' }, { header: 'CGST', key: 'cgst' },
    { header: 'SGST', key: 'sgst' }, { header: 'IGST', key: 'igst' },
    { header: 'ADVANCE', key: 'advance', width: 14 }, { header: 'BALANCE', key: 'balance', width: 14 },
  ];
  return createExcelBuffer('FullFile', cols, rows, {
    db, title: 'Full File', metaLines: auctionMeta(db, auctionId),
  });
}

// ── Export Type 7: Collection (invoice register) ─────────────
// Mirrors COLLECTION.pdf: one row per sales invoice issued, grouped by buyer
// state. Columns: SALE+INVO | TRADE NAME (firm) | NAME (buyer) | QTY | VALUE.
async function exportCollection(db, auctionId) {
  return newCollectionXlsx(db, auctionId);
}

// ── Export Type 8: Dealer List ────────────────────────────────
// Dealers are sellers whose `cr` field stores a GSTIN. Storage is
// inconsistent across imports — values appear as "GSTIN.<15>", "gstin <15>",
// "gstin<15>", or bare 15-char alphanumeric. The earlier query
// `WHERE cr LIKE '%GST%'` skipped the bare-15 case (silently returning
// an empty XLSX) and SUBSTR(cr,7,15) hard-coded a 6-char prefix.
//
// Fix: compute a clean GSTIN inline (strip any 'gstin' prefix +
// punctuation/whitespace, uppercase) and filter on its length being
// exactly 15. Works for every storage form.
async function exportDealerList(db, auctionId) {
  // No `amount > 0` filter: a dealer list enumerates the registered
  // dealers (sellers with a GSTIN) participating in the trade and is a
  // PRE-trade safety-net export — at that point lots are imported but
  // not yet priced (amount = 0), so filtering on amount returned an
  // empty sheet. Bags/qty are entered at import time, so they're
  // accurate regardless of pricing.
  const rows = db.all(
    `WITH cleaned AS (
       SELECT state, name, lot_no, bags, qty,
              UPPER(TRIM(
                CASE
                  WHEN LOWER(SUBSTR(TRIM(cr),1,5)) = 'gstin'
                    THEN LTRIM(SUBSTR(TRIM(cr),6), '. :-')
                  ELSE TRIM(cr)
                END
              )) AS gstin
         FROM lots
        WHERE auction_id = ?
     )
     SELECT state, name, gstin,
            COUNT(lot_no) as lots, SUM(bags) as bags, SUM(qty) as qty
       FROM cleaned
      WHERE LENGTH(gstin) = 15
      GROUP BY state, name, gstin
      ORDER BY state, name`, [auctionId]
  );
  const cols = [
    { header: 'STATE', key: 'state', width: 12 },
    { header: 'NAME', key: 'name', width: 30 },
    { header: 'GSTIN', key: 'gstin', width: 18 },
    { header: 'LOTS', key: 'lots', width: 6 },
    { header: 'BAGS', key: 'bags', width: 6 },
    { header: 'QTY', key: 'qty', width: 12 },
  ];
  return createExcelBuffer('DealerList', cols, rows, {
    db, title: 'Dealer List', metaLines: auctionMeta(db, auctionId),
  });
}

// ── Export Type 9: Sales & Taxes ─────────────────────────────
async function exportSalesTaxes(db, auctionId) {
  const rows = db.all(
    `SELECT state, sale, invo, buyer1 as tradername, bags as bag, qty, 
      amount as cardamom_cost, gunny as gunny_cost,
      cgst, sgst, igst, tcs, pava_hc as transport, ins as insurance, tot as total
     FROM invoices WHERE ano = (SELECT ano FROM auctions WHERE id = ?)
     ORDER BY sale, invo`, [auctionId]
  );
  const cols = [
    { header: 'STATE', key: 'state' }, { header: 'SALE', key: 'sale' },
    { header: 'INVO', key: 'invo' }, { header: 'TRADERNAME', key: 'tradername', width: 25 },
    { header: 'BAG', key: 'bag', width: 6 }, { header: 'QTY', key: 'qty', width: 12 },
    { header: 'CARDAMOM', key: 'cardamom_cost', width: 14 },
    { header: 'GUNNY', key: 'gunny_cost', width: 10 },
    { header: 'CGST', key: 'cgst', width: 12 }, { header: 'SGST', key: 'sgst', width: 12 },
    { header: 'IGST', key: 'igst', width: 12 }, { header: 'TCS', key: 'tcs', width: 10 },
    { header: 'TRANSPORT', key: 'transport', width: 10 },
    { header: 'INSURANCE', key: 'insurance', width: 10 },
    { header: 'TOTAL', key: 'total', width: 14 },
  ];
  return createExcelBuffer('SalesTaxes', cols, rows, {
    db, title: 'Sales & Taxes', metaLines: auctionMeta(db, auctionId),
  });
}

// ── Export: Payment Summary ──────────────────────────────────
async function exportPaymentSummary(db, auctionId, cfg, _state, opts) {
  // Match getPaymentSummary semantics: discount includes BOTH the per-lot
  // policy discount AND any manual debit_notes for this auction's sellers.
  // We compute it per-row by adding debit_notes (joined by ano + name).
  const mode = (cfg && cfg.business_mode || 'e-Trade').toLowerCase();
  const discountCol = (mode === 'auction') ? 'advance' : 'refund';
  // GST 5% is the GST-on-discount the calc already stored in `advance`
  // (cgst+sgst+igst). In auction mode `advance` is repurposed as the
  // discount itself, so there's no separate GST column → 0.
  const gstCol = (mode === 'auction') ? '0' : 'advance';
  const auction = db.get('SELECT ano FROM auctions WHERE id = ?', [auctionId]);
  const ano = auction ? auction.ano : null;
  // Build name → manual debit total map (debit_notes can have multiple
  // rows per seller per auction; we sum)
  const debitMap = {};
  if (ano) {
    const debits = db.all(
      'SELECT name, SUM(amount) as total FROM debit_notes WHERE ano = ? GROUP BY name',
      [ano]
    );
    for (const d of debits) debitMap[d.name] = Number(d.total) || 0;
  }
  // TDS comes from the seller's stamped purchase invoice (Section 194Q),
  // spread per row ∝ that lot's puramt — same as the Payments tab. 0 until the
  // invoice is generated / below the ₹50-lakh threshold.
  const { paymentTdsContext, distributeRoundedPayable } = require('./calculations');
  const tdsCtx = paymentTdsContext(db, auctionId);
  // Optional seller-name filter — "Export Payment XLSX (Selected)"
  // limits the rows to the ticked sellers. We push the filter into the
  // SQL with a `name IN (…)` clause so we don't waste the SELECT on rows
  // we'll throw away. Names are matched case-insensitively to be
  // resilient to slight casing drift between the UI and the DB.
  const filterNames = (opts && Array.isArray(opts.names) && opts.names.length)
    ? opts.names.map(n => String(n || '').trim()).filter(Boolean)
    : null;
  let whereExtra = '';
  const params = [auctionId];
  if (filterNames && filterNames.length) {
    const placeholders = filterNames.map(() => '?').join(',');
    whereExtra = ` AND UPPER(TRIM(name)) IN (${placeholders})`;
    for (const n of filterNames) params.push(n.toUpperCase());
  }
  let rows = db.all(
    `SELECT name as poolername, cr, lot_no as lot, bags as bag, qty, price, amount,
      pqty, prate, puramt, ${discountCol} as lot_discount, ${gstCol} as lot_gst, balance as payable
     FROM lots WHERE auction_id = ? AND amount > 0${whereExtra}
     ORDER BY state, name`, params
  );
  // Optional per-seller lot-picks AND already-exported exclusions.
  // Same shape as exportBankPayment:
  //   opts.lots[seller]        — keep ONLY these lot rows for the seller
  //   opts.excludeLots[seller] — drop these lots (already shipped before)
  // Both filters compose: if a seller has both, the row must satisfy
  // BOTH conditions to survive. Match name case-insensitively to be
  // tolerant of slight casing drift between localStorage and the DB.
  const lotPicks    = (opts && opts.lots         && typeof opts.lots         === 'object') ? opts.lots         : null;
  const excludeLots = (opts && opts.excludeLots  && typeof opts.excludeLots  === 'object') ? opts.excludeLots  : null;
  if (lotPicks || excludeLots) {
    const picksUpper   = {};
    const excludeUpper = {};
    if (lotPicks) for (const k of Object.keys(lotPicks)) {
      const arr = Array.isArray(lotPicks[k]) ? lotPicks[k] : [];
      if (arr.length) picksUpper[k.trim().toUpperCase()] = new Set(arr.map(x => String(x)));
    }
    if (excludeLots) for (const k of Object.keys(excludeLots)) {
      const arr = Array.isArray(excludeLots[k]) ? excludeLots[k] : [];
      if (arr.length) excludeUpper[k.trim().toUpperCase()] = new Set(arr.map(x => String(x)));
    }
    rows = rows.filter(r => {
      const key = String(r.poolername || '').trim().toUpperCase();
      const lotKey = String(r.lot);
      const picks = picksUpper[key];
      if (picks && !picks.has(lotKey)) return false;
      const excl = excludeUpper[key];
      if (excl && excl.has(lotKey)) return false;
      return true;
    });
  }
  // Spread debit_notes amount across the seller's lots proportionally so
  // every row totals to the same SUM as the payments view. Simpler approach:
  // attribute the FULL manual debit on the FIRST row for each seller; later
  // rows show only the lot policy discount. Avoids per-row arithmetic but
  // still preserves the seller-level total.
  const seenSellers = new Set();
  const enrichedFlat = rows.map(r => {
    const lotDisc = Number(r.lot_discount) || 0;
    const firstForSeller = !seenSellers.has(r.poolername);
    const manualDisc = firstForSeller ? (Number(debitMap[r.poolername]) || 0) : 0;
    // Each row carries its share of the seller's stamped purchase-invoice TDS,
    // spread ∝ this lot's puramt.
    const tds = tdsCtx.share(r.poolername, r.puramt);
    seenSellers.add(r.poolername);
    const discount = lotDisc + manualDisc;
    return {
      ...r,
      discount,
      // GST 5% = the stored GST-on-discount (`advance`). It's already
      // netted inside `balance`/PAYABLE, so we only DISPLAY it here and
      // never subtract it again.
      gst5: Number(r.lot_gst) || 0,
      tds,
      payable: (Number(r.payable) || 0) - manualDisc - tds,
    };
  });

  // When invoice rounding is on (cfg.flag_round), round each seller's per-lot
  // Payable to whole rupees — distributed so the lines sum EXACTLY to that
  // seller's rounded total (so subtotals + grand total stay whole and foot).
  // Rows are contiguous per seller (ORDER BY state, name).
  const roundPay = cfg.flag_round === true || String(cfg.flag_round || '').toLowerCase() === 'true';
  if (roundPay) {
    let gi = 0;
    while (gi < enrichedFlat.length) {
      let gj = gi;
      while (gj < enrichedFlat.length && enrichedFlat[gj].poolername === enrichedFlat[gi].poolername) gj++;
      const grp = enrichedFlat.slice(gi, gj);
      const rounded = distributeRoundedPayable(grp.map(r => r.payable));
      for (let k = 0; k < grp.length; k++) grp[k].payable = rounded[k];
      gi = gj;
    }
  }

  // Interleave per-pooler subtotal rows after each name group — mirrors
  // the PDF's groupByKey:'poolername' subtotalKeys behaviour. Rows are
  // already sorted by (state, name) so a single linear pass groups them.
  const SUB_KEYS = ['bag', 'qty', 'amount', 'pqty', 'puramt', 'discount', 'gst5', 'tds', 'payable'];
  const enriched = [];
  let curName = null;
  let acc = null;
  const flushSub = () => {
    if (!acc || curName == null) return;
    const sub = { _isSubtotal: true, poolername: `${curName} TOTAL` };
    SUB_KEYS.forEach(k => { sub[k] = acc[k] || 0; });
    enriched.push(sub);
  };
  for (const r of enrichedFlat) {
    const k = r.poolername || '';
    if (k !== curName) {
      flushSub();
      curName = k;
      acc = Object.fromEntries(SUB_KEYS.map(x => [x, 0]));
    }
    SUB_KEYS.forEach(x => { acc[x] += Number(r[x]) || 0; });
    enriched.push(r);
  }
  flushSub();
  const cols = [
    { header: 'POOLERNAME', key: 'poolername', width: 30 },
    { header: 'LOT', key: 'lot', width: 8 }, { header: 'BAG', key: 'bag', width: 6 },
    { header: 'QTY', key: 'qty', width: 12 }, { header: 'PRICE', key: 'price', width: 10 },
    { header: 'AMOUNT', key: 'amount', width: 14 }, { header: 'PQTY', key: 'pqty', width: 12 },
    { header: 'PRATE', key: 'prate', width: 10 }, { header: 'PURAMT', key: 'puramt', width: 14 },
    { header: 'DISCOUNT', key: 'discount', width: 14 },
    { header: 'GST 5%', key: 'gst5', width: 12 },
    { header: 'TDS', key: 'tds', width: 12 },
    { header: 'PAYABLE', key: 'payable', width: 14 },
  ];
  // Footer totals — sum every numeric column. The earlier export had no
  // totals row, so users had to compute payable/discount sums manually
  // in Excel before reconciling with bank transfers. PRICE/PRATE are
  // omitted from the sum (averaging rates makes no business sense; a
  // sum would mislead readers). Sum over enrichedFlat (data rows only)
  // so interleaved subtotals don't double-count.
  const sum = (key) => enrichedFlat.reduce((s, r) => s + (Number(r[key]) || 0), 0);
  const grandTotal = {
    label: 'GRAND TOTAL',
    values: {
      bag:     sum('bag'),
      qty:     sum('qty'),
      amount:  sum('amount'),
      pqty:    sum('pqty'),
      puramt:  sum('puramt'),
      discount:sum('discount'),
      gst5:    sum('gst5'),
      tds:     sum('tds'),
      payable: sum('payable'),
    },
  };
  return createExcelBuffer('Payment', cols, enriched, {
    db, title: 'Payment Summary', metaLines: auctionMeta(db, auctionId),
    grandTotal,
  });
}

// ── Export: Payment Summary — Party wise ─────────────────────
// One aggregated row PER POOLER (seller), grouped by state with a
// per-state subtotal and a grand total. Mirrors the legacy DOS
// "party-wise payment" report: columns POOLER NAME | P_QTY | P_RATE |
// PURAMOUNT | DISCOUNT | GST 5% | PAYABLE. P_RATE is intentionally blank
// (a party spans many lots at different rates, so a party-level rate is
// meaningless). DISCOUNT = refund, GST 5% = stored GST-on-discount
// (`advance`), PAYABLE = balance (GST already netted in).
async function exportPaymentPartywise(db, auctionId, cfg, _state, opts) {
  const mode = (cfg && cfg.business_mode || 'e-Trade').toLowerCase();
  const discountCol = (mode === 'auction') ? 'advance' : 'refund';
  const gstCol = (mode === 'auction') ? '0' : 'advance';
  const rows = db.all(
    `SELECT state,
            name           as poolername,
            MAX(cr)        as cr,
            SUM(pqty)      as pqty,
            SUM(puramt)    as puramt,
            SUM(${discountCol}) as discount,
            SUM(${gstCol}) as gst5,
            SUM(balance)   as payable
       FROM lots
      WHERE auction_id = ? AND amount > 0
      GROUP BY state, name
      ORDER BY state, name`, [auctionId]);

  // Per-seller TDS — the stamped Section 194Q purchase-invoice TDS for this
  // trade (0 until the invoice is generated / below threshold). Same as the
  // Payments tab. Subtracted straight off PAYABLE (Payable = PurAmt −
  // Discount − GST 5% − TDS).
  const { paymentTdsContext, round0 } = require('./calculations');
  const tdsCtx = paymentTdsContext(db, auctionId);
  // Round each seller's Payable to whole rupees when invoice rounding is on
  // (cfg.flag_round), matching the Payments tab + bank file.
  const roundPay = cfg.flag_round === true || String(cfg.flag_round || '').toLowerCase() === 'true';
  for (const r of rows) {
    r.tds = tdsCtx.share(r.poolername, r.puramt);
    const raw = (Number(r.payable) || 0) - r.tds;
    r.payable = roundPay ? round0(raw) : raw;
  }

  // Interleave a per-state header row and a per-state subtotal row,
  // matching the reference layout (STATE … parties … STATE TOTAL).
  const SUB_KEYS = ['pqty', 'puramt', 'discount', 'gst5', 'tds', 'payable'];
  const out = [];
  let curState = null;
  let acc = null;
  const flush = () => {
    if (!acc || curState == null) return;
    const sub = { _isSubtotal: true, poolername: `${curState} TOTAL` };
    SUB_KEYS.forEach(k => { sub[k] = acc[k] || 0; });
    out.push(sub);
  };
  for (const r of rows) {
    const st = r.state || '';
    if (st !== curState) {
      flush();
      curState = st;
      acc = Object.fromEntries(SUB_KEYS.map(k => [k, 0]));
      out.push({ poolername: st });   // state header line
    }
    SUB_KEYS.forEach(k => { acc[k] += Number(r[k]) || 0; });
    out.push(r);
  }
  flush();

  const cols = [
    { header: 'POOLER NAME', key: 'poolername', width: 34 },
    { header: 'P_QTY',     key: 'pqty',    width: 12, numFmt: '#,##0.000' },
    { header: 'P_RATE',    key: 'prate',   width: 10, align: 'right' },
    { header: 'PURAMOUNT', key: 'puramt',  width: 16, numFmt: '#,##0.00' },
    { header: 'DISCOUNT',  key: 'discount',width: 14, numFmt: '#,##0.00' },
    { header: 'GST 5%',    key: 'gst5',    width: 12, numFmt: '#,##0.00' },
    { header: 'TDS',       key: 'tds',     width: 12, numFmt: '#,##0.00' },
    { header: 'PAYABLE',   key: 'payable', width: 16, numFmt: '#,##0.00' },
  ];
  const sum = (k) => rows.reduce((s, r) => s + (Number(r[k]) || 0), 0);
  const grandTotal = {
    label: 'TOTAL',
    values: {
      pqty: sum('pqty'), puramt: sum('puramt'), discount: sum('discount'),
      gst5: sum('gst5'), tds: sum('tds'), payable: sum('payable'),
    },
  };
  return createExcelBuffer('PaymentPartywise', cols, out, {
    db, title: 'Payment Summary - Party wise', metaLines: auctionMeta(db, auctionId),
    grandTotal,
  });
}

// ── Export: TDS Return ───────────────────────────────────────
async function exportTDSReturn(db, fromDate, toDate) {
  const { getTDSReturnData } = require('./calculations');
  const rows = getTDSReturnData(db, fromDate, toDate, 'invoice');
  const cols = [
    { header: 'INVOICE', key: 'invoice', width: 10 },
    { header: 'DATE', key: 'date', width: 12 },
    { header: 'NAME', key: 'name', width: 30 },
    { header: 'PAN', key: 'pan', width: 12 },
    { header: 'ASSESS_VALUE', key: 'assess_value', width: 14 },
    { header: 'TDS', key: 'tds', width: 12 },
  ];
  return createExcelBuffer('TDSReturn', cols, rows, {
    db, title: 'TDS Return', metaLines: [`From: ${fromDate}`, `To: ${toDate}`],
  });
}

// ── Export: Tally format (TALY.PRG — purchase data for accounting)
async function exportTallyPurchase(db, auctionId, cfg) {
  const mode = (cfg && cfg.business_mode || 'e-Trade').toLowerCase();
  const discountCol = (mode === 'auction') ? 'advance' : 'refund';
  const rows = db.all(
    `SELECT name, padd as add, ppla as place, cr as gstin, tel,
      lot_no as lot, bags as bag, pqty as qty, prate as price, puramt as amount,
      cgst, sgst, igst, ${discountCol} as discount, puramt as bilamt
     FROM lots WHERE auction_id = ? AND amount > 0
      AND cr NOT LIKE 'GSTIN.%'
     ORDER BY name`, [auctionId]
  );
  const cols = [
    { header: 'NAME', key: 'name', width: 30 }, { header: 'ADD', key: 'add', width: 30 },
    { header: 'PLACE', key: 'place', width: 15 }, { header: 'GSTIN', key: 'gstin', width: 20 },
    { header: 'TEL', key: 'tel', width: 14 }, { header: 'LOT', key: 'lot', width: 8 },
    { header: 'BAG', key: 'bag', width: 6 }, { header: 'QTY', key: 'qty', width: 12 },
    { header: 'PRICE', key: 'price', width: 10 }, { header: 'AMOUNT', key: 'amount', width: 14 },
    { header: 'CGST', key: 'cgst', width: 12 }, { header: 'SGST', key: 'sgst', width: 12 },
    { header: 'IGST', key: 'igst', width: 12 }, { header: 'DISCOUNT', key: 'discount', width: 14 },
    { header: 'BILAMT', key: 'bilamt', width: 14 },
  ];
  return createExcelBuffer('TallyPurchase', cols, rows, {
    db, title: 'Tally Purchase', metaLines: auctionMeta(db, auctionId),
  });
}

// ── Export: Sales Journal (JOUR.PRG) ────────────────────────
// Trade-based: filters by auction_id; dates rendered dd/mm/yyyy.
async function exportSalesJournal(db, auctionId, saleType) {
  const { getSalesJournal } = require('./calculations');
  const rows = getSalesJournal(db, auctionId, saleType);
  const cols = [
    { header: 'DATE', key: 'date', width: 12 },
    { header: 'SALE', key: 'sale', width: 6 },
    { header: 'INV#', key: 'invo', width: 8 },
    { header: 'BUYER', key: 'buyer', width: 8 },
    { header: 'TRADE NAME', key: 'buyer1', width: 30 },
    { header: 'GSTIN', key: 'gstin', width: 20 },
    { header: 'PLACE', key: 'place', width: 15 },
    { header: 'BAGS', key: 'bag', width: 6 },
    { header: 'QTY', key: 'qty', width: 12 },
    { header: 'CARDAMOM', key: 'cardamom', width: 14 },
    { header: 'GUNNY', key: 'gunny', width: 10 },
    { header: 'TRANSPORT', key: 'transport', width: 10 },
    { header: 'INSURANCE', key: 'insurance', width: 10 },
    { header: 'CGST', key: 'cgst', width: 10 },
    { header: 'SGST', key: 'sgst', width: 10 },
    { header: 'IGST', key: 'igst', width: 10 },
    { header: 'TCS', key: 'tcs', width: 10 },
    { header: 'ROUND', key: 'rund', width: 8 },
    { header: 'TOTAL', key: 'total', width: 14 },
  ];
  return createExcelBuffer('SalesJournal', cols, rows, {
    db, title: 'Sales Journal',
    metaLines: [...auctionMeta(db, auctionId), saleType ? `Type: ${saleType}` : ''].filter(Boolean),
  });
}

// ── Export: Purchase Journal (PUJOUR.PRG / PPUJOUR.PRG) ────
// Trade-based: filters by auction_id (or ano for legacy bills);
// dates rendered dd/mm/yyyy.
async function exportPurchaseJournal(db, auctionId, type) {
  const { getPurchaseJournal } = require('./calculations');
  const rows = getPurchaseJournal(db, auctionId, type);
  const cols = type === 'agri' ? [
    { header: 'DATE', key: 'date', width: 12 },
    { header: 'BILL#', key: 'bill_no', width: 8 },
    { header: 'NAME', key: 'name', width: 30 },
    { header: 'ADDRESS', key: 'address', width: 30 },
    { header: 'PLACE', key: 'place', width: 15 },
    { header: 'STATE', key: 'state', width: 12 },
    { header: 'CR', key: 'cr', width: 15 },
    { header: 'PAN', key: 'pan', width: 12 },
    { header: 'QTY', key: 'qty', width: 12 },
    { header: 'COST', key: 'cost', width: 14 },
    { header: 'IGST', key: 'igst', width: 10 },
    { header: 'NET', key: 'net', width: 14 },
  ] : [
    { header: 'DATE', key: 'date', width: 12 },
    { header: 'INV#', key: 'invoice_no', width: 8 },
    { header: 'NAME', key: 'name', width: 30 },
    { header: 'ADDRESS', key: 'address', width: 30 },
    { header: 'PLACE', key: 'place', width: 15 },
    { header: 'STATE', key: 'state', width: 12 },
    { header: 'GSTIN', key: 'gstin', width: 20 },
    { header: 'QTY', key: 'qty', width: 12 },
    { header: 'AMOUNT', key: 'amount', width: 14 },
    { header: 'CGST', key: 'cgst', width: 10 },
    { header: 'SGST', key: 'sgst', width: 10 },
    { header: 'IGST', key: 'igst', width: 10 },
    { header: 'ROUND', key: 'rund', width: 8 },
    { header: 'TOTAL', key: 'total', width: 14 },
    { header: 'TDS', key: 'tds', width: 10 },
  ];
  const name = type === 'agri' ? 'AgriBillJournal' : 'PurchaseJournal';
  return createExcelBuffer(name, cols, rows, {
    db,
    title: type === 'agri' ? 'Agri Bill Journal' : 'Purchase Journal',
    metaLines: auctionMeta(db, auctionId),
  });
}

// Header meta lines for the Registers — trade (when scoped to one) or a
// date range (when spanning trades), plus an optional sale-type note.
function registerMeta(db, opts) {
  const lines = [];
  if (opts && opts.auctionId) lines.push(...auctionMeta(db, opts.auctionId));
  else if (opts && opts.from && opts.to) lines.push(`Period: ${opts.from} to ${opts.to}`);
  else lines.push('All trades');
  if (opts && opts.saleType) lines.push(`Sale: ${opts.saleType}`);
  return lines.filter(Boolean);
}

// ── Export: Purchase Register (lot-wise) ───────────────────
async function exportPurchaseRegister(db, opts = {}) {
  const { getPurchaseRegister } = require('./calculations');
  const rows = getPurchaseRegister(db, opts);
  const cols = [
    { header: 'STATE',  key: 'state',  width: 14 },
    { header: 'TNO',    key: 'tno',    width: 6  },
    { header: 'DATE',   key: 'date',   width: 12 },
    { header: 'LOT',    key: 'lot',    width: 8  },
    { header: 'BRANCH', key: 'branch', width: 10 },
    { header: 'NAME',   key: 'name',   width: 28 },
    { header: 'PLACE',  key: 'place',  width: 14 },
    { header: 'GSTIN',  key: 'gstin',  width: 18 },
    { header: 'BAG',    key: 'bag',    width: 6  },
    { header: 'QTY',    key: 'qty',    width: 11, numFmt: '#,##0.000' },
    { header: 'PRICE',  key: 'price',  width: 10, numFmt: '#,##0.00' },
    { header: 'AMOUNT', key: 'amount', width: 14, numFmt: '#,##0.00' },
    { header: 'PQTY',   key: 'pqty',   width: 11, numFmt: '#,##0.000' },
    { header: 'PRATE',  key: 'prate',  width: 10, numFmt: '#,##0.00' },
    { header: 'PURAMT', key: 'puramt', width: 14, numFmt: '#,##0.00' },
    { header: 'DISCOUNT', key: 'discount', width: 12, numFmt: '#,##0.00' },
    { header: 'GST5',   key: 'gst5',   width: 11, numFmt: '#,##0.00' },
    { header: 'PAYABLE', key: 'payable', width: 14, numFmt: '#,##0.00' },
  ];
  const sum = (k) => rows.reduce((s, r) => s + (Number(r[k]) || 0), 0);
  const grandTotal = { label: 'TOTAL', values: {
    bag: sum('bag'), qty: sum('qty'), amount: sum('amount'), pqty: sum('pqty'),
    puramt: sum('puramt'), discount: sum('discount'), gst5: sum('gst5'), payable: sum('payable'),
  }};
  return createExcelBuffer('PurchaseRegister', cols, rows, {
    db, title: 'Purchase Register', metaLines: registerMeta(db, opts), grandTotal,
  });
}

// ── Export: Sales Register (invoice-wise) ──────────────────
async function exportSalesRegister(db, opts = {}) {
  const { getSalesRegister } = require('./calculations');
  const rows = getSalesRegister(db, opts);
  const cols = [
    { header: 'STATE',  key: 'state',  width: 14 },
    { header: 'TNO',    key: 'tno',    width: 6  },
    { header: 'DATE',   key: 'date',   width: 12 },
    { header: 'SALE',   key: 'sale',   width: 6  },
    { header: 'INVO',   key: 'invo',   width: 8  },
    { header: 'TRADERNAME', key: 'tradername', width: 30 },
    { header: 'BIDDER', key: 'bidder', width: 10 },
    { header: 'BAG',    key: 'bag',    width: 6  },
    { header: 'QTY',    key: 'qty',    width: 11, numFmt: '#,##0.000' },
    { header: 'AMOUNT', key: 'amount', width: 14, numFmt: '#,##0.00' },
    { header: 'LORRY',  key: 'lorry',  width: 10, numFmt: '#,##0.00' },
    { header: 'GUNNY',  key: 'gunny',  width: 10, numFmt: '#,##0.00' },
    { header: 'IGST',   key: 'igst',   width: 10, numFmt: '#,##0.00' },
    { header: 'CGST',   key: 'cgst',   width: 10, numFmt: '#,##0.00' },
    { header: 'SGST',   key: 'sgst',   width: 10, numFmt: '#,##0.00' },
    { header: 'INS',    key: 'ins',    width: 10, numFmt: '#,##0.00' },
    { header: 'INVAMT', key: 'invamt', width: 14, numFmt: '#,##0.00' },
  ];
  const sum = (k) => rows.reduce((s, r) => s + (Number(r[k]) || 0), 0);
  const grandTotal = { label: 'TOTAL', values: {
    bag: sum('bag'), qty: sum('qty'), amount: sum('amount'), lorry: sum('lorry'),
    gunny: sum('gunny'), igst: sum('igst'), cgst: sum('cgst'), sgst: sum('sgst'),
    ins: sum('ins'), invamt: sum('invamt'),
  }};
  return createExcelBuffer('SalesRegister', cols, rows, {
    db, title: 'Sales Register', metaLines: registerMeta(db, opts), grandTotal,
  });
}

// ── Export: Per-party "Individual" Registers (cross-auction) ───────
// Pooler / Seller / Merchant statements, one section per party. Shares the
// createExcelBuffer section-grouped mode: each party becomes a banded
// section (name + GSTIN) followed by its rows, a bold TOTAL subtotal, and a
// summary line (Sold/Not Sold for poolers, Closing Balance for the others).
// `labelKey` is the first column the TOTAL/summary labels land in.
const INDIVIDUAL_REG_DEFS = {
  pooler: {
    sheet: 'PoolerRegister', title: 'Pooler Register', labelKey: 'tno',
    cols: [
      { header: 'TNO',    key: 'tno',    width: 8  },
      { header: 'DATE',   key: 'date',   width: 12 },
      { header: 'LOT',    key: 'lot',    width: 8  },
      { header: 'QTY',    key: 'qty',    width: 12, numFmt: '#,##0.000' },
      { header: 'RATE',   key: 'rate',   width: 11, numFmt: '#,##0.00'  },
      { header: 'VALUE',  key: 'value',  width: 16, numFmt: '#,##0.00'  },
      { header: 'P_QTY',  key: 'pqty',   width: 12, numFmt: '#,##0.000' },
      { header: 'P_RATE', key: 'prate',  width: 11, numFmt: '#,##0.00'  },
      { header: 'PURAMT', key: 'puramt', width: 16, numFmt: '#,##0.00'  },
    ],
    summaryRows: (p) => ([
      { _isSubtotal: true, tno: 'Total',     qty: p.summary.qty,     value: p.summary.value, pqty: p.summary.pqty, puramt: p.summary.puramt },
      { _isSubtotal: true, tno: 'Sold',      qty: p.summary.soldQty, value: p.summary.soldValue },
      { _isSubtotal: true, tno: 'Withdrawn', qty: p.summary.wdQty,   value: p.summary.wdValue },
    ]),
    grandKeys: ['qty', 'value', 'pqty', 'puramt'],
  },
  seller: {
    sheet: 'SellerRegister', title: 'Sellers Individual', labelKey: 'date',
    cols: [
      { header: 'DATE',    key: 'date',    width: 12 },
      { header: 'ANO',     key: 'ano',     width: 8  },
      { header: 'INVO',    key: 'invo',    width: 8,  numFmt: '#,##0' },
      { header: 'QTY',     key: 'qty',     width: 12, numFmt: '#,##0.000' },
      { header: 'INVOICE', key: 'invoice', width: 16, numFmt: '#,##0.00' },
    ],
    summaryRows: (p) => ([
      { _isSubtotal: true, date: 'Total',           qty: p.summary.qty, invoice: p.summary.invoice },
      { _isSubtotal: true, date: 'Closing Balance', invoice: p.summary.closing },
    ]),
    grandKeys: ['qty', 'invoice'],
  },
  merchant: {
    sheet: 'MerchantRegister', title: 'Merchants Individual', labelKey: 'date',
    cols: [
      { header: 'DATE',    key: 'date',    width: 12 },
      { header: 'TNO',     key: 'tno',     width: 8  },
      { header: 'INVO',    key: 'invo',    width: 8  },
      { header: 'RECP',    key: 'recp',    width: 8  },
      { header: 'QTY',     key: 'qty',     width: 12, numFmt: '#,##0.000' },
      { header: 'INVOICE', key: 'invoice', width: 16, numFmt: '#,##0.00' },
      { header: 'RECEIPT', key: 'receipt', width: 16, numFmt: '#,##0.00' },
    ],
    summaryRows: (p) => ([
      { _isSubtotal: true, date: 'Total',           qty: p.summary.qty, invoice: p.summary.invoice, receipt: p.summary.receipt },
      { _isSubtotal: true, date: 'Closing Balance', invoice: p.summary.closing },
    ]),
    grandKeys: ['qty', 'invoice', 'receipt'],
  },
};

function individualRegisterData(db, kind, opts) {
  const { getPoolerRegister, getSellerRegister, getMerchantRegister } = require('./calculations');
  if (kind === 'seller')   return getSellerRegister(db, opts);
  if (kind === 'merchant') return getMerchantRegister(db, opts);
  return getPoolerRegister(db, opts);
}

async function exportIndividualRegister(db, kind, opts = {}) {
  const def = INDIVIDUAL_REG_DEFS[kind];
  if (!def) throw new Error(`Unknown individual register kind: ${kind}`);
  const data = individualRegisterData(db, kind, opts);
  const sections = data.parties.map(p => ({
    title: p.name + (p.gstin ? `      GSTIN: ${p.gstin}` : ''),
    rows: [...p.rows, ...def.summaryRows(p)],
  }));
  // Grand total across every party in the file.
  const gv = {};
  def.grandKeys.forEach(k => {
    gv[k] = data.parties.reduce((s, p) => s + (Number(p.summary[k]) || 0), 0);
  });
  gv[def.labelKey] = 'GRAND TOTAL';
  return createExcelBuffer(def.sheet, def.cols, [], {
    db, title: def.title, metaLines: registerMeta(db, opts),
    sections, spacerBetween: true,
    grandTotal: { values: gv },
  });
}

// ── Export: Praman CSV (Lot Slip in Praman auction platform format) ──
// Produces a CSV (NOT xlsx) matching the column layout required by Praman's
// lot-upload interface. Returns a Buffer of CSV text.
//
// "Lot Company" column (col 2) is the registered Praman uploader identity —
// resolved from company_settings (`short_name` → `logo` short code). NO
// hardcoded fallback: if neither is configured the cell is left blank,
// surfacing the misconfiguration rather than leaking a stale literal.
async function exportPramanCSV(db, auctionId, cfg, state) {
  // Praman expects PER-LOT planter info — the seller (and their GSTIN)
  // for each individual lot, not a single legal-entity stamp on every
  // row. The earlier export used `getCompanyIdentity(cfg)` and wrote
  // the company's own name + GSTIN on every row, which surfaced as
  // "VANDANMEDU SPICES" (or whatever trade_name was set) repeated for
  // every lot — wrong for the Praman upload, which uses these fields
  // to identify each lot's actual seller.
  //
  // Fix: pull lots.name (seller per lot) and the trader's `cr`
  // (stored as the GSTIN). Falls back to the company identity ONLY if
  // a lot has no associated seller record (legacy data, partial
  // imports).
  // Join the trader row PER LOT by trader_id (the FK each lot stores
  // when it's created from the seller picker). Joining by name —
  // which the older query did — multiplies rows whenever a seller
  // has more than one row in `traders` (legitimate multi-branch
  // sellers, or accidental dupes), producing duplicate lots in the
  // Praman CSV. The name-based fallback subquery only kicks in for
  // legacy rows that pre-date the trader_id column.
  const rows = db.all(
    `SELECT l.lot_no, l.branch, l.grade, l.name, l.cr, l.qty, l.litre, l.bags, l.tel, l.reserved_price,
            COALESCE(
              t.cr,
              (SELECT cr  FROM traders WHERE UPPER(TRIM(name)) = UPPER(TRIM(l.name)) LIMIT 1)
            ) AS trader_cr,
            COALESCE(
              t.tel,
              (SELECT tel FROM traders WHERE UPPER(TRIM(name)) = UPPER(TRIM(l.name)) LIMIT 1)
            ) AS trader_tel
       FROM lots l
       LEFT JOIN traders t ON t.id = l.trader_id
      WHERE l.auction_id = ? ${state ? 'AND l.state = ?' : ''}
      ORDER BY CAST(l.lot_no AS INTEGER), l.lot_no`,
    state ? [auctionId, state] : [auctionId]
  );

  const header = [
    'Lot Number', 'Lot Company', 'Collection Centre', 'Planter/Dealer',
    'Planter Name', 'CRNO/SBL No', 'Quantity(Kg)', 'Litre Weight(Gms)',
    'Bags', 'Grade Type', 'Grade', 'Reserved Price', 'Auction Start Price(Rs)',
    'Immature Seeds(%)', 'Moisture Content(%)', 'Planter Mobile Number',
    'Youtube Video Link'
  ];

  // Escape a CSV field: wrap in quotes if it contains comma/quote/newline,
  // and double-up any embedded quotes. Undefined/null → empty.
  const csvEscape = (v) => {
    if (v === null || v === undefined) return '';
    const s = String(v);
    if (/[",\n\r]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
    return s;
  };

  // GSTIN extractor — `cr` may be stored as "GSTIN.<15>", "gstin.<15>",
  // bare 15-char, or empty. Strip the prefix if present.
  const stripGstinPrefix = (raw) => {
    let s = String(raw || '').trim();
    if (/^gstin\.?/i.test(s)) s = s.replace(/^gstin\.?/i, '');
    return s;
  };

  // Lot company short code — first check for a dedicated Praman value
  // (Settings → Integrations → Praman Lot Company Code). If unset,
  // fall back to the company-wide `short_name` (Settings → Company →
  // Short Name) via the identity resolver. This split lets the user
  // register a different short code with Praman than what they use
  // elsewhere (e.g. invoice prefixes, logo derivation) without
  // touching every other code path.
  const identity = getCompanyIdentity(cfg);
  const lotCompany = String(cfg.praman_company || '').trim() || identity.shortName || '';

  // Praman classifies sellers as 1=Planter (URD/agriculturist) or
  // 2=Dealer (registered, with GSTIN). Per-lot decision based on
  // whether the seller has a GSTIN attached.
  const classify = (gstin) => (gstin && gstin.length >= 15) ? 2 : 1;

  const lines = [header.join(',')];
  for (const r of rows) {
    // Per-lot planter: name from lots.name, GSTIN from trader's `cr`
    // (preferred — master data) with the lot's own `cr` as a fallback
    // when traders join misses.
    const planterName   = (r.name || '').trim();
    const planterGstin  = stripGstinPrefix(r.trader_cr || r.cr);
    const planterMobile = (r.trader_tel || r.tel || '').trim();
    const planterDealer = classify(planterGstin);

    lines.push([
      r.lot_no || '',
      lotCompany,
      r.branch || '',
      planterDealer,
      planterName,
      planterGstin,
      r.qty || '',
      r.litre || '',
      r.bags || '',
      '', // Grade Type (not captured — blank as per sample)
      '', // Grade (Praman's own grade codes, not ours — blank)
      Number(r.reserved_price) > 0 ? Number(r.reserved_price).toFixed(2) : '', // Reserved Price (seller's minimum, per lot)
      '', // Auction Start Price (blank)
      '', // Immature Seeds (blank)
      '', // Moisture Content (blank)
      planterMobile,
      '', // Youtube link (blank)
    ].map(csvEscape).join(','));
  }

  // CSV text → Buffer. Prefix with BOM so Excel on Windows opens with
  // UTF-8 correctly (otherwise accented characters break).
  return Buffer.from('\uFEFF' + lines.join('\r\n'), 'utf8');
}

// ── Export Type 12: Trade Report (BUYERS LIST FOR VERIFICATION) ──
async function exportTradeReport(db, auctionId) {
  return tradeReportXlsx(db, auctionId);
}

// ── Lot Buyer / Lot Name / Lot Payment ──────────────────────
// Three pre-printed auction-floor sheets. Each shares the lot list
// for the auction; the difference is which columns are filled in vs.
// left blank for hand entry on the day:
//   • Lot Buyer   → BUYER blank (auction not yet held)
//   • Lot Name    → NAME pre-filled (seller), PRICE + CONTROL blank
//   • Lot Payment → fully populated post-auction payment summary
// `br` is derived from the lot's state column (KERALA → KL,
// TAMIL NADU → TN) to match the printed-form layout the user
// already uses.
const BR_FROM_STATE_SQL = `
  CASE UPPER(COALESCE(state,''))
    WHEN 'KERALA' THEN 'KL'
    WHEN 'TAMIL NADU' THEN 'TN'
    ELSE UPPER(SUBSTR(COALESCE(state,''), 1, 2))
  END`;

async function exportLotBuyer(db, auctionId) {
  const rows = db.all(
    `SELECT lot_no AS lot, COALESCE(buyer,'') AS buyer,
            ${BR_FROM_STATE_SQL} AS br,
            bags AS bag, qty
     FROM lots WHERE auction_id = ? ORDER BY lot_no`,
    [auctionId]
  );
  const cols = [
    { header: 'LOT',   key: 'lot',   width: 8  },
    { header: 'BUYER', key: 'buyer', width: 24 },
    { header: 'BR',    key: 'br',    width: 6  },
    { header: 'BAG',   key: 'bag',   width: 6  },
    { header: 'QTY',   key: 'qty',   width: 12 },
  ];
  return createExcelBuffer('LotBuyer', cols, rows, {
    db, title: 'Lot Buyer', metaLines: auctionMeta(db, auctionId),
  });
}

async function exportLotName(db, auctionId) {
  // PRICE is blanked when 0 (lot not yet priced) so the column reads
  // empty rather than "0.00" — matching exportPriceListBefore.
  const rows = db.all(
    `SELECT lot_no AS lot, COALESCE(name,'') AS name,
            ${BR_FROM_STATE_SQL} AS br,
            bags AS bag, qty,
            CASE WHEN COALESCE(price,0) = 0 THEN '' ELSE price END AS price,
            '' AS control
     FROM lots WHERE auction_id = ? ORDER BY lot_no`,
    [auctionId]
  );
  const cols = [
    { header: 'LOT',     key: 'lot',     width: 8  },
    { header: 'NAME',    key: 'name',    width: 30 },
    { header: 'BR',      key: 'br',      width: 6  },
    { header: 'BAG',     key: 'bag',     width: 6  },
    { header: 'QTY',     key: 'qty',     width: 12 },
    { header: 'PRICE',   key: 'price',   width: 10 },
    { header: 'CONTROL', key: 'control', width: 12 },
  ];
  return createExcelBuffer('LotName', cols, rows, {
    db, title: 'Lot Name', metaLines: auctionMeta(db, auctionId),
  });
}

async function exportLotPayment(db, auctionId) {
  // Order by branch then name so the natural printed layout (branch
  // header followed by that branch's lots) emerges from the row
  // sequence, even though the generic table renderer doesn't draw
  // a separator row between branches.
  const rows = db.all(
    `SELECT COALESCE(branch,'') AS branch,
            lot_no AS lot, qty, price AS rate, amount AS cost,
            pqty, prate, puramt AS purchamt,
            COALESCE(name,'') AS seller_name
     FROM lots WHERE auction_id = ? ORDER BY branch, name, lot_no`,
    [auctionId]
  );
  const cols = [
    { header: 'BRANCH',      key: 'branch',      width: 14 },
    { header: 'LOT',         key: 'lot',         width: 6  },
    { header: 'QTY',         key: 'qty',         width: 10 },
    { header: 'RATE',        key: 'rate',        width: 10 },
    { header: 'COST',        key: 'cost',        width: 14 },
    { header: 'PQTY',        key: 'pqty',        width: 10 },
    { header: 'PRATE',       key: 'prate',       width: 10 },
    { header: 'PURCHAMT',    key: 'purchamt',    width: 14 },
    { header: 'SELLER NAME', key: 'seller_name', width: 26 },
  ];
  return createExcelBuffer('LotPayment', cols, rows, {
    db, title: 'Lot Payment Summary', metaLines: auctionMeta(db, auctionId),
  });
}

// ── Export: Dealer List — Party wise ─────────────────────────
// Same registered-dealer set as Dealer List (sellers whose `cr` resolves
// to a clean 15-char GSTIN), but presented PARTY-WISE: one row per dealer,
// grouped under each STATE with a per-state subtotal and a grand total.
// Mirrors exportPaymentPartywise's presentation (state header line +
// "<STATE> TOTAL" subtotal + GRAND TOTAL).
async function exportDealerListPartywise(db, auctionId) {
  const rows = db.all(
    `WITH cleaned AS (
       SELECT state, name, lot_no, bags, qty,
              UPPER(TRIM(
                CASE
                  WHEN LOWER(SUBSTR(TRIM(cr),1,5)) = 'gstin'
                    THEN LTRIM(SUBSTR(TRIM(cr),6), '. :-')
                  ELSE TRIM(cr)
                END
              )) AS gstin
         FROM lots
        WHERE auction_id = ?
     )
     SELECT state, name, gstin,
            COUNT(lot_no) as lots, SUM(bags) as bags, SUM(qty) as qty
       FROM cleaned
      WHERE LENGTH(gstin) = 15
      GROUP BY state, name, gstin
      ORDER BY state, name`, [auctionId]
  );
  const SUB_KEYS = ['lots', 'bags', 'qty'];
  const out = [];
  let curState = null, acc = null;
  const flush = () => {
    if (!acc || curState == null) return;
    const sub = { _isSubtotal: true, name: `${curState} TOTAL` };
    SUB_KEYS.forEach(k => { sub[k] = acc[k] || 0; });
    out.push(sub);
  };
  for (const r of rows) {
    const st = r.state || '';
    if (st !== curState) {
      flush();
      curState = st;
      acc = Object.fromEntries(SUB_KEYS.map(k => [k, 0]));
      out.push({ name: st });   // state header line
    }
    SUB_KEYS.forEach(k => { acc[k] += Number(r[k]) || 0; });
    out.push(r);
  }
  flush();
  const cols = [
    { header: 'NAME',  key: 'name',  width: 34 },
    { header: 'GSTIN', key: 'gstin', width: 18 },
    { header: 'LOTS',  key: 'lots',  width: 8 },
    { header: 'BAGS',  key: 'bags',  width: 8 },
    { header: 'QTY',   key: 'qty',   width: 12, numFmt: '#,##0.000' },
  ];
  const sum = (k) => rows.reduce((s, r) => s + (Number(r[k]) || 0), 0);
  const grandTotal = { label: 'GRAND TOTAL', values: { lots: sum('lots'), bags: sum('bags'), qty: sum('qty') } };
  return createExcelBuffer('DealerListPartywise', cols, out, {
    db, title: 'Dealer List - Party wise', metaLines: auctionMeta(db, auctionId), grandTotal,
  });
}

// ── Export: Pooler List Consolidated — Party wise ────────────
// Consolidates the Pooler Register (one row per lot) into one aggregated
// row PER POOLER, grouped by STATE with a per-state subtotal + grand total.
// PRICE / PRATE are intentionally omitted — a party spans many lots at
// different rates, so a party-level rate is meaningless. Non-withdrawn
// lots only (matches the Pooler Register's `code != 'WD'` filter).
async function exportPoolerListConsolidated(db, auctionId) {
  const rows = db.all(
    `SELECT state, name as poolername,
            COUNT(lot_no) as lots, SUM(qty) as qty, SUM(amount) as amount,
            SUM(pqty) as pqty, SUM(puramt) as puramt
       FROM lots
      WHERE auction_id = ? AND UPPER(TRIM(COALESCE(code,''))) != 'WD'
      GROUP BY state, name
      ORDER BY state, name`, [auctionId]
  );
  const SUB_KEYS = ['lots', 'qty', 'amount', 'pqty', 'puramt'];
  const out = [];
  let curState = null, acc = null;
  const flush = () => {
    if (!acc || curState == null) return;
    const sub = { _isSubtotal: true, poolername: `${curState} TOTAL` };
    SUB_KEYS.forEach(k => { sub[k] = acc[k] || 0; });
    out.push(sub);
  };
  for (const r of rows) {
    const st = r.state || '';
    if (st !== curState) {
      flush();
      curState = st;
      acc = Object.fromEntries(SUB_KEYS.map(k => [k, 0]));
      out.push({ poolername: st });   // state header line
    }
    SUB_KEYS.forEach(k => { acc[k] += Number(r[k]) || 0; });
    out.push(r);
  }
  flush();
  const cols = [
    { header: 'POOLER NAME', key: 'poolername', width: 34 },
    { header: 'LOTS',   key: 'lots',   width: 8 },
    { header: 'QTY',    key: 'qty',    width: 12, numFmt: '#,##0.000' },
    { header: 'AMOUNT', key: 'amount', width: 16, numFmt: '#,##0.00' },
    { header: 'PQTY',   key: 'pqty',   width: 12, numFmt: '#,##0.000' },
    { header: 'PURAMT', key: 'puramt', width: 16, numFmt: '#,##0.00' },
  ];
  const sum = (k) => rows.reduce((s, r) => s + (Number(r[k]) || 0), 0);
  const grandTotal = { label: 'GRAND TOTAL', values: {
    lots: sum('lots'), qty: sum('qty'), amount: sum('amount'), pqty: sum('pqty'), puramt: sum('puramt'),
  } };
  return createExcelBuffer('PoolerListConsolidated', cols, out, {
    db, title: 'Pooler List Consolidated - Party wise', metaLines: auctionMeta(db, auctionId), grandTotal,
  });
}

// Render an already-built XLSX buffer as a simple HTML table. Used by the
// Reports "Preview" action for XLSX-ONLY reports (the ones with no PDF
// renderer — Full File, Voucher Payment, Bank Payment Before) so the
// operator still gets an on-screen look. It reads the exact cells the
// download contains, so the preview matches the file faithfully — including
// any header/brand bands the generator added.
async function xlsxBufferToHtml(buffer) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  const ws = wb.worksheets[0];
  if (!ws) return '<div style="padding:16px;color:#666">Nothing to preview.</div>';
  const esc = s => String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const cellText = (cell) => {
    const v = cell.value;
    if (v == null) return '';
    if (typeof v !== 'object') return String(v);
    if (Array.isArray(v.richText)) return v.richText.map(t => t.text).join('');
    if (v.text != null) return String(v.text);
    if (v.result != null) return String(v.result);
    if (v.hyperlink != null) return String(v.hyperlink);
    if (v instanceof Date) return v.toISOString().slice(0, 10);
    return '';
  };
  const maxCol = ws.columnCount || 1;
  let html = '<table style="border-collapse:collapse;width:100%;font-size:12px;font-family:Arial,Helvetica,sans-serif">';
  ws.eachRow({ includeEmpty: false }, (row) => {
    html += '<tr>';
    for (let c = 1; c <= maxCol; c++) {
      const cell = row.getCell(c);
      const text = cellText(cell);
      const bold = !!(cell.font && cell.font.bold);
      const num = typeof cell.value === 'number';
      const align = (cell.alignment && cell.alignment.horizontal) || (num ? 'right' : 'left');
      html += `<td style="border:1px solid #d5d5d5;padding:3px 7px;text-align:${align};${bold ? 'font-weight:700;' : ''}white-space:nowrap">${esc(text)}</td>`;
    }
    html += '</tr>';
  });
  html += '</table>';
  return html;
}

// ── Export router ────────────────────────────────────────────
const EXPORT_TYPES = {
  lot_slip:           { fn: exportLotSlip,           name: 'LotSlip' },
  lot_slip_after:     { fn: exportLotSlipAfter,      name: 'LotSlipAfter' },
  lot_buyer:          { fn: exportLotBuyer,          name: 'LotBuyer' },
  lot_name:           { fn: exportLotName,           name: 'LotName' },
  lot_payment:        { fn: exportLotPayment,        name: 'LotPayment' },
  praman_csv:         { fn: exportPramanCSV,         name: 'eTrade_Praman', ext: 'csv', mime: 'text/csv', needsCfg: true },
  price_list:         { fn: exportPriceList,         name: 'PriceList' },
  price_list_before:  { fn: exportPriceListBefore,   name: 'PriceListBefore' },
  bank_payment_before:{ fn: exportBankPaymentBefore, name: 'BankPaymentBefore', needsCfg: true },
  bank_payment:       { fn: exportBankPayment,       name: 'BankPayment',       needsCfg: true },
  voucher_payment:    { fn: exportVoucherPayment,    name: 'VoucherPayment',    needsCfg: true },
  pooler_register:    { fn: exportPoolerRegister,    name: 'PoolerRegister' },
  pooler_list_consolidated: { fn: exportPoolerListConsolidated, name: 'PoolerListConsolidated' },
  full_file:          { fn: exportFullFile,          name: 'FullFile' },
  collection:         { fn: exportCollection,        name: 'Collection' },
  trade_report:       { fn: exportTradeReport,       name: 'TradeReport' },
  dealer_list:        { fn: exportDealerList,        name: 'DealerList' },
  dealer_list_partywise: { fn: exportDealerListPartywise, name: 'DealerListPartywise' },
  sales_taxes:        { fn: exportSalesTaxes,        name: 'SalesTaxes' },
  payment:            { fn: exportPaymentSummary,    name: 'Payment',           needsCfg: true },
  payment_partywise:  { fn: exportPaymentPartywise,  name: 'PaymentPartywise',  needsCfg: true },
  tally_purchase:     { fn: exportTallyPurchase,     name: 'TallyPurchase',     needsCfg: true },
};

module.exports = {
  EXPORT_TYPES,
  // Reusable XLSX builder — exposed so other modules (lorry-reports.js etc.)
  // can route through the same standardized brand band + column-header
  // styling instead of building their own ExcelJS workbook.
  createExcelBuffer,
  xlsxBufferToHtml,
  exportLotSlip, exportLotSlipAfter, exportLotBuyer, exportLotName, exportLotPayment,
  exportPramanCSV, exportPriceList, exportPriceListBefore,
  exportBankPayment, exportBankPaymentBefore, exportVoucherPayment,
  exportPoolerRegister, exportPoolerListConsolidated, exportFullFile, exportCollection, exportTradeReport,
  exportDealerList, exportDealerListPartywise,
  exportSalesTaxes, exportPaymentSummary, exportPaymentPartywise, exportTDSReturn, exportTallyPurchase,
  exportSalesJournal, exportPurchaseJournal,
  exportPurchaseRegister, exportSalesRegister,
  exportIndividualRegister,
};
