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

// ── Export Type 4: Bank Payment (RTGS/NEFT — bank-import format) ─
// Bare 13-column sheet (no brand band, no totals) matching the bank's
// upload template:
//   TRANSACT_A | MESSAGETYP | DEBITACCOU | PAYMENTAMO | TRANSACT_B |
//   VALUEDATE  | BENEFICI_A | BENEFIARYN | BENEFIARYB | BENEFIARYE |
//   BENEFICI_B | REMARKS    | CLIENTCODE
// Header on row 1, data from row 2. Bank software auto-ingests this
// shape — adding a brand band would break the import.
async function exportBankPayment(db, auctionId, cfg, _state, opts) {
  const { getBankPaymentData, formatLotList } = require('./calculations');
  let payments = getBankPaymentData(db, auctionId, cfg);
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
                GROUP_CONCAT(l.lot_no) AS lot_nos
           FROM lots l
          WHERE l.auction_id = ? AND l.amount > 0
            AND (l.paid IS NULL OR l.paid = '')
            AND UPPER(TRIM(l.name)) = ?${extraWhere}`,
        params
      ) || { payable: 0, puramt: 0, lot_nos: '' };
      const rawAmount = Number(sub.payable) || 0;
      const roundedAmount = cfg.flag_round ? Math.round(rawAmount) : rawAmount;
      const isRTGS = roundedAmount >= 200000;
      payments[idx] = {
        ...payments[idx],
        amount: roundedAmount,
        transactionType: isRTGS ? 'RTGS' : 'NEFT',
        // Re-derive the covered-lots list from the same picked/excluded
        // subset so REMARKS lists exactly the lots this row pays for.
        lots: formatLotList(sub.lot_nos),
      };
    }
    // Drop any zero-amount rows produced by the recompute — banks reject
    // zero-value RTGS rows, and a seller whose remaining lots all net to
    // zero (everything already exported, or only zero-balance lots
    // picked) shouldn't appear at all.
    payments = payments.filter(p => Number(p.amount) > 0);
  }

  // Sender-side context (state-aware): debit account, IFSC for BT/LBT
  // detection, and the email used in BENEFIARYE.
  const isKL = String(cfg.business_state || cfg.state || '').toUpperCase().includes('KERALA');
  const senderAcct  = (isKL ? cfg.bank_kl_acct  : cfg.bank_tn_acct)  || cfg.bank_tn_acct  || cfg.bank_kl_acct  || '';
  const senderIfsc  = (isKL ? cfg.bank_kl_ifsc  : cfg.bank_tn_ifsc)  || cfg.bank_tn_ifsc  || cfg.bank_kl_ifsc  || '';
  const senderEmail = (isKL ? cfg.kl_email      : cfg.tn_email)      || cfg.tn_email      || cfg.kl_email      || '';
  const senderBankPrefix = String(senderIfsc).slice(0, 4).toUpperCase();
  // Short tag inserted into REMARKS (e.g. "VSTL" → "5 ANN MARIA SPICES VSTL PAYMENT 5945275.00 Credited").
  // Falls back to the leading word of trade_name when short_name isn't set.
  const shortTag = String(cfg.short_name || (cfg.trade_name || '').split(/\s+/)[0] || '').toUpperCase();

  // Auction context: ano (REMARKS prefix) + value date (DD/MM/YYYY).
  const a = db.get('SELECT ano, date FROM auctions WHERE id = ?', [auctionId]) || {};
  const ano = a.ano || '';
  const valueDate = fmtUserDate(String(a.date || '').slice(0, 10));

  const rows = payments.map(p => {
    const amount = Number(p.amount) || 0;
    const beneIfsc = String(p.ifsc || '').toUpperCase();
    const benePrefix = beneIfsc.slice(0, 4);
    // BT  = book transfer (same bank as sender)
    // LBT = local bank transfer (different bank, RTGS/NEFT routed)
    const transactA = (senderBankPrefix && benePrefix === senderBankPrefix) ? 'BT' : 'LBT';
    return {
      TRANSACT_A:  transactA,
      MESSAGETYP:  p.transactionType || 'RTGS',
      DEBITACCOU:  senderAcct,
      PAYMENTAMO:  amount,
      TRANSACT_B:  'INR',
      VALUEDATE:   valueDate,
      BENEFICI_A:  p.accountNo || '',
      BENEFIARYN:  String(p.beneficiaryName || '').toUpperCase(),
      BENEFIARYB:  beneIfsc,
      BENEFIARYE:  senderEmail,
      BENEFICI_B:  '',
      REMARKS:     `${ano} ${String(p.beneficiaryName || '').toUpperCase()}${shortTag ? ' ' + shortTag : ''} PAYMENT ${amount.toFixed(2)} Credited${p.lots ? ` for lot${p.lots.includes(',') ? 's' : ''} ${p.lots}` : ''}`,
      CLIENTCODE:  '',
    };
  });

  // Build the sheet directly (bypass createExcelBuffer's brand-band).
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('BANK_PAYMENT');
  const cols = [
    { key: 'TRANSACT_A',  header: 'TRANSACT_A',  width: 12 },
    { key: 'MESSAGETYP',  header: 'MESSAGETYP',  width: 12 },
    { key: 'DEBITACCOU',  header: 'DEBITACCOU',  width: 18 },
    { key: 'PAYMENTAMO',  header: 'PAYMENTAMO',  width: 14, numFmt: '#,##0.00' },
    { key: 'TRANSACT_B',  header: 'TRANSACT_B',  width: 10 },
    { key: 'VALUEDATE',   header: 'VALUEDATE',   width: 12 },
    { key: 'BENEFICI_A',  header: 'BENEFICI_A',  width: 22 },
    { key: 'BENEFIARYN',  header: 'BENEFIARYN',  width: 32 },
    { key: 'BENEFIARYB',  header: 'BENEFIARYB',  width: 16 },
    { key: 'BENEFIARYE',  header: 'BENEFIARYE',  width: 28 },
    { key: 'BENEFICI_B',  header: 'BENEFICI_B',  width: 12 },
    { key: 'REMARKS',     header: 'REMARKS',     width: 60 },
    { key: 'CLIENTCODE',  header: 'CLIENTCODE',  width: 14 },
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

// ── Export Type 4b: Bank Payment (Before discount) ───────────
// Same data shape as bank_payment except `amount` is the pre-discount
// puramt (raw purchase amount before refund/GST). Per the e-Trade spec
// the Amount + SendertoRcvrInfo columns are omitted from this variant.
async function exportBankPaymentBefore(db, auctionId, cfg) {
  const { getBankPaymentData } = require('./calculations');
  const payments = getBankPaymentData(db, auctionId, cfg, { before: true });
  const cols = [
    { header: 'TransactionType', key: 'transactionType', width: 16 },
    { header: 'BeneIFSCode',     key: 'ifsc',            width: 14 },
    { header: 'BeneAcctNo',      key: 'accountNo',       width: 20 },
    { header: 'BeneName',        key: 'beneficiaryName', width: 30 },
    { header: 'BeneAddLine1',    key: 'address1',        width: 30 },
    { header: 'BeneAddLine2',    key: 'address2',        width: 20 },
    { header: 'BeneAddLine3',    key: 'pin',             width: 10 },
  ];
  return createExcelBuffer('BankPaymentBefore', cols, payments, {
    db, title: 'Bank Payment (Before)', metaLines: auctionMeta(db, auctionId),
  });
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
    `SELECT name as poolername, lot_no as lot, bags as bag, qty, price, amount,
      pqty, prate, puramt, ${discountCol} as lot_discount, balance as payable
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
    const manualDisc = (!seenSellers.has(r.poolername))
      ? (Number(debitMap[r.poolername]) || 0)
      : 0;
    seenSellers.add(r.poolername);
    return {
      ...r,
      discount: lotDisc + manualDisc,
      payable: (Number(r.payable) || 0) - manualDisc,
    };
  });

  // Interleave per-pooler subtotal rows after each name group — mirrors
  // the PDF's groupByKey:'poolername' subtotalKeys behaviour. Rows are
  // already sorted by (state, name) so a single linear pass groups them.
  const SUB_KEYS = ['bag', 'qty', 'amount', 'pqty', 'puramt', 'discount', 'payable'];
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
      payable: sum('payable'),
    },
  };
  return createExcelBuffer('Payment', cols, enriched, {
    db, title: 'Payment Summary', metaLines: auctionMeta(db, auctionId),
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
  const rows = db.all(
    `SELECT lot_no AS lot, COALESCE(name,'') AS name,
            ${BR_FROM_STATE_SQL} AS br,
            bags AS bag, qty, price, '' AS control
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
  pooler_register:    { fn: exportPoolerRegister,    name: 'PoolerRegister' },
  full_file:          { fn: exportFullFile,          name: 'FullFile' },
  collection:         { fn: exportCollection,        name: 'Collection' },
  trade_report:       { fn: exportTradeReport,       name: 'TradeReport' },
  dealer_list:        { fn: exportDealerList,        name: 'DealerList' },
  sales_taxes:        { fn: exportSalesTaxes,        name: 'SalesTaxes' },
  payment:            { fn: exportPaymentSummary,    name: 'Payment',           needsCfg: true },
  tally_purchase:     { fn: exportTallyPurchase,     name: 'TallyPurchase',     needsCfg: true },
};

module.exports = {
  EXPORT_TYPES,
  // Reusable XLSX builder — exposed so other modules (lorry-reports.js etc.)
  // can route through the same standardized brand band + column-header
  // styling instead of building their own ExcelJS workbook.
  createExcelBuffer,
  exportLotSlip, exportLotSlipAfter, exportLotBuyer, exportLotName, exportLotPayment,
  exportPramanCSV, exportPriceList, exportPriceListBefore,
  exportBankPayment, exportBankPaymentBefore,
  exportPoolerRegister, exportFullFile, exportCollection, exportTradeReport, exportDealerList,
  exportSalesTaxes, exportPaymentSummary, exportTDSReturn, exportTallyPurchase,
  exportSalesJournal, exportPurchaseJournal,
};
