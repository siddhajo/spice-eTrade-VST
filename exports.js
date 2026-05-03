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

// Build an XLSX buffer with a unified brand band on top and Indian-format
// numeric columns. `opts.title` is the report title shown in the middle of
// the band; `opts.metaLines` is an array of right-aligned meta strings
// (e.g. ["Trade #3", "15/04/2026", "ASP"]).
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

  // Brand band: logo + name + address (left), title (middle), meta (right).
  const header = opts.companyHeader || getCompanyHeader(opts.db);
  const startRow = writeXlsxCompanyHeader(wb, ws, header, {
    colCount: columns.length,
    title: opts.title || sheetName,
    metaLines: opts.metaLines || [],
  });

  // Column-header row (right after the brand band, with the spacer row).
  const headerRow = ws.getRow(startRow);
  columns.forEach((c, i) => {
    headerRow.getCell(i + 1).value = c.header;
  });
  headerRow.font = { bold: true, size: 10 };
  headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8E4DD' } };
  headerRow.eachCell((cell) => {
    cell.border = { top: { style: 'thin' }, bottom: { style: 'thin' } };
    cell.alignment = { horizontal: 'center' };
  });

  // Resolve per-column numFmt + alignment ONCE so we can apply them to
  // both data rows and the grand-total row uniformly.
  //
  // Default alignment policy (matches Lorry export):
  //   - explicit `align` wins
  //   - numeric columns (have a numFmt) → right
  //   - everything else → left (cell default; we don't set explicitly)
  const colMeta = columns.map(c => {
    const fmt = c.numFmt || xlsxNumFmtForHeader(c.header);
    const align = c.align || (fmt ? 'right' : null);
    return { fmt, align };
  });

  // Apply numFmt at the worksheet-column level so every data row picks it
  // up automatically. Alignment is also applied per-column for consistency.
  colMeta.forEach((m, i) => {
    const colObj = ws.getColumn(i + 1);
    if (m.fmt) colObj.numFmt = m.fmt;
    if (m.align) colObj.alignment = { horizontal: m.align };
  });

  // Helper to emit a single data row honouring numeric coercion + per-col align
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
      // Per-cell alignment guard — column-level alignment doesn't always
      // win over default cell alignment in some ExcelJS versions, so set
      // it on the cell too when explicit.
      if (colMeta[i].align) cell.alignment = { horizontal: colMeta[i].align };
    });
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
      titleRow.font = { bold: true, size: 11 };
      titleRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD4EDDA' } };
      titleRow.alignment = { horizontal: 'center' };
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
    const fill = gt.fillArgb || 'FFFFF3CD';
    gRow.eachCell((cell, ci) => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fill } };
      cell.border = { top: { style: 'double' }, bottom: { style: 'double' } };
      const m = colMeta[ci - 1];
      if (m && m.fmt)   cell.numFmt = m.fmt;
      if (m && m.align) cell.alignment = { horizontal: m.align };
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
    const dt = String(a.date || '').slice(0, 10).split('-').reverse().join('/');
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

// ── Export Type 4: Bank Payment (RTGS/NEFT format) ───────────
async function exportBankPayment(db, auctionId, cfg) {
  const { getBankPaymentData } = require('./calculations');
  const payments = getBankPaymentData(db, auctionId, cfg);
  const cols = [
    { header: 'TransactionType', key: 'transactionType', width: 16 },
    { header: 'BeneIFSCode', key: 'ifsc', width: 14 },
    { header: 'BeneAcctNo', key: 'accountNo', width: 20 },
    { header: 'BeneName', key: 'beneficiaryName', width: 30 },
    { header: 'BeneAddLine1', key: 'address1', width: 30 },
    { header: 'BeneAddLine2', key: 'address2', width: 20 },
    { header: 'BeneAddLine3', key: 'pin', width: 10 },
    { header: 'Amount', key: 'amount', width: 14 },
    { header: 'SendertoRcvrInfo', key: 'remarks', width: 50 },
  ];
  return createExcelBuffer('BankPayment', cols, payments, {
    db, title: 'Bank Payment (RTGS/NEFT)', metaLines: auctionMeta(db, auctionId),
  });
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
  const rows = db.all(
    `SELECT state, lot_no as lot, name as poolername, branch as br, qty, price, amount, pqty, prate, puramt
     FROM lots WHERE auction_id = ? AND amount > 0
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
async function exportDealerList(db, auctionId) {
  const rows = db.all(
    `SELECT state, name, SUBSTR(cr, 7, 15) as gstin, 
      COUNT(lot_no) as lots, SUM(bags) as bags, SUM(qty) as qty
     FROM lots WHERE auction_id = ? AND cr LIKE '%GST%' AND amount > 0
     GROUP BY state, name, cr ORDER BY state, name`, [auctionId]
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
async function exportPaymentSummary(db, auctionId, cfg) {
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
  const rows = db.all(
    `SELECT name as poolername, lot_no as lot, bags as bag, qty, price, amount,
      pqty, prate, puramt, ${discountCol} as lot_discount, balance as payable
     FROM lots WHERE auction_id = ? AND amount > 0
     ORDER BY state, name`, [auctionId]
  );
  // Spread debit_notes amount across the seller's lots proportionally so
  // every row totals to the same SUM as the payments view. Simpler approach:
  // attribute the FULL manual debit on the FIRST row for each seller; later
  // rows show only the lot policy discount. Avoids per-row arithmetic but
  // still preserves the seller-level total.
  const seenSellers = new Set();
  const enriched = rows.map(r => {
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
  const cols = [
    { header: 'POOLERNAME', key: 'poolername', width: 30 },
    { header: 'LOT', key: 'lot', width: 8 }, { header: 'BAG', key: 'bag', width: 6 },
    { header: 'QTY', key: 'qty', width: 12 }, { header: 'PRICE', key: 'price', width: 10 },
    { header: 'AMOUNT', key: 'amount', width: 14 }, { header: 'PQTY', key: 'pqty', width: 12 },
    { header: 'PRATE', key: 'prate', width: 10 }, { header: 'PURAMT', key: 'puramt', width: 14 },
    { header: 'DISCOUNT', key: 'discount', width: 14 },
    { header: 'PAYABLE', key: 'payable', width: 14 },
  ];
  return createExcelBuffer('Payment', cols, enriched, {
    db, title: 'Payment Summary', metaLines: auctionMeta(db, auctionId),
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
// Special rule (item #9): Grade 1 lots → Lot Company = 'ASP' on the CSV
// output only (doesn't change stored data). All other grades → 'ISPL'.
// Rationale: Grade 1 (pooler) lots are routed to ASP for tax/accounting
// reasons, but they still appear as ISPL lots in the local DB.
async function exportPramanCSV(db, auctionId, cfg, state) {
  const rows = db.all(
    `SELECT lot_no, branch, grade, name, cr, qty, litre, bags, tel
     FROM lots WHERE auction_id = ? ${state ? 'AND state = ?' : ''}
     ORDER BY CAST(lot_no AS INTEGER), lot_no`,
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

  // Praman planter identity — derived from the live company config, not
  // hardcoded. The Praman portal expects a single planter (the registered
  // legal seller) per upload; for single-company e-Trade that's THIS
  // company, so we use the central identity resolver.
  const identity = getCompanyIdentity(cfg);
  const planterName   = identity.name;
  const planterGstin  = identity.gstin;
  const planterMobile = (cfg && (cfg.kl_phone || cfg.tn_phone || cfg.phone)) || '';
  const planterDealer = 2; // 2 = Dealer (always — registered company is the legal seller)
  // Lot company short code shown in the per-lot column. Matches the user's
  // configured short_name / logo code rather than a hardcoded 'ISPL'.
  const lotCompany = identity.shortName || 'COMPANY';

  const lines = [header.join(',')];
  for (const r of rows) {
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
      '', // Reserved Price (blank)
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

// ── Export router ────────────────────────────────────────────
const EXPORT_TYPES = {
  lot_slip:           { fn: exportLotSlip,           name: 'LotSlip' },
  lot_slip_after:     { fn: exportLotSlipAfter,      name: 'LotSlipAfter' },
  praman_csv:         { fn: exportPramanCSV,         name: 'eTrade_Praman', ext: 'csv', mime: 'text/csv', needsCfg: true },
  price_list:         { fn: exportPriceList,         name: 'PriceList' },
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
  exportLotSlip, exportLotSlipAfter, exportPramanCSV, exportPriceList,
  exportBankPayment, exportBankPaymentBefore,
  exportPoolerRegister, exportFullFile, exportCollection, exportTradeReport, exportDealerList,
  exportSalesTaxes, exportPaymentSummary, exportTDSReturn, exportTallyPurchase,
  exportSalesJournal, exportPurchaseJournal,
};
