/**
 * dbf-exports.js — DBF (FoxPro) format exports
 * 
 * Exports spice-config data in FoxPro-compatible DBF format so the legacy
 * application can continue to consume records during the transition period.
 * 
 * Field types:
 *   C = Character (text)
 *   N = Numeric
 *   D = Date
 *   L = Logical (boolean)
 * 
 * Key DBF rules learned from the previous chat:
 *   - LotNo, grade, pst_code, ppin, litre → store as TEXT (preserves leading zeros)
 *   - DATE columns are written as real DBF 'D' (date) fields via toDbfDate()
 *     (built at UTC midnight so IST doesn't shift the day). fmtDate() is still
 *     used for human-readable DD/MM/YYYY strings in XLSX header/meta lines.
 *   - Qty: 3 decimal places
 *   - Amount: 2 decimal places
 */

const { DBFFile } = require('dbffile');
const path = require('path');
const fs = require('fs');

const TMP_DIR = path.join(__dirname, 'data', 'tmp-dbf');
if (!fs.existsSync(TMP_DIR)) fs.mkdirSync(TMP_DIR, { recursive: true });

// Format date as DD/MM/YYYY string (FoxPro-friendly)
function fmtDate(d) {
  if (!d) return '';
  const s = String(d).trim();
  if (s.includes('/')) return s;
  // ISO format YYYY-MM-DD → DD/MM/YYYY
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return `${m[3]}/${m[2]}/${m[1]}`;
  return s;
}

// Parse a stored date (ISO YYYY-MM-DD or DD/MM/YYYY) into a JS Date at UTC
// midnight, for DBF 'D' (date) fields. The dbffile library formats a 'D'
// value via Date#toISOString (UTC), so we build the Date in UTC — a local
// Date would shift back a day in IST (UTC+5:30) and write the wrong date.
// Returns null for blank/unparseable input, which dbffile writes as an
// empty date field.
function toDbfDate(v) {
  if (!v) return null;
  const s = String(v).trim();
  let y, mo, d, m;
  if ((m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/)))      { y = +m[1]; mo = +m[2]; d = +m[3]; }
  else if ((m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/))) { d = +m[1]; mo = +m[2]; y = +m[3]; }
  else return null;
  if (!y || !mo || !d) return null;
  return new Date(Date.UTC(y, mo - 1, d));
}

// Safely trim string values to fit DBF field size
function fit(val, maxLen) {
  if (val === null || val === undefined) return '';
  return String(val).substring(0, maxLen);
}

// Round number to N decimals
function num(val, dec = 2) {
  const n = parseFloat(val);
  if (isNaN(n)) return 0;
  return parseFloat(n.toFixed(dec));
}

/**
 * Create a DBF file and write records to it
 * Returns a Buffer containing the DBF file contents
 */
async function writeDbfBuffer(fields, records) {
  const tmpFile = path.join(TMP_DIR, `export-${Date.now()}-${Math.random().toString(36).slice(2)}.dbf`);
  try {
    const dbf = await DBFFile.create(tmpFile, fields);
    if (records.length) await dbf.appendRecords(records);
    const buffer = fs.readFileSync(tmpFile);
    return buffer;
  } finally {
    // Cleanup temp file
    try { fs.unlinkSync(tmpFile); } catch(e) {}
  }
}

// Resolve the trade-wise filter value (ano) for transaction tables that
// key on the auction NUMBER rather than its row id. The DBF/Excel UI uses
// a single Trade dropdown whose value is the auction id, so when an
// auctionId is supplied we look the ano up here; an explicit ano wins.
function anoFilter(db, opts) {
  if (opts && opts.ano) return opts.ano;
  if (opts && opts.auctionId) {
    const a = db.get('SELECT ano FROM auctions WHERE id = ?', [opts.auctionId]);
    return a ? a.ano : null;
  }
  return null;
}

// Map DBF field defs → ExcelJS column defs so the XLSX twin of each export
// carries the same headers, order, and numeric precision as the .dbf.
function dbfFieldsToCols(fields) {
  return fields.map(f => {
    const col = {
      key: f.name,
      header: f.name,
      width: Math.min(Math.max(Math.round((f.size || 12) * 0.95), 8), 42),
    };
    if (f.type === 'N') {
      col.align = 'right';
      const dp = f.decimalPlaces || 0;
      col.numFmt = dp > 0 ? ('#,##0.' + '0'.repeat(dp)) : '#,##0';
    } else if (f.type === 'D') {
      // DATE columns carry real JS Date objects (see toDbfDate). Give the
      // XLSX twin a dd/mm/yyyy format so it shows a date, not a serial.
      col.align = 'left';
      col.numFmt = 'dd/mm/yyyy';
    } else {
      col.align = 'left';
    }
    return col;
  });
}

// Lazy require to avoid any load-order coupling with exports.js (which is a
// large module). Reuses the shared branded XLSX builder.
let _createExcelBuffer = null;
function getExcelWriter() {
  if (!_createExcelBuffer) ({ createExcelBuffer: _createExcelBuffer } = require('./exports'));
  return _createExcelBuffer;
}

// Human-readable meta lines (trade + date range) shown in the XLSX header
// band so the recipient can see exactly what the sheet was filtered to.
function metaForOpts(db, opts) {
  const lines = [];
  if (opts && opts.auctionId) {
    const a = db.get('SELECT ano, date FROM auctions WHERE id = ?', [opts.auctionId]);
    if (a) lines.push(`Trade: ${a.ano}${a.date ? ' (' + fmtDate(a.date) + ')' : ''}`);
  } else if (opts && opts.ano) {
    lines.push(`Trade: ${opts.ano}`);
  }
  if (opts && opts.from && opts.to) lines.push(`Period: ${fmtDate(opts.from)} – ${fmtDate(opts.to)}`);
  return lines;
}

// Generic writers — every module exposes a build(db, opts) → {fields,
// records}; these turn that into a .dbf or .xlsx buffer respectively.
async function exportDbf(db, type, opts = {}) {
  const def = DBF_EXPORTS[type];
  if (!def) throw new Error(`Unknown DBF export type: ${type}`);
  const { fields, records } = def.build(db, opts);
  return writeDbfBuffer(fields, records);
}
async function exportXlsx(db, type, opts = {}) {
  const def = DBF_EXPORTS[type];
  if (!def) throw new Error(`Unknown export type: ${type}`);
  const { fields, records } = def.build(db, opts);
  const createExcelBuffer = getExcelWriter();
  return createExcelBuffer(def.name, dbfFieldsToCols(fields), records, {
    db, title: def.label, metaLines: metaForOpts(db, opts),
  });
}

// ── LOTS (CPA1.DBF structure) ─────────────────────────────────
function buildLots(db, opts = {}) {
  // Trade-wise (auctionId) and/or date-wise (from/to over the auction
  // date). With neither filter, every lot across all trades is returned.
  let query = `
    SELECT l.*, a.ano as trade_no, a.date as trade_date
    FROM lots l JOIN auctions a ON a.id = l.auction_id
    WHERE 1=1`;
  const params = [];
  if (opts.auctionId) {
    // A specific trade already pins the date, so the date range is
    // redundant here — applying it too would wrongly drop the trade's
    // lots whenever the (optional) range doesn't cover the trade's date.
    query += ' AND l.auction_id = ?'; params.push(opts.auctionId);
  } else if (opts.from && opts.to) {
    query += ' AND a.date BETWEEN ? AND ?'; params.push(opts.from, opts.to);
  }
  query += ' ORDER BY a.date, CAST(l.lot_no AS INTEGER), l.lot_no';
  const rows = db.all(query, params);

  const fields = [
    { name: 'ANO',      type: 'C', size: 10 },
    { name: 'DATE',     type: 'D', size: 8 },
    { name: 'LOT',      type: 'C', size: 10 },
    { name: 'CROP',     type: 'C', size: 10 },
    { name: 'GRADE',    type: 'C', size: 10 },
    { name: 'CRPT',     type: 'C', size: 10 },
    { name: 'BR',       type: 'C', size: 30 },
    { name: 'STATE',    type: 'C', size: 20 },
    { name: 'NAME',     type: 'C', size: 50 },
    { name: 'PADD',     type: 'C', size: 80 },
    { name: 'PPLA',     type: 'C', size: 30 },
    { name: 'PPIN',     type: 'C', size: 10 },
    { name: 'PSTATE',   type: 'C', size: 20 },
    { name: 'PST_CODE', type: 'C', size: 10 },
    { name: 'CR',       type: 'C', size: 40 },
    { name: 'PAN',      type: 'C', size: 14 },
    { name: 'TEL',      type: 'C', size: 20 },
    { name: 'AADHAR',   type: 'C', size: 20 },
    { name: 'BAG',      type: 'N', size: 6, decimalPlaces: 0 },
    { name: 'LITRE',    type: 'C', size: 10 },
    { name: 'QTY',      type: 'N', size: 12, decimalPlaces: 3 },
    { name: 'PRICE',    type: 'N', size: 12, decimalPlaces: 2 },
    { name: 'AMOUNT',   type: 'N', size: 14, decimalPlaces: 2 },
    { name: 'CODE',     type: 'C', size: 10 },
    { name: 'BUYER',    type: 'C', size: 10 },
    { name: 'BUYER1',   type: 'C', size: 50 },
    { name: 'SALE',     type: 'C', size: 2 },
    { name: 'INVO',     type: 'C', size: 10 },
    { name: 'PQTY',     type: 'N', size: 12, decimalPlaces: 3 },
    { name: 'PRATE',    type: 'N', size: 12, decimalPlaces: 2 },
    { name: 'PURAMT',   type: 'N', size: 14, decimalPlaces: 2 },
    { name: 'COM',      type: 'N', size: 14, decimalPlaces: 2 },
    { name: 'CGST',     type: 'N', size: 12, decimalPlaces: 2 },
    { name: 'SGST',     type: 'N', size: 12, decimalPlaces: 2 },
    { name: 'IGST',     type: 'N', size: 12, decimalPlaces: 2 },
    { name: 'ADVANCE',  type: 'N', size: 14, decimalPlaces: 2 },
    { name: 'BALANCE',  type: 'N', size: 14, decimalPlaces: 2 },
  ];

  const records = rows.map(r => ({
    ANO: fit(r.trade_no, 10),
    DATE: toDbfDate(r.trade_date),
    LOT: fit(r.lot_no, 10),
    CROP: '',
    GRADE: fit(r.grade, 10),
    CRPT: fit(r.crpt, 10),
    BR: fit(r.branch, 30),
    STATE: fit(r.state, 20),
    NAME: fit(r.name, 50),
    PADD: fit(r.padd, 80),
    PPLA: fit(r.ppla, 30),
    PPIN: fit(r.ppin, 10),
    PSTATE: fit(r.pstate, 20),
    PST_CODE: fit(r.pst_code, 10),
    CR: fit(r.cr, 40),
    PAN: fit(r.pan, 14),
    TEL: fit(r.tel, 20),
    AADHAR: fit(r.aadhar, 20),
    BAG: parseInt(r.bags) || 0,
    LITRE: fit(r.litre, 10),
    QTY: num(r.qty, 3),
    PRICE: num(r.price, 2),
    AMOUNT: num(r.amount, 2),
    CODE: fit(r.code, 10),
    BUYER: fit(r.buyer, 10),
    BUYER1: fit(r.buyer1, 50),
    SALE: fit(r.sale, 2),
    INVO: fit(r.invo, 10),
    PQTY: num(r.pqty, 3),
    PRATE: num(r.prate, 2),
    PURAMT: num(r.puramt, 2),
    COM: num(r.com, 2),
    CGST: num(r.cgst, 2),
    SGST: num(r.sgst, 2),
    IGST: num(r.igst, 2),
    ADVANCE: num(r.advance, 2),
    BALANCE: num(r.balance, 2),
  }));

  return { fields, records };
}

// ── SALES INVOICES (INV.DBF structure) ────────────────────────
function buildInvoices(db, opts = {}) {
  let query = 'SELECT * FROM invoices WHERE 1=1';
  const params = [];
  // Trade-wise OR date-wise: a chosen trade scopes the export on its own,
  // so the date range only applies when no specific trade is selected.
  const ano = anoFilter(db, opts);
  if (ano) { query += ' AND ano = ?'; params.push(ano); }
  else if (opts.from && opts.to) { query += ' AND date BETWEEN ? AND ?'; params.push(opts.from, opts.to); }
  query += ' ORDER BY date, sale, invo';
  const rows = db.all(query, params);

  const fields = [
    { name: 'ANO',     type: 'C', size: 10 },
    { name: 'DATE',    type: 'D', size: 8 },
    { name: 'STATE',   type: 'C', size: 20 },
    { name: 'SALE',    type: 'C', size: 2 },
    { name: 'INVO',    type: 'C', size: 10 },
    { name: 'BUYER',   type: 'C', size: 10 },
    { name: 'BUYER1',  type: 'C', size: 50 },
    { name: 'GSTIN',   type: 'C', size: 20 },
    { name: 'PLACE',   type: 'C', size: 30 },
    { name: 'BAG',     type: 'N', size: 6, decimalPlaces: 0 },
    { name: 'QTY',     type: 'N', size: 12, decimalPlaces: 3 },
    { name: 'AMOUNT',  type: 'N', size: 14, decimalPlaces: 2 },
    { name: 'GUNNY',   type: 'N', size: 12, decimalPlaces: 2 },
    { name: 'PAVA_HC', type: 'N', size: 12, decimalPlaces: 2 },
    { name: 'INS',     type: 'N', size: 12, decimalPlaces: 2 },
    { name: 'CGST',    type: 'N', size: 12, decimalPlaces: 2 },
    { name: 'SGST',    type: 'N', size: 12, decimalPlaces: 2 },
    { name: 'IGST',    type: 'N', size: 12, decimalPlaces: 2 },
    { name: 'TCS',     type: 'N', size: 12, decimalPlaces: 2 },
    { name: 'RUND',    type: 'N', size: 8, decimalPlaces: 2 },
    { name: 'TOT',     type: 'N', size: 14, decimalPlaces: 2 },
  ];

  const records = rows.map(r => ({
    ANO: fit(r.ano, 10),
    DATE: toDbfDate(r.date),
    STATE: fit(r.state, 20),
    SALE: fit(r.sale, 2),
    INVO: fit(r.invo, 10),
    BUYER: fit(r.buyer, 10),
    BUYER1: fit(r.buyer1, 50),
    GSTIN: fit(r.gstin, 20),
    PLACE: fit(r.place, 30),
    BAG: parseInt(r.bag) || 0,
    QTY: num(r.qty, 3),
    AMOUNT: num(r.amount, 2),
    GUNNY: num(r.gunny, 2),
    PAVA_HC: num(r.pava_hc, 2),
    INS: num(r.ins, 2),
    CGST: num(r.cgst, 2),
    SGST: num(r.sgst, 2),
    IGST: num(r.igst, 2),
    TCS: num(r.tcs, 2),
    RUND: num(r.rund, 2),
    TOT: num(r.tot, 2),
  }));

  return { fields, records };
}

// ── PURCHASES (PURCHASE.DBF structure) ────────────────────────
function buildPurchases(db, opts = {}) {
  let query = 'SELECT * FROM purchases WHERE 1=1';
  const params = [];
  // Trade-wise OR date-wise: a chosen trade scopes the export on its own,
  // so the date range only applies when no specific trade is selected.
  const ano = anoFilter(db, opts);
  if (ano) { query += ' AND ano = ?'; params.push(ano); }
  else if (opts.from && opts.to) { query += ' AND date BETWEEN ? AND ?'; params.push(opts.from, opts.to); }
  query += ' ORDER BY date, invo';
  const rows = db.all(query, params);

  const fields = [
    { name: 'ANO',     type: 'C', size: 10 },
    { name: 'DATE',    type: 'D', size: 8 },
    { name: 'STATE',   type: 'C', size: 20 },
    { name: 'BR',      type: 'C', size: 30 },
    { name: 'NAME',    type: 'C', size: 50 },
    { name: 'ADD_LINE',type: 'C', size: 80 },
    { name: 'PLACE',   type: 'C', size: 30 },
    { name: 'GSTIN',   type: 'C', size: 40 },
    { name: 'INVO',    type: 'C', size: 10 },
    { name: 'QTY',     type: 'N', size: 12, decimalPlaces: 3 },
    { name: 'AMOUNT',  type: 'N', size: 14, decimalPlaces: 2 },
    { name: 'CGST',    type: 'N', size: 12, decimalPlaces: 2 },
    { name: 'SGST',    type: 'N', size: 12, decimalPlaces: 2 },
    { name: 'IGST',    type: 'N', size: 12, decimalPlaces: 2 },
    { name: 'RUND',    type: 'N', size: 8, decimalPlaces: 2 },
    { name: 'TOTAL',   type: 'N', size: 14, decimalPlaces: 2 },
    { name: 'TDS',     type: 'N', size: 12, decimalPlaces: 2 },
  ];

  const records = rows.map(r => ({
    ANO: fit(r.ano, 10),
    DATE: toDbfDate(r.date),
    STATE: fit(r.state, 20),
    BR: fit(r.br, 30),
    NAME: fit(r.name, 50),
    ADD_LINE: fit(r.add_line, 80),
    PLACE: fit(r.place, 30),
    GSTIN: fit(r.gstin, 40),
    INVO: fit(r.invo, 10),
    QTY: num(r.qty, 3),
    AMOUNT: num(r.amount, 2),
    CGST: num(r.cgst, 2),
    SGST: num(r.sgst, 2),
    IGST: num(r.igst, 2),
    RUND: num(r.rund, 2),
    TOTAL: num(r.total, 2),
    TDS: num(r.tds, 2),
  }));

  return { fields, records };
}

// ── BILLS of SUPPLY (BILL.DBF structure) ──────────────────────
function buildBills(db, opts = {}) {
  let query = 'SELECT * FROM bills WHERE 1=1';
  const params = [];
  // Trade-wise OR date-wise: a chosen trade scopes the export on its own,
  // so the date range only applies when no specific trade is selected.
  const ano = anoFilter(db, opts);
  if (ano) { query += ' AND ano = ?'; params.push(ano); }
  else if (opts.from && opts.to) { query += ' AND date BETWEEN ? AND ?'; params.push(opts.from, opts.to); }
  query += ' ORDER BY date, bil';
  const rows = db.all(query, params);

  const fields = [
    { name: 'ANO',     type: 'C', size: 10 },
    { name: 'DATE',    type: 'D', size: 8 },
    { name: 'STATE',   type: 'C', size: 20 },
    { name: 'BR',      type: 'C', size: 30 },
    { name: 'CRPT',    type: 'C', size: 10 },
    { name: 'BIL',     type: 'N', size: 8, decimalPlaces: 0 },
    { name: 'NAME',    type: 'C', size: 50 },
    { name: 'ADD_LINE',type: 'C', size: 80 },
    { name: 'PLA',     type: 'C', size: 30 },
    { name: 'PSTATE',  type: 'C', size: 20 },
    { name: 'ST_CODE', type: 'C', size: 10 },
    { name: 'CRR',     type: 'C', size: 20 },
    { name: 'PAN',     type: 'C', size: 14 },
    { name: 'QTY',     type: 'N', size: 12, decimalPlaces: 3 },
    { name: 'COST',    type: 'N', size: 14, decimalPlaces: 2 },
    { name: 'IGST',    type: 'N', size: 12, decimalPlaces: 2 },
    { name: 'NET',     type: 'N', size: 14, decimalPlaces: 2 },
  ];

  const records = rows.map(r => ({
    ANO: fit(r.ano, 10),
    DATE: toDbfDate(r.date),
    STATE: fit(r.state, 20),
    BR: fit(r.br, 30),
    CRPT: fit(r.crpt, 10),
    BIL: parseInt(r.bil) || 0,
    NAME: fit(r.name, 50),
    ADD_LINE: fit(r.add_line, 80),
    PLA: fit(r.pla, 30),
    PSTATE: fit(r.pstate, 20),
    ST_CODE: fit(r.st_code, 10),
    CRR: fit(r.crr, 20),
    PAN: fit(r.pan, 14),
    QTY: num(r.qty, 3),
    COST: num(r.cost, 2),
    IGST: num(r.igst, 2),
    NET: num(r.net, 2),
  }));

  return { fields, records };
}

// ── TRADERS / SELLERS (NAM.DBF structure) ─────────────────────
function buildTraders(db) {
  const rows = db.all('SELECT * FROM traders ORDER BY name');

  const fields = [
    { name: 'NAME',       type: 'C', size: 50 },
    { name: 'CR',         type: 'C', size: 40 },
    { name: 'PAN',        type: 'C', size: 14 },
    { name: 'TEL',        type: 'C', size: 20 },
    { name: 'AADHAR',     type: 'C', size: 20 },
    { name: 'PADD',       type: 'C', size: 80 },
    { name: 'PPLA',       type: 'C', size: 30 },
    { name: 'PIN',        type: 'C', size: 10 },
    { name: 'PSTATE',     type: 'C', size: 20 },
    { name: 'PST_CODE',   type: 'C', size: 10 },
    { name: 'IFSC',       type: 'C', size: 15 },
    { name: 'ACCTNUM',    type: 'C', size: 25 },
    { name: 'HOLDER_NM',  type: 'C', size: 50 },
  ];

  const records = rows.map(r => ({
    NAME: fit(r.name, 50),
    CR: fit(r.cr, 40),
    PAN: fit(r.pan, 14),
    TEL: fit(r.tel, 20),
    AADHAR: fit(r.aadhar, 20),
    PADD: fit(r.padd, 80),
    PPLA: fit(r.ppla, 30),
    PIN: fit(r.pin, 10),
    PSTATE: fit(r.pstate, 20),
    PST_CODE: fit(r.pst_code, 10),
    IFSC: fit(r.ifsc, 15),
    ACCTNUM: fit(r.acctnum, 25),
    HOLDER_NM: fit(r.holder_name, 50),
  }));

  return { fields, records };
}

// ── BUYERS / DEALERS (SBL.DBF structure) ──────────────────────
function buildBuyers(db) {
  const rows = db.all('SELECT * FROM buyers ORDER BY buyer');

  const fields = [
    { name: 'BUYER',    type: 'C', size: 10 },
    { name: 'BUYER1',   type: 'C', size: 50 },
    { name: 'ADD1',     type: 'C', size: 80 },
    { name: 'ADD2',     type: 'C', size: 80 },
    { name: 'PLA',      type: 'C', size: 30 },
    { name: 'PIN',      type: 'C', size: 10 },
    { name: 'STATE',    type: 'C', size: 20 },
    { name: 'ST_CODE',  type: 'C', size: 10 },
    { name: 'GSTIN',    type: 'C', size: 20 },
    { name: 'PAN',      type: 'C', size: 14 },
    { name: 'TEL',      type: 'C', size: 20 },
    { name: 'TI',       type: 'C', size: 20 },
    { name: 'SALE',     type: 'C', size: 2 },
  ];

  const records = rows.map(r => ({
    BUYER: fit(r.buyer, 10),
    BUYER1: fit(r.buyer1, 50),
    ADD1: fit(r.add1, 80),
    ADD2: fit(r.add2, 80),
    PLA: fit(r.pla, 30),
    PIN: fit(r.pin, 10),
    STATE: fit(r.state, 20),
    ST_CODE: fit(r.st_code, 10),
    GSTIN: fit(r.gstin, 20),
    PAN: fit(r.pan, 14),
    TEL: fit(r.tel, 20),
    TI: fit(r.ti, 20),
    SALE: fit(r.sale, 2) || 'L',
  }));

  return { fields, records };
}

// ── DEBIT NOTES ───────────────────────────────────────────────
function buildDebitNotes(db, opts = {}) {
  let query = 'SELECT * FROM debit_notes WHERE 1=1';
  const params = [];
  // Trade-wise OR date-wise: a chosen trade scopes the export on its own,
  // so the date range only applies when no specific trade is selected.
  const ano = anoFilter(db, opts);
  if (ano) { query += ' AND ano = ?'; params.push(ano); }
  else if (opts.from && opts.to) { query += ' AND date BETWEEN ? AND ?'; params.push(opts.from, opts.to); }
  query += ' ORDER BY date, note_no';
  const rows = db.all(query, params);

  const fields = [
    { name: 'ANO',     type: 'C', size: 10 },
    { name: 'DATE',    type: 'D', size: 8 },
    { name: 'STATE',   type: 'C', size: 20 },
    { name: 'NAME',    type: 'C', size: 50 },
    { name: 'NOTE_NO', type: 'C', size: 10 },
    { name: 'AMOUNT',  type: 'N', size: 14, decimalPlaces: 2 },
    { name: 'CGST',    type: 'N', size: 12, decimalPlaces: 2 },
    { name: 'SGST',    type: 'N', size: 12, decimalPlaces: 2 },
    { name: 'IGST',    type: 'N', size: 12, decimalPlaces: 2 },
    { name: 'TOTAL',   type: 'N', size: 14, decimalPlaces: 2 },
  ];

  const records = rows.map(r => ({
    ANO: fit(r.ano, 10),
    DATE: toDbfDate(r.date),
    STATE: fit(r.state, 20),
    NAME: fit(r.name, 50),
    NOTE_NO: fit(r.note_no, 10),
    AMOUNT: num(r.amount, 2),
    CGST: num(r.cgst, 2),
    SGST: num(r.sgst, 2),
    IGST: num(r.igst, 2),
    TOTAL: num(r.total, 2),
  }));

  return { fields, records };
}

// ── Registry for easy routing ─────────────────────────────────
// Each entry exposes a build(db, opts) → {fields, records}; the generic
// exportDbf / exportXlsx writers above turn that into the requested
// format. Capability flags drive the UI:
//   trade     — a Trade dropdown applies (auctionId → ano where needed)
//   dateRange — a From/To date filter applies
// Masters (traders, buyers) carry neither; they're full-table dumps.
const DBF_EXPORTS = {
  lots:         { build: buildLots,       name: 'CPA1',     trade: true,  dateRange: true,  label: 'Lots (CPA1.DBF)' },
  invoices:     { build: buildInvoices,   name: 'INV',      trade: true,  dateRange: true,  label: 'Sales Invoices (INV.DBF)' },
  purchases:    { build: buildPurchases,  name: 'PURCHASE', trade: true,  dateRange: true,  label: 'Purchases (PURCHASE.DBF)' },
  bills:        { build: buildBills,      name: 'BILL',     trade: true,  dateRange: true,  label: 'Bills of Supply (BILL.DBF)' },
  debit_notes:  { build: buildDebitNotes, name: 'DEBIT',    trade: true,  dateRange: true,  label: 'Debit Notes' },
  traders:      { build: buildTraders,    name: 'NAM',      trade: false, dateRange: false, label: 'Sellers (NAM.DBF)' },
  buyers:       { build: buildBuyers,     name: 'SBL',      trade: false, dateRange: false, label: 'Buyers (SBL.DBF)' },
};

module.exports = {
  DBF_EXPORTS,
  exportDbf,
  exportXlsx,
};
