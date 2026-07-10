/**
 * auction-reports.js — Specialized auction reports that don't fit the generic
 * column-based table renderer in exports-pdf.js. Each one mirrors a specific
 * FoxPro report layout the customer is replacing.
 *
 *   1. Lot Slip (PDF)         — pre-trade carbon-copy slip with LOT|BAG|QTY|PRICE
 *   2. Collection (XLSX+PDF)  — invoice register: INVO|TRADE NAME|NAME|QTY|VALUE
 *   3. Trade Report (XLSX+PDF)— BUYERS LIST FOR VERIFICATION grouped by state
 *
 * The XLSX path goes through a custom worksheet builder (not createExcelBuffer)
 * so we can emit headers/subtotals/group-rows that the generic builder can't.
 */

const ExcelJS = require('exceljs');
const PDFDocument = require('pdfkit');
const {
  fmtMoney, fmtQty, fmtPrice,
  getCompanyHeader, drawCompanyHeader,
  writeXlsxCompanyHeader,
} = require('./report-formatters');
// Defensive identity resolver — see _company-identity-fallback.js for
// rationale. Decoupled from the destructure above so a stale
// report-formatters.js doesn't blank `getCompanyIdentity` to undefined.
const getCompanyIdentity = require('./_company-identity-fallback').resolve();
// getSettingsFlat returns the cfg object that getCompanyIdentity expects.
// Loaded lazily inside each report so the module load order doesn't matter.
function _loadSettings(db) {
  try { return require('./company-config').getSettingsFlat(db); }
  catch(_) { return {}; }
}

// ── Number formatting ────────────────────────────────────────
// fmtMoney / fmtQty / fmtPrice come from report-formatters.js (Indian comma
// grouping; 2 decimals for rupees, 3 for kilos).
// Date format honours the user's Settings → Display → Date format
// choice via the shared ./date-format module.
const { fmtDate: fmtDateDMY } = require('./date-format');

// Manually truncate `text` so doc.widthOfString(out) <= maxWidth, appending an
// ellipsis when truncated. PDFKit 0.15's `lineBreak: false` + `ellipsis: true`
// is unreliable for long single tokens — we ellipsize ourselves so multi-word
// names like "EMPEROR SPICES PRIVATE LIMITED" never wrap into the next row.
// Caller must already have set the desired font/size on `doc` before invoking.
function fitText(doc, text, maxWidth) {
  const s = String(text == null ? '' : text);
  if (!s) return '';
  if (doc.widthOfString(s) <= maxWidth) return s;
  const ell = '…';
  const ellW = doc.widthOfString(ell);
  if (ellW >= maxWidth) return '';
  // Binary search for longest prefix that fits with the ellipsis appended.
  let lo = 0, hi = s.length;
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    if (doc.widthOfString(s.slice(0, mid)) + ellW <= maxWidth) lo = mid;
    else hi = mid - 1;
  }
  return s.slice(0, lo).trimEnd() + ell;
}

// Wrap `text` into one or more lines, each fitting within `maxWidth`. Tries
// to break on word boundaries; if a single token is wider than maxWidth, it
// breaks at character level so nothing ever overflows. Caller must have
// already set the desired font/size on `doc` before invoking. Returns an
// array of strings (one per line), at least one element.
function wrapText(doc, text, maxWidth) {
  const s = String(text == null ? '' : text).trim();
  if (!s) return [''];
  if (doc.widthOfString(s) <= maxWidth) return [s];

  const words = s.split(/\s+/);
  const lines = [];
  let cur = '';
  for (const w of words) {
    const probe = cur ? cur + ' ' + w : w;
    if (doc.widthOfString(probe) <= maxWidth) {
      cur = probe;
      continue;
    }
    // probe doesn't fit. Push current line (if any) and start fresh with w.
    if (cur) { lines.push(cur); cur = ''; }
    // If the word itself is too wide, break it at character level.
    if (doc.widthOfString(w) > maxWidth) {
      let chunk = '';
      for (const ch of w) {
        if (doc.widthOfString(chunk + ch) <= maxWidth) {
          chunk += ch;
        } else {
          if (chunk) lines.push(chunk);
          chunk = ch;
        }
      }
      cur = chunk;
    } else {
      cur = w;
    }
  }
  if (cur) lines.push(cur);
  return lines.length ? lines : [''];
}

function getAuctionHeader(db, auctionId) {
  const a = db.get('SELECT id, ano, date, crop_type, state FROM auctions WHERE id = ?', [auctionId]);
  if (!a) throw new Error('Auction not found');
  return a;
}

// ════════════════════════════════════════════════════════════
// REPORT 1 — LOT SLIP (PDF carbon-copy, pre-trade)
// ════════════════════════════════════════════════════════════
//
// Mirrors LOT_SLIP.pdf. Two identical halves per page (carbon copy), 4 cols
// per half: LOT | BAG | QTY | PRICE — PRICE is intentionally blank because
// this slip is printed BEFORE the auction so the auctioneer can fill in
// hammer prices by hand. Total bags + qty at the end.

function getLotSlipPreRows(db, auctionId, state) {
  const where = state ? 'AND state = ?' : '';
  const params = state ? [auctionId, state] : [auctionId];
  return db.all(
    `SELECT lot_no AS lot, bags AS bag, qty FROM lots
     WHERE auction_id = ? ${where}
     ORDER BY CAST(lot_no AS INTEGER), lot_no`, params
  );
}

// Generic carbon-copy two-up slip renderer. Two identical halves per A4
// page so the operator can tear the page down the gutter and hand one
// half to the seller / buyer / etc. lotSlipPdf (pre-trade) was the only
// caller originally; lot-slip-after, lot-buyer and lot-name now share
// the same layout so the office staff don't have to remember which
// report uses which format.
//
// opts = {
//   title,      — optional string shown under the brand band on each half
//   columns,    — [{ key, header, width, align?, fmt?, blank? }]
//                 width is a relative weight, normalised to half-page width.
//                 blank=true renders an empty cell (pre-trade PRICE column).
//                 fmt is an optional formatter; falls back to header-driven
//                 defaults (QTY → 3-decimal, money → 2-decimal-comma, etc.)
//   rows,       — data rows (objects keyed by columns[*].key)
//   totalKeys,  — column keys to sum and surface in the totals strip
//   totalLabel, — text shown in the first column of the totals strip
// }
async function twoUpSlipPdf(db, auctionId, opts) {
  const auction = getAuctionHeader(db, auctionId);
  const {
    title = '',
    columns,
    rows,
    totalKeys = [],
    totalLabel = 'Total',
  } = opts;
  if (!Array.isArray(columns) || !columns.length) {
    throw new Error('twoUpSlipPdf: columns is required');
  }

  const doc = new PDFDocument({ size: 'A4', layout: 'portrait', margin: 18 });
  const buffers = [];
  doc.on('data', b => buffers.push(b));

  const m = 18;
  // Wide gutter between the two halves so the page can be torn or cut down
  // the middle to give two physical copies (one for the auction office, one
  // for the seller). A dashed cut-line is drawn in the gutter as a guide.
  const gutter = 40;
  const pageW = doc.page.width;
  const pageH = doc.page.height;
  const halfW = (pageW - m * 2 - gutter) / 2;

  // Single font size used everywhere inside the table — header,
  // body, totals — so the report reads as one consistent unit.
  // Brand band / meta lines above the table keep their own sizes.
  const BASE_FONT = 8;
  const LINE_H = 9.5;
  const PAD = 6;
  const MIN_ROW_H = 14;
  const HEAD_H = 18;
  const TOP_H = 50 + (title ? 12 : 0);
  const BODY_TOP = m + TOP_H;

  // Per-column minimum width = whatever the bold header at BASE_FONT
  // actually measures, plus 6pt padding. Earlier this was a single
  // hard-coded MIN_COL=18, which left headers like AMOUNT (~28pt) and
  // PRICE (~22pt) wider than their proportional column allotments —
  // they truncated to "AMO…" / "PRI…" on the ~260pt half-page. By
  // measuring the actual header width up front, every column is
  // guaranteed to fit its label without an ellipsis at BASE_FONT.
  doc.font('Helvetica-Bold').fontSize(BASE_FONT);
  const HARD_MIN_COL = 14;   // floor for empty-header / icon-only columns
  const minColW = columns.map(c => {
    const w = doc.widthOfString(String(c.header || ''));
    return Math.max(HARD_MIN_COL, Math.ceil(w + 6));
  });
  // The first column also has to hold the totals-row label (e.g.
  // "Total"). Without this bump it would still get clipped to "Tot…"
  // even though the header itself fits.
  if (totalKeys.length && columns.length) {
    const labelW = doc.widthOfString(String(totalLabel || ''));
    minColW[0] = Math.max(minColW[0], Math.ceil(labelW + 6));
  }

  // Proportional column widths normalised to halfW. Narrow columns get
  // bumped to their per-header minimum and the wider columns donate
  // the deficit pro-rata so headers stay readable.
  const totalWeight = columns.reduce((s, c) => s + (c.width || 12), 0);
  let colW = columns.map(c => (c.width || 12) / totalWeight * halfW);
  const deficit = colW.reduce((s, w, i) => s + Math.max(0, minColW[i] - w), 0);
  if (deficit > 0) {
    const donatePool = colW.reduce((s, w, i) => s + Math.max(0, w - minColW[i]), 0);
    if (donatePool > 0) {
      colW = colW.map((w, i) => {
        if (w < minColW[i]) return minColW[i];
        const share = (w - minColW[i]) / donatePool;
        return w - deficit * share;
      });
    }
  }
  colW = colW.map((w, i) => Math.max(minColW[i], Math.floor(w)));
  // Final pixel-correction so column widths sum exactly to halfW. Apply
  // it to the widest column (not the last) so we don't squeeze code/br
  // back below its header minimum.
  const diff = halfW - colW.reduce((s, w) => s + w, 0);
  if (diff !== 0) {
    let widestIdx = 0;
    for (let i = 1; i < colW.length; i++) if (colW[i] > colW[widestIdx]) widestIdx = i;
    colW[widestIdx] = Math.max(minColW[widestIdx], colW[widestIdx] + diff);
  }
  // Reserve a little extra space at the bottom for the totals strip when
  // present — otherwise it collides with the page margin on dense pages.
  const BODY_MAX_Y = pageH - m - (totalKeys.length ? 26 : 8);

  // Resolve company branding once. The lot-slip carbon-copy halves are
  // narrow, so the brand band uses a small (22pt) logo and skips the
  // multi-line address — only the company name fits.
  const companyHeader = getCompanyHeader(db);

  // Header-driven default formatter — QTY → 3-decimal, money columns →
  // 2-decimal Indian commas, otherwise stringify. Per-column .fmt wins.
  function fmtCell(val, col) {
    if (col.blank) return '';
    if (val == null || val === '') return '';
    if (typeof col.fmt === 'function') return col.fmt(val);
    const h = String(col.header || '').toUpperCase();
    if (typeof val === 'number') {
      if (h === 'QTY' || h === 'KILOS' || h === 'QUANTITY' || h === 'PQTY' || h === 'LITRE') return fmtQty(val);
      if (Number.isInteger(val) && (h === 'BAG' || h === 'BAGS' || h === 'LOTS')) return String(val);
      if (h === 'LOT') return String(val);
      return fmtMoney(val);
    }
    return String(val);
  }
  function alignFor(col, ci) {
    if (col.align) return col.align;
    return ci === 0 ? 'center' : 'right';
  }
  // Match the renderTablePdf semantics: numeric headers render single-line
  // with font auto-shrink, everything else word-wraps. Without this, a
  // 6-digit AMOUNT or a long NAME on a ~260pt half-page would either get
  // chopped by ellipsis (the reported bug) or wrap into a meaningless
  // fragment.
  function isNumericCol(col) {
    const h = String(col.header || '').toUpperCase();
    return /^(QTY|BAG|BAGS|PRICE|RATE|AMOUNT|PQTY|PRATE|PURAMT|PURCHAMT|CGST|SGST|IGST|TCS|TOTAL|DISCOUNT|PAYABLE|ADVANCE|BALANCE|LITRE|LOTS|TDS|COST|NET|GUNNY|TRANSPORT|INSURANCE|CARDAMOM|ROUND|BILAMT|COM|KILOS)$/.test(h);
  }
  // Shrink the font for a numeric value until the rendered string fits on
  // one line — `fmtMoney(537540) → "5,37,540.00"` is wider than the half-
  // page AMOUNT column at the base size. Falls back to 5pt minimum (below
  // that the digit shapes start melting into each other).
  function fitNumericFontSize(text, maxWidth, baseSize, isBold) {
    doc.font(isBold ? 'Helvetica-Bold' : 'Helvetica');
    let size = baseSize;
    doc.fontSize(size);
    while (size > 5 && doc.widthOfString(text) > maxWidth) {
      size -= 0.5;
      doc.fontSize(size);
    }
    return size;
  }

  // Pre-measure rows: text cells wrap onto multiple lines if needed,
  // numeric cells stay single-line. Row height is driven by the tallest
  // cell in each row.
  doc.font('Helvetica').fontSize(BASE_FONT);
  const measured = rows.map(r => {
    const wrapped = columns.map((c, ci) => {
      const text = fmtCell(r[c.key], c);
      const cellW = colW[ci] - 4;
      if (isNumericCol(c)) return [String(text)];   // single-line; shrink at draw
      return wrapText(doc, String(text), cellW);
    });
    const maxLines = Math.max(1, ...wrapped.map(ls => ls.length));
    const rowH = Math.max(MIN_ROW_H, maxLines * LINE_H + PAD);
    return { wrapped, rowH };
  });

  // Group rows into pages — fit as many as the half-page can hold given
  // each row's measured height. Carbon-copy halves get the SAME slice on
  // both sides so the tear-off duplicate is identical to the original.
  const pages = [];
  let currentPage = [];
  let pageHeightUsed = 0;
  const usableYPerPage = BODY_MAX_Y - (BODY_TOP + HEAD_H);
  measured.forEach((mr, idx) => {
    if (pageHeightUsed + mr.rowH > usableYPerPage && currentPage.length) {
      pages.push(currentPage);
      currentPage = [];
      pageHeightUsed = 0;
    }
    currentPage.push({ row: rows[idx], measured: mr });
    pageHeightUsed += mr.rowH;
  });
  if (currentPage.length) pages.push(currentPage);
  if (!pages.length) pages.push([]);   // empty result still gets one page

  function drawHalfHeader(xOrigin, page) {
    // Compact company brand band (no full address — too narrow).
    const afterY = drawCompanyHeader(doc, companyHeader, {
      x: xOrigin, y: m, width: halfW,
      logoH: 22, logoW: 22, showAddress: false,
    });
    // Page / e-TRADE / Date metadata stacks beneath the brand band.
    doc.font('Helvetica').fontSize(8).fillColor('#000')
       .text(`Page: ${page}`, xOrigin, afterY, { width: halfW, align: 'right' });
    doc.font('Helvetica-Bold').fontSize(9)
       .text(`e-TRADE No:${auction.ano}`, xOrigin, afterY + 12, { width: halfW / 2, align: 'left' });
    doc.text(`Date:${fmtDateDMY(auction.date)}`, xOrigin + halfW / 2, afterY + 12, { width: halfW / 2, align: 'right' });
    if (title) {
      doc.font('Helvetica-Bold').fontSize(BASE_FONT)
         .text(title, xOrigin, afterY + 24, { width: halfW, align: 'center' });
    }

    const hy = BODY_TOP;
    doc.rect(xOrigin, hy, halfW, HEAD_H).fillAndStroke('#E8E4DD', '#444');
    // BASE_FONT for the header strip — matches body cells so the table
    // reads as one consistent unit. Column widths are guaranteed to
    // accommodate the header at this size (see minColW above), so a
    // plain text() call is enough; no fitText / ellipsis fallback.
    doc.fillColor('#000').font('Helvetica-Bold').fontSize(BASE_FONT);
    let cx = xOrigin;
    columns.forEach((c, i) => {
      doc.text(String(c.header || ''), cx + 2, hy + 5, {
        width: colW[i] - 4, align: 'center', lineBreak: false,
      });
      cx += colW[i];
    });
  }

  function drawHalfRows(xOrigin, pageEntries, isLastPage) {
    let ry = BODY_TOP + HEAD_H;
    const tblTop = BODY_TOP;

    pageEntries.forEach((entry, i) => {
      const { measured: mr } = entry;
      // Inset 0.4pt so the stripe fill doesn't cover the adjacent row
      // separator lines (see exports-pdf.js drawRow for the rationale).
      if (i % 2 === 1) doc.rect(xOrigin, ry + 0.4, halfW, mr.rowH - 0.8).fill('#F7F5F2');
      let cx = xOrigin;
      columns.forEach((c, ci) => {
        const lines = mr.wrapped[ci];
        const cellW = colW[ci] - 4;
        const align = alignFor(c, ci);
        if (isNumericCol(c) && lines.length === 1) {
          // Numeric: single-line + shrink-to-fit so totals never get
          // chopped to "…" mid-number.
          const size = fitNumericFontSize(lines[0], cellW, BASE_FONT, false);
          doc.fillColor('#000').font('Helvetica').fontSize(size);
          doc.text(lines[0], cx + 2, ry + 4, {
            width: cellW, align, lineBreak: false,
          });
        } else {
          // Text: render every wrapped line so long names don't get
          // truncated. NAME / BUYER columns are the main consumers.
          doc.fillColor('#000').font('Helvetica').fontSize(BASE_FONT);
          lines.forEach((line, li) => {
            doc.text(line, cx + 2, ry + 4 + li * LINE_H, {
              width: cellW, align, lineBreak: false,
            });
          });
        }
        cx += colW[ci];
      });
      // Horizontal row separator
      doc.moveTo(xOrigin, ry + mr.rowH).lineTo(xOrigin + halfW, ry + mr.rowH)
         .lineWidth(0.25).strokeColor('#999').stroke();
      ry += mr.rowH;
    });

    // Vertical column separators
    let vx = xOrigin;
    for (let ci = 0; ci < colW.length - 1; ci++) {
      vx += colW[ci];
      doc.moveTo(vx, tblTop).lineTo(vx, ry)
         .lineWidth(0.25).strokeColor('#888').stroke();
    }
    // Outer table border
    doc.rect(xOrigin, tblTop, halfW, ry - tblTop)
       .lineWidth(0.5).strokeColor('#444').stroke();

    // Grand total on the last page only
    if (isLastPage && totalKeys.length) {
      const tot = {};
      for (const k of totalKeys) tot[k] = rows.reduce((s, r) => s + (Number(r[k]) || 0), 0);
      const TOT_H = 16;
      const ty = ry + 2;
      doc.rect(xOrigin, ty, halfW, TOT_H + 2).fillAndStroke('#FFF3CD', '#E0B020');
      let cx = xOrigin;
      columns.forEach((c, ci) => {
        let txt;
        if (ci === 0) txt = totalLabel;
        else if (totalKeys.includes(c.key)) txt = fmtCell(tot[c.key], c);
        else txt = '';
        const cellW = colW[ci] - 4;
        const align = alignFor(c, ci);
        if (isNumericCol(c) && txt) {
          // Bold totals also auto-shrink — at the base size, the SUM of
          // AMOUNT for a busy trade can easily exceed the column width.
          const size = fitNumericFontSize(txt, cellW, BASE_FONT, true);
          doc.fillColor('#000').font('Helvetica-Bold').fontSize(size);
          doc.text(txt, cx + 2, ty + 5, { width: cellW, align, lineBreak: false });
        } else {
          doc.fillColor('#000').font('Helvetica-Bold').fontSize(BASE_FONT);
          doc.text(fitText(doc, txt, cellW), cx + 2, ty + 5, {
            width: cellW, align, lineBreak: false,
          });
        }
        cx += colW[ci];
      });
      // Verticals through the total row too
      let vx2 = xOrigin;
      for (let ci = 0; ci < colW.length - 1; ci++) {
        vx2 += colW[ci];
        doc.moveTo(vx2, ty).lineTo(vx2, ty + TOT_H + 2)
           .lineWidth(0.25).strokeColor('#E0B020').stroke();
      }
    }
  }

  pages.forEach((pageRows, i) => {
    if (i > 0) doc.addPage();
    const isLast = (i === pages.length - 1);
    drawHalfHeader(m, i + 1);
    drawHalfRows(m, pageRows, isLast);
    drawHalfHeader(m + halfW + gutter, i + 1);
    drawHalfRows(m + halfW + gutter, pageRows, isLast);

    // Vertical dashed cut-line down the middle of the gutter — a tearing
    // guide so the page can be split into two copies. A small "✂" hint at
    // the top makes the intent obvious at a glance.
    const cutX = m + halfW + gutter / 2;
    doc.save();
    doc.dash(3, { space: 3 });
    doc.moveTo(cutX, m).lineTo(cutX, pageH - m)
       .lineWidth(0.5).strokeColor('#888').stroke();
    doc.undash();
    doc.restore();
    doc.font('Helvetica').fontSize(10).fillColor('#888')
       .text('✂', cutX - 4, m - 12, { lineBreak: false });
  });

  return new Promise(resolve => {
    doc.on('end', () => resolve(Buffer.concat(buffers)));
    doc.end();
  });
}

// Pre-trade lot slip — LOT | BAG | QTY | PRICE (PRICE blank for hand-fill).
// Now a thin wrapper around twoUpSlipPdf so the layout stays in sync with
// the after-trade / by-buyer / by-name variants.
async function lotSlipPdf(db, auctionId, _cfg, extra) {
  const rows = getLotSlipPreRows(db, auctionId, extra && extra.state);
  return twoUpSlipPdf(db, auctionId, {
    columns: [
      { key: 'lot',   header: 'LOT',   width: 20, align: 'center' },
      { key: 'bag',   header: 'BAG',   width: 18, align: 'right' },
      { key: 'qty',   header: 'QTY',   width: 30, align: 'right', fmt: fmtQty },
      { key: 'price', header: 'PRICE', width: 32, align: 'right', blank: true },
    ],
    rows,
    totalKeys: ['bag', 'qty'],
  });
}

// Post-trade lot slip — adds PRICE, AMOUNT, CODE. NAME stays alongside so
// the operator can audit per-seller hammer rates without flipping reports.
// STATE is dropped from the slip-style layout (state is either filtered
// via the URL or uniform per row — adds noise without adding info on a
// narrow half-page).
async function lotSlipAfterPdf(db, auctionId, _cfg, extra) {
  const state = extra && extra.state;
  const rows = db.all(
    `SELECT lot_no AS lot, COALESCE(name,'') AS name,
            bags AS bag, qty, price, amount, COALESCE(code,'') AS code
       FROM lots
      WHERE auction_id = ? ${state ? 'AND state = ?' : ''}
      ORDER BY CAST(lot_no AS INTEGER), lot_no`,
    state ? [auctionId, state] : [auctionId]
  );
  return twoUpSlipPdf(db, auctionId, {
    title: 'LOT SLIP (AFTER TRADE)',
    columns: [
      { key: 'lot',    header: 'LOT',    width: 10, align: 'center' },
      { key: 'name',   header: 'NAME',   width: 36, align: 'left' },
      { key: 'bag',    header: 'BAG',    width: 8,  align: 'right' },
      { key: 'qty',    header: 'QTY',    width: 16, align: 'right', fmt: fmtQty },
      { key: 'price',  header: 'PRICE',  width: 12, align: 'right', fmt: fmtPrice },
      { key: 'amount', header: 'AMOUNT', width: 20, align: 'right', fmt: fmtMoney },
      { key: 'code',   header: 'CODE',   width: 10, align: 'center' },
    ],
    rows,
    totalKeys: ['bag', 'qty', 'amount'],
  });
}

// Lot Buyer — LOT | BUYER | BR | BAG | QTY. BR is the state abbreviation
// (KL / TN / etc.) so a tear-off slip is small enough to stuff into a
// trade pouch without folding.
async function lotBuyerPdf(db, auctionId) {
  const rows = db.all(
    `SELECT lot_no AS lot, COALESCE(buyer,'') AS buyer,
            CASE UPPER(COALESCE(state,''))
              WHEN 'KERALA'     THEN 'KL'
              WHEN 'TAMIL NADU' THEN 'TN'
              ELSE UPPER(SUBSTR(COALESCE(state,''), 1, 2))
            END AS br,
            bags AS bag, qty
       FROM lots
      WHERE auction_id = ?
      ORDER BY CAST(lot_no AS INTEGER), lot_no`,
    [auctionId]
  );
  return twoUpSlipPdf(db, auctionId, {
    title: 'LOT BUYER',
    columns: [
      { key: 'lot',   header: 'LOT',   width: 10, align: 'center' },
      { key: 'buyer', header: 'BUYER', width: 36, align: 'left' },
      { key: 'br',    header: 'BR',    width: 8,  align: 'center' },
      { key: 'bag',   header: 'BAG',   width: 8,  align: 'right' },
      { key: 'qty',   header: 'QTY',   width: 16, align: 'right', fmt: fmtQty },
    ],
    rows,
    totalKeys: ['bag', 'qty'],
  });
}

// Lot Name — LOT | NAME | BR | BAG | QTY | PRICE | CONTROL.
// CONTROL is intentionally blank — it's a hand-write column for the
// auctioneer's control number (matches the FoxPro report layout).
async function lotNamePdf(db, auctionId) {
  const rows = db.all(
    `SELECT lot_no AS lot, COALESCE(name,'') AS name,
            CASE UPPER(COALESCE(state,''))
              WHEN 'KERALA'     THEN 'KL'
              WHEN 'TAMIL NADU' THEN 'TN'
              ELSE UPPER(SUBSTR(COALESCE(state,''), 1, 2))
            END AS br,
            bags AS bag, qty,
            CASE WHEN COALESCE(price,0) = 0 THEN '' ELSE price END AS price
       FROM lots
      WHERE auction_id = ?
      ORDER BY CAST(lot_no AS INTEGER), lot_no`,
    [auctionId]
  );
  return twoUpSlipPdf(db, auctionId, {
    title: 'LOT NAME',
    columns: [
      { key: 'lot',     header: 'LOT',     width: 10, align: 'center' },
      { key: 'name',    header: 'NAME',    width: 36, align: 'left' },
      { key: 'br',      header: 'BR',      width: 7,  align: 'center' },
      { key: 'bag',     header: 'BAG',     width: 7,  align: 'right' },
      { key: 'qty',     header: 'QTY',     width: 14, align: 'right', fmt: fmtQty },
      { key: 'price',   header: 'PRICE',   width: 10, align: 'right', fmt: fmtPrice },
      { key: 'control', header: 'CONTROL', width: 12, align: 'center', blank: true },
    ],
    rows,
    totalKeys: ['bag', 'qty'],
  });
}

// ════════════════════════════════════════════════════════════
// REPORT 2 — COLLECTION (XLSX + PDF, invoice register)
// ════════════════════════════════════════════════════════════
//
// Mirrors COLLECTION.pdf. One row per sales invoice issued:
//   SALE+INVO | TRADE NAME (firm) | NAME (buyer) | QUANTITY | VALUE
// Grouped by buyer state — TAMIL NADU header etc. Total at the end.
//
// We pull from the `invoices` table (one row per invoice, sale='L'/'I').

function getCollectionRows(db, auctionId) {
  // Each row in this report is ONE sales invoice. Earlier the SQL used
  // a LEFT JOIN buyers ... ON OR ... which fanned out duplicate output
  // rows whenever the buyers master held multiple rows matching one
  // invoice. We then tried a correlated subquery, but sql.js (the
  // SQLite WASM build used in production) raises "no such column:
  // i.buyer" on outer-alias references inside nested expressions —
  // engine limitation, not our SQL.
  //
  // Final fix: pull two flat result sets and join in JavaScript:
  //   1. all invoices for this auction (one row each — no fanout)
  //   2. all buyers (small master table — pulled once)
  // Then for each invoice pick the single best-matching buyer
  // (preferring code-match `buyers.buyer = invoices.buyer` over
  // name-match `buyers.buyer1 = invoices.buyer1`). Deterministic via
  // buyers.id ASC tiebreak.
  const invoices = db.all(
    `SELECT id, sale, invo, buyer, buyer1, qty, tot
       FROM invoices
      WHERE auction_id = ?
      ORDER BY sale, CAST(invo AS INTEGER), invo`,
    [auctionId]
  );
  if (!invoices.length) return [];

  // Pull every buyer in the master table once. The buyers master is
  // typically <500 rows for a working set, so this is cheap.
  const buyers = db.all(`SELECT id, buyer, buyer1, sbl, state FROM buyers ORDER BY id`);

  // Index buyers by uppercase-trimmed `buyer` (code) and `buyer1`
  // (trade name) for O(1) lookup. First write wins (lowest id),
  // matching the ORDER BY id ASC tiebreak we'd want from a SQL
  // correlated subquery.
  const byCode = {}, byName = {};
  for (const b of buyers) {
    const code = String(b.buyer  || '').trim().toUpperCase();
    const name = String(b.buyer1 || '').trim().toUpperCase();
    if (code && byCode[code] == null) byCode[code] = b;
    if (name && byName[name] == null) byName[name] = b;
  }

  return invoices.map(i => {
    const iCode = String(i.buyer  || '').trim().toUpperCase();
    const iName = String(i.buyer1 || '').trim().toUpperCase();
    // Code match wins; fall back to name match for legacy invoices
    // that pre-date code stamping.
    const b = (iCode && byCode[iCode]) || (iName && byName[iName]) || null;
    const buyerName = b
      ? (b.buyer || b.sbl || i.buyer || '')
      : (i.buyer || '');
    return {
      sale:        i.sale,
      invo:        i.invo,
      trade_name:  i.buyer1 || '',
      buyer_name:  buyerName,
      qty:         i.qty,
      value:       i.tot,
      buyer_state: (b && b.state) || '',
    };
  });
}

function classifyByState(rows, auctionState) {
  // Group rows by buyer state. Auction's home state goes last (after any
  // out-of-state buyers — matches the FoxPro layout where TAMIL NADU comes
  // first only because it's the only state in this trade).
  const groups = new Map();
  const auctionSt = String(auctionState || '').trim().toUpperCase();
  for (const r of rows) {
    const st = (r.buyer_state || '').trim().toUpperCase() || auctionSt;
    if (!groups.has(st)) groups.set(st, []);
    groups.get(st).push(r);
  }
  return [...groups.entries()].sort(([a], [b]) => {
    // Auction's home state first
    if (a === auctionSt) return -1;
    if (b === auctionSt) return 1;
    return a.localeCompare(b);
  });
}

async function collectionXlsx(db, auctionId) {
  const auction = getAuctionHeader(db, auctionId);
  const rows = getCollectionRows(db, auctionId);
  const groups = classifyByState(rows, auction.state);

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Collection');
  ws.columns = [
    { width: 12 }, { width: 30 }, { width: 24 }, { width: 14 }, { width: 18 },
  ];

  // Brand band: logo (60×60) + company name + address (left), title
  // centered, trade meta right. Was a hand-rolled 2-row text merge with
  // no logo and inconsistent alignment vs other exports — replaced with
  // the same writer every other XLSX uses, so all exports look the
  // same and the Trade Name set in Settings flows through correctly.
  const companyHeader = getCompanyHeader(db);
  const headerStartRow = writeXlsxCompanyHeader(wb, ws, companyHeader, {
    colCount: 5,
    title: 'COLLECTION',
    metaLines: [
      `e-TRADE No: ${auction.ano}`,
      `Date: ${fmtDateDMY(auction.date)}`,
    ],
  });

  // Column-header row sits where the brand band reserved space.
  const head = ws.getRow(headerStartRow);
  ['INVO', 'TRADE NAME', 'NAME', 'QUANTITY', 'VALUE']
    .forEach((label, i) => { head.getCell(i + 1).value = label; });
  head.font = { bold: true };
  head.height = 20;
  head.eachCell(c => {
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8E4DD' } };
    c.border = { top: { style: 'thin' }, bottom: { style: 'thin' } };
    c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  });

  let gQty = 0, gValue = 0;
  groups.forEach(([state, items]) => {
    const sec = ws.addRow([state || 'OTHER']);
    ws.mergeCells(`A${sec.number}:E${sec.number}`);
    sec.font = { bold: true, size: 10 };
    sec.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF3CD' } };
    sec.alignment = { horizontal: 'left' };

    items.forEach(it => {
      const r = ws.addRow([
        `${it.sale || ''} ${it.invo || ''}`.trim(),
        it.trade_name || '',
        it.buyer_name || '',
        Number(it.qty) || 0,
        Number(it.value) || 0,
      ]);
      r.getCell(4).numFmt = '#,##0.000';
      r.getCell(4).alignment = { horizontal: 'right' };
      r.getCell(5).numFmt = '#,##,##0.00';
      r.getCell(5).alignment = { horizontal: 'right' };
      gQty += Number(it.qty) || 0;
      gValue += Number(it.value) || 0;
    });
  });

  const tot = ws.addRow(['', '', 'Total', gQty, gValue]);
  tot.font = { bold: true };
  tot.getCell(4).numFmt = '#,##0.000';
  tot.getCell(5).numFmt = '#,##,##0.00';
  tot.eachCell((c) => {
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF3CD' } };
    c.border = { top: { style: 'thin' }, bottom: { style: 'double' } };
  });

  return wb.xlsx.writeBuffer();
}

async function collectionPdf(db, auctionId) {
  const auction = getAuctionHeader(db, auctionId);
  const rows = getCollectionRows(db, auctionId);
  const groups = classifyByState(rows, auction.state);

  const doc = new PDFDocument({ size: 'A4', layout: 'portrait', margin: 24 });
  const buffers = [];
  doc.on('data', b => buffers.push(b));

  const m = 24;
  const pageW = doc.page.width;
  const pageH = doc.page.height;
  const usableW = pageW - m * 2;

  // Cols: INVO | TRADE NAME | NAME | QTY | VALUE
  // Portrait A4 → ~547pt usable. Long firm names like "PERFECT CARDAMOM AND
  // SPICES MARKETING COMPANY PVT LTD" wrap to multiple lines via wrapText
  // (drawDataRow grows the row height to fit the tallest cell).
  const colW = [
    Math.floor(usableW * 0.10),  // INVO
    Math.floor(usableW * 0.30),  // TRADE NAME
    Math.floor(usableW * 0.30),  // NAME
    Math.floor(usableW * 0.13),  // QUANTITY
    0,                           // VALUE absorbs rounding (~17%)
  ];
  colW[4] = usableW - colW[0] - colW[1] - colW[2] - colW[3];
  const colX = [m];
  for (let i = 0; i < colW.length - 1; i++) colX.push(colX[i] + colW[i]);

  const ROW_H = 16;
  const HEAD_H = 18;
  const SECT_H = 16;
  let y;
  let pageStartY;

  // Segment tracking — verticals only inside data segments, never through
  // full-width section header strips or the Total row strip.
  const dataSegments = [];   // per-page segments
  let curSegStart = null;

  function startSegment() { if (curSegStart === null) curSegStart = y; }
  function closeSegment() {
    if (curSegStart === null) return;
    if (y > curSegStart + 0.5) {
      const segs = dataSegments[dataSegments.length - 1];
      segs.push({ top: curSegStart, bottom: y });
    }
    curSegStart = null;
  }

  // Resolve company branding once (logo + name + address) for use across pages
  const companyHeader = getCompanyHeader(db);

  function drawTopHeader() {
    // Three-column brand band: company on left, "Collection" centered,
    // trade meta right-aligned. The page number changes per page.
    const afterY = drawCompanyHeader(doc, companyHeader, {
      x: m, y: m, width: usableW,
      title: 'COLLECTION',
      metaLines: [
        `e-TRADE No: ${auction.ano}`,
        `Date: ${fmtDateDMY(auction.date)}`,
        `Page: ${doc.bufferedPageRange().count}`,
      ],
    });
    y = afterY;
    pageStartY = y;
    dataSegments.push([]);
  }

  function drawColHeader() {
    const headTop = y;
    doc.rect(m, y, usableW, HEAD_H).fillAndStroke('#E8E4DD', '#444');
    doc.fillColor('#000').font('Helvetica-Bold').fontSize(9.5);
    const heads = ['INVO', 'TRADE NAME', 'NAME', 'QUANTITY', 'VALUE'];
    heads.forEach((h, i) => {
      const align = (i === 0) ? 'center' : (i <= 2 ? 'left' : 'right');
      doc.text(fitText(doc, h, colW[i] - 8), colX[i] + 4, y + 5, {
        width: colW[i] - 8, align, lineBreak: false,
      });
    });
    y += HEAD_H;
    curSegStart = headTop;  // header strip + following data rows = one segment
  }

  function drawSectionRow(label) {
    if (y + SECT_H > pageH - m - 12) {
      finishPage(); doc.addPage(); drawTopHeader(); drawColHeader();
    }
    closeSegment();  // section header is NOT in a data segment
    doc.rect(m, y, usableW, SECT_H).fillAndStroke('#FFF3CD', '#9A6700');
    doc.fillColor('#000').font('Helvetica-Bold').fontSize(10);
    doc.text(label, m + 6, y + 4, { width: usableW - 12, align: 'left', lineBreak: false });
    y += SECT_H;
  }

  function drawDataRow(it, idx) {
    doc.font('Helvetica').fontSize(9);
    const invoDisplay = `${it.sale || ''} ${it.invo || ''}`.trim();
    const cells = [invoDisplay, it.trade_name || '', it.buyer_name || '',
                   fmtQty(it.qty), fmtMoney(it.value)];
    const aligns = ['center', 'left', 'left', 'right', 'right'];

    const LINE_H = 12;
    const PAD_TOP = 4, PAD_BOT = 4;
    const wrapped = cells.map((v, ci) => wrapText(doc, v, colW[ci] - 8));
    const maxLines = Math.max(1, ...wrapped.map(ls => ls.length));
    const rowH = maxLines * LINE_H + PAD_TOP + PAD_BOT;

    if (y + rowH > pageH - m - 12) {
      finishPage(); doc.addPage(); drawTopHeader(); drawColHeader();
    }
    startSegment();
    // Inset 0.4pt so the stripe fill doesn't cover the row separator lines.
    if (idx % 2 === 1) doc.rect(m, y + 0.4, usableW, rowH - 0.8).fill('#F7F5F2');
    doc.fillColor('#000').font('Helvetica').fontSize(9);

    cells.forEach((v, ci) => {
      const lines = wrapped[ci];
      lines.forEach((line, li) => {
        doc.text(line, colX[ci] + 4, y + PAD_TOP + li * LINE_H, {
          width: colW[ci] - 8, align: aligns[ci], lineBreak: false,
        });
      });
    });
    doc.moveTo(m, y + rowH).lineTo(m + usableW, y + rowH)
       .lineWidth(0.3).strokeColor('#999').stroke();
    y += rowH;
  }

  // Draw verticals (only inside data segments) + outer border for current page.
  function finishPage() {
    if (pageStartY === null || pageStartY === undefined) return;
    closeSegment();
    const top = pageStartY;
    const bottom = y;
    const segs = dataSegments[dataSegments.length - 1] || [];
    for (let ci = 0; ci < colW.length - 1; ci++) {
      const vx = colX[ci] + colW[ci];
      segs.forEach(s => {
        doc.moveTo(vx, s.top).lineTo(vx, s.bottom).lineWidth(0.3).strokeColor('#888').stroke();
      });
    }
    doc.rect(m, top, usableW, bottom - top).lineWidth(0.5).strokeColor('#444').stroke();
    pageStartY = null;
  }

  drawTopHeader();
  drawColHeader();

  let gQty = 0, gValue = 0;
  let rowIdx = 0;
  groups.forEach(([state, items]) => {
    drawSectionRow(state || 'OTHER');
    items.forEach(it => {
      drawDataRow(it, rowIdx++);
      gQty += Number(it.qty) || 0;
      gValue += Number(it.value) || 0;
    });
  });

  // Total row — full-width strip, close segment first
  if (y + ROW_H + 4 > pageH - m - 12) {
    finishPage(); doc.addPage(); drawTopHeader(); drawColHeader();
  }
  closeSegment();
  doc.rect(m, y, usableW, ROW_H + 4).fillAndStroke('#FFF3CD', '#E0B020');
  doc.fillColor('#000').font('Helvetica-Bold').fontSize(10);
  // The total occupies cols: empty | empty | "Total" | qty | value
  doc.text('Total', colX[2] + 4, y + 6, { width: colW[2] - 8, align: 'right', lineBreak: false });
  doc.text(fmtQty(gQty),   colX[3] + 4, y + 6, { width: colW[3] - 8, align: 'right', lineBreak: false });
  doc.text(fmtMoney(gValue), colX[4] + 4, y + 6, { width: colW[4] - 8, align: 'right', lineBreak: false });
  y += ROW_H + 4;

  finishPage();

  return new Promise(resolve => {
    doc.on('end', () => resolve(Buffer.concat(buffers)));
    doc.end();
  });
}

// ════════════════════════════════════════════════════════════
// REPORT 3 — TRADE REPORT (XLSX + PDF, BUYERS LIST FOR VERIFICATION)
// ════════════════════════════════════════════════════════════
//
// Mirrors TRADE_REPORT.pdf. Per buyer-code in the auction:
//   SALE | BIDDER (short name) | TRADE NAME (firm) | BAG | QUANTITY | AMOUNT | INV.AMOUNT | CODE
// Grouped by state, with subtotals:
//   - INTER-STATE SALES (within state)
//   - INTRA STATE SALES (within state)
//   - <state> STATE TOTAL
//   - GRAND TOTAL
// Plus a footer block with TOTAL ARRIVALS / WITHDRAWN / SOLD / NOT eTRADED
// counts and MAX/MIN/AVERAGE prices and COST OF CARDAMOM.

function getTradeReportData(db, auctionId) {
  const auction = getAuctionHeader(db, auctionId);

  // One row per buyer code, summed across all lots.
  // INV.AMOUNT comes from the invoices table (sum of `tot` per buyer).
  // We join on buyers master to get state and full buyer info.
  //
  // Column mapping (matches user's data convention):
  //   TRADE NAME ← b.buyer1     (firm/trade name as user stores it —
  //                                "EMPEROR SPICES PRIVATE LIMITED",
  //                                "MAR TRADERS", "VARDHAN TRADING COMPANY")
  //   BUYER NAME ← b.buyer       (the buyer's short name as the user
  //                                stores it on the buyers master — the
  //                                "name" column the rest of the app
  //                                surfaces. Falls back to lots.buyer
  //                                then to the short code so the column
  //                                never goes blank for orphan rows.)
  const rows = db.all(`
    SELECT
      l.code                                                  AS code,
      COALESCE(NULLIF(b.buyer, ''),
               NULLIF(l.buyer, ''),
               NULLIF(b.code,  ''),
               '')                                            AS bidder,
      COALESCE(b.buyer1, l.buyer1, '')                        AS trade_name,
      COALESCE(b.state,  '')                                  AS state,
      COALESCE(b.gstin,  '')                                  AS gstin,
      COALESCE(l.sale,   '')                                  AS lot_sale,
      SUM(l.bags)                                             AS bag,
      SUM(l.qty)                                              AS qty,
      SUM(l.amount)                                           AS amount
    FROM lots l
    LEFT JOIN buyers b
      ON UPPER(TRIM(b.code))  = UPPER(TRIM(l.code))
      OR UPPER(TRIM(b.buyer)) = UPPER(TRIM(l.buyer))
    WHERE l.auction_id = ?
      AND l.amount > 0
    GROUP BY l.code, b.buyer, l.buyer, b.buyer1, b.code, b.state, b.gstin, l.sale
    ORDER BY UPPER(COALESCE(b.buyer1, l.buyer1, l.code)), l.code
  `, [auctionId]);

  // INV.AMOUNT is COMPUTED per buyer from their aggregated lots + the
  // configured rates — NOT read from the invoices table — so the report
  // shows the invoice value even before sales invoices are generated. The
  // math mirrors the sales-invoice calc (calculations.js buildSalesInvoice),
  // including its rounding and its transport/insurance flags:
  //   cardamomGunny = Σ amount + Σ bags × gunny_rate
  //   transport     = round2(Σ qty × transport_rate)                    (Local only *)
  //   insurance     = round2(cardamomGunny × (1 + gst%) / 1000 × insurance_rate) (Local only *)
  //   GST           = per-component round2 (CGST+SGST intra, or IGST inter)
  //   INV.AMOUNT    = round0(cardamomGunny + transport + insurance + GST) → nearest ₹
  //   * and only when the matching flag (flag_local_transport /
  //     flag_local_insurance) is ON — exactly like the generated invoice.
  // "Local" = intra-state (sale 'L'); inter-state buyers ('I') cover their own
  // freight, so transport + insurance are 0 for them.
  const cfg = _loadSettings(db);
  const _num = (v, d) => { const n = parseFloat(v); return Number.isFinite(n) ? n : d; };
  // round2 / round0 replicated from calculations.js so the components + total
  // round EXACTLY like the invoice does (string-shift rounding).
  const _round2 = (n) => { const x = Number(n); if (!isFinite(x) || x === 0) return 0; const s = x < 0 ? -1 : 1; return s * Number(Math.round(Number(Math.abs(x) + 'e2')) + 'e-2'); };
  const _round0 = (n) => { const x = Number(n); if (!isFinite(x) || x === 0) return 0; return (x < 0 ? -1 : 1) * Math.round(Math.abs(x)); };
  const _flagOn = (k) => { const v = cfg[k]; if (v === undefined || v === null || v === '') return true; return v === true || String(v).toLowerCase() === 'true'; };
  const gunnyRate     = _num(cfg.gunny_rate, 165);
  const gstGoods      = _num(cfg.gst_goods, 5);
  const halfRate      = gstGoods / 2;
  // Local transport / insurance rates — forced to 0 when their flag is off,
  // so the report drops the same components the generated invoice would.
  const transportRate = _flagOn('flag_local_transport') ? _num(cfg.local_transport, _num(cfg.transport, 2.5)) : 0;
  const insuranceRate = _flagOn('flag_local_insurance') ? _num(cfg.local_insurance, _num(cfg.insurance, 0.75)) : 0;

  const auctionState = String(auction.state || '').trim().toUpperCase();
  rows.forEach(r => {
    const buyerSt = String(r.state || '').trim().toUpperCase();
    // Inter-state if buyer state ≠ auction state. Empty buyer state defaults
    // to intra-state (matches the FoxPro fallback).
    r.sale = (buyerSt && buyerSt !== auctionState) ? 'I' : 'L';
    const isLocal = r.sale === 'L';
    const isInter = !isLocal;
    const sumAmount = Number(r.amount) || 0;
    const sumBags   = Number(r.bag) || 0;
    const sumQty    = Number(r.qty) || 0;
    const gunnyCost     = sumBags * gunnyRate;
    const cardamomGunny = sumAmount + gunnyCost;               // = invoice subtotalGoods
    const transport = isLocal ? _round2(sumQty * transportRate) : 0;
    const gstOnGoods = cardamomGunny * gstGoods / 100;
    const insurance = isLocal ? _round2((cardamomGunny + gstOnGoods) / 1000 * insuranceRate) : 0;
    const taxable = cardamomGunny + transport + insurance;
    // GST: tax each component (cardamom / gunny / transport / insurance)
    // with round2, then sum — same order the invoice uses. Intra = CGST+SGST
    // (each = round2(Σ half-rate component tax)); inter = single IGST line.
    let gst;
    if (isInter) {
      gst = _round2(
        _round2(sumAmount * gstGoods / 100) +
        _round2(gunnyCost * gstGoods / 100) +
        _round2(transport * gstGoods / 100) +
        _round2(insurance * gstGoods / 100)
      );
    } else {
      const half = _round2(
        _round2(sumAmount * halfRate / 100) +
        _round2(gunnyCost * halfRate / 100) +
        _round2(transport * halfRate / 100) +
        _round2(insurance * halfRate / 100)
      );
      gst = half * 2;   // CGST + SGST
    }
    r.inv_amount = _round0(taxable + gst);
  });

  // Group by buyer-state (column on the report). Within each state, split into
  // inter-state vs intra-state sub-buckets for the section subtotals.
  const stateGroups = new Map();
  for (const r of rows) {
    const st = (r.state || '').trim().toUpperCase() || auctionState;
    if (!stateGroups.has(st)) stateGroups.set(st, { inter: [], intra: [] });
    if (r.sale === 'I') stateGroups.get(st).inter.push(r);
    else                stateGroups.get(st).intra.push(r);
  }
  // Sort: auction's home state first
  const sortedStates = [...stateGroups.entries()].sort(([a], [b]) => {
    if (a === auctionState) return -1;
    if (b === auctionState) return 1;
    return a.localeCompare(b);
  });

  // Statistics for the footer
  const allLots = db.all(
    `SELECT bags, qty, price, amount FROM lots WHERE auction_id = ?`, [auctionId]
  );
  const sold = allLots.filter(l => Number(l.amount) > 0);
  const notSold = allLots.filter(l => !(Number(l.amount) > 0));
  const sumLots = (xs, k) => xs.reduce((s, r) => s + (Number(r[k]) || 0), 0);
  const stats = {
    arrivals_qty:  sumLots(allLots, 'qty'),
    arrivals_bags: sumLots(allLots, 'bags'),
    arrivals_lots: allLots.length,
    withdrawn_qty: 0, withdrawn_bags: 0, withdrawn_lots: 0,
    sold_qty:      sumLots(sold, 'qty'),
    sold_bags:     sumLots(sold, 'bags'),
    sold_lots:     sold.length,
    not_qty:       sumLots(notSold, 'qty'),
    not_bags:      sumLots(notSold, 'bags'),
    not_lots:      notSold.length,
    cost:          sumLots(sold, 'amount'),
    max_price:     sold.length ? Math.max(...sold.map(l => Number(l.price) || 0)) : 0,
    min_price:     sold.length ? Math.min(...sold.map(l => Number(l.price) || 0)) : 0,
    avg_price:     0,
  };
  if (stats.sold_qty > 0) stats.avg_price = stats.cost / stats.sold_qty;

  return { auction, sortedStates, stats };
}

async function tradeReportXlsx(db, auctionId) {
  const { auction, sortedStates, stats } = getTradeReportData(db, auctionId);

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('TradeReport');
  ws.columns = [
    { width: 6 },   // SALE
    { width: 22 },  // BIDDER
    { width: 26 },  // TRADE NAME
    { width: 6 },   // BAG
    { width: 12 },  // QUANTITY
    { width: 16 },  // AMOUNT
    { width: 16 },  // INV.AMOUNT
    { width: 8 },   // CODE
  ];

  // Brand band: logo + company on left, "BUYERS LIST FOR VERIFICATION"
  // centered, trade meta on right. Replaces the previous text-only
  // 3-row merge that had no logo and showed `identity.name` only when
  // `company_name` was set (now reads `trade_name` via getCompanyHeader).
  const companyHeader = getCompanyHeader(db);
  const headerStartRow = writeXlsxCompanyHeader(wb, ws, companyHeader, {
    colCount: 8,
    title: 'BUYERS LIST FOR VERIFICATION',
    metaLines: [
      `e-TRADE No: ${auction.ano}`,
      `DATE: ${fmtDateDMY(auction.date)}`,
    ],
  });

  // Column-header row sits where the brand band reserved space.
  const head = ws.getRow(headerStartRow);
  ['SALE', 'BUYER NAME', 'TRADE NAME', 'BAG', 'QUANTITY', 'AMOUNT', 'INV.AMOUNT', 'CODE']
    .forEach((label, i) => { head.getCell(i + 1).value = label; });
  head.font = { bold: true };
  head.height = 20;
  head.eachCell(c => {
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8E4DD' } };
    c.border = { top: { style: 'thin' }, bottom: { style: 'thin' } };
    c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  });

  function emitRow(r) {
    const row = ws.addRow([
      r.sale,
      r.bidder || '',
      r.trade_name || '',
      Number(r.bag) || 0,
      Number(r.qty) || 0,
      Number(r.amount) || 0,
      Number(r.inv_amount) || 0,
      r.code || '',
    ]);
    row.getCell(4).alignment = { horizontal: 'right' };
    row.getCell(5).alignment = { horizontal: 'right' };
    row.getCell(5).numFmt = '#,##0.000';
    row.getCell(6).alignment = { horizontal: 'right' };
    row.getCell(6).numFmt = '#,##,##0.00';
    row.getCell(7).alignment = { horizontal: 'right' };
    row.getCell(7).numFmt = '#,##,##0.00';
  }

  function emitSubtotal(label, items) {
    const sumKey = k => items.reduce((s, r) => s + (Number(r[k]) || 0), 0);
    const tot = ws.addRow([
      '', label, '',
      sumKey('bag'), sumKey('qty'), sumKey('amount'), '', '',
    ]);
    ws.mergeCells(`B${tot.number}:C${tot.number}`);
    tot.font = { bold: true };
    tot.getCell(5).numFmt = '#,##0.000';
    tot.getCell(6).numFmt = '#,##,##0.00';
    tot.eachCell((c) => {
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF3CD' } };
      c.border = { top: { style: 'thin' }, bottom: { style: 'thin' } };
    });
    return { bag: sumKey('bag'), qty: sumKey('qty'), amount: sumKey('amount') };
  }

  let gBag = 0, gQty = 0, gAmt = 0;
  sortedStates.forEach(([state, group]) => {
    const stateRow = ws.addRow([state]);
    ws.mergeCells(`A${stateRow.number}:H${stateRow.number}`);
    stateRow.font = { bold: true, size: 11 };
    stateRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF3CD' } };
    stateRow.alignment = { horizontal: 'left' };

    let stBag = 0, stQty = 0, stAmt = 0;
    if (group.inter.length) {
      group.inter.forEach(emitRow);
      const t = emitSubtotal('INTER-STATE SALES', group.inter);
      stBag += t.bag; stQty += t.qty; stAmt += t.amount;
    }
    if (group.intra.length) {
      group.intra.forEach(emitRow);
      const t = emitSubtotal('INTRA STATE SALES', group.intra);
      stBag += t.bag; stQty += t.qty; stAmt += t.amount;
    }
    // State total
    const stRow = ws.addRow(['', `${state} STATE TOTAL`, '', stBag, stQty, stAmt, '', '']);
    ws.mergeCells(`B${stRow.number}:C${stRow.number}`);
    stRow.font = { bold: true };
    stRow.getCell(5).numFmt = '#,##0.000';
    stRow.getCell(6).numFmt = '#,##,##0.00';
    stRow.eachCell((c) => {
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF3CD' } };
      c.border = { top: { style: 'thin' }, bottom: { style: 'thin' } };
    });
    gBag += stBag; gQty += stQty; gAmt += stAmt;
  });

  // Grand total
  const grand = ws.addRow(['', 'GRAND TOTAL', '', gBag, gQty, gAmt, '', '']);
  ws.mergeCells(`B${grand.number}:C${grand.number}`);
  grand.font = { bold: true, size: 11 };
  grand.getCell(5).numFmt = '#,##0.000';
  grand.getCell(6).numFmt = '#,##,##0.00';
  grand.eachCell((c) => {
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF3CD' } };
    c.border = { top: { style: 'double' }, bottom: { style: 'double' } };
  });

  // ── Stats footer — two clean tabular blocks ──
  // Replaces the previous single-line prose rows ("TOTAL ARRIVALS Kgs. xxx
  // Bags. xxx Lot. xxx" packed into one merged cell) with proper tables
  // so values land in their own columns and align with the numeric data
  // above. The 8-column worksheet is partitioned:
  //   Block A (cols A–D): Quantity stats  | LABEL | Kgs | Bags | Lot |
  //   Block B (cols F–H): Price stats           | METRIC | Rs. |
  // A blank gutter column (E) separates the two blocks.
  ws.addRow([]);

  // Helper to set a cell — minimal repetition.
  const setCell = (rowNum, col, value, opts = {}) => {
    const c = ws.getCell(`${col}${rowNum}`);
    c.value = value;
    if (opts.font)      c.font = opts.font;
    if (opts.fill)      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: opts.fill } };
    if (opts.align)     c.alignment = opts.align;
    if (opts.numFmt)    c.numFmt = opts.numFmt;
    if (opts.border)    c.border = opts.border;
  };
  const HDR_FILL = 'FFE8E4DD';     // same as column-header strip above
  const HDR_FONT = { bold: true, size: 10 };
  const HDR_ALIGN = { horizontal: 'center', vertical: 'middle' };
  const HDR_BORDER = { top: { style: 'thin' }, bottom: { style: 'thin' } };
  const CELL_BORDER = { top: { style: 'hair' }, bottom: { style: 'hair' },
                        left: { style: 'hair' }, right: { style: 'hair' } };
  const NUM_ALIGN  = { horizontal: 'right', vertical: 'middle' };
  const TEXT_ALIGN = { horizontal: 'left',  vertical: 'middle' };

  // Block-A header row
  const aHdr = ws.addRow([]);
  setCell(aHdr.number, 'A', '',     { font: HDR_FONT, fill: HDR_FILL, align: HDR_ALIGN, border: HDR_BORDER });
  setCell(aHdr.number, 'B', 'Kgs',  { font: HDR_FONT, fill: HDR_FILL, align: HDR_ALIGN, border: HDR_BORDER });
  setCell(aHdr.number, 'C', 'Bags', { font: HDR_FONT, fill: HDR_FILL, align: HDR_ALIGN, border: HDR_BORDER });
  setCell(aHdr.number, 'D', 'Lot',  { font: HDR_FONT, fill: HDR_FILL, align: HDR_ALIGN, border: HDR_BORDER });
  // Block-B header (same row): price metrics
  setCell(aHdr.number, 'F', 'Metric', { font: HDR_FONT, fill: HDR_FILL, align: HDR_ALIGN, border: HDR_BORDER });
  setCell(aHdr.number, 'G', 'Rs.',    { font: HDR_FONT, fill: HDR_FILL, align: HDR_ALIGN, border: HDR_BORDER });
  ws.mergeCells(`G${aHdr.number}:H${aHdr.number}`);

  // Pair the four quantity rows with the four price rows side-by-side.
  const qtyStats = [
    ['TOTAL ARRIVALS', stats.arrivals_qty,  stats.arrivals_bags,  stats.arrivals_lots],
    ['WITHDRAWN',      stats.withdrawn_qty, stats.withdrawn_bags, stats.withdrawn_lots],
    ['SOLD',           stats.sold_qty,      stats.sold_bags,      stats.sold_lots],
    ['NOT eTRADED',    stats.not_qty,       stats.not_bags,       stats.not_lots],
  ];
  const priceStats = [
    ['MAXIMUM',          stats.max_price],
    ['MINIMUM',          stats.min_price],
    ['AVERAGE',          stats.avg_price],
    ['COST OF CARDAMOM', stats.cost],
  ];
  for (let i = 0; i < 4; i++) {
    const r = ws.addRow([]);
    const [lbl, kg, bg, lt] = qtyStats[i];
    setCell(r.number, 'A', lbl, { font: { bold: true }, align: TEXT_ALIGN, border: CELL_BORDER });
    setCell(r.number, 'B', Number(kg) || 0, { align: NUM_ALIGN, numFmt: '#,##0.000', border: CELL_BORDER });
    setCell(r.number, 'C', Number(bg) || 0, { align: NUM_ALIGN, numFmt: '#,##0',     border: CELL_BORDER });
    setCell(r.number, 'D', Number(lt) || 0, { align: NUM_ALIGN, numFmt: '#,##0',     border: CELL_BORDER });

    const [pLbl, pVal] = priceStats[i];
    setCell(r.number, 'F', pLbl, { font: { bold: true }, align: TEXT_ALIGN, border: CELL_BORDER });
    setCell(r.number, 'G', Number(pVal) || 0, { align: NUM_ALIGN, numFmt: '#,##,##0.00', border: CELL_BORDER });
    ws.mergeCells(`G${r.number}:H${r.number}`);
  }

  return wb.xlsx.writeBuffer();
}

async function tradeReportPdf(db, auctionId) {
  const { auction, sortedStates, stats } = getTradeReportData(db, auctionId);

  const doc = new PDFDocument({ size: 'A4', layout: 'portrait', margin: 18 });
  const buffers = [];
  doc.on('data', b => buffers.push(b));

  const m = 18;
  const pageW = doc.page.width;
  const pageH = doc.page.height;
  const usableW = pageW - m * 2;

  // Cols: SL.NO | SALE | BIDDER | TRADE NAME | BAG | QTY | AMOUNT | INV.AMOUNT | CODE
  // Portrait A4 → ~559pt usable. Long firm/person names wrap to multiple
  // lines via wrapText (drawDataRow grows the row height to fit). Numeric
  // columns are sized to fit Indian-format money values without truncation.
  const colW = [
    Math.floor(usableW * 0.06),  // SL.NO
    Math.floor(usableW * 0.06),  // SALE
    Math.floor(usableW * 0.16),  // BIDDER
    Math.floor(usableW * 0.18),  // TRADE NAME
    Math.floor(usableW * 0.05),  // BAG
    Math.floor(usableW * 0.11),  // QTY
    Math.floor(usableW * 0.16),  // AMOUNT
    Math.floor(usableW * 0.16),  // INV.AMOUNT
    0,                           // CODE absorbs rounding
  ];
  colW[8] = usableW - colW.slice(0, 8).reduce((s, w) => s + w, 0);
  const colX = [m];
  for (let i = 0; i < colW.length - 1; i++) colX.push(colX[i] + colW[i]);

  const ROW_H = 14;
  const HEAD_H = 18;
  const SECT_H = 16;
  let y;
  let pageStartY;
  let pageNum = 0;

  // Segment tracking — verticals only inside column-header + data-row regions,
  // never through full-width section headers, subtotal strips, state-total
  // strips, or the grand-total strip.
  const dataSegments = [];   // per-page list of {top, bottom} segments
  let curSegStart = null;

  function startSegment() { if (curSegStart === null) curSegStart = y; }
  function closeSegment() {
    if (curSegStart === null) return;
    if (y > curSegStart + 0.5) {
      const segs = dataSegments[dataSegments.length - 1];
      segs.push({ top: curSegStart, bottom: y });
    }
    curSegStart = null;
  }

  // Resolve company branding once (logo + name + address) for use across pages
  const companyHeader = getCompanyHeader(db);

  function drawTopHeader() {
    pageNum += 1;
    // Three-column brand band: company on left, "BUYERS LIST FOR VERIFICATION"
    // centered, trade meta right-aligned. Page number updates per page.
    const afterY = drawCompanyHeader(doc, companyHeader, {
      x: m, y: m, width: usableW,
      title: 'BUYERS LIST FOR VERIFICATION',
      metaLines: [
        `e-TRADE No: ${auction.ano}`,
        `Date: ${fmtDateDMY(auction.date)}`,
        `Page: ${pageNum}`,
      ],
    });
    y = afterY;
    pageStartY = y;
    dataSegments.push([]);
  }

  function drawColHeader() {
    const headTop = y;
    doc.rect(m, y, usableW, HEAD_H).fillAndStroke('#E8E4DD', '#444');
    doc.fillColor('#000').font('Helvetica-Bold').fontSize(8.5);
    const heads = ['S.NO', 'SALE', 'BUYER NAME', 'TRADE NAME', 'BAG', 'QUANTITY', 'AMOUNT', 'INV.AMOUNT', 'CODE'];
    const aligns = ['center', 'center', 'left', 'left', 'center', 'right', 'right', 'right', 'center'];
    heads.forEach((h, i) => {
      doc.text(fitText(doc, h, colW[i] - 8), colX[i] + 4, y + 5, {
        width: colW[i] - 8, align: aligns[i], lineBreak: false,
      });
    });
    y += HEAD_H;
    curSegStart = headTop;
  }

  // Draws verticals + outer table border on the current page.
  function finishPage() {
    if (pageStartY === null || pageStartY === undefined) return;
    closeSegment();
    const top = pageStartY, bottom = y;
    // Draw verticals only inside data segments (header strip + lot rows)
    const segs = dataSegments[dataSegments.length - 1] || [];
    for (let ci = 0; ci < colW.length - 1; ci++) {
      const vx = colX[ci] + colW[ci];
      segs.forEach(s => {
        doc.moveTo(vx, s.top).lineTo(vx, s.bottom).lineWidth(0.3).strokeColor('#888').stroke();
      });
    }
    // Outer table border still spans the whole page table area
    doc.rect(m, top, usableW, bottom - top).lineWidth(0.5).strokeColor('#444').stroke();
    pageStartY = null;
  }

  function ensureRoom(needed) {
    if (y + needed > pageH - m - 12) {
      finishPage(); doc.addPage(); drawTopHeader(); drawColHeader();
    }
  }

  function drawStateRow(label) {
    ensureRoom(SECT_H + ROW_H);
    closeSegment();  // section header is NOT in a data segment
    doc.rect(m, y, usableW, SECT_H).fillAndStroke('#FFF3CD', '#9A6700');
    doc.fillColor('#000').font('Helvetica-Bold').fontSize(10);
    doc.text(label, m + 6, y + 4, { width: usableW - 12, align: 'left', lineBreak: false });
    y += SECT_H;
  }

  function drawDataRow(r, idx) {
    doc.font('Helvetica').fontSize(8.5);
    const cells = [
      String(idx + 1),         // SL.NO
      r.sale,
      r.bidder || '',
      r.trade_name || '',
      String(r.bag || 0),
      fmtQty(r.qty),
      fmtMoney(r.amount),
      fmtMoney(r.inv_amount),
      r.code || '',
    ];
    const aligns = ['center', 'center', 'left', 'left', 'center', 'right', 'right', 'right', 'center'];

    // Wrap each cell text into lines that fit its column width. Use the
    // tallest cell to set this row's height so nothing overflows.
    const LINE_H = 11;
    const PAD_TOP = 3, PAD_BOT = 3;
    const wrapped = cells.map((v, ci) => wrapText(doc, v, colW[ci] - 8));
    const maxLines = Math.max(1, ...wrapped.map(ls => ls.length));
    const rowH = maxLines * LINE_H + PAD_TOP + PAD_BOT;

    ensureRoom(rowH);
    startSegment();
    // Inset 0.4pt so the stripe fill doesn't cover the row separator lines.
    if (idx % 2 === 1) doc.rect(m, y + 0.4, usableW, rowH - 0.8).fill('#F7F5F2');
    doc.fillColor('#000').font('Helvetica').fontSize(8.5);

    cells.forEach((v, ci) => {
      const lines = wrapped[ci];
      lines.forEach((line, li) => {
        doc.text(line, colX[ci] + 4, y + PAD_TOP + li * LINE_H, {
          width: colW[ci] - 8, align: aligns[ci], lineBreak: false,
        });
      });
    });
    doc.moveTo(m, y + rowH).lineTo(m + usableW, y + rowH)
       .lineWidth(0.2).strokeColor('#CCC').stroke();
    y += rowH;
  }

  function drawSubtotal(label, items, color) {
    ensureRoom(ROW_H + 2);
    closeSegment();  // subtotal strip is NOT in a data segment
    doc.rect(m, y, usableW, ROW_H + 2).fillAndStroke(color || '#E6F4EA', '#5A8F62');
    doc.fillColor('#000').font('Helvetica-Bold').fontSize(9);
    const sum = k => items.reduce((s, r) => s + (Number(r[k]) || 0), 0);
    // Label spans SL.NO + SALE + BIDDER + TRADE NAME (cols 0..3)
    const labelW = colW[0] + colW[1] + colW[2] + colW[3] - 8;
    doc.text(fitText(doc, label, labelW), colX[0] + 4, y + 4, {
      width: labelW, align: 'left', lineBreak: false,
    });
    doc.text(fitText(doc, String(sum('bag')),      colW[4] - 8), colX[4] + 4, y + 4, { width: colW[4] - 8, align: 'center', lineBreak: false });
    doc.text(fitText(doc, fmtQty(sum('qty')),      colW[5] - 8), colX[5] + 4, y + 4, { width: colW[5] - 8, align: 'right',  lineBreak: false });
    doc.text(fitText(doc, fmtMoney(sum('amount')), colW[6] - 8), colX[6] + 4, y + 4, { width: colW[6] - 8, align: 'right',  lineBreak: false });
    y += ROW_H + 2;
    return { bag: sum('bag'), qty: sum('qty'), amount: sum('amount') };
  }

  drawTopHeader();
  drawColHeader();

  let gBag = 0, gQty = 0, gAmt = 0;
  sortedStates.forEach(([state, group]) => {
    drawStateRow(state);
    // Serial number (and zebra striping) restart at the beginning of each
    // state so the rows count Tamil Nadu 1..N independently of Kerala 1..M.
    let rowIdx = 0;
    let stBag = 0, stQty = 0, stAmt = 0;
    if (group.inter.length) {
      group.inter.forEach(r => drawDataRow(r, rowIdx++));
      const t = drawSubtotal('INTER-STATE SALES', group.inter, '#E6F4EA');
      stBag += t.bag; stQty += t.qty; stAmt += t.amount;
    }
    if (group.intra.length) {
      group.intra.forEach(r => drawDataRow(r, rowIdx++));
      const t = drawSubtotal('INTRA STATE SALES', group.intra, '#E6F4EA');
      stBag += t.bag; stQty += t.qty; stAmt += t.amount;
    }
    // State total (use a slightly different colour) — also full-width strip,
    // so close the data segment first.
    ensureRoom(ROW_H + 2);
    closeSegment();
    doc.rect(m, y, usableW, ROW_H + 2).fillAndStroke('#FFF3CD', '#9A6700');
    doc.fillColor('#000').font('Helvetica-Bold').fontSize(9);
    const stLabelW = colW[0] + colW[1] + colW[2] + colW[3] - 8;
    doc.text(fitText(doc, `${state} STATE TOTAL`, stLabelW), colX[0] + 4, y + 4, {
      width: stLabelW, align: 'left', lineBreak: false,
    });
    doc.text(fitText(doc, String(stBag),       colW[4] - 8), colX[4] + 4, y + 4, { width: colW[4] - 8, align: 'center', lineBreak: false });
    doc.text(fitText(doc, fmtQty(stQty),       colW[5] - 8), colX[5] + 4, y + 4, { width: colW[5] - 8, align: 'right',  lineBreak: false });
    doc.text(fitText(doc, fmtMoney(stAmt),     colW[6] - 8), colX[6] + 4, y + 4, { width: colW[6] - 8, align: 'right',  lineBreak: false });
    y += ROW_H + 2;
    gBag += stBag; gQty += stQty; gAmt += stAmt;
  });

  // Grand total — full-width strip, so close segment first
  ensureRoom(ROW_H + 4);
  closeSegment();
  doc.rect(m, y, usableW, ROW_H + 4).fillAndStroke('#FFF3CD', '#7A4400');
  doc.fillColor('#000').font('Helvetica-Bold').fontSize(10);
  const gtLabelW = colW[0] + colW[1] + colW[2] + colW[3] - 8;
  doc.text(fitText(doc, 'GRAND TOTAL', gtLabelW), colX[0] + 4, y + 5, {
    width: gtLabelW, align: 'left', lineBreak: false,
  });
  doc.text(fitText(doc, String(gBag),         colW[4] - 8), colX[4] + 4, y + 5, { width: colW[4] - 8, align: 'center', lineBreak: false });
  doc.text(fitText(doc, fmtQty(gQty),         colW[5] - 8), colX[5] + 4, y + 5, { width: colW[5] - 8, align: 'right',  lineBreak: false });
  doc.text(fitText(doc, fmtMoney(gAmt),       colW[6] - 8), colX[6] + 4, y + 5, { width: colW[6] - 8, align: 'right',  lineBreak: false });
  y += ROW_H + 4;

  finishPage();

  // ── Footer: stats summary as two clean tables ──────────────────
  // Left  table — INVENTORY: STATUS | KGS | BAGS | LOTS  (4 rows)
  // Right table — PRICE STATS: METRIC | RS              (3 rows)
  // Bottom — COST OF CARDAMOM full-width, prominent.
  const FOOT_GAP        = 16;
  const FOOT_HEAD_H     = 18;
  const FOOT_ROW_H      = 16;
  const FOOT_COST_H     = 20;
  const FOOT_BLOCK_H    = FOOT_HEAD_H + FOOT_ROW_H * 4 + FOOT_COST_H + 8;

  if (y + FOOT_GAP + FOOT_BLOCK_H > pageH - m) { doc.addPage(); y = m; }
  y += FOOT_GAP;

  // Left & right tables share the page width with a small gutter
  const gutter = 12;
  const leftW  = Math.floor((usableW - gutter) * 0.62);
  const rightW = usableW - gutter - leftW;
  const leftX  = m;
  const rightX = m + leftW + gutter;

  // Left INVENTORY table: STATUS (45%) | KGS (22%) | BAGS (16%) | LOTS (17%)
  const lcW = [
    Math.floor(leftW * 0.45),
    Math.floor(leftW * 0.22),
    Math.floor(leftW * 0.16),
    0,
  ];
  lcW[3] = leftW - lcW[0] - lcW[1] - lcW[2];
  const lcX = [leftX];
  for (let i = 0; i < lcW.length - 1; i++) lcX.push(lcX[i] + lcW[i]);

  // Header strip
  doc.rect(leftX, y, leftW, FOOT_HEAD_H).fillAndStroke('#E8E4DD', '#444');
  doc.fillColor('#000').font('Helvetica-Bold').fontSize(9);
  ['STATUS', 'KGS', 'BAGS', 'LOTS'].forEach((h, i) => {
    const align = i === 0 ? 'left' : (i === 1 ? 'right' : 'center');
    doc.text(h, lcX[i] + 6, y + 5, { width: lcW[i] - 12, align, lineBreak: false });
  });
  let ly = y + FOOT_HEAD_H;

  function leftRow(label, kgs, bags, lots, idx) {
    if (idx % 2 === 1) doc.rect(leftX, ly + 0.4, leftW, FOOT_ROW_H - 0.8).fill('#F7F5F2');
    doc.fillColor('#000').font('Helvetica').fontSize(9);
    doc.text(label, lcX[0] + 6, ly + 4, { width: lcW[0] - 12, align: 'left', lineBreak: false });
    doc.text(kgs,   lcX[1] + 6, ly + 4, { width: lcW[1] - 12, align: 'right', lineBreak: false });
    doc.text(bags,  lcX[2] + 6, ly + 4, { width: lcW[2] - 12, align: 'center', lineBreak: false });
    doc.text(lots,  lcX[3] + 6, ly + 4, { width: lcW[3] - 12, align: 'center', lineBreak: false });
    doc.moveTo(leftX, ly + FOOT_ROW_H).lineTo(leftX + leftW, ly + FOOT_ROW_H)
       .lineWidth(0.25).strokeColor('#CCC').stroke();
    ly += FOOT_ROW_H;
  }
  leftRow('TOTAL ARRIVALS',  fmtQty(stats.arrivals_qty),  String(stats.arrivals_bags), String(stats.arrivals_lots), 0);
  leftRow('WITHDRAWN',       fmtQty(stats.withdrawn_qty), String(stats.withdrawn_bags), String(stats.withdrawn_lots), 1);
  leftRow('SOLD',            fmtQty(stats.sold_qty),      String(stats.sold_bags),     String(stats.sold_lots), 2);
  leftRow('NOT eTRADED',     fmtQty(stats.not_qty),       String(stats.not_bags),      String(stats.not_lots), 3);

  // Vertical column separators inside the left table
  for (let ci = 0; ci < lcW.length - 1; ci++) {
    const vx = lcX[ci] + lcW[ci];
    doc.moveTo(vx, y).lineTo(vx, ly).lineWidth(0.3).strokeColor('#888').stroke();
  }
  doc.rect(leftX, y, leftW, ly - y).lineWidth(0.5).strokeColor('#444').stroke();

  // Right PRICE STATS table: METRIC (50%) | RS (50%)
  const rcW = [Math.floor(rightW * 0.5), 0];
  rcW[1] = rightW - rcW[0];
  const rcX = [rightX, rightX + rcW[0]];

  doc.rect(rightX, y, rightW, FOOT_HEAD_H).fillAndStroke('#E8E4DD', '#444');
  doc.fillColor('#000').font('Helvetica-Bold').fontSize(9);
  doc.text('PRICE',  rcX[0] + 6, y + 5, { width: rcW[0] - 12, align: 'left',  lineBreak: false });
  doc.text('Rs.',    rcX[1] + 6, y + 5, { width: rcW[1] - 12, align: 'right', lineBreak: false });
  let ry2 = y + FOOT_HEAD_H;

  function priceRow(label, val, idx) {
    if (idx % 2 === 1) doc.rect(rightX, ry2 + 0.4, rightW, FOOT_ROW_H - 0.8).fill('#F7F5F2');
    doc.fillColor('#000').font('Helvetica').fontSize(9);
    doc.text(label, rcX[0] + 6, ry2 + 4, { width: rcW[0] - 12, align: 'left',  lineBreak: false });
    doc.text(val,   rcX[1] + 6, ry2 + 4, { width: rcW[1] - 12, align: 'right', lineBreak: false });
    doc.moveTo(rightX, ry2 + FOOT_ROW_H).lineTo(rightX + rightW, ry2 + FOOT_ROW_H)
       .lineWidth(0.25).strokeColor('#CCC').stroke();
    ry2 += FOOT_ROW_H;
  }
  priceRow('MAXIMUM', fmtMoney(stats.max_price), 0);
  priceRow('MINIMUM', fmtMoney(stats.min_price), 1);
  priceRow('AVERAGE', fmtMoney(stats.avg_price), 2);
  // Pad with empty row so right table aligns with left's 4 rows
  if (ry2 < ly) {
    // Inset top 0.4pt so the pad fill keeps the last price row's separator.
    doc.rect(rightX, ry2 + 0.4, rightW, ly - ry2 - 0.4).fill('#F7F5F2');
    ry2 = ly;
  }

  // Verticals + outer for right table
  for (let ci = 0; ci < rcW.length - 1; ci++) {
    const vx = rcX[ci] + rcW[ci];
    doc.moveTo(vx, y).lineTo(vx, ry2).lineWidth(0.3).strokeColor('#888').stroke();
  }
  doc.rect(rightX, y, rightW, ry2 - y).lineWidth(0.5).strokeColor('#444').stroke();

  y = Math.max(ly, ry2) + 8;

  // Cost of cardamom — full-width prominent strip
  doc.rect(m, y, usableW, FOOT_COST_H).fillAndStroke('#FFF3CD', '#9A6700');
  doc.fillColor('#000').font('Helvetica-Bold').fontSize(11);
  doc.text('COST OF CARDAMOM', m + 8, y + 5, { width: usableW * 0.55, align: 'left', lineBreak: false });
  doc.text(`Rs. ${fmtMoney(stats.cost)}`, m + usableW * 0.55, y + 5, {
    width: usableW * 0.45 - 8, align: 'right', lineBreak: false,
  });
  y += FOOT_COST_H;

  return new Promise(resolve => {
    doc.on('end', () => resolve(Buffer.concat(buffers)));
    doc.end();
  });
}

module.exports = {
  // Lot slip — PDF only (XLSX stays in exports.js)
  lotSlipPdf,
  // After-trade / by-buyer / by-name variants share the carbon-copy
  // two-up layout via twoUpSlipPdf — exported so exports-pdf.js can
  // route the corresponding export types to them instead of the
  // generic single-column renderTablePdf.
  lotSlipAfterPdf,
  lotBuyerPdf,
  lotNamePdf,
  // Collection — both formats
  collectionXlsx,
  collectionPdf,
  // Trade report — both formats
  tradeReportXlsx,
  tradeReportPdf,
};
