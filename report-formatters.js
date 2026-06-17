/**
 * report-formatters.js
 *
 * Shared utilities for export PDFs / XLSX:
 *   - Indian-format number formatters (rupees, kilos, prices)
 *   - Company header (logo + name + address) drawn at the top of every PDF
 *
 * The number formatters apply Indian-style comma grouping (last 3 digits, then
 * pairs going left): 1,98,000.00 — not 198,000.00 — for amounts in rupees,
 * and 1,100.000 for kilos.
 */

const fs = require('fs');
const path = require('path');

// ── Indian-format number formatters ─────────────────────────

// Group an integer string (no sign) using Indian comma convention:
// last 3 digits keep together, every preceding pair gets a comma.
//   "1198000" -> "11,98,000"
//   "100"     -> "100"
//   "1000"    -> "1,000"
function groupIndian(digits) {
  if (digits.length <= 3) return digits;
  const last3 = digits.slice(-3);
  const rest  = digits.slice(0, -3);
  const restGrouped = rest.replace(/\B(?=(\d{2})+(?!\d))/g, ',');
  return `${restGrouped},${last3}`;
}

// Format a number with `decimals` decimal places and Indian comma grouping
// on the integer part.
function fmtIndian(n, decimals) {
  const num = Number(n) || 0;
  const fixed = num.toFixed(decimals);
  const [intPart, decPart] = fixed.split('.');
  const sign = intPart.startsWith('-') ? '-' : '';
  const digits = sign ? intPart.slice(1) : intPart;
  const grouped = groupIndian(digits);
  return decPart != null ? `${sign}${grouped}.${decPart}` : `${sign}${grouped}`;
}

// Rupees: 2 decimals, Indian commas.  Examples:
//   198000     -> "1,98,000.00"
//   1100.5     -> "1,100.50"
//   89.1       -> "89.10"
function fmtMoney(n) { return fmtIndian(n, 2); }

// Kilos: 3 decimals, Indian commas.  Examples:
//   1100       -> "1,100.000"
//   90.5       -> "90.500"
//   123456.789 -> "1,23,456.789"
function fmtQty(n) { return fmtIndian(n, 3); }

// Per-kg price (also rupees): same as money — 2 decimals, Indian commas.
function fmtPrice(n) { return fmtIndian(n, 2); }

// ── Company header (logo + name + address) ──────────────────

// Resolve the active company branding from company_settings. Falls back to
// hard-coded ISP defaults if the settings table doesn't exist or returns
// nothing — useful for export tools that run before the DB is initialized.
function getCompanyHeader(db) {
  // Defensive defaults so the header still renders if anything is missing.
  // Single-company e-Trade build: we read from `company_settings` only.
  // The legacy ISP/ASP preset switch + `s_*` sister-company fields were
  // dropped in v3, so this function reads the canonical keys directly.
  let name = '';
  let logoFile = '';
  let address1 = '';
  let branch = '';

  if (db && typeof db.get === 'function') {
    try {
      const get = (k) => {
        try {
          const r = db.get('SELECT value FROM company_settings WHERE key = ?', [k]);
          return r && r.value ? String(r.value) : '';
        } catch (_) { return ''; }
      };
      // Name resolution priority — first non-empty wins. Earlier the
      // first read was `company_name` which doesn't exist in the seed
      // data; the canonical name field in this build is `trade_name`,
      // edited under Settings → Company → Trade Name. Reading
      // `company_name` first meant the header silently fell through to
      // `tally_company_name` (or worse, the literal 'Company' default)
      // even when the user had filled in Trade Name.
      //   1. trade_name          — canonical, user-edited
      //   2. company_name        — legacy/alias (kept for backward compat)
      //   3. tally_company_name  — Tally-export name
      //   4. short_name          — short label (e.g. "VST")
      //   5. hardcoded default   — only if everything's blank
      name = get('trade_name') || get('company_name') || get('tally_company_name') || get('short_name') || '';

      // Logo file: derived from `logo` setting (the user's short logo code,
      // (whatever short code the user configured). Convention: lowercased + `.png` extension,
      // located in /public. Falls back to `logo-ispl.png` only if the
      // user's logo file is missing on disk (handled below).
      const logoCode = get('logo').toLowerCase();
      if (logoCode) logoFile = `logo-${logoCode}.png`;

      // Address: TN-prefixed in the e-Trade schema (single-state default).
      // We also accept the bare `address1` key — some installs migrated
      // values there during cleanup.
      address1 = get('tn_address1') || get('address1') || '';
      branch   = get('tn_branch')   || get('branch')   || '';
    } catch (_) {
      // Ignore — fall through to defaults
    }
  }

  // Resolve the logo to either a Buffer (uploaded BLOB from company_logos)
  // or an absolute path on disk (bundled default). Returns null if
  // missing — a fresh install with no Logo Code configured renders
  // without a logo rather than the legacy ISP image.
  // getLogoSource prefers DB BLOB so cloud-uploaded logos survive
  // redeploys without needing a persistent filesystem mount.
  const { getLogoSource } = require('./logo-paths');
  const tryPaths = [logoFile, 'logo.png'].filter(Boolean);
  const logoOnDisk = getLogoSource('ispl', tryPaths);

  // The header object exposes `address1` and `address2` for backward
  // compatibility with PDF/XLSX renderers — `address2` carries the
  // office branch (single line in the brand band).
  return { name, logoPath: logoOnDisk, address1, address2: branch, branch };
}

// ─────────────────────────────────────────────────────────────────────────
// getCompanyIdentity — central resolver for company name + identity fields.
//
// Single source of truth for every export (PDF / XLSX / XML / CSV). Reads
// from the runtime cfg object (already-flattened company_settings) and
// returns ALL the fields any downstream renderer needs, with NO hardcoded
// fallback names — empty strings if a field isn't configured. This way a
// fresh install renders blank fields rather than stale "IDEAL SPICES" or
// "AMAZING SPICE PARK" text from the legacy dual-company scaffolding.
//
// Field resolution priority (first non-empty wins):
//   name      : company_name → tally_company_name → short_name
//   shortName : short_name → logo (uppercased) → first word of name
//   address1  : tn_address1 → address1 → address
//   address2  : tn_address2 → address2 → tn_branch → branch
//   gstin     : gstin → tn_gstin → business_gstin
//   pan       : pan → tn_pan → business_pan → derived from gstin (positions 3-12)
//   state     : tn_state → business_state → state (uppercased)
//   stateCode : tally_state_code → derived from gstin first 2 chars → ''
//
// Called at the top of every export. Cheap — pure object lookups, no DB.
function getCompanyIdentity(cfg) {
  cfg = cfg || {};
  const pick = (...keys) => {
    for (const k of keys) {
      const v = cfg[k];
      if (v != null && String(v).trim() !== '') return String(v).trim();
    }
    return '';
  };
  // Name resolution — `trade_name` is the canonical user-edited field
  // in this build (Settings → Company → Trade Name). Reading
  // `company_name` first means the header silently falls through to
  // `tally_company_name` or `short_name` even when Trade Name is set.
  const name      = pick('trade_name', 'company_name', 'tally_company_name', 'short_name');
  const logoCode  = pick('logo');
  const shortName = pick('short_name') || logoCode.toUpperCase() || name.split(/\s+/)[0] || '';
  const address1  = pick('tn_address1', 'address1', 'address');
  const address2  = pick('tn_address2', 'address2', 'tn_branch', 'branch');
  const gstin     = pick('gstin', 'tn_gstin', 'business_gstin');
  // MSME / Udyam registration — single company-wide field. Printed
  // beside the company GSTIN on the Sales Invoice.
  const msme      = pick('msme');
  const pan       = pick('pan', 'tn_pan', 'business_pan')
                  || (gstin && gstin.length >= 12 ? gstin.slice(2, 12) : '');
  const state     = pick('tn_state', 'business_state', 'state').toUpperCase();
  const stateCode = pick('tally_state_code')
                  || (gstin && gstin.length >= 2 ? gstin.slice(0, 2) : '');
  // Partnership / CIN identity line.
  // Toggle in Settings → Company → "Partnership Firm". When ON, the
  // PDFs that previously printed "CIN: <value>" switch to
  // "Partnership: <partnership_name>". When OFF, the existing CIN
  // field renders as before.
  // Surface as `idLine` so PDF code doesn't need to know about the
  // toggle — just renders `${idLine.label}: ${idLine.value}` when value
  // is non-empty.
  const isPartnership = String(cfg.is_partnership || '').toLowerCase() === 'true';
  const partnershipName = pick('partnership_name');
  const cin             = pick('cin');
  const idLine = isPartnership
    ? { label: 'Partnership', value: partnershipName, isPartnership: true }
    : { label: 'CIN',         value: cin,             isPartnership: false };
  return { name, shortName, logoCode, address1, address2, gstin, msme, pan, state, stateCode, cin, idLine };
}

// Truncate `text` so doc.widthOfString(out) <= maxWidth, appending an ellipsis
// when it doesn't fit. PDFKit 0.15 doesn't reliably honor `lineBreak: false`
// when text exceeds the column — pre-fitting avoids visual overlap.
function fitText(doc, text, maxWidth) {
  if (text == null) return '';
  const s = String(text);
  if (doc.widthOfString(s) <= maxWidth) return s;
  const ELLIP = '…';
  let lo = 0, hi = s.length;
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    const candidate = s.slice(0, mid) + ELLIP;
    if (doc.widthOfString(candidate) <= maxWidth) lo = mid;
    else hi = mid - 1;
  }
  return lo > 0 ? s.slice(0, lo) + ELLIP : ELLIP;
}

// Draw a unified three-column header band on the current page:
//
//   ┌─────────────────────────────────────────────────────────────┐
//   │ ┌──────┐ COMPANY NAME              REPORT       Trade #3    │
//   │ │ LOGO │ Address line 1            TITLE        15/04/2026  │
//   │ └──────┘ Address line 2                                     │
//   └─────────────────────────────────────────────────────────────┘
//      (left block)            (middle, big)    (right, meta)
//
// The logo is sized to vertically span the company-text block so it acts as
// a proper brand mark instead of a floating ornament. Returns the y-coord
// just below the band — caller continues drawing from there.
//
// opts:
//   x, y, width — band position and width
//   title       — report title (e.g. "Bank Payment") shown in the middle
//   metaLines   — array of strings shown right-aligned (e.g. ["e-TRADE No: 3", "Date: 15/04/2026"])
//   showAddress — whether to show the address (default true; false on narrow halves)
//   logoH       — explicit logo height. Defaults to fit the company-text block.
function drawCompanyHeader(doc, header, opts) {
  const { x, y, width } = opts;
  const showAddress = opts.showAddress !== false;
  const title = opts.title || '';
  const metaLines = Array.isArray(opts.metaLines) ? opts.metaLines : [];
  const startY = y;
  const NAME_FONT = 10;
  const ADDR_FONT = 7;
  const TITLE_FONT = 16;
  const META_FONT = 9;
  const NAME_LINE_H = 13;
  const ADDR_LINE_H = 9;
  const GAP_LOGO_TEXT = 6;

  // ── Compute company-text block height so the logo can match it ──
  // 1 line of name + (0/1/2) lines of address.
  const addressLineCount = showAddress
    ? (header.address1 ? 1 : 0) + (header.address2 ? 1 : 0)
    : 0;
  const textBlockH = NAME_LINE_H + addressLineCount * ADDR_LINE_H;

  // Logo sized to match the text block height (so they sit flush together).
  // Caller can override with opts.logoH for special cases (carbon-copy halves).
  const logoH = opts.logoH || textBlockH;
  const logoW = opts.logoW || logoH;  // square by default; aspect-fit handles rectangular logos

  // ── Three-column geometry (adaptive, symmetric) ──
  // For the title to actually appear centered on the page, the left and
  // right reservations must be equal — so middle is symmetric around the
  // page center. When there's no title we don't need symmetry.
  const hasTitle = !!title;
  const hasMeta = metaLines.length > 0;
  let leftW, midW, rightW;
  if (hasTitle) {
    // Reserve equal-width side columns so the middle title column is
    // page-centered. Side width is sized to fit the company name (measured
    // at NAME_FONT) plus the logo + gap, with a percentage cap so the title
    // always has reasonable room.
    doc.font('Helvetica-Bold').fontSize(NAME_FONT);
    const nameW = doc.widthOfString(header.name || '');
    const desiredLeftContent = logoW + GAP_LOGO_TEXT + nameW + 8;
    // Clamp side width between 28% (so title gets >=44%) and 38% (so name
    // doesn't overshoot if it's very long).
    const sideW = Math.max(width * 0.28, Math.min(width * 0.38, desiredLeftContent));
    leftW = sideW;
    rightW = sideW;
    midW = width - leftW - rightW;
  } else if (hasMeta) {
    // Meta-only: small right column, left takes the rest.
    rightW = Math.min(width * 0.22, 110);
    leftW = width - rightW;
    midW = 0;
  } else {
    // No title and no meta — left block uses the full width.
    leftW = width;
    rightW = 0;
    midW = 0;
  }
  const midX = x + leftW;
  const rightX = x + width - rightW;

  // ── LEFT block: logo + company text ──
  let logoDrawn = false;
  if (header.logoPath) {
    try {
      doc.image(header.logoPath, x, startY, { fit: [logoW, logoH] });
      logoDrawn = true;
    } catch (_) { /* fall through to text-only */ }
  }
  const textX = logoDrawn ? x + logoW + GAP_LOGO_TEXT : x;
  const textW = leftW - (logoDrawn ? logoW + GAP_LOGO_TEXT : 0);

  // Company name (top of left text block). Use fitText so a long name doesn't
  // overrun and overlap into the middle column.
  doc.fillColor('#000').font('Helvetica-Bold').fontSize(NAME_FONT)
     .text(fitText(doc, header.name || '', textW - 2), textX, startY, {
       width: textW, align: 'left', lineBreak: false,
     });

  // Address lines beneath the name, in the same left column. Use fitText to
  // ellipsize lines that exceed the column width — PDFKit's `lineBreak: false`
  // is unreliable with long single-token strings.
  if (showAddress) {
    let aY = startY + NAME_LINE_H;
    doc.font('Helvetica').fontSize(ADDR_FONT).fillColor('#444');
    if (header.address1) {
      doc.text(fitText(doc, header.address1, textW - 2), textX, aY, {
        width: textW, align: 'left', lineBreak: false,
      });
      aY += ADDR_LINE_H;
    }
    if (header.address2) {
      doc.text(fitText(doc, header.address2, textW - 2), textX, aY, {
        width: textW, align: 'left', lineBreak: false,
      });
    }
  }

  // ── MIDDLE: report title, vertically centered against the brand band ──
  if (title && midW > 30) {
    // Auto-shrink the title font size until it fits in one line within the
    // middle column. PDFKit 0.15 does not reliably honor `lineBreak: false`
    // when the text exceeds the column width — measuring + scaling avoids
    // ugly wrap.
    let titleSize = TITLE_FONT;
    doc.font('Helvetica-Bold');
    while (titleSize > 9 && (doc.fontSize(titleSize), doc.widthOfString(title)) > midW - 6) {
      titleSize -= 0.5;
    }
    const titleY = startY + (Math.max(logoH, textBlockH) - titleSize) / 2;
    doc.fillColor('#000').font('Helvetica-Bold').fontSize(titleSize)
       .text(title, midX, titleY, {
         width: midW, align: 'center', lineBreak: false,
       });
  }

  // ── RIGHT: meta lines (Trade #, date, page, etc.) ──
  // Vertically center the meta block against the brand band.
  if (metaLines.length) {
    const totalMetaH = metaLines.length * (META_FONT + 4);
    let mY = startY + (Math.max(logoH, textBlockH) - totalMetaH) / 2;
    doc.fillColor('#000').font('Helvetica-Bold').fontSize(META_FONT);
    metaLines.forEach((line) => {
      doc.text(fitText(doc, line, rightW - 2), rightX, mY, {
        width: rightW, align: 'right', lineBreak: false,
      });
      mY += META_FONT + 4;
    });
  }

  // Total band height + small gap before report body begins
  const bandH = Math.max(logoH, textBlockH);
  doc.fillColor('#000');
  return startY + bandH + 6;
}

// ── XLSX helpers ────────────────────────────────────────────

// Number-format hint for ExcelJS based on a column header label.
// Matches the PDF formatters: rupees → 2 decimals + Indian commas,
// kilos → 3 decimals + Indian commas, integer counts stay plain.
function xlsxNumFmtForHeader(header) {
  const h = String(header || '').toUpperCase();
  if (h === 'QTY' || h === 'PQTY' || h === 'LITRE' || h === 'KILOS' || h === 'QUANTITY') {
    return '#,##0.000';   // Excel will format Indian-style with the right locale; pattern is the standard 3-decimal lakh format
  }
  if (h === 'BAG' || h === 'BAGS' || h === 'LOTS' || h === 'SL.NO' || h === 'S.NO') {
    return '#,##0';        // integer
  }
  // 'LOT' is intentionally NOT here — lot numbers can be alphanumeric
  // ('001A', '12B') and even purely numeric lot strings ('001') must
  // preserve their leading zeros. Forcing #,##0 would coerce '001'→1
  // and right-align it like a number. Returning null leaves it as text,
  // and the column-level alignment policy in createExcelBuffer aligns
  // text columns to the left.
  // Treat all other numerics as rupees (price/amount/total/etc.)
  if (
    h === 'PRICE'   || h === 'AMOUNT' || h === 'PURAMT' || h === 'PURCHAMT' ||
    h === 'RATE'    || h === 'PRATE'  || h === 'PAYABLE'|| h === 'DISCOUNT' ||
    h === 'BALANCE' || h === 'ADVANCE'|| h === 'VALUE' ||
    h === 'INV.AMOUNT' || h === 'TOTAL' || h === 'COMMISSION' || h === 'COM' ||
    h === 'CGST'    || h === 'SGST'   || h === 'IGST' ||
    h === 'REFUND'  || h === 'BILAMT' || h === 'BIL.AMT' ||
    // Sales-journal column headers — were previously unformatted, now
    // grouped Indian-style for readability across the wide journal.
    h === 'CARDAMOM' || h === 'GUNNY'  || h === 'TRANSPORT' ||
    h === 'INSURANCE'|| h === 'TCS'    || h === 'ROUND' ||
    // Purchase-journal extras
    h === 'COST'     || h === 'NET'    || h === 'TDS' ||
    // Sales-taxes export uses these compound headers
    h === 'CARDAMOM_COST' || h === 'CARDAMOM COST' ||
    h === 'GUNNY_COST'    || h === 'GUNNY COST'    ||
    h === 'TAX'      || h === 'ASSESS_VALUE' || h === 'ASSESS VALUE'
  ) {
    return '#,##,##0.00';  // Indian-style 2-decimal with lakh grouping
  }
  // Fallback: catch any money-like header by suffix (covers user-added
  // exports that don't show up in the explicit list above). 2 decimals
  // is the safe default for rupee columns.
  if (/(AMT|AMOUNT|COST|PRICE|RATE|TOTAL|VALUE|CHG|TAX|GST|TDS|TCS|REFUND|DISCOUNT|PAYABLE|BALANCE|ADVANCE|NET|COMMISSION|ROUND)$/.test(h)) {
    return '#,##,##0.00';
  }
  return null;             // not numeric — no format
}

// Draw a unified three-column brand band at the top of an XLSX worksheet.
// Mirrors the PDF drawCompanyHeader layout: logo + name + address (left),
// report title (middle), meta lines (right). Returns the row number of the
// first row AFTER the brand band — caller continues writing from there.
//
// `ws` is an ExcelJS worksheet. `wb` is the workbook (needed to register
// the logo image). `colCount` is how many columns the band should span
// (typically the worksheet's data column count).
//
// Layout: 4 rows tall.
//   Row 1: [logo cell merged 1×3] [company name] [report title centered] [meta line 1]
//   Row 2: [logo cell continued]  [address 1]                            [meta line 2]
//   Row 3: [logo cell continued]  [address 2]                            [meta line 3]
//   Row 4: blank spacer
//
// The brand-band rows are merged into 3 horizontal bands so name/title/meta
// each get their own column groups regardless of the underlying data column
// widths. We split the available columns: ~30% left, ~40% middle, ~30% right.
function writeXlsxCompanyHeader(wb, ws, header, opts) {
  // ── Style: matches the sample collection_6.xlsx exactly ──
  // Row 1: company name, merged A..lastCol, bold 14pt, centered
  // Row 2: meta lines joined by 4 spaces, merged, bold 11pt, centered
  // Row 3: blank spacer
  // Row 4: returned to caller for column headers
  // No logo, no 3-zone split, no separate report title — strict match.
  const colCount = Math.max(opts.colCount || 1, 1);
  const metaLines = Array.isArray(opts.metaLines) ? opts.metaLines : [];

  function colLetter(n) {
    let s = '';
    while (n > 0) {
      const r = (n - 1) % 26;
      s = String.fromCharCode(65 + r) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }
  const lastCol = colLetter(colCount);

  // Row 1 — company name. Set alignment on EVERY cell in the merge range
  // so a downstream `column.alignment` assignment can't overwrite the
  // master cell's alignment via cascade.
  ws.mergeCells(`A1:${lastCol}1`);
  const name = header && header.name ? String(header.name) : '';
  for (let i = 1; i <= colCount; i++) {
    const c = ws.getCell(`${colLetter(i)}1`);
    if (i === 1) c.value = name;
    c.font = { bold: true, size: 14 };
    c.alignment = { horizontal: 'center', vertical: 'middle' };
  }
  ws.getRow(1).height = 22;

  // Row 2 — meta lines joined by 4 spaces (matches sample
  // "e-TRADE No: 6    Date: 02/05/2026")
  ws.mergeCells(`A2:${lastCol}2`);
  const metaText = metaLines.filter(Boolean).map(String).join('    ');
  for (let i = 1; i <= colCount; i++) {
    const c = ws.getCell(`${colLetter(i)}2`);
    if (i === 1) c.value = metaText;
    c.font = { bold: true, size: 11 };
    c.alignment = { horizontal: 'center', vertical: 'middle' };
  }
  ws.getRow(2).height = 18;

  // Row 3 — spacer
  ws.getRow(3).height = 6;

  return 4;  // first row available for callers (column-header row)
}

module.exports = {
  fmtMoney, fmtQty, fmtPrice, fmtIndian,
  getCompanyHeader, getCompanyIdentity, drawCompanyHeader,
  xlsxNumFmtForHeader, writeXlsxCompanyHeader,
};
