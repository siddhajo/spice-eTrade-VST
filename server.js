// Load env vars from .env (if present) BEFORE any other require runs.
// Anything downstream — db.js, mobile-bridge, the tenant-preset admin
// gate, the Anthropic API client — can then read process.env.<KEY>
// without caring whether the value came from .env, the shell, or
// Railway's variables panel. The {silent} option means a missing
// .env file is fine (production on Railway uses dashboard vars, not
// a checked-in .env).
try { require('dotenv').config(); } catch (_) { /* dotenv not installed — fine */ }

const express = require('express');
const crypto = require('crypto');
const path = require('path');
const fs = require('fs');
const multer = require('multer');
const bcrypt = require('bcryptjs');
const rateLimit = require('express-rate-limit');
const ExcelJS = require('exceljs');
const XLSX = require('xlsx');
const { initDb, getDb, DB_PATH, replaceFromBuffer } = require('./db');
const tradeFair = require('./trade-fair');
const { initCompanySettings, CATEGORIES, getSetting, getAllSettings, updateSettings, getSettingsFlat, getGSTRates } = require('./company-config');
const { calculateLot, buildSalesInvoice, buildPurchaseInvoice, buildAgriBill, listAgriSellers, getPaymentSummary, getBankPaymentData, getTDSReturnData, getSalesJournal, getPurchaseJournal, round2, round0, distributeRoundedPayable } = require('./calculations');
const { generatePurchaseInvoicePDF, generateCropReceiptPDF, generateLotReceiptPDF, generateAgriBillPDF, generateSalesInvoicePDF, generateSalesInvoicesBatchPDF, generatePurchaseInvoicesBatchPDF, generateAgriBillsBatchPDF } = require('./invoice-pdf');
const { EXPORT_TYPES, createExcelBuffer } = require('./exports');
const { getCompanyHeader, writeXlsxCompanyHeader } = require('./report-formatters');
const { exportPdf: exportAnyPdf } = require('./exports-pdf');
const { DBF_EXPORTS, exportDbf, exportXlsx } = require('./dbf-exports');
const { REPORTS: LORRY_REPORTS } = require('./lorry-reports');
// Defensive resolution — see _company-identity-fallback.js. Uses the
// real getCompanyIdentity from report-formatters.js when available,
// falls through to an inline fallback otherwise. Fixes
// "getCompanyIdentity is not a function" on partial deploys.
const getCompanyIdentity = require('./_company-identity-fallback').resolve();
const {
  generSalesXML, generSalesIspXML, generSalesAspXML, generIspPurchaseXML,
  generRDPurchaseXML, generURDPurchaseXML, generDebitNoteXML, generLedgerXML,
  buildSalesRows, buildSalesIspRows, buildSalesAspRows,
  buildRDPurchaseRows, buildURDPurchaseRows, buildDebitNoteRows, buildLedgerRows,
  buildSalesPartyLedgerRows, buildRDPartyLedgerRows, buildURDPartyLedgerRows,
  listAuctionParties,
} = require('./tally-xml');
// Per-install time-bombed licensing — see license.js for the model.
// Token signing/verification + the license_state row helpers live there
// so server.js, the CLI minter, and any future tooling all share one
// codec.
const license = require('./license');

const app = express();
// Behind Railway's edge proxy, the real client IP arrives in the
// X-Forwarded-For header. Trust exactly ONE proxy hop so:
//   • req.ip resolves to the actual client (not the proxy), which the
//     login rate-limiter keys on; and
//   • express-rate-limit stops throwing ERR_ERL_UNEXPECTED_X_FORWARDED_FOR.
// We use 1 (not `true`): trusting all hops would let a client spoof
// X-Forwarded-For and dodge the per-IP login throttle.
app.set('trust proxy', 1);
// `verify` stashes the raw request bytes so the WhatsApp webhook can
// validate Meta's X-Hub-Signature-256 (HMAC-SHA256 of the raw body with
// the app secret) — JSON.stringify(req.body) is NOT byte-identical to
// what Meta signed, so the parsed body can't be used for verification.
app.use(express.json({ limit: '50mb', verify: (req, _res, buf) => { req.rawBody = buf; } }));

// Disable caching of HTML files so users always get the latest UI without
// needing a hard-reload. This is critical for ngrok-tunnelled deployments
// where intermediate proxies may cache aggressively. JavaScript/CSS/
// images can still be cached normally (handled by the static middleware
// after this).
app.use((req, res, next) => {
  if (req.path === '/' || req.path.endsWith('.html')) {
    res.set('Cache-Control', 'no-store, no-cache, must-revalidate, private');
    res.set('Pragma', 'no-cache');
    res.set('Expires', '0');
  }
  next();
});

// Logo resolution helper — uploads go to a persistent volume
// (SPICE_DATA_DIR/logos) so they survive Railway/Heroku redeploys,
// which wipe everything outside the mounted volume. Without this the
// uploaded logo would vanish on every container restart and fall back
// to the bundled default (logo_kj.png). See logo-paths.js.
const { resolveLogoPath: _resolveLogoPath, getLogoSource: _getLogoSource } = require('./logo-paths');

// Uploaded logos live in the company_logos BLOB table so they persist
// across Railway/Heroku redeploys. These routes try the DB first, then
// fall back to the bundled /public file (logo_kj.png is the ultimate
// default when no custom upload exists). GET /api/company-settings/logo/:which
// still returns exists:false until the user uploads, so the "Upload your
// logo" UI state stays accurate.
function _serveLogoBlob(req, res, which, nextOrFallback) {
  try {
    const row = getDb().get('SELECT mime, data FROM company_logos WHERE key = ?', [which]);
    if (row && row.data && row.data.length) {
      res.setHeader('Content-Type', row.mime || 'image/png');
      res.setHeader('Cache-Control', 'no-cache');
      return res.end(Buffer.from(row.data));
    }
  } catch (_) { /* DB not ready yet — fall through to filesystem default */ }
  nextOrFallback();
}
app.get('/logo-ispl.png', (req, res) => {
  _serveLogoBlob(req, res, 'ispl', () => {
    const bundled = _resolveLogoPath('logo-ispl.png') || _resolveLogoPath('logo_kj.png');
    if (bundled) return res.sendFile(bundled);
    res.status(404).end();
  });
});
app.get('/logo-asp.png', (req, res, next) => {
  _serveLogoBlob(req, res, 'asp', () => next());
});

app.use(express.static(path.join(__dirname, 'public')));

// Prevent browser/proxy caching of API responses so Refresh buttons actually
// fetch fresh data (without this, fetch() may return stale cached JSON)
app.use('/api', (req, res, next) => {
  res.set('Cache-Control', 'no-store, no-cache, must-revalidate, private');
  res.set('Pragma', 'no-cache');
  res.set('Expires', '0');
  next();
});

// Health check — used by the Electron wrapper to wait until the server
// is ready to accept requests before loading the window URL. Returns a
// minimal 200 with no auth required.
app.get('/api/health', (req, res) => {
  res.json({ ok: true, ts: Date.now() });
});

// File upload setup
// Honor SPICE_DATA_DIR so uploads also land in userData when packaged.
const uploadDir = path.join(process.env.SPICE_DATA_DIR || path.join(__dirname, 'data'), 'uploads');
if (!fs.existsSync(uploadDir)) fs.mkdirSync(uploadDir, { recursive: true });
const upload = multer({ dest: uploadDir, limits: { fileSize: 20 * 1024 * 1024 } });

// ── Password hashing ──────────────────────────────────────────────
// Bcrypt cost factor. 12 ≈ 250ms per hash on modern hardware — slow enough
// to make brute-force expensive but fast enough that login latency is fine.
const BCRYPT_ROUNDS = 12;

// Pre-computed dummy bcrypt hash for unknown-user logins. Running a real
// bcrypt comparison against this when the username isn't found prevents
// timing-based user enumeration (otherwise unknown-user paths return in
// microseconds while real-user paths take ~250ms).
const DUMMY_BCRYPT_HASH = bcrypt.hashSync('__never_a_real_password__', BCRYPT_ROUNDS);

// Hash a plaintext password with bcrypt. Async — every callsite is in a
// route handler that can be made async.
async function hashPassword(plain) {
  return bcrypt.hash(String(plain), BCRYPT_ROUNDS);
}

// Detect legacy SHA-256 password hashes (64 lowercase hex chars). Bcrypt
// hashes start with "$2a$" / "$2b$" / "$2y$" so the two are unambiguous.
const LEGACY_SHA256_RE = /^[a-f0-9]{64}$/i;
function isLegacyHash(stored) { return LEGACY_SHA256_RE.test(stored || ''); }

// Verify a plaintext password against a stored hash, supporting both new
// bcrypt rows and legacy SHA-256 rows from before the migration. Returns
// true/false; never throws on bad input.
async function verifyPassword(plain, stored) {
  if (!stored || typeof stored !== 'string') return false;
  if (isLegacyHash(stored)) {
    const sha = crypto.createHash('sha256').update(String(plain)).digest('hex');
    return sha === stored;
  }
  try { return await bcrypt.compare(String(plain), stored); }
  catch { return false; }
}

// ── Date helpers ──
// Convert any date-ish input (Date object, Excel serial number, dd/mm/yyyy,
// yyyy-mm-dd, etc.) to canonical ISO yyyy-mm-dd for storage.
function normalizeDate(v) {
  if (v === null || v === undefined || v === '') return '';
  // Date object
  if (v instanceof Date && !isNaN(v)) {
    return v.toISOString().slice(0, 10);
  }
  // Number = Excel serial (days since 1900-01-01, with the famous 1900 leap-year bug)
  if (typeof v === 'number' && v > 0 && v < 80000) {
    // Excel epoch: 1899-12-30 (accounts for the 1900 leap-year bug)
    const ms = (v - 25569) * 86400 * 1000;
    const d = new Date(ms);
    if (!isNaN(d)) return d.toISOString().slice(0, 10);
  }
  const s = String(v).trim();
  // ISO yyyy-mm-dd (or yyyy-mm-dd HH:MM:SS)
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.substring(0, 10);
  // dd/mm/yyyy or dd-mm-yyyy
  const m1 = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  if (m1) return `${m1[3]}-${m1[2].padStart(2, '0')}-${m1[1].padStart(2, '0')}`;
  // Pure numeric string Excel serial
  if (/^\d+(\.\d+)?$/.test(s)) {
    const n = Number(s);
    if (n > 0 && n < 80000) {
      const ms = (n - 25569) * 86400 * 1000;
      const d = new Date(ms);
      if (!isNaN(d)) return d.toISOString().slice(0, 10);
    }
  }
  // Last resort: try Date parsing
  const d = new Date(s);
  if (!isNaN(d)) return d.toISOString().slice(0, 10);
  return s;
}

// Add N days to an ISO date string (YYYY-MM-DD). String-based to avoid
// timezone drift — operating on a plain Date() in JS will silently shift
// the day across DST boundaries or when the server is in a non-UTC TZ.
// Uses Date in UTC mode purely for the calendar math (month/year rollover,
// leap years), then re-extracts the YYYY-MM-DD components.
//
// Edge cases handled:
//   - Month rollover:  '2026-01-31' + 1  → '2026-02-01'
//   - Year rollover:   '2026-12-31' + 1  → '2027-01-01'
//   - Leap year:       '2028-02-28' + 1  → '2028-02-29'
//   - Non-leap year:   '2026-02-28' + 1  → '2026-03-01'
//   - Negative offset: '2026-03-01' + -1 → '2026-02-28'
//   - Empty input:                       → '' (no throw)
function addDays(isoDate, days) {
  const iso = normalizeDate(isoDate);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(iso)) return '';
  const [y, m, d] = iso.split('-').map(Number);
  // Date.UTC anchors the math at midnight UTC so DST never enters the
  // calculation — month rollover, year rollover, and leap years all
  // resolve correctly via the underlying calendar arithmetic.
  const ms = Date.UTC(y, m - 1, d) + (Number(days) || 0) * 86400 * 1000;
  const out = new Date(ms);
  const yy = out.getUTCFullYear();
  const mm = String(out.getUTCMonth() + 1).padStart(2, '0');
  const dd = String(out.getUTCDate()).padStart(2, '0');
  return `${yy}-${mm}-${dd}`;
}

// Shared date-format helpers live in ./date-format.js so PDF / XLSX
// generators in other modules can import the same single source of
// truth (avoids drift between local copies of fmtDate that all
// hardcode DD/MM/YYYY).
const { fmtDate, todayLocalISO, invalidateDateFormatCache } = require('./date-format');

function withFmtDate(rows, field = 'date') {
  if (!Array.isArray(rows)) return rows;
  return rows.map(r => ({ ...r, date_fmt: fmtDate(r[field]) }));
}

// Auth middleware: verify a valid session, attach req.user/req.session.
// DOES NOT check role — use this for endpoints that any logged-in user
// (admin OR regular user) should be able to hit (GET list endpoints mostly).
// Endpoints reachable while a user is in the forced-password-change state.
// Anything outside this set returns 403 until they rotate their password.
// Endpoints reachable while a user is in the forced-password-change
// state. Includes both desktop (/api/me*) and mobile (/api/auth/*)
// paths — without /api/auth/change-password here, mobile users with
// must_change_password=1 get a 403 on every endpoint INCLUDING the
// one they need to clear the flag, leaving them permanently stuck.
const FORCED_CHANGE_ALLOWED = new Set([
  '/api/me', '/api/me/password',
  '/api/auth/me', '/api/auth/change-password', '/api/auth/logout',
]);

function requireAuth(req, res, next) {
  const token = (req.headers.authorization || '').replace('Bearer ', '');
  if (!token) return res.status(401).json({ error: 'No token' });
  const db = getDb();
  // Strict expiry: rows pre-dating the migration get a NULL expires_at and
  // are grandfathered in. New rows have a hard cap — once expires_at is in
  // the past, the token is rejected even if last_used_at is recent. Compare
  // in SQLite (datetime('now','localtime')) so it matches the same TZ
  // semantics the INSERT used; mixing JS UTC with SQLite localtime would
  // wrongly expire sessions on any non-UTC host.
  const session = db.get(
    `SELECT * FROM sessions
     WHERE token = ?
       AND (expires_at IS NULL OR expires_at >= datetime('now','localtime'))`,
    [token]
  );
  if (!session) {
    // Best-effort cleanup: drop the row if it existed but expired.
    db.run('DELETE FROM sessions WHERE token = ?', [token]);
    return res.status(403).json({ error: 'Session expired — please sign in again' });
  }
  const user = db.get('SELECT * FROM users WHERE id = ?', [session.user_id]);
  if (!user) return res.status(403).json({ error: 'Unauthorized' });
  // Block every other endpoint until the seeded/reset password is rotated —
  // closes the default-creds attack window even if rate-limiting is missed.
  if (user.must_change_password && !FORCED_CHANGE_ALLOWED.has(req.path)) {
    return res.status(403).json({
      error: 'Password change required before continuing',
      must_change_password: true
    });
  }
  // Touch last_used_at for cleanup / activity display
  db.run(`UPDATE sessions SET last_used_at = datetime('now','localtime') WHERE token = ?`, [token]);
  req.user = user;
  req.session = session;
  next();
}

// Admin-only middleware: gates mutations, settings, deletes, user management.
// Runs requireAuth first, then verifies role.
function requireAdmin(req, res, next) {
  requireAuth(req, res, (err) => {
    if (err) return next(err);
    if (!req.user || req.user.role !== 'admin') {
      return res.status(403).json({ error: 'Admin access required for this action' });
    }
    next();
  });
}

// ══════════════════════════════════════════════════════════════
// ROLE-BASED PERMISSIONS
// ══════════════════════════════════════════════════════════════
// Four pre-defined role tiers, each granting a fixed set of capabilities.
// Capabilities are referenced by name everywhere in the code that needs
// permission gating, so adding/removing a capability touches one place.
//
// Hierarchy (least → most privileged):
//   viewer    — read-only
//   operator  — daily auction-floor work (lots, invoices, buyers/traders)
//   manager   — branch oversight (auctions, settings, revert)
//   admin     — full control (delete, user management, business state)
//
// Capability names are short snake_case strings. New capabilities are
// added by including the name in the appropriate role(s) below.
const ROLE_PERMISSIONS = {
  viewer: new Set([
    'view',           // read any list / detail
    'export',         // download XLSX / PDF / CSV exports
    'self_password', // change own password
    'lot_entry_view'  // also see the Lot Entry tab + its data (read-only)
                      // — viewers don't get lot_write so they can't modify,
                      //   but they can SEE the shared trade/lot data, which
                      //   is the expected behaviour for "authorised users
                      //   viewing shared valid data".
  ]),
  // Field-staff role for the auction-hall lot entry workflow. Sees only
  // the Lot Entry tab in the sidebar (everything else is gated by
  // `view`, which lot_entry intentionally lacks). The narrow scope —
  // create trades, search sellers, create/edit own lots — matches what
  // the original PWA admin exposed to non-admin users in the field.
  lot_entry: new Set([
    'self_password',
    'view',            // read shared trade/lot data — needed so multi-user
                       // sessions can all see the same in-progress entries.
                       // Without this, a second lot_entry user opening the
                       // same trade hit "do not have permission to view"
                       // because some shared endpoints require general view.
    'lot_entry_view',  // Lot Entry tab + its endpoints (auctions, lots, traders)
    'lot_write',       // create/edit own lots; trader search via lot-entry views
    'auction_write'    // create new trades on the fly during an auction day
  ]),
  operator: new Set([
    'view', 'export', 'self_password',
    'lot_entry_view', // operators can also use the Lot Entry tab if they want
    'lot_write',     // create/edit lots, calculate, validate, price-import
    'invoice_write', // generate sales/purchase/bills + edit
    'trader_write',  // create/edit/delete-bank traders
    'buyer_write'    // create/edit buyers (per user decision: tax fields editable)
  ]),
  manager: new Set([
    'view', 'export', 'self_password', 'lot_entry_view',
    'lot_write', 'invoice_write', 'trader_write', 'buyer_write',
    'auction_write',  // create/edit auctions (trades)
    'invoice_revert', // revert sales/purchase/bills (undo invoice)
    'settings_write', // edit company settings (rates, addresses, flags)
    'state_toggle'    // toggle business state TN ↔ KL
  ]),
  admin: new Set([
    'view', 'export', 'self_password', 'lot_entry_view',
    'lot_write', 'invoice_write', 'trader_write', 'buyer_write',
    'auction_write', 'invoice_revert', 'settings_write', 'state_toggle',
    'delete',       // delete any individual record
    'delete_all',   // bulk Delete All (sales, purchases, lots, etc.)
    'user_manage'   // create/delete users, reset passwords, revoke sessions
  ])
};

// Best-effort capability lookup. Unknown roles get treated as 'viewer'
// (safest default — fails closed instead of open).
function userHas(role, capability) {
  const perms = ROLE_PERMISSIONS[role] || ROLE_PERMISSIONS.viewer;
  return perms.has(capability);
}

// ══════════════════════════════════════════════════════════════
// MOBILE PWA MOUNT
// ══════════════════════════════════════════════════════════════
// Field-user lot entry UI ships as a Progressive Web App at /mobile/
// (installable to phone home screen via the manifest in public-mobile/).
// The PWA was originally a standalone app with its own server; the
// `mobile-bridge` module is a compatibility shim that exposes the API
// surface the PWA's app.html expects (/api/auth/*, /api/config,
// /api/lots query-string form, trader-bank CRUD, etc.) while sharing
// the SAME SQLite database as this admin server.
//
// All deltas (~30 endpoints + the static mount) live in one file
// (`mobile-bridge.js`) so the boundary is easy to audit. The shim's
// auth uses verifyPassword/hashPassword — same bcrypt flow as desktop
// /api/login — so a user can log in to either UI with the same creds.
//
// Mounted BEFORE the admin's own routes so /mobile and /api/auth/* are
// claimed first; path-pattern collisions (e.g. /api/lots/print-batch
// vs /api/lots/:auctionId) win the bridge's specific handler.
const { mountMobile } = require('./mobile-bridge');
mountMobile(app, {
  getDb,
  requireAuth,
  verifyPassword,
  hashPassword,
  isLegacyHash,
  ROLE_PERMISSIONS,
});

// Middleware factory: returns an Express middleware that requires the
// authenticated user to have a specific capability.
//
// Usage:
//   app.post('/api/lots', requirePermission('lot_write'), handler)
//   app.delete('/api/invoices/:id', requirePermission('delete'), handler)
//
// Falls through to next() on success; sends 403 with a clear message
// indicating both the user's current role AND the capability required
// (helps the client show a useful error rather than a generic "denied").
function requirePermission(capability) {
  return (req, res, next) => {
    requireAuth(req, res, (err) => {
      if (err) return next(err);
      if (!req.user) return res.status(401).json({ error: 'Authentication required' });
      if (!userHas(req.user.role, capability)) {
        return res.status(403).json({
          error: `Your role (${req.user.role}) does not allow this action`,
          required: capability,
          role: req.user.role
        });
      }
      next();
    });
  };
}

// Same idea but accepts ANY of a list of capabilities. Used for endpoints
// that serve multiple roles — e.g., the trader search endpoint, which
// general operators reach through 'view' and lot-entry users reach
// through their own lot_entry_view capability.
function requireAnyPermission(...capabilities) {
  return (req, res, next) => {
    requireAuth(req, res, (err) => {
      if (err) return next(err);
      if (!req.user) return res.status(401).json({ error: 'Authentication required' });
      const hasAny = capabilities.some(c => userHas(req.user.role, c));
      if (!hasAny) {
        return res.status(403).json({
          error: `Your role (${req.user.role}) does not allow this action`,
          required: capabilities.join(' or '),
          role: req.user.role
        });
      }
      next();
    });
  };
}

// Convenience aliases — readable names for the most common gate points.
// Encapsulates the permission name so callers don't repeat string literals.
const requireView          = requirePermission('view');
const requireViewOrLotEntry = requireAnyPermission('view', 'lot_entry_view');
const requireLotWrite      = requirePermission('lot_write');
const requireInvoiceWrite  = requirePermission('invoice_write');
const requireInvoiceRevert = requirePermission('invoice_revert');
const requireTraderWrite   = requirePermission('trader_write');
const requireBuyerWrite    = requirePermission('buyer_write');
const requireAuctionWrite  = requirePermission('auction_write');
const requireSettingsWrite = requirePermission('settings_write');
const requireStateToggle   = requirePermission('state_toggle');
const requireDelete        = requirePermission('delete');
const requireDeleteAll     = requirePermission('delete_all');
const requireUserManage    = requirePermission('user_manage');
const requireExport        = requirePermission('export');

// ══════════════════════════════════════════════════════════════
// AUTH
// ══════════════════════════════════════════════════════════════

// Public branding (no auth) — login screen / topbar pulls company
// name + logo from settings. Returns only the safe-to-expose subset:
// trade name, short name, branch, GSTIN, and a logo URL when present.
// Theme + theme_custom_color are also exposed so the per-install
// branding (which customer this is) is applied BEFORE login — gives
// every device on this install the same look-and-feel even before
// the user signs in.
app.get('/api/branding', (req, res) => {
  try {
    const cfg = getSettingsFlat(getDb());
    const isKL = String(cfg.business_state || '').toUpperCase().includes('KERALA');
    const branch = (isKL ? cfg.kl_branch : cfg.tn_branch) || cfg.tn_branch || cfg.kl_branch || '';
    const gstin  = (isKL ? cfg.kl_gstin  : cfg.tn_gstin)  || cfg.tn_gstin  || cfg.kl_gstin  || '';
    // Preset = the full white-label bundle (color + density + font +
    // hide-appearance flag) configured by you via /admin/branding.
    // Empty string when no preset has been set — the frontend then
    // treats it as the legacy "default" state with Appearance visible.
    let presetConfig = null;
    if (cfg.preset_config) {
      try { presetConfig = JSON.parse(cfg.preset_config); }
      catch (_) { presetConfig = null; }
    }
    // /logo-ispl.png is always served (falls through to logo_kj.png when
    // the user hasn't uploaded a custom logo — see the route handler
    // registered before the static middleware).
    res.json({
      tradeName: cfg.trade_name || cfg.short_name || '',
      shortName: cfg.short_name || cfg.trade_name || '',
      branch,
      gstin,
      logoUrl: '/logo-ispl.png',
      // Per-install branding choices — empty string when the admin
      // hasn't picked one yet (frontend then falls back to localStorage
      // and finally to 'emerald').
      theme: cfg.theme || '',
      themeCustomColor: cfg.theme_custom_color || '',
      // Preset bundle. preset is the named slug (e.g. 'bluehill');
      // presetConfig is the full {theme, density, font, hideAppearance}
      // payload the frontend applies at boot.
      preset: cfg.tenant_preset || '',
      presetConfig,
    });
  } catch (e) {
    res.json({ tradeName: '', shortName: '', branch: '', gstin: '', logoUrl: null, theme: '', themeCustomColor: '', preset: '', presetConfig: null });
  }
});

// Persist per-install branding (theme + custom color). Upserts directly
// into company_settings — bypasses updateSettings() which only UPDATEs
// existing rows (and these two keys may not be in DEFAULTS yet on
// installs that pre-date the branding feature).
//
// Whitelisted to ONLY these two keys so a misbehaving client can't
// write arbitrary settings through this endpoint. Larger settings
// changes still go through PUT /api/company-settings.
app.put('/api/branding', requireSettingsWrite, (req, res) => {
  try {
    const db = getDb();
    const body = req.body || {};
    const allowed = {};
    if (typeof body.theme === 'string') {
      // Validate against the same theme list the frontend uses
      const THEMES = ['emerald','coral','violet','sunshine','electric','ocean','tech','minimal','trust','rose','indigo','teal','slate','custom'];
      if (THEMES.includes(body.theme)) allowed.theme = body.theme;
    }
    if (typeof body.themeCustomColor === 'string' && /^#[0-9a-fA-F]{6}$/.test(body.themeCustomColor)) {
      allowed.theme_custom_color = body.themeCustomColor;
    }
    if (!Object.keys(allowed).length) {
      return res.status(400).json({ error: 'No valid branding fields supplied' });
    }
    // INSERT OR REPLACE so the row is created on first write and
    // updated on subsequent writes. category='branding' so the row
    // is grouped sensibly if it ever lands in the Settings categories
    // view (currently appearance lives in its own card).
    const stmt = db.prepare(
      `INSERT INTO company_settings (key, value, category, label, field_type)
       VALUES (?, ?, 'branding', ?, 'text')
       ON CONFLICT(key) DO UPDATE SET value = excluded.value`
    );
    const labels = { theme: 'Theme', theme_custom_color: 'Custom primary color' };
    for (const [k, v] of Object.entries(allowed)) {
      stmt.run(k, String(v), labels[k] || k);
    }
    res.json({ success: true, updated: Object.keys(allowed).length });
  } catch (e) {
    res.status(500).json({ error: 'Branding save failed: ' + (e.message || e) });
  }
});

// ══════════════════════════════════════════════════════════════════
// TENANT PRESET (white-label switcher) — DEVELOPER ONLY
// ══════════════════════════════════════════════════════════════════
// The "preset" is a named bundle of color + density + font + hide
// flags. Lets Vinoth configure each customer install to look visually
// distinct without changes to the customer's data. End users never
// see this — the Appearance card in Settings is hidden once a preset
// is active.
//
// Access is gated by ADMIN_BRANDING_KEY (env var). The /admin/branding
// page expects ?key=<secret> on every request. Default key for dev is
// 'change-me' so a non-secured deploy is obvious in the URL.
//
// Pre-baked presets — extend this list to add more. Each preset is a
// pure config object; the frontend reads presetConfig from GET
// /api/branding and applies it at boot.
const TENANT_PRESETS = {
  cardamom: {
    label: 'Cardamom (default — premium spacious)',
    theme: 'emerald',
    customColor: '',
    density: 'roomy',
    font: 'jakarta',
    hideAppearance: true,
  },
  bluehill: {
    label: 'Bluehill (indigo + dense + inter)',
    theme: 'indigo',
    customColor: '',
    density: 'compact',
    font: 'inter',
    hideAppearance: true,
  },
  'western-ghats': {
    label: 'Western Ghats (teal + spacious + outfit)',
    theme: 'teal',
    customColor: '',
    density: 'spacious',
    font: 'outfit',
    hideAppearance: true,
  },
  slate: {
    label: 'Slate (corporate grey + dense + system)',
    theme: 'slate',
    customColor: '',
    density: 'compact',
    font: 'system',
    hideAppearance: true,
  },
  marigold: {
    label: 'Marigold (sunshine + roomy + jakarta)',
    theme: 'sunshine',
    customColor: '',
    density: 'roomy',
    font: 'jakarta',
    hideAppearance: true,
  },
  ocean: {
    label: 'Ocean (cool blue + roomy + inter)',
    theme: 'ocean',
    customColor: '',
    density: 'roomy',
    font: 'inter',
    hideAppearance: true,
  },
};

// Available font slugs that the frontend knows how to apply. Keep in
// sync with the <link>s and CSS variable handling in index.html.
const TENANT_FONTS = ['jakarta', 'inter', 'outfit', 'system'];

// Available density slugs. Frontend maps each to a `data-density` attr
// + CSS rules that adjust padding / corner radius / row heights.
const TENANT_DENSITIES = ['compact', 'roomy', 'spacious'];

// Gatekeeper. Compares against ADMIN_BRANDING_KEY env var, falling back
// to 'change-me' so an unsecured deploy is loud — running ?key=change-me
// in production logs is a clear signal to set the env var.
function checkAdminKey(req) {
  const expected = process.env.ADMIN_BRANDING_KEY || 'change-me';
  const supplied = String(req.query.key || req.headers['x-admin-key'] || '');
  return supplied === expected && supplied.length > 0;
}

// GET /admin/branding?key=… — hidden HTML control panel. Lists presets,
// shows current selection, has a small form for the "custom" preset
// (manual color/density/font picker). NOT linked from anywhere in the
// app; you reach it by typing the URL.
app.get('/admin/branding', (req, res) => {
  if (!checkAdminKey(req)) {
    return res.status(403).type('html').send(
      '<html><body style="font-family:system-ui;padding:40px;text-align:center;color:#666"><h2>403 — Access denied</h2><p>This page requires a valid <code>?key=</code> in the URL.</p></body></html>'
    );
  }
  const db = getDb();
  const cfg = getSettingsFlat(db);
  const currentPreset = cfg.tenant_preset || '';
  let currentConfig = null;
  try { currentConfig = JSON.parse(cfg.preset_config || 'null'); } catch (_) {}

  const presetOptions = Object.entries(TENANT_PRESETS).map(([slug, p]) => {
    const sel = slug === currentPreset ? 'selected' : '';
    return `<option value="${slug}" ${sel}>${slug} — ${p.label}</option>`;
  }).join('');
  const customSel = currentPreset === 'custom' ? 'selected' : '';
  const noneSel = !currentPreset ? 'selected' : '';

  // Custom-preset form field defaults — populated from current config
  // if the active preset IS 'custom', otherwise blank.
  const c = (currentPreset === 'custom' && currentConfig) ? currentConfig : {};
  const cTheme    = c.theme || 'emerald';
  const cColor    = c.customColor || '';
  const cDensity  = c.density || 'roomy';
  const cFont     = c.font || 'jakarta';
  const cHide     = c.hideAppearance !== false; // default true

  const themeOpts = ['emerald','coral','violet','sunshine','electric','ocean','tech','minimal','trust','rose','indigo','teal','slate','custom']
    .map(t => `<option value="${t}" ${t === cTheme ? 'selected' : ''}>${t}</option>`).join('');
  const densityOpts = TENANT_DENSITIES.map(d => `<option value="${d}" ${d === cDensity ? 'selected' : ''}>${d}</option>`).join('');
  const fontOpts = TENANT_FONTS.map(f => `<option value="${f}" ${f === cFont ? 'selected' : ''}>${f}</option>`).join('');

  const keyEsc = String(req.query.key).replace(/[<>'"&]/g, '');

  res.type('html').send(`<!DOCTYPE html>
<html><head><meta charset="utf-8"><title>Tenant Branding — Admin</title>
<style>
  body { font-family: system-ui, -apple-system, sans-serif; max-width: 720px; margin: 40px auto; padding: 20px; color: #1f2937; background: #f9fafb; }
  h1 { font-size: 22px; margin: 0 0 4px; }
  h2 { font-size: 16px; margin: 24px 0 10px; color: #374151; }
  .sub { color: #6b7280; font-size: 13px; margin-bottom: 24px; }
  .card { background: #fff; border: 1px solid #e5e7eb; border-radius: 8px; padding: 20px; margin-bottom: 16px; }
  .current { background: #f0fdf4; border-color: #86efac; }
  label { display: block; font-size: 12px; font-weight: 600; color: #374151; margin: 12px 0 4px; text-transform: uppercase; letter-spacing: .3px; }
  select, input[type=text], input[type=color] { width: 100%; padding: 8px 10px; border: 1px solid #d1d5db; border-radius: 6px; font-size: 14px; box-sizing: border-box; }
  input[type=color] { height: 36px; padding: 2px; }
  input[type=checkbox] { margin-right: 6px; }
  button { background: #166534; color: #fff; border: 0; padding: 10px 18px; border-radius: 6px; font-size: 14px; font-weight: 600; cursor: pointer; }
  button.secondary { background: #6b7280; }
  button:hover { opacity: .9; }
  .row { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }
  pre { background: #f3f4f6; padding: 10px; border-radius: 6px; font-size: 12px; overflow-x: auto; }
  .msg { display: none; padding: 10px; border-radius: 6px; margin-bottom: 16px; }
  .msg.ok { display: block; background: #dcfce7; color: #14532d; }
  .msg.err { display: block; background: #fee2e2; color: #991b1b; }
</style></head>
<body>
  <h1>Tenant branding</h1>
  <p class="sub">Hidden admin panel. Sets the white-label preset for this install. End users never see this URL or page.</p>

  <div id="msg" class="msg"></div>

  <div class="card current">
    <h2>Current setting</h2>
    <p><strong>Active preset:</strong> ${currentPreset || '<em>none (legacy mode — Appearance card visible to users)</em>'}</p>
    ${currentConfig ? `<pre>${JSON.stringify(currentConfig, null, 2)}</pre>` : ''}
  </div>

  <div class="card">
    <h2>Pick a pre-baked preset</h2>
    <label>Preset</label>
    <select id="preset-pick">
      <option value="" ${noneSel}>— none (show Appearance card to users) —</option>
      ${presetOptions}
      <option value="custom" ${customSel}>custom (use the form below)</option>
    </select>
    <p style="font-size: 12px; color: #6b7280; margin: 10px 0 0">Pick a named preset OR pick "custom" and fill in the form below.</p>
  </div>

  <div class="card">
    <h2>Custom preset (used only when preset = "custom")</h2>
    <div class="row">
      <div>
        <label>Color theme</label>
        <select id="custom-theme">${themeOpts}</select>
      </div>
      <div>
        <label>Custom hex (used when theme = "custom")</label>
        <input type="color" id="custom-color" value="${cColor || '#166534'}">
      </div>
      <div>
        <label>Density</label>
        <select id="custom-density">${densityOpts}</select>
      </div>
      <div>
        <label>Font</label>
        <select id="custom-font">${fontOpts}</select>
      </div>
    </div>
    <label style="margin-top: 16px;">
      <input type="checkbox" id="custom-hide" ${cHide ? 'checked' : ''}>
      Hide Appearance card from users
    </label>
  </div>

  <div style="display: flex; gap: 10px;">
    <button onclick="apply()">Apply preset</button>
    <button class="secondary" onclick="clearPreset()">Clear (revert to legacy mode)</button>
  </div>

  <script>
    const KEY = ${JSON.stringify(keyEsc)};
    const FONTS = ${JSON.stringify(TENANT_FONTS)};
    const DENSITIES = ${JSON.stringify(TENANT_DENSITIES)};
    function msg(text, ok) {
      const el = document.getElementById('msg');
      el.className = 'msg ' + (ok ? 'ok' : 'err');
      el.textContent = text;
    }
    async function apply() {
      const preset = document.getElementById('preset-pick').value;
      const body = { preset };
      if (preset === 'custom') {
        body.config = {
          theme: document.getElementById('custom-theme').value,
          customColor: document.getElementById('custom-color').value,
          density: document.getElementById('custom-density').value,
          font: document.getElementById('custom-font').value,
          hideAppearance: document.getElementById('custom-hide').checked,
        };
      }
      try {
        const r = await fetch('/api/admin/preset?key=' + encodeURIComponent(KEY), {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(body),
        });
        const d = await r.json();
        if (!r.ok) return msg(d.error || 'Failed', false);
        msg('Applied. Reload your customer-facing app to see the change.', true);
        setTimeout(() => location.reload(), 800);
      } catch (e) { msg(e.message, false); }
    }
    async function clearPreset() {
      if (!confirm('Clear the active preset and show Appearance to users again?')) return;
      try {
        const r = await fetch('/api/admin/preset?key=' + encodeURIComponent(KEY), {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ preset: '' }),
        });
        if (!r.ok) return msg('Failed to clear', false);
        msg('Cleared.', true);
        setTimeout(() => location.reload(), 800);
      } catch (e) { msg(e.message, false); }
    }
  </script>
</body></html>`);
});

// POST /api/admin/preset — receives the preset pick from the admin page.
// Key-gated like the GET. Writes tenant_preset (slug) + preset_config
// (full bundle as JSON) to company_settings so the frontend can pick
// it up via GET /api/branding without further auth.
app.post('/api/admin/preset', (req, res) => {
  if (!checkAdminKey(req)) return res.status(403).json({ error: 'Invalid admin key' });
  const db = getDb();
  const { preset, config } = req.body || {};
  const slug = String(preset || '').trim();

  let finalConfig = null;
  if (slug === '') {
    // Clear — empty preset + clear stored config so the frontend reverts
    // to legacy mode and the Appearance card reappears.
    finalConfig = null;
  } else if (slug === 'custom') {
    if (!config || typeof config !== 'object') {
      return res.status(400).json({ error: 'custom preset requires a config object' });
    }
    // Whitelist fields so an attacker can't shove arbitrary JSON.
    finalConfig = {
      theme: String(config.theme || 'emerald'),
      customColor: /^#[0-9a-fA-F]{6}$/.test(config.customColor || '') ? config.customColor : '',
      density: TENANT_DENSITIES.includes(config.density) ? config.density : 'roomy',
      font: TENANT_FONTS.includes(config.font) ? config.font : 'jakarta',
      hideAppearance: !!config.hideAppearance,
    };
  } else if (TENANT_PRESETS[slug]) {
    finalConfig = TENANT_PRESETS[slug];
  } else {
    return res.status(400).json({ error: `Unknown preset: ${slug}` });
  }

  // Upsert both keys. Same INSERT OR REPLACE pattern as PUT /api/branding
  // because tenant_preset / preset_config don't live in DEFAULTS.
  const stmt = db.prepare(
    `INSERT INTO company_settings (key, value, category, label, field_type)
     VALUES (?, ?, 'branding', ?, 'text')
     ON CONFLICT(key) DO UPDATE SET value = excluded.value`
  );
  stmt.run('tenant_preset', slug, 'Tenant preset');
  stmt.run('preset_config', finalConfig ? JSON.stringify(finalConfig) : '', 'Tenant preset config (JSON)');
  res.json({ success: true, preset: slug, config: finalConfig });
});

// Login rate limiter — 10 attempts per 15-minute window per IP. Returns
// 429 on overflow. Skip-on-success ensures legitimate users who type a
// wrong password once aren't penalised on every later request after they
// log in correctly. Defeats credential-stuffing / online brute-force.
const loginLimiter = rateLimit({
  windowMs: 15 * 60 * 1000,
  max: 10,
  standardHeaders: true,
  legacyHeaders: false,
  skipSuccessfulRequests: true,
  message: { error: 'Too many login attempts. Please wait a few minutes and try again.' }
});

// Strict cap on session lifetime. A leaked Authorization header is now
// only valid for at most this many days from creation, even if the
// attacker keeps the session "active" by hitting endpoints (which
// previously kept it alive forever via the sliding-window sweep).
const SESSION_TTL_DAYS = 30;

// ══════════════════════════════════════════════════════════════
// LICENSING — install ID, expiry probe, token apply
// ══════════════════════════════════════════════════════════════
//
// Both endpoints are deliberately auth-free: the renewal page lives
// outside the normal authenticated flow (the operator is locked out
// of /api/login while expired, so they need a way to apply a new
// token from a logged-out state). status is also used by the topbar
// countdown pill on the main UI.
app.get('/api/license/status', (req, res) => {
  res.set('Cache-Control', 'no-store');
  try {
    const status = license.getStatus(getDb());
    res.json(status);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});
app.post('/api/license/apply', (req, res) => {
  const token = (req.body && (req.body.token || req.body.license_token) || '').trim();
  if (!token) return res.status(400).json({ error: 'token required' });
  const result = license.applyToken(getDb(), token);
  if (!result.ok) return res.status(400).json(result);
  res.json(result);
});

// ─────────────────────────────────────────────────────────────
// Admin endpoint — back-channel for ops without shell access.
//
// Authenticated by the same LICENSE_SECRET used to sign tokens
// (only the developer knows it). When LICENSE_SECRET is unset on
// the server, this endpoint refuses ALL requests so the fallback
// development secret in license.js can never double as a remote
// admin backdoor on a misconfigured production deploy.
//
// Usage from anywhere (no shell required):
//
//   curl -X POST https://<app>/api/license/admin/set-expiry \
//     -H "content-type: application/json" \
//     -H "X-License-Secret: $LICENSE_SECRET" \
//     -d '{"expires_at":"2020-01-01"}'      # force expired (test)
//
//   curl -X POST https://<app>/api/license/admin/set-expiry \
//     -H "content-type: application/json" \
//     -H "X-License-Secret: $LICENSE_SECRET" \
//     -d '{"expires_at":"2026-07-02"}'      # restore
//
// Sets the row's active_token to NULL so the audit trail doesn't
// falsely attribute the new expiry to a previously-applied token.
// For token-driven history use POST /api/license/apply.
// ─────────────────────────────────────────────────────────────
function _requireLicenseAdmin(req) {
  const envSecret = String(process.env.LICENSE_SECRET || '').trim();
  if (!envSecret) {
    return { status: 403, error: 'admin disabled: LICENSE_SECRET env var is not set on this server' };
  }
  const provided = String(req.headers['x-license-secret'] || '').trim();
  if (!provided) {
    return { status: 403, error: 'X-License-Secret header required' };
  }
  // Constant-time compare on equal-length buffers. Bail before the
  // compare when lengths differ so we don't leak the secret length
  // (and so timingSafeEqual doesn't throw on a length mismatch).
  const a = Buffer.from(envSecret, 'utf8');
  const b = Buffer.from(provided, 'utf8');
  if (a.length !== b.length || !crypto.timingSafeEqual(a, b)) {
    return { status: 403, error: 'invalid X-License-Secret' };
  }
  return null;
}

app.post('/api/license/admin/set-expiry', (req, res) => {
  res.set('Cache-Control', 'no-store');
  const denied = _requireLicenseAdmin(req);
  if (denied) return res.status(denied.status).json({ error: denied.error });

  const body = req.body || {};
  const expires_at = String(body.expires_at || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(expires_at)) {
    return res.status(400).json({ error: 'expires_at must be YYYY-MM-DD (e.g. 2026-07-31)' });
  }
  // Sanity bound — anything outside 2020-2099 is almost certainly a typo.
  // Tokens already issued past 2099 would also be unusable.
  const yr = Number(expires_at.slice(0, 4));
  if (yr < 2020 || yr > 2099) {
    return res.status(400).json({ error: 'expires_at year out of range (2020-2099)' });
  }

  try {
    const db = getDb();
    // Make sure the row exists — on a fresh install this endpoint may
    // be hit before any normal traffic has triggered the bootstrap.
    license.ensureState(db);
    db.run(
      'UPDATE license_state SET expires_at = ?, active_token = NULL WHERE id = 1',
      [expires_at]
    );
    res.json({ ok: true, status: license.getStatus(db) });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.post('/api/login', loginLimiter, async (req, res) => {
  // Hard licence gate before the credential check — when the install
  // is expired we don't even pretend to authenticate. The frontend
  // login script catches the 451 and redirects to /renew.html with
  // the install ID pre-filled.
  try {
    const lstatus = license.getStatus(getDb());
    if (lstatus.expired) {
      return res.status(451).json({
        error: 'license_expired',
        message: `Your access window ended on ${lstatus.expires_at}. Send the install ID below to your provider to receive a renewal token.`,
        install_id: lstatus.install_id,
        expires_at: lstatus.expires_at,
      });
    }
  } catch (_) { /* license check failed → let login through; fail-open is friendlier than a brick */ }
  const { username, password, device_label } = req.body || {};
  if (!username || !password) return res.status(401).json({ error: 'Invalid credentials' });
  const db = getDb();
  const user = db.get('SELECT * FROM users WHERE username = ?', [username]);
  // Always run a bcrypt verify — against the real hash if the user exists,
  // against a fixed dummy hash if not — so the response time doesn't leak
  // whether the username is registered. Single uniform error message.
  const ok = await verifyPassword(password, user ? user.password_hash : DUMMY_BCRYPT_HASH);
  if (!user || !ok) return res.status(401).json({ error: 'Invalid credentials' });

  // Opportunistic rehash: if the stored hash is legacy SHA-256, upgrade it
  // to bcrypt now that we have the plaintext. New users + admin resets
  // already write bcrypt directly; this catches existing rows on first login.
  if (isLegacyHash(user.password_hash)) {
    try {
      const upgraded = await hashPassword(password);
      db.run('UPDATE users SET password_hash = ? WHERE id = ?', [upgraded, user.id]);
    } catch (_) { /* non-fatal — login continues with legacy hash, retry next time */ }
  }

  const token = crypto.randomBytes(32).toString('hex');
  // Create a new session row WITHOUT deleting any existing sessions —
  // this lets the same user stay logged in on multiple devices simultaneously.
  db.run(
    `INSERT INTO sessions (token, user_id, device_label, expires_at)
     VALUES (?, ?, ?, datetime('now','localtime','+${SESSION_TTL_DAYS} days'))`,
    [token, user.id, device_label || '']
  );
  // Clean up very old sessions (expired OR > 30d idle) so the table doesn't grow forever
  db.run(
    `DELETE FROM sessions
     WHERE (expires_at IS NOT NULL AND expires_at < datetime('now','localtime'))
        OR last_used_at < datetime('now','-30 days')`
  );
  // Return the user's capabilities array so the client can hide buttons
  // they're not allowed to use. Server still validates every request.
  const permissions = Array.from(ROLE_PERMISSIONS[user.role] || ROLE_PERMISSIONS.viewer);
  res.json({
    token, role: user.role, username: user.username, permissions,
    must_change_password: !!user.must_change_password
  });
});
app.post('/api/logout', (req, res) => {
  const t = (req.headers.authorization||'').replace('Bearer ','');
  if (t) getDb().run('DELETE FROM sessions WHERE token = ?', [t]);
  res.json({ success: true });
});
app.get('/api/me', requireAnyPermission('view', 'lot_entry_view', 'self_password'), (req, res) => {
  const permissions = Array.from(ROLE_PERMISSIONS[req.user.role] || ROLE_PERMISSIONS.viewer);
  res.json({
    username: req.user.username,
    role: req.user.role,
    permissions,
    must_change_password: !!req.user.must_change_password
  });
});

// ══════════════════════════════════════════════════════════════
// USER MANAGEMENT (admin-only)
// ══════════════════════════════════════════════════════════════
app.get('/api/users', requireUserManage, (req, res) => {
  const db = getDb();
  const users = db.all(`
    SELECT u.id, u.username, u.role, u.created_at,
      (SELECT COUNT(*) FROM sessions s WHERE s.user_id = u.id) as active_sessions,
      (SELECT MAX(last_used_at) FROM sessions s WHERE s.user_id = u.id) as last_active
    FROM users u ORDER BY u.id ASC
  `);
  res.json(users);
});

app.post('/api/users', requireUserManage, async (req, res) => {
  const { username, password, role } = req.body || {};
  if (!username || !password) return res.status(400).json({ error: 'username and password are required' });
  if (password.length < 4) return res.status(400).json({ error: 'Password must be at least 4 characters' });
  if (!/^[a-zA-Z0-9_.-]{3,30}$/.test(username)) return res.status(400).json({ error: 'Username: 3–30 chars, letters/digits/._- only' });
  // Validate role against the known set. Default to 'operator' (the most
  // common and least privileged write-capable role) if missing or invalid.
  // Legacy 'user' role is mapped to 'viewer' for backward compat.
  const VALID_ROLES = ['viewer', 'lot_entry', 'operator', 'manager', 'admin'];
  let finalRole = (role || '').toLowerCase();
  if (finalRole === 'user') finalRole = 'viewer';
  if (!VALID_ROLES.includes(finalRole)) finalRole = 'operator';
  const db = getDb();
  const existing = db.get('SELECT id FROM users WHERE username = ?', [username]);
  if (existing) return res.status(400).json({ error: 'Username already exists' });
  // New users start with must_change_password=1 — the admin-supplied
  // password is a one-time hand-off, not the user's chosen credential.
  const pwHash = await hashPassword(password);
  db.run(
    'INSERT INTO users (username, password_hash, role, must_change_password) VALUES (?, ?, ?, 1)',
    [username, pwHash, finalRole]
  );
  const created = db.get('SELECT id, username, role FROM users WHERE username = ?', [username]);
  res.json({ success: true, id: created ? created.id : null, username, role: finalRole });
});

// Update an existing user's role (for promoting/demoting users without
// recreating them). Admin-only — same gate as creating users.
app.put('/api/users/:id/role', requireUserManage, (req, res) => {
  const { role } = req.body || {};
  const VALID_ROLES = ['viewer', 'lot_entry', 'operator', 'manager', 'admin'];
  let finalRole = String(role || '').toLowerCase();
  if (finalRole === 'user') finalRole = 'viewer';
  if (!VALID_ROLES.includes(finalRole)) {
    return res.status(400).json({ error: `Role must be one of: ${VALID_ROLES.join(', ')}` });
  }
  const db = getDb();
  const target = db.get('SELECT id, username, role FROM users WHERE id = ?', [req.params.id]);
  if (!target) return res.status(404).json({ error: 'User not found' });
  // Safety: don't let admins demote the last remaining admin (would lock
  // everyone out of user management).
  if (target.role === 'admin' && finalRole !== 'admin') {
    const adminCount = db.get(`SELECT COUNT(*) as n FROM users WHERE role = 'admin'`).n;
    if (adminCount <= 1) {
      return res.status(400).json({ error: 'Cannot demote the last admin — promote someone else first' });
    }
  }
  db.run('UPDATE users SET role = ? WHERE id = ?', [finalRole, target.id]);
  res.json({ success: true, username: target.username, role: finalRole });
});

app.put('/api/users/:id/password', requireUserManage, async (req, res) => {
  const { password } = req.body || {};
  if (!password || password.length < 4) return res.status(400).json({ error: 'Password must be at least 4 characters' });
  const db = getDb();
  const user = db.get('SELECT id, username FROM users WHERE id = ?', [req.params.id]);
  if (!user) return res.status(404).json({ error: 'User not found' });
  // Force the target to pick a new password on next login — the admin-set
  // value is a one-time hand-off, not the user's chosen credential.
  const pwHash = await hashPassword(password);
  db.run(
    'UPDATE users SET password_hash = ?, must_change_password = 1 WHERE id = ?',
    [pwHash, user.id]
  );
  // Invalidate all sessions of that user (force re-login after password change)
  db.run('DELETE FROM sessions WHERE user_id = ?', [user.id]);
  res.json({ success: true, username: user.username });
});

app.delete('/api/users/:id', requireUserManage, (req, res) => {
  const db = getDb();
  const target = db.get('SELECT id, username, role FROM users WHERE id = ?', [req.params.id]);
  if (!target) return res.status(404).json({ error: 'User not found' });
  // Safety: don't let admin delete themselves
  if (target.id === req.user.id) return res.status(400).json({ error: 'You cannot delete your own account while signed in' });
  // Safety: never allow deleting the last remaining user
  const total = db.get('SELECT COUNT(*) as c FROM users').c;
  if (total <= 1) return res.status(400).json({ error: 'Cannot delete the last remaining user' });
  db.run('DELETE FROM sessions WHERE user_id = ?', [target.id]);
  db.run('DELETE FROM users WHERE id = ?', [target.id]);
  res.json({ success: true, username: target.username });
});

// Change own password — shortcut that doesn't require user id
app.put('/api/me/password', requirePermission('self_password'), async (req, res) => {
  const { current_password, new_password } = req.body || {};
  if (!current_password || !new_password) return res.status(400).json({ error: 'Both current and new password required' });
  if (new_password.length < 4) return res.status(400).json({ error: 'New password must be at least 4 characters' });
  const db = getDb();
  const user = db.get('SELECT * FROM users WHERE id = ?', [req.user.id]);
  if (!user || !(await verifyPassword(current_password, user.password_hash))) {
    return res.status(401).json({ error: 'Current password is incorrect' });
  }
  if (await verifyPassword(new_password, user.password_hash)) {
    return res.status(400).json({ error: 'New password must be different from current password' });
  }
  // Clear must_change_password — this is the gate exit for the seeded
  // admin (and any user the admin force-resets, see PUT /api/users/:id/password).
  const newHash = await hashPassword(new_password);
  db.run(
    'UPDATE users SET password_hash = ?, must_change_password = 0 WHERE id = ?',
    [newHash, user.id]
  );
  // Kill all OTHER sessions (keep current one)
  db.run('DELETE FROM sessions WHERE user_id = ? AND token != ?', [user.id, req.session.token]);
  res.json({ success: true });
});

// See my active sessions (all devices signed in as me)
app.get('/api/me/sessions', requireView, (req, res) => {
  const db = getDb();
  const sessions = db.all(
    `SELECT token, device_label, created_at, last_used_at,
            CASE WHEN token = ? THEN 1 ELSE 0 END as is_current
     FROM sessions WHERE user_id = ? ORDER BY last_used_at DESC`,
    [req.session.token, req.user.id]
  );
  // Mask tokens — only show last 8 chars
  res.json(sessions.map(s => ({ ...s, token: '…' + (s.token || '').slice(-8) })));
});

// Revoke (log out) another session I own
app.delete('/api/me/sessions/:tokenSuffix', requireView, (req, res) => {
  const suffix = req.params.tokenSuffix;
  const db = getDb();
  // Find session by matching suffix, for THIS user only
  const sessions = db.all('SELECT token FROM sessions WHERE user_id = ?', [req.user.id]);
  const match = sessions.find(s => (s.token || '').endsWith(suffix));
  if (!match) return res.status(404).json({ error: 'Session not found' });
  if (match.token === req.session.token) return res.status(400).json({ error: 'Use Logout to end your current session' });
  db.run('DELETE FROM sessions WHERE token = ?', [match.token]);
  res.json({ success: true });
});

// ══════════════════════════════════════════════════════════════
// COMPANY SETTINGS
// ══════════════════════════════════════════════════════════════
app.get('/api/company-settings', requireViewOrLotEntry, (req, res) => {
  res.json({ categories: CATEGORIES, settings: getAllSettings(getDb()) });
});
// Settings whose value feeds calculateLot — i.e. changing any of them makes
// every already-calculated lot's stamped planter columns (prate/puramt/refund/
// cgst/sgst/igst/advance/balance) stale. The classic symptom: turning on
// "Discount includes GST" (flag_disc_gst) AFTER lots were calculated left the
// per-lot GST stamps at 0 while the Payments tab computed GST live, so the two
// disagreed. Re-running the calc on a settings change keeps the stamps honest.
const CALC_AFFECTING_SETTINGS = [
  'flag_disc_gst', 'discount_gst', 'gst_service',
  'flag_sample', 'discount_pct', 'discount_days', 'dealer_days',
  'deduction1', 'deduction2', 'deduction1_inclusive', 'flag_discount_in_prate',
  'refund', 'business_state',
];
// Re-run calculateLot over every UNLOCKED, calculable lot (all auctions) so
// the stamped columns reflect the current settings. Same row set + UPDATE the
// manual "Calculate All" uses; locked (finalised) lots are left untouched.
function recalcUnlockedLots(db, cfg) {
  const lots = db.all(
    `SELECT * FROM lots
       WHERE locked_at IS NULL
         AND ( amount > 0
               OR (qty > 0 AND price > 0)
               OR puramt > 0 OR prate > 0
               OR cgst > 0 OR sgst > 0 OR igst > 0 )`
  );
  let n = 0;
  for (const lot of lots) {
    const c = calculateLot(lot, cfg);
    db.run(`UPDATE lots SET amount=?,pqty=?,prate=?,puramt=?,com=?,sertax=?,cgst=?,sgst=?,igst=?,advance=?,balance=?,bilamt=?,refund=?,refud=?,isp_pqty=?,isp_prate=?,isp_puramt=?,asp_pqty=?,asp_prate=?,asp_puramt=? WHERE id=?`,
      [c.amount,c.pqty,c.prate,c.puramt,c.com,c.sertax,c.cgst,c.sgst,c.igst,c.advance,c.balance,c.bilamt,c.refund||0,c.refud||0,c.isp_pqty||0,c.isp_prate||0,c.isp_puramt||0,c.asp_pqty||0,c.asp_prate||0,c.asp_puramt||0,lot.id]);
    n++;
  }
  return n;
}

// Rates & charges whose changes are logged to settings_history and shown
// on hover in the Settings → Rates panel. `insurance` is the rates-section
// key (the To Tally panel's tally_insurance_rate is a separate field).
const TRACKED_HISTORY_KEYS = ['gunny_rate', 'transport', 'insurance', 'discount_pct', 'discount_days', 'dealer_days'];

app.put('/api/company-settings', requireSettingsWrite, (req, res) => {
  const db = getDb();
  const incoming = req.body.settings || {};
  // Snapshot the calc-affecting settings BEFORE the write so we can tell
  // whether a recalc is actually needed (avoid re-stamping every lot on an
  // unrelated change like a phone number or theme).
  const before = getSettingsFlat(db);
  const count = updateSettings(db, incoming);
  // Drop the cached date-format whenever settings change so fmtDate
  // picks up the new value on the next call instead of serving the
  // stale one for the rest of the process lifetime.
  invalidateDateFormatCache();
  const after = getSettingsFlat(db);

  // Log rates/charges changes to settings_history. getSettingsFlat
  // normalises numbers (parseFloat), so "2.5" vs "2.50" won't register
  // as a change — only a real value difference is recorded. Other panels
  // never touch these keys, so before===after and nothing is logged.
  try {
    const histStmt = db.prepare(
      `INSERT INTO settings_history (key, old_value, new_value)
       VALUES (?, ?, ?)`
    );
    for (const k of TRACKED_HISTORY_KEYS) {
      const ov = String(before[k] ?? '');
      const nv = String(after[k] ?? '');
      if (ov !== nv) histStmt.run(k, ov, nv);
    }
  } catch (e) {
    console.warn('[settings history] log failed:', e.message);
  }
  const calcChanged = CALC_AFFECTING_SETTINGS.some(
    k => String(before[k] ?? '') !== String(after[k] ?? '')
  );
  let recalculated = 0;
  if (calcChanged) {
    try { recalculated = recalcUnlockedLots(db, after); }
    catch (e) { console.error('[settings recalc] failed:', e && (e.stack || e.message || e)); }
  }
  res.json({ success: true, updated: count, recalculated });
});
app.get('/api/company-settings/flat', requireViewOrLotEntry, (req, res) => res.json(getSettingsFlat(getDb())));

// Change history for the tracked rates/charges settings. Returns a map
// of { key: [{ old_value, new_value, changed_at }, ...newest first] }.
// Powers the hover-to-view history panel under Settings → Rates.
app.get('/api/settings-history', requireView, (req, res) => {
  try {
    const db = getDb();
    const stmt = db.prepare(
      `SELECT old_value, new_value, changed_at
         FROM settings_history
        WHERE key = ?
        ORDER BY id DESC
        LIMIT 50`
    );
    const out = {};
    for (const k of TRACKED_HISTORY_KEYS) out[k] = stmt.all(k);
    res.json(out);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── Company identity presets — REMOVED in e-Trade-only build ──────────
// The original Spice Config app had ISP/ASP preset switching tied to the
// Logo Code dropdown. This build is a single-company app, so the endpoints
// return harmless empty payloads to keep older clients from erroring out
// while they're still cached. The `active` field reflects the user's
// configured short name (Logo Code) rather than a hardcoded literal.
function _activeLabel() {
  try {
    const id = getCompanyIdentity(getSettingsFlat(getDb()));
    return id.shortName || id.logoCode || '';
  } catch (_) { return ''; }
}
app.get('/api/company-presets', requireView, (_req, res) => {
  const a = _activeLabel();
  res.json({ [a || 'default']: {}, active: a });
});
app.put('/api/company-presets/active', requireStateToggle, (_req, res) => {
  res.json({ success: true, active: _activeLabel() });
});
app.put('/api/company-presets/:code', requireSettingsWrite, (_req, res) => {
  res.json({ success: true });
});

// ── Logo upload/delete ────────────────────────────────────────
// Logos are stored as BLOBs in the company_logos table so they persist
// wherever the SQLite DB persists — no separate volume mount needed for
// the filesystem upload directory on Railway / Heroku / Fly. The
// historical /public/logo-*.png defaults still ship in the bundle and
// are served as a fallback by the /logo-ispl.png and /logo-asp.png
// routes when no DB row exists for the slot.
// e-Trade-only build: a single 'ispl' logo. The 'asp' slot is preserved so
// older client code that still POSTs there gets a clean error rather than
// an unhandled exception.
const LOGO_SLOTS = new Set(['ispl', 'asp']);
function _logoMimeForExt(ext) {
  return (ext === 'jpg' || ext === 'jpeg') ? 'image/jpeg' : 'image/png';
}
app.post('/api/company-settings/logo/:which', requireSettingsWrite, upload.single('file'), (req, res) => {
  const which = req.params.which;
  if (!LOGO_SLOTS.has(which)) return res.status(400).json({ error: 'Invalid logo type (use ispl or asp)' });
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
  const ext = (req.file.originalname.split('.').pop() || '').toLowerCase();
  if (!['png', 'jpg', 'jpeg'].includes(ext)) {
    fs.unlinkSync(req.file.path);
    return res.status(400).json({ error: 'Only PNG or JPG images allowed' });
  }
  // Read the uploaded file into memory and persist as a BLOB. Multer's
  // tmp file is removed afterwards regardless of DB success so we don't
  // leak disk space.
  let bytes;
  try {
    bytes = fs.readFileSync(req.file.path);
  } finally {
    try { fs.unlinkSync(req.file.path); } catch (_) {}
  }
  const mime = _logoMimeForExt(ext);
  const db = getDb();
  db.run(
    `INSERT INTO company_logos (key, mime, data, updated_at)
     VALUES (?, ?, ?, datetime('now','localtime'))
     ON CONFLICT(key) DO UPDATE SET
       mime = excluded.mime,
       data = excluded.data,
       updated_at = excluded.updated_at`,
    [which, mime, bytes]
  );
  res.json({ success: true, path: `/logo-${which}.png`, size: bytes.length });
});
app.delete('/api/company-settings/logo/:which', requireSettingsWrite, (req, res) => {
  const which = req.params.which;
  if (!LOGO_SLOTS.has(which)) return res.status(400).json({ error: 'Invalid logo type' });
  getDb().run('DELETE FROM company_logos WHERE key = ?', [which]);
  res.json({ success: true });
});
// Quick probe so the UI knows whether a logo is uploaded
app.get('/api/company-settings/logo/:which', requireView, (req, res) => {
  const which = req.params.which;
  if (!LOGO_SLOTS.has(which)) return res.status(400).json({ error: 'Invalid logo type' });
  const row = getDb().get(
    'SELECT LENGTH(data) AS size, updated_at AS mtime FROM company_logos WHERE key = ?',
    [which]
  );
  if (!row || !row.size) return res.json({ exists: false });
  res.json({ exists: true, size: row.size, mtime: row.mtime });
});

app.get('/api/company-settings/export', requireExport, (req, res) => {
  res.setHeader('Content-Disposition', 'attachment; filename="company-settings.json"');
  res.json(getSettingsFlat(getDb()));
});
app.post('/api/company-settings/import', requireSettingsWrite, (req, res) => {
  const count = updateSettings(getDb(), req.body.settings || {});
  res.json({ success: true, imported: count });
});

// ══════════════════════════════════════════════════════════════
// BULK DELETE ROUTES (DELETE ALL records from a given table)
// ══════════════════════════════════════════════════════════════
// Every wipe is wrapped in three safeguards:
//   1. Auto-backup — the live SQLite file is copied to data/backups/
//      with a timestamped, resource-tagged filename BEFORE any DELETE
//      runs. One stray click is recoverable by restoring the snapshot.
//   2. Audit log — a row in delete_log captures who/when/what/how-many
//      plus the backup path, so a wipe can be traced months later.
//   3. Cascade-count preflight — the client can hit /preflight to show
//      "About to delete 1,247 invoices…" with the real number before
//      asking the operator to type DELETE.
//
// `requireDeleteAll` is still enforced by Express middleware on every
// route below; these protections are layered on top of it, not in
// place of it.

// Map of resource → { table, cascade: [tables wiped alongside], scope }.
// `scope` is 'global' (wipes the whole table) or 'trade' (requires
// ?ano= and only deletes rows matching that ano).
const DELETE_ALL_RESOURCES = {
  traders:      { table: 'traders',     cascade: ['trader_banks'],                                          scope: 'global' },
  buyers:       { table: 'buyers',      cascade: [],                                                        scope: 'global' },
  // Combined master wipe — clears BOTH the sellers (traders + their
  // bank rows) and buyers masters in one audited action, so a full
  // master reset doesn't need two separate Delete All clicks.
  masters:      { table: 'traders',     cascade: ['trader_banks','buyers'],                                 scope: 'global' },
  invoices:     { table: 'invoices',    cascade: [],                                                        scope: 'global' },
  purchases:    { table: 'purchases',   cascade: [],                                                        scope: 'global' },
  bills:        { table: 'bills',       cascade: [],                                                        scope: 'global' },
  auctions:     { table: 'auctions',    cascade: ['lots','invoices','purchases','bills','debit_notes','lot_allocations'], scope: 'global' },
  'debit-notes': { table: 'debit_notes', cascade: [],                                                       scope: 'trade' },
};

// Snapshot the live SQLite file before any destructive operation.
// Returns the absolute backup path so the audit row can point at it.
// Failure is fatal — we'd rather refuse the wipe than do it without a
// safety net.
function _snapshotBackupBeforeDelete(resource) {
  const backupDir = path.join(process.env.SPICE_DATA_DIR || path.join(__dirname, 'data'), 'backups');
  try { fs.mkdirSync(backupDir, { recursive: true }); } catch (_) {}
  const stamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
  const filename = `before-delete-${resource}-${stamp}.db`;
  const target = path.join(backupDir, filename);
  // Persist any pending in-memory state first so the snapshot is consistent.
  try { require('./db').flushSave(); } catch (_) {}
  fs.copyFileSync(DB_PATH, target);
  return target;
}

function _logDelete(db, { resource, deletedCount, cascadeCounts, backupPath, req }) {
  try {
    db.run(
      `INSERT INTO delete_log (resource, deleted_count, cascade_counts, backup_path, user_id, username, ip)
       VALUES (?,?,?,?,?,?,?)`,
      [
        resource,
        deletedCount,
        JSON.stringify(cascadeCounts || {}),
        backupPath || '',
        (req.user && req.user.id) || null,
        (req.user && req.user.username) || '',
        String(req.ip || req.headers['x-forwarded-for'] || '').slice(0, 64),
      ],
    );
  } catch (_) { /* best-effort */ }
}

// Count helper used both by the preflight endpoint and by the wipe
// itself (so we record the exact row count being deleted).
function _countDeleteAllImpact(db, resource, ano) {
  const def = DELETE_ALL_RESOURCES[resource];
  if (!def) return null;
  const counts = {};
  if (def.scope === 'trade') {
    counts[def.table] = db.get(`SELECT COUNT(*) as c FROM ${def.table} WHERE ano = ?`, [ano || '']).c;
  } else {
    counts[def.table] = db.get(`SELECT COUNT(*) as c FROM ${def.table}`).c;
    for (const t of def.cascade) {
      try { counts[t] = db.get(`SELECT COUNT(*) as c FROM ${t}`).c; }
      catch (_) { counts[t] = 0; }
    }
  }
  return counts;
}

// GET /api/admin/delete-all/preflight?resource=invoices[&ano=16]
// Returns the row count of the target table + each cascade table, so
// the client can show "About to delete 1,247 invoices" before asking
// for typed confirmation.
app.get('/api/admin/delete-all/preflight', requireDeleteAll, (req, res) => {
  const resource = String(req.query.resource || '').trim();
  const def = DELETE_ALL_RESOURCES[resource];
  if (!def) return res.status(400).json({ error: 'Unknown resource', available: Object.keys(DELETE_ALL_RESOURCES) });
  if (def.scope === 'trade' && !String(req.query.ano || '').trim()) {
    return res.status(400).json({ error: 'ano query param required for trade-scoped delete' });
  }
  const counts = _countDeleteAllImpact(getDb(), resource, req.query.ano);
  res.json({ resource, scope: def.scope, counts });
});

// GET /api/admin/delete-log — recent Delete All audit entries.
app.get('/api/admin/delete-log', requireDeleteAll, (req, res) => {
  const rows = getDb().all(
    `SELECT id, resource, deleted_count, cascade_counts, backup_path, username, ip, created_at
       FROM delete_log ORDER BY id DESC LIMIT 200`
  );
  res.json(rows.map(r => ({
    ...r,
    cascade_counts: (() => { try { return JSON.parse(r.cascade_counts || '{}'); } catch (_) { return {}; } })(),
  })));
});

function makeDeleteAll(resource) {
  const def = DELETE_ALL_RESOURCES[resource];
  return (req, res) => {
    try {
      const db = getDb();
      // Snapshot first so the operator can always roll back. If this
      // fails (disk full, permission), bail before touching any data.
      let backupPath = '';
      try { backupPath = _snapshotBackupBeforeDelete(resource); }
      catch (e) {
        return res.status(500).json({ error: 'Backup snapshot failed; refusing to delete: ' + (e.message || e) });
      }
      const counts = _countDeleteAllImpact(db, resource);
      const before = counts[def.table] || 0;
      for (const t of def.cascade) {
        try { db.run(`DELETE FROM ${t}`); } catch (_) {}
        try { db.exec(`DELETE FROM sqlite_sequence WHERE name = '${t}'`); } catch (_) {}
      }
      db.run(`DELETE FROM ${def.table}`);
      try { db.exec(`DELETE FROM sqlite_sequence WHERE name = '${def.table}'`); } catch (_) {}
      _logDelete(db, { resource, deletedCount: before, cascadeCounts: counts, backupPath, req });
      res.json({ success: true, deleted: before, cascadeCounts: counts, backupPath });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  };
}
app.delete('/api/traders/delete-all',     requireDeleteAll, makeDeleteAll('traders'));
app.delete('/api/buyers/delete-all',      requireDeleteAll, makeDeleteAll('buyers'));
app.delete('/api/masters/delete-all',     requireDeleteAll, makeDeleteAll('masters'));
app.delete('/api/invoices/delete-all',    requireDeleteAll, makeDeleteAll('invoices'));
app.delete('/api/purchases/delete-all',   requireDeleteAll, makeDeleteAll('purchases'));
app.delete('/api/bills/delete-all',       requireDeleteAll, makeDeleteAll('bills'));
app.delete('/api/auctions/delete-all',    requireDeleteAll, makeDeleteAll('auctions'));
// Delete All Debit Notes — two modes:
//   • Trade-scoped:  ?ano=<trade-number> deletes only that trade's DNs
//     (preferred for the per-trade "Delete All" button in the DN tab)
//   • Global:        no ?ano= wipes EVERY debit note across all trades
//     (used by the Backup → Delete All admin page, like other resources)
// Both modes snapshot the DB first so a misclick is recoverable.
app.delete('/api/debit-notes/delete-all', requireDeleteAll, (req, res) => {
  try {
    const db = getDb();
    const ano = String(req.query.ano || '').trim();
    const scope = ano ? ('debit-notes-' + ano) : 'debit-notes';
    let backupPath = '';
    try { backupPath = _snapshotBackupBeforeDelete(scope); }
    catch (e) {
      return res.status(500).json({ error: 'Backup snapshot failed; refusing to delete: ' + (e.message || e) });
    }
    let before;
    if (ano) {
      before = db.get('SELECT COUNT(*) as c FROM debit_notes WHERE ano = ?', [ano]).c;
      db.run('DELETE FROM debit_notes WHERE ano = ?', [ano]);
    } else {
      before = db.get('SELECT COUNT(*) as c FROM debit_notes').c;
      db.run('DELETE FROM debit_notes');
      try { db.exec(`DELETE FROM sqlite_sequence WHERE name = 'debit_notes'`); } catch (_) {}
    }
    _logDelete(db, {
      resource: 'debit-notes',
      deletedCount: before,
      cascadeCounts: ano ? { ano, debit_notes: before } : { debit_notes: before },
      backupPath,
      req,
    });
    res.json({ success: true, deleted: before, ano: ano || null, backupPath });
  } catch (e) {
    res.status(500).json({ error: 'Delete All failed: ' + (e.message || e) });
  }
});

// ══════════════════════════════════════════════════════════════
// GST LOOKUP — fetch trade name/address/state from GSTIN
// Uses gstincheck.co.in if an API key is configured in settings
// (company-config: gst_api_key). Falls back to structural validation.
// ══════════════════════════════════════════════════════════════
const STATE_CODES = {
  '01':'JAMMU AND KASHMIR','02':'HIMACHAL PRADESH','03':'PUNJAB','04':'CHANDIGARH','05':'UTTARAKHAND',
  '06':'HARYANA','07':'DELHI','08':'RAJASTHAN','09':'UTTAR PRADESH','10':'BIHAR','11':'SIKKIM',
  '12':'ARUNACHAL PRADESH','13':'NAGALAND','14':'MANIPUR','15':'MIZORAM','16':'TRIPURA','17':'MEGHALAYA',
  '18':'ASSAM','19':'WEST BENGAL','20':'JHARKHAND','21':'ODISHA','22':'CHATTISGARH','23':'MADHYA PRADESH',
  '24':'GUJARAT','25':'DAMAN AND DIU','26':'DADRA AND NAGAR HAVELI','27':'MAHARASHTRA','28':'ANDHRA PRADESH',
  '29':'KARNATAKA','30':'GOA','31':'LAKSHADWEEP','32':'KERALA','33':'TAMIL NADU','34':'PUDUCHERRY',
  '35':'ANDAMAN AND NICOBAR ISLANDS','36':'TELANGANA','37':'ANDHRA PRADESH','38':'LADAKH'
};
const GSTIN_RE = /^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][1-9A-Z]Z[0-9A-Z]$/;

// Defaults for the low-credit warning thresholds. Operator can override
// via settings keys `gst_warn_below` and `gst_critical_below` once the
// recharge habits stabilise.
const GST_WARN_BELOW_DEFAULT = 50;
const GST_CRITICAL_BELOW_DEFAULT = 10;

// gstincheck.co.in ships credit-related fields under several names
// depending on plan and endpoint version. Probe each common one and
// return the first non-empty match. Adding new aliases here is safe —
// the probe is pure, just data lookups.
function _gstPickField(obj, names) {
  if (!obj || typeof obj !== 'object') return null;
  for (const n of names) {
    if (obj[n] != null && obj[n] !== '') return obj[n];
  }
  return null;
}
// Some gstincheck.co.in plans report quota via response headers, not
// the JSON body. Pull headers into a plain {lowercase-name → value}
// object so the probe can search them alongside the body.
function _headersToObj(h) {
  const out = {};
  if (!h) return out;
  try {
    // Node fetch returns a Headers object (iterable). Older shims
    // surface a plain object — handle both.
    if (typeof h.forEach === 'function') {
      h.forEach((v, k) => { out[String(k).toLowerCase()] = v; });
    } else if (typeof h === 'object') {
      for (const k of Object.keys(h)) out[String(k).toLowerCase()] = h[k];
    }
  } catch (_) {}
  return out;
}
// Scan a "message" string for inline credit counts. gstincheck.co.in
// returns sentences like "847 Searches Left" or "Total 1000 Credits"
// depending on the package; capture whichever pattern matches first.
function _parseCreditsFromMessage(msg) {
  const s = String(msg || '');
  if (!s) return {};
  const out = {};
  // "847 searches left", "847 credits remaining", "847 hits remaining"
  let m = s.match(/(\d[\d,]*)\s*(?:searches?|credits?|hits?|requests?)\s*(?:left|remaining|available|balance)\b/i);
  if (m) out.remaining = Number(m[1].replace(/,/g, ''));
  // "balance: 847" / "left: 847"
  if (out.remaining == null) {
    m = s.match(/(?:balance|left|remaining|available)\s*[:=]\s*(\d[\d,]*)/i);
    if (m) out.remaining = Number(m[1].replace(/,/g, ''));
  }
  // "of 1000" / "total 1000"
  m = s.match(/(?:of|total|out\s+of)\s*(\d[\d,]*)/i);
  if (m) out.total = Number(m[1].replace(/,/g, ''));
  // "valid till 2026-12-31" / "expires on 2026-12-31"
  m = s.match(/(?:valid|expir(?:es|y))[^0-9]*?(\d{4}-\d{2}-\d{2})/i);
  if (m) out.expires = m[1];
  return out;
}
function _gstExtractCredits(rawBody, rawHeaders) {
  // Search the body at the top level AND inside common metadata
  // wrappers, the response headers, AND the human-readable message
  // field. Adding new aliases here is safe — every probe is pure
  // data lookup or regex, no side effects.
  const candidates = [];
  if (rawBody && typeof rawBody === 'object') {
    candidates.push(rawBody);
    for (const k of ['userInfo', 'quota', 'meta', 'pkg', 'plan', 'subscription', 'account']) {
      if (rawBody[k] && typeof rawBody[k] === 'object') candidates.push(rawBody[k]);
    }
  }
  if (rawHeaders && typeof rawHeaders === 'object') candidates.push(rawHeaders);
  let remaining = null, total = null, expires = null;
  for (const c of candidates) {
    remaining = remaining ?? _gstPickField(c, [
      // body field names
      'gstinAvailableSearch', 'availableSearch', 'remainingHits',
      'remainingSearches', 'creditsLeft', 'credits_remaining', 'remaining',
      'balance', 'searchLeft', 'searches_left', 'available_credits',
      'creditBalance', 'credit_balance', 'apiCallsRemaining',
      // header names (lowercased by _headersToObj)
      'x-credits-remaining', 'x-credit-remaining', 'x-credits-left',
      'x-searches-remaining', 'x-quota-remaining', 'x-ratelimit-remaining',
    ]);
    total = total ?? _gstPickField(c, [
      'gstinTotalSearch', 'totalSearch', 'totalHits',
      'totalSearches', 'creditsTotal', 'credits_total', 'total',
      'totalCredits', 'total_credits', 'plan_size', 'planSize',
      'x-credits-total', 'x-quota-limit', 'x-ratelimit-limit',
    ]);
    expires = expires ?? _gstPickField(c, [
      'gstinValidity', 'searchValidUpto', 'validUpto', 'validity',
      'planExpiresAt', 'plan_expires_at', 'expiryDate', 'expiry_date',
      'subscriptionEnd', 'subscription_end', 'expiresOn', 'expires_on',
      'x-plan-expires', 'x-subscription-expires',
    ]);
  }
  // Last-resort: parse the human-readable `message` field.
  if (remaining == null || total == null || expires == null) {
    const msgScan = _parseCreditsFromMessage(rawBody && (rawBody.message || rawBody.msg || rawBody.note));
    if (remaining == null) remaining = msgScan.remaining != null ? msgScan.remaining : null;
    if (total     == null) total     = msgScan.total     != null ? msgScan.total     : null;
    if (expires   == null) expires   = msgScan.expires   != null ? msgScan.expires   : null;
  }
  return {
    remaining: remaining == null ? null : Number(remaining),
    total:     total     == null ? null : Number(total),
    expires:   expires   == null ? null : String(expires),
  };
}
function _gstSaveState(db, rawBody, rawHeaders) {
  const headerObj = _headersToObj(rawHeaders);
  const credits = _gstExtractCredits(rawBody, headerObj);
  // Persist a trimmed envelope so the operator can audit what the API
  // returned without keeping the (potentially large) `data` blob.
  // Includes both the body (minus `data`) and any headers that LOOK
  // credit-related, plus the `_extracted` snapshot so the UI can show
  // exactly what the probe matched (or didn't).
  const meta = { _body: {}, _headers: {}, _extracted: credits };
  if (rawBody && typeof rawBody === 'object') {
    for (const k of Object.keys(rawBody)) {
      if (k === 'data') continue;
      meta._body[k] = rawBody[k];
    }
  }
  // Save only header names that look credit/quota-related so we don't
  // bloat the row with cookies / CDN noise. Anything matching common
  // prefixes (x-credit, x-search, x-quota, x-ratelimit, x-plan) stays.
  for (const k of Object.keys(headerObj)) {
    if (/^(x-credit|x-search|x-quota|x-ratelimit|x-plan|x-subscription|x-balance)/.test(k)) {
      meta._headers[k] = headerObj[k];
    }
  }
  const now = new Date().toISOString();
  // Ensure the single row exists then UPDATE — simpler than UPSERT
  // and avoids any quirks with sql.js's INSERT OR REPLACE on a
  // CHECK-constrained id column.
  const exists = db.get('SELECT id FROM gst_api_state WHERE id = 1');
  if (!exists) {
    db.run(
      `INSERT INTO gst_api_state
        (id, credits_remaining, credits_total, plan_expires_at, last_checked_at, last_response_raw)
       VALUES (1, ?, ?, ?, ?, ?)`,
      [credits.remaining, credits.total, credits.expires, now, JSON.stringify(meta)]
    );
  } else {
    // Only overwrite fields we actually observed — keep prior data on
    // misses so a partial response doesn't blank the cached value.
    const cur = db.get('SELECT * FROM gst_api_state WHERE id = 1') || {};
    db.run(
      `UPDATE gst_api_state SET
         credits_remaining = ?,
         credits_total     = ?,
         plan_expires_at   = ?,
         last_checked_at   = ?,
         last_response_raw = ?
       WHERE id = 1`,
      [
        credits.remaining ?? cur.credits_remaining ?? null,
        credits.total     ?? cur.credits_total     ?? null,
        credits.expires   ?? cur.plan_expires_at   ?? null,
        now,
        JSON.stringify(meta),
      ]
    );
  }
  return credits;
}
function _gstStatusFor(db, cfg) {
  const row = db.get('SELECT * FROM gst_api_state WHERE id = 1') || {};
  const cfgg = cfg || getSettingsFlat(db);
  const warnBelow     = Number(cfgg.gst_warn_below)     || GST_WARN_BELOW_DEFAULT;
  const criticalBelow = Number(cfgg.gst_critical_below) || GST_CRITICAL_BELOW_DEFAULT;
  const remaining = row.credits_remaining == null ? null : Number(row.credits_remaining);
  let level = 'unknown';
  if (remaining == null)              level = 'unknown';
  else if (remaining <= 0)            level = 'exhausted';
  else if (remaining <  criticalBelow) level = 'critical';
  else if (remaining <  warnBelow)     level = 'warning';
  else                                 level = 'ok';
  // Parse the persisted envelope back so callers can see WHAT the
  // API returned (body keys + saved headers + what the probe matched).
  // Useful when level === 'unknown' and we need to teach the probe
  // about a new field name without re-running the lookup.
  let lastEnvelope = null;
  if (row.last_response_raw) {
    try { lastEnvelope = JSON.parse(row.last_response_raw); } catch (_) { lastEnvelope = null; }
  }
  return {
    has_api_key:        !!(cfgg.gst_api_key && String(cfgg.gst_api_key).trim()),
    credits_remaining:  remaining,
    credits_total:      row.credits_total == null ? null : Number(row.credits_total),
    plan_expires_at:    row.plan_expires_at || null,
    last_checked_at:    row.last_checked_at || null,
    warn_below:         warnBelow,
    critical_below:     criticalBelow,
    level,            // 'ok' | 'warning' | 'critical' | 'exhausted' | 'unknown'
    recharge_url:       'https://gstincheck.co.in/',
    last_envelope:      lastEnvelope,
  };
}

// ── Cheap status probe (registered BEFORE the dynamic :gstin route) ──
// Express matches routes in registration order, so the literal path
// /api/gst-lookup/status MUST be declared first — otherwise the
// :gstin handler below grabs the word "status", fails the GSTIN
// regex, and returns 400 "Invalid GSTIN format" (the bug you hit).
//
// Used by the Settings → Integrations card and the topbar pill to
// render the credit count without burning a real lookup. Numbers
// reflect the most recent observation; do a real GSTIN lookup to
// force a refresh.
app.get('/api/gst-lookup/status', requireView, (req, res) => {
  res.set('Cache-Control', 'no-store');
  try {
    const status = _gstStatusFor(getDb(), null);
    res.json(status);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.get('/api/gst-lookup/:gstin', requireView, async (req, res) => {
  const gstin = String(req.params.gstin || '').toUpperCase().trim();
  if (!GSTIN_RE.test(gstin)) {
    return res.status(400).json({ valid: false, error: 'Invalid GSTIN format' });
  }
  const stCode = gstin.substring(0, 2);
  const pan    = gstin.substring(2, 12);
  const state  = STATE_CODES[stCode] || '';

  const cfg = getSettingsFlat(getDb());
  const apiKey = cfg.gst_api_key || '';

  // No API key → return structural details only
  if (!apiKey) {
    return res.json({
      valid: true, gstin, pan, st_code: stCode, state,
      source: 'structural',
      note: 'Set "gst_api_key" in settings to auto-fetch trade name/address.'
    });
  }

  // With API key → attempt live lookup. Every successful response
  // opportunistically refreshes gst_api_state so the Settings card +
  // topbar pill stay current without a separate "ping" cost.
  try {
    const url = `https://sheet.gstincheck.co.in/check/${apiKey}/${gstin}`;
    const r = await fetch(url);
    if (!r.ok) throw new Error(`HTTP ${r.status}`);
    const body = await r.json();
    // Persist credit info regardless of whether `data` came back —
    // even a "no records found" response usually still carries the
    // quota envelope, and we want that visible to the operator. Pass
    // both body AND response headers so the probe can pick up plans
    // that report credits via X-Credits-Remaining etc.
    const db = getDb();
    try { _gstSaveState(db, body, r.headers); } catch (_) { /* persistence is best-effort */ }
    const apiStatus = _gstStatusFor(db, cfg);
    if (body && body.flag && body.data) {
      const d = body.data;
      const addr = (d.pradr && d.pradr.addr) || {};
      return res.json({
        valid: true, gstin, pan, st_code: stCode,
        name:     d.lgnm || d.tradeNam || '',
        tradeName:d.tradeNam || d.lgnm || '',
        address:  [addr.bno, addr.bnm, addr.st, addr.loc].filter(Boolean).join(', '),
        place:    addr.dst || addr.loc || '',
        pin:      addr.pncd || '',
        state:    addr.stcd || state,
        status:   d.sts || '',
        regDate:  d.rgdt || '',
        source:   'live',
        // Surface the freshly-refreshed credit state so the UI can
        // show the "X searches left" hint on each lookup. Carried in
        // a nested `api` object to avoid colliding with the existing
        // `status` (GST registration status) field.
        api:      apiStatus,
      });
    }
    return res.json({
      valid: true, gstin, pan, st_code: stCode, state,
      source: 'structural',
      note: body && body.message ? body.message : 'GST portal returned no data',
      api:  apiStatus,
    });
  } catch (e) {
    return res.json({
      valid: true, gstin, pan, st_code: stCode, state,
      source: 'structural',
      note: 'GST lookup failed: ' + e.message
    });
  }
});

// ══════════════════════════════════════════════════════════════
// TRADERS (NAM.DBF — sellers/poolers)
// ══════════════════════════════════════════════════════════════
app.get('/api/traders', requireViewOrLotEntry, (req, res) => {
  const { search, limit } = req.query;
  const db = getDb();
  // Helper: attach the `banks` array to each trader row (from trader_banks
  // table). Kept as a post-query hydration step so we don't bloat the main
  // query with joins or GROUP_CONCAT — easier to read, and the N+1 is fine
  // for the small trader counts this app handles (~few hundred max).
  const hydrateBanks = (rows) => {
    if (!rows.length) return rows;
    const ids = rows.map(r => r.id);
    const placeholders = ids.map(() => '?').join(',');
    const banks = db.all(
      `SELECT id, trader_id, bank_name, branch, acctnum, ifsc, holder_name, is_default
       FROM trader_banks WHERE trader_id IN (${placeholders})
       ORDER BY trader_id, is_default DESC, id`, ids
    );
    const byTrader = new Map();
    for (const b of banks) {
      if (!byTrader.has(b.trader_id)) byTrader.set(b.trader_id, []);
      byTrader.get(b.trader_id).push(b);
    }
    for (const r of rows) r.banks = byTrader.get(r.id) || [];
    return rows;
  };

  // ── Pagination ──────────────────────────────────────────────
  // `?page=` (1-based) + `?pageSize=` cap page-window size.
  //   • `?page=` or `?pageSize=` present  → paged response
  //         { rows: [...], total: N, page: P, pageSize: S }
  //   • Neither present                   → legacy plain-array response,
  //         capped at 500 rows (back-compat for old UI / dropdowns)
  //   • Legacy `?limit=` is preserved as an alias for `pageSize` so the
  //         lot-entry seller picker (which sends `&limit=100`) keeps
  //         working unchanged.
  // Search filters apply BEFORE paging, so search hits the whole table —
  // critical on 7,000-row masters where the legacy LIMIT 500 was hiding
  // matches past row 500.
  const page     = Math.max(1, parseInt(req.query.page, 10) || 1);
  const pageSize = Math.min(500, Math.max(1, parseInt(req.query.pageSize || req.query.limit, 10) || 50));
  const offset   = (page - 1) * pageSize;
  const wantPaged = req.query.page != null || req.query.pageSize != null;

  // Build WHERE clause once — shared by COUNT and page query so the
  // total reported in the pager matches what'll appear when the user
  // pages through.
  let where = '';
  let params = [];
  if (search) {
    const q = `%${search}%`;
    where = 'WHERE name LIKE ? OR tel LIKE ? OR cr LIKE ? OR pan LIKE ? OR ppla LIKE ? OR aadhar LIKE ?';
    params = [q, q, q, q, q, q];
  }

  // Dedupe-by-(name,CR) for legacy responses is intentionally NOT applied
  // to paged responses: pagination needs stable row counts that match
  // total, and the LIMIT/OFFSET windowing won't reach all rows reliably
  // if we collapse duplicates client-side. Fixing legacy dupes is a
  // separate data-cleanup task — see _maybeDedupeTradersByNameCR below.
  if (wantPaged) {
    const total = db.get(`SELECT COUNT(*) as c FROM traders ${where}`, params).c;
    const rows = db.all(
      `SELECT * FROM traders ${where} ORDER BY name LIMIT ? OFFSET ?`,
      [...params, pageSize, offset]
    );
    return res.json({ rows: hydrateBanks(rows), total, page, pageSize });
  }

  // ── Legacy (un-paged) path — preserved for back-compat ──────
  // For search: returns up to `pageSize` matches (default 50) but the
  // dedupe-by-(name,CR) collapse still runs, matching the previous
  // behaviour used by the lot-entry seller picker.
  if (search) {
    const rows = db.all(
      `SELECT * FROM traders ${where} ORDER BY name LIMIT ?`,
      [...params, pageSize]
    );
    const normalize = (s) => String(s || '').trim().toUpperCase();
    const stripGstinPrefix = (s) => {
      let v = normalize(s);
      if (v.startsWith('GSTIN.')) v = v.slice(6);
      else if (v.startsWith('GSTIN')) v = v.slice(5);
      return v;
    };
    const seen = new Map();
    for (const r of rows) {
      const key = normalize(r.name) + '|' + stripGstinPrefix(r.cr);
      const prev = seen.get(key);
      if (!prev || (r.id || 0) > (prev.id || 0)) seen.set(key, r);
    }
    return res.json(hydrateBanks([...seen.values()]));
  }
  res.json(hydrateBanks(db.all('SELECT * FROM traders ORDER BY name LIMIT 500')));
});
// Lookup a trader (seller) by exact name — used by the WhatsApp "send"
// flow on Purchase / Payments rows where we have the seller name and
// need their phone number. Case-insensitive match. Returns the first hit.
app.get('/api/traders/by-name/:name', requireViewOrLotEntry, (req, res) => {
  const db = getDb();
  const nm = String(req.params.name || '').trim();
  if (!nm) return res.status(400).json({ error: 'name required' });
  const row = db.get('SELECT id, name, tel FROM traders WHERE LOWER(name) = LOWER(?) LIMIT 1', [nm]);
  if (!row) return res.status(404).json({ error: 'Not found' });
  res.json(row);
});

// ── WhatsApp Cloud API (Meta) ─────────────────────────────────
// Production send path: every Cloud message goes through a Meta-APPROVED
// TEMPLATE, so we can message anyone — including contacts who have never
// messaged us (no 24h customer-service-window restriction). PDFs ride
// inside a document-header template. When credentials are unset the send
// endpoints return 501 so the client falls back to its manual
// (web.whatsapp.com link / Web Share API) flow without erroring.
//
// Credentials resolve ENV-FIRST, then the DB (whatsapp_config table, set
// via the in-app Settings card). This lets Railway use dashboard vars and
// the desktop/in-app user paste creds without touching the environment.
//
// Setup: https://developers.facebook.com/docs/whatsapp/cloud-api
//   1. Meta Business → WhatsApp → API Setup: capture Phone number ID + WABA ID.
//   2. Generate a PERMANENT system-user token (scopes whatsapp_business_messaging
//      + whatsapp_business_management). App secret from App → Settings → Basic.
//   3. Create + get approval for two templates (see /api/whatsapp/config card):
//      a DOCUMENT template (header=Document) for PDF sends, and a TEXT template.
//      Both use body "Dear {{1}}, {{2}} Regards, {{3}}." → name, summary, company.
//   4. Webhook: set Callback URL to <host>/api/whatsapp/webhook + a verify token,
//      subscribe to the "messages" field. Delivery receipts then flow back in.
//   5. Set the values via env (below) or the in-app Settings → Integrations card.
//      Env keys: WHATSAPP_TOKEN, WHATSAPP_PHONE_ID, WHATSAPP_WABA_ID,
//      WHATSAPP_APP_SECRET, WHATSAPP_VERIFY_TOKEN, WHATSAPP_TPL_DOCUMENT,
//      WHATSAPP_TPL_TEXT.

// Resolve effective config: process.env wins over the DB row. Returns a
// flat object plus `configured` (token + phoneId both present) and
// `source` ('env' | 'db' | 'none') indicating where the live creds came from.
function _waConfig(db) {
  let row = {};
  try { row = db.get('SELECT * FROM whatsapp_config WHERE id = 1') || {}; } catch (_) { row = {}; }
  const envv = (k) => (process.env[k] && String(process.env[k]).trim()) || '';
  const dbv = (c) => (row[c] != null && String(row[c]).trim()) || '';
  const pick = (k, c) => envv(k) || dbv(c);
  const token = pick('WHATSAPP_TOKEN', 'access_token');
  const phoneId = pick('WHATSAPP_PHONE_ID', 'phone_id');
  let source = 'none';
  if (token && phoneId) source = (envv('WHATSAPP_TOKEN') && envv('WHATSAPP_PHONE_ID')) ? 'env' : 'db';
  return {
    token, phoneId,
    wabaId:          pick('WHATSAPP_WABA_ID', 'waba_id'),
    appSecret:       pick('WHATSAPP_APP_SECRET', 'app_secret'),
    verifyToken:     pick('WHATSAPP_VERIFY_TOKEN', 'verify_token'),
    displayNumber:   dbv('display_number'),
    tplDocument:     pick('WHATSAPP_TPL_DOCUMENT', 'tpl_document'),
    tplDocumentLang: dbv('tpl_document_lang') || 'en',
    tplText:         pick('WHATSAPP_TPL_TEXT', 'tpl_text'),
    tplTextLang:     dbv('tpl_text_lang') || 'en',
    configured: Boolean(token && phoneId),
    source,
  };
}
function _waNormalizePhone(tel) {
  const d = String(tel || '').replace(/\D/g, '');
  if (!d) return '';
  return d.length === 10 ? '91' + d : d; // default IN country code
}
// Parse positional template body params from a request. JSON bodies send an
// array directly; multipart sends a JSON-encoded string (or a bare string).
function _waParseParams(raw) {
  if (Array.isArray(raw)) return raw;
  if (raw == null || raw === '') return [];
  if (typeof raw === 'string') {
    const s = raw.trim();
    if (s.startsWith('[')) { try { const a = JSON.parse(s); return Array.isArray(a) ? a : [s]; } catch (_) { return [s]; } }
    return [s];
  }
  return [];
}
// POST to the Graph API messages-or-other endpoint with the resolved creds.
async function _waGraphPost(cfg, path, body) {
  const url = `https://graph.facebook.com/v18.0/${cfg.phoneId}${path}`;
  const r = await fetch(url, {
    method: 'POST',
    headers: { 'Authorization': 'Bearer ' + cfg.token, 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });
  const j = await r.json().catch(() => ({}));
  if (!r.ok) throw new Error(j.error && j.error.message ? j.error.message : `Cloud API error ${r.status}`);
  return j;
}
// Upload a PDF buffer to Meta /media → returns a media_id usable for ~30 days.
async function _waUploadMedia(cfg, buffer, filename) {
  const fd = new FormData();
  fd.append('messaging_product', 'whatsapp');
  fd.append('type', 'application/pdf');
  fd.append('file', new Blob([buffer], { type: 'application/pdf' }), filename || 'document.pdf');
  const r = await fetch(`https://graph.facebook.com/v18.0/${cfg.phoneId}/media`, {
    method: 'POST', headers: { 'Authorization': 'Bearer ' + cfg.token }, body: fd,
  });
  const j = await r.json().catch(() => ({}));
  if (!r.ok || !j.id) throw new Error((j.error && j.error.message) || `Media upload failed ${r.status}`);
  return j.id;
}
// Send an approved template. Includes the document header component only
// when a media id is supplied; the body component carries positional params.
async function _waSendTemplate(cfg, { phone, template, lang, bodyParams = [], documentMediaId = null, filename = 'document.pdf' }) {
  const components = [];
  if (documentMediaId) {
    components.push({ type: 'header', parameters: [{ type: 'document', document: { id: documentMediaId, filename } }] });
  }
  if (bodyParams.length) {
    // Meta rejects template body params containing newlines, tabs, or 5+
    // consecutive spaces. Flatten multi-line summaries (e.g. lot receipts)
    // into a single line so any caller's text is accepted.
    const clean = (t) => String(t == null ? '' : t)
      .replace(/\r?\n+/g, ' · ')
      .replace(/\t+/g, ' ')
      .replace(/ {2,}/g, ' ')
      .trim();
    components.push({ type: 'body', parameters: bodyParams.map((t) => ({ type: 'text', text: clean(t) })) });
  }
  const tpl = { name: template, language: { code: lang || 'en' } };
  if (components.length) tpl.components = components;
  const out = await _waGraphPost(cfg, '/messages', {
    messaging_product: 'whatsapp', to: phone, type: 'template', template: tpl,
  });
  return out.messages && out.messages[0] && out.messages[0].id;
}
// Append a row to the send log. Best-effort — never throws into a handler.
function _waLog(db, f) {
  try {
    db.run(
      `INSERT INTO whatsapp_messages (wamid, direction, phone, msg_type, caption, status, error, ref_type, ref_id)
       VALUES (?,?,?,?,?,?,?,?,?)`,
      [f.wamid || '', f.direction || 'out', f.phone || '', f.msg_type || '',
       f.caption || '', f.status || 'queued', f.error || '', f.ref_type || '', f.ref_id || '']
    );
  } catch (_) { /* logging must not break sends */ }
}

// ── Status / config / diagnostics ─────────────────────────────
// Connection status for the client capability probe + Settings card. NEVER
// returns the token/app-secret/verify-token — only booleans + non-secret
// identifiers. When configured, best-effort pings Graph for live health.
app.get('/api/whatsapp/status', requireView, async (req, res) => {
  const cfg = _waConfig(getDb());
  const result = {
    configured: cfg.configured,
    source: cfg.source,
    phoneId: cfg.phoneId,
    wabaId: cfg.wabaId,
    displayNumber: cfg.displayNumber,
    hasToken: !!cfg.token,
    hasAppSecret: !!cfg.appSecret,
    hasVerifyToken: !!cfg.verifyToken,
    tplDocument: cfg.tplDocument,
    tplDocumentLang: cfg.tplDocumentLang,
    tplText: cfg.tplText,
    tplTextLang: cfg.tplTextLang,
    webhookReady: !!(cfg.verifyToken && cfg.appSecret),
    qualityRating: null,
    displayPhone: null,
    verifiedName: null,
    live: false,
    liveError: null,
  };
  if (cfg.configured) {
    try {
      const ctrl = new AbortController();
      const to = setTimeout(() => ctrl.abort(), 4000);
      const r = await fetch(
        `https://graph.facebook.com/v18.0/${cfg.phoneId}?fields=display_phone_number,quality_rating,verified_name`,
        { headers: { 'Authorization': 'Bearer ' + cfg.token }, signal: ctrl.signal }
      );
      clearTimeout(to);
      const j = await r.json().catch(() => ({}));
      if (r.ok) {
        result.qualityRating = j.quality_rating || null;
        result.displayPhone = j.display_phone_number || null;
        result.verifiedName = j.verified_name || null;
        result.live = true;
      } else {
        result.liveError = (j.error && j.error.message) || `HTTP ${r.status}`;
      }
    } catch (e) {
      result.liveError = e.name === 'AbortError' ? 'Timed out contacting Meta' : e.message;
    }
  }
  res.json(result);
});

// Upsert WhatsApp credentials + template config. Secrets (token, app
// secret, verify token) are write-only: a blank/absent value leaves the
// stored secret UNCHANGED, so saving other fields never wipes them. The
// client can resend non-secret fields (it gets them back from /status).
app.put('/api/whatsapp/config', requireSettingsWrite, (req, res) => {
  const db = getDb();
  const b = req.body || {};
  const sets = [];
  const vals = [];
  const put = (col, val) => { sets.push(`${col} = ?`); vals.push(val); };
  const putSecret = (col, val) => { if (typeof val === 'string' && val.trim()) put(col, val.trim()); };
  const putField = (col, val, dflt) => { if (val !== undefined) put(col, String(val).trim() || (dflt || '')); };
  putSecret('access_token', b.token);
  putSecret('app_secret', b.appSecret);
  putSecret('verify_token', b.verifyToken);
  putField('phone_id', b.phoneId);
  putField('waba_id', b.wabaId);
  putField('display_number', b.displayNumber);
  putField('tpl_document', b.tplDocument);
  putField('tpl_document_lang', b.tplDocumentLang, 'en');
  putField('tpl_text', b.tplText);
  putField('tpl_text_lang', b.tplTextLang, 'en');
  if (b.enabled !== undefined) put('enabled', b.enabled ? 1 : 0);
  if (!sets.length) return res.json({ ok: true, updated: 0 });
  try {
    db.run(`UPDATE whatsapp_config SET ${sets.join(', ')}, updated_at = datetime('now','localtime') WHERE id = 1`, vals);
  } catch (e) { return res.status(500).json({ error: e.message }); }
  res.json({ ok: true, updated: sets.length });
});

// Fire a test send (text template) to verify the whole pipeline. Surfaces
// the exact Meta error so the user can self-diagnose from the Settings card.
app.post('/api/whatsapp/test', requireSettingsWrite, async (req, res) => {
  const db = getDb();
  const cfg = _waConfig(db);
  if (!cfg.configured) return res.status(400).json({ error: 'Token and Phone number ID are required first.' });
  if (!cfg.tplText) return res.status(400).json({ error: 'Set an approved Text template name first.' });
  const phone = _waNormalizePhone(req.body.phone);
  if (!phone) return res.status(400).json({ error: 'Recipient phone required.' });
  const params = _waParseParams(req.body.params);
  const bodyParams = params.length ? params
    : ['Test', 'this is a test message confirming your WhatsApp Cloud API setup works.', cfg.verifiedName || 'Spice e-Trade'];
  try {
    const wamid = await _waSendTemplate(cfg, { phone, template: cfg.tplText, lang: cfg.tplTextLang, bodyParams });
    _waLog(db, { wamid, phone, msg_type: 'template-text', caption: '[test]', status: 'sent', ref_type: 'test' });
    res.json({ ok: true, id: wamid });
  } catch (e) {
    _waLog(db, { phone, msg_type: 'template-text', caption: '[test]', status: 'failed', error: e.message, ref_type: 'test' });
    res.status(502).json({ error: e.message });
  }
});

// ── Primary send paths (always template) ──────────────────────
// Text-only template send. Body: { phone, params:[…], ref_type?, ref_id? }.
app.post('/api/whatsapp/send-template-text', requireView, async (req, res) => {
  const db = getDb();
  const cfg = _waConfig(db);
  if (!cfg.configured) return res.status(501).json({ error: 'WhatsApp Cloud API not configured', fallback: true });
  if (!cfg.tplText) return res.status(400).json({ error: 'No text template configured', fallback: true });
  const phone = _waNormalizePhone(req.body.phone);
  if (!phone) return res.status(400).json({ error: 'phone required' });
  const bodyParams = _waParseParams(req.body.params);
  try {
    const wamid = await _waSendTemplate(cfg, { phone, template: cfg.tplText, lang: cfg.tplTextLang, bodyParams });
    _waLog(db, { wamid, phone, msg_type: 'template-text', caption: bodyParams[1] || '', status: 'sent',
      ref_type: req.body.ref_type || '', ref_id: req.body.ref_id || '' });
    res.json({ ok: true, id: wamid });
  } catch (e) {
    _waLog(db, { phone, msg_type: 'template-text', caption: bodyParams[1] || '', status: 'failed', error: e.message,
      ref_type: req.body.ref_type || '', ref_id: req.body.ref_id || '' });
    res.status(502).json({ error: e.message });
  }
});

// Document template send — delivers a locally-generated PDF to ANYONE
// (cold) via an approved document-header template. Multipart body:
//   file     — the PDF blob (required)
//   phone    — recipient (required; 10-digit IN auto-prefixed with 91)
//   params   — JSON array of positional body params (name, summary, company)
//   filename — display name for the attachment (optional)
//   ref_type / ref_id — optional traceability back to the source record
// Same 501/502 contract so the client falls back to the manual flow.
app.post('/api/whatsapp/send-template-document', requireView, upload.single('file'), async (req, res) => {
  const db = getDb();
  const cfg = _waConfig(db);
  const cleanup = () => { if (req.file) fs.unlink(req.file.path, () => {}); };
  if (!cfg.configured) { cleanup(); return res.status(501).json({ error: 'WhatsApp Cloud API not configured', fallback: true }); }
  if (!cfg.tplDocument) { cleanup(); return res.status(400).json({ error: 'No document template configured', fallback: true }); }
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
  const phone = _waNormalizePhone(req.body.phone);
  const filename = String(req.body.filename || req.file.originalname || 'document.pdf');
  const bodyParams = _waParseParams(req.body.params);
  if (!phone) { cleanup(); return res.status(400).json({ error: 'phone required' }); }
  try {
    const buffer = fs.readFileSync(req.file.path);
    const mediaId = await _waUploadMedia(cfg, buffer, filename);
    const wamid = await _waSendTemplate(cfg, {
      phone, template: cfg.tplDocument, lang: cfg.tplDocumentLang,
      bodyParams, documentMediaId: mediaId, filename,
    });
    _waLog(db, { wamid, phone, msg_type: 'template-document', caption: bodyParams[1] || filename, status: 'sent',
      ref_type: req.body.ref_type || '', ref_id: req.body.ref_id || '' });
    res.json({ ok: true, id: wamid, mediaId });
  } catch (e) {
    _waLog(db, { phone, msg_type: 'template-document', caption: bodyParams[1] || filename, status: 'failed', error: e.message,
      ref_type: req.body.ref_type || '', ref_id: req.body.ref_id || '' });
    res.status(502).json({ error: e.message });
  } finally {
    cleanup();
  }
});

// ── Send log ───────────────────────────────────────────────────
app.get('/api/whatsapp/messages', requireView, (req, res) => {
  const db = getDb();
  const rows = db.all(
    `SELECT id, wamid, direction, phone, msg_type, caption, status, error, ref_type, ref_id, created_at, updated_at
     FROM whatsapp_messages ORDER BY id DESC LIMIT 100`
  );
  res.json(rows);
});

// ── Webhook (PUBLIC — Meta calls these unauthenticated) ────────
// GET: verification handshake. Echo hub.challenge when the verify token matches.
app.get('/api/whatsapp/webhook', (req, res) => {
  const cfg = _waConfig(getDb());
  const mode = req.query['hub.mode'];
  const token = req.query['hub.verify_token'];
  const challenge = req.query['hub.challenge'];
  if (mode === 'subscribe' && cfg.verifyToken && token === cfg.verifyToken) {
    return res.status(200).send(String(challenge == null ? '' : challenge));
  }
  return res.sendStatus(403);
});
// POST: delivery-status + inbound-message events. Verify Meta's HMAC
// signature over the RAW body (captured by the express.json verify hook),
// then update the send log / record inbound replies. Always 200 fast —
// Meta retries aggressively on any non-2xx.
app.post('/api/whatsapp/webhook', (req, res) => {
  const db = getDb();
  const cfg = _waConfig(db);
  if (cfg.appSecret) {
    const sig = req.get('X-Hub-Signature-256') || '';
    const expected = 'sha256=' + crypto.createHmac('sha256', cfg.appSecret).update(req.rawBody || Buffer.from('')).digest('hex');
    const a = Buffer.from(sig), b = Buffer.from(expected);
    if (a.length !== b.length || !crypto.timingSafeEqual(a, b)) return res.sendStatus(401);
  }
  try {
    for (const entry of (req.body && req.body.entry) || []) {
      for (const change of (entry.changes || [])) {
        const v = change.value || {};
        for (const st of (v.statuses || [])) {
          const wamid = st.id || '';
          const status = st.status || '';
          const err = (st.errors && st.errors[0] && (st.errors[0].title || st.errors[0].message)) || '';
          if (wamid && status) {
            try {
              db.run("UPDATE whatsapp_messages SET status = ?, error = ?, updated_at = datetime('now','localtime') WHERE wamid = ?",
                [status, err, wamid]);
            } catch (_) {}
          }
        }
        for (const m of (v.messages || [])) {
          const body = (m.text && m.text.body) || m.type || '';
          try { db.run('INSERT INTO whatsapp_inbound (wamid, phone, body) VALUES (?,?,?)', [m.id || '', m.from || '', String(body)]); } catch (_) {}
        }
      }
    }
  } catch (_) { /* never let a malformed payload 500 — Meta would retry forever */ }
  res.sendStatus(200);
});

app.get('/api/traders/:id', requireViewOrLotEntry, (req, res) => {
  const db = getDb();
  const row = db.get('SELECT * FROM traders WHERE id = ?', [req.params.id]);
  if (!row) return res.status(404).json({ error: 'Not found' });
  // Attach banks array so the edit modal sees all bank accounts.
  row.banks = db.all(
    'SELECT id, trader_id, bank_name, branch, acctnum, ifsc, holder_name, is_default FROM trader_banks WHERE trader_id = ? ORDER BY is_default DESC, id',
    [row.id]
  );
  res.json(row);
});
// Sync a trader's banks array into the trader_banks table.
// Strategy: clear existing rows for this trader, reinsert. Simple and
// correct; the number of banks per trader is tiny (typically 1-3) so
// the delete+reinsert cost is negligible.
// Also mirrors the FIRST bank back into the parent traders.ifsc/acctnum/
// holder_name columns so older code paths that haven't been migrated to
// read trader_banks yet still see a valid primary account.
function syncTraderBanks(db, traderId, banks) {
  const arr = Array.isArray(banks) ? banks.filter(b => b && (b.acctnum || b.ifsc)) : [];
  db.run('DELETE FROM trader_banks WHERE trader_id = ?', [traderId]);
  for (const b of arr) {
    db.run(
      'INSERT INTO trader_banks (trader_id, bank_name, branch, acctnum, ifsc, holder_name) VALUES (?,?,?,?,?,?)',
      [traderId, b.bank_name||'', b.branch||'', String(b.acctnum||''), String(b.ifsc||''), b.holder_name||'']
    );
  }
  // Mirror first bank into traders row for legacy compatibility
  const first = arr[0] || {};
  db.run(
    'UPDATE traders SET ifsc=?, acctnum=?, holder_name=? WHERE id=?',
    [first.ifsc||'', first.acctnum||'', first.holder_name||'', traderId]
  );
}

// Duplicate-PAN guard for sellers. PAN is the legally-unique tax id
// on a trader so two rows sharing one is almost always an accidental
// re-entry of the same person. The check is case-insensitive and
// trims whitespace so "abc123" / " ABC123 " are treated the same.
// Hard block — there is no override path; the client cannot save a
// duplicate. Edit the existing seller instead.
function _findTraderDuplicateByPan(db, pan, excludeId) {
  const norm = String(pan || '').trim().toUpperCase();
  if (!norm) return null;
  const params = [norm];
  let sql = 'SELECT id, name, cr, pan, tel FROM traders WHERE UPPER(TRIM(pan)) = ?';
  if (excludeId) { sql += ' AND id != ?'; params.push(excludeId); }
  sql += ' LIMIT 1';
  return db.get(sql, params);
}
app.post('/api/traders', requireTraderWrite, (req, res) => {
  const t = req.body;
  const db = getDb();
  const dup = _findTraderDuplicateByPan(db, t.pan);
  if (dup) return res.status(409).json({
    duplicate: true, field: 'pan', existing: dup,
    error: `A seller with PAN "${dup.pan}" already exists: ${dup.name || '(unnamed)'}`,
  });
  const info = db.run(`INSERT INTO traders (name,cr,pan,tel,aadhar,padd,ppla,pin,pstate,pst_code,ifsc,acctnum,holder_name)
    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)`,
    [t.name,t.cr||'',t.pan||'',t.tel||'',t.aadhar||'',t.padd||'',t.ppla||'',t.pin||'',t.pstate||'',t.pst_code||'',t.ifsc||'',t.acctnum||'',t.holder_name||'']);
  // If the client sent a banks array (new multi-bank UI), persist them.
  // Otherwise honor the legacy single-bank fields already inserted above.
  if (Array.isArray(t.banks)) {
    syncTraderBanks(db, info.lastInsertRowid, t.banks);
  }
  res.json({ success: true, id: info.lastInsertRowid });
});
app.put('/api/traders/:id', requireTraderWrite, (req, res) => {
  const t = req.body;
  const db = getDb();
  const tid = parseInt(req.params.id, 10);
  const dup = _findTraderDuplicateByPan(db, t.pan, tid);
  if (dup) return res.status(409).json({
    duplicate: true, field: 'pan', existing: dup,
    error: `Another seller with PAN "${dup.pan}" already exists: ${dup.name || '(unnamed)'}`,
  });
  db.run(`UPDATE traders SET name=?,cr=?,pan=?,tel=?,aadhar=?,padd=?,ppla=?,pin=?,pstate=?,pst_code=?,ifsc=?,acctnum=?,holder_name=? WHERE id=?`,
    [t.name,t.cr||'',t.pan||'',t.tel||'',t.aadhar||'',t.padd||'',t.ppla||'',t.pin||'',t.pstate||'',t.pst_code||'',t.ifsc||'',t.acctnum||'',t.holder_name||'',tid]);
  if (Array.isArray(t.banks)) {
    syncTraderBanks(db, tid, t.banks);
  }
  res.json({ success: true });
});
app.delete('/api/traders/:id', requireDelete, (req, res) => {
  const db = getDb();
  // Clear child rows first (trader_banks FK) before deleting the parent
  db.run('DELETE FROM trader_banks WHERE trader_id = ?', [req.params.id]);
  db.run('DELETE FROM traders WHERE id = ?', [req.params.id]);
  res.json({ success: true });
});

// ── LOT-ENTRY QUICK-ADD SELLER ─────────────────────────────────
// Field-staff endpoint for adding a seller on the fly during the auction
// hall workflow. Same shape as POST /api/traders but reachable by the
// lot_entry role (which doesn't have full trader_write capability).
// Validates only the minimum required fields — admins can fill out
// missing details later from the Sellers tab.
app.post('/api/traders/quick', requireAnyPermission('trader_write', 'lot_write'), (req, res) => {
  const t = req.body || {};
  if (!t.name || !String(t.name).trim()) {
    return res.status(400).json({ error: 'Name is required' });
  }
  const db = getDb();
  // Hard duplicate-PAN block — same logic as POST /api/traders so an
  // auction-hall user (lot_entry role) cannot accidentally re-create a
  // seller that already exists with the same PAN under a different
  // name spelling. Soft (name+cr / name+tel) de-dupe below still
  // applies when no PAN is provided.
  const panDup = _findTraderDuplicateByPan(db, t.pan);
  if (panDup) return res.status(409).json({
    duplicate: true, field: 'pan', existing: panDup,
    error: `A seller with PAN "${panDup.pan}" already exists: ${panDup.name || '(unnamed)'}`,
  });
  // De-dupe: if a seller with the same name AND (CR or phone) already
  // exists, return that one instead of creating a duplicate. Helps when
  // multiple field users create the same seller around the same time.
  let existing = null;
  if (t.cr && String(t.cr).trim()) {
    existing = db.get('SELECT * FROM traders WHERE name = ? AND cr = ? LIMIT 1',
      [String(t.name).trim(), String(t.cr).trim()]);
  }
  if (!existing && t.tel && String(t.tel).trim()) {
    existing = db.get('SELECT * FROM traders WHERE name = ? AND tel = ? LIMIT 1',
      [String(t.name).trim(), String(t.tel).trim()]);
  }
  if (existing) {
    return res.json({ success: true, id: existing.id, deduped: true, trader: existing });
  }
  const info = db.run(`INSERT INTO traders (name,cr,pan,tel,aadhar,padd,ppla,pin,pstate,pst_code,ifsc,acctnum,holder_name)
    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)`,
    [
      String(t.name).trim().toUpperCase(),
      (t.cr || '').toString().trim(),
      (t.pan || '').toString().trim().toUpperCase(),
      (t.tel || '').toString().trim(),
      (t.aadhar || '').toString().trim(),
      (t.padd || '').toString().trim(),
      (t.ppla || '').toString().trim().toUpperCase(),
      (t.pin || '').toString().trim(),
      (t.pstate || 'TAMIL NADU').toString().trim().toUpperCase(),
      (t.pst_code || '33').toString().trim(),
      '', '', ''
    ]);
  // Read back the row to return the full trader shape (matches what the
  // search endpoint returns so the client can drop it straight into the
  // selectedSeller state).
  const created = db.get('SELECT * FROM traders WHERE id = ?', [info.lastInsertRowid]);
  if (created) created.banks = [];
  res.json({ success: true, id: info.lastInsertRowid, trader: created });
});

// Set a bank as the trader's default. Used by the Lot Entry bank-picker
// so picking a bank on a lot save updates the trader's default for next
// time. Also syncs the legacy traders.acctnum/ifsc/holder_name fields
// since several existing exports read directly from the traders row.
app.put('/api/traders/:id/bank-default/:bankId', requireAnyPermission('trader_write', 'lot_write'), (req, res) => {
  const traderId = parseInt(req.params.id, 10);
  const bankId   = parseInt(req.params.bankId, 10);
  if (!Number.isFinite(traderId) || !Number.isFinite(bankId)) {
    return res.status(400).json({ error: 'Invalid trader or bank id' });
  }
  const db = getDb();
  const bank = db.get('SELECT * FROM trader_banks WHERE id = ? AND trader_id = ?', [bankId, traderId]);
  if (!bank) return res.status(404).json({ error: 'Bank not found for this trader' });
  // Clear is_default on every bank for this trader, then set it on the
  // chosen one. Done in two statements (sql.js doesn't support RETURNING
  // and we want both updates atomic-feeling without extra wiring).
  db.run('UPDATE trader_banks SET is_default = 0 WHERE trader_id = ?', [traderId]);
  db.run('UPDATE trader_banks SET is_default = 1 WHERE id = ?', [bankId]);
  // Sync the legacy single-bank fields on traders. Existing exports
  // (DBF, XLSX) read these columns directly so we keep them in lockstep
  // with the chosen default.
  db.run('UPDATE traders SET acctnum = ?, ifsc = ?, holder_name = ? WHERE id = ?',
    [bank.acctnum || '', bank.ifsc || '', bank.holder_name || '', traderId]);
  res.json({ success: true });
});

// ── Import Sellers from XLS/XLSX ──────────────────────────────
app.post('/api/traders/import', requireTraderWrite, upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
  try {
    const workbook = XLSX.readFile(req.file.path);
    const ws = workbook.Sheets[workbook.SheetNames[0]];
    if (!ws) throw new Error('No worksheet found');
    const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
    if (!rows.length) throw new Error('File is empty');

    const db = getDb();
    const mode = req.body.mode || 'append';
    if (mode === 'replace') {
      // Wipe child rows (trader_banks FK) before parents — avoids FK error
      db.run('DELETE FROM trader_banks');
      db.run('DELETE FROM traders');
    }

    // Build flexible header map — normalize all keys to uppercase
    const mapCol = (row, ...names) => {
      for (const n of names) { if (row[n] !== undefined) return String(row[n]).trim(); }
      // Also try uppercase/lowercase variants
      const keys = Object.keys(row);
      for (const n of names) {
        const found = keys.find(k => k.toUpperCase() === n.toUpperCase());
        if (found && row[found] !== undefined) return String(row[found]).trim();
      }
      return '';
    };

    let imported = 0, skipped = 0;
    for (const row of rows) {
      const name = mapCol(row, 'NAME', 'SELLER', 'POOLER', 'TRADER');
      if (!name) { skipped++; continue; }

      const cr = mapCol(row, 'CR', 'GSTIN', 'CR_NO', 'CRNO');
      if (mode === 'append') {
        const existing = db.get('SELECT id FROM traders WHERE name = ? AND cr = ?', [name, cr]);
        if (existing) { skipped++; continue; }
      }

      db.run(`INSERT INTO traders (name,cr,pan,tel,aadhar,padd,ppla,pin,pstate,pst_code,ifsc,acctnum,holder_name) 
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)`,
        [name, cr,
         mapCol(row, 'PAN', 'PAN_NO'),
         mapCol(row, 'TEL', 'PHONE', 'MOBILE', 'CONTACT'),
         mapCol(row, 'AADHAR', 'AADHAAR', 'AADHAR_NO'),
         mapCol(row, 'PADD', 'ADDRESS', 'ADD', 'ADD1', 'ADDRESS1'),
         mapCol(row, 'PPLA', 'PLACE', 'PLA', 'CITY'),
         mapCol(row, 'PIN', 'PPIN', 'PINCODE', 'ZIP'),
         mapCol(row, 'PSTATE', 'STATE'),
         mapCol(row, 'PST_CODE', 'ST_CODE', 'STATE_CODE', 'STATECODE'),
         mapCol(row, 'IFSC', 'IFS_CODE', 'IFSCODE', 'IFS'),
         mapCol(row, 'ACCTNUM', 'ACCOUNT', 'ACCNO', 'ACC_NO', 'ACCOUNT_NO', 'ACCOUNTNO'),
         mapCol(row, 'HOLDER_NAME', 'HOLDER', 'ACCOUNT_HOLDER')]);
      imported++;
    }

    fs.unlink(req.file.path, () => {});
    res.json({ success: true, imported, skipped, total: rows.length });
  } catch (e) {
    if (req.file) fs.unlink(req.file.path, () => {});
    res.status(400).json({ error: e.message });
  }
});

// ── Download Seller template XLSX ────────────────────────────
app.get('/api/traders/template', requireExport, async (req, res) => {
  // Use the central createExcelBuffer so the template gets the same brand
  // band, header alignment, and Indian numFmts as every other XLSX export.
  const db = getDb();
  const cols = [
    { header: 'NAME',         key: 'name',         width: 30 },
    { header: 'CR',           key: 'cr',           width: 28 },
    { header: 'PAN',          key: 'pan',          width: 14 },
    { header: 'TEL',          key: 'tel',          width: 16 },
    { header: 'AADHAR',       key: 'aadhar',       width: 16 },
    { header: 'PADD',         key: 'padd',         width: 50 },
    { header: 'PPLA',         key: 'ppla',         width: 20 },
    { header: 'PIN',          key: 'pin',          width: 10 },
    { header: 'PSTATE',       key: 'pstate',       width: 14 },
    { header: 'PST_CODE',     key: 'pst_code',     width: 10, align: 'left' },
    { header: 'IFSC',         key: 'ifsc',         width: 18 },
    { header: 'ACCTNUM',      key: 'acctnum',      width: 24, align: 'left' },
    { header: 'HOLDER_NAME',  key: 'holder_name',  width: 22 },
  ];
  // Sample row uses the configured business state — no hardcoded 'TAMIL NADU'
  const bizState = (getSetting(db, 'business_state') || 'TAMIL NADU').toUpperCase();
  const stCode = bizState === 'KERALA' ? '32' : '33';
  const sample = [{
    name: 'SAMPLE SELLER', cr: 'CR.12345', pan: 'ABCDE1234F', tel: '9876543210',
    aadhar: '', padd: '123 MAIN STREET', ppla: '', pin: '',
    pstate: bizState, pst_code: stCode, ifsc: '', acctnum: '', holder_name: 'SAMPLE SELLER',
  }];
  const buf = await createExcelBuffer('Sellers', cols, sample, {
    db, title: 'SELLERS TEMPLATE',
    metaLines: [`Date: ${fmtDate(todayLocalISO())}`],
  });
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename="sellers-template.xlsx"');
  res.send(Buffer.from(buf));
});

// ══════════════════════════════════════════════════════════════
// BUYERS (SBL.DBF — dealers/traders)
// ══════════════════════════════════════════════════════════════
app.get('/api/buyers', requireView, (req, res) => {
  const { search } = req.query;
  const db = getDb();

  // ── Three response modes ────────────────────────────────────
  //   • `?all=1`                          → every buyer, no limit
  //         (used by the lot-edit modal's multi-code picker to seed
  //         its in-memory cache — needs >500 rows on installs with
  //         large buyer masters)
  //   • `?page=` + `?pageSize=`           → { rows, total, page, pageSize }
  //   • neither                           → legacy plain-array, capped 500
  const wantAll   = String(req.query.all || '') === '1';
  const wantPaged = req.query.page != null || req.query.pageSize != null;

  let where = '';
  let params = [];
  if (search) {
    const q = `%${search}%`;
    where = `WHERE buyer LIKE ? OR buyer1 LIKE ? OR tel LIKE ? OR gstin LIKE ? OR pan LIKE ? OR pla LIKE ? OR ti LIKE ? OR code LIKE ?`;
    params = [q, q, q, q, q, q, q, q];
  }

  if (wantAll) {
    // No LIMIT — caller takes responsibility for the row count.
    return res.json(db.all(`SELECT * FROM buyers ${where} ORDER BY buyer1`, params));
  }

  if (wantPaged) {
    const page     = Math.max(1, parseInt(req.query.page, 10) || 1);
    const pageSize = Math.min(500, Math.max(1, parseInt(req.query.pageSize, 10) || 50));
    const offset   = (page - 1) * pageSize;
    const total = db.get(`SELECT COUNT(*) as c FROM buyers ${where}`, params).c;
    const rows = db.all(
      `SELECT * FROM buyers ${where} ORDER BY buyer1 LIMIT ? OFFSET ?`,
      [...params, pageSize, offset]
    );
    return res.json({ rows, total, page, pageSize });
  }

  // Legacy plain-array response — capped to keep old UI fast.
  if (search) {
    return res.json(db.all(
      `SELECT * FROM buyers ${where} ORDER BY buyer1 LIMIT 50`,
      params
    ));
  }
  res.json(db.all('SELECT * FROM buyers ORDER BY buyer1 LIMIT 500'));
});
// Duplicate-key guard for buyers. Both the full buyer code (`buyer`)
// and the short alias (`code`) are operator-typed identifiers used
// across lots / invoices / price files — two rows sharing either is
// almost always an accidental re-entry. Returns the first match,
// preferring a `buyer` hit so the toast names the more-prominent
// collision when both fields conflict.
// Hard block — there is no override path; the client cannot save a
// duplicate. Edit the existing buyer instead.
function _findBuyerDuplicate(db, buyer, code, excludeId) {
  const nb = String(buyer || '').trim().toUpperCase();
  const nc = String(code  || '').trim().toUpperCase();
  const buildSql = (col, val) => {
    const params = [val];
    let sql = `SELECT id, buyer, buyer1, code FROM buyers WHERE UPPER(TRIM(${col})) = ?`;
    if (excludeId) { sql += ' AND id != ?'; params.push(excludeId); }
    sql += ' LIMIT 1';
    return { sql, params };
  };
  if (nb) {
    const q = buildSql('buyer', nb);
    const hit = db.get(q.sql, q.params);
    if (hit) return { row: hit, field: 'buyer' };
  }
  if (nc) {
    const q = buildSql('code', nc);
    const hit = db.get(q.sql, q.params);
    if (hit) return { row: hit, field: 'code' };
  }
  return null;
}
app.post('/api/buyers', requireBuyerWrite, (req, res) => {
  const b = req.body;
  const db = getDb();
  const dup = _findBuyerDuplicate(db, b.buyer, b.code);
  if (dup) return res.status(409).json({
    duplicate: true, field: dup.field, existing: dup.row,
    error: `A buyer with ${dup.field === 'buyer' ? `code "${dup.row.buyer}"` : `short alias "${dup.row.code}"`} already exists${dup.row.buyer1 ? `: ${dup.row.buyer1}` : ''}`,
  });
  db.run(`INSERT INTO buyers (
      buyer, buyer1, code, sbl, add1, add2, pla, pin, state, st_code,
      gstin, pan, tel, ti, sale, email, tdsq,
      cbuyer1, cadd1, cadd2, cpla, cpin, cstate, cst_code, cgstin
    ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`,
    [b.buyer, b.buyer1||'', b.code||'', b.sbl||'', b.add1||'', b.add2||'', b.pla||'', b.pin||'', b.state||'', b.st_code||'',
     b.gstin||'', b.pan||'', b.tel||'', b.ti||'', b.sale||'L', b.email||'', b.tdsq||'',
     b.cbuyer1||'', b.cadd1||'', b.cadd2||'', b.cpla||'', b.cpin||'', b.cstate||'', b.cst_code||'', b.cgstin||'']);
  res.json({ success: true });
});
app.put('/api/buyers/:id', requireBuyerWrite, (req, res) => {
  const b = req.body;
  const db = getDb();
  const bid = parseInt(req.params.id, 10);
  const dup = _findBuyerDuplicate(db, b.buyer, b.code, bid);
  if (dup) return res.status(409).json({
    duplicate: true, field: dup.field, existing: dup.row,
    error: `Another buyer with ${dup.field === 'buyer' ? `code "${dup.row.buyer}"` : `short alias "${dup.row.code}"`} already exists${dup.row.buyer1 ? `: ${dup.row.buyer1}` : ''}`,
  });
  db.run(`UPDATE buyers SET
      buyer=?, buyer1=?, code=?, sbl=?, add1=?, add2=?, pla=?, pin=?, state=?, st_code=?,
      gstin=?, pan=?, tel=?, ti=?, sale=?, email=?, tdsq=?,
      cbuyer1=?, cadd1=?, cadd2=?, cpla=?, cpin=?, cstate=?, cst_code=?, cgstin=?
    WHERE id=?`,
    [b.buyer, b.buyer1||'', b.code||'', b.sbl||'', b.add1||'', b.add2||'', b.pla||'', b.pin||'', b.state||'', b.st_code||'',
     b.gstin||'', b.pan||'', b.tel||'', b.ti||'', b.sale||'L', b.email||'', b.tdsq||'',
     b.cbuyer1||'', b.cadd1||'', b.cadd2||'', b.cpla||'', b.cpin||'', b.cstate||'', b.cst_code||'', b.cgstin||'',
     bid]);
  res.json({ success: true });
});
app.delete('/api/buyers/:id', requireDelete, (req, res) => {
  getDb().run('DELETE FROM buyers WHERE id = ?', [req.params.id]); res.json({ success: true });
});

// ── Import Buyers from XLS/XLSX ───────────────────────────────
app.post('/api/buyers/import', requireBuyerWrite, upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
  try {
    const workbook = XLSX.readFile(req.file.path);
    const ws = workbook.Sheets[workbook.SheetNames[0]];
    if (!ws) throw new Error('No worksheet found');
    const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
    if (!rows.length) throw new Error('File is empty');

    const db = getDb();
    const mode = req.body.mode || 'append';
    if (mode === 'replace') db.run('DELETE FROM buyers');

    const mapCol = (row, ...names) => {
      for (const n of names) { if (row[n] !== undefined) return String(row[n]).trim(); }
      const keys = Object.keys(row);
      for (const n of names) {
        const found = keys.find(k => k.toUpperCase() === n.toUpperCase());
        if (found && row[found] !== undefined) return String(row[found]).trim();
      }
      return '';
    };

    let imported = 0, skipped = 0;
    for (const row of rows) {
      // BUYER = full buyer code (primary key in lot.buyer → matches invoice lookup)
      // CODE  = short alias printed on tags (e.g. RSH, TE, SL) — used by post-auction price files
      // The two may be the same value in some files, or different. Treat them as distinct columns.
      const buyer = mapCol(row, 'BUYER', 'BUYER_CODE', 'BUYERCODE');
      const code  = mapCol(row, 'CODE', 'SHORT_CODE', 'ALIAS');
      if (!buyer && !code) { skipped++; continue; }
      // If BUYER column missing, fall back to CODE, then trade name
      const buyerVal = buyer || code || mapCol(row, 'BUYER1', 'TRADE_NAME', 'TRADENAME', 'NAME');
      if (!buyerVal) { skipped++; continue; }

      if (mode === 'append') {
        const existing = db.get('SELECT id FROM buyers WHERE buyer = ?', [buyerVal]);
        if (existing) { skipped++; continue; }
      }

      db.run(`INSERT INTO buyers (
        buyer, buyer1, code, sbl, add1, add2, pla, pin, state, st_code,
        gstin, pan, tel, ti, sale, email, tdsq,
        cbuyer1, cadd1, cadd2, cpla, cpin, cstate, cst_code, cgstin
      ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`,
        [buyerVal,
         mapCol(row, 'BUYER1', 'TRADE_NAME', 'TRADENAME', 'NAME'),
         code,
         mapCol(row, 'SBL', 'SBLNO'),
         mapCol(row, 'ADD1', 'ADDRESS1', 'ADDRESS'),
         mapCol(row, 'ADD2', 'ADDRESS2'),
         mapCol(row, 'PLA', 'PLACE', 'CITY'),
         mapCol(row, 'PIN', 'PINCODE', 'ZIP'),
         mapCol(row, 'STATE'),
         mapCol(row, 'ST_CODE', 'STATE_CODE', 'STATECODE'),
         mapCol(row, 'GSTIN', 'GST', 'GSTNO', 'GST_NO'),
         mapCol(row, 'PAN', 'PAN_NO'),
         mapCol(row, 'TEL', 'PHONE', 'MOBILE', 'CONTACT'),
         mapCol(row, 'TI'),
         mapCol(row, 'SALE', 'SALE_TYPE') || 'L',
         mapCol(row, 'EMAIL', 'E_MAIL', 'MAIL'),
         mapCol(row, 'TDSQ', 'TDS_Q', 'TDS'),
         // Consignee (ship-to) details
         mapCol(row, 'CBUYER1', 'CONSIGNEE', 'CONSIGNEE_NAME'),
         mapCol(row, 'CADD1', 'CONS_ADD1', 'CONSIGNEE_ADDRESS1'),
         mapCol(row, 'CADD2', 'CONS_ADD2', 'CONSIGNEE_ADDRESS2'),
         mapCol(row, 'CPLA', 'CONS_PLA', 'CONSIGNEE_PLACE'),
         mapCol(row, 'CPIN', 'CONS_PIN', 'CONSIGNEE_PIN'),
         mapCol(row, 'CSTATE', 'CONS_STATE'),
         mapCol(row, 'CST_CODE', 'CONS_ST_CODE'),
         mapCol(row, 'CGSTIN', 'CONS_GSTIN')]);
      imported++;
    }

    fs.unlink(req.file.path, () => {});
    res.json({ success: true, imported, skipped, total: rows.length });
  } catch (e) {
    if (req.file) fs.unlink(req.file.path, () => {});
    res.status(400).json({ error: e.message });
  }
});

// ── Download Buyer template XLSX ─────────────────────────────
app.get('/api/buyers/template', requireExport, async (req, res) => {
  const db = getDb();
  const cols = [
    { header: 'BUYER',    key: 'buyer',    width: 14 },
    { header: 'BUYER1',   key: 'buyer1',   width: 30 },
    { header: 'ADD1',     key: 'add1',     width: 30 },
    { header: 'ADD2',     key: 'add2',     width: 30 },
    { header: 'PLA',      key: 'pla',      width: 18 },
    { header: 'PIN',      key: 'pin',      width: 10, align: 'left' },
    { header: 'STATE',    key: 'state',    width: 16 },
    { header: 'ST_CODE',  key: 'st_code',  width: 10, align: 'left' },
    { header: 'GSTIN',    key: 'gstin',    width: 18 },
    { header: 'PAN',      key: 'pan',      width: 14 },
    { header: 'TEL',      key: 'tel',      width: 14 },
    { header: 'TI',       key: 'ti',       width: 10 },
    { header: 'SALE',     key: 'sale',     width: 8  },
  ];
  const bizState = (getSetting(db, 'business_state') || 'TAMIL NADU').toUpperCase();
  const stCode = bizState === 'KERALA' ? '32' : '33';
  const sample = [{
    buyer: 'ABC', buyer1: 'ABC TRADERS', add1: '10 MARKET ROAD', add2: '', pla: '',
    pin: '', state: bizState, st_code: stCode, gstin: '', pan: '', tel: '', ti: '', sale: 'L',
  }];
  const buf = await createExcelBuffer('Buyers', cols, sample, {
    db, title: 'BUYERS TEMPLATE',
    metaLines: [`Date: ${fmtDate(todayLocalISO())}`],
  });
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename="buyers-template.xlsx"');
  res.send(Buffer.from(buf));
});

// ══════════════════════════════════════════════════════════════
// BUYERS — lookup helpers used by the Lot edit modal so the operator
// can pick from every buyer code that shares a trade name. The
// `buyer1` column is the canonical "trade name"; multiple buyer rows
// can share one — different GSTINs, different consignees, different
// sale-type defaults — and after a Price List mapping the operator
// needs to pick the right one per lot.
// ══════════════════════════════════════════════════════════════
app.get('/api/buyers/by-tradename', requireView, (req, res) => {
  const name = String(req.query.name || '').trim();
  if (!name) return res.json([]);
  // Case-insensitive exact match on buyer1 (primary) OR buyer (the
  // full buyer-code string can also be passed as a trade name from
  // free-text fields). Sort by code so the picker is stable.
  // `tel` is included so the WhatsApp share buttons can look up a
  // buyer's phone with one round-trip (same shape sellers use via
  // /api/traders/by-name).
  const rows = getDb().all(
    `SELECT id, buyer, buyer1, code, ti, sale, gstin, pla, tel
       FROM buyers
      WHERE UPPER(TRIM(buyer1)) = UPPER(TRIM(?))
         OR UPPER(TRIM(buyer))  = UPPER(TRIM(?))
      ORDER BY code, buyer`,
    [name, name]
  );
  res.json(rows);
});

// ══════════════════════════════════════════════════════════════
// PRICE LIST (BEFORE) — code mapping tool
// ══════════════════════════════════════════════════════════════
// Operator workflow:
//   1. Export an empty Price List (Before) sheet (Exports tab)
//   2. Print → buyers write their TRADE NAME (and prices) by hand
//   3. Type the trade names back into the file
//   4. Upload here → server resolves CODE from the buyers master
//   5. Preview the matches; download the updated file
//   6. Feed the downloaded file into Lots → Price Import
//
// Two endpoints share the same parsing/matching code:
//   POST /api/price-list/map-preview  → JSON summary only
//   POST /api/price-list/map-download → updated XLSX (Buffer)
//
// `ExcelJS` is used end-to-end so the brand header / total row /
// column widths from the original export survive the round-trip.
function _plLocateColumns(ws) {
  // Find the header row containing both "TRADE NAME" and "CODE".
  // Match is case-insensitive + whitespace-tolerant so renamed headers
  // (e.g. "Trade Name", "trade_name") still resolve.
  const normalize = s => String(s == null ? '' : s).trim().toUpperCase().replace(/[\s_\-]+/g, ' ');
  let headerRow = 0, tradeCol = 0, codeCol = 0, anoCol = 0, dateCol = 0, lotCol = 0;
  const maxRow = ws.rowCount || 0;
  for (let r = 1; r <= maxRow; r++) {
    const row = ws.getRow(r);
    const cells = {};
    row.eachCell({ includeEmpty: false }, (cell, col) => {
      cells[normalize(cell.value)] = col;
    });
    if (cells['TRADE NAME'] && cells['CODE']) {
      headerRow = r;
      tradeCol = cells['TRADE NAME'];
      codeCol = cells['CODE'];
      anoCol = cells['AUCTION NO'] || cells['ANO'] || cells['TNO'] || 0;
      dateCol = cells['DATE'] || 0;
      lotCol = cells['LOT'] || cells['LOT NO'] || cells['LOTNO'] || 0;
      break;
    }
  }
  return { headerRow, tradeCol, codeCol, anoCol, dateCol, lotCol };
}
function _plBuildTradeIndex(db) {
  // Pre-index every buyer by their trade name so a 1000-row file is one
  // DB query, not 1000. We index BOTH buyer1 and buyer because operators
  // sometimes write the full buyer-code string in the TRADE NAME column.
  const buyers = db.all('SELECT id, buyer, buyer1, code, ti, sale, gstin FROM buyers');
  const idx = new Map();
  const push = (key, row) => {
    if (!key) return;
    const k = key.trim().toUpperCase();
    if (!k) return;
    if (!idx.has(k)) idx.set(k, []);
    // Avoid duplicate entries when buyer === buyer1.
    const arr = idx.get(k);
    if (!arr.some(b => b.id === row.id)) arr.push(row);
  };
  for (const b of buyers) {
    push(b.buyer1, b);
    push(b.buyer, b);
  }
  return idx;
}
async function _plProcessFile(filePath) {
  // Returns { wb, ws, cols, perRow: [{row, tradeName, status, pickedCode, candidates}], summary }
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(filePath);
  const ws = wb.worksheets[0];
  if (!ws) throw new Error('Workbook has no worksheets.');
  const cols = _plLocateColumns(ws);
  if (!cols.headerRow || !cols.tradeCol || !cols.codeCol) {
    throw new Error('Could not locate TRADE NAME and CODE columns. The sheet must have both headers.');
  }
  const idx = _plBuildTradeIndex(getDb());
  const perRow = [];
  let matched = 0, unmatched = 0, ambiguous = 0, blank = 0;
  const maxRow = ws.rowCount || 0;
  for (let r = cols.headerRow + 1; r <= maxRow; r++) {
    const row = ws.getRow(r);
    const tradeRaw = row.getCell(cols.tradeCol).value;
    const tradeName = String(tradeRaw == null ? '' : tradeRaw).trim();
    const entry = {
      row: r,
      tradeName,
      currentCode: String(row.getCell(cols.codeCol).value || '').trim(),
      ano:   cols.anoCol  ? String(row.getCell(cols.anoCol).value  || '').trim() : '',
      date:  cols.dateCol ? String(row.getCell(cols.dateCol).value || '').trim() : '',
      lot:   cols.lotCol  ? String(row.getCell(cols.lotCol).value  || '').trim() : '',
      status: 'blank',
      pickedCode: '',
      candidates: [],
    };
    if (!tradeName) {
      blank++;
      perRow.push(entry);
      continue;
    }
    const key = tradeName.toUpperCase();
    const cands = idx.get(key) || [];
    entry.candidates = cands.map(b => ({
      id: b.id, code: b.code, buyer: b.buyer, buyer1: b.buyer1, sale: b.sale, gstin: b.gstin,
    }));
    if (cands.length === 0) {
      entry.status = 'unmatched';
      unmatched++;
    } else if (cands.length === 1) {
      entry.status = 'matched';
      entry.pickedCode = cands[0].code || '';
      matched++;
    } else {
      // Ambiguous — multiple buyers share this trade name. Pick the
      // first by code-sort order (same order as /api/buyers/by-tradename
      // so the UI and the file agree). Operator resolves per-lot later
      // using the multi-code picker in the Lot edit modal.
      entry.status = 'ambiguous';
      // Prefer a candidate with a non-blank code; fall back to first.
      const withCode = cands.find(c => c.code && String(c.code).trim());
      entry.pickedCode = (withCode || cands[0]).code || '';
      ambiguous++;
    }
    perRow.push(entry);
  }
  const summary = {
    total: perRow.length,
    matched, ambiguous, unmatched, blank,
    uniqueTradeNames: Array.from(new Set(perRow.filter(p => p.tradeName).map(p => p.tradeName))).length,
  };
  return { wb, ws, cols, perRow, summary };
}
app.post('/api/price-list/map-preview', requireView, upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
  try {
    const { perRow, summary } = await _plProcessFile(req.file.path);
    res.json({ ...summary, rows: perRow });
  } catch (e) {
    res.status(400).json({ error: e.message });
  } finally {
    if (req.file) fs.unlink(req.file.path, () => {});
  }
});
app.post('/api/price-list/map-download', requireView, upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
  try {
    const { wb, ws, cols, perRow } = await _plProcessFile(req.file.path);
    // Write the resolved code back into each row's CODE cell. We
    // explicitly set the cell value so the existing column-level numFmt
    // (which Excel uses to right-pad short codes like "RSH") still
    // applies — modifying `.value` keeps the format intact.
    for (const entry of perRow) {
      if (!entry.pickedCode) continue;
      ws.getRow(entry.row).getCell(cols.codeCol).value = entry.pickedCode;
    }
    const buf = await wb.xlsx.writeBuffer();
    const baseName = (req.file.originalname || 'price-list-before.xlsx')
      .replace(/\.xlsx?$/i, '')
      .replace(/[^A-Za-z0-9._-]+/g, '-');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${baseName}-mapped.xlsx"`);
    res.send(Buffer.from(buf));
  } catch (e) {
    res.status(400).json({ error: e.message });
  } finally {
    if (req.file) fs.unlink(req.file.path, () => {});
  }
});

// ══════════════════════════════════════════════════════════════
// PRICE CHECK — reconcile an external price sheet against the
// `lots` table. Replaces the legacy PriceCheck_VSTL.xlsm
// (VBA macros `priceCheck` + `checkRate`). See price-check.js for
// the matching logic and round-trip XLSX annotation.
//
//   POST /api/price-check/verify    → JSON preview (per-row + summary)
//   POST /api/price-check/download  → XLSX with red/amber highlighting
//
// Optional form field `auction_id` forces every row to match against
// that auction — used when the operator pre-picks an auction in the
// Price Check tab. Without it, rows resolve per-row via TNO + DATE.
// ══════════════════════════════════════════════════════════════
const priceCheck = require('./price-check');

// ── Price-check gate helpers ─────────────────────────────────
// Tri-state gate. The auction carries two timestamps:
//   - price_check_first_passed_at: set on the FIRST successful verify
//     and never cleared. Tells us the operator has reconciled this
//     auction at least once.
//   - price_checked_at: set on every successful verify AND cleared by
//     any endpoint that mutates a lot's price/code. Tells us the
//     reconciliation is still current.
//
// Gate states (pcGateState):
//   'off'   — feature flag disabled, treat as clean
//   'never' — auction has never passed verify → hard 412 block
//   'stale' — verified at least once, but lots changed since →
//             allow the action, surface a soft warning in the UI
//   'clean' — verified and no edits since → green light, no warning
//
// Everything below is gated by the `flag_price_check` feature flag.
// When the flag is OFF: stamp / clear become no-ops, state is always
// 'off', and requirePriceChecked just calls next(). The whole
// subsystem (tab, banners, button disables) collapses to a no-op,
// matching the off-by-default rollout we want.
function pcFlagOn(db) {
  try {
    const cfg = getSettingsFlat(db || getDb());
    return String(cfg.flag_price_check || '').toLowerCase() === 'true';
  } catch (_) { return false; }
}
function pcStampGate(db, auctionId) {
  if (!auctionId) return;
  if (!pcFlagOn(db)) return;
  // Stamp the current-verify timestamp on every successful pass, and
  // set the first-pass timestamp only the first time (COALESCE keeps
  // the original first-pass date when re-stamping).
  db.run(
    `UPDATE auctions
        SET price_checked_at = datetime('now','localtime'),
            price_check_first_passed_at = COALESCE(NULLIF(price_check_first_passed_at, ''), datetime('now','localtime'))
      WHERE id = ?`,
    [auctionId]
  );
}
function pcClearGate(db, auctionId) {
  if (!auctionId) return;
  if (!pcFlagOn(db)) return;
  // Only clear the current-verify timestamp. The first-pass stamp is
  // permanent for the life of the auction — once verified, edits drop
  // the gate to 'stale' (soft warning) instead of 'never' (hard block).
  db.run(`UPDATE auctions SET price_checked_at = '' WHERE id = ?`, [auctionId]);
}
function pcGateState(db, auctionId) {
  if (!auctionId)   return 'never';
  if (!pcFlagOn(db)) return 'off';
  const row = db.get(
    'SELECT price_checked_at, price_check_first_passed_at FROM auctions WHERE id = ?',
    [auctionId]
  );
  if (!row) return 'never';
  if (!row.price_check_first_passed_at) return 'never';
  return row.price_checked_at ? 'clean' : 'stale';
}
// Kept for any legacy callers — true only when the auction is fully
// reconciled (clean OR feature off). 'stale' is NOT ready by this
// definition; ready-vs-allowed are intentionally different concepts.
function pcGateReady(db, auctionId) {
  const s = pcGateState(db, auctionId);
  return s === 'clean' || s === 'off';
}
// Express middleware factory: drops a 412 PRECONDITION FAILED only on
// the 'never' state. 'stale' is allowed through (the operator has
// verified at least once; UI surfaces a soft warning instead).
// `getAuctionId(req)` returns the auction id to gate on.
function requirePriceChecked(getAuctionId) {
  return (req, res, next) => {
    if (!pcFlagOn()) return next();        // feature disabled
    const aid = getAuctionId(req);
    if (!aid) return next();               // can't gate a global call
    if (pcGateState(getDb(), aid) !== 'never') return next();
    return res.status(412).json({
      error: 'Price check required',
      detail: 'Run Reports → Price Check against the auction (and apply any code fixes) before this action.',
      auctionId: aid,
      gate: 'price_check',
    });
  };
}

// ══════════════════════════════════════════════════════════════
// LOT VALIDATION GATE — pre-price-import "validate entered lots"
// ══════════════════════════════════════════════════════════════
// Mirrors the price-check gate above, but sits on the OTHER side of
// the workflow: it guards PRICE IMPORT, not the post-import generate
// actions. The incident it prevents: a seller entered with no GSTIN
// silently drops out of GSTIN-keyed reports (Dealer List filters
// LENGTH(gstin)=15), so the Lots screen and the Dealer List disagree
// on lot count. This validates the ENTERED lots and forces the
// operator to see (and acknowledge) those gaps before importing prices.
//
// Gate states (lvGateState):
//   'off'   — flag_lot_validation disabled → no-op, never blocks
//   'never' — not validated since the last lot change → 412 block
//   'clean' — validated (errors=0, warnings acknowledged), no edits since
//
// lots_validated_at is stamped on a clean confirm and cleared by ANY
// lot insert/edit/delete — re-validation is required after every change
// (it's one cheap click), which is stricter than the price-check gate's
// stale-allowed behaviour, on purpose.
function lvFlagOn(db) {
  try {
    const cfg = getSettingsFlat(db || getDb());
    return String(cfg.flag_lot_validation || '').toLowerCase() === 'true';
  } catch (_) { return false; }
}
function lvStampGate(db, auctionId) {
  if (!auctionId) return;
  if (!lvFlagOn(db)) return;
  db.run(`UPDATE auctions SET lots_validated_at = datetime('now','localtime') WHERE id = ?`, [auctionId]);
}
function lvClearGate(db, auctionId) {
  if (!auctionId) return;
  if (!lvFlagOn(db)) return;
  db.run(`UPDATE auctions SET lots_validated_at = '' WHERE id = ?`, [auctionId]);
}
function lvGateState(db, auctionId) {
  if (!auctionId)    return 'never';
  if (!lvFlagOn(db)) return 'off';
  const row = db.get('SELECT lots_validated_at FROM auctions WHERE id = ?', [auctionId]);
  if (!row) return 'never';
  return row.lots_validated_at ? 'clean' : 'never';
}

// Clean a stored GSTIN the same way exportDealerList() does, so the
// validation's "will this lot appear in the Dealer List?" verdict
// matches the report exactly. Storage is inconsistent: "GSTIN.<15>",
// "gstin <15>", bare 15-char, etc. A valid GSTIN cleans to length 15.
function cleanGstin(cr) {
  let s = String(cr == null ? '' : cr).trim();
  if (s.slice(0, 5).toLowerCase() === 'gstin') {
    s = s.slice(5).replace(/^[.\s:-]+/, '');
  }
  return s.trim().toUpperCase();
}
const hasValidGstin = (cr) => cleanGstin(cr).length === 15;

// Pure, read-only: build the validation report for one auction.
// Returns { ok, errors[], warnings[], reconciliation, totals }.
//   errors   — block import (duplicate lot numbers, lots with no seller)
//   warnings — acknowledge to proceed (missing GSTIN / bank / PAN / phone)
function validateAuctionLots(db, auctionId) {
  const lots = db.all(
    `SELECT id, lot_no, trader_id, bank_id, name, cr, pan, tel, branch,
            litre, COALESCE(bags,0) AS bags, COALESCE(qty,0) AS qty
       FROM lots WHERE auction_id = ?`,
    [auctionId]
  );
  // Trader IDs that actually have a bank account on file — one query,
  // O(1) lookup. A lot's payment routes via lots.bank_id or the trader's
  // default bank, so "no bank" = the trader has zero trader_banks rows.
  const tradersWithBank = new Set(
    db.all(`SELECT DISTINCT trader_id FROM trader_banks`).map(r => r.trader_id)
  );

  // Display projection — the fields the UI shows for each flagged lot,
  // matching the Lot Entry "Recent entries" row (Lot / Seller / Branch /
  // Bags / Litre / Qty). Used for every error/warning lot list.
  const disp = (l) => ({
    id: l.id, lot_no: l.lot_no, name: l.name || '', branch: l.branch || '',
    bags: l.bags, litre: l.litre || '', qty: l.qty, trader_id: l.trader_id || null,
  });

  const errors = [];
  const warnings = [];

  // ── Duplicate lot numbers (padding-insensitive: "1" === "001") ──
  const byKey = new Map();
  for (const l of lots) {
    const p = parseLotNo(l.lot_no);
    const key = p ? p.prefix + ':' + p.num : String(l.lot_no).trim().toUpperCase();
    if (!byKey.has(key)) byKey.set(key, []);
    byKey.get(key).push(l);
  }
  for (const [, group] of byKey) {
    if (group.length > 1) {
      errors.push({
        type: 'duplicate_lot',
        title: 'Duplicate lot',
        message: `Lot #${group[0].lot_no} entered ${group.length} times`,
        lots: group.map(disp),
      });
    }
  }

  // ── Lots with no seller (blocks: every lot must have a seller) ──
  const noSeller = lots.filter(l => !l.trader_id || !String(l.name || '').trim());
  if (noSeller.length) {
    errors.push({
      type: 'no_seller',
      title: 'No seller',
      message: `${noSeller.length} lot(s) have no seller linked`,
      lots: noSeller.map(disp),
    });
  }

  // ── Warnings — missing seller details (acknowledge to proceed) ──
  const pushWarn = (type, title, label, predicate) => {
    const hit = lots.filter(predicate);
    if (hit.length) {
      warnings.push({
        type, title, label, count: hit.length,
        lots: hit.map(disp),
      });
    }
  };
  // Flag the GSTIN issue ONLY when the column is blank/empty — not when a
  // GSTIN is present but malformed (wrong length / stray text). cleanGstin
  // strips a leading "GSTIN" label, so a prefix-only value ("GSTIN.") still
  // counts as blank. (The Dealer List reconciliation below keeps its own
  // strict 15-char rule via hasValidGstin — that's the real exclusion gap.)
  pushWarn('no_gstin', 'No GSTIN',       'Seller has no GSTIN (excluded from Dealer List)', l => !cleanGstin(l.cr));
  pushWarn('no_bank',  'No bank account', 'Seller has no bank account on file',             l => l.trader_id && !tradersWithBank.has(l.trader_id));
  pushWarn('no_pan',   'No PAN',          'Seller has no PAN',                              l => !String(l.pan || '').trim());
  pushWarn('no_phone', 'No phone',        'Seller has no phone number',                     l => !String(l.tel || '').trim());

  // ── Reconciliation (the "tally") ──────────────────────────────
  const totalLots = lots.length;
  const totalBags = lots.reduce((s, l) => s + Number(l.bags || 0), 0);
  const totalQty  = Math.round(lots.reduce((s, l) => s + Number(l.qty || 0), 0) * 1000) / 1000;
  const gstinLots = lots.filter(l => hasValidGstin(l.cr)).length;

  // Per-seller breakdown (grouped by trader_id, falling back to name)
  const sellerMap = new Map();
  for (const l of lots) {
    const key = l.trader_id != null ? 't' + l.trader_id : 'n:' + String(l.name || '').trim().toUpperCase();
    if (!sellerMap.has(key)) {
      sellerMap.set(key, { name: l.name || '(no name)', hasGstin: hasValidGstin(l.cr), lots: 0, bags: 0, qty: 0 });
    }
    const s = sellerMap.get(key);
    s.lots += 1; s.bags += Number(l.bags || 0); s.qty += Number(l.qty || 0);
  }
  const sellers = Array.from(sellerMap.values())
    .map(s => ({ ...s, qty: Math.round(s.qty * 1000) / 1000 }))
    .sort((a, b) => Number(a.hasGstin) - Number(b.hasGstin) || a.name.localeCompare(b.name));

  return {
    ok: errors.length === 0,
    errors,
    warnings,
    totals: { lots: totalLots, bags: totalBags, qty: totalQty },
    reconciliation: {
      totalLots,
      dealerListLots: gstinLots,            // appear in the Dealer List
      excludedLots: totalLots - gstinLots,  // the gap — non-GSTIN sellers
      sellerCount: sellers.length,
      sellers,
    },
  };
}

app.post('/api/price-check/verify', requireView, upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
  try {
    const auctionId = req.body.auction_id ? Number(req.body.auction_id) : null;
    const { perRow, summary } = await priceCheck.processFile(
      req.file.path, getDb(),
      { auctionId }
    );
    // Auto-stamp the gate when verify ran against a specific auction
    // AND there are no fixable code issues left. Operators can either
    // (a) upload a clean file, or (b) re-run verify after applying
    // every fixable row — both paths converge on gateReady=true.
    if (auctionId && summary.gateReady) {
      pcStampGate(getDb(), auctionId);
    }
    res.json({ ...summary, rows: perRow });
  } catch (e) {
    res.status(400).json({ error: e.message });
  } finally {
    if (req.file) fs.unlink(req.file.path, () => {});
  }
});

// Lightweight status probe — the client polls this to enable/disable
// transaction CTAs without re-uploading a file.
app.get('/api/auctions/:id/price-check-status', requireView, (req, res) => {
  const id = parseInt(req.params.id, 10);
  if (!id) return res.status(400).json({ error: 'invalid auction id' });
  const db = getDb();
  const row = db.get(
    'SELECT id, ano, date, price_checked_at, price_check_first_passed_at FROM auctions WHERE id = ?', [id]
  );
  if (!row) return res.status(404).json({ error: 'auction not found' });
  // `checked` stays true ONLY for 'clean' so legacy callers keep their
  // strict semantics. `state` is the tri-state field new UI uses to
  // distinguish 'stale' (soft warning) from 'never' (hard block).
  const state = pcGateState(db, id);
  res.json({
    auctionId: id,
    ano: row.ano, date: row.date,
    state,                                     // 'off' | 'never' | 'stale' | 'clean'
    checked: state === 'clean' || state === 'off',
    everPassed: !!row.price_check_first_passed_at,
    checkedAt: row.price_checked_at || null,
    firstPassedAt: row.price_check_first_passed_at || null,
  });
});

app.post('/api/price-check/download', requireView, upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
  try {
    const { wb, ws, cols, perRow } = await priceCheck.processFile(
      req.file.path, getDb(),
      { auctionId: req.body.auction_id || null }
    );
    priceCheck.annotateWorkbook(wb, ws, cols, perRow);
    const buf = await wb.xlsx.writeBuffer();
    // Filename: prefer the auction's ANO if pre-picked, otherwise
    // append "-checked" to the uploaded filename (mirrors the legacy
    // VBA `saveAs` which named files Price{TNO}.xls).
    const aid = req.body.auction_id;
    let baseName;
    if (aid) {
      const auc = getDb().get('SELECT ano FROM auctions WHERE id = ?', [aid]);
      baseName = `Price${auc && auc.ano ? auc.ano : aid}-checked`;
    } else {
      baseName = (req.file.originalname || 'price-check.xlsx')
        .replace(/\.xlsx?$/i, '')
        .replace(/[^A-Za-z0-9._-]+/g, '-') + '-checked';
    }
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${baseName}.xlsx"`);
    res.send(Buffer.from(buf));
  } catch (e) {
    res.status(400).json({ error: e.message });
  } finally {
    if (req.file) fs.unlink(req.file.path, () => {});
  }
});

// ── LOT VALIDATION ENDPOINTS ──────────────────────────────────
// Read-only validation report for the auction's entered lots. Does
// NOT stamp the gate — the operator confirms separately (below) once
// they've reviewed errors + acknowledged warnings.
app.get('/api/auctions/:id/validate-lots', requireViewOrLotEntry, (req, res) => {
  const id = parseInt(req.params.id, 10);
  if (!id) return res.status(400).json({ error: 'invalid auction id' });
  const db = getDb();
  const auc = db.get('SELECT id, ano, date FROM auctions WHERE id = ?', [id]);
  if (!auc) return res.status(404).json({ error: 'auction not found' });
  const report = validateAuctionLots(db, id);
  res.json({
    auctionId: id, ano: auc.ano, date: auc.date,
    flagOn: lvFlagOn(db),
    state: lvGateState(db, id),   // 'off' | 'never' | 'clean'
    ...report,
  });
});

// Confirm validation → stamp the gate. Refuses if there are still
// hard errors; warnings are acknowledged implicitly by confirming
// (the UI only enables this after the operator ticks the box).
app.post('/api/auctions/:id/validate-lots/confirm', requireLotWrite, (req, res) => {
  const id = parseInt(req.params.id, 10);
  if (!id) return res.status(400).json({ error: 'invalid auction id' });
  const db = getDb();
  const auc = db.get('SELECT id FROM auctions WHERE id = ?', [id]);
  if (!auc) return res.status(404).json({ error: 'auction not found' });
  const report = validateAuctionLots(db, id);
  if (!report.ok) {
    return res.status(409).json({
      error: 'Cannot validate — unresolved errors remain',
      detail: 'Fix all errors (duplicate lots, lots with no seller) before marking the trade validated.',
      ...report,
      state: lvGateState(db, id),
    });
  }
  lvStampGate(db, id);
  res.json({ success: true, state: lvGateState(db, id), ...report });
});

// Returns null when the price-import lot-validation gate is satisfied
// (or doesn't apply), otherwise { status, body } describing the block.
// Shared by the upload middleware (below) and the Trade Fair sync, so
// both enforce the gate identically. Only gates mode='price'; mode='full'
// (the step that CREATES the lots) passes through.
function lotValidationGateBlock(db, body) {
  body = body || {};
  if (!lvFlagOn(db)) return null;
  if (String(body.mode || '') !== 'price') return null;
  let aid = body.auction_id ? parseInt(body.auction_id, 10) : null;
  if (!aid && body.ano) {
    const d = normalizeDate(body.date);
    const auc = d
      ? db.get('SELECT id FROM auctions WHERE ano = ? AND date = ?', [body.ano, d])
      : db.get('SELECT id FROM auctions WHERE ano = ? ORDER BY date DESC LIMIT 1', [body.ano]);
    if (auc) aid = auc.id;
  }
  if (!aid) {
    return { status: 412, body: {
      error: 'Validate entered lots first',
      detail: 'Pick the specific trade (ANO + date) so its entered lots can be validated before price import.',
      gate: 'lot_validation',
    } };
  }
  if (lvGateState(db, aid) === 'clean') return null;
  return { status: 412, body: {
    error: 'Validate entered lots first',
    detail: 'Open Lots → Validate Entered Lots, resolve all errors and acknowledge the warnings, then import prices.',
    auctionId: aid,
    gate: 'lot_validation',
  } };
}

// Express middleware: block PRICE IMPORT until the trade's entered lots
// have been validated. MUST be mounted AFTER upload.single so req.body
// is populated.
function requireLotsValidatedForPriceImport(req, res, next) {
  const block = lotValidationGateBlock(getDb(), req.body || {});
  if (!block) return next();
  return res.status(block.status).json(block.body);
}

// ══════════════════════════════════════════════════════════════
// AUCTIONS
// ══════════════════════════════════════════════════════════════
app.get('/api/auctions', requireViewOrLotEntry, (req, res) => {
  const db = getDb();
  const search   = String(req.query.search || '').trim();
  const wantPaged = req.query.page != null || req.query.pageSize != null;

  // Search filter — single-token match against the ano column. Trades
  // are normally identified by trade number, so name/state/crop searches
  // weren't valuable here. Add OR clauses if that ever changes.
  let where = '';
  let params = [];
  if (search) {
    where = 'WHERE ano LIKE ?';
    params = [`%${search}%`];
  }

  // Lot count is computed as a correlated subquery so each row gets its
  // current count without a separate round-trip. SQLite handles this in
  // a single query plan.
  const sel = `SELECT *, (SELECT COUNT(*) FROM lots WHERE auction_id=auctions.id) as lot_count
               FROM auctions ${where}
               ORDER BY date DESC, ano DESC`;

  if (wantPaged) {
    const page     = Math.max(1, parseInt(req.query.page, 10) || 1);
    const pageSize = Math.min(500, Math.max(1, parseInt(req.query.pageSize, 10) || 50));
    const offset   = (page - 1) * pageSize;
    const total = db.get(`SELECT COUNT(*) as c FROM auctions ${where}`, params).c;
    const rows = db.all(sel + ' LIMIT ? OFFSET ?', [...params, pageSize, offset]);
    return res.json({ rows: withFmtDate(rows), total, page, pageSize });
  }

  // Legacy plain-array response — capped at 100 like before.
  const rows = db.all(sel + ' LIMIT 100', params);
  res.json(withFmtDate(rows));
});
app.post('/api/auctions', requireAuctionWrite, (req, res) => {
  const { ano, date, crop_type, state } = req.body;
  const db = getDb();
  const d = normalizeDate(date);
  const defaultCrop  = getSetting(db, 'default_crop_type') || 'VST';
  const defaultState = getSetting(db, 'business_state')    || 'TAMIL NADU';
  db.run('INSERT INTO auctions (ano,date,crop_type,state) VALUES (?,?,?,?)', [ano, d, crop_type||defaultCrop, state||defaultState]);
  const created = db.get('SELECT id FROM auctions WHERE ano = ? AND date = ? ORDER BY id DESC LIMIT 1', [ano, d]);
  res.json({ success: true, id: created ? created.id : null });
});
// Re-sync the denormalized trade-number (`ano`) copies after a trade is
// renumbered. The dependent rows carry `ano` as a denormalized copy of the
// trade number set at creation time; editing a trade's `ano` left them stale,
// which orphaned a trade's purchases / invoices / bills / debit notes under
// the OLD number. The visible symptom was "Generate Debit Notes" reporting
// "no eligible purchases" because the lookup is by the new `ano`.
//
// purchases / invoices / bills are bridged by the STABLE auction_id. debit_notes
// have no auction_id, so they're bridged via the OLD ano still stamped on this
// trade's purchases, scoped by dealer name so a different trade that happens to
// share the old number isn't touched. Idempotent — safe on every save.
function _resyncTradeAno(db, auctionId, currentAno) {
  const aid = Number(auctionId);
  const cur = String(currentAno == null ? '' : currentAno).trim();
  if (!aid || cur === '') return;
  try {
    // Old numbers still stamped on this trade's id-linked rows (pre-heal).
    const oldAnos = new Set();
    for (const tbl of ['purchases', 'invoices', 'bills']) {
      try {
        for (const r of db.all(
          `SELECT DISTINCT ano FROM ${tbl} WHERE auction_id = ? AND ano IS NOT NULL AND TRIM(ano) <> '' AND ano <> ?`,
          [aid, cur])) oldAnos.add(String(r.ano));
      } catch (_) { /* table/column may be absent on a partial migration */ }
    }
    if (oldAnos.size) {
      // Dealer names on this trade — scopes the debit_notes bridge.
      const names = [];
      try {
        for (const r of db.all('SELECT DISTINCT name FROM purchases WHERE auction_id = ?', [aid])) {
          if (r.name != null && String(r.name).trim() !== '') names.push(String(r.name));
        }
      } catch (_) {}
      if (names.length) {
        const ph = names.map(() => '?').join(',');
        for (const old of oldAnos) {
          try { db.run(`UPDATE debit_notes SET ano = ? WHERE ano = ? AND name IN (${ph})`, [cur, old, ...names]); } catch (_) {}
        }
      }
    }
    // Re-sync the id-linked tables to the current trade number.
    for (const tbl of ['purchases', 'invoices', 'bills']) {
      try { db.run(`UPDATE ${tbl} SET ano = ? WHERE auction_id = ?`, [cur, aid]); } catch (_) {}
    }
  } catch (e) { console.warn('[resync trade ano]', e.message); }
}

app.put('/api/auctions/:id', requireAuctionWrite, (req, res) => {
  const { ano, date, crop_type, state } = req.body;
  const db = getDb();
  const defaultCrop  = getSetting(db, 'default_crop_type') || 'VST';
  const defaultState = getSetting(db, 'business_state')    || 'TAMIL NADU';
  db.run('UPDATE auctions SET ano=?, date=?, crop_type=?, state=? WHERE id=?',
    [ano, normalizeDate(date), crop_type||defaultCrop, state||defaultState, req.params.id]);
  // Cascade the (possibly changed) trade number to dependent rows so a rename
  // never orphans this trade's purchases / debit notes. Also heals any trade
  // already desynced by a rename made before this cascade existed.
  _resyncTradeAno(db, req.params.id, ano);
  res.json({ success: true });
});
app.delete('/api/auctions/:id', requireDelete, (req, res) => {
  const db = getDb();
  // Cascade-delete every record that belongs to this trade. The
  // dependent tables reference the trade either by `auction_id` (FK,
  // newer rows) or by `ano` (legacy / debit_notes, which never had a
  // FK). We resolve `ano` once up-front so both link styles are
  // covered consistently. Wrapped in a transaction so a partial
  // failure (e.g. one table missing on a not-yet-migrated DB) doesn't
  // leave orphan rows behind.
  const auction = db.get('SELECT id, ano FROM auctions WHERE id = ?', [req.params.id]);
  if (!auction) return res.status(404).json({ error: 'Trade not found' });
  const ano = String(auction.ano || '').trim();

  const removed = { lots: 0, lot_allocations: 0, invoices: 0, purchases: 0, bills: 0, debit_notes: 0 };
  const cnt = (sql, params) => {
    try { const r = db.get(sql, params); return Number(r && (r.c ?? r.count ?? 0)) || 0; } catch (_) { return 0; }
  };
  // Pre-count for the response payload so the UI can show "Deleted N
  // invoices, M purchases, …".
  removed.lots            = cnt('SELECT COUNT(*) AS c FROM lots            WHERE auction_id = ?', [auction.id]);
  removed.lot_allocations = cnt('SELECT COUNT(*) AS c FROM lot_allocations WHERE auction_id = ?', [auction.id]);
  removed.invoices        = cnt('SELECT COUNT(*) AS c FROM invoices        WHERE auction_id = ? OR ano = ?', [auction.id, ano]);
  removed.purchases       = cnt('SELECT COUNT(*) AS c FROM purchases       WHERE auction_id = ? OR ano = ?', [auction.id, ano]);
  removed.bills           = cnt('SELECT COUNT(*) AS c FROM bills           WHERE auction_id = ? OR ano = ?', [auction.id, ano]);
  removed.debit_notes     = cnt('SELECT COUNT(*) AS c FROM debit_notes     WHERE ano = ?',                  [ano]);

  // Actual cascade. Each statement is wrapped in try/catch so a stray
  // missing column on a partially-migrated DB doesn't abort the rest;
  // any failure is logged but doesn't block the trade row's deletion.
  const safeRun = (sql, params) => { try { db.run(sql, params); } catch (e) { console.warn('[delete trade]', e.message); } };
  safeRun('DELETE FROM lot_allocations WHERE auction_id = ?', [auction.id]);
  safeRun('DELETE FROM lots            WHERE auction_id = ?', [auction.id]);
  safeRun('DELETE FROM invoices        WHERE auction_id = ? OR ano = ?', [auction.id, ano]);
  safeRun('DELETE FROM purchases       WHERE auction_id = ? OR ano = ?', [auction.id, ano]);
  safeRun('DELETE FROM bills           WHERE auction_id = ? OR ano = ?', [auction.id, ano]);
  safeRun('DELETE FROM debit_notes     WHERE ano = ?',                  [ano]);
  safeRun('DELETE FROM auctions        WHERE id  = ?', [auction.id]);

  // Payments are NOT a stored entity — they're computed live from the
  // lots table by getPaymentSummary, so once lots are gone the
  // Payments tab automatically returns no rows for this trade. No
  // separate delete needed.
  res.json({ success: true, deleted: removed });
});

// ══════════════════════════════════════════════════════════════
// LOT ALLOCATIONS — per-trade, per-branch lot-number ranges
// ══════════════════════════════════════════════════════════════
// Ported from spice-auction-pwa. Each allocation row reserves a
// contiguous range of lot numbers (e.g. "001"–"080" or "A001"–"A080")
// for one branch within one trade. Lot Entry uses these ranges to:
//   1. validate every saved lot's lot_no falls inside its branch's range
//   2. suggest the next free lot_no when picking a seller
//   3. show used/free progress per branch on the entry screen
//
// Lot numbers are parsed as `[A-Za-z]*\d+` so both pure-numeric and
// prefixed schemes work. Padding length is preserved when generating
// candidate lot numbers from a range.
// ──────────────────────────────────────────────────────────────

function parseLotNo(lot) {
  const match = String(lot).match(/^([A-Za-z]*)(\d+)$/);
  if (!match) return null;
  return { prefix: match[1].toUpperCase(), num: parseInt(match[2], 10), padLen: match[2].length };
}

function buildLotNo(prefix, num, padLen) {
  return prefix + String(num).padStart(padLen, '0');
}

function isLotInRange(lotNo, startLot, endLot) {
  const lot = parseLotNo(lotNo);
  const s = parseLotNo(startLot);
  const e = parseLotNo(endLot);
  if (!lot || !s || !e) return false;
  if (lot.prefix !== s.prefix || s.prefix !== e.prefix) return false;
  return lot.num >= s.num && lot.num <= e.num;
}

function rangeSize(startLot, endLot) {
  const s = parseLotNo(startLot);
  const e = parseLotNo(endLot);
  if (!s || !e) return 0;
  return e.num - s.num + 1;
}

// Get allocations for a trade
app.get('/api/auctions/:id/allocations', requireViewOrLotEntry, (req, res) => {
  const db = getDb();
  const auctionId = parseInt(req.params.id, 10);
  const allocations = db.all(
    'SELECT * FROM lot_allocations WHERE auction_id = ? ORDER BY branch, start_lot',
    [auctionId]
  );
  res.json({ allocations });
});

// Save allocations for a trade (bulk replace)
// Validates: required fields, parseable formats, matching prefixes, no
// overlapping ranges, and refuses to drop an existing allocation whose
// lot range still has lots in the lots table.
app.post('/api/auctions/:id/allocations', requireAuctionWrite, (req, res) => {
  const db = getDb();
  const auctionId = parseInt(req.params.id, 10);
  const { allocations } = req.body;
  if (!allocations || !Array.isArray(allocations) || !allocations.length) {
    return res.status(400).json({ error: 'At least one allocation is required' });
  }

  for (const a of allocations) {
    if (!a.branch || !a.start_lot || !a.end_lot) {
      return res.status(400).json({ error: 'Branch, start_lot, end_lot required for each allocation' });
    }
    const s = parseLotNo(a.start_lot);
    const e = parseLotNo(a.end_lot);
    if (!s || !e) return res.status(400).json({ error: `Invalid lot format: ${a.start_lot} or ${a.end_lot}. Use format like 001, A001` });
    if (s.prefix !== e.prefix) return res.status(400).json({ error: `Prefix mismatch: ${a.start_lot} vs ${a.end_lot}` });
    if (s.num > e.num) return res.status(400).json({ error: `Start (${a.start_lot}) must be <= End (${a.end_lot})` });
  }

  for (let i = 0; i < allocations.length; i++) {
    for (let j = i + 1; j < allocations.length; j++) {
      const a = allocations[i], b = allocations[j];
      const ap = parseLotNo(a.start_lot), ae = parseLotNo(a.end_lot);
      const bp = parseLotNo(b.start_lot), be = parseLotNo(b.end_lot);
      if (ap.prefix === bp.prefix && ap.num <= be.num && bp.num <= ae.num) {
        return res.status(400).json({ error: `Ranges overlap: ${a.branch} (${a.start_lot}-${a.end_lot}) and ${b.branch} (${b.start_lot}-${b.end_lot})` });
      }
    }
  }

  // Safety: a save must not strip coverage from a lot that is currently
  // covered. We compute orphans BEFORE and AFTER the proposed change.
  // Only lots that become NEWLY orphaned (covered now, uncovered after)
  // block the save. Lots already orphaned before the save — typically
  // the import-before-allocate case where lots entered the DB without
  // any matching allocation — stay orphaned and don't block; the user
  // can fix those separately by adding ranges that cover them.
  const existing = db.all('SELECT * FROM lot_allocations WHERE auction_id = ?', [auctionId]);
  const usedLots = db.all('SELECT lot_no, branch FROM lots WHERE auction_id = ?', [auctionId]);
  const orphansBefore = new Set();
  for (const ul of usedLots) {
    const coveredBefore = existing.some(a => isLotInRange(ul.lot_no, a.start_lot, a.end_lot));
    if (!coveredBefore) orphansBefore.add(ul.lot_no);
  }
  const newlyOrphaned = [];
  for (const ul of usedLots) {
    const coveredAfter = allocations.some(a => isLotInRange(ul.lot_no, a.start_lot, a.end_lot));
    if (!coveredAfter && !orphansBefore.has(ul.lot_no)) newlyOrphaned.push(ul.lot_no);
  }
  if (newlyOrphaned.length > 0) {
    return res.status(400).json({
      error: `Cannot save — ${newlyOrphaned.length} lot${newlyOrphaned.length === 1 ? '' : 's'} that ${newlyOrphaned.length === 1 ? 'is' : 'are'} currently covered would lose ${newlyOrphaned.length === 1 ? 'its' : 'their'} branch assignment: ${newlyOrphaned.slice(0, 5).join(', ')}${newlyOrphaned.length > 5 ? '…' : ''}. Adjust the ranges so they still include ${newlyOrphaned.length === 1 ? 'this lot' : 'these lots'}.`
    });
  }

  db.run('DELETE FROM lot_allocations WHERE auction_id = ?', [auctionId]);
  for (const a of allocations) {
    db.run(
      'INSERT INTO lot_allocations (auction_id, branch, start_lot, end_lot) VALUES (?, ?, ?, ?)',
      [auctionId, a.branch, String(a.start_lot).trim(), String(a.end_lot).trim()]
    );
  }

  const saved = db.all(
    'SELECT * FROM lot_allocations WHERE auction_id = ? ORDER BY branch, start_lot',
    [auctionId]
  );
  res.json({ allocations: saved });
});

// ══════════════════════════════════════════════════════════════
// GENERATION LOCK — once EVERY eligible party in a trade has its
// invoice/purchase/bill/debit-note, the matching "Generate" actions
// are blocked. While at least one party is still un-documented, the
// operator can keep generating (single OR bulk) freely.
//
// Admin can grant ONE regeneration at a time by inserting a row
// into generation_overrides; the row is consumed (deleted) the first
// time the corresponding generate endpoint runs.
//
// Doc types: 'invoices' | 'purchases' | 'bills' | 'debit_notes'.
// Payments aren't generated (they're a derived view of lots) so
// they're outside this gate.
// ══════════════════════════════════════════════════════════════
const _GEN_TABLE = {
  invoices:    'invoices',
  purchases:   'purchases',
  bills:       'bills',
  debit_notes: 'debit_notes',
};
function _hasGeneratedDocs(db, docType, auctionId) {
  const table = _GEN_TABLE[docType];
  if (!table || !auctionId) return false;
  return !!db.get(`SELECT 1 FROM ${table} WHERE auction_id = ? LIMIT 1`, [auctionId]);
}
// True iff at least one eligible party in the trade is still missing
// its doc. Mirrors each generate-all endpoint's "what's left to do"
// query, so the gate engages exactly when those endpoints would
// return "nothing to do".
//
// Withdrawn lots (code = 'WD') are excluded — they don't need any
// invoice / purchase / bill, so a party whose only lots are WD must
// not count as "remaining work" or the gate would never engage.
const _NOT_WD = `UPPER(TRIM(COALESCE(l.code,''))) != 'WD'`;
function _hasRemainingParties(db, docType, auctionId) {
  if (!auctionId) return false;
  if (docType === 'invoices') {
    const cfg = getSettingsFlat(db);
    const isASPState = String(cfg.business_state || '').toUpperCase() === 'KERALA';
    const uninvoicedExpr = isASPState
      ? `(l.invo IS NULL OR l.invo = '')`
      : `(l.invo IS NULL OR l.invo = '' OR (l.asp_invo IS NOT NULL AND l.asp_invo != '' AND l.invo = l.asp_invo))`;
    return !!db.get(
      `SELECT 1 FROM lots l
       WHERE l.auction_id = ? AND l.buyer IS NOT NULL AND l.buyer != ''
         AND l.amount > 0 AND l.locked_at IS NULL AND ${_NOT_WD}
         AND ${uninvoicedExpr}
       LIMIT 1`,
      [auctionId]
    );
  }
  if (docType === 'purchases') {
    return !!db.get(
      `SELECT 1 FROM lots l
       WHERE l.auction_id = ? AND l.amount > 0 AND ${_NOT_WD}
         AND l.name IS NOT NULL AND l.name != ''
         AND (UPPER(l.cr) LIKE 'GSTIN%' OR (l.cr GLOB '[0-9][0-9]*' AND LENGTH(l.cr) >= 15))
         AND NOT EXISTS (
           SELECT 1 FROM purchases p WHERE p.auction_id = l.auction_id AND p.name = l.name
         )
       LIMIT 1`,
      [auctionId]
    );
  }
  if (docType === 'bills') {
    return !!db.get(
      `SELECT 1 FROM lots l
       WHERE l.auction_id = ? AND l.amount > 0 AND ${_NOT_WD}
         AND l.name IS NOT NULL AND l.name != ''
         AND (l.cr IS NULL OR l.cr = ''
              OR (UPPER(l.cr) NOT LIKE 'GSTIN%' AND l.cr NOT GLOB '[0-9][0-9]*'))
         AND NOT EXISTS (
           SELECT 1 FROM bills b WHERE b.auction_id = l.auction_id AND b.name = l.name
         )
       LIMIT 1`,
      [auctionId]
    );
  }
  if (docType === 'debit_notes') {
    return !!db.get(
      `SELECT 1 FROM purchases p
       WHERE p.auction_id = ?
         AND NOT EXISTS (
           SELECT 1 FROM debit_notes d WHERE d.auction_id = p.auction_id AND d.name = p.name
         )
       LIMIT 1`,
      [auctionId]
    );
  }
  return false;
}
// "Fully generated" = at least one doc exists AND no remaining party.
// Empty trades report false (nothing to lock against).
function _isFullyGenerated(db, docType, auctionId) {
  if (!_hasGeneratedDocs(db, docType, auctionId)) return false;
  return !_hasRemainingParties(db, docType, auctionId);
}
function _getGenerationOverride(db, docType, auctionId) {
  if (!auctionId) return null;
  return db.get(
    'SELECT auction_id, doc_type, granted_at, granted_by FROM generation_overrides WHERE auction_id = ? AND doc_type = ?',
    [auctionId, docType]
  );
}
function _consumeGenerationOverride(db, docType, auctionId) {
  if (!auctionId) return;
  db.run('DELETE FROM generation_overrides WHERE auction_id = ? AND doc_type = ?', [auctionId, docType]);
}
// Pre-flight gate check used at the top of every generate endpoint.
// Returns { allowed: true } when (a) at least one party still needs
// a doc, or (b) an admin override is present (consumed here, one-shot).
// Returns { allowed: false, error } with a 412 payload to send.
function _checkGenerationGate(db, docType, auctionId) {
  if (!auctionId) return { allowed: true };
  if (!_isFullyGenerated(db, docType, auctionId)) return { allowed: true };
  const override = _getGenerationOverride(db, docType, auctionId);
  if (override) {
    _consumeGenerationOverride(db, docType, auctionId);
    return { allowed: true, usedOverride: true };
  }
  const labels = {
    invoices: 'Invoices', purchases: 'Purchases',
    bills: 'Bills of Supply', debit_notes: 'Debit Notes',
  };
  const label = labels[docType] || docType;
  return {
    allowed: false,
    error: {
      error: 'Generation locked',
      detail: `${label} have already been generated for every party in this trade. An admin must click 🔓 Allow regeneration before another generate can run.`,
      auctionId, docType, gate: 'generation',
    },
  };
}

// Status endpoint — drives the client's button-disable state.
// `has` is true only when EVERY eligible party has its doc (= the
// gate is engaged). Field name kept for client compatibility.
app.get('/api/auctions/:id/generation-status', requireView, (req, res) => {
  const db = getDb();
  const auctionId = parseInt(req.params.id, 10);
  if (!auctionId) return res.status(400).json({ error: 'Invalid auction id' });
  const out = {};
  for (const docType of Object.keys(_GEN_TABLE)) {
    out[docType] = {
      has: _isFullyGenerated(db, docType, auctionId),
      override: !!_getGenerationOverride(db, docType, auctionId),
    };
  }
  res.json({ auctionId, ...out });
});

// Admin: grant one regeneration for (auction, doc_type). Idempotent —
// re-granting before consumption is a no-op (PRIMARY KEY collision is
// handled by INSERT OR REPLACE).
app.post('/api/auctions/:id/generation-override', requireAdmin, (req, res) => {
  const db = getDb();
  const auctionId = parseInt(req.params.id, 10);
  const docType = String(req.body.docType || '').trim();
  if (!auctionId) return res.status(400).json({ error: 'Invalid auction id' });
  if (!_GEN_TABLE[docType]) return res.status(400).json({ error: `Invalid docType: ${docType}` });
  const grantedBy = (req.user && (req.user.username || req.user.name)) || 'admin';
  db.run(
    `INSERT OR REPLACE INTO generation_overrides
       (auction_id, doc_type, granted_at, granted_by)
       VALUES (?, ?, datetime('now','localtime'), ?)`,
    [auctionId, docType, grantedBy]
  );
  res.json({ ok: true, auctionId, docType, grantedBy });
});

// Admin: revoke an unused override. Useful if admin clicks unlock by
// mistake and wants to take it back before the operator regenerates.
app.delete('/api/auctions/:id/generation-override', requireAdmin, (req, res) => {
  const db = getDb();
  const auctionId = parseInt(req.params.id, 10);
  const docType = String(req.body.docType || '').trim();
  if (!auctionId) return res.status(400).json({ error: 'Invalid auction id' });
  if (!_GEN_TABLE[docType]) return res.status(400).json({ error: `Invalid docType: ${docType}` });
  _consumeGenerationOverride(db, docType, auctionId);
  res.json({ ok: true, auctionId, docType });
});

// Auto-fill allocations from existing lots. Walks every lot in the
// auction, groups uncovered ones by (branch, prefix), and APPENDS one
// new allocation per group covering 001..999 (or higher if any
// uncovered lot exceeds that). Existing allocations are kept as-is so
// the operator's manual setup isn't overwritten.
//
// Idempotent: if every lot is already covered, returns the existing
// list with `created: 0`. Useful as a one-click rescue for trades that
// arrived with imported lots and no allocations — the same fill the
// import endpoint now applies, available after the fact for older
// data.
app.post('/api/auctions/:id/allocations/auto-fill', requireAuctionWrite, (req, res) => {
  const db = getDb();
  const auctionId = parseInt(req.params.id, 10);
  if (!auctionId) return res.status(400).json({ error: 'Invalid auction id' });
  const auc = db.get('SELECT id FROM auctions WHERE id = ?', [auctionId]);
  if (!auc) return res.status(404).json({ error: 'Auction not found' });

  const existing = db.all('SELECT * FROM lot_allocations WHERE auction_id = ?', [auctionId]);
  const lots = db.all('SELECT lot_no, branch FROM lots WHERE auction_id = ?', [auctionId]);

  // Same grouping logic as the post-import auto-allocate so the two
  // paths converge on identical ranges.
  const groups = new Map(); // key = `${branch}||${prefix}` → { branch, prefix, padLen, max }
  for (const l of lots) {
    const covered = existing.some(a => isLotInRange(l.lot_no, a.start_lot, a.end_lot));
    if (covered) continue;
    const p = parseLotNo(l.lot_no);
    if (!p) continue;
    const branch = String(l.branch || '').trim();
    const key = `${branch}||${p.prefix}`;
    const g = groups.get(key);
    if (!g) groups.set(key, { branch, prefix: p.prefix, padLen: p.padLen, max: p.num });
    else if (p.num > g.max) g.max = p.num;
  }

  let created = 0;
  for (const g of groups.values()) {
    const end = g.max > 999 ? 9999 : 999;
    const padLen = Math.max(g.padLen, String(end).length);
    const startLot = buildLotNo(g.prefix, 1, padLen);
    const endLot   = buildLotNo(g.prefix, end, padLen);
    db.run(
      'INSERT INTO lot_allocations (auction_id, branch, start_lot, end_lot) VALUES (?, ?, ?, ?)',
      [auctionId, g.branch, startLot, endLot]
    );
    created++;
  }

  const allocations = db.all(
    'SELECT * FROM lot_allocations WHERE auction_id = ? ORDER BY branch, start_lot',
    [auctionId]
  );
  res.json({ created, allocations });
});

// Allocation stats (used/total per branch + per-lot grid)
// Drives both the admin Allocations modal and the Lot Entry status bar.
app.get('/api/auctions/:id/allocation-stats', requireViewOrLotEntry, (req, res) => {
  const db = getDb();
  const auctionId = parseInt(req.params.id, 10);
  const allocations = db.all(
    'SELECT * FROM lot_allocations WHERE auction_id = ? ORDER BY branch, start_lot',
    [auctionId]
  );
  // Pull seller + amount alongside lot_no so the tile UI can show who
  // booked each lot and color "booked" (amount > 0) vs "allocated"
  // (entered but not booked) differently. Ports the per-tile state
  // model from spice-config-merged-with-pwa source.
  const lots = db.all(
    'SELECT lot_no, branch, name, amount FROM lots WHERE auction_id = ?',
    [auctionId]
  );
  // lot_no → { branch, seller, booked } — O(1) lookup while building the grid.
  // Key by a PADDING-NORMALISED form (prefix + numeric value), NOT the raw
  // stored lot_no. The grid below generates lot numbers via buildLotNo()
  // padded to the allocation's width (e.g. "005"), but a lot may be stored
  // unpadded ("5") or with different padding. A raw-string key then missed
  // the match, so a booked/entered lot showed up as a FREE, selectable chip
  // in the "Select Lot" grid. isLotInRange (used for the count) already
  // compares numerically, so this aligns the per-tile flags with the count.
  const lotKey = (lotNo) => {
    const p = parseLotNo(lotNo);
    return p ? p.prefix + ':' + p.num : String(lotNo).trim().toUpperCase();
  };
  const lotInfo = {};
  for (const l of lots) {
    lotInfo[lotKey(l.lot_no)] = {
      branch: l.branch || '',
      seller: l.name   || '',
      booked: Number(l.amount) > 0,
    };
  }

  const stats = {};
  for (const a of allocations) {
    if (!stats[a.branch]) stats[a.branch] = { branch: a.branch, total: 0, used: 0, ranges: [] };
    const total = rangeSize(a.start_lot, a.end_lot);
    const usedInRange = lots.filter(l => isLotInRange(l.lot_no, a.start_lot, a.end_lot));
    stats[a.branch].total += total;
    stats[a.branch].used += usedInRange.length;

    const s = parseLotNo(a.start_lot);
    const e = parseLotNo(a.end_lot);
    const lotGrid = [];
    if (s && e) {
      for (let n = s.num; n <= e.num; n++) {
        const lotNo = buildLotNo(s.prefix, n, s.padLen);
        const info = lotInfo[lotKey(lotNo)];
        // State machine:
        //   booked    — present in lots table AND has a sale amount
        //               (a real bid landed on this lot — can't reassign)
        //   allocated — present in lots table, amount=0
        //               (lot row created but not yet sold — still locked
        //                because the field user has committed this number)
        //   free      — not in lots table at all (safe to reassign / delete)
        let state = 'free';
        if (info && info.booked) state = 'booked';
        else if (info)           state = 'allocated';
        lotGrid.push({
          lot: lotNo,
          used: !!info,
          booked: !!(info && info.booked),
          seller: info ? info.seller : '',
          branch: a.branch,
          state,
        });
      }
    }
    stats[a.branch].ranges.push({
      start: a.start_lot, end: a.end_lot, total,
      used: usedInRange.length, lots: lotGrid
    });
  }

  res.json({ stats: Object.values(stats), allocations });
});

// Reassign an unused range from one branch to another
// Splits the source allocations around the reassigned range, then
// inserts a single allocation covering the same range under the
// destination branch. Refuses to act if any lot in the range is already
// saved (lots can only be reassigned by deleting + re-entering).
app.post('/api/auctions/:id/reassign-lots', requireAuctionWrite, (req, res) => {
  const db = getDb();
  const auctionId = parseInt(req.params.id, 10);
  const { from_branch, to_branch, start_lot, end_lot } = req.body;

  if (!from_branch || !to_branch || !start_lot || !end_lot) {
    return res.status(400).json({ error: 'All fields required: from_branch, to_branch, start_lot, end_lot' });
  }
  if (from_branch === to_branch) return res.status(400).json({ error: 'FROM and TO branch must be different' });

  const s = parseLotNo(start_lot);
  const e = parseLotNo(end_lot);
  if (!s || !e) return res.status(400).json({ error: 'Invalid lot number format' });
  if (s.prefix !== e.prefix) return res.status(400).json({ error: 'Start and end must have same prefix' });
  if (s.num > e.num) return res.status(400).json({ error: 'Start must be <= End' });

  // Every lot in the range must currently belong to from_branch
  const fromAllocs = db.all('SELECT * FROM lot_allocations WHERE auction_id = ? AND branch = ?', [auctionId, from_branch]);
  for (let n = s.num; n <= e.num; n++) {
    const lotNo = buildLotNo(s.prefix, n, s.padLen);
    const inRange = fromAllocs.some(a => isLotInRange(lotNo, a.start_lot, a.end_lot));
    if (!inRange) return res.status(400).json({ error: `Lot ${lotNo} is not allocated to ${from_branch}` });
  }

  // None of the lots may already be saved
  const usedLots = db.all('SELECT lot_no FROM lots WHERE auction_id = ?', [auctionId]).map(l => l.lot_no);
  const usedSet = new Set(usedLots);
  const usedInRange = [];
  for (let n = s.num; n <= e.num; n++) {
    const lotNo = buildLotNo(s.prefix, n, s.padLen);
    if (usedSet.has(lotNo)) usedInRange.push(lotNo);
  }
  if (usedInRange.length > 0) {
    return res.status(400).json({
      error: `Cannot reassign — ${usedInRange.length} lot(s) already used: ${usedInRange.slice(0, 5).join(', ')}${usedInRange.length > 5 ? '...' : ''}`
    });
  }

  // Rebuild from_branch allocations excluding the reassigned range
  const fromAllocsAll = db.all('SELECT * FROM lot_allocations WHERE auction_id = ? AND branch = ?', [auctionId, from_branch]);
  db.run('DELETE FROM lot_allocations WHERE auction_id = ? AND branch = ?', [auctionId, from_branch]);

  for (const alloc of fromAllocsAll) {
    const as = parseLotNo(alloc.start_lot);
    const ae = parseLotNo(alloc.end_lot);
    if (!as || !ae) continue;
    const overlapStart = Math.max(as.num, s.num);
    const overlapEnd = Math.min(ae.num, e.num);

    if (overlapStart > overlapEnd) {
      // No overlap — keep entire allocation
      db.run('INSERT INTO lot_allocations (auction_id, branch, start_lot, end_lot) VALUES (?, ?, ?, ?)',
        [auctionId, from_branch, alloc.start_lot, alloc.end_lot]);
    } else {
      // Has overlap — keep slices before/after the reassigned range
      if (as.num < overlapStart) {
        db.run('INSERT INTO lot_allocations (auction_id, branch, start_lot, end_lot) VALUES (?, ?, ?, ?)',
          [auctionId, from_branch, buildLotNo(as.prefix, as.num, as.padLen), buildLotNo(as.prefix, overlapStart - 1, as.padLen)]);
      }
      if (ae.num > overlapEnd) {
        db.run('INSERT INTO lot_allocations (auction_id, branch, start_lot, end_lot) VALUES (?, ?, ?, ?)',
          [auctionId, from_branch, buildLotNo(ae.prefix, overlapEnd + 1, ae.padLen), buildLotNo(ae.prefix, ae.num, ae.padLen)]);
      }
    }
  }

  db.run('INSERT INTO lot_allocations (auction_id, branch, start_lot, end_lot) VALUES (?, ?, ?, ?)',
    [auctionId, to_branch, String(start_lot).trim(), String(end_lot).trim()]);

  const allocs = db.all('SELECT * FROM lot_allocations WHERE auction_id = ? ORDER BY branch, start_lot', [auctionId]);
  res.json({ success: true, allocations: allocs, message: `Lots ${start_lot}-${end_lot} reassigned from ${from_branch} to ${to_branch}` });
});

// Validate a single lot number against (a) duplicates and (b) the
// branch's allocation. Returns { valid: bool, error?: string }.
app.get('/api/auctions/:id/validate-lot', requireViewOrLotEntry, (req, res) => {
  const db = getDb();
  const auctionId = parseInt(req.params.id, 10);
  const lotNo = String(req.query.lot_no || '').trim();
  const branch = String(req.query.branch || '').trim();
  if (!lotNo) return res.json({ valid: false, error: 'Enter lot number' });

  const dup = db.get('SELECT id FROM lots WHERE auction_id = ? AND lot_no = ?', [auctionId, lotNo]);
  if (dup) return res.json({ valid: false, error: 'Lot #' + lotNo + ' already exists' });

  const allocs = db.all('SELECT * FROM lot_allocations WHERE auction_id = ? AND branch = ?', [auctionId, branch]);
  if (allocs.length > 0) {
    const inRange = allocs.some(a => isLotInRange(lotNo, a.start_lot, a.end_lot));
    if (!inRange) {
      const ranges = allocs.map(a => a.start_lot + '-' + a.end_lot).join(', ');
      return res.json({ valid: false, error: 'Outside allocation (' + ranges + ')' });
    }
  }
  res.json({ valid: true });
});

// Next available lot number for a branch — used by Lot Entry to
// auto-suggest after every save and after a seller is picked.
app.get('/api/auctions/:id/next-lot/:branch', requireViewOrLotEntry, (req, res) => {
  const db = getDb();
  const auctionId = parseInt(req.params.id, 10);
  const branch = req.params.branch;
  const allocations = db.all(
    'SELECT * FROM lot_allocations WHERE auction_id = ? AND branch = ? ORDER BY start_lot',
    [auctionId, branch]
  );
  if (!allocations.length) return res.json({ next_lot: null, error: 'No allocation for this branch' });

  const usedLots = db.all('SELECT lot_no FROM lots WHERE auction_id = ?', [auctionId]).map(l => l.lot_no);
  const usedSet = new Set(usedLots);

  for (const a of allocations) {
    const s = parseLotNo(a.start_lot);
    const e = parseLotNo(a.end_lot);
    if (!s || !e) continue;
    for (let n = s.num; n <= e.num; n++) {
      const lotNo = buildLotNo(s.prefix, n, s.padLen);
      if (!usedSet.has(lotNo)) return res.json({ next_lot: lotNo });
    }
  }
  res.json({ next_lot: null, error: 'All lots in this branch are used' });
});

// ── Import Auction + Lots from XLS/XLSX (replaces APPA.PRG) ──
// Core lot/price import — shared by the manual upload route
// (/api/auctions/import) and the Trade Fair sync (/api/trade-fair/import).
// Reads the .xls/.xlsx at `filePath`, applies mode 'full' (insert new
// lots) or 'price' (update price/qty/code/buyer on existing lots), and
// returns the summary object. Does NOT delete the file or touch any HTTP
// response — callers own those.
function runLotImport(db, filePath, body) {
    body = body || {};
    const workbook = XLSX.readFile(filePath);
    const ws = workbook.Sheets[workbook.SheetNames[0]];
    if (!ws) throw new Error('No worksheet found');
    const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
    if (!rows.length) throw new Error('File is empty');

    const mode = body.mode || 'full'; // 'full' = new lots, 'price' = update price/buyer only

    const mapCol = (row, ...names) => {
      for (const n of names) { if (row[n] !== undefined) return String(row[n]).trim(); }
      const keys = Object.keys(row);
      for (const n of names) {
        const found = keys.find(k => k.toUpperCase() === n.toUpperCase());
        if (found && row[found] !== undefined) return String(row[found]).trim();
      }
      return '';
    };
    const mapNum = (row, ...names) => parseFloat(mapCol(row, ...names)) || 0;

    // If user specified ano/date in the form → that OVERRIDES every row (single-auction import)
    // Otherwise → resolve auction per row from its own ANO/DATE columns (multi-auction import)
    const overrideAno = body.ano;
    const overrideDate = normalizeDate(body.date);
    const cropType = body.crop_type || mapCol(rows[0], 'CRPT', 'CROP_TYPE', 'CROPTYPE') || (getSetting(db, 'default_crop_type') || 'VST');
    const state = body.state || mapCol(rows[0], 'STATE') || 'TAMIL NADU';

    // Cache of resolved auctions so we don't query the DB for every row
    const auctionCache = new Map(); // key = "ano|date" → {id, ano, date}
    const resolveAuction = (ano, dateStr) => {
      const key = `${ano}|${dateStr}`;
      if (auctionCache.has(key)) return auctionCache.get(key);
      let auc = db.get('SELECT * FROM auctions WHERE ano = ? AND date = ?', [ano, dateStr]);
      if (!auc) {
        db.run('INSERT INTO auctions (ano, date, crop_type, state) VALUES (?,?,?,?)',
          [ano, dateStr || new Date().toISOString().slice(0, 10), cropType, state]);
        auc = db.get('SELECT * FROM auctions WHERE ano = ? AND date = ? ORDER BY id DESC LIMIT 1', [ano, dateStr]);
      }
      auctionCache.set(key, auc);
      return auc;
    };

    // Pre-validate: if no form override AND no ANO column anywhere, bail early with a clear message
    if (!overrideAno) {
      const firstAno = rows.length ? mapCol(rows[0], 'ANO', 'TNO', 'TRADE', 'TRADE_NO', 'TRADENO') : '';
      if (!firstAno) throw new Error('No ANO column found in file. Add ANO/TRADE/TRADE_NO column, or specify Trade No in the form to override.');
    }

    let imported = 0, updated = 0, skipped = 0;
    const skipReasons = []; // [{row, lot, reason}]
    const auctionStats = new Map(); // key = "ano|date" → count

    // Helper: check if row is completely empty (all values blank/undefined)
    const isBlankRow = (row) => {
      const vals = Object.values(row);
      return !vals.length || vals.every(v => v === '' || v === null || v === undefined);
    };

    if (mode === 'price') {
      // Price update mode — only update price, amount, code, buyer fields on existing lots
      for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        const rowNum = i + 2; // +2 because: 1-based + header row
        if (isBlankRow(row)) continue; // truly empty rows — don't count as skipped
        
        // Resolve this row's auction (form override OR read from row's ANO/DATE columns)
        const rowAno = overrideAno || mapCol(row, 'ANO', 'TNO', 'TRADE', 'TRADE_NO', 'TRADENO');
        // Read raw (un-stringified) DATE cell so Excel Date objects / serial numbers normalize correctly
        const rawDate = row.DATE !== undefined ? row.DATE
                      : row.date !== undefined ? row.date
                      : row.TRADE_DATE !== undefined ? row.TRADE_DATE
                      : '';
        const rowDate = overrideDate || normalizeDate(rawDate);
        if (!rowAno) { skipped++; skipReasons.push({row: rowNum, lot: '', reason: 'Missing ANO/TRADE_NO for this row'}); continue; }
        const auc = resolveAuction(rowAno, rowDate);
        const auctionId = auc.id;

        const lotNo = mapCol(row, 'LOT', 'LOT_NO', 'LOTNO');
        if (!lotNo) { skipped++; skipReasons.push({row: rowNum, lot: '', reason: 'Missing LOT / LOT_NO column value'}); continue; }

        // (price mode continues below in original code)
        const existing = db.get('SELECT id FROM lots WHERE auction_id = ? AND lot_no = ?', [auctionId, lotNo]);
        if (!existing) { skipped++; skipReasons.push({row: rowNum, lot: lotNo, reason: `Lot ${lotNo} does not exist in Trade ${rowAno} (price-update requires existing lot)`}); continue; }

        try {
          // Parse each field from the row using generous synonyms so different XLSX layouts work
          const price = mapNum(row, 'PRICE');
          const qty   = mapNum(row, 'QTY', 'QUANTITY', 'WEIGHT', 'WT');
          const bag   = mapNum(row, 'BAG', 'BAGS', 'NO_OF_BAGS');
          // If file didn't provide AMOUNT, compute qty × price (common in post-auction price sheets)
          let amount  = mapNum(row, 'AMOUNT', 'AMT', 'VALUE', 'TOTAL');
          if (!amount && qty && price) amount = qty * price;

          // Build UPDATE dynamically — only touch fields the file provided, so a sparse "price-only"
          // file doesn't wipe pre-existing bag/qty/buyer values
          const sets = []; const vals = [];
          if (row.PRICE !== undefined || row.price !== undefined) { sets.push('price=?');  vals.push(price); }
          if (amount)                                              { sets.push('amount=?'); vals.push(amount); }
          if (row.QTY !== undefined || row.qty !== undefined)      { sets.push('qty=?');    vals.push(qty); }
          if (row.BAG !== undefined || row.bag !== undefined ||
              row.BAGS !== undefined || row.bags !== undefined)    { sets.push('bags=?');   vals.push(bag); }
          const codeVal  = mapCol(row, 'CODE', 'BUYER_CODE');
          if (codeVal)                                             { sets.push('code=?');   vals.push(codeVal); }

          // Auto-resolve short CODE (e.g. RSH, TE, SL) to the full buyer record.
          // Priority: explicit BUYER/BIDDER column in file → matching buyers.code → matching buyers.ti → matching buyers.buyer
          let resolvedBuyer  = mapCol(row, 'BUYER', 'BIDDER', 'BUYER_NAME');
          let resolvedBuyer1 = mapCol(row, 'BUYER1', 'TRADE_NAME', 'TRADENAME');
          let resolvedSale   = mapCol(row, 'SALE', 'SALE_TYPE');

          if (codeVal && (!resolvedBuyer || !resolvedBuyer1)) {
            // Look the code up in the buyers master (case-insensitive match on code, ti, or buyer)
            const match = db.get(
              `SELECT buyer, buyer1, sale FROM buyers
               WHERE UPPER(TRIM(code))  = UPPER(TRIM(?))
                  OR UPPER(TRIM(ti))    = UPPER(TRIM(?))
                  OR UPPER(TRIM(buyer)) = UPPER(TRIM(?))
               LIMIT 1`,
              [codeVal, codeVal, codeVal]
            );
            if (match) {
              if (!resolvedBuyer)  resolvedBuyer  = match.buyer  || '';
              if (!resolvedBuyer1) resolvedBuyer1 = match.buyer1 || '';
              if (!resolvedSale)   resolvedSale   = match.sale   || '';
            } else {
              // Not found — record a warning but DON'T fail the row (we still update price/qty/bag)
              skipReasons.push({
                row: rowNum, lot: lotNo,
                reason: `Warning: CODE "${codeVal}" not found in Buyers master — price updated but buyer NOT assigned. Add this code to Buyers to enable invoicing.`
              });
            }
          }

          if (resolvedBuyer)  { sets.push('buyer=?');  vals.push(resolvedBuyer); }
          if (resolvedBuyer1) { sets.push('buyer1=?'); vals.push(resolvedBuyer1); }
          if (resolvedSale)   { sets.push('sale=?');   vals.push(resolvedSale); }

          if (!sets.length) { skipped++; skipReasons.push({row: rowNum, lot: lotNo, reason: 'Row has no updatable fields (price/qty/bag/code/buyer/sale)'}); continue; }

          vals.push(existing.id);
          db.run(`UPDATE lots SET ${sets.join(', ')} WHERE id=?`, vals);
          updated++;
          const key = `${rowAno}|${rowDate}`;
          auctionStats.set(key, (auctionStats.get(key) || 0) + 1);
        } catch (e) {
          skipped++;
          skipReasons.push({row: rowNum, lot: lotNo, reason: `DB error: ${e.message}`});
        }
      }
    } else {
      // Full import — insert new lots (skip if lot_no already exists for this auction)
      for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        const rowNum = i + 2;
        if (isBlankRow(row)) continue;
        
        // Resolve this row's auction (form override OR read from row's ANO/DATE columns)
        const rowAno = overrideAno || mapCol(row, 'ANO', 'TNO', 'TRADE', 'TRADE_NO', 'TRADENO');
        // Read raw DATE cell (may be Date object or Excel serial number) and normalize
        const rawDate = row.DATE !== undefined ? row.DATE
                      : row.date !== undefined ? row.date
                      : row.TRADE_DATE !== undefined ? row.TRADE_DATE
                      : '';
        const rowDate = overrideDate || normalizeDate(rawDate);
        if (!rowAno) { skipped++; skipReasons.push({row: rowNum, lot: '', reason: 'Missing ANO/TRADE_NO for this row'}); continue; }
        const auc = resolveAuction(rowAno, rowDate);
        const auctionId = auc.id;
        
        const lotNo = mapCol(row, 'LOT', 'LOT_NO', 'LOTNO');
        if (!lotNo) { skipped++; skipReasons.push({row: rowNum, lot: '', reason: 'Missing LOT / LOT_NO column value'}); continue; }

        const existing = db.get('SELECT id FROM lots WHERE auction_id = ? AND lot_no = ?', [auctionId, lotNo]);
        if (existing) { skipped++; skipReasons.push({row: rowNum, lot: lotNo, reason: `Duplicate — lot ${lotNo} already exists in Trade ${rowAno}`}); continue; }

        // Try to find trader by name for linking
        const sellerName = mapCol(row, 'NAME', 'SELLER', 'POOLER', 'TRADER');
        let traderId = null;
        if (sellerName) {
          const trader = db.get('SELECT id FROM traders WHERE name = ?', [sellerName]);
          if (trader) traderId = trader.id;
        }

        try {
          db.run(`INSERT INTO lots (auction_id, lot_no, crop, grade, crpt, branch, state, trader_id,
            name, padd, ppla, ppin, pstate, pst_code, cr, pan, tel, aadhar,
            bags, litre, qty, price, amount, code, buyer, buyer1, sale, invo,
            pqty, prate, puramt, com, sertax, cgst, sgst, igst, advance, balance, bilamt, user_id)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`,
            [auctionId, lotNo,
             mapCol(row, 'CROP'),
             mapCol(row, 'GRADE'),
             mapCol(row, 'CRPT', 'CROP_TYPE') || cropType,
             mapCol(row, 'BR', 'BRANCH', 'DEPOT'),
             mapCol(row, 'STATE') || state,
             traderId,
             sellerName,
             mapCol(row, 'PADD', 'ADDRESS', 'ADD', 'ADD1'),
             mapCol(row, 'PPLA', 'PLACE', 'PLA'),
             mapCol(row, 'PPIN', 'PIN', 'PINCODE'),
             mapCol(row, 'PSTATE'),
             mapCol(row, 'PST_CODE', 'ST_CODE', 'STATE_CODE'),
             mapCol(row, 'CR', 'GSTIN', 'CR_NO'),
             mapCol(row, 'PAN', 'PAN_NO'),
             mapCol(row, 'TEL', 'PHONE', 'MOBILE'),
             mapCol(row, 'AADHAR', 'AADHAAR'),
             mapNum(row, 'BAG', 'BAGS'),
             mapCol(row, 'LITRE', 'LITRE_WT'),
             mapNum(row, 'QTY', 'QUANTITY', 'NET_QTY'),
             mapNum(row, 'PRICE', 'RATE'),
             mapNum(row, 'AMOUNT'),
             mapCol(row, 'CODE', 'BUYER_CODE'),
             mapCol(row, 'BUYER', 'BIDDER'),
             mapCol(row, 'BUYER1', 'TRADE_NAME', 'TRADENAME'),
             mapCol(row, 'SALE', 'SALE_TYPE'),
             mapCol(row, 'INVO', 'INVOICE'),
             mapNum(row, 'PQTY', 'PUR_QTY'),
             mapNum(row, 'PRATE', 'PUR_RATE'),
             mapNum(row, 'PURAMT', 'PUR_AMT', 'PURCHASE_AMT'),
             mapNum(row, 'COM', 'COMMISSION'),
             mapNum(row, 'SERTAX', 'HPC'),
             mapNum(row, 'CGST'),
             mapNum(row, 'SGST'),
             mapNum(row, 'IGST'),
             mapNum(row, 'ADVANCE', 'DISCOUNT'),
             mapNum(row, 'BALANCE', 'PAYABLE'),
             mapNum(row, 'BILAMT', 'BILL_AMT'),
             mapCol(row, 'USER_ID', 'USER') || 'import']);
          imported++;
          const key = `${rowAno}|${rowDate}`;
          auctionStats.set(key, (auctionStats.get(key) || 0) + 1);
        } catch (e) {
          skipped++;
          skipReasons.push({row: rowNum, lot: lotNo, reason: `DB error: ${e.message}`});
        }
      }
    }

    // Auto-allocate (full import only). For each auction that received
    // NEW lots in this import AND has no existing allocations, create
    // one allocation per (branch, lot-prefix) pair covering 001..999 of
    // that prefix (or 0001..9999 if any imported lot exceeds 999). This
    // saves operators from the orphan-on-first-allocation-edit trap —
    // without it, imported trades start with zero coverage and any
    // later attempt to add allocations runs into the "lots would be
    // orphaned" guard.
    //
    // Skipped in 'price' mode (it only updates existing lots, never
    // adds new ones, so coverage status doesn't change).
    //
    // The empty-string branch is included so trades imported without a
    // BR column still get covered. The lot-entry guard treats blank
    // branch as "no allocation enforcement" anyway, but having a row
    // means the Allocations modal isn't blank for the user.
    let autoAllocCreated = 0;
    if (mode !== 'price') {
      for (const auc of auctionCache.values()) {
        if (!auc || !auc.id) continue;
        const hasAny = db.get('SELECT 1 FROM lot_allocations WHERE auction_id = ? LIMIT 1', [auc.id]);
        if (hasAny) continue;     // respect existing setup
        const lots = db.all('SELECT lot_no, branch FROM lots WHERE auction_id = ?', [auc.id]);
        // Group lots by (branch, prefix) → track the max numeric part.
        const groups = new Map(); // key = `${branch}||${prefix}` → { branch, prefix, padLen, max }
        for (const l of lots) {
          const p = parseLotNo(l.lot_no);
          if (!p) continue;
          const branch = String(l.branch || '').trim();
          const key = `${branch}||${p.prefix}`;
          const g = groups.get(key);
          if (!g) groups.set(key, { branch, prefix: p.prefix, padLen: p.padLen, max: p.num });
          else if (p.num > g.max) g.max = p.num;
        }
        for (const g of groups.values()) {
          // End range = 999 (or 9999 if anything went above), so future
          // hand-entered lots up to that ceiling are covered without
          // needing a second allocation save.
          const end = g.max > 999 ? 9999 : 999;
          const padLen = Math.max(g.padLen, String(end).length);
          const startLot = buildLotNo(g.prefix, 1, padLen);
          const endLot   = buildLotNo(g.prefix, end, padLen);
          db.run(
            'INSERT INTO lot_allocations (auction_id, branch, start_lot, end_lot) VALUES (?, ?, ?, ?)',
            [auc.id, g.branch, startLot, endLot]
          );
          autoAllocCreated++;
        }
      }
    }

    // Build auction breakdown for the response
    const auctionBreakdown = [];
    for (const [key, count] of auctionStats) {
      const [ano, date] = key.split('|');
      const auc = auctionCache.get(key);
      auctionBreakdown.push({ id: auc?.id, ano, date, count });
    }
    auctionBreakdown.sort((a,b) => String(a.ano).localeCompare(String(b.ano), undefined, {numeric:true}));

    return {
      success: true,
      imported, updated, skipped, total: rows.length,
      auctionCount: auctionBreakdown.length,
      auctionBreakdown,
      autoAllocCreated,
      skipReasons
    };
}

// Manual price/lot import from an uploaded .xls/.xlsx file.
app.post('/api/auctions/import', requireAuctionWrite, upload.single('file'), requireLotsValidatedForPriceImport, (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
  try {
    const result = runLotImport(getDb(), req.file.path, req.body);
    fs.unlink(req.file.path, () => {});
    res.json(result);
  } catch (e) {
    if (req.file) fs.unlink(req.file.path, () => {});
    res.status(400).json({ error: e.message });
  }
});

// ══════════════════════════════════════════════════════════════
// TRADE FAIR SYNC (tradefair.intelloinsights.com)
// ──────────────────────────────────────────────────────────────
// Pull auction history + per-auction price-list .xlsx from the
// trade-fair site (using a pasted _kcpmc_rails_session cookie) and feed
// the price list straight through the SAME import pipeline a manual
// upload uses (runLotImport, mode='price'). Config + secret cookie live
// in the trade_fair_config table; see trade-fair.js for the client.
// ══════════════════════════════════════════════════════════════

// Resolve effective trade-fair config: process.env wins over the DB
// row (lets a hosted deploy set TRADEFAIR_* vars; desktop pastes into
// the Settings card). Returns the cookie too — callers keep it
// server-side and never echo it to the browser.
function _tfConfig(db) {
  let row = {};
  try { row = db.get('SELECT * FROM trade_fair_config WHERE id = 1') || {}; } catch (_) { row = {}; }
  const envv = (k) => (process.env[k] && String(process.env[k]).trim()) || '';
  const dbv  = (c) => (row[c] != null && String(row[c]).trim()) || '';
  const pick = (k, c, dflt) => envv(k) || dbv(c) || (dflt || '');
  const sessionCookie = pick('TRADEFAIR_COOKIE', 'session_cookie', '');
  return {
    baseUrl:     pick('TRADEFAIR_BASE_URL',     'base_url',     'https://tradefair.intelloinsights.com'),
    historyPath: pick('TRADEFAIR_HISTORY_PATH', 'history_path', '/spices/admin/get_trade_fair_history'),
    pricePath:   pick('TRADEFAIR_PRICE_PATH',   'price_path',   '/spices/reports/download_price_list_report_excel/'),
    idField:     pick('TRADEFAIR_ID_FIELD',     'id_field',     'trade_session_id'),
    priceParam:  pick('TRADEFAIR_PRICE_PARAM',  'price_param',  'auction_id'),
    cookieName:  dbv('cookie_name') || '_kcpmc_rails_session',
    sessionCookie,
    enabled:     row.enabled == null ? 1 : Number(row.enabled),
    cookieUpdatedAt: dbv('cookie_updated_at'),
    updatedAt:   dbv('updated_at'),
    cookieFromEnv: !!envv('TRADEFAIR_COOKIE'),
    configured:  !!sessionCookie,
  };
}

// Read the (possibly nested) id value used for the price-list download
// from a history row, honouring the configured id_field with sensible
// fallbacks.
function _tfRowId(cfg, r) {
  let v = r[cfg.idField];
  if (v == null) v = r.auction_id != null ? r.auction_id : r.trade_session_id;
  if (v != null && typeof v === 'object') v = (v.id != null ? v.id : v.value);
  return v == null ? '' : v;
}

// Non-secret status for the Settings card. NEVER returns the cookie —
// only whether one is saved + when.
app.get('/api/trade-fair/status', requireView, (req, res) => {
  const cfg = _tfConfig(getDb());
  res.json({
    configured:      cfg.configured,
    enabled:         !!cfg.enabled,
    baseUrl:         cfg.baseUrl,
    historyPath:     cfg.historyPath,
    pricePath:       cfg.pricePath,
    idField:         cfg.idField,
    priceParam:      cfg.priceParam,
    cookieName:      cfg.cookieName,
    hasCookie:       cfg.configured,
    cookieFromEnv:   cfg.cookieFromEnv,
    cookieUpdatedAt: cfg.cookieUpdatedAt || null,
    source:          cfg.cookieFromEnv ? 'env' : (cfg.configured ? 'db' : 'none'),
  });
});

// Upsert trade-fair config. The session cookie is write-only: a blank/
// absent value leaves the stored cookie UNCHANGED (so saving paths
// doesn't wipe it). Saving a new cookie stamps cookie_updated_at.
app.put('/api/trade-fair/config', requireSettingsWrite, (req, res) => {
  const db = getDb();
  const b = req.body || {};
  const sets = [], vals = [];
  const put = (col, val) => { sets.push(`${col} = ?`); vals.push(val); };
  const putField = (col, val) => { if (val !== undefined) put(col, String(val).trim()); };
  putField('base_url', b.baseUrl);
  putField('history_path', b.historyPath);
  putField('price_path', b.pricePath);
  putField('id_field', b.idField);
  putField('price_param', b.priceParam);
  putField('cookie_name', b.cookieName);
  if (b.enabled !== undefined) put('enabled', b.enabled ? 1 : 0);
  // Secret cookie — only overwrite when a non-blank value is supplied.
  if (typeof b.sessionCookie === 'string' && b.sessionCookie.trim()) {
    put('session_cookie', b.sessionCookie.trim());
    put('cookie_updated_at', new Date().toISOString());
  } else if (b.clearCookie) {
    put('session_cookie', '');
    put('cookie_updated_at', '');
  }
  if (!sets.length) return res.json({ ok: true, updated: 0 });
  try {
    db.run(`UPDATE trade_fair_config SET ${sets.join(', ')}, updated_at = datetime('now','localtime') WHERE id = 1`, vals);
  } catch (e) { return res.status(500).json({ error: e.message }); }
  res.json({ ok: true, updated: sets.length });
});

// List recent trade-fair auctions/sessions (newest first). Doubles as
// the "test connection" probe — a bad/expired cookie surfaces here as a
// clear session-expired error.
app.get('/api/trade-fair/history', requireAuctionWrite, async (req, res) => {
  const cfg = _tfConfig(getDb());
  const limit = Math.min(Math.max(parseInt(req.query.limit, 10) || 100, 1), 1000);
  const search = req.query.search ? String(req.query.search) : '';
  try {
    const { rows, total } = await tradeFair.fetchHistory(cfg, { limit, search });
    const auctions = rows.map(r => ({
      idValue:        _tfRowId(cfg, r),
      ano:            r.auction_number != null ? String(r.auction_number) : '',
      sbAuction:      r.sb_auction_number || '',
      date:           r.auction_date || '',
      tradeType:      r.trade_type || '',
      status:         r.trade_session_status || '',
      lotCount:       r.auction_lot_count,
      lotQty:         r.auction_lot_qty,
      sellQty:        r.trade_session_sell_qty,
      auctionId:      r.auction_id,
      tradeSessionId: r.trade_session_id,
    }));
    res.json({ ok: true, total, count: auctions.length, idField: cfg.idField, priceParam: cfg.priceParam, auctions });
  } catch (e) {
    res.status(502).json({ error: e.message });
  }
});

// Download one auction's raw price-list .xlsx for inspection / manual
// import — handy when the auto-import column mapping needs checking.
app.get('/api/trade-fair/price-list', requireAuctionWrite, async (req, res) => {
  const cfg = _tfConfig(getDb());
  const idValue = req.query.id || req.query.idValue;
  if (!idValue) return res.status(400).json({ error: 'Missing ?id= (the auction/session id)' });
  try {
    const { buffer, ext } = await tradeFair.downloadPriceList(cfg, idValue);
    const name = `tradefair-price-${String(idValue).replace(/[^A-Za-z0-9._-]+/g, '')}${ext}`;
    res.setHeader('Content-Type', ext === '.xls'
      ? 'application/vnd.ms-excel'
      : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${name}"`);
    res.send(buffer);
  } catch (e) {
    res.status(502).json({ error: e.message });
  }
});

// Download a trade-fair auction's price list and run it through the
// price-import pipeline (mode='price' by default). The operator maps
// the trade-fair auction (idValue) onto an app trade (ano + date),
// which override every row so prices land on the right lots. Respects
// the same lot-validation gate as the manual upload.
app.post('/api/trade-fair/import', requireAuctionWrite, async (req, res) => {
  const db = getDb();
  const cfg = _tfConfig(db);
  const b = req.body || {};
  const idValue = b.idValue != null ? b.idValue : b.id;
  if (idValue == null || idValue === '') return res.status(400).json({ error: 'Missing idValue (the trade-fair auction/session id).' });

  // Build the import body. Default to 'price' (update existing lots) —
  // that's the post-auction price sync the trade-fair sheet is for.
  const importBody = {
    mode: b.mode === 'full' ? 'full' : 'price',
    ano: b.ano != null ? String(b.ano).trim() : '',
    date: b.date || '',
    crop_type: b.crop_type,
    state: b.state,
  };

  // Same lot-validation gate the manual upload enforces.
  const block = lotValidationGateBlock(db, importBody);
  if (block) return res.status(block.status).json(block.body);

  let tmpPath = null;
  try {
    const { buffer, ext } = await tradeFair.downloadPriceList(cfg, idValue);
    try { fs.mkdirSync(uploadDir, { recursive: true }); } catch (_) {}
    tmpPath = path.join(uploadDir, 'tf-' + crypto.randomBytes(6).toString('hex') + ext);
    fs.writeFileSync(tmpPath, buffer);
    const result = runLotImport(db, tmpPath, importBody);
    res.json({ ...result, source: 'trade-fair', idValue });
  } catch (e) {
    res.status(400).json({ error: e.message });
  } finally {
    if (tmpPath) fs.unlink(tmpPath, () => {});
  }
});

// ── Download Auction/Lots template XLSX ──────────────────────
app.get('/api/auctions/template', requireExport, async (req, res) => {
  const db = getDb();
  const headers = ['ANO','DATE','LOT','CROP','GRADE','CRPT','BR','STATE','NAME','PADD','PPLA','PPIN','PSTATE','PST_CODE',
    'CR','PAN','TEL','AADHAR','BAG','LITRE','QTY','PRICE','AMOUNT','CODE','BUYER','BUYER1','SALE','INVO',
    'PQTY','PRATE','PURAMT','CGST','SGST','IGST','ADVANCE','BALANCE'];
  const cols = headers.map(h => ({ header: h, key: h.toLowerCase(), width: h.length < 5 ? 9 : 14 }));
  // Pull dynamic defaults from settings — no hardcoded 'ASP' / 'TAMIL NADU'
  const defaultCrop  = getSetting(db, 'default_crop_type') || '';
  const bizState     = (getSetting(db, 'business_state') || 'TAMIL NADU').toUpperCase();
  const stCode       = bizState === 'KERALA' ? '32' : '33';
  const sample = [{
    ano: '1', date: '2026-04-15', lot: '001', crop: '', grade: '1',
    crpt: defaultCrop, br: '', state: bizState,
    name: 'SAMPLE SELLER', padd: '123 MAIN ST', ppla: '', ppin: '',
    pstate: bizState, pst_code: stCode, cr: 'CR.001', pan: 'ABCDE1234F', tel: '9876543210', aadhar: '',
    bag: 5, litre: '380', qty: 100.567, price: 0, amount: 0, code: '', buyer: '', buyer1: '', sale: '', invo: '',
    pqty: 0, prate: 0, puramt: 0, cgst: 0, sgst: 0, igst: 0, advance: 0, balance: 0,
  }];
  const buf = await createExcelBuffer('Lots', cols, sample, {
    db, title: 'AUCTION / LOTS TEMPLATE',
    metaLines: [`Date: ${fmtDate(todayLocalISO())}`],
  });
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename="auction-lots-template.xlsx"');
  res.send(Buffer.from(buf));
});

// ══════════════════════════════════════════════════════════════
// LOTS (CPA1.DBF — main data)
// ══════════════════════════════════════════════════════════════
app.get('/api/lots/:auctionId', requireViewOrLotEntry, (req, res) => {
  const { branch, name, buyer, limit, offset, paginated, summary, search } = req.query;
  const db = getDb();
  // Correlated subquery (not LEFT JOIN) to avoid any risk of row duplication
  // if the same buyer code exists multiple times in the buyers table.
  let q = `SELECT lots.*,
             (SELECT b.code FROM buyers b WHERE b.buyer = lots.buyer LIMIT 1) AS buyer_code
           FROM lots
           WHERE lots.auction_id = ?`;
  const p = [req.params.auctionId];
  if (branch) { q += ' AND lots.branch = ?'; p.push(branch); }
  if (name)   { q += ' AND lots.name LIKE ?'; p.push(`%${name}%`); }
  if (buyer)  { q += ' AND lots.buyer = ?'; p.push(buyer); }
  // Free-text search within the trade — lot no, seller name, buyer
  // (short trade name), code, invoice no, branch. The "Code" column
  // shown in the Lots table is sourced from either `lots.code` (set by
  // price-import / set-buyer flows) or a join to `buyers.code` keyed
  // by `lots.buyer`. Both surfaces are searched so a user typing the
  // short alias they see on screen reliably gets a hit.
  //
  // `lots.buyer1` (the denormalised long buyer name) is intentionally
  // NOT searched. The Lots table only displays the short `buyer`
  // column; matching against `buyer1` produced rows that looked
  // irrelevant to the operator (e.g. searching "spices" hit lots
  // whose short buyer was "CHINNAMUTHU1" because their *full* name was
  // "CHINNAMUTHU SPICES"). Long-name search lives on the Buyers tab.
  const searchTerm = String(search || '').trim();
  if (searchTerm) {
    const wild = `%${searchTerm}%`;
    q += ` AND (
            COALESCE(lots.lot_no,'') LIKE ?
            OR COALESCE(lots.name,'')   LIKE ?
            OR COALESCE(lots.buyer,'')  LIKE ?
            OR COALESCE(lots.code,'')   LIKE ?
            OR COALESCE(lots.invo,'')   LIKE ?
            OR COALESCE(lots.branch,'') LIKE ?
            OR EXISTS (
              SELECT 1 FROM buyers b
               WHERE b.buyer = lots.buyer
                 AND COALESCE(b.code,'') LIKE ?
            )
          )`;
    p.push(wild, wild, wild, wild, wild, wild, wild);
  }

  // Summary mode — returns aggregate counts only (cheap, no row data).
  // Used by the Lot Entry stats badge so it shows true totals even when
  // only a 25-row window of lots is loaded client-side. The WHERE
  // clauses below MUST stay in lockstep with the placeholders pushed
  // onto `p` above (auction_id + branch + name + buyer + search×7) —
  // a mismatch silently drops or swallows parameters.
  if (summary === '1') {
    let aggSql =
      `SELECT COUNT(*) AS n,
              COALESCE(SUM(CAST(bags AS INTEGER)), 0) AS bags,
              COALESCE(SUM(qty), 0)                  AS qty,
              COALESCE(SUM(CASE WHEN price > 0 THEN 1 ELSE 0 END), 0) AS priced
         FROM lots
        WHERE lots.auction_id = ?`
      + (branch ? ' AND lots.branch = ?' : '')
      + (name   ? ' AND lots.name LIKE ?' : '')
      + (buyer  ? ' AND lots.buyer = ?' : '');
    if (searchTerm) {
      // Same column set as the main GET — buyer1 intentionally excluded
      // so the summary count matches what the operator sees in the table.
      aggSql += ` AND (
            COALESCE(lots.lot_no,'') LIKE ?
            OR COALESCE(lots.name,'')   LIKE ?
            OR COALESCE(lots.buyer,'')  LIKE ?
            OR COALESCE(lots.code,'')   LIKE ?
            OR COALESCE(lots.invo,'')   LIKE ?
            OR COALESCE(lots.branch,'') LIKE ?
            OR EXISTS (
              SELECT 1 FROM buyers b
               WHERE b.buyer = lots.buyer
                 AND COALESCE(b.code,'') LIKE ?
            )
          )`;
    }
    const row = db.get(aggSql, p) || { n:0, bags:0, qty:0, priced:0 };
    return res.json({ n: row.n, bags: row.bags, qty: row.qty, priced: row.priced });
  }

  // Pagination — opt-in via `paginated=1` so the existing callers
  // (Lots tab, exports, etc.) keep getting the full list as a flat array.
  // The Lot Entry "Recent entries" panel passes paginated=1&limit=25&offset=N
  // and expects { rows, total } so it can show a "Load more" / page count.
  // ORDER BY lot_no DESC for the recent panel — newest entries first;
  // existing flat-list callers keep ascending order.
  if (paginated === '1') {
    const lim = Math.min(Math.max(parseInt(limit, 10) || 25, 1), 200);
    const off = Math.max(parseInt(offset, 10) || 0, 0);
    // Total count — needed so the client can show "Showing X of Y" and
    // know when to hide "Load more". Cheap because the WHERE clause is
    // narrow (auction_id is indexed).
    let cq = q.replace(
      /^SELECT[\s\S]+?FROM lots/,
      'SELECT COUNT(*) AS n FROM lots'
    );
    const total = (db.get(cq, p) || {}).n || 0;

    q += ' ORDER BY CAST(lots.lot_no AS INTEGER) DESC, lots.lot_no DESC LIMIT ? OFFSET ?';
    const rows = db.all(q, [...p, lim, off]);
    return res.json({ rows, total, limit: lim, offset: off });
  }

  q += ' ORDER BY lots.lot_no';
  res.json(db.all(q, p));
});

// Append a row to audit_log for a lot create / edit / delete so the Lot
// Entry activity feed can show WHO did WHAT and WHEN. The auction_id is
// always written as the FIRST key of the JSON details so the feed
// endpoint can filter by trade with a collision-free LIKE prefix
// ('{"auction_id":<id>,'). Audit is best-effort — a logging failure must
// never break the actual lot mutation, so the whole thing is wrapped.
function logLotActivity(db, req, action, lotId, details) {
  try {
    const user = (req && req.user && req.user.username) || 'unknown';
    db.run(
      'INSERT INTO audit_log (user_id, action, entity, entity_id, details) VALUES (?,?,?,?,?)',
      [String(user), action, 'lot', lotId != null ? Number(lotId) : null, JSON.stringify(details || {})]
    );
  } catch (e) { /* best-effort audit; ignore */ }
}

app.post('/api/lots', requireLotWrite, (req, res) => {
  const l = req.body;
  const db = getDb();
  const auctionId = parseInt(l.auction_id, 10);
  const lotNoStr  = String(l.lot_no || '').trim();
  const branch    = String(l.branch || '').trim();

  if (!auctionId || !lotNoStr) {
    return res.status(400).json({ error: 'auction_id and lot_no are required' });
  }

  // Reject duplicates within the same trade so two field-staff users
  // can't end up with the same lot number on different rows.
  const existing = db.get(
    'SELECT id FROM lots WHERE auction_id = ? AND lot_no = ?',
    [auctionId, lotNoStr]
  );
  if (existing) {
    return res.status(409).json({ error: `Lot #${lotNoStr} already exists in this auction` });
  }

  // The lot must be bound to the exact trader_id the UI picked. When
  // a seller has multiple GSTIN/branch records in the traders table
  // (legitimate or legacy), this guarantees the lot doesn't silently
  // pick up a different one's denormalised values via name fallback.
  // We rebuild name/cr/pan/etc. from the chosen trader row so a stale
  // payload (e.g. user picked seller A then their session cached B's
  // fields) can't write inconsistent data.
  const traderId = l.trader_id != null ? parseInt(l.trader_id, 10) : null;
  let trader = null;
  if (traderId) {
    trader = db.get('SELECT * FROM traders WHERE id = ?', [traderId]);
    if (!trader) {
      return res.status(400).json({ error: `Selected seller (trader id ${traderId}) no longer exists. Re-pick from the search.` });
    }
  }
  if (trader) {
    // Force-coerce the denormalised fields to the picked trader's row
    // so the saved lot is unambiguous about which GSTIN / branch /
    // address it belongs to. The client already sends these values,
    // but we override on the server as the authoritative source.
    l.name   = trader.name   || l.name   || '';
    l.cr     = trader.cr     || '';
    l.pan    = trader.pan    || '';
    l.tel    = trader.tel    || '';
    l.aadhar = trader.aadhar || '';
    l.padd   = trader.padd   || '';
    l.ppla   = trader.ppla   || '';
    l.ppin   = trader.pin    || '';
    l.pstate = trader.pstate || l.pstate || '';
    l.pst_code = trader.pst_code || l.pst_code || '';
  }

  // Allocation enforcement — only kicks in when the trade has explicit
  // allocations for this branch. Trades created before allocations were
  // configured (or branches without one) skip the check, preserving
  // backward compatibility with existing data.
  if (branch) {
    const allocs = db.all(
      'SELECT * FROM lot_allocations WHERE auction_id = ? AND branch = ?',
      [auctionId, branch]
    );
    if (allocs.length > 0) {
      const inRange = allocs.some(a => isLotInRange(lotNoStr, a.start_lot, a.end_lot));
      if (!inRange) {
        const ranges = allocs.map(a => a.start_lot + '-' + a.end_lot).join(', ');
        return res.status(400).json({ error: `Lot #${lotNoStr} is outside ${branch} allocation (${ranges})` });
      }
    }
  }

  // Crop receipt number — per-trade running counter. Auto-assigned at
  // INSERT time so it's atomic and race-free across concurrent saves
  // (the SELECT MAX + INSERT runs inside a single JS turn, and SQLite
  // serialises writes). If the client supplied an explicit value, honour
  // it (allows backfill / manual override during edits / imports).
  let cropReceiptNo = null;
  if (l.crop_receipt_no != null && l.crop_receipt_no !== '') {
    const n = parseInt(l.crop_receipt_no, 10);
    if (Number.isFinite(n) && n > 0) cropReceiptNo = n;
  }
  if (cropReceiptNo == null) {
    const maxRow = db.get(
      'SELECT MAX(crop_receipt_no) AS m FROM lots WHERE auction_id = ?',
      [auctionId]
    );
    cropReceiptNo = ((maxRow && maxRow.m) || 0) + 1;
  }

  // Per-lot bank account (FK trader_banks.id). When the operator picked an
  // account, honour it. Otherwise ALWAYS default to the seller's default
  // account (is_default=1, else first row) so lots never land untagged —
  // an untagged lot alongside tagged ones used to (falsely) trip the
  // "multiple banks" badge on the Payments tab. Stays null only when the
  // seller has no banks on file at all.
  let bankId = (l.bank_id != null && l.bank_id !== '') ? (Number(l.bank_id) || null) : null;
  if (bankId == null && l.trader_id) {
    try {
      const defBank = db.get(
        `SELECT id FROM trader_banks WHERE trader_id = ?
          ORDER BY is_default DESC, id ASC LIMIT 1`,
        [l.trader_id]
      );
      if (defBank && defBank.id != null) bankId = Number(defBank.id) || null;
    } catch (_) { /* trader_banks may not exist on partial migrations */ }
  }
  const ins = db.run(`INSERT INTO lots (auction_id,lot_no,crop,grade,crpt,branch,state,trader_id,name,padd,ppla,ppin,pstate,pst_code,cr,pan,tel,aadhar,bags,litre,qty,gross_wt,sample_wt,weight_with_gunny,moisture,crop_receipt_no,reserved_price,bank_id,user_id)
    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`,
    [auctionId,lotNoStr,l.crop||'',l.grade||'',l.crpt||'',branch,l.state||'TAMIL NADU',l.trader_id||null,l.name||'',l.padd||'',l.ppla||'',l.ppin||'',l.pstate||'',l.pst_code||'',l.cr||'',l.pan||'',l.tel||'',l.aadhar||'',l.bags||0,l.litre||'',l.qty||0,l.gross_wt||0,l.sample_wt||0,Number(l.weight_with_gunny)||0,l.moisture||'',cropReceiptNo,Number(l.reserved_price)||0,bankId,l.user_id||'']);
  pcClearGate(db, auctionId);
  lvClearGate(db, auctionId);   // new lot → trade must be re-validated before price import
  logLotActivity(db, req, 'create', ins && ins.lastInsertRowid, { auction_id: auctionId, lot_no: lotNoStr, name: l.name || '' });
  res.json({ success: true, crop_receipt_no: cropReceiptNo });
});

// Whitelist of columns that PUT /api/lots/:id is allowed to update.
// Anything outside this set is silently dropped — keeps untrusted keys out
// of the dynamically-built UPDATE statement (was an SQL-injection vector
// when the request-body keys were spliced raw into SQL) and stops clients
// from rewriting attribution / server-calculated fields they shouldn't
// touch (id, auction_id, created_at are excluded by construction).
const LOT_UPDATE_COLUMNS = new Set([
  'lot_no','crop','grade','crpt','branch','state','trader_id',
  'name','padd','ppla','ppin','pstate','pst_code','cr','pan','tel','aadhar',
  'bags','litre','qty','gross_wt','sample_wt','weight_with_gunny','moisture',
  'crop_receipt_no','reserved_price','bank_id',
  'price','amount','code','buyer','buyer1','sale','invo',
  'pqty','prate','puramt','com','sertax','cgst','sgst','igst',
  'dcgst','dsgst','digst','refud','refund','advance','balance','bilamt','paid',
  'user_id','asp_invo',
  'isp_pqty','isp_prate','isp_puramt','asp_pqty','asp_prate','asp_puramt'
]);

// Fields that feed calculateLot OR that imply the planter math should
// be re-evaluated. When a PUT touches any of these, the endpoint
// re-runs the calc so the dependent columns (prate/puramt/com/cgst/
// sgst/igst/bilamt/advance/balance) stay consistent. `code` is here
// because a code change to 'WD' (withdrawn) typically zeroes price
// elsewhere in the flow — and we need to clear stale planter values
// regardless of what price ends up as.
const CALC_TRIGGER_FIELDS = ['price','qty','grade','refud','cr','code'];

app.put('/api/lots/:id', requireLotWrite, (req, res) => {
  const l = req.body; const sets = []; const vals = [];
  const db = getDb();
  const lotId = parseInt(req.params.id, 10);

  // If lot_no or branch is being changed, re-run the allocation +
  // duplicate validation that POST /api/lots applies. Skipping these
  // on edit was a previous gap that let users move a lot outside its
  // branch's range or collide with another lot's number.
  const current = db.get('SELECT auction_id, lot_no, branch, locked_at FROM lots WHERE id = ?', [lotId]);
  // Lock guard — admins always pass; non-admins are blocked the moment
  // the row is marked locked, regardless of which fields they're trying
  // to change. Returning 423 (Locked) so the client can show a tailored
  // message instead of the generic "forbidden".
  if (current && current.locked_at && !isAdmin(req)) {
    return res.status(423).json({ error: 'This lot is locked — only an admin can edit it.' });
  }
  if (current) {
    const newLotNo = (l.lot_no != null) ? String(l.lot_no).trim() : current.lot_no;
    const newBranch = (l.branch != null) ? String(l.branch).trim() : current.branch;
    if (newLotNo !== current.lot_no) {
      const dup = db.get(
        'SELECT id FROM lots WHERE auction_id = ? AND lot_no = ? AND id != ?',
        [current.auction_id, newLotNo, lotId]
      );
      if (dup) return res.status(409).json({ error: `Lot #${newLotNo} already exists in this auction` });
    }
    if (newBranch && (newLotNo !== current.lot_no || newBranch !== current.branch)) {
      const allocs = db.all(
        'SELECT * FROM lot_allocations WHERE auction_id = ? AND branch = ?',
        [current.auction_id, newBranch]
      );
      if (allocs.length > 0) {
        const inRange = allocs.some(a => isLotInRange(newLotNo, a.start_lot, a.end_lot));
        if (!inRange) {
          const ranges = allocs.map(a => a.start_lot + '-' + a.end_lot).join(', ');
          return res.status(400).json({ error: `Lot #${newLotNo} is outside ${newBranch} allocation (${ranges})` });
        }
      }
    }
  }

  for (const [k,v] of Object.entries(l)) {
    if (!LOT_UPDATE_COLUMNS.has(k)) continue;
    sets.push(`${k}=?`); vals.push(v);
  }
  if (sets.length === 0) return res.json({ success: true });
  vals.push(lotId);
  db.run(`UPDATE lots SET ${sets.join(',')} WHERE id=?`, vals);
  // If a field that feeds calculateLot changed, refresh the planter-side
  // columns in-place. Otherwise `prate`/`puramt`/`com`/`cgst`/`sgst`/
  // `igst`/`bilamt`/`advance`/`balance` stay frozen at their pre-edit
  // values until the operator clicks Calculate All — and the generate
  // endpoints' auto-calc only heals rows where `puramt = 0`, so a
  // stale-but-non-zero puramt would silently flow into purchase
  // invoices / bills / debit notes after a price tweak.
  //
  // Recalc runs even when the lot ends up with zero inputs (e.g.
  // code → WD, which zeroes price): calculateLot produces zeros for
  // all planter fields in that case, which is the correct outcome.
  // Without this, marking a lot as withdrawn left stale prate/puramt
  // values that propagated into bills + debit notes.
  //
  // Skip locked lots (same rule Calculate All uses — finalised planter
  // numbers must not be overwritten). Admins editing a locked lot must
  // unlock + Calculate All explicitly if they want the cascade.
  if (current && !current.locked_at &&
      CALC_TRIGGER_FIELDS.some(f => Object.prototype.hasOwnProperty.call(l, f))) {
    const fresh = db.get('SELECT * FROM lots WHERE id = ?', [lotId]);
    if (fresh) {
      const cfg = getSettingsFlat(db);
      const calc = calculateLot(fresh, cfg);
      db.run(`UPDATE lots SET amount=?,pqty=?,prate=?,puramt=?,com=?,sertax=?,cgst=?,sgst=?,igst=?,advance=?,balance=?,bilamt=?,refund=?,refud=?,isp_pqty=?,isp_prate=?,isp_puramt=?,asp_pqty=?,asp_prate=?,asp_puramt=? WHERE id=?`,
        [calc.amount,calc.pqty,calc.prate,calc.puramt,calc.com,calc.sertax,calc.cgst,calc.sgst,calc.igst,calc.advance,calc.balance,calc.bilamt,calc.refund||0,calc.refud||0,calc.isp_pqty||0,calc.isp_prate||0,calc.isp_puramt||0,calc.asp_pqty||0,calc.asp_prate||0,calc.asp_puramt||0,lotId]);
    }
  }
  // Any lot edit invalidates a previous price-check stamp — the data
  // the verify run looked at has changed. Operator must re-verify
  // before the next calculate/invoice/etc.
  if (current && current.auction_id) { pcClearGate(db, current.auction_id); lvClearGate(db, current.auction_id); }
  if (current) {
    const changedFields = Object.keys(l).filter(k => LOT_UPDATE_COLUMNS.has(k));
    const newLotNo = (l.lot_no != null) ? String(l.lot_no).trim() : current.lot_no;
    logLotActivity(db, req, 'update', lotId, { auction_id: current.auction_id, lot_no: newLotNo, name: l.name, fields: changedFields });
  }
  res.json({ success: true });
});

app.delete('/api/lots/:id', requireDelete, (req, res) => {
  const db = getDb();
  const cur = db.get('SELECT auction_id, lot_no, name, locked_at FROM lots WHERE id = ?', [req.params.id]);
  if (cur && cur.locked_at && !isAdmin(req)) {
    return res.status(423).json({ error: 'This lot is locked — only an admin can delete it.' });
  }
  db.run('DELETE FROM lots WHERE id = ?', [req.params.id]);
  if (cur && cur.auction_id) { pcClearGate(db, cur.auction_id); lvClearGate(db, cur.auction_id); }
  if (cur) logLotActivity(db, req, 'delete', Number(req.params.id), { auction_id: cur.auction_id, lot_no: cur.lot_no, name: cur.name });
  res.json({ success: true });
});

// ── Lot activity feed ──────────────────────────────────────────
// Powers the collapsible "Activity Log" panel in the Lot Entry screen.
// Returns who created / edited / deleted lots, newest first, paginated.
// Optional ?auctionId= scopes to one trade via a collision-free LIKE on
// the auction_id stored as the first key of each details JSON.
app.get('/api/lot-activity', requireAuth, (req, res) => {
  try {
    const db = getDb();
    const auctionId = req.query.auctionId ? parseInt(req.query.auctionId, 10) : null;
    const page = Math.max(1, parseInt(req.query.page, 10) || 1);
    const pageSize = Math.min(100, Math.max(1, parseInt(req.query.pageSize, 10) || 20));
    let where = "entity = 'lot'";
    const params = [];
    if (auctionId) { where += ' AND details LIKE ?'; params.push(`{"auction_id":${auctionId},%`); }
    const totalRow = db.get(`SELECT COUNT(*) AS n FROM audit_log WHERE ${where}`, params);
    const total = (totalRow && totalRow.n) || 0;
    const rows = db.all(
      `SELECT id, user_id, action, entity_id, details, created_at
         FROM audit_log WHERE ${where}
        ORDER BY id DESC LIMIT ? OFFSET ?`,
      [...params, pageSize, (page - 1) * pageSize]
    );
    const items = rows.map(r => {
      let d = {};
      try { d = JSON.parse(r.details || '{}'); } catch (e) { /* tolerate legacy rows */ }
      return {
        id: r.id, user: r.user_id, action: r.action, lot_id: r.entity_id,
        lot_no: d.lot_no || '', name: d.name || '',
        fields: Array.isArray(d.fields) ? d.fields : null,
        auction_id: d.auction_id || null, created_at: r.created_at,
      };
    });
    res.json({ items, total, page, pageSize });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Clear the lot activity feed. Scoped to a single trade when auctionId is
// passed (the "Clear All" button on the Lot Entry activity panel only ever
// clears the currently-picked trade). Deletes the matching audit_log rows —
// destructive, so the client guards it behind a confirm(). The WHERE mirrors
// the GET above so exactly the rows the user can see get wiped.
app.delete('/api/lot-activity', requireAuth, (req, res) => {
  try {
    const db = getDb();
    const auctionId = req.query.auctionId ? parseInt(req.query.auctionId, 10) : null;
    let where = "entity = 'lot'";
    const params = [];
    if (auctionId) { where += ' AND details LIKE ?'; params.push(`{"auction_id":${auctionId},%`); }
    const info = db.run(`DELETE FROM audit_log WHERE ${where}`, params);
    res.json({ success: true, deleted: (info && info.changes) || 0 });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ══════════════════════════════════════════════════════════════
// LOT LOCK / UNLOCK
// ══════════════════════════════════════════════════════════════
// Lock = mark a lot record as finalized. Locked lots become uneditable
// for non-admins everywhere they get touched: direct PUT/DELETE on the
// lot itself, every bulk-* helper, the calculate endpoints, and any
// dependent transaction (sales invoice, purchase, debit note) that
// would mutate or invalidate the lot. Only an admin can edit a locked
// lot or clear its lock.
//
// Schema columns:
//   lots.locked_at — ISO timestamp set at lock time, NULL when unlocked
//   lots.locked_by — username of the user who set the lock (audit only;
//                    permission is decided from req.user.role, not this)
//
// Any user with `lot_write` can apply a lock (the natural "I'm confirming
// this row" action). Unlock is admin-only.
//
// Endpoints:
//   POST /api/lots/lock     body: { ids: [int, ...] }   — anyone with lot_write
//   POST /api/lots/unlock   body: { ids: [int, ...] }   — admin only
// Both endpoints are batch: pass [singleId] for the per-row UI button.
function isAdmin(req) {
  return !!(req && req.user && req.user.role === 'admin');
}

// Filter a list of lot ids down to those whose row is NOT locked.
// Used by every bulk lot mutation to skip locked rows silently — see
// the "Skip locked, process the rest" design choice. Returns
// { allowed: [...numericIds], skipped: [...numericIds] }.
function filterLockedLotIds(db, ids) {
  const nums = (ids || []).map(x => Number(x)).filter(Number.isFinite);
  if (!nums.length) return { allowed: [], skipped: [] };
  const placeholders = nums.map(() => '?').join(',');
  const locked = db.all(
    `SELECT id FROM lots WHERE id IN (${placeholders}) AND locked_at IS NOT NULL`,
    nums
  );
  const lockedSet = new Set(locked.map(r => Number(r.id)));
  const allowed = nums.filter(id => !lockedSet.has(id));
  const skipped = nums.filter(id =>  lockedSet.has(id));
  return { allowed, skipped };
}

// Cascade lock checks for dependent transactions. A sales invoice is
// considered "lock-cascaded" if ANY lot in the same auction + buyer
// has locked_at set. Purchase / debit-note use auction + seller name.
// Used to block non-admin PUT/DELETE/revert on dependent rows.
function lotsLockedForInvoice(db, invoiceId) {
  const inv = db.get('SELECT auction_id, buyer FROM invoices WHERE id = ?', [invoiceId]);
  if (!inv) return false;
  const hit = db.get(
    `SELECT 1 FROM lots
       WHERE auction_id = ? AND buyer = ? AND locked_at IS NOT NULL
       LIMIT 1`,
    [inv.auction_id, inv.buyer || '']
  );
  return !!hit;
}
function lotsLockedForPurchase(db, purchaseId) {
  const pur = db.get('SELECT auction_id, name FROM purchases WHERE id = ?', [purchaseId]);
  if (!pur) return false;
  const hit = db.get(
    `SELECT 1 FROM lots
       WHERE auction_id = ? AND name = ? AND locked_at IS NOT NULL
       LIMIT 1`,
    [pur.auction_id, pur.name || '']
  );
  return !!hit;
}
function lotsLockedForDebitNote(db, noteId) {
  const dn = db.get('SELECT auction_id, name FROM debit_notes WHERE id = ?', [noteId]);
  if (!dn) return false;
  const hit = db.get(
    `SELECT 1 FROM lots
       WHERE auction_id = ? AND name = ? AND locked_at IS NOT NULL
       LIMIT 1`,
    [dn.auction_id, dn.name || '']
  );
  return !!hit;
}

app.post('/api/lots/lock', requireLotWrite, (req, res) => {
  const ids = Array.isArray(req.body && req.body.ids) ? req.body.ids : [];
  const numericIds = ids.map(x => Number(x)).filter(Number.isFinite);
  if (!numericIds.length) {
    return res.status(400).json({ error: 'ids[] is required' });
  }
  const db = getDb();
  const username = (req.user && req.user.username) || '';
  const placeholders = numericIds.map(() => '?').join(',');
  // Only lock currently-unlocked rows so the locked_at/locked_by audit
  // pair reflects the most-recent confirm action — re-locking an already-
  // locked row would otherwise overwrite the original locker silently.
  const info = db.run(
    `UPDATE lots
        SET locked_at = datetime('now','localtime'),
            locked_by = ?
      WHERE id IN (${placeholders}) AND locked_at IS NULL`,
    [username, ...numericIds]
  );
  res.json({ success: true, locked: (info && info.changes) || 0, requested: numericIds.length });
});

app.post('/api/lots/unlock', requireAdmin, (req, res) => {
  const ids = Array.isArray(req.body && req.body.ids) ? req.body.ids : [];
  const numericIds = ids.map(x => Number(x)).filter(Number.isFinite);
  if (!numericIds.length) {
    return res.status(400).json({ error: 'ids[] is required' });
  }
  const db = getDb();
  const placeholders = numericIds.map(() => '?').join(',');
  const info = db.run(
    `UPDATE lots SET locked_at = NULL, locked_by = NULL
      WHERE id IN (${placeholders}) AND locked_at IS NOT NULL`,
    numericIds
  );
  res.json({ success: true, unlocked: (info && info.changes) || 0, requested: numericIds.length });
});

// ── Calculate all lots for an auction (GENERATE.PRG) ─────────
app.post('/api/lots/calculate/:auctionId',
  requireLotWrite,
  requirePriceChecked(req => parseInt(req.params.auctionId, 10)),
  (req, res) => {
  const db = getDb(); const cfg = getSettingsFlat(db);
  // Locked lots are treated as finalized — skip them so a recalculate
  // doesn't overwrite the planter-side values a user already confirmed.
  // Admins also skip locked rows here intentionally: recalc is a batch
  // refresh, not a "force-edit one row" action; admins can still hit
  // PUT /api/lots/:id directly if they need to override.
  // Gate on qty*price>0 instead of amount>0 — a price edit through Price
  // Check (bulk-set-buyer only writes `price`) can leave `amount` stale
  // at 0 while qty and the new price are both valid; we want to recalc
  // those rows and heal `amount` rather than skip them.
  //
  // Also pick up rows whose price/amount are now zero but still carry
  // non-zero planter values (e.g. a lot marked WD after its planter
  // math was computed). Without this branch, Calculate All would leave
  // stale prate/puramt/cgst/sgst/igst on withdrawn lots — those values
  // would then leak into bills, purchases, and debit notes.
  const lots = db.all(
    `SELECT * FROM lots
       WHERE auction_id = ?
         AND locked_at IS NULL
         AND ( amount > 0
               OR (qty > 0 AND price > 0)
               OR puramt > 0 OR prate > 0
               OR cgst > 0 OR sgst > 0 OR igst > 0 )`,
    [req.params.auctionId]
  );
  let count = 0;
  for (const lot of lots) {
    const calc = calculateLot(lot, cfg);
    db.run(`UPDATE lots SET amount=?,pqty=?,prate=?,puramt=?,com=?,sertax=?,cgst=?,sgst=?,igst=?,advance=?,balance=?,bilamt=?,refund=?,refud=?,isp_pqty=?,isp_prate=?,isp_puramt=?,asp_pqty=?,asp_prate=?,asp_puramt=? WHERE id=?`,
      [calc.amount,calc.pqty,calc.prate,calc.puramt,calc.com,calc.sertax,calc.cgst,calc.sgst,calc.igst,calc.advance,calc.balance,calc.bilamt,calc.refund||0,calc.refud||0,calc.isp_pqty||0,calc.isp_prate||0,calc.isp_puramt||0,calc.asp_pqty||0,calc.asp_prate||0,calc.asp_puramt||0,lot.id]);
    count++;
  }
  res.json({ success: true, calculated: count });
});

// Recalculate every lot in every auction with the CURRENT business
// settings. Used by the client when business_state changes — calculations
// like CGST/SGST/IGST and prate are state-sensitive (intra vs inter), so
// the saved values become stale on a state flip and must be refreshed.
//
// Picks up lots with `amount > 0` OR `qty*price > 0` so a stale `amount`
// (e.g. after a Price Check price write) doesn't hide a row that has
// otherwise valid qty + price. Also picks up rows whose price/amount
// are now zero but still carry stale planter values (e.g. lots marked
// WD after their planter math was computed) — see the per-auction
// calculate above for the same rationale. Returns total lots
// calculated across all auctions.
app.post('/api/lots/calculate-all', requireLotWrite, (req, res) => {
  const db = getDb(); const cfg = getSettingsFlat(db);
  // Skip locked lots (same rationale as the per-auction calculate above).
  const lots = db.all(
    `SELECT * FROM lots
       WHERE locked_at IS NULL
         AND ( amount > 0
               OR (qty > 0 AND price > 0)
               OR puramt > 0 OR prate > 0
               OR cgst > 0 OR sgst > 0 OR igst > 0 )`
  );
  let count = 0;
  for (const lot of lots) {
    const calc = calculateLot(lot, cfg);
    db.run(`UPDATE lots SET amount=?,pqty=?,prate=?,puramt=?,com=?,sertax=?,cgst=?,sgst=?,igst=?,advance=?,balance=?,bilamt=?,refund=?,refud=?,isp_pqty=?,isp_prate=?,isp_puramt=?,asp_pqty=?,asp_prate=?,asp_puramt=? WHERE id=?`,
      [calc.amount,calc.pqty,calc.prate,calc.puramt,calc.com,calc.sertax,calc.cgst,calc.sgst,calc.igst,calc.advance,calc.balance,calc.bilamt,calc.refund||0,calc.refud||0,calc.isp_pqty||0,calc.isp_prate||0,calc.isp_puramt||0,calc.asp_pqty||0,calc.asp_prate||0,calc.asp_puramt||0,lot.id]);
    count++;
  }
  res.json({ success: true, calculated: count });
});

// ── Data validation (PRICHECK.PRG) ───────────────────────────
app.get('/api/lots/validate/:auctionId', requireViewOrLotEntry, (req, res) => {
  const rows = getDb().all(
    `SELECT * FROM lots WHERE auction_id = ? AND (price = 0 OR amount = 0 OR buyer = '' OR code = '' OR ROUND(qty*price,2) <> ROUND(amount,2))`,
    [req.params.auctionId]);
  res.json(rows);
});

// Bulk grade update — paired with the Lots-screen "Set Grade" button.
// Body: { ids: [1, 2, …], grade: '1' }
// Grade is whitelisted to a short fixed set so a malformed client can't
// poison the column. Empty grade is allowed (clears it). Returns the
// count of updated rows.
app.post('/api/lots/bulk-grade', requireLotWrite, (req, res) => {
  try {
    const { ids, grade } = req.body || {};
    if (!Array.isArray(ids) || ids.length === 0) {
      return res.status(400).json({ error: 'ids[] is required' });
    }
    const validGrades = new Set(['', '1', '2', '3']);
    const g = String(grade == null ? '' : grade).trim();
    if (!validGrades.has(g)) {
      return res.status(400).json({ error: `grade must be one of: ${[...validGrades].map(v => v || '(blank)').join(', ')}` });
    }
    const numericIds = ids.map(x => Number(x)).filter(Number.isFinite);
    if (!numericIds.length) {
      return res.status(400).json({ error: 'ids[] contains no valid numeric ids' });
    }
    const db = getDb();
    const { allowed, skipped } = filterLockedLotIds(db, numericIds);
    if (!allowed.length) {
      return res.json({ success: true, updated: 0, grade: g, skipped_locked: skipped.length });
    }
    const placeholders = allowed.map(() => '?').join(',');
    db.run(`UPDATE lots SET grade = ? WHERE id IN (${placeholders})`, [g, ...allowed]);
    res.json({ success: true, updated: allowed.length, grade: g, skipped_locked: skipped.length });
  } catch (e) {
    res.status(500).json({ error: 'Bulk grade update failed: ' + (e.message || e) });
  }
});

// Bulk buyer-code update — paired with the Lots-screen "Set Buyer"
// button. Body: { ids: [1, 2, …], buyer: 'CODE' }
// Resolves the buyer in the buyers table to fetch buyer1 (trade name)
// and code (short code) so all three columns on the lot stay in sync
// — otherwise the lot would carry the new buyer code with a stale
// trade name from the previous buyer, which breaks invoice generation
// and Tally export party lookups. Match is case-insensitive on
// `buyer` to tolerate uppercase/lowercase entries in legacy data.
// /api/lots/bulk-set-buyer — flexible bulk update used by both the
// generic "set buyer code" bulk action and the Price Check "Apply"
// flow. Takes an `ids` array plus any combination of `code` / `buyer`
// / `buyer1` / `sale` / `price` to write. Empty / undefined fields are
// left untouched, so a caller can update just `price` (Price Check's
// price-only fix) or just `code` (post-mapping cleanup) without
// blasting unrelated fields. Server also clears the price-check stamp
// on every touched auction so the next verify run picks up the change.
app.post('/api/lots/bulk-set-buyer', requireLotWrite, (req, res) => {
  const ids = Array.isArray(req.body.ids) ? req.body.ids.map(Number).filter(Number.isFinite) : [];
  if (!ids.length) return res.status(400).json({ error: 'No lot ids provided' });
  const code   = String(req.body.code   || '').trim();
  const buyer  = String(req.body.buyer  || '').trim();
  const buyer1 = String(req.body.buyer1 || '').trim();
  const sale   = String(req.body.sale   || '').trim();
  const hasPrice = req.body.price !== undefined && req.body.price !== null && req.body.price !== '';
  const priceNum = hasPrice ? Number(req.body.price) : null;
  if (hasPrice && !Number.isFinite(priceNum)) {
    return res.status(400).json({ error: 'price must be a number' });
  }
  if (!code && !hasPrice) {
    return res.status(400).json({ error: 'At least one of code or price is required' });
  }
  const db = getDb();
  // Drop any locked rows up-front so the chunked UPDATE below sees only
  // mutable ids — matches the "skip locked, process the rest" behaviour
  // of bulk-grade / bulk-buyer / bulk-seller.
  const { allowed: mutableIds, skipped: lockedIds } = filterLockedLotIds(db, ids);
  const sets = [];
  const vals = [];
  if (code)   { sets.push('code = ?');   vals.push(code);   }
  if (buyer)  { sets.push('buyer = ?');  vals.push(buyer);  }
  if (buyer1) { sets.push('buyer1 = ?'); vals.push(buyer1); }
  if (req.body.sale !== undefined) { sets.push('sale = ?'); vals.push(sale); }
  if (hasPrice) { sets.push('price = ?'); vals.push(priceNum); }
  const CHUNK = 500;
  let updated = 0;
  // Capture every auction touched so we can clear their price-check
  // stamps — codes/prices just changed, so an earlier verify is stale.
  const touchedAuctions = new Set();
  for (let i = 0; i < mutableIds.length; i += CHUNK) {
    const slice = mutableIds.slice(i, i + CHUNK);
    const placeholders = slice.map(() => '?').join(',');
    const affectedRows = db.all(
      `SELECT DISTINCT auction_id FROM lots WHERE id IN (${placeholders})`,
      slice
    );
    affectedRows.forEach(r => { if (r.auction_id) touchedAuctions.add(r.auction_id); });
    const info = db.run(
      `UPDATE lots SET ${sets.join(', ')} WHERE id IN (${placeholders})`,
      [...vals, ...slice]
    );
    if (info && typeof info.changes === 'number') updated += info.changes;
  }
  for (const aid of touchedAuctions) { pcClearGate(db, aid); lvClearGate(db, aid); }
  res.json({ success: true, updated, requested: ids.length, skipped_locked: lockedIds.length });
});

app.post('/api/lots/bulk-buyer', requireLotWrite, (req, res) => {
  try {
    const { ids, buyer } = req.body || {};
    if (!Array.isArray(ids) || ids.length === 0) {
      return res.status(400).json({ error: 'ids[] is required' });
    }
    const buyerCode = String(buyer == null ? '' : buyer).trim();
    if (!buyerCode) {
      return res.status(400).json({ error: 'buyer code is required' });
    }
    const numericIds = ids.map(x => Number(x)).filter(Number.isFinite);
    if (!numericIds.length) {
      return res.status(400).json({ error: 'ids[] contains no valid numeric ids' });
    }
    const db = getDb();
    // Special sentinel: WD = withdrawn. Not a real buyer in the master,
    // so skip the lookup and stamp the withdrawn marker + zero the price
    // (amount is healed to 0 by the follow-up Calculate the client runs).
    // Mirrors the single-lot edit modal's WD handling.
    if (buyerCode.toUpperCase() === 'WD') {
      const { allowed: mutableIds, skipped: lockedIds } = filterLockedLotIds(db, numericIds);
      if (!mutableIds.length) {
        return res.json({
          success: true, updated: 0,
          buyer: 'WD', buyer1: 'Withdrawn', code: 'WD', sale: 'W',
          skipped_locked: lockedIds.length,
        });
      }
      const placeholders = mutableIds.map(() => '?').join(',');
      const touchedAuctions = new Set();
      const affectedRows = db.all(
        `SELECT DISTINCT auction_id FROM lots WHERE id IN (${placeholders})`,
        mutableIds
      );
      affectedRows.forEach(r => { if (r.auction_id) touchedAuctions.add(r.auction_id); });
      db.run(
        `UPDATE lots SET buyer = ?, buyer1 = ?, code = ?, sale = ?, price = 0, amount = 0 WHERE id IN (${placeholders})`,
        ['WD', 'Withdrawn', 'WD', 'W', ...mutableIds]
      );
      for (const aid of touchedAuctions) { pcClearGate(db, aid); lvClearGate(db, aid); }
      return res.json({
        success: true,
        updated: mutableIds.length,
        buyer: 'WD', buyer1: 'Withdrawn', code: 'WD', sale: 'W',
        skipped_locked: lockedIds.length,
      });
    }
    const b = db.get(
      `SELECT buyer, buyer1, code, sale FROM buyers
        WHERE UPPER(TRIM(buyer)) = UPPER(TRIM(?))
        LIMIT 1`,
      [buyerCode]
    );
    if (!b) {
      return res.status(404).json({ error: `No buyer found with code "${buyerCode}". Register the buyer first in the Buyers tab.` });
    }
    // Apply the buyer's default sale type (L/I/E) onto the lot so the
    // GST regime on subsequent invoice generation matches what the buyer
    // is registered for — otherwise a Local buyer reassigned from an
    // Inter-state lot would still carry the IGST flag from the prior
    // assignment.
    const buyerSale = String(b.sale || '').trim().toUpperCase();
    const saleVal = ['L', 'I', 'E'].includes(buyerSale) ? buyerSale : 'L';
    const { allowed: mutableIds, skipped: lockedIds } = filterLockedLotIds(db, numericIds);
    if (!mutableIds.length) {
      return res.json({
        success: true, updated: 0,
        buyer: b.buyer, buyer1: b.buyer1 || '', code: b.code || '', sale: saleVal,
        skipped_locked: lockedIds.length,
      });
    }
    const placeholders = mutableIds.map(() => '?').join(',');
    // Capture which auctions are affected so we can clear their price-
    // check gates — the buyer codes have changed, so an earlier verify
    // run no longer reflects the current state.
    const touchedAuctions = new Set();
    const affectedRows = db.all(
      `SELECT DISTINCT auction_id FROM lots WHERE id IN (${placeholders})`,
      mutableIds
    );
    affectedRows.forEach(r => { if (r.auction_id) touchedAuctions.add(r.auction_id); });
    db.run(
      `UPDATE lots SET buyer = ?, buyer1 = ?, code = ?, sale = ? WHERE id IN (${placeholders})`,
      [b.buyer, b.buyer1 || '', b.code || '', saleVal, ...mutableIds]
    );
    for (const aid of touchedAuctions) { pcClearGate(db, aid); lvClearGate(db, aid); }
    res.json({
      success: true,
      updated: mutableIds.length,
      buyer: b.buyer,
      buyer1: b.buyer1 || '',
      code: b.code || '',
      sale: saleVal,
      skipped_locked: lockedIds.length,
    });
  } catch (e) {
    res.status(500).json({ error: 'Bulk buyer update failed: ' + (e.message || e) });
  }
});

// Bulk seller-reassign — paired with the Lot Entry "Change Seller"
// action. Body: { ids: [1, 2, …], trader_id: 42 }
// Resolves the trader once and updates every selected lot's
// trader_id + all denormalised seller columns (name, cr, pan, tel,
// padd, ppla, ppin, pstate, pst_code, aadhar). Without this, a lot
// would keep its previous seller's name in `lots.name` even though
// `trader_id` was changed — that mismatch breaks invoice/bill
// generation, exports, and Tally XML (which read directly from the
// denormalised columns).
app.post('/api/lots/bulk-seller', requireLotWrite, (req, res) => {
  try {
    const { ids, trader_id } = req.body || {};
    if (!Array.isArray(ids) || ids.length === 0) {
      return res.status(400).json({ error: 'ids[] is required' });
    }
    const tid = parseInt(trader_id, 10);
    if (!Number.isFinite(tid)) {
      return res.status(400).json({ error: 'trader_id must be a numeric id' });
    }
    const numericIds = ids.map(x => Number(x)).filter(Number.isFinite);
    if (!numericIds.length) {
      return res.status(400).json({ error: 'ids[] contains no valid numeric ids' });
    }
    const db = getDb();
    const t = db.get('SELECT * FROM traders WHERE id = ?', [tid]);
    if (!t) {
      return res.status(404).json({ error: `No seller found with id ${tid}. Refresh the seller list and try again.` });
    }
    const { allowed: mutableIds, skipped: lockedIds } = filterLockedLotIds(db, numericIds);
    if (!mutableIds.length) {
      return res.json({
        success: true, updated: 0, trader_id: tid, name: t.name || '',
        skipped_locked: lockedIds.length,
      });
    }
    const placeholders = mutableIds.map(() => '?').join(',');
    // Mirror the same column set POST/PUT /api/lots populate when a
    // trader_id is supplied — keeps the source of truth single.
    db.run(
      `UPDATE lots SET
         trader_id = ?,
         name      = ?,
         cr        = ?,
         pan       = ?,
         tel       = ?,
         aadhar    = ?,
         padd      = ?,
         ppla      = ?,
         ppin      = ?,
         pstate    = ?,
         pst_code  = ?
       WHERE id IN (${placeholders})`,
      [
        tid,
        t.name || '',
        t.cr || '',
        t.pan || '',
        t.tel || '',
        t.aadhar || '',
        t.padd || '',
        t.ppla || '',
        t.pin || '',
        t.pstate || '',
        t.pst_code || '',
        ...mutableIds,
      ]
    );
    res.json({
      success: true,
      updated: mutableIds.length,
      trader_id: tid,
      name: t.name || '',
      skipped_locked: lockedIds.length,
    });
  } catch (e) {
    res.status(500).json({ error: 'Bulk seller update failed: ' + (e.message || e) });
  }
});

// ══════════════════════════════════════════════════════════════
// INVOICES — Sales (GSTIN.PRG / KGSTIN.PRG)
// ══════════════════════════════════════════════════════════════
app.get('/api/invoices', requireView, (req, res) => {
  const { ano, auction_id, from, to, sale, search } = req.query;
  const db = getDb();
  const cfg = getSettingsFlat(db);
  // Filter list by active business context: when state=KERALA show only
  // ASP-stamped invoices, when state=TAMIL NADU show only ISP-stamped.
  // This avoids the "two rows per buyer" confusion in users who run the
  // ASP→ISP flow on the same auction.
  const businessState = String(cfg.business_state || 'TAMIL NADU').toUpperCase();
  let q = 'SELECT * FROM invoices WHERE 1=1'; const p = [];
  if (auction_id) { q += ' AND auction_id = ?'; p.push(parseInt(auction_id)); }
  if (ano) { q += ' AND ano = ?'; p.push(ano); }
  // Free-text search within the selected trade — invoice no, buyer
  // code, trade name, GSTIN, and lorry/vehicle no. Trade no isn't
  // included because the user already picked one in the auction dropdown.
  const searchTerm = String(search || '').trim();
  if (searchTerm) {
    const wild = `%${searchTerm}%`;
    q += ` AND (
            COALESCE(invo,'')     LIKE ?
            OR COALESCE(buyer,'')    LIKE ?
            OR COALESCE(buyer1,'')   LIKE ?
            OR COALESCE(gstin,'')    LIKE ?
            OR COALESCE(lorry_no,'') LIKE ?
          )`;
    p.push(wild, wild, wild, wild, wild);
  }
  if (from && to) { q += ' AND date BETWEEN ? AND ?'; p.push(from, to); }
  // Sale-type filter — L=Local, I=Inter-state, E=Export. Whitelisted
  // to those three values; anything else is ignored (so a malformed
  // client doesn't poison the WHERE clause). The `sale` column is
  // populated at invoice generation from the underlying lots.
  if (sale && ['L','I','E'].includes(String(sale).toUpperCase())) {
    q += ' AND UPPER(sale) = ?';
    p.push(String(sale).toUpperCase());
  }
  // e-Trade-only build: every invoice belongs to the single company
  // (ISP), regardless of what's stamped in the `state` column. The
  // original Spice Config app used `state` to split rows between ISP
  // (TN) and ASP (KL) views — that split no longer exists. So no
  // state filter here; anything for the requested auction_id ships
  // back. (`businessState` is left unused but kept above so future
  // callers can reintroduce a filter without restructuring.)
  q += ' ORDER BY date DESC, invo DESC LIMIT 500';
  const rows = db.all(q, p);
  // Hydrate asp_invo: for each invoice, find the ASP invoice number
  // recorded on its lots. Multiple distinct asp_invos are concatenated
  // (rare — usually one ASP invoice maps 1:1 to one ISP invoice for the
  // same buyer/auction). Empty for ASP invoices themselves.
  const aspStmt = db.prepare(
    `SELECT DISTINCT asp_invo FROM lots
     WHERE auction_id = ? AND buyer = ? AND invo = ?
       AND asp_invo IS NOT NULL AND asp_invo != ''`
  );
  for (const r of rows) {
    // For ASP invoices (state contains "Kerala"), the asp_invo column
    // would just be a copy of `invo` — show blank instead of duplicating.
    const isASPRow = String(r.state || '').toLowerCase().includes('kerala');
    if (isASPRow) { r.asp_invo = ''; }
    else {
      const aspRows = aspStmt.all(r.auction_id, r.buyer, r.invo);
      r.asp_invo = aspRows.map(x => x.asp_invo).filter(Boolean).join(', ');
    }
  }

  // Hydrate the resolved e-way bill distance per invoice. Resolution
  // order matches the Tally generator (so what the user sees in the
  // Sales tab is exactly what ships in the voucher):
  //   1. invoices.distance_km — per-invoice override, wins if set
  //   2. route_distances[(dispatch_pin, buyer_pin)] — saved route value
  //   3. blank
  //
  // Done in JS rather than a JOIN because the (dispatch_pin → buyer_pin)
  // pair varies per invoice (each buyer has a different PIN). A single
  // JOIN would need a CTE for normalised pairs; cleaner to look up.
  try {
    const dispatchPin = String(
      cfg.tally_dispatch_pin ||
      cfg.s_pin ||
      cfg.kl_pin ||
      cfg.tn_pin ||
      ''
    ).trim();
    const routeStmt = db.prepare(
      // Routes are stored under a normalised (min, max) pair, so a
      // lookup must match the OR of both directions. Without this the
      // lookup would silently miss every route stored in the opposite
      // order.
      `SELECT km FROM route_distances
       WHERE (from_pin = ? AND to_pin = ?)
          OR (from_pin = ? AND to_pin = ?)
       LIMIT 1`
    );
    // Per spec: SOURCE PINCODE PRIORITY
    //   1. ship-to (consignee) PIN — buyers.cpin
    //   2. bill-to (buyer) PIN     — buyers.pin
    // The ship-to address is where the goods physically arrive (and
    // therefore the correct e-way bill destination); bill-to is the
    // legal billing address, which can differ for buyers with multiple
    // delivery sites. Distance must be computed to the actual ship-to
    // when one is registered. Falls back to bill-to so legacy buyers
    // without a separate consignee block still resolve.
    const buyerPinStmt = db.prepare('SELECT pin, cpin FROM buyers WHERE buyer = ? LIMIT 1');
    for (const r of rows) {
      // Per-invoice override always wins
      if (r.distance_km != null && r.distance_km !== '') {
        r.resolved_distance_km = Number(r.distance_km);
        continue;
      }
      // Route table lookup: need the buyer's PIN
      const b = buyerPinStmt.get(r.buyer);
      const shipPin = b && b.cpin ? String(b.cpin).trim() : '';
      const billPin = b && b.pin  ? String(b.pin ).trim() : '';
      const buyerPin = shipPin || billPin;
      if (!buyerPin || !dispatchPin) { r.resolved_distance_km = null; continue; }
      const hit = routeStmt.get(dispatchPin, buyerPin, buyerPin, dispatchPin);
      r.resolved_distance_km = hit && hit.km != null ? Number(hit.km) : null;
    }
  } catch (e) {
    // route_distances table may be missing on a partially-migrated DB.
    // Don't fail the whole endpoint — just leave resolved_distance_km null.
    console.warn('[invoices] resolved distance hydration failed:', e.message);
  }
  res.json(rows);
});

app.post('/api/invoices/generate/:auctionId',
  requireInvoiceWrite,
  requirePriceChecked(req => parseInt(req.params.auctionId, 10)),
  (req, res) => {
  const db = getDb(); const cfg = getSettingsFlat(db);
  const auctionIdForGate = parseInt(req.params.auctionId, 10);
  const _gen = _checkGenerationGate(db, 'invoices', auctionIdForGate);
  if (!_gen.allowed) return res.status(412).json(_gen.error);
  const { saleType, buyerCode, invoiceNo } = req.body;
  // Per-invoice "No Transport & Insurance" — when set, this invoice is
  // generated with transport + insurance forced to 0.
  const noTI = (req.body.noTI === true || String(req.body.noTI || '').toLowerCase() === 'true' || Number(req.body.noTI) === 1) ? 1 : 0;

  if (!saleType || !buyerCode || !invoiceNo) {
    return res.status(400).json({ error: 'saleType, buyerCode, and invoiceNo are required' });
  }
  
  // Auto-calculate lots if puramt is missing (user might not have clicked Calculate)
  // Locked lots are skipped — recalculating would overwrite the planter-side
  // values they were confirmed at. The buildSalesInvoice call below also
  // filters locked rows, so this is consistent.
  const uncalc = db.all(`SELECT * FROM lots WHERE auction_id = ? AND amount > 0 AND (puramt IS NULL OR puramt = 0) AND locked_at IS NULL`, [req.params.auctionId]);
  for (const lot of uncalc) {
    const c = calculateLot(lot, cfg);
    db.run(`UPDATE lots SET pqty=?,prate=?,puramt=?,com=?,sertax=?,cgst=?,sgst=?,igst=?,advance=?,balance=?,bilamt=?,refund=?,refud=?,isp_pqty=?,isp_prate=?,isp_puramt=?,asp_pqty=?,asp_prate=?,asp_puramt=? WHERE id=?`,
      [c.pqty,c.prate,c.puramt,c.com,c.sertax,c.cgst,c.sgst,c.igst,c.advance,c.balance,c.bilamt,c.refund||0,c.refud||0,c.isp_pqty||0,c.isp_prate||0,c.isp_puramt||0,c.asp_pqty||0,c.asp_prate||0,c.asp_puramt||0,lot.id]);
  }

  const invoice = buildSalesInvoice(db, req.params.auctionId, buyerCode, saleType, cfg, { noTI });
  if (!invoice) return res.status(404).json({ error: `No lots found for buyer "${buyerCode}" in this auction. Make sure lots have this buyer code assigned.` });

  const auction = db.get('SELECT * FROM auctions WHERE id = ?', [req.params.auctionId]);
  // Dedup guards — refuse to create a second invoice for the same buyer
  // + sale type in this trade, or to reuse an invoice number that's
  // already on file for this trade. Both checks return 409 so the
  // caller can show a precise reason.
  const dupBuyer = db.get(
    'SELECT id, invo FROM invoices WHERE auction_id = ? AND sale = ? AND buyer = ? LIMIT 1',
    [req.params.auctionId, saleType, buyerCode]
  );
  if (dupBuyer) {
    return res.status(409).json({
      error: `Invoice already exists for buyer "${buyerCode}" (${saleType}) in this trade — invoice #${dupBuyer.invo}.`,
      existingId: dupBuyer.id, existingInvo: dupBuyer.invo,
    });
  }
  const dupNo = db.get(
    'SELECT id, buyer FROM invoices WHERE auction_id = ? AND sale = ? AND invo = ? LIMIT 1',
    [req.params.auctionId, saleType, String(invoiceNo)]
  );
  if (dupNo) {
    return res.status(409).json({
      error: `Invoice number ${invoiceNo} (${saleType}) is already used in this trade by buyer "${dupNo.buyer}".`,
      existingId: dupNo.id, existingBuyer: dupNo.buyer,
    });
  }
  const s = invoice.summary;
  // Store the BUSINESS context state (TAMIL NADU=ISP, KERALA=ASP), not
  // the auction's physical state. This lets us distinguish ASP invoices
  // from ISP invoices in the same auction, which matters for the sales
  // list cross-reference (ASP Inv# column).
  const invoiceState = cfg.business_state || auction.state || '';
  db.run(`INSERT INTO invoices (auction_id,ano,date,state,sale,invo,buyer,buyer1,gstin,place,bag,qty,amount,gunny,pava_hc,ins,cgst,sgst,igst,tcs,rund,tot,addl_chg,addl_name,no_ti)
    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`,
    [req.params.auctionId,auction.ano,auction.date,invoiceState,saleType,String(invoiceNo),buyerCode,invoice.buyer.buyer1||'',
     invoice.buyer.gstin||'',invoice.buyer.pla||'',s.totalBags,s.totalQty,s.totalAmount,s.gunnyCost,s.transportCost,s.insuranceCost,
     s.cgst,s.sgst,s.igst,0,s.roundDiff,s.grandTotal,s.addlCharge||0,s.addlChargeName||'',noTI]);

  // Update lots with sale type and invoice number.
  // Workflow trace:
  //   - In Kerala (ASP) context: set `invo` AND `asp_invo` to the new ASP
  //     invoice number. DON'T update `lots.sale` — sale type is determined
  //     by the ISP→external transaction (could be Local/Inter-state/Export
  //     depending on buyer's GST state). ASP→ISP is a fixed intra-Kerala
  //     transfer, so any `sale` value would constrain the later ISP step.
  //     EXCEPT: if `invo` already holds a non-ASP value (i.e., an ISP
  //     invoice was already generated), preserve `invo` and only refresh
  //     `asp_invo`. This prevents accidentally destroying the ISP invoice
  //     number when re-running ASP after the full ASP→ISP cycle.
  //   - In Tamil Nadu (ISP) context: update `sale` and `invo` — `asp_invo`
  //     retains the prior ASP number from the earlier ASP-state generation.
  const isASPState = String(cfg.business_state || '').toUpperCase() === 'KERALA';
  for (const li of invoice.lineItems) {
    if (isASPState) {
      const existing = db.get(
        'SELECT invo, asp_invo FROM lots WHERE auction_id=? AND lot_no=? AND buyer=? LIMIT 1',
        [req.params.auctionId, li.lot, buyerCode]
      );
      const hasIspInvo = existing && existing.invo && existing.invo !== existing.asp_invo;
      if (hasIspInvo) {
        // Preserve invo (ISP); refresh only asp_invo. Sale stays as set by ISP step.
        db.run('UPDATE lots SET asp_invo=? WHERE auction_id=? AND lot_no=? AND buyer=?',
          [String(invoiceNo), req.params.auctionId, li.lot, buyerCode]);
      } else {
        // First-time ASP: don't touch `sale` so ISP step has a clean slate
        db.run('UPDATE lots SET invo=?, asp_invo=? WHERE auction_id=? AND lot_no=? AND buyer=?',
          [String(invoiceNo), String(invoiceNo), req.params.auctionId, li.lot, buyerCode]);
      }
    } else {
      db.run('UPDATE lots SET sale=?, invo=? WHERE auction_id=? AND lot_no=? AND buyer=? AND locked_at IS NULL',
        [saleType, String(invoiceNo), req.params.auctionId, li.lot, buyerCode]);
    }
  }
  res.json({ success: true, invoice: invoice.summary });
});

// List eligible buyers for an auction (distinct buyers with lots in amount > 0)
app.get('/api/invoices/eligible-buyers/:auctionId', requireView, (req, res) => {
  const { saleType } = req.query;
  const db = getDb();
  const cfg = getSettingsFlat(db);
  const params = [req.params.auctionId];

  // Match buyers by sale type via their default (b.sale) when a type is specified.
  // A buyer is eligible when any of their lots in this auction isn't yet invoiced
  // for the current state context (so user can always see/regenerate; server
  // endpoint has stricter filter).
  let saleClause = '';
  if (saleType) {
    saleClause = ` AND (COALESCE(NULLIF(l.sale,''), b.sale, 'L') = ?)`;
    params.push(saleType);
  }

  // State-aware eligibility:
  //   - In Kerala (ASP) context: a lot is eligible if no `invo` set yet,
  //     OR if `invo == asp_invo` (i.e., it was previously invoiced in
  //     ASP — user is regenerating).
  //   - In Tamil Nadu (ISP) context: a lot is eligible if `invo` is empty
  //     OR if `invo == asp_invo` (lot only has its ASP invoice, still
  //     needs ISP invoicing). This is the key case that was broken.
  // In both states, lots with `invo != asp_invo AND invo != ''` are
  // considered "fully invoiced" for the current state and excluded.
  const isASPState = String(cfg.business_state || '').toUpperCase() === 'KERALA';
  // Both states share the same eligibility expression — what differs is
  // the meaning. The expression: lot is eligible if no `invo` OR `invo
  // matches asp_invo` (meaning the only existing invoice on this lot is
  // an ASP one, which doesn't count toward "ISP-invoiced" status).
  const eligibleExpr = isASPState
    ? `(l.invo IS NULL OR l.invo = '')`
    : `(l.invo IS NULL OR l.invo = '' OR (l.asp_invo IS NOT NULL AND l.asp_invo != '' AND l.invo = l.asp_invo))`;

  res.json(db.all(
    `SELECT l.buyer, COALESCE(b.buyer1, MAX(l.buyer1), l.buyer) as buyer1,
        b.code as code,
        COUNT(*) as lot_count, SUM(l.qty) as total_qty, SUM(l.amount) as total_amount,
        b.gstin, b.sale as default_sale
     FROM lots l
     LEFT JOIN buyers b ON b.buyer = l.buyer
     WHERE l.auction_id = ?
       AND l.buyer IS NOT NULL AND l.buyer != ''
       AND l.locked_at IS NULL
       ${saleClause}
     GROUP BY l.buyer
     HAVING COUNT(CASE WHEN ${eligibleExpr} THEN 1 END) > 0
     ORDER BY l.buyer`,
    params
  ));
});

// ── Diagnostic: show EVERYTHING about buyers in an auction ──
// Helps troubleshoot why eligible-buyers returns an unexpected count.
app.get('/api/invoices/eligibility-debug/:auctionId', requireView, (req, res) => {
  const db = getDb();
  const aid = req.params.auctionId;
  const auction = db.get('SELECT * FROM auctions WHERE id = ?', [aid]);
  if (!auction) return res.status(404).json({ error: 'Auction not found' });

  // Every distinct value in lots.buyer (including blanks), with counts
  const allBuyerGroups = db.all(
    `SELECT
       COALESCE(NULLIF(TRIM(l.buyer),''), '<BLANK>') as buyer_raw,
       COUNT(*) as total_lots,
       COUNT(CASE WHEN l.invo IS NULL OR l.invo = '' THEN 1 END) as uninvoiced_lots,
       COUNT(CASE WHEN l.amount > 0 THEN 1 END) as priced_lots,
       SUM(l.amount) as total_amount,
       MAX(l.buyer1) as lot_buyer1,
       (SELECT buyer1 FROM buyers WHERE buyer = l.buyer LIMIT 1) as master_buyer1,
       (SELECT sale    FROM buyers WHERE buyer = l.buyer LIMIT 1) as master_sale,
       (SELECT gstin   FROM buyers WHERE buyer = l.buyer LIMIT 1) as master_gstin,
       (SELECT id      FROM buyers WHERE buyer = l.buyer LIMIT 1) as master_id
     FROM lots l
     WHERE l.auction_id = ?
     GROUP BY TRIM(l.buyer)
     ORDER BY total_lots DESC`,
    [aid]
  );

  const total_lots      = db.get('SELECT COUNT(*) as c FROM lots WHERE auction_id = ?', [aid]).c;
  const lots_no_buyer   = db.get(`SELECT COUNT(*) as c FROM lots WHERE auction_id = ? AND (buyer IS NULL OR TRIM(buyer) = '')`, [aid]).c;
  const lots_invoiced   = db.get(`SELECT COUNT(*) as c FROM lots WHERE auction_id = ? AND invo IS NOT NULL AND invo != ''`, [aid]).c;
  const distinct_buyers_in_lots = db.get(
    `SELECT COUNT(*) as c FROM (
       SELECT DISTINCT TRIM(buyer) as b FROM lots
       WHERE auction_id = ? AND buyer IS NOT NULL AND TRIM(buyer) != ''
     )`, [aid]).c;

  res.json({
    auction: { id: auction.id, ano: auction.ano, date: auction.date, crop_type: auction.crop_type },
    totals: {
      total_lots,
      lots_with_blank_buyer: lots_no_buyer,
      lots_already_invoiced: lots_invoiced,
      distinct_buyer_codes_in_lots: distinct_buyers_in_lots,
      buyers_table_total: db.get('SELECT COUNT(*) as c FROM buyers').c,
    },
    breakdown: allBuyerGroups.map(r => ({
      buyer_code: r.buyer_raw,
      master_match: r.master_id ? 'yes' : 'NO — not in buyers table',
      master_buyer1: r.master_buyer1 || null,
      lot_buyer1:    r.lot_buyer1 || null,
      master_sale:   r.master_sale || null,
      total_lots:      r.total_lots,
      uninvoiced_lots: r.uninvoiced_lots,
      priced_lots:     r.priced_lots,
      total_amount:    r.total_amount,
      eligible: r.buyer_raw !== '<BLANK>' && r.uninvoiced_lots > 0 ? 'yes' : 'NO',
      eligibility_reason: r.buyer_raw === '<BLANK>' ? 'buyer code is blank'
        : r.uninvoiced_lots === 0 ? 'all lots already invoiced'
        : 'eligible'
    }))
  });
});

// Batch: generate sales invoice for ALL buyers in an auction
app.post('/api/invoices/generate-all/:auctionId',
  requireInvoiceWrite,
  requirePriceChecked(req => parseInt(req.params.auctionId, 10)),
  (req, res) => {
  const db = getDb(); const cfg = getSettingsFlat(db);
  const auctionIdForGate = parseInt(req.params.auctionId, 10);
  const _gen = _checkGenerationGate(db, 'invoices', auctionIdForGate);
  if (!_gen.allowed) return res.status(412).json(_gen.error);
  const { startInvoiceNo, saleType } = req.body;
  // Per-invoice "No Transport & Insurance" — applies to every invoice in
  // this bulk run when set.
  const noTI = (req.body.noTI === true || String(req.body.noTI || '').toLowerCase() === 'true' || Number(req.body.noTI) === 1) ? 1 : 0;

  let nextNo = parseInt(startInvoiceNo);
  if (!nextNo || nextNo < 1) return res.status(400).json({ error: 'startInvoiceNo must be a positive integer' });
  
  // Auto-calculate uncalculated lots — skip locked rows (recalc would
  // overwrite the planter-side values they were finalised at).
  const uncalc = db.all(`SELECT * FROM lots WHERE auction_id = ? AND amount > 0 AND (puramt IS NULL OR puramt = 0) AND locked_at IS NULL`, [req.params.auctionId]);
  for (const lot of uncalc) {
    const c = calculateLot(lot, cfg);
    db.run(`UPDATE lots SET pqty=?,prate=?,puramt=?,com=?,sertax=?,cgst=?,sgst=?,igst=?,advance=?,balance=?,bilamt=?,refund=?,refud=?,isp_pqty=?,isp_prate=?,isp_puramt=?,asp_pqty=?,asp_prate=?,asp_puramt=? WHERE id=?`,
      [c.pqty,c.prate,c.puramt,c.com,c.sertax,c.cgst,c.sgst,c.igst,c.advance,c.balance,c.bilamt,c.refund||0,c.refud||0,c.isp_pqty||0,c.isp_prate||0,c.isp_puramt||0,c.asp_pqty||0,c.asp_prate||0,c.asp_puramt||0,lot.id]);
  }
  
  // Get distinct buyers. When saleType filter is set, only buyers whose
  // default sale matches (or whose lots already have that sale assigned) are included.
  // The "un-invoiced" check is state-aware:
  //   - In Tamil Nadu (ISP): a lot is un-invoiced if `invo` is empty OR
  //     if the only existing invoice is the ASP one (invo == asp_invo).
  //   - In Kerala (ASP): un-invoiced means `invo` is empty.
  const isASPState = String(cfg.business_state || '').toUpperCase() === 'KERALA';
  const uninvoicedExpr = isASPState
    ? `(l.invo IS NULL OR l.invo = '')`
    : `(l.invo IS NULL OR l.invo = '' OR (l.asp_invo IS NOT NULL AND l.asp_invo != '' AND l.invo = l.asp_invo))`;
  const params = [req.params.auctionId];
  let saleClause = '';
  if (saleType) {
    saleClause = ` AND (COALESCE(NULLIF(l.sale,''), b.sale, 'L') = ?)`;
    params.push(saleType);
  }
  const buyers = db.all(
    `SELECT DISTINCT l.buyer, b.sale as default_sale
     FROM lots l LEFT JOIN buyers b ON b.buyer = l.buyer
     WHERE l.auction_id = ? AND l.buyer IS NOT NULL AND l.buyer != '' AND l.amount > 0
       AND l.locked_at IS NULL
       AND ${uninvoicedExpr}
       ${saleClause}`,
    params
  );
  
  if (!buyers.length) return res.status(404).json({ error: saleType ? `No un-invoiced buyers for sale type ${saleType}` : 'No un-invoiced buyers with lots in this auction' });
  
  const auction = db.get('SELECT * FROM auctions WHERE id = ?', [req.params.auctionId]);
  const results = [];
  const errors = [];
  
  for (const row of buyers) {
    const useSaleType = saleType || row.default_sale || 'L';
    try {
      // Dedup — should be filtered upstream by the un-invoiced SQL
      // expression, but double-check before INSERT so a race or stale
      // lot.invo state can't produce a duplicate invoice row.
      const dupBuyer = db.get(
        'SELECT id, invo FROM invoices WHERE auction_id = ? AND sale = ? AND buyer = ? LIMIT 1',
        [req.params.auctionId, useSaleType, row.buyer]
      );
      if (dupBuyer) {
        errors.push({ buyer: row.buyer, error: `Already invoiced as #${dupBuyer.invo}` });
        continue;
      }
      const invoice = buildSalesInvoice(db, req.params.auctionId, row.buyer, useSaleType, cfg, { noTI });
      if (!invoice) { errors.push({ buyer: row.buyer, error: 'No matching lots' }); continue; }
      const s = invoice.summary;
      const invoNo = String(nextNo);
      // Store BUSINESS context state — see single-invoice handler for rationale
      const invoiceState = cfg.business_state || auction.state || '';
      // Skip if the chosen invoice number is already taken in this trade.
      const dupNo = db.get(
        'SELECT id, buyer FROM invoices WHERE auction_id = ? AND sale = ? AND invo = ? LIMIT 1',
        [req.params.auctionId, useSaleType, invoNo]
      );
      if (dupNo) {
        errors.push({ buyer: row.buyer, error: `Invoice #${invoNo} already used by ${dupNo.buyer}` });
        // Don't increment nextNo for a skipped row so the next iteration
        // tries the same number; the SELECT will reflect any insert from
        // an earlier iteration.
        continue;
      }
      db.run(`INSERT INTO invoices (auction_id,ano,date,state,sale,invo,buyer,buyer1,gstin,place,bag,qty,amount,gunny,pava_hc,ins,cgst,sgst,igst,tcs,rund,tot,addl_chg,addl_name,no_ti)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`,
        [req.params.auctionId,auction.ano,auction.date,invoiceState,useSaleType,invoNo,row.buyer,invoice.buyer.buyer1||'',
         invoice.buyer.gstin||'',invoice.buyer.pla||'',s.totalBags,s.totalQty,s.totalAmount,s.gunnyCost,s.transportCost,s.insuranceCost,
         s.cgst,s.sgst,s.igst,0,s.roundDiff,s.grandTotal,s.addlCharge||0,s.addlChargeName||'',noTI]);
      // ASP-aware lot update: see single-invoice handler above for rationale.
      const isASPStateBulk = String(cfg.business_state || '').toUpperCase() === 'KERALA';
      for (const li of invoice.lineItems) {
        if (isASPStateBulk) {
          const existing = db.get(
            'SELECT invo, asp_invo FROM lots WHERE auction_id=? AND lot_no=? AND buyer=? LIMIT 1',
            [req.params.auctionId, li.lot, row.buyer]
          );
          const hasIspInvo = existing && existing.invo && existing.invo !== existing.asp_invo;
          if (hasIspInvo) {
            db.run('UPDATE lots SET asp_invo=? WHERE auction_id=? AND lot_no=? AND buyer=? AND locked_at IS NULL',
              [invoNo, req.params.auctionId, li.lot, row.buyer]);
          } else {
            // Don't set `sale` in ASP context — ISP step decides
            db.run('UPDATE lots SET invo=?, asp_invo=? WHERE auction_id=? AND lot_no=? AND buyer=? AND locked_at IS NULL',
              [invoNo, invoNo, req.params.auctionId, li.lot, row.buyer]);
          }
        } else {
          db.run('UPDATE lots SET sale=?, invo=? WHERE auction_id=? AND lot_no=? AND buyer=? AND locked_at IS NULL',
            [useSaleType, invoNo, req.params.auctionId, li.lot, row.buyer]);
        }
      }
      results.push({ buyer: row.buyer, invoiceNo: invoNo, sale: useSaleType, grandTotal: s.grandTotal });
      nextNo++;
    } catch (e) { errors.push({ buyer: row.buyer, error: e.message }); }
  }
  
  res.json({ success: true, generated: results.length, results, errors });
});

// ── LORRY / VEHICLE NUMBER — bulk-set on selected invoices ──
// Stored on invoices.lorry_no (added in db.js). The Tally sales voucher
// generator reads this column and emits it as the e-way bill VehicleNo.
// UI ticks invoices in the Sales tab → enters one lorry no → applies
// to all ticked rows in a single round-trip.
//
// Body: { ids: [1,2,3], lorry_no: 'TN66H1234' }  (lorry_no '' = clear)
//
// CRITICAL: this MUST be declared before `app.put('/api/invoices/:id')`
// below — Express matches routes in declaration order, so a generic
// `:id` route declared first would match `lorry-no` as the id and route
// the request through the wrong handler.
app.put('/api/invoices/lorry-no', requireInvoiceWrite, (req, res) => {
  const { ids, lorry_no } = req.body || {};
  if (!Array.isArray(ids) || !ids.length) {
    return res.status(400).json({ error: 'ids array required' });
  }
  const cleanIds = ids.map(Number).filter(Number.isFinite);
  if (!cleanIds.length) return res.status(400).json({ error: 'No valid invoice IDs' });
  // Normalise: trim, uppercase, strip spaces. Empty → NULL (clear).
  // We intentionally DON'T validate format — Indian vehicle plates have
  // many regional variants (TN-66-H-1234, KL07AB1234, BH-prefixed,
  // commercial XX0000XX, etc.) and rejecting valid plates is worse than
  // accepting a typo.
  let v = null;
  if (lorry_no != null && String(lorry_no).trim() !== '') {
    v = String(lorry_no).trim().toUpperCase().replace(/\s+/g, '');
    if (v.length > 20) return res.status(400).json({ error: 'Lorry no too long (max 20 chars)' });
  }
  try {
    const db = getDb();
    // Cascade lock — drop any invoice whose buyer has a locked lot in
    // the same auction (non-admin). Editing lorry_no doesn't touch the
    // lot row directly, but the user-stated rule is that locking a lot
    // makes its dependent invoice "uneditable in other transactions",
    // so we apply it consistently here too.
    let mutableIds = cleanIds;
    let skippedLocked = 0;
    if (!isAdmin(req)) {
      mutableIds = [];
      for (const id of cleanIds) {
        if (lotsLockedForInvoice(db, id)) { skippedLocked++; continue; }
        mutableIds.push(id);
      }
    }
    if (!mutableIds.length) {
      return res.json({ ok: true, updated: 0, lorry_no: v, skipped_locked: skippedLocked });
    }
    const placeholders = mutableIds.map(() => '?').join(',');
    const r = db.run(
      `UPDATE invoices SET lorry_no = ? WHERE id IN (${placeholders})`,
      [v, ...mutableIds]
    );
    res.json({ ok: true, updated: r.changes, lorry_no: v, skipped_locked: skippedLocked });
  } catch (e) {
    console.error('[lorry-no] bulk update failed:', e);
    res.status(500).json({ error: e.message });
  }
});

// Update invoice fields (edit)
app.put('/api/invoices/:id', requireInvoiceWrite, (req, res) => {
  const i = req.body;
  const db = getDb();
  // Cascade lock — if any lot in (auction_id, buyer) for this invoice
  // is locked, the invoice is considered finalised and only an admin
  // can edit it. Mirrors the same gate on PUT /api/lots/:id.
  if (!isAdmin(req) && lotsLockedForInvoice(db, req.params.id)) {
    return res.status(423).json({ error: 'This invoice is locked because at least one of its lots is locked — only an admin can edit it.' });
  }
  const fields = ['ano','date','state','sale','invo','buyer','buyer1','gstin','place',
    'bag','qty','amount','gunny','pava_hc','ins','cgst','sgst','igst','tcs','rund','tot'];
  const sets = []; const vals = [];
  for (const f of fields) {
    if (i[f] !== undefined) { sets.push(`${f}=?`); vals.push(i[f]); }
  }
  if (!sets.length) return res.status(400).json({ error: 'No fields to update' });
  vals.push(req.params.id);
  db.run(`UPDATE invoices SET ${sets.join(',')} WHERE id=?`, vals);
  res.json({ success: true });
});

// Delete invoice
app.delete('/api/invoices/:id', requireDelete, (req, res) => {
  const db = getDb();
  const inv = db.get('SELECT * FROM invoices WHERE id=?', [req.params.id]);
  if (!inv) return res.status(404).json({ error: 'Invoice not found' });
  // Cascade lock — even admins reach this path via requireDelete (which
  // is already admin-only), so this is a no-op in practice today; it
  // stays in for defence in depth if `delete` is ever opened up to
  // managers, and to keep the cascade rule consistent across endpoints.
  if (!isAdmin(req) && lotsLockedForInvoice(db, req.params.id)) {
    return res.status(423).json({ error: 'This invoice is locked because at least one of its lots is locked — only an admin can delete it.' });
  }
  // Clear sale/invo from the related lots so they're eligible again
  let lotsFreed = 0;
  if (inv.auction_id) {
    const before = db.get('SELECT COUNT(*) as c FROM lots WHERE auction_id=? AND sale=? AND invo=? AND buyer=?',
      [inv.auction_id, inv.sale, inv.invo, inv.buyer]).c;
    db.run(`UPDATE lots SET sale='', invo='' WHERE auction_id=? AND sale=? AND invo=? AND buyer=?`,
      [inv.auction_id, inv.sale, inv.invo, inv.buyer]);
    lotsFreed = before;
  }
  db.run('DELETE FROM invoices WHERE id=?', [req.params.id]);
  res.json({ success: true, invoiceId: Number(req.params.id), lotsFreed });
});

// Explicit revert route (same effect as DELETE but returns richer info)
app.post('/api/invoices/:id/revert', requireInvoiceRevert, (req, res) => {
  const db = getDb();
  const inv = db.get('SELECT * FROM invoices WHERE id=?', [req.params.id]);
  if (!inv) return res.status(404).json({ error: 'Invoice not found' });
  // Cascade lock — revert clears lots.sale/invo, which is a mutation
  // on the lot row itself; locked lots reject that for non-admins.
  if (!isAdmin(req) && lotsLockedForInvoice(db, req.params.id)) {
    return res.status(423).json({ error: 'This invoice is locked because at least one of its lots is locked — only an admin can revert it.' });
  }
  let lotsFreed = 0;
  if (inv.auction_id) {
    const affected = db.all(
      'SELECT lot_no FROM lots WHERE auction_id=? AND sale=? AND invo=? AND buyer=?',
      [inv.auction_id, inv.sale, inv.invo, inv.buyer]
    );
    lotsFreed = affected.length;
    db.run(`UPDATE lots SET sale='', invo='' WHERE auction_id=? AND sale=? AND invo=? AND buyer=?`,
      [inv.auction_id, inv.sale, inv.invo, inv.buyer]);
    db.run('DELETE FROM invoices WHERE id=?', [req.params.id]);
    return res.json({
      success: true,
      invoice: { sale: inv.sale, invo: inv.invo, buyer: inv.buyer, buyer1: inv.buyer1 },
      lotsFreed,
      lots: affected.map(r => r.lot_no),
    });
  }
  db.run('DELETE FROM invoices WHERE id=?', [req.params.id]);
  res.json({ success: true, lotsFreed: 0 });
});

// Bulk revert: revert ALL invoices in an auction
app.post('/api/invoices/revert-all/:auctionId', requireInvoiceRevert, (req, res) => {
  const db = getDb();
  const aid = req.params.auctionId;
  const invoices = db.all('SELECT * FROM invoices WHERE auction_id = ?', [aid]);
  const admin = isAdmin(req);
  let lotsFreed = 0;
  let skippedLocked = 0;
  const revertedIds = [];
  for (const inv of invoices) {
    // Cascade lock: non-admins skip any invoice whose buyer has a
    // locked lot in this auction. Bulk-action design = "skip locked,
    // process the rest" so a single locked record doesn't block the
    // whole batch.
    if (!admin && lotsLockedForInvoice(db, inv.id)) {
      skippedLocked++;
      continue;
    }
    const n = db.get('SELECT COUNT(*) as c FROM lots WHERE auction_id=? AND sale=? AND invo=? AND buyer=? AND locked_at IS NULL',
      [inv.auction_id, inv.sale, inv.invo, inv.buyer]).c;
    lotsFreed += n;
    db.run(`UPDATE lots SET sale='', invo='' WHERE auction_id=? AND sale=? AND invo=? AND buyer=? AND locked_at IS NULL`,
      [inv.auction_id, inv.sale, inv.invo, inv.buyer]);
    db.run('DELETE FROM invoices WHERE id=?', [inv.id]);
    revertedIds.push(inv.id);
  }
  // Safety net: clear any orphan invo values from unlocked lots in this
  // auction. Locked lots are intentionally preserved.
  const orphan = db.get(
    `SELECT COUNT(*) as c FROM lots WHERE auction_id = ? AND invo IS NOT NULL AND invo != '' AND locked_at IS NULL`, [aid]
  ).c;
  if (orphan) {
    db.run(`UPDATE lots SET sale='', invo='' WHERE auction_id = ? AND locked_at IS NULL`, [aid]);
    lotsFreed += orphan;
  }
  res.json({ success: true, invoicesReverted: revertedIds.length, lotsFreed, skipped_locked: skippedLocked });
});

// Toggle "No Transport & Insurance" on an existing invoice. Recomputes
// transport/insurance/GST/round/total under the new flag and persists the
// financial columns + no_ti so the list, PDF and Tally voucher all stay
// consistent. Body: { value: 0|1 }.
app.post('/api/invoices/:id/no-ti', requireInvoiceWrite, (req, res) => {
  const db = getDb(); const cfg = getSettingsFlat(db);
  const inv = db.get('SELECT * FROM invoices WHERE id=?', [req.params.id]);
  if (!inv) return res.status(404).json({ error: 'Invoice not found' });
  // Cascade lock — same rule as edit: a locked lot finalises the invoice.
  if (!isAdmin(req) && lotsLockedForInvoice(db, req.params.id)) {
    return res.status(423).json({ error: 'This invoice is locked because at least one of its lots is locked — only an admin can change it.' });
  }
  const value = (req.body.value === true || String(req.body.value || '').toLowerCase() === 'true' || Number(req.body.value) === 1) ? 1 : 0;
  // Rebuild from the invoice's lots so transport/insurance/GST/total are
  // recomputed under the new flag. Requires the underlying lots to still
  // carry this (auction, buyer, sale) — they do unless the invoice was
  // reverted, in which case there's nothing to recompute.
  const rebuilt = inv.auction_id
    ? buildSalesInvoice(db, inv.auction_id, inv.buyer, inv.sale, cfg, { noTI: value })
    : null;
  if (!rebuilt) {
    return res.status(409).json({ error: 'Cannot recompute this invoice — its lots are no longer available. Revert and regenerate instead.' });
  }
  const s = rebuilt.summary;
  db.run(
    `UPDATE invoices SET no_ti=?, pava_hc=?, ins=?, cgst=?, sgst=?, igst=?, tcs=?, rund=?, tot=? WHERE id=?`,
    [value, s.transportCost, s.insuranceCost, s.cgst, s.sgst, s.igst, s.tdsAmount || 0, s.roundDiff, s.grandTotal, req.params.id]
  );
  res.json({ success: true, value, summary: s });
});

// Sales Invoice PDF
app.get('/api/invoices/pdf/:id', requireView, async (req, res) => {
  try {
    const db = getDb();
    const cfg = getSettingsFlat(db);
    const stored = db.get('SELECT * FROM invoices WHERE id=?', [req.params.id]);
    if (!stored) return res.status(404).json({ error: 'Invoice not found' });

    // Try to rebuild fresh from lots (gives line-item detail), fall back to stored summary
    let invoice = stored.auction_id
      ? buildSalesInvoice(db, stored.auction_id, stored.buyer, stored.sale, cfg, { noTI: stored.no_ti })
      : null;

    // Defensive: even when lots exist, if buyer lookup missed, enrich from stored invoice fields
    // so BILL TO / SHIPPED TO isn't blank.
    const enrichBuyer = (buyer) => {
      if (!buyer) buyer = {};
      // If buyer has no recognizable display field, try several fallbacks
      if (!buyer.buyer1 && !buyer.buyer) {
        const looked = db.get('SELECT * FROM buyers WHERE buyer=? OR buyer1=? LIMIT 1',
          [stored.buyer, stored.buyer1 || stored.tradername || '']);
        if (looked) buyer = looked;
      }
      // Last-resort fill from stored invoice row
      if (!buyer.buyer1 && stored.buyer1)   buyer.buyer1   = stored.buyer1;
      if (!buyer.buyer1 && stored.tradername) buyer.buyer1 = stored.tradername;
      if (!buyer.buyer  && stored.buyer)    buyer.buyer    = stored.buyer;
      if (!buyer.gstin  && stored.gstin)    buyer.gstin    = stored.gstin;
      if (!buyer.pla    && stored.place)    buyer.pla      = stored.place;
      if (!buyer.state  && stored.state)    buyer.state    = stored.state;
      if (!buyer.add1   && stored.add_line) buyer.add1     = stored.add_line;
      return buyer;
    };

    if (invoice) {
      invoice.buyer = enrichBuyer(invoice.buyer);
    } else {
      // Build a minimal invoice object from stored fields (lots may have been deleted)
      const buyer = enrichBuyer(db.get('SELECT * FROM buyers WHERE buyer=? LIMIT 1', [stored.buyer]));
      invoice = {
        buyer,
        lineItems: [{ lot: '—', grade: '', bags: stored.bag || 0, qty: stored.qty || 0, price: 0, amount: stored.amount || 0 }],
        summary: {
          totalBags: stored.bag || 0,
          totalQty: stored.qty || 0,
          totalAmount: stored.amount || 0,
          gunnyCost: stored.gunny || 0,
          transportCost: stored.pava_hc || 0,
          insuranceCost: stored.ins || 0,
          cgst: stored.cgst || 0,
          sgst: stored.sgst || 0,
          igst: stored.igst || 0,
          tcs: stored.tcs || 0,
          roundDiff: stored.rund || 0,
          subtotalRounded: (stored.tot || 0) - (stored.addl_chg || 0),
          addlCharge: stored.addl_chg || 0,
          addlChargeName: stored.addl_name || '',
          grandTotal: stored.tot || 0,
          isInterState: stored.sale === 'I',
        }
      };
    }

    // Optional dispatched-through override from print modal (URL-encoded)
    const dispatchedThrough = req.query.dispatchedThrough || '';
    if (dispatchedThrough) invoice.dispatchedThrough = dispatchedThrough;

    // Look up the ASP invoice number from lots so the ISP PDF can show
    // the cross-reference under "Other References" as ASP/I-{asp}/{season}.
    // When this invoice IS an ASP one (state=KERALA), aspInvo stays empty.
    if (String(stored.state || '').toUpperCase() !== 'KERALA') {
      const aspRow = db.get(
        `SELECT asp_invo FROM lots
         WHERE auction_id = ? AND buyer = ? AND invo = ?
           AND asp_invo IS NOT NULL AND asp_invo != ''
         LIMIT 1`,
        [stored.auction_id, stored.buyer, stored.invo]
      );
      if (aspRow && aspRow.asp_invo) invoice.aspInvo = aspRow.asp_invo;
    }

    const pdf = await generateSalesInvoicePDF(invoice, cfg, stored.sale, stored.invo, stored.date);
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `inline; filename="Invoice_${stored.sale}_${stored.invo}.pdf"`);
    res.send(pdf);
  } catch (e) {
    console.error('Sales invoice PDF error:', e);
    res.status(500).json({ error: 'PDF generation failed: ' + e.message });
  }
});

// ASP Purchase-view PDF — renders the mirror-image of an ASP sales invoice
// with ISPL at the top (issuer), ASP as Seller (Bill from), TN bank details.
// Only valid when current business_mode is e-Trade and business_state is KERALA;
// otherwise returns 400. Uses the same `generateSalesInvoicePDF` code path with
// variant='purchase' so the math (P_Rate, PurAmt, totals, HSN) stays identical
// to the source sales invoice.
app.get('/api/invoices/purchase-pdf/:id', requireView, async (req, res) => {
  try {
    const db = getDb();
    const cfg = getSettingsFlat(db);
    // Guard: purchase view is meaningful only for ASP invoices
    const isASPContext = (String(cfg.business_mode || '').toLowerCase() === 'e-trade')
                      && (String(cfg.business_state || '').toUpperCase() === 'KERALA');
    if (!isASPContext) {
      return res.status(400).json({
        error: 'Purchase view is only available when business state is Kerala (e-Trade mode).'
      });
    }
    const stored = db.get('SELECT * FROM invoices WHERE id=?', [req.params.id]);
    if (!stored) return res.status(404).json({ error: 'Invoice not found' });

    // Same enrichment pattern as the sales-invoice endpoint
    let invoice = stored.auction_id
      ? buildSalesInvoice(db, stored.auction_id, stored.buyer, stored.sale, cfg, { noTI: stored.no_ti })
      : null;

    const enrichBuyer = (buyer) => {
      if (!buyer) buyer = {};
      if (!buyer.buyer1 && !buyer.buyer) {
        const looked = db.get('SELECT * FROM buyers WHERE buyer=? OR buyer1=? LIMIT 1',
          [stored.buyer, stored.buyer1 || stored.tradername || '']);
        if (looked) buyer = looked;
      }
      if (!buyer.buyer1 && stored.buyer1)   buyer.buyer1   = stored.buyer1;
      if (!buyer.buyer1 && stored.tradername) buyer.buyer1 = stored.tradername;
      if (!buyer.buyer  && stored.buyer)    buyer.buyer    = stored.buyer;
      if (!buyer.gstin  && stored.gstin)    buyer.gstin    = stored.gstin;
      if (!buyer.pla    && stored.place)    buyer.pla      = stored.place;
      if (!buyer.state  && stored.state)    buyer.state    = stored.state;
      if (!buyer.add1   && stored.add_line) buyer.add1     = stored.add_line;
      return buyer;
    };

    if (invoice) {
      invoice.buyer = enrichBuyer(invoice.buyer);
    } else {
      const buyer = enrichBuyer(db.get('SELECT * FROM buyers WHERE buyer=? LIMIT 1', [stored.buyer]));
      invoice = {
        buyer,
        lineItems: [{ lot: '—', grade: '', bags: stored.bag || 0, qty: stored.qty || 0, price: 0, amount: stored.amount || 0 }],
        summary: {
          totalBags: stored.bag || 0,
          totalQty: stored.qty || 0,
          totalAmount: stored.amount || 0,
          gunnyCost: stored.gunny || 0,
          transportCost: stored.pava_hc || 0,
          insuranceCost: stored.ins || 0,
          cgst: stored.cgst || 0,
          sgst: stored.sgst || 0,
          igst: stored.igst || 0,
          tcs: stored.tcs || 0,
          roundDiff: stored.rund || 0,
          subtotalRounded: (stored.tot || 0) - (stored.addl_chg || 0),
          addlCharge: stored.addl_chg || 0,
          addlChargeName: stored.addl_name || '',
          grandTotal: stored.tot || 0,
          isInterState: stored.sale === 'I',
        }
      };
    }

    // variant='purchase' flips the display: ISPL at top, ASP as seller, TN bank
    const pdf = await generateSalesInvoicePDF(invoice, cfg, stored.sale, stored.invo, stored.date, undefined, 'purchase');
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `inline; filename="PurchaseView_${stored.sale}_${stored.invo}.pdf"`);
    res.send(pdf);
  } catch (e) {
    console.error('Purchase-view PDF error:', e);
    res.status(500).json({ error: 'PDF generation failed: ' + e.message });
  }
});

// Bulk Sales Invoice PDF — merges N invoices into a single PDF
// Body: { ids: [1, 2, 3, ...] }
// Returns: one PDF with each invoice on fresh page(s), in the order given.
app.post('/api/invoices/pdf-bulk', requireView, async (req, res) => {
  try {
    const ids = Array.isArray(req.body?.ids) ? req.body.ids.map(Number).filter(Boolean) : [];
    if (!ids.length) return res.status(400).json({ error: 'No invoice IDs provided' });
    // Optional dispatched-through override applied to every invoice in the batch.
    const dispatchedThrough = (req.body?.dispatchedThrough || '').toString();

    const db = getDb();
    const cfg = getSettingsFlat(db);

    // Enrich-buyer helper (same logic as single-invoice endpoint) — ensures
    // Bill-To / Ship-To have values even when rebuilding from stored data.
    const enrichBuyer = (buyer, stored) => {
      if (!buyer) buyer = {};
      if (!buyer.buyer1 && !buyer.buyer) {
        const looked = db.get('SELECT * FROM buyers WHERE buyer=? OR buyer1=? LIMIT 1',
          [stored.buyer, stored.buyer1 || stored.tradername || '']);
        if (looked) buyer = looked;
      }
      if (!buyer.buyer1 && stored.buyer1)   buyer.buyer1   = stored.buyer1;
      if (!buyer.buyer1 && stored.tradername) buyer.buyer1 = stored.tradername;
      if (!buyer.buyer  && stored.buyer)    buyer.buyer    = stored.buyer;
      if (!buyer.gstin  && stored.gstin)    buyer.gstin    = stored.gstin;
      if (!buyer.pla    && stored.place)    buyer.pla      = stored.place;
      if (!buyer.state  && stored.state)    buyer.state    = stored.state;
      if (!buyer.add1   && stored.add_line) buyer.add1     = stored.add_line;
      return buyer;
    };

    // Build each invoice's data
    const payloads = [];
    for (const id of ids) {
      const stored = db.get('SELECT * FROM invoices WHERE id=?', [id]);
      if (!stored) continue; // silently skip missing IDs
      let invoice = stored.auction_id
        ? buildSalesInvoice(db, stored.auction_id, stored.buyer, stored.sale, cfg, { noTI: stored.no_ti })
        : null;
      if (invoice) {
        invoice.buyer = enrichBuyer(invoice.buyer, stored);
      } else {
        const buyer = enrichBuyer(db.get('SELECT * FROM buyers WHERE buyer=? LIMIT 1', [stored.buyer]), stored);
        invoice = {
          buyer,
          lineItems: [{ lot: '—', grade: '', bags: stored.bag || 0, qty: stored.qty || 0, price: 0, amount: stored.amount || 0 }],
          summary: {
            totalBags: stored.bag || 0, totalQty: stored.qty || 0,
            totalAmount: stored.amount || 0, gunnyCost: stored.gunny || 0,
            transportCost: stored.pava_hc || 0, insuranceCost: stored.ins || 0,
            taxableValue: (stored.amount || 0) + (stored.gunny || 0) + (stored.pava_hc || 0) + (stored.ins || 0),
            cgst: stored.cgst || 0, sgst: stored.sgst || 0, igst: stored.igst || 0,
            roundDiff: stored.rund || 0,
            subtotalRounded: (stored.tot || 0) - (stored.addl_chg || 0),
            addlCharge: stored.addl_chg || 0,
            addlChargeName: stored.addl_name || '',
            grandTotal: stored.tot || 0,
            isInterState: stored.sale === 'I',
          }
        };
      }
      if (dispatchedThrough) invoice.dispatchedThrough = dispatchedThrough;
      // ASP cross-reference for ISP invoices — see single endpoint for rationale
      if (String(stored.state || '').toUpperCase() !== 'KERALA') {
        const aspRow = db.get(
          `SELECT asp_invo FROM lots
           WHERE auction_id = ? AND buyer = ? AND invo = ?
             AND asp_invo IS NOT NULL AND asp_invo != ''
           LIMIT 1`,
          [stored.auction_id, stored.buyer, stored.invo]
        );
        if (aspRow && aspRow.asp_invo) invoice.aspInvo = aspRow.asp_invo;
      }
      payloads.push({
        invoiceData: invoice,
        saleType: stored.sale,
        invoiceNo: stored.invo,
        invoiceDate: stored.date,
      });
    }

    if (!payloads.length) return res.status(404).json({ error: 'No invoices resolved from the provided IDs' });

    const pdf = await generateSalesInvoicesBatchPDF(payloads, cfg);
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `inline; filename="Invoices_Batch_${payloads.length}.pdf"`);
    res.send(pdf);
  } catch (e) {
    console.error('Bulk sales invoice PDF error:', e);
    res.status(500).json({ error: 'Batch PDF generation failed: ' + e.message });
  }
});

// Bulk Purchase-View PDF — like /pdf-bulk but renders each invoice with
// variant='purchase' so the buyer (ISPL) appears as the issuing company
// and the active ASP company appears as the seller. ASP context only.
// Body: { ids: [1, 2, 3, ...] }
app.post('/api/invoices/purchase-pdf-bulk', requireView, async (req, res) => {
  try {
    const ids = Array.isArray(req.body?.ids) ? req.body.ids.map(Number).filter(Boolean) : [];
    if (!ids.length) return res.status(400).json({ error: 'No invoice IDs provided' });

    const db = getDb();
    const cfg = getSettingsFlat(db);

    const isASPContext = (String(cfg.business_mode || '').toLowerCase() === 'e-trade')
                      && (String(cfg.business_state || '').toUpperCase() === 'KERALA');
    if (!isASPContext) {
      return res.status(400).json({
        error: 'Purchase view is only available when business state is Kerala (e-Trade mode).'
      });
    }

    // Same enrichBuyer pattern as /pdf-bulk
    const enrichBuyer = (buyer, stored) => {
      if (!buyer) buyer = {};
      if (!buyer.buyer1 && !buyer.buyer) {
        const looked = db.get('SELECT * FROM buyers WHERE buyer=? OR buyer1=? LIMIT 1',
          [stored.buyer, stored.buyer1 || stored.tradername || '']);
        if (looked) buyer = looked;
      }
      if (!buyer.buyer1 && stored.buyer1)   buyer.buyer1   = stored.buyer1;
      if (!buyer.buyer1 && stored.tradername) buyer.buyer1 = stored.tradername;
      if (!buyer.buyer  && stored.buyer)    buyer.buyer    = stored.buyer;
      if (!buyer.gstin  && stored.gstin)    buyer.gstin    = stored.gstin;
      if (!buyer.pla    && stored.place)    buyer.pla      = stored.place;
      if (!buyer.state  && stored.state)    buyer.state    = stored.state;
      if (!buyer.add1   && stored.add_line) buyer.add1     = stored.add_line;
      return buyer;
    };

    const payloads = [];
    for (const id of ids) {
      const stored = db.get('SELECT * FROM invoices WHERE id=?', [id]);
      if (!stored) continue;
      let invoice = stored.auction_id
        ? buildSalesInvoice(db, stored.auction_id, stored.buyer, stored.sale, cfg, { noTI: stored.no_ti })
        : null;
      if (invoice) {
        invoice.buyer = enrichBuyer(invoice.buyer, stored);
      } else {
        const buyer = enrichBuyer(db.get('SELECT * FROM buyers WHERE buyer=? LIMIT 1', [stored.buyer]), stored);
        invoice = {
          buyer,
          lineItems: [{ lot: '—', grade: '', bags: stored.bag || 0, qty: stored.qty || 0, price: 0, amount: stored.amount || 0 }],
          summary: {
            totalBags: stored.bag || 0, totalQty: stored.qty || 0,
            totalAmount: stored.amount || 0, gunnyCost: stored.gunny || 0,
            transportCost: stored.pava_hc || 0, insuranceCost: stored.ins || 0,
            taxableValue: (stored.amount || 0) + (stored.gunny || 0) + (stored.pava_hc || 0) + (stored.ins || 0),
            cgst: stored.cgst || 0, sgst: stored.sgst || 0, igst: stored.igst || 0,
            roundDiff: stored.rund || 0,
            subtotalRounded: (stored.tot || 0) - (stored.addl_chg || 0),
            addlCharge: stored.addl_chg || 0,
            addlChargeName: stored.addl_name || '',
            grandTotal: stored.tot || 0,
            isInterState: stored.sale === 'I',
          }
        };
      }
      payloads.push({
        invoiceData: invoice,
        saleType: stored.sale,
        invoiceNo: stored.invo,
        invoiceDate: stored.date,
      });
    }
    if (!payloads.length) return res.status(404).json({ error: 'No invoices resolved from the provided IDs' });

    // 'purchase' variant: ISPL at top, ASP as seller, TN bank — applied to every page
    const pdf = await generateSalesInvoicesBatchPDF(payloads, cfg, 'purchase');
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `inline; filename="PurchaseView_Batch_${payloads.length}.pdf"`);
    res.send(pdf);
  } catch (e) {
    console.error('Bulk purchase-view PDF error:', e);
    res.status(500).json({ error: 'Batch PDF generation failed: ' + e.message });
  }
});

// ══════════════════════════════════════════════════════════════
// PURCHASES (GSTKBILT.PRG — registered dealer invoices)
// ══════════════════════════════════════════════════════════════
app.get('/api/purchases', requireView, (req, res) => {
  const { auction_id, ano, from, to, sale, search } = req.query;
  let q = 'SELECT * FROM purchases WHERE 1=1'; const p = [];
  if (auction_id) { q += ' AND auction_id = ?'; p.push(parseInt(auction_id)); }
  if (ano) { q += ' AND ano = ?'; p.push(ano); }
  // Free-text search within the selected trade — invoice no, seller
  // name, GSTIN.
  const searchTerm = String(search || '').trim();
  if (searchTerm) {
    const wild = `%${searchTerm}%`;
    q += ` AND (
            COALESCE(invo,'')  LIKE ?
            OR COALESCE(name,'')  LIKE ?
            OR COALESCE(gstin,'') LIKE ?
          )`;
    p.push(wild, wild, wild);
  }
  if (from && to) { q += ' AND date BETWEEN ? AND ?'; p.push(from, to); }
  // Sale-type filter (L / I / E). Purchases don't carry a sale column,
  // but the GST split on each row is deterministic:
  //   • L (Local / intra-state)  → CGST + SGST > 0, IGST = 0
  //   • I (Inter-state)          → IGST > 0, CGST = SGST = 0
  //   • E (Export)               → IGST > 0 AND dealer has ≥1 lot
  //                                tagged sale='E' in the same auction
  // Inferring directly from lots.sale would fail because that column
  // is blank on most installs — L would match everything, I nothing.
  // The GST-column approach is data-driven and always correct.
  const saleNorm = String(sale || '').trim().toUpperCase();
  if (saleNorm === 'L') {
    q += ' AND COALESCE(igst,0) = 0 AND (COALESCE(cgst,0) > 0 OR COALESCE(sgst,0) > 0)';
  } else if (saleNorm === 'I') {
    q += ' AND COALESCE(igst,0) > 0';
    q += ` AND NOT EXISTS (
            SELECT 1 FROM lots l
             WHERE l.auction_id = purchases.auction_id
               AND UPPER(TRIM(COALESCE(l.name,''))) = UPPER(TRIM(COALESCE(purchases.name,'')))
               AND UPPER(TRIM(COALESCE(l.sale,''))) = 'E'
          )`;
  } else if (saleNorm === 'E') {
    q += ' AND COALESCE(igst,0) > 0';
    q += ` AND EXISTS (
            SELECT 1 FROM lots l
             WHERE l.auction_id = purchases.auction_id
               AND UPPER(TRIM(COALESCE(l.name,''))) = UPPER(TRIM(COALESCE(purchases.name,'')))
               AND UPPER(TRIM(COALESCE(l.sale,''))) = 'E'
          )`;
  }
  q += ' ORDER BY date DESC LIMIT 500';
  res.json(getDb().all(q, p));
});

// Normalize a GSTIN string for equality comparison. Strips the legacy
// "GSTIN."/"GSTIN" prefix and any whitespace, uppercases. Two values
// that both resolve to the same 15-char body compare equal even if
// one was stored in the UI-prefix format and the other was the bare
// 15-char Excel-import format.
function _normGstin(s) {
  let v = String(s == null ? '' : s).trim().toUpperCase();
  if (v.startsWith('GSTIN.')) v = v.slice(6);
  else if (v.startsWith('GSTIN')) v = v.slice(5);
  return v.trim();
}

// Resolve the company's (buyer's) GSTIN. State-aware: in a Kerala
// install the GSTIN lives under Settings → Address (Kerala) → GSTIN
// (`kl_gstin`); in a Tamil Nadu install under Address (Tamil Nadu)
// (`tn_gstin`). We probe BOTH slots regardless of state and also
// fall through to the generic top-level keys, so a user who typed
// the GSTIN under the wrong section still gets a usable value.
function _resolveBuyerGstin(cfg) {
  const stateUpper = String(cfg && cfg.business_state || '').toUpperCase();
  const isKerala = stateUpper === 'KERALA';
  const candidates = isKerala
    ? [cfg.kl_gstin, cfg.tn_gstin, cfg.gstin, cfg.business_gstin]
    : [cfg.tn_gstin, cfg.kl_gstin, cfg.gstin, cfg.business_gstin];
  for (const c of candidates) {
    const norm = _normGstin(c);
    if (norm) return norm;
  }
  // Last resort: identity helper (reads its own pick chain).
  try {
    const id = getCompanyIdentity(cfg);
    return _normGstin(id && id.gstin);
  } catch (_) { return ''; }
}

// Normalize a company / seller name for equality comparison. Trims,
// collapses internal whitespace, uppercases. Two values that mean the
// same name (different spacing / case) compare equal.
function _normName(s) {
  return String(s == null ? '' : s).trim().toUpperCase().replace(/\s+/g, ' ');
}

// Resolve the company's (buyer's) trade name. Reads `trade_name` first
// (the canonical user-edited field), falls back to short_name and the
// central identity helper. Used as a SECONDARY same-entity signal —
// when a seller record happens to share the company's name but has a
// blank GSTIN, we still want to block the purchase (it's almost
// certainly the same legal entity with a configuration gap).
function _resolveBuyerName(cfg) {
  const direct = _normName(cfg && (cfg.trade_name || cfg.short_name || cfg.company_name || cfg.tally_company_name));
  if (direct) return direct;
  try {
    const id = getCompanyIdentity(cfg);
    return _normName(id && (id.name || id.shortName));
  } catch (_) { return ''; }
}

// Same-entity check: returns the reason string (for the API response)
// when the seller and the company should be treated as the same legal
// entity, otherwise returns null. Priority:
//   1. GSTIN match (definitive — same legal entity)
//   2. Name match (fallback — covers blank-GSTIN sellers that share
//      the company's trade name)
// Both `invoice.seller` and `cfg` are required.
function _checkSameEntity(invoice, cfg) {
  if (!invoice || !invoice.seller) return null;
  const sellerGstin = _normGstin(invoice.seller.cr || invoice.seller.gstin);
  const buyerGstin  = _resolveBuyerGstin(cfg);
  if (sellerGstin && buyerGstin && sellerGstin === buyerGstin) {
    return { reason: `Seller GSTIN (${sellerGstin}) matches the buyer/company GSTIN — same legal entity, no purchase invoice can be raised.`, by: 'gstin', sellerGstin, buyerGstin };
  }
  const sellerName = _normName(invoice.seller.name);
  const buyerName  = _resolveBuyerName(cfg);
  if (sellerName && buyerName && sellerName === buyerName) {
    return { reason: `Seller name (${invoice.seller.name}) matches the company name — same legal entity, no purchase invoice can be raised.`, by: 'name', sellerName, buyerName };
  }
  return null;
}

app.post('/api/purchases/generate/:auctionId',
  requireInvoiceWrite,
  requirePriceChecked(req => parseInt(req.params.auctionId, 10)),
  (req, res) => {
  const db = getDb(); const cfg = getSettingsFlat(db);
  const auctionIdForGate = parseInt(req.params.auctionId, 10);
  const _gen = _checkGenerationGate(db, 'purchases', auctionIdForGate);
  if (!_gen.allowed) return res.status(412).json(_gen.error);
  const { sellerName, invoiceNo } = req.body;
  const invoice = buildPurchaseInvoice(db, req.params.auctionId, sellerName, cfg);
  if (!invoice) return res.status(404).json({ error: 'No data for this seller' });

  // Skip same-entity transactions: if the seller's GSTIN equals our
  // company's GSTIN, the buyer and seller are the same legal entity
  // (e.g. inter-branch transfer under one GSTIN). Issuing a tax invoice
  // in that case isn't valid — refuse with a clear error so the operator
  // knows to handle it as a stock transfer instead.
  const sellerGstin = _normGstin(invoice.seller && (invoice.seller.cr || invoice.seller.gstin));
  const buyerGstin  = _resolveBuyerGstin(cfg);
  // Soft guardrail: if the company GSTIN couldn't be resolved at all
  // we surface a clear configuration error instead of silently allowing
  // the generation. Settings → Address (Kerala/Tamil Nadu) → GSTIN is
  // the canonical place to set it; without that filled in, the gate
  // can't compare anything.
  if (sellerGstin && !buyerGstin) {
    console.warn('[purchases/generate] Buyer GSTIN could not be resolved from cfg; same-entity check skipped. Set Settings → Address → GSTIN.');
  }
  if (sellerGstin && buyerGstin && sellerGstin === buyerGstin) {
    return res.status(400).json({
      error: `Seller GSTIN (${sellerGstin}) matches the buyer/company GSTIN — same legal entity, no purchase invoice can be raised. Treat as an internal stock transfer.`,
      sellerGstin, buyerGstin,
    });
  }

  const auction = db.get('SELECT * FROM auctions WHERE id = ?', [req.params.auctionId]);
  // Dedup — refuse a second purchase for the same dealer in this trade,
  // and refuse to reuse an invoice number that's already taken.
  const dupSeller = db.get(
    'SELECT id, invo FROM purchases WHERE auction_id = ? AND name = ? LIMIT 1',
    [req.params.auctionId, invoice.seller.name]
  );
  if (dupSeller) {
    return res.status(409).json({
      error: `Purchase already exists for dealer "${invoice.seller.name}" in this trade — invoice #${dupSeller.invo}.`,
      existingId: dupSeller.id, existingInvo: dupSeller.invo,
    });
  }
  const dupNo = db.get(
    'SELECT id, name FROM purchases WHERE auction_id = ? AND invo = ? LIMIT 1',
    [req.params.auctionId, String(invoiceNo)]
  );
  if (dupNo) {
    return res.status(409).json({
      error: `Purchase invoice number ${invoiceNo} is already used in this trade by dealer "${dupNo.name}".`,
      existingId: dupNo.id, existingDealer: dupNo.name,
    });
  }
  const s = invoice.summary;
  db.run(`INSERT INTO purchases (auction_id,ano,date,state,br,name,add_line,place,gstin,invo,qty,amount,cgst,sgst,igst,rund,total,tds)
    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`,
    [req.params.auctionId,auction.ano,auction.date,auction.state||'','',invoice.seller.name,invoice.seller.address||'',
     invoice.seller.place||'',invoice.seller.cr||'',String(invoiceNo),s.totalQty,s.totalPuramt,
     s.totalCgst,s.totalSgst,s.totalIgst,s.roundDiff,s.grandTotal,s.tdsAmount]);
  res.json({ success: true, invoice: s });
});

// List eligible sellers for purchase invoices (with GSTIN, amount > 0)
app.get('/api/purchases/eligible-sellers/:auctionId', requireView, (req, res) => {
  // A lot is eligible for a Purchase Invoice when it has a GSTIN-bearing
  // seller. The `cr` column stores GSTINs in two known formats:
  //   1. "GSTIN.<15-char>" — legacy UI format
  //   2. "<15-char>"        — bare format from Excel imports / API
  // Earlier this endpoint only matched format #1 (UPPER(cr) LIKE 'GSTIN%')
  // and silently excluded every seller whose `cr` was the bare form,
  // producing the reported "No eligible dealer(s) found" empty state on
  // installs where dealers were imported via XLSX rather than typed in
  // through the UI.
  //
  // Fix: accept either form. The bare-GSTIN check uses GLOB '[0-9][0-9]*'
  // (cr starts with two digits — matches the state-code prefix of every
  // valid GSTIN) plus a length >= 15 guard. Both forms fall through the
  // same downstream `gstinStateCode` helper so intra/inter logic is
  // unaffected.
  const cfgEs = getSettingsFlat(getDb());
  const buyerGstinEs = _resolveBuyerGstin(cfgEs);
  const allSellers = getDb().all(
    `SELECT name, COUNT(*) as lot_count, SUM(qty) as total_qty, SUM(amount) as total_amount, MAX(cr) as cr
     FROM lots
     WHERE auction_id = ? AND name IS NOT NULL AND name != ''
       AND amount > 0
       AND (
         UPPER(cr) LIKE 'GSTIN%'
         OR (cr GLOB '[0-9][0-9]*' AND LENGTH(cr) >= 15)
       )
     GROUP BY name
     ORDER BY name`,
    [req.params.auctionId]
  );
  // Drop sellers whose GSTIN equals the buyer/company GSTIN — those
  // sales are internal stock transfers, not purchases. Keeps them out
  // of the dropdown AND the bulk "Generate for All" set.
  const eligible = buyerGstinEs
    ? allSellers.filter(r => _normGstin(r.cr) !== buyerGstinEs)
    : allSellers;
  res.json(eligible);
});

// Batch: generate purchase invoice for ALL registered dealers in an auction
app.post('/api/purchases/generate-all/:auctionId',
  requireInvoiceWrite,
  requirePriceChecked(req => parseInt(req.params.auctionId, 10)),
  (req, res) => {
  const db = getDb(); const cfg = getSettingsFlat(db);
  const auctionIdForGate = parseInt(req.params.auctionId, 10);
  const _gen = _checkGenerationGate(db, 'purchases', auctionIdForGate);
  if (!_gen.allowed) return res.status(412).json(_gen.error);
  const { startInvoiceNo } = req.body;
  
  let nextNo = parseInt(startInvoiceNo);
  if (!nextNo || nextNo < 1) return res.status(400).json({ error: 'startInvoiceNo must be a positive integer' });
  
  // Auto-calculate
  const uncalc = db.all(`SELECT * FROM lots WHERE auction_id = ? AND amount > 0 AND (puramt IS NULL OR puramt = 0)`, [req.params.auctionId]);
  for (const lot of uncalc) {
    const c = calculateLot(lot, cfg);
    db.run(`UPDATE lots SET pqty=?,prate=?,puramt=?,com=?,sertax=?,cgst=?,sgst=?,igst=?,advance=?,balance=?,bilamt=?,refund=?,refud=?,isp_pqty=?,isp_prate=?,isp_puramt=?,asp_pqty=?,asp_prate=?,asp_puramt=? WHERE id=?`,
      [c.pqty,c.prate,c.puramt,c.com,c.sertax,c.cgst,c.sgst,c.igst,c.advance,c.balance,c.bilamt,c.refund||0,c.refud||0,c.isp_pqty||0,c.isp_prate||0,c.isp_puramt||0,c.asp_pqty||0,c.asp_prate||0,c.asp_puramt||0,lot.id]);
  }
  
  // Same dual-format GSTIN filter as the eligible-sellers endpoint —
  // ensures bare-GSTIN imports aren't silently skipped during batch
  // generation.
  const sellers = db.all(
    `SELECT DISTINCT name FROM lots
     WHERE auction_id = ?
       AND amount > 0
       AND name IS NOT NULL AND name != ''
       AND (
         UPPER(cr) LIKE 'GSTIN%'
         OR (cr GLOB '[0-9][0-9]*' AND LENGTH(cr) >= 15)
       )`,
    [req.params.auctionId]
  );
  
  if (!sellers.length) return res.status(404).json({ error: 'No registered dealers (with GSTIN) in this auction' });
  
  const auction = db.get('SELECT * FROM auctions WHERE id = ?', [req.params.auctionId]);
  const results = [];
  const errors = [];
  const buyerGstinAll = _resolveBuyerGstin(cfg);
  if (!buyerGstinAll) {
    console.warn('[purchases/generate-all] Buyer GSTIN could not be resolved from cfg; same-entity skip will be inactive. Set Settings → Address → GSTIN.');
  }

  for (const row of sellers) {
    try {
      // Skip dealers who already have a purchase invoice in this trade.
      const dupSeller = db.get(
        'SELECT id, invo FROM purchases WHERE auction_id = ? AND name = ? LIMIT 1',
        [req.params.auctionId, row.name]
      );
      if (dupSeller) {
        errors.push({ seller: row.name, error: `Already invoiced as #${dupSeller.invo}` });
        continue;
      }
      const invoice = buildPurchaseInvoice(db, req.params.auctionId, row.name, cfg);
      if (!invoice) { errors.push({ seller: row.name, error: 'Build failed' }); continue; }
      // Same-entity skip: seller GSTIN equals buyer/company GSTIN.
      // See the single-generate handler for rationale.
      const sellerGstinAll = _normGstin(invoice.seller && (invoice.seller.cr || invoice.seller.gstin));
      if (sellerGstinAll && buyerGstinAll && sellerGstinAll === buyerGstinAll) {
        errors.push({ seller: row.name, error: `Same-entity (GSTIN ${sellerGstinAll}) — no purchase invoice raised` });
        continue;
      }
      const s = invoice.summary;
      const invoNo = String(nextNo);
      // Skip if this number is already taken by another dealer in the trade.
      const dupNo = db.get(
        'SELECT id, name FROM purchases WHERE auction_id = ? AND invo = ? LIMIT 1',
        [req.params.auctionId, invoNo]
      );
      if (dupNo) {
        errors.push({ seller: row.name, error: `Invoice #${invoNo} already used by ${dupNo.name}` });
        continue;
      }
      db.run(`INSERT INTO purchases (auction_id,ano,date,state,br,name,add_line,place,gstin,invo,qty,amount,cgst,sgst,igst,rund,total,tds)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`,
        [req.params.auctionId,auction.ano,auction.date,auction.state||'','',invoice.seller.name,invoice.seller.address||'',
         invoice.seller.place||'',invoice.seller.cr||'',invoNo,s.totalQty,s.totalPuramt,
         s.totalCgst,s.totalSgst,s.totalIgst,s.roundDiff,s.grandTotal,s.tdsAmount]);
      results.push({ seller: row.name, invoiceNo: invoNo, grandTotal: s.grandTotal });
      nextNo++;
    } catch (e) { errors.push({ seller: row.name, error: e.message }); }
  }
  
  res.json({ success: true, generated: results.length, results, errors });
});

// Update purchase (edit)
app.put('/api/purchases/:id', requireInvoiceWrite, (req, res) => {
  const p = req.body;
  const db = getDb();
  // Cascade lock — if any lot for this purchase's seller in this trade
  // is locked, the purchase is treated as finalised; only an admin can
  // edit it. Same pattern as PUT /api/invoices/:id.
  if (!isAdmin(req) && lotsLockedForPurchase(db, req.params.id)) {
    return res.status(423).json({ error: 'This purchase is locked because at least one of its lots is locked — only an admin can edit it.' });
  }
  const fields = ['ano','date','state','br','name','add_line','place','gstin','invo',
    'qty','amount','cgst','sgst','igst','rund','total','tds'];
  const sets = []; const vals = [];
  for (const f of fields) {
    if (p[f] !== undefined) { sets.push(`${f}=?`); vals.push(p[f]); }
  }
  if (!sets.length) return res.status(400).json({ error: 'No fields to update' });
  vals.push(req.params.id);
  db.run(`UPDATE purchases SET ${sets.join(',')} WHERE id=?`, vals);
  res.json({ success: true });
});

// Delete purchase
app.delete('/api/purchases/:id', requireDelete, (req, res) => {
  const db = getDb();
  if (!isAdmin(req) && lotsLockedForPurchase(db, req.params.id)) {
    return res.status(423).json({ error: 'This purchase is locked because at least one of its lots is locked — only an admin can delete it.' });
  }
  db.run('DELETE FROM purchases WHERE id=?', [req.params.id]);
  res.json({ success: true });
});

// ── Purchase Invoice PDF ─────────────────────────────────────
app.get('/api/purchases/pdf/:auctionId/:sellerName', requireView, async (req, res) => {
  try {
    const db = getDb();
    const cfg = getSettingsFlat(db);
    const sellerName = decodeURIComponent(req.params.sellerName);
    const auctionId = req.params.auctionId;
    const invoiceNo = req.query.invoiceNo || '001';
    
    // Try to build fresh invoice from lots
    let invoice = buildPurchaseInvoice(db, auctionId, sellerName, cfg);
    
    // Fallback: if lots data missing, rebuild from stored purchase record
    if (!invoice) {
      // Try with auction_id first, then fall back to name+invo match (for older records)
      let stored = db.get(
        `SELECT * FROM purchases WHERE auction_id = ? AND name = ? AND invo = ? LIMIT 1`,
        [auctionId, sellerName, String(invoiceNo)]
      );
      if (!stored) {
        stored = db.get(
          `SELECT * FROM purchases WHERE name = ? AND invo = ? LIMIT 1`,
          [sellerName, String(invoiceNo)]
        );
      }
      if (!stored) {
        return res.status(404).json({ 
          error: `No purchase data found for seller "${sellerName}" with invoice ${invoiceNo}. Lots may have been deleted.` 
        });
      }
      // Reconstruct minimal invoice object from stored data
      invoice = {
        seller: { name: stored.name, address: stored.add_line, place: stored.place, cr: stored.gstin, pan: '', state: stored.state },
        lineItems: [{ lot: '—', qty: stored.qty, pqty: stored.qty, price: 0, prate: 0, amount: stored.amount, puramt: stored.amount, com: 0, sertax: 0, cgst: stored.cgst, sgst: stored.sgst, igst: stored.igst }],
        summary: {
          totalQty: stored.qty, totalPuramt: stored.amount,
          totalCgst: stored.cgst, totalSgst: stored.sgst, totalIgst: stored.igst,
          roundDiff: stored.rund, grandTotal: stored.total,
          tdsAmount: stored.tds, invoiceAmount: stored.total - stored.tds,
          isInter: stored.igst > 0
        }
      };
    }
    
    const pdf = await generatePurchaseInvoicePDF(
      enrichPurchaseForPDF(invoice, cfg, db, auctionId), cfg, invoiceNo
    );
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `inline; filename="PurchaseInvoice_${sellerName.replace(/[^\w]/g, '_')}_${invoiceNo}.pdf"`);
    res.send(pdf);
  } catch (e) {
    console.error('PDF generation error:', e);
    res.status(500).json({ error: 'PDF generation failed: ' + e.message });
  }
});

// Attach buyer block + e-TRADE No + invoice date to a purchase invoice
// object so the new renderer can show the full BILLED/SHIPPED TO + top
// grid correctly. The buyer is whichever identity is currently active in
// the app (ASP when Kerala+e-Trade, else ISP from company_settings).
function enrichPurchaseForPDF(invoice, cfg, db, auctionId) {
  if (!invoice) return invoice;
  // Stamp auction date + e-TRADE no. The e-TRADE no shown on the invoice
  // must be the trade NUMBER (auctions.ano), never the internal row id.
  const auction = auctionId ? db.get('SELECT ano, date FROM auctions WHERE id = ?', [auctionId]) : null;
  if (!invoice.invoiceDate && auction && auction.date) {
    const d = new Date(auction.date);
    if (!isNaN(d)) invoice.invoiceDate = fmtDate(auction.date);
  }
  if (!invoice.invoiceDate) invoice.invoiceDate = fmtDate(todayLocalISO());
  if (!invoice.eTradeNo) invoice.eTradeNo = String((auction && auction.ano != null) ? auction.ano : (auctionId || ''));

  // Buyer block — populated from the central company identity (name /
  // PAN / etc.) and the state-specific address slot. Single-company
  // build: there is no sister/ASP. Kerala installs read `kl_*`, every
  // other state (TN by default) reads `tn_*`. The buyer GSTIN
  // specifically prefers the state's slot over `_ident.gstin` because
  // `getCompanyIdentity` only checks `tn_gstin`, which would print the
  // wrong GSTIN on a Kerala install whose user had filled `kl_gstin`.
  if (!invoice.buyer) {
    const _ident = getCompanyIdentity(cfg);
    const isKerala = String(cfg.business_state || '').toUpperCase() === 'KERALA';
    if (isKerala) {
      invoice.buyer = {
        name:    _ident.name      || cfg.short_name || cfg.trade_name || '',
        address: cfg.kl_address1   || _ident.address1 || '',
        place:   cfg.kl_place      || '',
        pin:     cfg.kl_pin        || '',
        state:   cfg.kl_state      || _ident.state || 'Kerala',
        st_code: '32',
        gstin:   cfg.kl_gstin      || _ident.gstin || '',
        pan:     _ident.pan        || cfg.pan || '',
      };
    } else {
      invoice.buyer = {
        name:    _ident.name      || cfg.short_name || cfg.trade_name || '',
        address: cfg.tn_address1   || _ident.address1 || '',
        place:   cfg.tn_place      || '',
        pin:     cfg.tn_pin        || '',
        state:   cfg.tn_state      || _ident.state || 'Tamil Nadu',
        st_code: cfg.tn_st_code    || _ident.stateCode || '33',
        gstin:   cfg.tn_gstin      || _ident.gstin || '',
        pan:     _ident.pan        || cfg.pan || '',
      };
    }
  }
  return invoice;
}

// Bulk Purchase Invoice PDF — merges N purchases into a single PDF
// Body: { ids: [1, 2, 3, ...] } — database row IDs from `purchases` table.
// Returns: one PDF with each purchase on its own page(s), in the order given.
// Same rebuild-from-lots OR fallback-to-stored pattern as the single route.
app.post('/api/purchases/pdf-bulk', requireView, async (req, res) => {
  try {
    const db = getDb();
    const cfg = getSettingsFlat(db);
    const ids = Array.isArray(req.body?.ids) ? req.body.ids.filter(Boolean) : [];
    if (!ids.length) return res.status(400).json({ error: 'No purchase IDs provided' });

    const placeholders = ids.map(() => '?').join(',');
    const rows = db.all(`SELECT * FROM purchases WHERE id IN (${placeholders})`, ids);
    if (!rows.length) return res.status(404).json({ error: 'No matching purchases found' });

    // Preserve the order the user ticked them in by looking up each ID in
    // the returned set (the IN query doesn't preserve order).
    const byId = new Map(rows.map(r => [r.id, r]));
    const ordered = ids.map(id => byId.get(Number(id))).filter(Boolean);

    const payloads = [];
    for (const stored of ordered) {
      // Try fresh rebuild from lots first (richer line-item detail)
      let invoice = stored.auction_id
        ? buildPurchaseInvoice(db, stored.auction_id, stored.name, cfg)
        : null;
      if (!invoice) {
        // Fallback: stored summary only (one line item)
        invoice = {
          seller: {
            name: stored.name, address: stored.add_line, place: stored.place,
            cr: stored.gstin, pan: '', state: stored.state
          },
          lineItems: [{
            lot: '—', qty: stored.qty, pqty: stored.qty, price: 0, prate: 0,
            amount: stored.amount, puramt: stored.amount,
            com: 0, sertax: 0, cgst: stored.cgst, sgst: stored.sgst, igst: stored.igst
          }],
          summary: {
            totalQty: stored.qty, totalPuramt: stored.amount,
            totalCgst: stored.cgst, totalSgst: stored.sgst, totalIgst: stored.igst,
            roundDiff: stored.rund, grandTotal: stored.total,
            tdsAmount: stored.tds, invoiceAmount: stored.total - stored.tds,
            isInter: stored.igst > 0
          }
        };
      }
      payloads.push({ invoiceData: enrichPurchaseForPDF(invoice, cfg, db, stored.auction_id), invoiceNo: stored.invo });
    }

    const pdf = await generatePurchaseInvoicesBatchPDF(payloads, cfg);
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `inline; filename="PurchaseInvoices_Batch_${payloads.length}.pdf"`);
    res.send(pdf);
  } catch (e) {
    console.error('Bulk purchase PDF error:', e);
    res.status(500).json({ error: 'Bulk PDF generation failed: ' + e.message });
  }
});

// ══════════════════════════════════════════════════════════════
// BILLS — Agriculturist Bills of Supply (GSTKBILP/GSTBILP)
// ══════════════════════════════════════════════════════════════
app.get('/api/bills', requireView, (req, res) => {
  const { auction_id, ano, from, to, branch, search } = req.query;
  let q = 'SELECT * FROM bills WHERE 1=1'; const p = [];
  if (auction_id) { q += ' AND auction_id = ?'; p.push(parseInt(auction_id)); }
  if (ano) { q += ' AND ano = ?'; p.push(ano); }
  // Free-text search within the selected trade — bill no, seller name.
  const searchTerm = String(search || '').trim();
  if (searchTerm) {
    const wild = `%${searchTerm}%`;
    q += ` AND (
            COALESCE(bil,'')  LIKE ?
            OR COALESCE(name,'') LIKE ?
          )`;
    p.push(wild, wild);
  }
  if (from && to) { q += ' AND date BETWEEN ? AND ?'; p.push(from, to); }
  // Branch filter — bills.br may not be set on legacy/auto-generated
  // rows (this codebase inserts it as empty), so the match must ALSO
  // infer the branch via the underlying lot rows for the same seller.
  //   bills.br = branch         (when populated)  OR
  //   lots.branch via JOIN      (auction_id + seller name match)
  // Both paths case-insensitive + trim-tolerant so legacy data with
  // mixed casing still matches.
  const branchFilter = String(branch || '').trim();
  if (branchFilter) {
    q += ` AND (
            UPPER(TRIM(COALESCE(br,''))) = UPPER(TRIM(?))
            OR EXISTS (
              SELECT 1 FROM lots l
               WHERE l.auction_id = COALESCE(
                       bills.auction_id,
                       (SELECT a.id FROM auctions a WHERE a.ano = bills.ano LIMIT 1)
                     )
                 AND UPPER(TRIM(COALESCE(l.name,''))) = UPPER(TRIM(COALESCE(bills.name,'')))
                 AND UPPER(TRIM(COALESCE(l.branch,''))) = UPPER(TRIM(?))
            )
          )`;
    p.push(branchFilter, branchFilter);
  }
  q += ' ORDER BY date DESC, bil DESC LIMIT 500';
  res.json(withFmtDate(getDb().all(q, p)));
});

// Generate agri bill for a seller
app.post('/api/bills/generate/:auctionId',
  requireInvoiceWrite,
  requirePriceChecked(req => parseInt(req.params.auctionId, 10)),
  (req, res) => {
  const db = getDb(); const cfg = getSettingsFlat(db);
  const auctionIdForGate = parseInt(req.params.auctionId, 10);
  const _gen = _checkGenerationGate(db, 'bills', auctionIdForGate);
  if (!_gen.allowed) return res.status(412).json(_gen.error);
  const { sellerName, billNo } = req.body;
  
  if (!sellerName || !billNo) {
    return res.status(400).json({ error: 'sellerName and billNo are required' });
  }
  
  // Auto-calculate if needed
  const uncalc = db.all(`SELECT * FROM lots WHERE auction_id = ? AND amount > 0 AND (puramt IS NULL OR puramt = 0)`, [req.params.auctionId]);
  for (const lot of uncalc) {
    const c = calculateLot(lot, cfg);
    db.run(`UPDATE lots SET pqty=?,prate=?,puramt=?,com=?,sertax=?,cgst=?,sgst=?,igst=?,advance=?,balance=?,bilamt=?,refund=?,refud=?,isp_pqty=?,isp_prate=?,isp_puramt=?,asp_pqty=?,asp_prate=?,asp_puramt=? WHERE id=?`,
      [c.pqty,c.prate,c.puramt,c.com,c.sertax,c.cgst,c.sgst,c.igst,c.advance,c.balance,c.bilamt,c.refund||0,c.refud||0,c.isp_pqty||0,c.isp_prate||0,c.isp_puramt||0,c.asp_pqty||0,c.asp_prate||0,c.asp_puramt||0,lot.id]);
  }
  
  const bill = buildAgriBill(db, req.params.auctionId, sellerName, cfg);
  if (!bill || bill.error) {
    return res.status(404).json({ error: bill?.error || 'No eligible lots found' });
  }

  const auction = db.get('SELECT * FROM auctions WHERE id = ?', [req.params.auctionId]);
  // Dedup — refuse a second bill for the same seller in this trade,
  // and refuse to reuse a bill number that's already taken.
  const dupSeller = db.get(
    'SELECT id, bil FROM bills WHERE auction_id = ? AND name = ? LIMIT 1',
    [req.params.auctionId, bill.seller.name]
  );
  if (dupSeller) {
    return res.status(409).json({
      error: `Bill of supply already exists for seller "${bill.seller.name}" in this trade — bill #${dupSeller.bil}.`,
      existingId: dupSeller.id, existingBil: dupSeller.bil,
    });
  }
  const dupNo = db.get(
    'SELECT id, name FROM bills WHERE auction_id = ? AND bil = ? LIMIT 1',
    [req.params.auctionId, parseInt(billNo)]
  );
  if (dupNo) {
    return res.status(409).json({
      error: `Bill number ${billNo} is already used in this trade by seller "${dupNo.name}".`,
      existingId: dupNo.id, existingSeller: dupNo.name,
    });
  }
  const s = bill.summary;
  db.run(`INSERT INTO bills (auction_id,ano,date,state,br,crpt,bil,name,add_line,pla,pstate,st_code,crr,pan,qty,cost,igst,net)
    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`,
    [req.params.auctionId,auction.ano,auction.date,auction.state||'','',auction.crop_type||(getSetting(getDb(),'default_crop_type')||'VST'),
     parseInt(billNo),bill.seller.name,bill.seller.address||'',bill.seller.place||'',
     bill.seller.state||'',bill.seller.st_code||'',bill.seller.cr||'',bill.seller.pan||'',
     s.totalQty,s.totalPuramt,0,s.netAmount]);
  
  res.json({ success: true, bill: s });
});

// List eligible agri sellers for an auction (no GSTIN + amount > 0)
app.get('/api/bills/eligible-sellers/:auctionId', requireView, (req, res) => {
  res.json(listAgriSellers(getDb(), req.params.auctionId));
});

// Batch: generate bill of supply for ALL agriculturists (no GSTIN) in an auction
app.post('/api/bills/generate-all/:auctionId',
  requireInvoiceWrite,
  requirePriceChecked(req => parseInt(req.params.auctionId, 10)),
  (req, res) => {
  const db = getDb(); const cfg = getSettingsFlat(db);
  const auctionIdForGate = parseInt(req.params.auctionId, 10);
  const _gen = _checkGenerationGate(db, 'bills', auctionIdForGate);
  if (!_gen.allowed) return res.status(412).json(_gen.error);
  const { startBillNo } = req.body;
  
  let nextNo = parseInt(startBillNo);
  if (!nextNo || nextNo < 1) return res.status(400).json({ error: 'startBillNo must be a positive integer' });
  
  // Auto-calculate
  const uncalc = db.all(`SELECT * FROM lots WHERE auction_id = ? AND amount > 0 AND (puramt IS NULL OR puramt = 0)`, [req.params.auctionId]);
  for (const lot of uncalc) {
    const c = calculateLot(lot, cfg);
    db.run(`UPDATE lots SET pqty=?,prate=?,puramt=?,com=?,sertax=?,cgst=?,sgst=?,igst=?,advance=?,balance=?,bilamt=?,refund=?,refud=?,isp_pqty=?,isp_prate=?,isp_puramt=?,asp_pqty=?,asp_prate=?,asp_puramt=? WHERE id=?`,
      [c.pqty,c.prate,c.puramt,c.com,c.sertax,c.cgst,c.sgst,c.igst,c.advance,c.balance,c.bilamt,c.refund||0,c.refud||0,c.isp_pqty||0,c.isp_prate||0,c.isp_puramt||0,c.asp_pqty||0,c.asp_prate||0,c.asp_puramt||0,lot.id]);
  }
  
  const sellers = listAgriSellers(db, req.params.auctionId);
  if (!sellers.length) return res.status(404).json({ error: 'No agriculturist sellers (without GSTIN) with lots in this auction' });
  
  const auction = db.get('SELECT * FROM auctions WHERE id = ?', [req.params.auctionId]);
  const results = [];
  const errors = [];
  
  for (const row of sellers) {
    try {
      // Skip sellers who already have a bill in this trade.
      const dupSeller = db.get(
        'SELECT id, bil FROM bills WHERE auction_id = ? AND name = ? LIMIT 1',
        [req.params.auctionId, row.name]
      );
      if (dupSeller) {
        errors.push({ seller: row.name, error: `Already billed as #${dupSeller.bil}` });
        continue;
      }
      const bill = buildAgriBill(db, req.params.auctionId, row.name, cfg);
      if (!bill || bill.error) { errors.push({ seller: row.name, error: bill?.error || 'Build failed' }); continue; }
      const s = bill.summary;
      // Skip if this bill number is already taken in this trade.
      const dupNo = db.get(
        'SELECT id, name FROM bills WHERE auction_id = ? AND bil = ? LIMIT 1',
        [req.params.auctionId, nextNo]
      );
      if (dupNo) {
        errors.push({ seller: row.name, error: `Bill #${nextNo} already used by ${dupNo.name}` });
        continue;
      }
      db.run(`INSERT INTO bills (auction_id,ano,date,state,br,crpt,bil,name,add_line,pla,pstate,st_code,crr,pan,qty,cost,igst,net)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`,
        [req.params.auctionId,auction.ano,auction.date,auction.state||'','',auction.crop_type||(getSetting(getDb(),'default_crop_type')||'VST'),
         nextNo,bill.seller.name,bill.seller.address||'',bill.seller.place||'',
         bill.seller.state||'',bill.seller.st_code||'',bill.seller.cr||'',bill.seller.pan||'',
         s.totalQty,s.totalPuramt,0,s.netAmount]);
      results.push({ seller: row.name, billNo: nextNo, netAmount: s.netAmount });
      nextNo++;
    } catch (e) { errors.push({ seller: row.name, error: e.message }); }
  }
  
  res.json({ success: true, generated: results.length, results, errors });
});

// Agri bill PDF
app.get('/api/bills/pdf/:auctionId/:sellerName', requireView, async (req, res) => {
  try {
    const db = getDb(); const cfg = getSettingsFlat(db);
    const sellerName = decodeURIComponent(req.params.sellerName);
    const billNo = req.query.billNo || '001';
    
    let bill = buildAgriBill(db, req.params.auctionId, sellerName, cfg);
    if (!bill || bill.error) {
      // Fallback to stored record
      const stored = db.get('SELECT * FROM bills WHERE name = ? AND bil = ? LIMIT 1', [sellerName, parseInt(billNo)]);
      if (!stored) return res.status(404).json({ error: bill?.error || `No bill data found for "${sellerName}"` });
      bill = {
        seller: { name: stored.name, address: stored.add_line, place: stored.pla, state: stored.pstate, st_code: stored.st_code, cr: stored.crr, crno: stored.crr, pan: stored.pan },
        lineItems: [{ lot: '—', qty: stored.qty, pqty: stored.qty, prate: 0, amount: stored.cost, puramt: stored.cost }],
        summary: { totalQty: stored.qty, totalPuramt: stored.cost, roundDiff: 0, netAmount: stored.net, cgst: 0, sgst: 0, igst: 0, tax: 0 }
      };
    }
    // Enrich seller.crno so the new renderer can display "CR.<n>" in the
    // details block when CR/GSTIN-style id is available on the trader row
    if (bill.seller && !bill.seller.crno) bill.seller.crno = bill.seller.cr || '';
    // Stamp the bill date + e-TRADE number so the new layout can render them
    // in the top strip (Invoice No / e-TRADE No / Date).
    const auction = db.get('SELECT ano, date FROM auctions WHERE id = ?', [req.params.auctionId]);
    if (auction && auction.date) {
      const d = new Date(auction.date);
      if (!isNaN(d)) bill.billDate = fmtDate(auction.date);
    }
    if (!bill.billDate) bill.billDate = fmtDate(todayLocalISO());
    // e-TRADE No is the trade NUMBER (auctions.ano), not the row id.
    bill.eTradeNo = req.query.eTradeNo || (auction && auction.ano != null ? String(auction.ano) : req.params.auctionId);

    const pdf = await generateAgriBillPDF(bill, cfg, billNo);
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `inline; filename="BillOfSupply_${sellerName.replace(/[^\w]/g,'_')}_${billNo}.pdf"`);
    res.send(pdf);
  } catch(e) {
    console.error('Agri Bill PDF error:', e);
    res.status(500).json({ error: 'PDF generation failed: ' + e.message });
  }
});

// Bulk Agri Bill PDF — merges N bills into a single PDF.
// Body: { ids: [1, 2, 3, ...] } — DB row IDs from the `bills` table.
app.post('/api/bills/pdf-bulk', requireView, async (req, res) => {
  try {
    const db = getDb();
    const cfg = getSettingsFlat(db);
    const ids = Array.isArray(req.body?.ids) ? req.body.ids.filter(Boolean) : [];
    if (!ids.length) return res.status(400).json({ error: 'No bill IDs provided' });

    const placeholders = ids.map(() => '?').join(',');
    const rows = db.all(`SELECT * FROM bills WHERE id IN (${placeholders})`, ids);
    if (!rows.length) return res.status(404).json({ error: 'No matching bills found' });

    const byId = new Map(rows.map(r => [r.id, r]));
    const ordered = ids.map(id => byId.get(Number(id))).filter(Boolean);

    const payloads = [];
    for (const stored of ordered) {
      let bill = stored.auction_id
        ? buildAgriBill(db, stored.auction_id, stored.name, cfg)
        : null;
      if (!bill || bill.error) {
        bill = {
          seller: {
            name: stored.name, address: stored.add_line, place: stored.pla,
            state: stored.pstate, st_code: stored.st_code, cr: stored.crr, crno: stored.crr, pan: stored.pan
          },
          lineItems: [{
            lot: '—', qty: stored.qty, pqty: stored.qty, prate: 0,
            amount: stored.cost, puramt: stored.cost
          }],
          summary: {
            totalQty: stored.qty, totalPuramt: stored.cost, roundDiff: 0,
            netAmount: stored.net, cgst: 0, sgst: 0, igst: 0, tax: 0
          }
        };
      }
      // Enrich for new renderer layout (Invoice No / e-TRADE No / Date strip)
      if (bill.seller && !bill.seller.crno) bill.seller.crno = bill.seller.cr || '';
      // e-TRADE No is the trade NUMBER. Prefer the bill's own `ano`
      // column; fall back to the auction's ano — never the row id.
      let billAno = stored.ano != null ? String(stored.ano) : '';
      if (stored.auction_id) {
        const auction = db.get('SELECT ano, date FROM auctions WHERE id = ?', [stored.auction_id]);
        if (auction && auction.date) {
          const d = new Date(auction.date);
          if (!isNaN(d)) bill.billDate = fmtDate(auction.date);
        }
        if (!billAno && auction && auction.ano != null) billAno = String(auction.ano);
      }
      if (!bill.billDate) bill.billDate = fmtDate(todayLocalISO());
      bill.eTradeNo = billAno;
      payloads.push({ billData: bill, billNo: stored.bil });
    }

    const pdf = await generateAgriBillsBatchPDF(payloads, cfg);
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `inline; filename="BillsOfSupply_Batch_${payloads.length}.pdf"`);
    res.send(pdf);
  } catch (e) {
    console.error('Bulk bill PDF error:', e);
    res.status(500).json({ error: 'Bulk PDF generation failed: ' + e.message });
  }
});

app.delete('/api/bills/:id', requireDelete, (req, res) => {
  getDb().run('DELETE FROM bills WHERE id = ?', [req.params.id]);
  res.json({ success: true });
});

// Update bill (edit)
app.put('/api/bills/:id', requireInvoiceWrite, (req, res) => {
  const b = req.body;
  const fields = ['ano','date','state','br','crpt','bil','name','add_line','pla','pstate','st_code','crr','pan','qty','cost','igst','net'];
  const sets = []; const vals = [];
  for (const f of fields) {
    if (b[f] !== undefined) { sets.push(`${f}=?`); vals.push(b[f]); }
  }
  if (!sets.length) return res.status(400).json({ error: 'No fields to update' });
  vals.push(req.params.id);
  getDb().run(`UPDATE bills SET ${sets.join(',')} WHERE id=?`, vals);
  res.json({ success: true });
});

// ══════════════════════════════════════════════════════════════
// DEBIT NOTES (for discounts/adjustments)
// ══════════════════════════════════════════════════════════════
app.get('/api/debit-notes', requireView, (req, res) => {
  const { auction_id, ano, from, to, search } = req.query;
  let q = 'SELECT * FROM debit_notes WHERE 1=1'; const p = [];
  if (auction_id) { q += ' AND auction_id = ?'; p.push(parseInt(auction_id)); }
  if (ano) { q += ' AND ano = ?'; p.push(ano); }
  // Free-text search within the selected trade — note no, seller name.
  // The trade filter is applied client-side in loadDebitNotes (by ano)
  // since debit_notes has no auction_id column.
  const searchTerm = String(search || '').trim();
  if (searchTerm) {
    const wild = `%${searchTerm}%`;
    q += ` AND (
            COALESCE(note_no,'') LIKE ?
            OR COALESCE(name,'')    LIKE ?
          )`;
    p.push(wild, wild);
  }
  if (from && to) { q += ' AND date BETWEEN ? AND ?'; p.push(from, to); }
  q += ' ORDER BY date DESC, note_no DESC LIMIT 500';
  res.json(withFmtDate(getDb().all(q, p)));
});

// Single-shot DN generation. Input: purchase invoice number + explicit
// discount + note number. Source is ALWAYS purchases — DNs against sales
// transactions are not allowed (they're a buyer-side instrument issued
// to suppliers). The legacy `saleType` parameter is accepted for
// compatibility with older clients but ignored — purchases don't carry
// sale type the way sales invoices do.
//
// Sibling endpoints:
//   /generate-bulk  — auto-derive discount from settings (one purchase)
//   /generate-all   — same logic over every purchase in the DB
// Generate a SINGLE debit note. Inputs (new contract):
//   - purchno : the purchase invoice number (lots up to one row in `purchases`)
//   - ano     : the trade number (REQUIRED — invoice must belong to it)
//
// Strict trade-membership check: the resolved purchase row's `ano` MUST
// equal the supplied `ano`. Cross-trade attempts get a clean 400 with
// the actual trade the invoice belongs to, so the user can self-correct
// without inspecting the DB.
//
// Discount + note number are server-derived (same rules as /generate-bulk):
//   discount = round((amount/1000) × discount_days × discount_pct)
//              or amount × discount_pct/100 if days isn't set
//   note_no  = MAX(note_no) + 1
//
// Legacy callers still passing `{ discount, noteNo }` are accepted —
// explicit values override the derivation. `saleType` is accepted and
// ignored (kept for back-compat with very old API consumers).
app.post('/api/debit-notes/generate', requireInvoiceWrite, (req, res) => {
  const db = getDb();
  const cfg = getSettingsFlat(db);
  const purchno = String(req.body.purchno || req.body.invoiceNo || '').trim();
  const ano     = String(req.body.ano || '').trim();

  if (!purchno) return res.status(400).json({ error: 'purchno (purchase invoice number) is required' });
  if (!ano)     return res.status(400).json({ error: 'ano (trade number) is required' });

  // Price-check gate: resolve auction by ano (debit notes don't carry
  // the numeric id directly), then refuse if its lots haven't been
  // verified yet. Same 412 contract as the URL-bound gates so the
  // client can show a uniform "Run price check first" message.
  // Short-circuit when the feature flag is OFF — keeps debit notes
  // generation usable on installs that don't run price check.
  if (pcFlagOn(db)) {
    // Tri-state gate: only block when the auction has NEVER been verified.
    // Once it has passed at least once, subsequent lot edits drop it to
    // 'stale' (soft warning, allowed) — matching the URL-bound middleware.
    const gateAuction = db.get('SELECT id, price_check_first_passed_at FROM auctions WHERE ano = ? ORDER BY date DESC LIMIT 1', [ano]);
    if (gateAuction && !gateAuction.price_check_first_passed_at) {
      return res.status(412).json({
        error: 'Price check required',
        detail: 'Run Reports → Price Check against the auction (and apply any code fixes) before generating debit notes.',
        auctionId: gateAuction.id,
        gate: 'price_check',
      });
    }
  }

  // Generation lock: once every purchase in this trade has a debit
  // note, the generate is blocked until an admin grants a one-shot
  // override. While any purchase is still un-DN'd it stays open.
  const _dnAuc = db.get('SELECT id FROM auctions WHERE ano = ? ORDER BY date DESC LIMIT 1', [ano]);
  if (_dnAuc) {
    const _gen = _checkGenerationGate(db, 'debit_notes', _dnAuc.id);
    if (!_gen.allowed) return res.status(412).json(_gen.error);
  }

  // Look up the purchase. We find ALL rows with this purchno (in case of
  // legacy duplicates across trades) and pick the one that matches the
  // selected trade. Earlier code took the most-recent and silently
  // generated against whatever trade it landed on — that broke the
  // strict trade-scoping requirement.
  const candidates = db.all(
    `SELECT * FROM purchases WHERE invo = ? ORDER BY date DESC, id DESC`,
    [purchno]
  );
  if (!candidates.length) {
    // Distinguish "is a sales invoice" from "doesn't exist" for cleaner UX.
    const isSalesInv = db.get(
      `SELECT id FROM invoices WHERE invo = ? LIMIT 1`,
      [purchno]
    );
    if (isSalesInv) {
      return res.status(400).json({
        error: `${purchno} is a SALES invoice. Debit notes can only be generated against PURCHASE invoices.`
      });
    }
    return res.status(404).json({ error: `Purchase invoice ${purchno} not found` });
  }

  // Trade-membership filter — the invoice MUST belong to the requested
  // trade. If only mismatching candidates exist, surface the actual
  // ano(s) so the user knows which trade to switch to.
  const purchase = candidates.find(p => String(p.ano) === ano);
  if (!purchase) {
    const otherAnos = [...new Set(candidates.map(p => String(p.ano)))].join(', ');
    return res.status(400).json({
      error: `Purchase invoice ${purchno} does not belong to trade #${ano}. ` +
             `It belongs to trade #${otherAnos}.`
    });
  }

  // Idempotency: skip if a DN for (ano, dealer) already exists.
  const dealerName = purchase.name || '';
  const dupe = db.get(
    `SELECT id, note_no FROM debit_notes WHERE ano = ? AND name = ? LIMIT 1`,
    [ano, dealerName]
  );
  if (dupe) {
    return res.status(409).json({
      error: `Debit note #${dupe.note_no} already exists for ${dealerName} in trade #${ano}`,
      existingId: dupe.id,
      existingNoteNo: dupe.note_no,
    });
  }

  // Discount: explicit override > computed from settings.
  const baseAmt = Number(purchase.amount || 0);
  if (baseAmt <= 0) {
    return res.status(400).json({ error: 'Purchase amount is zero — cannot compute discount' });
  }
  let discountAmt = req.body.discount != null ? parseFloat(req.body.discount) : NaN;
  if (!Number.isFinite(discountAmt) || discountAmt <= 0) {
    // "Discount In Payments" master switch (flag_sample) gates auto-computed
    // discount. When OFF, don't calculate one — the caller can still pass an
    // explicit amount to record a manual discount.
    const discountEnabled = String(cfg.flag_sample || '').toLowerCase() === 'true' || cfg.flag_sample === true;
    if (!discountEnabled) {
      return res.status(400).json({ error: 'Discount In Payments is turned off — enable it in Settings → Feature Flags, or pass an explicit discount amount.' });
    }
    const discountPct  = Number(cfg.discount_pct)  || 0;
    // Days source depends on the seller type — same rule the per-lot
    // refund calc uses ([calculations.js calculateLot]):
    //   • Has GSTIN (registered dealer) → cfg.dealer_days
    //   • No GSTIN (CR / agriculturist)  → cfg.discount_days
    // Read the dealer's GSTIN from the purchase row's `gstin` column
    // (same source the intra/inter classification uses below).
    let _g = String(purchase.gstin || '').trim().toUpperCase();
    if (_g.startsWith('GSTIN.')) _g = _g.slice(6);
    else if (_g.startsWith('GSTIN')) _g = _g.slice(5);
    const sellerHasGstin = /^\d{2}/.test(_g);
    const discountDays = sellerHasGstin
      ? (Number(cfg.dealer_days)   || 0)
      : (Number(cfg.discount_days) || 0);
    if (discountPct <= 0) {
      return res.status(400).json({ error: 'Discount % not configured in settings' });
    }
    discountAmt = discountDays > 0
      ? Math.round((baseAmt / 1000) * discountDays * discountPct)
      : Math.round(baseAmt * discountPct / 100);
  }
  if (discountAmt <= 0) {
    return res.status(400).json({ error: 'Computed discount is zero — check settings or invoice amount' });
  }

  // GST split — intra/inter classification by the DEALER's GSTIN
  // state code (NOT by purchase.igst, which can be stale or
  // misclassified). See the equivalent block in /generate-bulk for
  // detailed rationale.
  // Dealer state — `purchases.gstin` (the column actually stored on the
  // row). Earlier code read `purchase.cr` which doesn't exist on the
  // purchases schema; that silently resolved to '' and forced every DN
  // to the intra-state branch.
  let dealerStateCode = '';
  {
    let g = String(purchase.gstin || '').trim().toUpperCase();
    if (g.startsWith('GSTIN.')) g = g.slice(6);
    else if (g.startsWith('GSTIN')) g = g.slice(5);
    if (/^\d{2}/.test(g)) dealerStateCode = g.slice(0, 2);
  }
  const companyStateCode = String(cfg.tally_state_code
      || (String(cfg.business_state || '').toUpperCase() === 'KERALA' ? '32' : '33'));
  const isInter = !!dealerStateCode && dealerStateCode !== companyStateCode;

  const dnGstRate = Number(cfg.discount_gst) || Number(cfg.gst_service) || 18;
  // Gate discount-GST on the flag itself, not on whether the source
  // purchase already happens to carry GST. Older code used the latter as
  // a proxy, but the proxy goes stale the moment a user toggles
  // flag_disc_gst — purchases generated in the previous flag state still
  // have non-zero cgst/sgst/igst and cause new DNs to incorrectly add
  // tax. Secondary `dealerCarriedGst` filter retained: URD / agriculturist
  // purchases (no GSTIN, no GST stamped) still produce exempt DNs.
  const flagDiscGst = String(cfg.flag_disc_gst || '').toLowerCase() === 'true' || cfg.flag_disc_gst === true;
  const dealerCarriedGst = Number(purchase.cgst) || Number(purchase.sgst) || Number(purchase.igst);
  let cgst = 0, sgst = 0, igst = 0;
  if (flagDiscGst && dealerCarriedGst) {
    if (isInter) {
      igst = round2(discountAmt * dnGstRate / 100);
    } else {
      const half = round2(discountAmt * (dnGstRate / 2) / 100);
      cgst = half; sgst = half;
    }
  }
  const total = round2(discountAmt + cgst + sgst + igst);

  // DN date = trade.date + 1
  const trade = db.get('SELECT date FROM auctions WHERE ano = ? LIMIT 1', [ano]);
  const dnDate = trade && trade.date
    ? addDays(trade.date, 1)
    : new Date().toISOString().slice(0, 10);

  // Note number: client-supplied `startNoteNo` (preferred) or legacy
  // `noteNo` (back-compat). When neither is provided, fall back to
  // MAX(note_no)+1. The user-supplied value is validated as a positive
  // integer and checked for uniqueness against debit_notes — a clean
  // 409 is returned if the number is already taken so the user can pick
  // a different start.
  //
  // TRADE-WISE numbering: every uniqueness/MAX check is scoped to the
  // current trade `ano`. Each trade has its own independent 1..N
  // sequence — DN #5 in trade 7 is unrelated to DN #5 in trade 8.
  const rawStart = req.body.startNoteNo != null ? req.body.startNoteNo : req.body.noteNo;
  let noteNo;
  if (rawStart != null && String(rawStart).trim() !== '') {
    const n = parseInt(String(rawStart).trim(), 10);
    if (!Number.isFinite(n) || n < 1) {
      return res.status(400).json({ error: 'Starting Number must be a positive integer' });
    }
    noteNo = String(n);
    // Uniqueness check — scoped to the SELECTED TRADE only. Note_no is
    // stored as TEXT but we compare as integer to handle rare cases
    // where one side has leading zeroes.
    const taken = db.get(
      `SELECT id FROM debit_notes WHERE ano = ? AND CAST(note_no AS INTEGER) = ? LIMIT 1`,
      [ano, n]
    );
    if (taken) {
      return res.status(409).json({
        error: `Debit note #${n} is already used in trade #${ano}. Choose a different number.`,
        suggested: (() => {
          const row = db.get(
            'SELECT MAX(CAST(note_no AS INTEGER)) AS mx FROM debit_notes WHERE ano = ?',
            [ano]
          );
          const mx = parseInt(row && row.mx, 10);
          return Number.isFinite(mx) && mx > 0 ? mx + 1 : 1;
        })(),
      });
    }
  } else {
    // No explicit start → bump the trade's own MAX.
    const row = db.get(
      'SELECT MAX(CAST(note_no AS INTEGER)) AS mx FROM debit_notes WHERE ano = ?',
      [ano]
    );
    const mx = parseInt(row && row.mx, 10);
    noteNo = String(Number.isFinite(mx) && mx > 0 ? mx + 1 : 1);
  }

  db.run(
    `INSERT INTO debit_notes (ano,date,state,name,note_no,amount,cgst,sgst,igst,total)
     VALUES (?,?,?,?,?,?,?,?,?,?)`,
    [ano, dnDate, purchase.state || '', dealerName,
     noteNo, discountAmt, cgst, sgst, igst, total]
  );

  res.json({
    success: true,
    created: 1,
    note_no: noteNo,
    purchno,
    ano,
    dealer: dealerName,
    amount: discountAmt,
    cgst, sgst, igst, total,
  });
});

// Bulk DN generation — single input: purchase invoice number. Finds the
// trade (`ano`) that purchase belongs to, then generates one DN per
// sales invoice in that trade. Idempotent: invoices that already have a
// DN matching their (ano, name) are skipped.
//
// Discount is auto-derived from settings:
//   discount_pct × invoice qty × purchase rate
// (mirrors the per-lot discount calculation in calculations.js — keeps
// the bulk DN consistent with what a user would have manually entered).
// Trade-wise DN generation. Input: a trade number (`ano`). The endpoint
// finds EVERY purchase invoice in that trade and creates one DN per
// purchase (one DN per registered dealer in that auction). Mirrors the
// natural data model:
//   - A trade is the unit of business activity (one auction day)
//   - Purchases roll up per-dealer for that trade
//   - Discount adjustments are negotiated trade-wide (per-day basis)
//   - DN date = trade.date + 1
//
// Backwards compatibility: legacy callers passing `{ purchno }` (single
// purchase invoice number) still work — we look up the purchase, derive
// its trade `ano`, and treat it as a trade-wide generation. This way
// older UI builds aren't broken if they're still in browser caches.
//
// Per-DN math (same as before):
//   discount  = round((amount / 1000) × discount_days × discount_pct)
//               or amount × discount_pct / 100 if days isn't set
//   GST split = mirrors source purchase (intra → CGST+SGST, inter → IGST)
//   GST rate  = cfg.discount_gst (typically 18%) — DN is a service charge
//
// Idempotency: skip per (ano, dealer name) pair already DN'd. Returns
// { created, skipped, skippedDetails[], generated[] }.
app.post('/api/debit-notes/generate-bulk', requireInvoiceWrite, (req, res) => {
  const db = getDb();
  const cfg = getSettingsFlat(db);

  // Resolve trade number (`ano`). Two input shapes:
  //   1. `{ ano }`     — preferred, trade-wise (new UI)
  //   2. `{ purchno }` — legacy, derive ano from the single purchase
  let ano = String(req.body.ano || '').trim();
  if (!ano) {
    const purchno = String(req.body.purchno || '').trim();
    if (!purchno) {
      return res.status(400).json({ error: 'Trade number (ano) is required' });
    }
    const p = db.get(
      `SELECT ano FROM purchases WHERE invo = ? ORDER BY date DESC, id DESC LIMIT 1`,
      [purchno]
    );
    if (!p) return res.status(404).json({ error: `Purchase invoice ${purchno} not found` });
    ano = String(p.ano || '').trim();
    if (!ano) return res.status(400).json({ error: 'Purchase row has no trade number' });
  }

  // Heal any trade-rename ano desync up-front so the purchase/DN lookups below
  // (all keyed by ano) see this trade's rows even if it was renumbered after
  // the purchases were created. Bridged by the stable auction_id.
  {
    const _ra = db.get('SELECT id FROM auctions WHERE ano = ? ORDER BY date DESC LIMIT 1', [ano]);
    if (_ra) _resyncTradeAno(db, _ra.id, ano);
  }

  // Price-check gate — only when the feature is enabled.
  // Tri-state: hard 412 only when the auction was never verified. After
  // a first pass, edits drop it to 'stale' (soft warning, allowed).
  if (pcFlagOn(db)) {
    const ga = db.get('SELECT id, price_check_first_passed_at FROM auctions WHERE ano = ? ORDER BY date DESC LIMIT 1', [ano]);
    if (ga && !ga.price_check_first_passed_at) {
      return res.status(412).json({
        error: 'Price check required',
        detail: 'Run Reports → Price Check against the auction (and apply any code fixes) before generating debit notes.',
        auctionId: ga.id,
        gate: 'price_check',
      });
    }
  }

  // Generation lock — admin must grant an override before re-running
  // generate-all once every purchase in this trade has a debit note.
  // While any purchase is still un-DN'd, generate-all stays open and
  // its dedupe layer skips the already-done rows.
  {
    const _dnAuc = db.get('SELECT id FROM auctions WHERE ano = ? ORDER BY date DESC LIMIT 1', [ano]);
    if (_dnAuc) {
      const _gen = _checkGenerationGate(db, 'debit_notes', _dnAuc.id);
      if (!_gen.allowed) return res.status(412).json(_gen.error);
    }
  }

  // Pull every purchase row for this trade. Each becomes one DN unless
  // already DN'd.
  const purchases = db.all(
    `SELECT * FROM purchases WHERE ano = ? ORDER BY id`,
    [ano]
  );
  if (!purchases.length) {
    return res.json({
      success: true, created: 0, skipped: 0, generated: [], skippedDetails: [],
      note: `No purchase invoices in trade #${ano}`,
    });
  }

  // Pre-load existing DN keys for this trade (single query — cheap).
  const existingKeys = new Set(
    db.all(
      `SELECT name FROM debit_notes WHERE ano = ?`,
      [ano]
    ).map(r => r.name || '')
  );

  // Resolve DN date once per trade — saves a query per purchase.
  const trade = db.get('SELECT date FROM auctions WHERE ano = ? LIMIT 1', [ano]);
  const dnDate = trade && trade.date
    ? addDays(trade.date, 1)
    : new Date().toISOString().slice(0, 10);

  // Discount math constants — read once, applied per-purchase.
  // Days source is per-purchase: GSTIN sellers use `dealer_days`,
  // non-GSTIN (CR / agriculturist) sellers use `discount_days`. Same
  // rule the per-lot refund calc applies.
  const discountPct = Number(cfg.discount_pct) || 0;
  const dealerDays  = Number(cfg.dealer_days)  || 0;
  const crDays      = Number(cfg.discount_days) || 0;
  const dnGstRate   = Number(cfg.discount_gst) || Number(cfg.gst_service) || 18;
  // Gate GST on the live `flag_disc_gst` flag. See the corresponding
  // comment in /api/debit-notes/generate — relying on the source
  // purchase's stamped GST as a proxy goes stale across flag toggles.
  const flagDiscGst = String(cfg.flag_disc_gst || '').toLowerCase() === 'true' || cfg.flag_disc_gst === true;
  // "Discount In Payments" master switch (flag_sample). When OFF, bulk
  // discount Debit Notes are not generated at all.
  const flagDiscount = String(cfg.flag_sample || '').toLowerCase() === 'true' || cfg.flag_sample === true;
  if (!flagDiscount) {
    return res.status(400).json({ error: 'Discount In Payments is turned off — enable it in Settings → Feature Flags to generate discount debit notes.' });
  }
  if (discountPct <= 0) {
    return res.status(400).json({ error: 'Discount % not configured in settings' });
  }

  // Resolve next note number. The user can supply `startNoteNo` to
  // explicitly anchor the sequence (each generated DN gets startNoteNo,
  // startNoteNo+1, +2, ...). When omitted, fall back to MAX+1.
  //
  // Concurrency: SQLite (sql.js / better-sqlite3) serializes writes, so
  // there's no within-request race. The only real risk is two SEPARATE
  // bulk requests overlapping their ranges. We mitigate that by claiming
  // the range up-front: count eligible purchases, then verify no number
  // in [start, start+eligibleCount) is already in `debit_notes`. If a
  // collision exists → 409 with the next safe start so the user can
  // retry. The check + INSERTs all run inside the same JS turn (no
  // await), so a competing bulk can't slip in between.
  const eligibleCount = purchases.filter(
    p => !existingKeys.has(p.name || '') && Number(p.amount || 0) > 0
  ).length;

  let nextNoteNo;
  const rawStart = req.body.startNoteNo != null ? req.body.startNoteNo : req.body.startInvoiceNo;
  if (rawStart != null && String(rawStart).trim() !== '') {
    const n = parseInt(String(rawStart).trim(), 10);
    if (!Number.isFinite(n) || n < 1) {
      return res.status(400).json({ error: 'Starting Number must be a positive integer' });
    }
    nextNoteNo = n;
    if (eligibleCount > 0) {
      // Range claim — scoped to THIS TRADE only. Numbering is per-trade
      // (trade #1's #5 doesn't conflict with trade #2's #5), so collision
      // check has `WHERE ano = ?`. Without this filter, starting trade 2
      // at #1 would falsely fail when trade 1 already has #1..N.
      const upper = nextNoteNo + eligibleCount - 1;
      const collisions = db.all(
        `SELECT CAST(note_no AS INTEGER) AS n
           FROM debit_notes
          WHERE ano = ?
            AND CAST(note_no AS INTEGER) BETWEEN ? AND ?
          ORDER BY n`,
        [ano, nextNoteNo, upper]
      );
      if (collisions.length) {
        const safe = (() => {
          const row = db.get(
            'SELECT MAX(CAST(note_no AS INTEGER)) AS mx FROM debit_notes WHERE ano = ?',
            [ano]
          );
          const mx = parseInt(row && row.mx, 10);
          return Number.isFinite(mx) && mx > 0 ? mx + 1 : 1;
        })();
        return res.status(409).json({
          error: `Starting Number ${nextNoteNo} would overlap existing debit note(s) in trade #${ano} ` +
                 `(${collisions.slice(0, 5).map(c => '#' + c.n).join(', ')}` +
                 `${collisions.length > 5 ? `, +${collisions.length - 5} more` : ''}). ` +
                 `Try ${safe} or higher.`,
          collisions: collisions.map(c => c.n),
          suggested: safe,
        });
      }
    }
  } else {
    // No explicit start → bump THIS TRADE's MAX. Each trade has its
    // own independent sequence.
    const row = db.get(
      'SELECT MAX(CAST(note_no AS INTEGER)) AS mx FROM debit_notes WHERE ano = ?',
      [ano]
    );
    const mx = parseInt(row && row.mx, 10);
    nextNoteNo = Number.isFinite(mx) && mx > 0 ? mx + 1 : 1;
  }

  const generated = [];
  const skipped   = [];

  for (const p of purchases) {
    const dealerName = p.name || '';
    if (existingKeys.has(dealerName)) {
      skipped.push({
        invo: p.invo, ano, buyer: dealerName,
        reason: 'duplicate (DN already exists for this dealer in this trade)',
      });
      continue;
    }
    const baseAmt = Number(p.amount || 0);
    if (baseAmt <= 0) {
      skipped.push({ invo: p.invo, ano, buyer: dealerName, reason: 'zero amount' });
      continue;
    }
    // Pick the right "days" config based on whether THIS purchase's
    // dealer carries a GSTIN. Mirrors calculateLot() so DN amount and
    // sum(lots.refund) stay in lockstep.
    let _dg = String(p.gstin || '').trim().toUpperCase();
    if (_dg.startsWith('GSTIN.')) _dg = _dg.slice(6);
    else if (_dg.startsWith('GSTIN')) _dg = _dg.slice(5);
    const sellerHasGstin = /^\d{2}/.test(_dg);
    const days = sellerHasGstin ? dealerDays : crDays;
    const discountAmt = days > 0
      ? Math.round((baseAmt / 1000) * days * discountPct)
      : Math.round(baseAmt * discountPct / 100);
    if (discountAmt <= 0) {
      skipped.push({ invo: p.invo, ano, buyer: dealerName, reason: 'computed discount is zero' });
      continue;
    }

    // Intra/inter classification — determined by the DEALER's GSTIN
    // state code, not by `purchases.igst > 0`. The earlier heuristic
    // failed for two real cases:
    //   1. The source purchase was booked as intra-state (igst=0) but
    //      we're issuing the DN to a registered dealer whose GSTIN
    //      starts with a different state code → DN must use IGST.
    //   2. Settings were updated after the purchase was created — the
    //      stale igst value lingers and misleads classification.
    //
    // Resolution: pull the dealer's `cr` (GSTIN field on purchases),
    // strip any "GSTIN."/"gstin." prefix, take the first two digits as
    // the state code, compare to the company's state code (intra).
    // Any non-match → inter-state → IGST applies.
    // Dealer state — `purchases.gstin` (real column on the row). The
    // older `p.cr` read resolved to undefined → empty → every DN was
    // misclassified as intra-state.
    let dealerStateCode = '';
    {
      let g = String(p.gstin || '').trim().toUpperCase();
      if (g.startsWith('GSTIN.')) g = g.slice(6);
      else if (g.startsWith('GSTIN')) g = g.slice(5);
      if (/^\d{2}/.test(g)) dealerStateCode = g.slice(0, 2);
    }
    const companyStateCode = String(cfg.tally_state_code
        || (String(cfg.business_state || '').toUpperCase() === 'KERALA' ? '32' : '33'));
    const isInter = !!dealerStateCode && dealerStateCode !== companyStateCode;

    // GST is only emitted when (a) the flag is on AND (b) the source
    // purchase carried GST (registered dealer). URD/agri purchases
    // produce exempt DNs regardless. The flag is the primary gate; the
    // dealer-carried-GST check still filters URD purchases out.
    const dealerCarriedGst = Number(p.cgst) || Number(p.sgst) || Number(p.igst);
    let cgst = 0, sgst = 0, igst = 0;
    if (flagDiscGst && dealerCarriedGst) {
      if (isInter) {
        igst = round2(discountAmt * dnGstRate / 100);
      } else {
        const half = round2(discountAmt * (dnGstRate / 2) / 100);
        cgst = half; sgst = half;
      }
    }
    const total = round2(discountAmt + cgst + sgst + igst);

    db.run(
      `INSERT INTO debit_notes (ano,date,state,name,note_no,amount,cgst,sgst,igst,total)
       VALUES (?,?,?,?,?,?,?,?,?,?)`,
      [ano, dnDate, p.state || '', dealerName,
       String(nextNoteNo), discountAmt, cgst, sgst, igst, total]
    );
    generated.push({ note_no: nextNoteNo, purchno: p.invo, dealer: dealerName, total });
    existingKeys.add(dealerName); // prevent same-loop duplicates
    nextNoteNo++;
  }

  res.json({
    success: true,
    created: generated.length,
    skipped: skipped.length,
    generated,
    skippedDetails: skipped,
    note: generated.length === 0 && skipped.length === 0
      ? `No eligible purchases in trade #${ano}`
      : undefined,
  });
});

// List purchases in a trade that don't yet have a DN. Used by the
// front-end preview in the Generate Debit Notes modal so the user can
// see exactly what will be created before they click Generate.
//
// Resolves `auctionId` (FK) → `ano` (trade number string) since DNs
// store ano, not the auction id.
app.get('/api/debit-notes/eligible-purchases/:auctionId', requireView, (req, res) => {
  const db = getDb();
  const auction = db.get('SELECT ano FROM auctions WHERE id = ?', [req.params.auctionId]);
  if (!auction) return res.status(404).json({ error: 'Auction not found' });

  const ano = auction.ano;
  // Anti-join: purchases in this trade where no DN exists for the same
  // (ano, dealer name) pair. The dealer name is the natural key here —
  // matches the idempotency rule used by /generate-bulk.
  //
  // Match purchases by the STABLE auction_id as well as `ano`, so a trade that
  // was renumbered after its purchases were created (leaving purchases.ano
  // stale) still surfaces them here instead of showing "no eligible purchases".
  // The anti-join stays on p.ano so a purchase and its own DN (which share the
  // same — possibly stale — number) still match correctly.
  const rows = db.all(
    `SELECT p.id, p.invo, p.name, p.amount, p.cgst, p.sgst, p.igst, p.total, p.date, p.state
       FROM purchases p
      WHERE (p.auction_id = ? OR p.ano = ?)
        AND p.amount > 0
        AND NOT EXISTS (
          SELECT 1 FROM debit_notes dn
           WHERE dn.ano = p.ano AND dn.name = p.name
        )
      ORDER BY p.id`,
    [Number(req.params.auctionId), ano]
  );
  res.json(rows);
});

// Return the next-available debit note number for a given trade.
// Numbering is TRADE-WISE: each trade has its own 1..N sequence.
// `ano` (trade number) is required — without it we can't know which
// trade's max to bump. The UI passes the currently-selected trade.
//
// Trade-wise reason: business reality + user expectation.
//   Trade 1 has DN #1..17 → starting trade 2 at #1 must work.
//   Earlier code did `MAX(note_no) FROM debit_notes` (global), so a
//   user entering #1 for trade 2 hit the "already in use" check
//   because trade 1 owned that number. The fix is to scope every
//   uniqueness check (single, range, MAX) to `WHERE ano = ?`.
app.get('/api/debit-notes/next-note-no', requireView, (req, res) => {
  const db = getDb();
  const ano = String(req.query.ano || '').trim();
  if (!ano) {
    return res.status(400).json({ error: 'ano (trade number) is required for trade-wise numbering' });
  }
  const row = db.get(
    'SELECT MAX(CAST(note_no AS INTEGER)) AS mx FROM debit_notes WHERE ano = ?',
    [ano]
  );
  const mx = parseInt(row && row.mx, 10);
  const next = Number.isFinite(mx) && mx > 0 ? mx + 1 : 1;
  res.json({ next, ano });
});

// /api/debit-notes/generate-all — DEPRECATED.
//
// This endpoint previously generated DNs across EVERY trade in a single
// sweep. The new spec (Generate Debit Note + Generate All) explicitly
// forbids cross-trade or global generation: every action must be scoped
// to one trade selected via the dn-auction dropdown.
//
// Hard-disabled with 410 Gone so any orphan client (older browser tab,
// scripted consumer) gets a clear migration message instead of silently
// creating cross-trade data.
app.post('/api/debit-notes/generate-all', requireInvoiceWrite, (req, res) => {
  res.status(410).json({
    error: 'Cross-trade DN generation is no longer supported. Use POST /api/debit-notes/generate-bulk with { ano } to generate DNs for a specific trade.',
  });
});

app.delete('/api/debit-notes/:id', requireDelete, (req, res) => {
  const db = getDb();
  // Cascade lock — match the invoice/purchase rules so all three
  // dependent docs respect the same "locked lot freezes the doc"
  // promise.
  if (!isAdmin(req) && lotsLockedForDebitNote(db, req.params.id)) {
    return res.status(423).json({ error: 'This debit note is locked because at least one of its lots is locked — only an admin can delete it.' });
  }
  db.run('DELETE FROM debit_notes WHERE id = ?', [req.params.id]);
  res.json({ success: true });
});

// Update debit note (edit)
app.put('/api/debit-notes/:id', requireInvoiceWrite, (req, res) => {
  const n = req.body;
  const db = getDb();
  if (!isAdmin(req) && lotsLockedForDebitNote(db, req.params.id)) {
    return res.status(423).json({ error: 'This debit note is locked because at least one of its lots is locked — only an admin can edit it.' });
  }
  const fields = ['ano','date','state','name','note_no','amount','cgst','sgst','igst','total'];
  const sets = []; const vals = [];
  for (const f of fields) {
    if (n[f] !== undefined) { sets.push(`${f}=?`); vals.push(n[f]); }
  }
  if (!sets.length) return res.status(400).json({ error: 'No fields to update' });
  vals.push(req.params.id);
  db.run(`UPDATE debit_notes SET ${sets.join(',')} WHERE id=?`, vals);
  res.json({ success: true });
});

// Render a Debit Note PDF for printing. Slim by design — DN is a single-
// row instrument (one note number, one amount block, one party). We keep
// the visual style consistent with the sales invoice (same header band,
// same value-table conventions) without pulling in the full multi-page
// invoice-pdf machinery.
app.get('/api/debit-notes/:id/pdf', requireView, (req, res) => {
  try {
    const PDFDocument = require('pdfkit');
    const db = getDb();
    const cfg = getSettingsFlat(db);
    const dn = db.get('SELECT * FROM debit_notes WHERE id = ?', [req.params.id]);
    if (!dn) return res.status(404).json({ error: 'Debit note not found' });

    // The DN layout (per the reference PDF) is BUYER-LETTERHEAD style:
    // the BUYER (the party benefiting from the discount, who issues the
    // credit/debit note in their books) prints on top with their address
    // + GSTIN, then "CREDIT NOTE/DEBIT NOTE" banner, then ISP company
    // appears as the recipient block in the middle.
    //
    // IMPORTANT: a Debit Note is a PURCHASE-side instrument issued by US
    // (the buyer) TO our supplier (the dealer). So `dn.name` stores the
    // DEALER name (the seller from `purchases.name`), not a buyer code.
    // This was the source of the qty/rate/value=0 bug: an earlier query
    // joined to `lots` by buyer1/buyer (matching nobody, since the DN
    // owner is on the SELLER side), so every row fell through to the
    // synthetic zero-value placeholder.
    const dealerName = String(dn.name || '').trim();
    // Look up the dealer's record (used for letterhead address/GSTIN).
    // Dealers live in `traders`, not `buyers` — buyers are the trade-
    // winners (our customers), traders are our suppliers. Match by
    // exact name with a UPPER() guard for case differences.
    const buyer = dealerName
      ? db.get(`SELECT * FROM traders WHERE UPPER(name) = UPPER(?) LIMIT 1`, [dealerName])
      : null;

    // Pull the per-lot line items from this trade for THIS dealer. The
    // DN's discount is allocated proportionally to puramt across the
    // dealer's lots (so the per-lot Discount column sums back to the
    // DN total). Match on `lots.name` (seller) — same column the
    // purchase invoice uses.
    const auction = db.get('SELECT * FROM auctions WHERE ano = ? LIMIT 1', [dn.ano]);
    let lots = [];
    if (auction) {
      lots = db.all(
        `SELECT lot_no, qty, prate, puramt, pqty
           FROM lots
          WHERE auction_id = ?
            AND UPPER(COALESCE(name,'')) = UPPER(?)
            AND amount > 0
          ORDER BY CAST(lot_no AS INTEGER), lot_no`,
        [auction.id, dealerName]
      );
    }

    // Distribute the DN amount across lots proportionally to puramt so
    // the per-lot Discount column sums back to the DN total. If lots
    // is empty (rare — orphan DN), we render a single synthetic row.
    const totalPuramt = lots.reduce((s, l) => s + Number(l.puramt || 0), 0);
    const dnAmount    = Number(dn.amount || 0);
    if (lots.length && totalPuramt > 0) {
      let allocated = 0;
      lots = lots.map((l, idx) => {
        const isLast = idx === lots.length - 1;
        const share = isLast
          ? Math.round((dnAmount - allocated) * 100) / 100
          : Math.round((dnAmount * Number(l.puramt) / totalPuramt) * 100) / 100;
        allocated += share;
        return { ...l, discount: share, taxable: share };
      });
    } else {
      lots = [{ lot_no: '—', qty: 0, prate: 0, puramt: 0, discount: dnAmount, taxable: dnAmount }];
    }

    // ── Indian number-to-words formatter (lakhs/crores style) ──
    // Mirrors the "Rupees X Only" convention used on tax invoices.
    const ones  = ['','One','Two','Three','Four','Five','Six','Seven','Eight','Nine','Ten',
                   'Eleven','Twelve','Thirteen','Fourteen','Fifteen','Sixteen','Seventeen','Eighteen','Nineteen'];
    const tens  = ['','','Twenty','Thirty','Forty','Fifty','Sixty','Seventy','Eighty','Ninety'];
    const tw = (n) => { // 0..99
      if (n < 20) return ones[n];
      const t = Math.floor(n / 10), o = n % 10;
      return tens[t] + (o ? ' ' + ones[o] : '');
    };
    const th = (n) => { // 0..999
      if (n === 0) return '';
      const h = Math.floor(n / 100), r = n % 100;
      return (h ? ones[h] + ' Hundred' + (r ? ' And ' : '') : '') + (r ? tw(r) : '');
    };
    const numToIndianWords = (num) => {
      const n = Math.abs(Math.round(Number(num) || 0));
      if (n === 0) return 'Zero';
      const crore = Math.floor(n / 10000000);
      const lakh  = Math.floor((n % 10000000) / 100000);
      const thou  = Math.floor((n % 100000) / 1000);
      const rest  = n % 1000;
      const parts = [];
      if (crore) parts.push(tw(crore) + ' Crore');
      if (lakh)  parts.push(tw(lakh) + ' Lakh');
      if (thou)  parts.push(tw(thou) + ' Thousand');
      if (rest)  parts.push(th(rest));
      return parts.join(' ');
    };

    const fmtAmt = (n) => Number(n || 0).toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    const fmtQty = (n) => Number(n || 0).toLocaleString('en-IN', { minimumFractionDigits: 3, maximumFractionDigits: 3 });
    // fmtDate falls through to the module-level helper which honours
    // the user's chosen date format from Settings → Display.

    // ── Render ────────────────────────────────────────────────
    const doc = new PDFDocument({ size: 'A4', margin: 30 });
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `inline; filename="DebitNote_${dn.note_no || dn.id}.pdf"`);
    doc.pipe(res);

    const PAGE_L = 30, PAGE_R = 565, PAGE_W = PAGE_R - PAGE_L; // ≈535pt usable

    // Top-right ORIGINAL/DUPLICATE/TRIPLICATE strip
    doc.font('Helvetica').fontSize(8).text('ORIGINAL/DUPLICATE/TRIPLICATE', PAGE_L, 32, { width: PAGE_W, align: 'right' });

    // Dealer letterhead (top): name centered bold, address + GSTIN.
    // Address resolution from traders table:
    //   padd  → place-of-business address line
    //   ppla  → city/place
    //   pstate + pin → state + pin code
    // Each piece is optional; we render only what's present, comma-joined.
    let y = 50;
    doc.font('Helvetica-Bold').fontSize(15).text(dealerName.toUpperCase(), PAGE_L, y, { width: PAGE_W, align: 'center' });
    y = doc.y + 2;
    doc.font('Helvetica').fontSize(9);
    const dealerAddrParts = buyer ? [
      buyer.padd, buyer.ppla, buyer.pin, buyer.pstate,
    ].filter(s => s && String(s).trim()) : [];
    const dealerAddr = dealerAddrParts.join(', ');
    if (dealerAddr) { doc.text(dealerAddr, PAGE_L, y, { width: PAGE_W, align: 'center' }); y = doc.y; }
    // Dealer GSTIN: traders.cr stores it (legacy "GSTIN.<15>" or bare 15-char).
    // Strip the 'GSTIN.' prefix for clean rendering.
    let dealerGstin = (buyer && buyer.cr) || '';
    if (/^GSTIN\.?/i.test(dealerGstin)) dealerGstin = dealerGstin.replace(/^GSTIN\.?/i, '');
    if (dealerGstin) { doc.font('Helvetica-Bold').text(`GSTIN: ${dealerGstin}`, PAGE_L, y, { width: PAGE_W, align: 'center' }); y = doc.y; }
    y += 6;

    // Outer frame starts here
    const FRAME_TOP = y;
    doc.lineWidth(0.7).moveTo(PAGE_L, FRAME_TOP).lineTo(PAGE_R, FRAME_TOP).stroke();
    y += 4;

    // CREDIT NOTE / DEBIT NOTE banner
    doc.font('Helvetica-Bold').fontSize(11).text('CREDIT NOTE/DEBIT NOTE', PAGE_L, y, { width: PAGE_W, align: 'center' });
    y = doc.y + 4;
    doc.moveTo(PAGE_L, y).lineTo(PAGE_R, y).stroke(); y += 6;

    // Recipient block: OUR company on left, Note No / Date on right.
    // Reads from the central company-identity resolver so the address
    // appears below the company name reliably (the previous code looked
    // for `cfg.address1` / `cfg.r_address1`, neither of which exist in
    // the e-Trade single-company schema — both were always blank,
    // dropping the address line entirely).
    const _ident = getCompanyIdentity(cfg);
    const company       = _ident.name || cfg.short_name || '';
    const ispGstin      = _ident.gstin;
    const ispPan        = _ident.pan;
    const ispState      = _ident.state;
    const ispStateCode  = _ident.stateCode || (ispState === 'KERALA' ? '32' : '33');
    // Company address shown directly below the company name. Pull the
    // resolved identity first; fall back to raw cfg fields so the address
    // still renders when getCompanyIdentity returns blanks for either line.
    const ispAddrLine1  = _ident.address1
      || cfg.address1 || cfg.r_address1 || cfg.tn_address1 || cfg.kl_address1 || '';
    const ispAddrLine2  = _ident.address2
      || cfg.address2 || cfg.r_address2 || cfg.tn_address2 || cfg.kl_address2 || '';

    const noteRefSuffix = (cfg.season_short || cfg.tally_season || '26-27').replace(/[^0-9-]/g,'');

    doc.font('Helvetica-Bold').fontSize(10).text(company.toUpperCase(), PAGE_L + 4, y);
    doc.font('Helvetica-Bold').fontSize(10).text(`No.: ${dn.note_no || ''}/${noteRefSuffix}`, PAGE_R - 200, y, { width: 196, align: 'right' });
    y = doc.y + 2;
    // Company address sits directly under the company name. Always render
    // line 1 (uses a placeholder when blank so the layout stays predictable).
    doc.font('Helvetica').fontSize(9).text(ispAddrLine1 || '', PAGE_L + 4, y, { width: PAGE_R - 200 - (PAGE_L + 4) });
    doc.font('Helvetica-Bold').fontSize(9).text(`Date :${fmtDate(dn.date)}`, PAGE_R - 200, y, { width: 196, align: 'right' });
    y = doc.y + 2;
    if (ispAddrLine2) { doc.font('Helvetica').fontSize(9).text(ispAddrLine2, PAGE_L + 4, y); y = doc.y; }
    if (ispGstin) { doc.font('Helvetica').fontSize(9).text(`GSTIN : ${ispGstin}`, PAGE_L + 4, y); y = doc.y; }
    if (ispPan)   { doc.font('Helvetica').fontSize(9).text(`PAN   : ${ispPan}`,   PAGE_L + 4, y); y = doc.y; }
    if (ispState) { doc.font('Helvetica').fontSize(9).text(`STATE : ${ispState}    CODE:${ispStateCode}`, PAGE_L + 4, y); y = doc.y; }
    y += 4;
    doc.moveTo(PAGE_L, y).lineTo(PAGE_R, y).stroke(); y += 6;

    // "Discount on Sale of Cardamom HSN CODE:..." title
    const hsnCardamom = cfg.tally_hsn_cardamom || cfg.hsn_cardamom || '09083120';
    doc.font('Helvetica-Bold').fontSize(10).text(`Discount on Sale of Cardamom    HSN CODE:${hsnCardamom}`, PAGE_L + 4, y);
    y = doc.y + 4;
    doc.moveTo(PAGE_L, y).lineTo(PAGE_R, y).stroke(); y += 4;

    // ── Line items table ──
    // Column widths (sum ≈ 535pt usable). Match the reference layout:
    // Column layout:
    //   Sl | Lot | Quantity (kgs) | Rate/kg (Rs) | Value (Rs) | Discount (Rs) | [GST (Rs)?] | Taxable (Rs)
    //
    // The 7th column was previously a 40pt empty visual gap. It now shows
    // GST on the discount (taken from cfg.discount_gst, e.g. 18%) WHEN the
    // "Discount includes GST" feature flag (`flag_disc_gst`) is enabled.
    // When the flag is off, we DROP the column entirely (not just blank it)
    // so the surrounding columns redistribute the freed width — no visible
    // empty band on the PDF. The Taxable Value column also widens to keep
    // the table flush to PAGE_W.
    // Robust flag read — handles both real boolean and string-stored
    // forms. `!!cfg.flag_disc_gst` would treat the string 'false' as
    // truthy; this matches the same pattern used by the Payments-tab
    // live derivation at calculations.js so the PDF and the Payments
    // view always agree.
    const showGstCol = String(cfg.flag_disc_gst || '').toLowerCase() === 'true' || cfg.flag_disc_gst === true;
    const gstHeaderRate = Number(cfg.discount_gst) || Number(cfg.gst_service) || 18;
    const cols = [
      { key:'sl',       w: 32,  label1:'Sl',       label2:'No',     align:'center' },
      { key:'lot',      w: 38,  label1:'Lot',      label2:'No',     align:'center' },
      { key:'qty',      w: 76,  label1:'Quantity', label2:'(kgs)',  align:'right'  },
      { key:'rate',     w: 70,  label1:'Rate/kg',  label2:'(Rs)',   align:'right'  },
      { key:'value',    w: 100, label1:'Value',    label2:'(Rs)',   align:'right'  },
      { key:'discount', w: 80,  label1:'Discount', label2:'(Rs)',   align:'right'  },
    ];
    if (showGstCol) {
      // GST column header shows the active rate (e.g. "GST 18%") so the
      // user can immediately see what discount_gst is set to without
      // cross-checking Settings.
      cols.push({ key:'gst', w: 70,  label1:`GST ${gstHeaderRate}%`, label2:'(Rs)', align:'right' });
      cols.push({ key:'taxable', w: 69, label1:'Taxable', label2:'Value', align:'right' });
    } else {
      // Flag off — taxable column absorbs the entire 109pt slot.
      cols.push({ key:'taxable', w: 139, label1:'Taxable', label2:'Value', align:'right' });
    }
    // Compute x positions
    let xs = [PAGE_L];
    cols.forEach(c => xs.push(xs[xs.length - 1] + c.w));

    // Header rows
    const HEAD_H = 26;
    const headTop = y;
    // Vertical lines (header)
    xs.forEach(x => doc.moveTo(x, headTop).lineTo(x, headTop + HEAD_H).stroke());
    // Top + bottom of header
    doc.moveTo(PAGE_L, headTop).lineTo(PAGE_R, headTop).stroke();
    doc.moveTo(PAGE_L, headTop + HEAD_H).lineTo(PAGE_R, headTop + HEAD_H).stroke();
    doc.font('Helvetica-Bold').fontSize(9);
    cols.forEach((c, i) => {
      doc.text(c.label1, xs[i] + 3, headTop + 4,  { width: c.w - 6, align: c.align });
      doc.text(c.label2, xs[i] + 3, headTop + 14, { width: c.w - 6, align: c.align });
    });
    y = headTop + HEAD_H;

    // Resolve column indices by key — the layout is conditional on the
    // GST column flag, so we can't hardcode positions. `colIdx` returns
    // the array index of the named column or -1 if not present.
    const colIdx = (key) => cols.findIndex(c => c.key === key);
    const taxIdx = colIdx('taxable');
    const gstIdx = colIdx('gst');  // -1 when flag is off

    // Data rows
    const ROW_H = 14;
    const MAX_ROWS = 14; // matches reference pdf — page can hold ~14 line rows comfortably
    doc.font('Helvetica').fontSize(9);
    let totalQty = 0, totalValue = 0, totalDiscount = 0, totalTaxable = 0, totalGst = 0;
    // Per-row GST = discount × discount_gst%. discount_gst is the rate
    // (e.g. 18). Computed per-lot so the column sums back to the DN's
    // total GST when the flag is on.
    const gstRateFraction = gstHeaderRate / 100;
    for (let i = 0; i < Math.max(lots.length, MAX_ROWS); i++) {
      const lot = lots[i];
      // Vertical lines for this row
      xs.forEach(x => doc.moveTo(x, y).lineTo(x, y + ROW_H).stroke());
      if (lot) {
        const value    = Number(lot.qty || 0) * Number(lot.prate || 0);
        const discount = Number(lot.discount || 0);
        // GST on discount — only meaningful when the column is shown.
        const gstOnDiscount = showGstCol
          ? round2(discount * gstRateFraction)
          : 0;
        // Taxable Value historically equals the discount; with the GST
        // column visible we keep the same semantic so the existing
        // GRAND TOTAL math is unchanged.
        const taxable  = Number(lot.taxable || discount);
        totalQty      += Number(lot.qty || 0);
        totalValue    += value;
        totalDiscount += discount;
        totalGst      += gstOnDiscount;
        totalTaxable  += taxable;
        doc.text(String(i + 1),                 xs[0] + 3, y + 3, { width: cols[0].w - 6, align: cols[0].align });
        doc.text(String(lot.lot_no || ''),      xs[1] + 3, y + 3, { width: cols[1].w - 6, align: cols[1].align });
        doc.text(fmtQty(lot.qty),               xs[2] + 3, y + 3, { width: cols[2].w - 6, align: cols[2].align });
        doc.text(fmtAmt(lot.prate),             xs[3] + 3, y + 3, { width: cols[3].w - 6, align: cols[3].align });
        doc.text(fmtAmt(value),                 xs[4] + 3, y + 3, { width: cols[4].w - 6, align: cols[4].align });
        doc.text(fmtAmt(discount),              xs[5] + 3, y + 3, { width: cols[5].w - 6, align: cols[5].align });
        if (gstIdx >= 0) {
          doc.text(fmtAmt(gstOnDiscount),       xs[gstIdx] + 3, y + 3, { width: cols[gstIdx].w - 6, align: cols[gstIdx].align });
        }
        doc.text(fmtAmt(taxable),               xs[taxIdx] + 3, y + 3, { width: cols[taxIdx].w - 6, align: cols[taxIdx].align });
      }
      y += ROW_H;
    }
    // Bottom border of data area
    doc.moveTo(PAGE_L, y).lineTo(PAGE_R, y).stroke();

    // TOTAL row
    const TOT_H = 16;
    doc.font('Helvetica-Bold').fontSize(9);
    // Merged "TOTAL" cell across Sl + Lot
    doc.moveTo(PAGE_L, y).lineTo(PAGE_L, y + TOT_H).stroke();
    doc.moveTo(xs[2], y).lineTo(xs[2], y + TOT_H).stroke();
    doc.text('TOTAL', PAGE_L + 3, y + 4, { width: cols[0].w + cols[1].w - 6, align: 'center' });
    doc.text(fmtQty(totalQty), xs[2] + 3, y + 4, { width: cols[2].w - 6, align: 'right' });
    doc.moveTo(xs[3], y).lineTo(xs[3], y + TOT_H).stroke();
    // Skip Rate column on totals row (no average)
    doc.moveTo(xs[4], y).lineTo(xs[4], y + TOT_H).stroke();
    doc.text(fmtAmt(totalValue), xs[4] + 3, y + 4, { width: cols[4].w - 6, align: 'right' });
    doc.moveTo(xs[5], y).lineTo(xs[5], y + TOT_H).stroke();
    doc.text(fmtAmt(totalDiscount), xs[5] + 3, y + 4, { width: cols[5].w - 6, align: 'right' });
    if (gstIdx >= 0) {
      doc.moveTo(xs[gstIdx], y).lineTo(xs[gstIdx], y + TOT_H).stroke();
      doc.text(fmtAmt(totalGst), xs[gstIdx] + 3, y + 4, { width: cols[gstIdx].w - 6, align: 'right' });
    }
    doc.moveTo(xs[taxIdx], y).lineTo(xs[taxIdx], y + TOT_H).stroke();
    doc.moveTo(PAGE_R, y).lineTo(PAGE_R, y + TOT_H).stroke();
    doc.text(fmtAmt(totalTaxable), xs[taxIdx] + 3, y + 4, { width: cols[taxIdx].w - 6, align: 'right' });
    y += TOT_H;
    doc.moveTo(PAGE_L, y).lineTo(PAGE_R, y).stroke();

    // ── Tax calculation rows (mirrors purchase invoice) ──
    // Each row spans the same right-hand pair of columns as GRAND TOTAL:
    //   [discount column .. taxable column] holds the label + value.
    // Rows render only for non-zero figures so an intra-state DN doesn't
    // print an empty IGST line, and vice versa.
    const TAX_H = 14;
    const labelStartIdx = colIdx('discount');                 // left edge of label cell
    const valueIdx      = taxIdx;                              // taxable column = value
    const labelStartX   = xs[labelStartIdx];
    const labelW        = (xs[valueIdx] - xs[labelStartIdx]) - 6;
    const valueX        = xs[valueIdx] + 3;
    const valueW        = cols[valueIdx].w - 6;
    const taxableValue  = Number(dn.amount || totalTaxable) || 0;
    const cgstAmt       = Number(dn.cgst) || 0;
    const sgstAmt       = Number(dn.sgst) || 0;
    const igstAmt       = Number(dn.igst) || 0;
    const cgstRate      = Number(cfg.gst_cgst) || 2.5;
    const sgstRate      = Number(cfg.gst_sgst) || 2.5;
    const igstRate      = Number(cfg.gst_igst) || 5;
    const taxRows = [
      { label: 'TAXABLE VALUE',           amt: taxableValue, always: true },
      { label: `CGST  ${cgstRate}%`,      amt: cgstAmt },
      { label: `SGST  ${sgstRate}%`,      amt: sgstAmt },
      { label: `IGST  ${igstRate}%`,      amt: igstAmt },
    ].filter(r => r.always || r.amt > 0);
    doc.font('Helvetica').fontSize(9);
    for (const r of taxRows) {
      doc.moveTo(PAGE_L, y).lineTo(PAGE_L, y + TAX_H).stroke();
      doc.moveTo(labelStartX, y).lineTo(labelStartX, y + TAX_H).stroke();
      doc.moveTo(xs[valueIdx], y).lineTo(xs[valueIdx], y + TAX_H).stroke();
      doc.moveTo(PAGE_R, y).lineTo(PAGE_R, y + TAX_H).stroke();
      doc.text(r.label, labelStartX + 3, y + 3, { width: labelW, align: 'right' });
      doc.text(fmtAmt(r.amt), valueX, y + 3, { width: valueW, align: 'right' });
      y += TAX_H;
      doc.moveTo(PAGE_L, y).lineTo(PAGE_R, y).stroke();
    }

    // GRAND TOTAL row
    const GT_H = 18;
    doc.moveTo(PAGE_L, y).lineTo(PAGE_L, y + GT_H).stroke();
    doc.moveTo(xs[5], y).lineTo(xs[5], y + GT_H).stroke();
    doc.moveTo(xs[7], y).lineTo(xs[7], y + GT_H).stroke();
    doc.moveTo(PAGE_R, y).lineTo(PAGE_R, y + GT_H).stroke();
    doc.font('Helvetica-Bold').fontSize(10).text('GRAND TOTAL', xs[5] + 3, y + 5, { width: cols[5].w + cols[6].w - 6, align: 'right' });
    doc.text(fmtAmt(dn.total || (taxableValue + cgstAmt + sgstAmt + igstAmt) || totalTaxable), xs[7] + 3, y + 5, { width: cols[7].w - 6, align: 'right' });
    y += GT_H;
    doc.moveTo(PAGE_L, y).lineTo(PAGE_R, y).stroke();
    y += 14;

    // Amount in words
    const grandTotal = Math.round(Number(dn.total || totalTaxable) || 0);
    doc.font('Helvetica').fontSize(10).text(
      `Rupees ${numToIndianWords(grandTotal)} Only`,
      PAGE_L + 4, y,
      { width: PAGE_W - 8 }
    );
    y = doc.y + 14;

    // "For DEALER" + Authorised Signatory (right-aligned). The DN is
    // signed off by the dealer (the party crediting the discount).
    doc.font('Helvetica-Bold').fontSize(10).text(`For ${dealerName.toUpperCase()}`, PAGE_L, y, { width: PAGE_W - 8, align: 'right' });
    y += 50;
    doc.font('Helvetica').fontSize(9).text('Authorised Signatory', PAGE_L, y, { width: PAGE_W - 8, align: 'right' });
    y += 14;

    // Bottom frame line
    doc.moveTo(PAGE_L, y).lineTo(PAGE_R, y).stroke();
    // Left + right outer frame from FRAME_TOP to here
    doc.moveTo(PAGE_L, FRAME_TOP).lineTo(PAGE_L, y).stroke();
    doc.moveTo(PAGE_R, FRAME_TOP).lineTo(PAGE_R, y).stroke();

    doc.end();
  } catch (e) {
    console.error('[dn-pdf] failed:', e);
    if (!res.headersSent) res.status(500).json({ error: e.message });
  }
});

// ══════════════════════════════════════════════════════════════
// JOURNALS (JOUR.PRG, PUJOUR.PRG, PPUJOUR.PRG)
// ══════════════════════════════════════════════════════════════
// e-Trade-only build: journals filter by trade (auction_id), not by
// date range. Dates rendered dd/mm/yyyy by the calculations layer.
app.get('/api/journals/sales', requireView, (req, res) => {
  const { auctionId, saleType } = req.query;
  if (!auctionId) return res.status(400).json({ error: 'auctionId required' });
  res.json(getSalesJournal(getDb(), auctionId, saleType));
});

app.get('/api/journals/purchase', requireView, (req, res) => {
  const { auctionId, type } = req.query;
  if (!auctionId) return res.status(400).json({ error: 'auctionId required' });
  res.json(getPurchaseJournal(getDb(), auctionId, type || 'dealer'));
});

// Journal exports (XLSX only). Trade-based — dates dd/mm/yyyy.
app.get('/api/exports/sales-journal', requireExport, async (req, res) => {
  const { auctionId, saleType } = req.query;
  if (!auctionId) return res.status(400).json({ error: 'auctionId required' });
  const { exportSalesJournal } = require('./exports');
  const buffer = await exportSalesJournal(getDb(), auctionId, saleType);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename="SalesJournal.xlsx"');
  res.send(Buffer.from(buffer));
});

app.get('/api/exports/purchase-journal', requireExport, async (req, res) => {
  const { auctionId, type } = req.query;
  if (!auctionId) return res.status(400).json({ error: 'auctionId required' });
  const baseName = type === 'agri' ? 'AgriBillJournal' : 'PurchaseJournal';
  const { exportPurchaseJournal } = require('./exports');
  const buffer = await exportPurchaseJournal(getDb(), auctionId, type || 'dealer');
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename="${baseName}.xlsx"`);
  res.send(Buffer.from(buffer));
});

// ══════════════════════════════════════════════════════════════
// REGISTERS — lot-wise Purchase / invoice-wise Sales
// Sibling of Journals. Scope: a specific trade (auctionId) OR a date
// range across trades (from/to). Sales also takes an optional saleType.
// ══════════════════════════════════════════════════════════════
function registerOpts(req, cfg) {
  return {
    auctionId: req.query.auctionId || null,
    from: req.query.from || null,
    to: req.query.to || null,
    saleType: req.query.saleType || null,
    mode: (cfg && cfg.business_mode) || 'e-Trade',
  };
}

// Resolve an auction's human-facing trade number (ano) for use in export
// filenames. Downloads must be named by the trade no the user knows, NOT
// the internal auctions.id. Falls back to the raw id only if the auction
// can't be found, and sanitises for filesystem safety.
function anoForFile(db, auctionId) {
  if (auctionId == null || auctionId === '') return '';
  let val = String(auctionId);
  try {
    const a = db.get('SELECT ano FROM auctions WHERE id = ?', [auctionId]);
    if (a && a.ano != null && String(a.ano).trim() !== '') val = String(a.ano);
  } catch (e) { /* fall back to id */ }
  return val.replace(/[^\w.-]+/g, '_');
}

app.get('/api/registers/purchase', requireView, (req, res) => {
  const { getPurchaseRegister } = require('./calculations');
  const db = getDb();
  res.json(getPurchaseRegister(db, registerOpts(req, getSettingsFlat(db))));
});

app.get('/api/registers/sales', requireView, (req, res) => {
  const { getSalesRegister } = require('./calculations');
  res.json(getSalesRegister(getDb(), registerOpts(req, {})));
});

app.get('/api/exports/purchase-register', requireExport, async (req, res) => {
  try {
    const db = getDb();
    const cfg = getSettingsFlat(db);
    const opts = registerOpts(req, cfg);
    const format = String(req.query.format || 'xlsx').toLowerCase();
    if (format === 'pdf') {
      const buffer = await exportAnyPdf(db, 'purchase_register', opts.auctionId, cfg, { from: opts.from, to: opts.to });
      res.setHeader('Content-Type', 'application/pdf');
      res.setHeader('Content-Disposition', 'attachment; filename="PurchaseRegister.pdf"');
      return res.send(buffer);
    }
    const { exportPurchaseRegister } = require('./exports');
    const buffer = await exportPurchaseRegister(db, opts);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="PurchaseRegister.xlsx"');
    res.send(Buffer.from(buffer));
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.get('/api/exports/sales-register', requireExport, async (req, res) => {
  try {
    const db = getDb();
    const cfg = getSettingsFlat(db);
    const opts = registerOpts(req, cfg);
    const format = String(req.query.format || 'xlsx').toLowerCase();
    if (format === 'pdf') {
      const buffer = await exportAnyPdf(db, 'sales_register', opts.auctionId, cfg, { from: opts.from, to: opts.to, saleType: opts.saleType });
      res.setHeader('Content-Type', 'application/pdf');
      res.setHeader('Content-Disposition', 'attachment; filename="SalesRegister.pdf"');
      return res.send(buffer);
    }
    const { exportSalesRegister } = require('./exports');
    const buffer = await exportSalesRegister(db, opts);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="SalesRegister.xlsx"');
    res.send(Buffer.from(buffer));
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ── Per-party "Individual" registers (cross-auction, date-range) ──
// Pooler / Seller / Merchant party statements. Scoped by a From/To date
// range and an OPTIONAL single party (empty = every party, one section
// per party in the export).
const INDIVIDUAL_REG_KINDS = { pooler: 1, seller: 1, merchant: 1 };
function individualRegOpts(req) {
  return {
    from: req.query.from || null,
    to: req.query.to || null,
    party: req.query.party || null,
  };
}

app.get('/api/registers/individual', requireView, (req, res) => {
  try {
    const kind = String(req.query.kind || '').toLowerCase();
    if (!INDIVIDUAL_REG_KINDS[kind]) return res.status(400).json({ error: 'Unknown register kind' });
    const db = getDb();
    const { getPoolerRegister, getSellerRegister, getMerchantRegister } = require('./calculations');
    const fn = kind === 'seller' ? getSellerRegister
             : kind === 'merchant' ? getMerchantRegister : getPoolerRegister;
    res.json(fn(db, individualRegOpts(req)));
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.get('/api/registers/individual-parties', requireView, (req, res) => {
  try {
    const db = getDb();
    const { listRegisterParties } = require('./calculations');
    res.json(listRegisterParties(db, {
      kind: String(req.query.kind || '').toLowerCase(),
      from: req.query.from || null, to: req.query.to || null,
    }));
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.get('/api/exports/individual-register', requireExport, async (req, res) => {
  try {
    const kind = String(req.query.kind || '').toLowerCase();
    if (!INDIVIDUAL_REG_KINDS[kind]) return res.status(400).json({ error: 'Unknown register kind' });
    const db = getDb();
    const cfg = getSettingsFlat(db);
    const opts = individualRegOpts(req);
    const format = String(req.query.format || 'xlsx').toLowerCase();
    const fileBase = { pooler: 'PoolerRegister', seller: 'SellerRegister', merchant: 'MerchantRegister' }[kind];
    if (format === 'pdf') {
      const pdfType = { pooler: 'pooler_individual', seller: 'seller_individual', merchant: 'merchant_individual' }[kind];
      const buffer = await exportAnyPdf(db, pdfType, null, cfg, opts);
      res.setHeader('Content-Type', 'application/pdf');
      res.setHeader('Content-Disposition', `attachment; filename="${fileBase}.pdf"`);
      return res.send(buffer);
    }
    const { exportIndividualRegister } = require('./exports');
    const buffer = await exportIndividualRegister(db, kind, opts);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${fileBase}.xlsx"`);
    res.send(Buffer.from(buffer));
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ══════════════════════════════════════════════════════════════
// INVOICE PREVIEW (PREINVO.PRG) — dry-run, no save
// ══════════════════════════════════════════════════════════════
app.post('/api/invoices/preview/:auctionId', requireView, (req, res) => {
  const db = getDb(); const cfg = getSettingsFlat(db);
  const { saleType, buyerCode, type } = req.body;
  // Per-invoice "No Transport & Insurance" — preview reflects the toggle.
  const noTI = (req.body.noTI === true || String(req.body.noTI || '').toLowerCase() === 'true' || Number(req.body.noTI) === 1) ? 1 : 0;

  // Auto-calculate any uncalculated lots first (read-only would be better but we need the data)
  const uncalc = db.all(`SELECT * FROM lots WHERE auction_id = ? AND amount > 0 AND (puramt IS NULL OR puramt = 0)`, [req.params.auctionId]);
  for (const lot of uncalc) {
    const c = calculateLot(lot, cfg);
    db.run(`UPDATE lots SET pqty=?,prate=?,puramt=?,com=?,sertax=?,cgst=?,sgst=?,igst=?,advance=?,balance=?,bilamt=?,refund=?,refud=?,isp_pqty=?,isp_prate=?,isp_puramt=?,asp_pqty=?,asp_prate=?,asp_puramt=? WHERE id=?`,
      [c.pqty,c.prate,c.puramt,c.com,c.sertax,c.cgst,c.sgst,c.igst,c.advance,c.balance,c.bilamt,c.refund||0,c.refud||0,c.isp_pqty||0,c.isp_prate||0,c.isp_puramt||0,c.asp_pqty||0,c.asp_prate||0,c.asp_puramt||0,lot.id]);
  }
  
  let invoice;
  if (type === 'purchase') {
    invoice = buildPurchaseInvoice(db, req.params.auctionId, buyerCode, cfg); // buyerCode = sellerName for purchase
  } else if (type === 'agri') {
    invoice = buildAgriBill(db, req.params.auctionId, buyerCode, cfg);
    if (invoice && invoice.error) return res.status(404).json({ error: invoice.error });
  } else {
    invoice = buildSalesInvoice(db, req.params.auctionId, buyerCode, saleType, cfg, { noTI });
  }

  if (!invoice) return res.status(404).json({ error: 'No data found' });
  res.json({ preview: true, invoice });
});

// ══════════════════════════════════════════════════════════════
// PAYMENTS (PAYCHECK.PRG)
// ══════════════════════════════════════════════════════════════
app.get('/api/payments/:auctionId', requireView, (req, res) => {
  const db = getDb();
  const cfg = getSettingsFlat(db);
  const summary = getPaymentSummary(db, req.params.auctionId, req.query.state, cfg);
  res.json(summary);
});

// Delete the lots + DNs for a list of sellers in one auction. Powers
// the "Delete Selected" button on the Payments tab. Each row in the
// payments table is a roll-up of one seller's lots in the trade, so
// deleting a payment "row" means clearing those underlying lots.
//
// Body: { sellerNames: ['DealerA', 'DealerB', ...] }
//
// Trade-scoped — only touches data in this auction. Names are matched
// case-insensitively to be tolerant of legacy data inconsistencies.
app.post('/api/payments/:auctionId/delete-sellers', requireDelete, (req, res) => {
  try {
    const db = getDb();
    const auctionId = req.params.auctionId;
    const names = Array.isArray(req.body.sellerNames) ? req.body.sellerNames : [];
    if (!names.length) {
      return res.status(400).json({ error: 'sellerNames array is required' });
    }
    const auction = db.get('SELECT ano FROM auctions WHERE id = ?', [auctionId]);
    if (!auction) return res.status(404).json({ error: 'Auction not found' });

    // Build a parameterised IN clause. SQLite max parameters is 999;
    // we cap chunks at 500 names per query to stay well within limits.
    const CHUNK = 500;
    let lotsDeleted = 0, dnsDeleted = 0;
    for (let i = 0; i < names.length; i += CHUNK) {
      const batch = names.slice(i, i + CHUNK).map(s => String(s).trim()).filter(Boolean);
      if (!batch.length) continue;
      const placeholders = batch.map(() => '?').join(',');
      // Lots — case-insensitive match. UPPER() on both sides handles
      // legacy case mismatches (data sometimes imported as title case,
      // sometimes uppercased).
      const upperBatch = batch.map(n => n.toUpperCase());
      const lotsBefore = db.get(
        `SELECT COUNT(*) AS c FROM lots
          WHERE auction_id = ?
            AND UPPER(COALESCE(name,'')) IN (${placeholders})`,
        [auctionId, ...upperBatch]
      ).c;
      db.run(
        `DELETE FROM lots
          WHERE auction_id = ?
            AND UPPER(COALESCE(name,'')) IN (${placeholders})`,
        [auctionId, ...upperBatch]
      );
      lotsDeleted += lotsBefore;

      // Cascading DN cleanup — debit_notes for these sellers in this
      // trade are now orphaned references, so delete them too.
      const dnsBefore = db.get(
        `SELECT COUNT(*) AS c FROM debit_notes
          WHERE ano = ?
            AND UPPER(COALESCE(name,'')) IN (${placeholders})`,
        [auction.ano, ...upperBatch]
      ).c;
      db.run(
        `DELETE FROM debit_notes
          WHERE ano = ?
            AND UPPER(COALESCE(name,'')) IN (${placeholders})`,
        [auction.ano, ...upperBatch]
      );
      dnsDeleted += dnsBefore;
    }

    res.json({
      success: true,
      sellers: names.length,
      lotsDeleted,
      dnsDeleted,
    });
  } catch (e) {
    res.status(500).json({ error: 'Delete failed: ' + (e.message || e) });
  }
});

// ── Bank payment data (BANKPAY.PRG) ──────────────────────────
app.get('/api/payments/bank/:auctionId', requireView, (req, res) => {
  const cfg = getSettingsFlat(getDb());
  const data = getBankPaymentData(getDb(), req.params.auctionId, cfg);
  res.json(data);
});

// ── Payment Statement PDF (per-seller) ────────────────────────
// Lightweight A4 PDF showing the payment due to one seller for one
// auction: header, seller block, lots breakdown, totals. Powers both
// "Print" and "WhatsApp" actions on the Payments tab.
function _renderPaymentStatement(doc, db, auctionId, sellerName, cfg, lotIds) {
  const auction = db.get('SELECT * FROM auctions WHERE id = ?', [auctionId]) || { ano:'', date:'' };
  // The lots schema has no `rate` column — alias `prate` (per-kg purchase
  // rate post-calculation) as `rate` for the cols->key mapping below.
  // The previous query referenced a non-existent column, so the SELECT
  // threw AFTER doc.pipe(res) had begun, leaving the client with an empty
  // PDF (the catch block could only call doc.end() at that point).
  // Match seller by trimmed/case-insensitive name so legacy rows whose
  // `name` was stored with trailing whitespace or mixed case still pair
  // up — same fallback the trader lookup below already uses.
  //
  // Optional `lotIds` narrows the statement to a caller-chosen subset of
  // the seller's lots — supports the Payments tab's "partial payment"
  // flow where the operator picks specific lots to settle now and leaves
  // the rest for a later visit.
  const lotIdFilter = Array.isArray(lotIds)
    ? lotIds.map(n => parseInt(n, 10)).filter(Number.isFinite)
    : [];
  let lotSql = `SELECT id, lot_no, pqty, prate AS rate, puramt, refund, balance, cgst, sgst, igst
       FROM lots
      WHERE auction_id = ?
        AND TRIM(LOWER(COALESCE(name,''))) = TRIM(LOWER(?))
        AND amount > 0`;
  const lotParams = [auctionId, sellerName];
  if (lotIdFilter.length) {
    const placeholders = lotIdFilter.map(() => '?').join(',');
    lotSql += ` AND id IN (${placeholders})`;
    lotParams.push(...lotIdFilter);
  }
  lotSql += ' ORDER BY CAST(lot_no AS INTEGER), lot_no';
  const lots = db.all(lotSql, lotParams) || [];
  const trader = db.get('SELECT * FROM traders WHERE LOWER(name) = LOWER(?) LIMIT 1', [sellerName]);
  // GST and TDS are seller-level figures — pull them from the SAME source as
  // the Payments main list (getPaymentSummary) so the statement matches it
  // exactly: GST is computed LIVE from the discount (the per-lot cgst/sgst/igst
  // stamps are often 0/stale, which made the statement show GST 0 while the
  // list showed a value); TDS comes from the seller's purchase invoice(s).
  // Both are seller-level, so we spread them across the seller's lots in
  // proportion to puramt — using the seller's FULL puramt as the denominator
  // so a partial statement (lotIds subset) shows a proportionate share, and
  // the columns total back to the seller's main-list figures.
  let sellerGst = 0, sellerTds = 0;
  try {
    const ps = getPaymentSummary(db, auctionId, null, cfg);
    const psRow = (ps || []).find(r => String(r.name || '').trim().toLowerCase() === String(sellerName || '').trim().toLowerCase());
    if (psRow) { sellerGst = Number(psRow.total_tax) || 0; sellerTds = Number(psRow.total_tds) || 0; }
  } catch (_) { /* non-fatal — columns just read 0 */ }
  const fullPuramtRow = db.get(
    `SELECT SUM(COALESCE(puramt,0)) AS p FROM lots
      WHERE auction_id = ? AND TRIM(LOWER(COALESCE(name,''))) = TRIM(LOWER(?)) AND amount > 0`,
    [auctionId, sellerName]
  );
  const fullPuramt = fullPuramtRow ? (Number(fullPuramtRow.p) || 0) : 0;
  const gstRate = fullPuramt > 0 ? sellerGst / fullPuramt : 0;
  const tdsRate = fullPuramt > 0 ? sellerTds / fullPuramt : 0;
  const fmtAmt = (n) => Number(n || 0).toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const fmtQty = (n) => Number(n || 0).toLocaleString('en-IN', { minimumFractionDigits: 3, maximumFractionDigits: 3 });
  // fmtDate falls through to the module-level helper which honours
  // the user's chosen date format from Settings → Display.
  const company = (cfg.trade_name || cfg.short_name || cfg.tally_company_name || cfg.legal_name || 'Company').toString();

  const PAGE_L = 30, PAGE_R = 565, PAGE_W = PAGE_R - PAGE_L;
  let y = 40;
  doc.font('Helvetica-Bold').fontSize(16).text(company.toUpperCase(), PAGE_L, y, { width: PAGE_W, align: 'center' });
  y = doc.y + 4;
  doc.font('Helvetica-Bold').fontSize(13).text('PAYMENT STATEMENT', PAGE_L, y, { width: PAGE_W, align: 'center' });
  y = doc.y + 10;
  doc.moveTo(PAGE_L, y).lineTo(PAGE_R, y).lineWidth(1).stroke();
  y += 10;

  doc.font('Helvetica').fontSize(10);
  // Trade / Date sit right-aligned against the right margin (matching the
  // table edge below); the left-hand Seller / Phone text is width-capped to
  // stop short of that block so a long seller name can't overrun it. Single
  // line + ellipsis is a final guard against an extreme name.
  const RIGHT_W = 130, RIGHT_X = PAGE_R - RIGHT_W, LEFT_W = RIGHT_X - PAGE_L - 12;
  doc.text(`Seller: ${sellerName}`, PAGE_L, y, { width: LEFT_W, lineBreak: false, ellipsis: true });
  doc.text(`Trade: ${auction.ano}`, RIGHT_X, y, { width: RIGHT_W, align: 'right' });
  y += 14;
  doc.text(`Phone: ${trader && trader.tel ? trader.tel : '-'}`, PAGE_L, y, { width: LEFT_W, lineBreak: false, ellipsis: true });
  doc.text(`Date: ${fmtDate(auction.date)}`, RIGHT_X, y, { width: RIGHT_W, align: 'right' });
  y += 18;

  // Payment is the planter/purchase side: Qty/Rate/Amount source from
  // pqty / prate / puramt (NOT auction qty/price/amount) so this matches
  // the Payments screen + the per-seller Lots modal.
  // Re-laid out to fit a TDS column between GST and Payable; widths still
  // sum to PAGE_W (535) so the header rule and the right margin line up.
  const cols = [
    { k: 'lot_no', label: 'Lot#',     x: PAGE_L,        w: 50,  align: 'left' },
    { k: 'pqty',   label: 'Qty',      x: PAGE_L + 50,   w: 62,  align: 'right', fmt: fmtQty },
    { k: 'rate',   label: 'Rate',     x: PAGE_L + 112,  w: 52,  align: 'right', fmt: fmtAmt },
    { k: 'puramt', label: 'Amount',   x: PAGE_L + 164,  w: 74,  align: 'right', fmt: fmtAmt },
    { k: 'refund', label: 'Discount', x: PAGE_L + 238,  w: 64,  align: 'right', fmt: fmtAmt },
    { k: 'tax',    label: 'GST',      x: PAGE_L + 302,  w: 56,  align: 'right', fmt: fmtAmt },
    { k: 'tds',    label: 'TDS',      x: PAGE_L + 358,  w: 56,  align: 'right', fmt: fmtAmt },
    { k: 'payable',label: 'Payable',  x: PAGE_L + 414,  w: 121, align: 'right', fmt: fmtAmt },
  ];
  doc.font('Helvetica-Bold').fontSize(9);
  doc.rect(PAGE_L, y, PAGE_W, 18).fillAndStroke('#f3f4f6', '#999').fillColor('#000');
  for (const c of cols) doc.text(c.label, c.x + 2, y + 5, { width: c.w - 4, align: c.align });
  y += 18;
  doc.font('Helvetica').fontSize(9).fillColor('#000');
  let tQty=0,tAmt=0,tDisc=0,tTax=0,tTds=0,tPay=0;
  // Payable per lot = balance − the lot's TDS share. When invoice rounding is
  // on (cfg.flag_round), distribute whole-rupee rounding across the seller's
  // lots so each line is a whole rupee AND the lines sum exactly to the
  // seller's rounded Payable (matching the Payments tab + bank file).
  const roundPay = cfg.flag_round === true || String(cfg.flag_round || '').toLowerCase() === 'true';
  const rawPayables = lots.map(l => (Number(l.balance)||0) - round2((Number(l.puramt)||0) * tdsRate));
  const payables = roundPay ? distributeRoundedPayable(rawPayables) : rawPayables.map(round2);
  lots.forEach((l, i) => {
    // GST + TDS allocated from the seller's live totals (not the stale per-lot
    // stamps); Payable = balance − the lot's TDS share.
    const tax = round2((Number(l.puramt)||0) * gstRate);
    const lotTds = round2((Number(l.puramt)||0) * tdsRate);
    const payable = payables[i];
    const row = { ...l, tax, tds: lotTds, payable };
    tQty+=Number(l.pqty)||0; tAmt+=Number(l.puramt)||0; tDisc+=Number(l.refund)||0; tTax+=tax; tTds+=lotTds; tPay+=payable;
    if (y > 770) { doc.addPage(); y = 40; }
    for (const c of cols) {
      const v = c.fmt ? c.fmt(row[c.k]) : String(row[c.k] ?? '');
      doc.text(v, c.x + 2, y + 4, { width: c.w - 4, align: c.align });
    }
    y += 14;
    doc.moveTo(PAGE_L, y).lineTo(PAGE_R, y).strokeColor('#e5e7eb').lineWidth(0.5).stroke().strokeColor('#000');
  });
  doc.font('Helvetica-Bold').fontSize(10);
  doc.rect(PAGE_L, y, PAGE_W, 20).fillAndStroke('#f3f4f6', '#666').fillColor('#000');
  doc.text('TOTAL', PAGE_L + 2, y + 6);
  // Right-align each total under its column (x..x+w), reusing the cols layout.
  const totByKey = { pqty: fmtQty(tQty), puramt: fmtAmt(tAmt), refund: fmtAmt(tDisc), tax: fmtAmt(tTax), tds: fmtAmt(tTds), payable: fmtAmt(tPay) };
  for (const c of cols) {
    if (totByKey[c.k] == null) continue;
    doc.text(totByKey[c.k], c.x + 2, y + 6, { width: c.w - 4, align: 'right' });
  }
  y += 30;
  doc.font('Helvetica').fontSize(9).text(`Generated: ${new Date().toLocaleString('en-IN')}`, PAGE_L, y, { width: PAGE_W, align: 'right' });
  return tPay;
}

app.get('/api/payments/pdf/:auctionId/:sellerName', requireView, (req, res) => {
  let doc, piped = false;
  try {
    const db = getDb();
    const cfg = getSettingsFlat(db);
    const auctionId = req.params.auctionId;
    const sellerName = decodeURIComponent(req.params.sellerName);
    // Validate BEFORE piping — once we pipe, any subsequent throw can't
    // be turned into a JSON response (headers + body already going out).
    const auction = db.get('SELECT id FROM auctions WHERE id = ?', [auctionId]);
    if (!auction) return res.status(404).json({ error: 'Auction not found' });
    const PDFDocument = require('pdfkit');
    doc = new PDFDocument({ size: 'A4', margin: 30 });
    res.setHeader('Content-Type', 'application/pdf');
    // Sanitize the seller name before embedding in Content-Disposition.
    // Raw URL-encoded params or special chars (quotes, CR/LF, semicolons)
    // produce malformed headers — browsers then fail the print/download
    // silently. Same pattern the purchase-invoice + sales-invoice routes use.
    const safeName = String(sellerName || '').replace(/[^\w]/g, '_').slice(0, 80) || 'seller';
    res.setHeader('Content-Disposition', `inline; filename="Payment_${safeName}_${anoForFile(db, auctionId)}.pdf"`);
    doc.pipe(res); piped = true;
    res.on('close', () => { try { doc.destroy(); } catch(_){} });
    _renderPaymentStatement(doc, db, auctionId, sellerName, cfg);
    doc.end();
  } catch (e) {
    if (piped && doc) { try { doc.end(); } catch(_){} }
    else if (!res.headersSent) res.status(500).json({ error: e.message || 'PDF failed' });
  }
});

// Per-seller, lot-filtered PDF. Body { auction_id, seller_name, lot_ids: [...] }
// Powers the Payments tab's "partial payment" flow — operator opens the
// lots-detail modal for one seller, ticks the lots they're settling now,
// and prints a statement that only covers that subset. Remaining lots
// can be paid in a later run with the same flow.
app.post('/api/payments/pdf-lots', requireView, (req, res) => {
  let doc, piped = false;
  try {
    const db = getDb();
    const cfg = getSettingsFlat(db);
    const auctionId = Number(req.body.auction_id);
    const sellerName = String(req.body.seller_name || '').trim();
    const lotIds = Array.isArray(req.body.lot_ids) ? req.body.lot_ids : [];
    if (!auctionId || !sellerName || !lotIds.length) {
      return res.status(400).json({ error: 'auction_id, seller_name and lot_ids[] are required' });
    }
    const auction = db.get('SELECT id FROM auctions WHERE id = ?', [auctionId]);
    if (!auction) return res.status(404).json({ error: 'Auction not found' });
    const PDFDocument = require('pdfkit');
    doc = new PDFDocument({ size: 'A4', margin: 30 });
    res.setHeader('Content-Type', 'application/pdf');
    const safeName = sellerName.replace(/[^\w]/g, '_').slice(0, 80) || 'seller';
    res.setHeader('Content-Disposition', `inline; filename="Payment_${safeName}_${anoForFile(db, auctionId)}_partial.pdf"`);
    doc.pipe(res); piped = true;
    res.on('close', () => { try { doc.destroy(); } catch(_){} });
    _renderPaymentStatement(doc, db, auctionId, sellerName, cfg, lotIds);
    doc.end();
  } catch (e) {
    if (piped && doc) { try { doc.end(); } catch(_){} }
    else if (!res.headersSent) res.status(500).json({ error: e.message || 'PDF failed' });
  }
});

// Bulk: Body { auction_id, names: [...] } → one merged PDF, page-break per seller.
app.post('/api/payments/pdf-bulk', requireView, (req, res) => {
  let doc, piped = false;
  try {
    const db = getDb();
    const cfg = getSettingsFlat(db);
    const auctionId = Number(req.body.auction_id);
    const names = Array.isArray(req.body.names) ? req.body.names : [];
    if (!auctionId || !names.length) return res.status(400).json({ error: 'auction_id and names[] required' });
    const auction = db.get('SELECT id FROM auctions WHERE id = ?', [auctionId]);
    if (!auction) return res.status(404).json({ error: 'Auction not found' });
    const PDFDocument = require('pdfkit');
    doc = new PDFDocument({ size: 'A4', margin: 30 });
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `inline; filename="Payments_Batch_${names.length}.pdf"`);
    doc.pipe(res); piped = true;
    res.on('close', () => { try { doc.destroy(); } catch(_){} });
    names.forEach((nm, i) => {
      if (i > 0) doc.addPage();
      try { _renderPaymentStatement(doc, db, auctionId, nm, cfg); }
      catch (e) { try { doc.font('Helvetica').fontSize(10).text(`Error rendering ${nm}: ${e.message}`); } catch(_){} }
    });
    doc.end();
  } catch (e) {
    if (piped && doc) { try { doc.end(); } catch(_){} }
    else if (!res.headersSent) res.status(500).json({ error: e.message || 'PDF failed' });
  }
});

// ══════════════════════════════════════════════════════════════
// TDS RETURNS (TDSRETU.PRG)
// ══════════════════════════════════════════════════════════════
app.get('/api/tds-return', requireView, (req, res) => {
  const { from, to, orderBy } = req.query;
  if (!from || !to) return res.status(400).json({ error: 'from and to dates required' });
  const data = getTDSReturnData(getDb(), from, to, orderBy || 'invoice');
  res.json(data);
});

// ══════════════════════════════════════════════════════════════
// EXPORTS (EXP.PRG — all 11 types + TDS + Tally)
// ══════════════════════════════════════════════════════════════
app.get('/api/exports/:type/:auctionId', requireExport, async (req, res) => {
  const { type, auctionId } = req.params;
  const format = (req.query.format || 'xlsx').toLowerCase();

  if (format === 'pdf') {
    try {
      const db = getDb();
      const cfg = getSettingsFlat(db);
      const buffer = await exportAnyPdf(db, type, auctionId, cfg, { state: req.query.state });
      const niceName = (EXPORT_TYPES[type] && EXPORT_TYPES[type].name) || type;
      res.setHeader('Content-Type', 'application/pdf');
      res.setHeader('Content-Disposition', `attachment; filename="${niceName}_${anoForFile(db, auctionId)}.pdf"`);
      return res.send(buffer);
    } catch (e) {
      return res.status(500).json({ error: e.message });
    }
  }

  const exportDef = EXPORT_TYPES[type];
  if (!exportDef) return res.status(400).json({ error: 'Unknown export type', available: Object.keys(EXPORT_TYPES) });

  try {
    const db = getDb();
    let buffer;
    // Optional seller-name filter — drives the "Export Selected" buttons
    // in the Payments tab. The Bank Payment + Payment XLSX exporters
    // read this through their 5th arg; every other exporter ignores
    // unknown opts, so this is a no-op for them. Accept BOTH shapes:
    //   ?names=A&names=B&names=C   → req.query.names = ['A','B','C']
    //   ?names=A,B,C               → req.query.names = 'A,B,C'  → split
    let rawNames = req.query.names;
    if (typeof rawNames === 'string') rawNames = rawNames.split(',');
    if (!Array.isArray(rawNames)) rawNames = [];
    const names = rawNames.map(s => String(s || '').trim()).filter(Boolean);
    const opts = names.length ? { names } : undefined;
    if (exportDef.needsCfg) {
      const cfg = getSettingsFlat(db);
      // Pass state too so exports that need both (e.g. Praman) can filter
      // by state without losing cfg context. Backward-compatible: existing
      // needsCfg exports that ignore the 4th/5th args are unaffected.
      buffer = await exportDef.fn(db, auctionId, cfg, req.query.state, opts);
    } else {
      buffer = await exportDef.fn(db, auctionId, req.query.state, opts);
    }
    // Per-export-type content-type/extension override (defaults to xlsx).
    // CSV exports like Praman use ext:'csv', mime:'text/csv'.
    const ext  = exportDef.ext  || 'xlsx';
    const mime = exportDef.mime || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    // Tag the download filename with "_selected" when a names filter is
    // active so a partial-export file is obviously a subset on disk and
    // doesn't get confused with the full export later.
    const suffix = opts ? '_selected' : '';
    res.setHeader('Content-Type', mime);
    res.setHeader('Content-Disposition', `attachment; filename="${exportDef.name}${suffix}_${anoForFile(db, auctionId)}.${ext}"`);
    res.send(Buffer.from(buffer));
  } catch(e) {
    res.status(500).json({ error: e.message });
  }
});

// POST sibling of the GET export route — same response shape, but the
// "selected" filters live in the JSON body so:
//   (a) URL length isn't a constraint when many sellers are ticked,
//   (b) the per-seller lot-pick map ({ "Alice": ["12","15"] }) can be
//       sent as a normal nested object instead of being marshalled
//       through query-string bracket notation.
//
// Used by the Payments tab's "Export Bank Payment (Selected)" and
// "Export Payment XLSX (Selected)" buttons. The GET route stays the
// canonical path for the no-filter case (all sellers, all lots).
app.post('/api/exports/:type/:auctionId', requireExport, async (req, res) => {
  const { type, auctionId } = req.params;
  const format = String((req.body && req.body.format) || 'xlsx').toLowerCase();

  const exportDef = EXPORT_TYPES[type];
  if (!exportDef) return res.status(400).json({ error: 'Unknown export type', available: Object.keys(EXPORT_TYPES) });

  try {
    const db = getDb();
    // Normalise the body. names → array of trimmed non-empty strings.
    // lots → object mapping seller-name → array of lot_no strings.
    // excludeLots → same shape; lots that already shipped in earlier
    //               exports and must not be included again. Lets the
    //               client re-export a seller's remaining lots without
    //               accidentally double-paying the ones already paid.
    const body = req.body || {};
    const names = Array.isArray(body.names)
      ? body.names.map(s => String(s || '').trim()).filter(Boolean)
      : [];
    const lots        = (body.lots        && typeof body.lots        === 'object' && !Array.isArray(body.lots))        ? body.lots        : null;
    const excludeLots = (body.excludeLots && typeof body.excludeLots === 'object' && !Array.isArray(body.excludeLots)) ? body.excludeLots : null;
    const opts = (names.length || lots || excludeLots) ? {} : undefined;
    if (opts) {
      if (names.length) opts.names = names;
      const cleanMap = (src) => {
        const cleaned = {};
        for (const k of Object.keys(src)) {
          const arr = Array.isArray(src[k]) ? src[k].map(v => String(v || '').trim()).filter(Boolean) : [];
          if (arr.length) cleaned[k] = arr;
        }
        return Object.keys(cleaned).length ? cleaned : null;
      };
      if (lots) {
        const c = cleanMap(lots);
        if (c) opts.lots = c;
      }
      if (excludeLots) {
        const c = cleanMap(excludeLots);
        if (c) opts.excludeLots = c;
      }
    }
    // PDF variant of the selected export — thread `opts` through the PDF
    // generator the same way the GET route threads ?state=. The PDF row
    // source (exports-pdf.js getRowsForType) applies the names/lots/
    // excludeLots filter for the payment type.
    if (format === 'pdf') {
      const cfg = getSettingsFlat(db);
      const buffer = await exportAnyPdf(db, type, auctionId, cfg, { state: body.state || null, opts });
      const niceName = exportDef.name || type;
      const suffix = opts ? '_selected' : '';
      res.setHeader('Content-Type', 'application/pdf');
      res.setHeader('Content-Disposition', `attachment; filename="${niceName}${suffix}_${anoForFile(db, auctionId)}.pdf"`);
      return res.send(buffer);
    }
    let buffer;
    if (exportDef.needsCfg) {
      const cfg = getSettingsFlat(db);
      buffer = await exportDef.fn(db, auctionId, cfg, body.state || null, opts);
    } else {
      buffer = await exportDef.fn(db, auctionId, body.state || null, opts);
    }
    const ext  = exportDef.ext  || 'xlsx';
    const mime = exportDef.mime || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    const suffix = opts ? '_selected' : '';
    res.setHeader('Content-Type', mime);
    res.setHeader('Content-Disposition', `attachment; filename="${exportDef.name}${suffix}_${anoForFile(db, auctionId)}.${ext}"`);
    res.send(Buffer.from(buffer));
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// TDS export (supports ?format=pdf)
app.get('/api/exports/tds-return', requireExport, async (req, res) => {
  const { from, to } = req.query;
  const format = (req.query.format || 'xlsx').toLowerCase();
  if (!from || !to) return res.status(400).json({ error: 'from/to required' });
  if (format === 'pdf') {
    const buffer = await exportAnyPdf(getDb(), 'tds_return', null, null, { from, to });
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', 'attachment; filename="TDSReturn.pdf"');
    return res.send(buffer);
  }
  const { exportTDSReturn } = require('./exports');
  const buffer = await exportTDSReturn(getDb(), from, to);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename="TDSReturn.xlsx"');
  res.send(Buffer.from(buffer));
});

// ══════════════════════════════════════════════════════════════
// LORRY REPORTS (Lot Slip Code / Truck List / Buyer Lot Lorry)
// ══════════════════════════════════════════════════════════════
app.get('/api/lorry-reports/:type/:auctionId', requireExport, async (req, res) => {
  const { type, auctionId } = req.params;
  const format = (req.query.format || 'xlsx').toLowerCase();
  const def = LORRY_REPORTS[type];
  if (!def) return res.status(400).json({ error: 'Unknown lorry report', available: Object.keys(LORRY_REPORTS) });
  try {
    const db = getDb();
    if (format === 'pdf') {
      const buf = await def.pdf(db, auctionId);
      res.setHeader('Content-Type', 'application/pdf');
      res.setHeader('Content-Disposition', `attachment; filename="${def.name}_${anoForFile(db, auctionId)}.pdf"`);
      return res.send(buf);
    }
    const buf = await def.xlsx(db, auctionId);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${def.name}_${anoForFile(db, auctionId)}.xlsx"`);
    return res.send(Buffer.from(buf));
  } catch (e) {
    console.error('lorry-reports error:', e);
    return res.status(500).json({ error: e.message });
  }
});

// ══════════════════════════════════════════════════════════════
// TRADE REPORTS — JSON + PDF summaries for an auction
// ══════════════════════════════════════════════════════════════

// GET /api/reports/trade-summary/:auctionId
// JSON snapshot of one auction's activity: per-branch, per-seller,
// per-grade breakdowns, plus hourly entry rate. Drives the desktop
// "Reports → Trade Summary" view.
app.get('/api/reports/trade-summary/:auctionId', requireViewOrLotEntry, (req, res) => {
  const db = getDb();
  const auctionId = parseInt(req.params.auctionId, 10);
  const auction = db.get('SELECT * FROM auctions WHERE id = ?', [auctionId]);
  if (!auction) return res.status(404).json({ error: 'Trade not found' });

  // Optional branch filter — when set, the *aggregates* (Sold/Withdrawn/
  // Min/Max/Avg) restrict to lots from that branch. The branch list
  // itself is always returned in full so the UI can offer easy switching.
  const branchFilter = String(req.query.branch || '').trim();
  const bWhere = branchFilter ? ' AND branch = ?' : '';
  const bParams = branchFilter ? [branchFilter] : [];

  // Per-branch — primary view dad uses to compare today's branches.
  const branchWise = db.all(
    `SELECT branch,
            COUNT(*)             AS lot_count,
            SUM(bags)            AS total_bags,
            SUM(qty)             AS total_qty,
            COUNT(DISTINCT trader_id) AS seller_count,
            SUM(CASE WHEN amount > 0 THEN 1 ELSE 0 END) AS sold_lots,
            SUM(CASE WHEN amount > 0 THEN qty ELSE 0 END) AS sold_qty,
            SUM(CASE WHEN COALESCE(amount,0) <= 0 THEN 1 ELSE 0 END) AS withdrawn_lots,
            SUM(CASE WHEN COALESCE(amount,0) <= 0 THEN qty ELSE 0 END) AS withdrawn_qty,
            MAX(CASE WHEN amount > 0 THEN price END) AS max_price,
            MIN(CASE WHEN amount > 0 THEN price END) AS min_price,
            CASE WHEN SUM(CASE WHEN amount > 0 THEN qty ELSE 0 END) > 0
                 THEN SUM(CASE WHEN amount > 0 THEN amount ELSE 0 END) * 1.0
                      / SUM(CASE WHEN amount > 0 THEN qty ELSE 0 END)
                 ELSE 0 END AS avg_price
       FROM lots WHERE auction_id = ?
      GROUP BY branch ORDER BY total_qty DESC`,
    [auctionId]
  );

  // Sold/Withdrawn/Min/Max/Avg aggregates — branch-filtered if requested.
  const aggSold = db.get(
    `SELECT COUNT(*) AS lots, COALESCE(SUM(qty),0) AS qty,
            COALESCE(SUM(bags),0) AS bags, COALESCE(SUM(amount),0) AS cost
       FROM lots WHERE auction_id = ? AND amount > 0` + bWhere,
    [auctionId, ...bParams]
  ) || { lots:0, qty:0, bags:0, cost:0 };
  const aggWithdrawn = db.get(
    `SELECT COUNT(*) AS lots, COALESCE(SUM(qty),0) AS qty,
            COALESCE(SUM(bags),0) AS bags
       FROM lots WHERE auction_id = ? AND COALESCE(amount,0) <= 0` + bWhere,
    [auctionId, ...bParams]
  ) || { lots:0, qty:0, bags:0 };
  const aggPrice = db.get(
    `SELECT MIN(price) AS min_price, MAX(price) AS max_price
       FROM lots WHERE auction_id = ? AND amount > 0` + bWhere,
    [auctionId, ...bParams]
  ) || { min_price: 0, max_price: 0 };
  const avgPrice = aggSold.qty > 0 ? (Number(aggSold.cost) / Number(aggSold.qty)) : 0;
  const branchAggregates = {
    branch: branchFilter || null,
    sold:      { lots: aggSold.lots, bags: aggSold.bags, qty: aggSold.qty, cost: aggSold.cost },
    withdrawn: { lots: aggWithdrawn.lots, bags: aggWithdrawn.bags, qty: aggWithdrawn.qty },
    min: Number(aggPrice.min_price) || 0,
    max: Number(aggPrice.max_price) || 0,
    avg: avgPrice,
  };

  // Per-seller — pulls fresh names from traders master, falls back to
  // the denormalised lots.name field for legacy / mobile entries.
  const sellerWise = db.all(
    `SELECT COALESCE(t.name, l.name, 'Unknown') AS seller_name,
            l.trader_id,
            l.branch,
            COUNT(*)        AS lot_count,
            SUM(l.bags)     AS total_bags,
            SUM(l.qty)      AS total_qty
       FROM lots l
       LEFT JOIN traders t ON t.id = l.trader_id
      WHERE l.auction_id = ?
      GROUP BY COALESCE(l.trader_id, l.name)
      ORDER BY total_qty DESC`,
    [auctionId]
  );

  // Per-user (optional, gated by show_username setting).
  const userWise = db.all(
    `SELECT user_id,
            COUNT(*)    AS lot_count,
            SUM(bags)   AS total_bags,
            SUM(qty)    AS total_qty
       FROM lots WHERE auction_id = ? AND user_id != ''
      GROUP BY user_id ORDER BY lot_count DESC`,
    [auctionId]
  );

  // Hourly bucket — shows the rhythm of the auction day.
  const hourly = db.all(
    `SELECT substr(created_at, 12, 2) AS hour,
            COUNT(*)  AS lot_count,
            SUM(qty)  AS total_qty
       FROM lots WHERE auction_id = ?
      GROUP BY hour ORDER BY hour ASC`,
    [auctionId]
  );

  // Grade breakdown (1, 2, GRD, ungraded etc.)
  const gradeWise = db.all(
    `SELECT grade,
            COUNT(*)    AS lot_count,
            SUM(bags)   AS total_bags,
            SUM(qty)    AS total_qty,
            COUNT(DISTINCT trader_id) AS seller_count
       FROM lots WHERE auction_id = ?
      GROUP BY grade ORDER BY grade ASC`,
    [auctionId]
  );

  const totals = db.get(
    `SELECT COUNT(*)               AS lot_count,
            SUM(bags)              AS total_bags,
            SUM(qty)               AS total_qty,
            COUNT(DISTINCT trader_id) AS seller_count,
            COUNT(DISTINCT branch) AS branch_count
       FROM lots WHERE auction_id = ?`,
    [auctionId]
  ) || { lot_count: 0, total_bags: 0, total_qty: 0, seller_count: 0, branch_count: 0 };

  // show_username toggle controls whether the per-user list is sent.
  const showUserRow = db.get(`SELECT value FROM company_settings WHERE key = 'show_username'`);
  const showUsername = !!(showUserRow && showUserRow.value === 'true');

  res.json({
    auction, totals, branchWise, sellerWise,
    userWise: showUsername ? userWise : [],
    hourly, gradeWise, showUsername,
    branchAggregates,
  });
});

// GET /api/reports/branch-comparison
// Cross-trade comparison: how each branch has performed across all
// auctions in the DB. Heavy query — only runs on demand.
app.get('/api/reports/branch-comparison', requireViewOrLotEntry, (req, res) => {
  const db = getDb();
  const data = db.all(
    `SELECT l.branch, a.id AS auction_id, a.ano, a.date, a.crop_type,
            COUNT(*)    AS lot_count,
            SUM(l.bags) AS total_bags,
            SUM(l.qty)  AS total_qty
       FROM lots l JOIN auctions a ON a.id = l.auction_id
      GROUP BY l.branch, l.auction_id
      ORDER BY a.date DESC, l.branch ASC`
  );
  const overall = db.all(
    `SELECT branch,
            COUNT(*)              AS lot_count,
            SUM(bags)             AS total_bags,
            SUM(qty)              AS total_qty,
            COUNT(DISTINCT auction_id) AS trade_count,
            COUNT(DISTINCT trader_id)  AS seller_count
       FROM lots
      GROUP BY branch ORDER BY total_qty DESC`
  );
  res.json({ data, overall });
});

// GET /api/reports/summary-pdf/:auctionId
// One-page A4 PDF: headline totals, branch breakdown, top sellers.
// Window-opened via window.open() so we accept token via querystring
// (no Authorization header from window.open).
app.get('/api/reports/summary-pdf/:auctionId', (req, res, next) => {
  const hdr = (req.headers.authorization || '').replace('Bearer ', '');
  const tok = hdr || String(req.query.token || '');
  if (!tok) return res.status(401).json({ error: 'No token' });
  const db = getDb();
  const session = db.get('SELECT * FROM sessions WHERE token = ?', [tok]);
  if (!session) return res.status(403).json({ error: 'Session expired' });
  const user = db.get('SELECT * FROM users WHERE id = ?', [session.user_id]);
  if (!user) return res.status(403).json({ error: 'Unauthorized' });
  req.user = user;
  next();
}, (req, res) => {
  const db = getDb();
  const auctionId = parseInt(req.params.auctionId, 10);
  const auction = db.get('SELECT * FROM auctions WHERE id = ?', [auctionId]);
  if (!auction) return res.status(404).json({ error: 'Trade not found' });

  const titleRow = db.get(`SELECT value FROM company_settings WHERE key = 'trade_name'`);
  const appTitle = (titleRow && titleRow.value) || 'Spice Auction';
  const dateFmt  = auction.date ? fmtDate(auction.date) : '';

  const branchFilterPdf = String(req.query.branch || '').trim();
  const bWherePdf = branchFilterPdf ? ' AND branch = ?' : '';
  const bWherePdfL = branchFilterPdf ? ' AND l.branch = ?' : '';
  const bParamsPdf = branchFilterPdf ? [branchFilterPdf] : [];
  const totals = db.get(
    `SELECT COUNT(*) AS lots, SUM(bags) AS bags, SUM(qty) AS qty,
            COUNT(DISTINCT trader_id) AS sellers, COUNT(DISTINCT branch) AS branches
       FROM lots WHERE auction_id = ?` + bWherePdf,
    [auctionId, ...bParamsPdf]
  ) || { lots: 0, bags: 0, qty: 0, sellers: 0, branches: 0 };
  const branchWise = db.all(
    `SELECT branch, COUNT(*) AS lots, SUM(bags) AS bags, SUM(qty) AS qty,
            SUM(CASE WHEN amount > 0 THEN qty ELSE 0 END) AS sold_qty,
            SUM(CASE WHEN COALESCE(amount,0) <= 0 THEN qty ELSE 0 END) AS withdrawn_qty,
            MAX(CASE WHEN amount > 0 THEN price END) AS max_price,
            MIN(CASE WHEN amount > 0 THEN price END) AS min_price,
            CASE WHEN SUM(CASE WHEN amount > 0 THEN qty ELSE 0 END) > 0
                 THEN SUM(CASE WHEN amount > 0 THEN amount ELSE 0 END) * 1.0
                      / SUM(CASE WHEN amount > 0 THEN qty ELSE 0 END)
                 ELSE 0 END AS avg_price
       FROM lots WHERE auction_id = ? GROUP BY branch ORDER BY qty DESC`,
    [auctionId]
  );
  const sellerWise = db.all(
    `SELECT COALESCE(t.name, l.name, 'Unknown') AS name,
            l.branch,
            COUNT(*) AS lots, SUM(l.bags) AS bags, SUM(l.qty) AS qty
       FROM lots l LEFT JOIN traders t ON t.id = l.trader_id
      WHERE l.auction_id = ?` + bWherePdfL +
     ` GROUP BY COALESCE(l.trader_id, l.name) ORDER BY qty DESC`,
    [auctionId, ...bParamsPdf]
  );

  // Logo for the header — resolves persistent-volume upload first,
  // bundled default second, null if neither exists.
  const logoPath = _resolveLogoPath('logo-ispl.png') || _resolveLogoPath('logo_kj.png');

  const PDFDocument = require('pdfkit');
  const doc = new PDFDocument({ size: 'A4', margin: 40 });
  res.setHeader('Content-Type', 'application/pdf');
  res.setHeader('Content-Disposition',
    `inline; filename="Trade_${auction.ano}_Summary_${auction.date}.pdf"`);
  doc.pipe(res);

  const m = 40, w = 515;

  function drawRow(y, cols, font, size) {
    doc.font(font || 'Helvetica').fontSize(size || 9);
    cols.forEach(c => {
      doc.text(String(c.val == null ? '' : c.val), c.x, y,
        { width: c.w, align: c.align || 'left' });
    });
    return y + (size || 9) + 5;
  }

  // Header
  if (logoPath) {
    try { doc.image(logoPath, (595 - 45) / 2, doc.y, { width: 45, height: 45 }); doc.y += 50; } catch (_) {}
  }
  doc.font('Helvetica-Bold').fontSize(16).text(appTitle, m, doc.y, { width: w, align: 'center' });
  doc.moveDown(0.2);
  doc.font('Helvetica').fontSize(11).text(
    'Trade #' + auction.ano + '  |  ' + dateFmt + '  |  ' + (auction.crop_type || '') +
      (branchFilterPdf ? '  |  Branch: ' + branchFilterPdf : ''),
    m, doc.y, { width: w, align: 'center' }
  );
  doc.moveDown(0.5);
  doc.moveTo(m, doc.y).lineTo(m + w, doc.y).lineWidth(1).strokeColor('#166534').stroke();
  doc.strokeColor('#000');
  doc.moveDown(0.6);

  // Totals strip
  const sY = doc.y;
  const sItems = [
    { label: 'Lots',     val: totals.lots || 0 },
    { label: 'Bags',     val: totals.bags || 0 },
    { label: 'Qty (kg)', val: Number(totals.qty || 0).toFixed(3) },
    { label: 'Sellers',  val: totals.sellers || 0 },
    { label: 'Branches', val: totals.branches || 0 },
  ];
  const sW = w / sItems.length;
  sItems.forEach((s, i) => {
    const sx = m + i * sW;
    doc.font('Helvetica-Bold').fontSize(14).text(String(s.val), sx, sY, { width: sW, align: 'center' });
    doc.font('Helvetica').fontSize(8).fillColor('#666').text(s.label, sx, sY + 18, { width: sW, align: 'center' });
  });
  doc.fillColor('#000');
  doc.y = sY + 40;
  doc.moveDown(0.6);

  // Branch-wise table
  doc.font('Helvetica-Bold').fontSize(11).text('Branch-wise Breakdown', m);
  doc.moveDown(0.4);
  let y = doc.y;
  y = drawRow(y, [
    { x: m,       w: 200, val: 'Branch' },
    { x: m + 200, w: 70,  val: 'Lots', align: 'right' },
    { x: m + 270, w: 70,  val: 'Bags', align: 'right' },
    { x: m + 340, w: 100, val: 'Qty (kg)', align: 'right' },
  ], 'Helvetica-Bold', 9);
  doc.moveTo(m, y - 2).lineTo(m + 440, y - 2).lineWidth(0.5).stroke();
  branchWise.forEach(b => {
    if (y > 770) { doc.addPage(); y = 40; }
    y = drawRow(y, [
      { x: m,       w: 200, val: b.branch || '' },
      { x: m + 200, w: 70,  val: b.lots, align: 'right' },
      { x: m + 270, w: 70,  val: b.bags, align: 'right' },
      { x: m + 340, w: 100, val: Number(b.qty || 0).toFixed(3), align: 'right' },
    ]);
  });
  doc.y = y;
  doc.moveDown(0.8);

  // Top sellers table (cap at 30)
  if (sellerWise.length) {
    doc.font('Helvetica-Bold').fontSize(11).text('Top Sellers (up to 30)', m);
    doc.moveDown(0.4);
    y = doc.y;
    y = drawRow(y, [
      { x: m,       w: 170, val: 'Seller' },
      { x: m + 170, w: 110, val: 'Branch' },
      { x: m + 280, w: 50,  val: 'Lots', align: 'right' },
      { x: m + 330, w: 60,  val: 'Bags', align: 'right' },
      { x: m + 390, w: 80,  val: 'Qty (kg)', align: 'right' },
    ], 'Helvetica-Bold', 9);
    doc.moveTo(m, y - 2).lineTo(m + 470, y - 2).lineWidth(0.5).stroke();
    sellerWise.slice(0, 30).forEach(s => {
      if (y > 770) { doc.addPage(); y = 40; }
      y = drawRow(y, [
        { x: m,       w: 170, val: s.name || 'Unknown' },
        { x: m + 170, w: 110, val: s.branch || '' },
        { x: m + 280, w: 50,  val: s.lots, align: 'right' },
        { x: m + 330, w: 60,  val: s.bags, align: 'right' },
        { x: m + 390, w: 80,  val: Number(s.qty || 0).toFixed(3), align: 'right' },
      ]);
    });
    doc.y = y + 10;
  }

  // Footer
  doc.moveTo(m, doc.y).lineTo(m + w, doc.y).lineWidth(0.5).stroke();
  doc.moveDown(0.4);
  doc.font('Helvetica').fontSize(8).fillColor('#888')
     .text('Generated ' + new Date().toLocaleString('en-IN'), m, doc.y, { width: w, align: 'center' });

  doc.end();
});

// ══════════════════════════════════════════════════════════════
// DBF EXPORTS (FoxPro-compatible format)
// ══════════════════════════════════════════════════════════════

// List all available DBF/XLSX export types with labels + capability flags.
app.get('/api/dbf-exports/list', requireExport, (req, res) => {
  const list = {};
  for (const [key, def] of Object.entries(DBF_EXPORTS)) {
    list[key] = {
      label: def.label,
      name: def.name,
      trade: !!def.trade,
      dateRange: !!def.dateRange,
    };
  }
  res.json(list);
});

// Generic export endpoint — same data, two formats. ?format=xlsx returns
// the branded spreadsheet twin; anything else (default) returns the .dbf.
// Filters (all optional, applied where the module supports them):
//   ?auctionId=  trade-wise (lots key on it directly; transaction tables
//                resolve it to their ano)
//   ?ano=        explicit trade number (wins over auctionId)
//   ?from=&to=   inclusive date range
app.get('/api/dbf-exports/:type', requireExport, async (req, res) => {
  const { type } = req.params;
  const def = DBF_EXPORTS[type];
  if (!def) return res.status(400).json({ error: 'Unknown export type', available: Object.keys(DBF_EXPORTS) });

  const format = String(req.query.format || 'dbf').toLowerCase();

  try {
    const db = getDb();
    const opts = {};
    if (def.trade) {
      if (req.query.auctionId) opts.auctionId = req.query.auctionId;
      if (req.query.ano) opts.ano = req.query.ano;
    }
    if (def.dateRange && req.query.from && req.query.to) {
      opts.from = req.query.from;
      opts.to = req.query.to;
    }

    // Build a descriptive filename keyed by the trade NO (ano), not the
    // internal auctions.id: CPA1_4.xlsx, INV_2026-04-01_to_2026-04-30.dbf, NAM.dbf
    let filename = def.name;
    if (opts.auctionId) filename += `_${anoForFile(db, opts.auctionId)}`;
    else if (opts.ano) filename += `_${String(opts.ano).replace(/[^\w.-]+/g, '_')}`;
    if (opts.from && opts.to) filename += `_${opts.from}_to_${opts.to}`;

    if (format === 'xlsx') {
      const buffer = await exportXlsx(db, type, opts);
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', `attachment; filename="${filename}.xlsx"`);
      return res.send(Buffer.from(buffer));
    }

    const buffer = await exportDbf(db, type, opts);
    res.setHeader('Content-Type', 'application/x-dbase');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}.dbf"`);
    res.send(buffer);
  } catch(e) {
    console.error('Export error:', e);
    res.status(500).json({ error: 'Export failed: ' + e.message });
  }
});

// ══════════════════════════════════════════════════════════════
// TO TALLY — XML exports for Tally accounting software
// ══════════════════════════════════════════════════════════════

// Definitions of available Tally exports — keep in sync with frontend.
//
// `company` resolves which Tally company name goes in the XML's
// <SVCURRENTCOMPANY> tag. For now this routing only differs across the
// 3 party-ledger types:
//   • Sales Party Ledgers → ISP (sales-side parties = ISP customers)
//   • RD / URD Party Ledgers → ASP (purchase-side parties = ASP suppliers)
// All other exports (vouchers, the all-in-one ledger master) currently
// import into the ISP company. We can split vouchers later if dad asks.
const TALLY_EXPORTS = {
  ledger_sales:        { label: 'Sales Party Ledgers',                              name: 'SalesPartyLedgers',  builder: buildSalesPartyLedgerRows, generator: generLedgerXML, isLedger: true, company: 'isp' },
  ledger_rd_purchase:  { label: 'RD Purchase Party Ledgers',                        name: 'RDPartyLedgers',     builder: buildRDPartyLedgerRows,    generator: generLedgerXML, isLedger: true, company: 'isp' },
  ledger_urd_purchase: { label: 'URD Purchase Party Ledgers (Agriculturist)',       name: 'URDPartyLedgers',    builder: buildURDPartyLedgerRows,   generator: generLedgerXML, isLedger: true, company: 'isp' },
  ledger:              { label: 'All Ledger Masters (parties + tax + sales + purchase)', name: 'AllLedgers',  builder: buildLedgerRows,           generator: generLedgerXML, isLedger: true, company: 'isp' },
  // ── Sales Vouchers — split into two purpose-built exports ────
  // sales_isp = ISP→outside-customer sales (full e-way bill, dispatch
  //   from sister, BASICORDERREF to matching ASP voucher).
  // sales_asp = ASP→ISP internal transfers (lean format, no e-way
  //   bill, customer is always ISP, lot rates from asp_prate/asp_puramt).
  // The legacy `sales` key is kept as an alias for sales_isp so any old
  // bookmarks / API callers don't break; new UI buttons use the split keys.
  sales_isp:           { label: 'Sales Vouchers',                                   name: 'Sales',              builder: buildSalesIspRows,         generator: generSalesIspXML,     company: 'isp' },
  sales_asp:           { label: 'Sales Vouchers (Kerala / Intra-Company)',          name: 'SalesIntra',         builder: buildSalesAspRows,         generator: generSalesAspXML,     company: 'isp' },
  sales:               { label: 'Sales Vouchers (legacy alias)',                    name: 'Sales',              builder: buildSalesIspRows,         generator: generSalesIspXML,     company: 'isp' },
  // ISP Purchase = the buyer-side mirror of an ASP→ISP transfer. Each
  // sales_asp row produces one isp_purchase voucher into ISP's books with
  // the same VOUCHERNUMBER (e.g. ASP/I-61/26-27) for cross-reference. We
  // re-use buildSalesAspRows directly since the row shape is identical.
  isp_purchase:        { label: 'Intra-Company Purchase Vouchers',                  name: 'IntraPurchase',      builder: buildSalesAspRows,         generator: generIspPurchaseXML,  company: 'isp' },
  rd_purchase:         { label: 'RD Purchase Vouchers',                             name: 'RDPurchase',         builder: buildRDPurchaseRows,       generator: generRDPurchaseXML,   company: 'isp' },
  urd_purchase:        { label: 'URD Purchase Vouchers (Agriculturist)',            name: 'URDPurchase',        builder: buildURDPurchaseRows,      generator: generURDPurchaseXML,  company: 'isp' },
  debit_note:          { label: 'Debit Notes (Discount)',                           name: 'DebitNote',          builder: buildDebitNoteRows,        generator: generDebitNoteXML,    company: 'isp' },
};

// Resolve the Tally company name for a given export type.
// 'isp' → tally_company_name; 'asp' → tally_asp_company_name (falls
// back to ISP if the ASP name is blank, but logs a warning so misconfig
// is visible — silently falling back has caused confusion when a user
// sees ISP in <SVCURRENTCOMPANY> but expected ASP).
function resolveTallyCompanyName(cfg, target) {
  const isp = (cfg.tally_company_name || '').trim();
  const asp = (cfg.tally_asp_company_name || '').trim();
  if (target === 'asp') {
    if (!asp) {
      console.warn('[tally] tally_asp_company_name is empty — falling back to ISP company name. Set it via Settings → To Tally → "ASP Tally Company Name".');
    }
    return asp || isp;
  }
  return isp;
}

// Map a single-party `kind` (sales|rd_purchase|urd_purchase) to its
// dedicated builder + which Tally company its ledger belongs to.
const PARTY_LEDGER_BUILDERS = {
  sales:        { builder: buildSalesPartyLedgerRows, company: 'isp' },
  rd_purchase:  { builder: buildRDPartyLedgerRows,    company: 'isp' },
  urd_purchase: { builder: buildURDPartyLedgerRows,   company: 'isp' },
};

// List endpoint — used by the To Tally tab to render export buttons
app.get('/api/tally/list', requireExport, (req, res) => {
  const list = {};
  for (const [key, def] of Object.entries(TALLY_EXPORTS)) {
    list[key] = { label: def.label, name: def.name };
  }
  res.json(list);
});

// Preview endpoint — returns row counts so the user knows how many vouchers
// will be in the XML before downloading
app.get('/api/tally/preview/:type/:auctionId', requireExport, (req, res) => {
  const { type, auctionId } = req.params;
  const def = TALLY_EXPORTS[type];
  if (!def) return res.status(400).json({ error: 'Unknown Tally export', available: Object.keys(TALLY_EXPORTS) });
  try {
    const db = getDb();
    const cfg = getSettingsFlat(db);
    const rows = def.builder(db, auctionId, cfg);
    const targetCompany = resolveTallyCompanyName(cfg, def.company);
    if (def.isLedger) {
      // Ledger rows have a different shape — count by kind
      const byKind = rows.reduce((acc, r) => {
        const k = r.kind || 'other';
        acc[k] = (acc[k] || 0) + 1;
        return acc;
      }, {});
      return res.json({
        type, auctionId,
        ledgerCount: rows.length,
        byKind,
        targetCompany,
        sample: rows.slice(0, 6).map(r => ({ kind: r.kind, name: r.name, parent: r.parent, gstin: r.gstin || '' })),
      });
    }
    // Count DISTINCT lots, not the sum of per-voucher lot lists. A lot has
    // exactly one buyer and one seller, so it can never legitimately appear
    // in two vouchers — but a builder that scopes lots loosely (e.g. by
    // seller name when a seller has duplicate imported vouchers) could list
    // the same lot under multiple vouchers, which used to inflate this count
    // above the real lot total. Deduping by lot_no reports the true coverage.
    // Entries without a lot_no (shouldn't happen for lot-bearing types) are
    // counted individually as a safety net.
    const _lotSeen = new Set();
    let totalLots = 0;
    for (const r of rows) {
      if (!Array.isArray(r.lots)) continue;
      for (const l of r.lots) {
        const key = (l && l.lot != null && String(l.lot).trim() !== '') ? String(l.lot).trim() : null;
        if (key === null) { totalLots++; continue; }
        if (!_lotSeen.has(key)) { _lotSeen.add(key); totalLots++; }
      }
    }
    // For voucher types that have NO per-row lot list (debit notes,
    // journal entries, agri bills with single-line items), lotCount is
    // always 0 — that's a real-world correct value but it confused
    // users who saw "0 lots" and assumed the export was empty.
    // Also surface a `partyCount` (distinct dealer/buyer names) so the
    // preview can show that instead of the meaningless lot count for
    // these types. Caller (UI) decides which to display.
    const distinctParties = new Set();
    for (const r of rows) {
      const n = String(r.partyName || r.name || '').trim();
      if (n) distinctParties.add(n.toUpperCase());
    }
    res.json({
      type, auctionId,
      voucherCount: rows.length,
      lotCount: totalLots,
      partyCount: distinctParties.size,
      // Flag whether this voucher type carries a per-row lots array.
      // Lets the UI suppress "0 lots" for DN/journal/etc. without
      // hardcoding type names client-side.
      hasLots: rows.some(r => Array.isArray(r.lots) && r.lots.length > 0),
      targetCompany,
      sample: rows.slice(0, 3).map(r => ({
        ano: r.ano, date: r.date, name: r.partyName || r.name,
        voucher: r.voucherNum || r.invo,
        amount: r.total,
      })),
    });
  } catch (e) {
    console.error('tally preview error:', e);
    res.status(500).json({ error: e.message });
  }
});

// ══════════════════════════════════════════════════════════════
// E-WAY BILL DISTANCE — route table + per-invoice override
// ══════════════════════════════════════════════════════════════
// Two storage paths feed the <DISTANCE> field on ISP sales vouchers:
//
//   1. invoices.distance_km — per-invoice override. Wins when set.
//   2. route_distances[(from_pin, to_pin)] — saved route value, applied
//      to every invoice between the same two PINs.
//
// Workflow: user clicks NIC, looks up dispatch→consignee distance on
// the portal, types it in, clicks Save. We write to route_distances —
// every other invoice between the same two PINs (this auction and all
// future ones) auto-resolves.

// Resolve the configured dispatch PIN with the same fallback chain the
// voucher generator uses. Used for normalising route lookups.
function getDispatchPin(db) {
  const cfg = require('./company-config').getSettingsFlat(db);
  return String(
    cfg.tally_dispatch_pin || cfg.s_pin || cfg.kl_pin || cfg.tn_pin || ''
  ).trim();
}

// Normalise a (from, to) pair so A↔B share a single route_distances row.
// Always returns the lexicographically smaller PIN first.
function normalizeRouteKey(fromPin, toPin) {
  const a = String(fromPin || '').trim();
  const b = String(toPin || '').trim();
  return a < b ? [a, b] : [b, a];
}

// List ISP invoices for an auction with their resolved distance + source
// tag. The UI uses this to render the table — `km` is the value to
// display, `source` tells the user where it came from ('manual' = per-
// invoice override, 'route' = looked up by PIN pair, 'none' = blank).
app.get('/api/invoices/distances/:auctionId', requireView, (req, res) => {
  try {
    const db = getDb();
    const dispatchPin = getDispatchPin(db);
    const rows = db.all(
      `SELECT i.id, i.ano, i.invo, i.buyer, i.buyer1, i.gstin, i.state,
              b.pin AS buyer_bill_pin, b.cpin AS buyer_ship_pin,
              b.pla AS buyer_pla,
              i.distance_km
       FROM invoices i
       LEFT JOIN buyers b ON b.buyer = i.buyer
       WHERE i.auction_id = ?
       ORDER BY CAST(i.invo AS INTEGER), i.id`,
      [req.params.auctionId]
    );

    // Pre-fetch all route distances for this dispatch PIN — one query
    // instead of N. The set is small (a few dozen routes max) so it
    // fits comfortably in memory.
    const routes = {};
    try {
      const allRoutes = db.all(
        `SELECT from_pin, to_pin, km FROM route_distances
         WHERE from_pin = ? OR to_pin = ?`,
        [dispatchPin, dispatchPin]
      );
      for (const r of allRoutes) {
        // The "other PIN" — whichever side isn't the dispatch PIN
        const other = r.from_pin === dispatchPin ? r.to_pin : r.from_pin;
        routes[other] = r.km;
      }
    } catch (e) { /* table may not exist on very old DBs */ }

    // Annotate each row with resolved distance + source.
    //
    // SOURCE PINCODE PRIORITY (mirrors the voucher generator, the read-
    // hydration path, and the route-save path): ship-to (consignee) cpin
    // wins, bill-to pin is the fallback. The ship-to address is where the
    // goods physically arrive, so it's the correct e-way bill destination;
    // a buyer can bill to one PIN (e.g. head office) but take delivery at
    // another. `buyer_pin` is surfaced as this resolved value so display,
    // the NIC/Maps lookups, and the Save route key all key off the same PIN.
    const enriched = rows.map(r => {
      const shipPin = r.buyer_ship_pin ? String(r.buyer_ship_pin).trim() : '';
      const billPin = r.buyer_bill_pin ? String(r.buyer_bill_pin).trim() : '';
      const buyer_pin = shipPin || billPin;
      let km = null, source = 'none';
      if (r.distance_km != null) {
        km = r.distance_km;
        source = 'manual';
      } else if (buyer_pin && routes[buyer_pin] != null) {
        km = routes[buyer_pin];
        source = 'route';
      }
      return { ...r, buyer_pin, resolved_km: km, distance_source: source };
    });

    res.json({
      count: enriched.length,
      dispatchPin,
      invoices: enriched,
    });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// Save a route distance. Body: { from_pin, to_pin, km }. The pair gets
// normalised (smaller PIN first) before write, so subsequent lookups
// find it regardless of direction. Empty/null `km` deletes the row.
//
// Returns: how many ISP invoices now resolve via this route (so the
// UI can show "applied to N invoices" feedback).
app.put('/api/route-distances', requireExport, (req, res) => {
  const { from_pin, to_pin, km } = req.body || {};
  if (!from_pin || !to_pin) {
    return res.status(400).json({ error: 'from_pin and to_pin required' });
  }
  if (!/^\d{6}$/.test(String(from_pin).trim()) || !/^\d{6}$/.test(String(to_pin).trim())) {
    return res.status(400).json({ error: 'PINs must be 6-digit strings' });
  }
  const [k1, k2] = normalizeRouteKey(from_pin, to_pin);

  // Empty km = delete
  if (km === '' || km == null) {
    try {
      const r = getDb().run(
        'DELETE FROM route_distances WHERE from_pin = ? AND to_pin = ?',
        [k1, k2]
      );
      return res.json({ ok: true, deleted: r.changes > 0, from_pin: k1, to_pin: k2 });
    } catch (e) { return res.status(500).json({ error: e.message }); }
  }

  const v = Math.round(Number(km));
  if (!isFinite(v) || v < 0 || v > 5000) {
    return res.status(400).json({ error: 'km must be between 0 and 5000' });
  }

  try {
    const db = getDb();
    db.run(
      `INSERT INTO route_distances (from_pin, to_pin, km, updated_at)
       VALUES (?, ?, ?, datetime('now','localtime'))
       ON CONFLICT(from_pin, to_pin) DO UPDATE
         SET km = excluded.km, updated_at = excluded.updated_at`,
      [k1, k2, v]
    );

    // Saving a route is the user's signal that "this distance applies to
    // every invoice between these PINs." Clear any legacy per-invoice
    // overrides on matching invoices so the route value actually wins —
    // otherwise leftover invoices.distance_km values from earlier saves
    // would shadow the route forever. (Per-invoice overrides have higher
    // priority by design; if we ever add a UI to set a true per-invoice
    // override, this clearing step would need to be opt-in.)
    //
    // PIN-resolution mirrors the read path: ship-to (cpin) wins, with
    // bill-to (pin) as fallback. Otherwise routes saved against a
    // ship-to PIN wouldn't apply to invoices whose buyer has a
    // separate consignee address.
    const dispatchPin = getDispatchPin(db);
    const otherPin = k1 === dispatchPin ? k2 : (k2 === dispatchPin ? k1 : null);
    let clearedOverrides = 0;
    if (otherPin) {
      const r = db.run(
        `UPDATE invoices SET distance_km = NULL
         WHERE id IN (
           SELECT i.id FROM invoices i
           LEFT JOIN buyers b ON b.buyer = i.buyer
           WHERE COALESCE(NULLIF(TRIM(b.cpin), ''), TRIM(b.pin)) = ?
             AND i.distance_km IS NOT NULL
         )`,
        [otherPin]
      );
      clearedOverrides = r.changes || 0;
    }

    // How many invoices now resolve via this route? Now that we cleared
    // the legacy overrides, every invoice with the matching destination
    // PIN counts (ship-to first, bill-to fallback).
    let appliedCount = 0;
    if (otherPin) {
      const r = db.get(
        `SELECT COUNT(*) AS n FROM invoices i
         LEFT JOIN buyers b ON b.buyer = i.buyer
         WHERE COALESCE(NULLIF(TRIM(b.cpin), ''), TRIM(b.pin)) = ?`,
        [otherPin]
      );
      appliedCount = r ? r.n : 0;
    }

    res.json({ ok: true, from_pin: k1, to_pin: k2, km: v, appliedCount, clearedOverrides });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// Bulk-clear all per-invoice distance overrides. Used to wipe legacy
// invoices.distance_km values from before the route-table refactor —
// after running this, every ISP invoice resolves via the route table
// (or stays blank if no route exists). Confirmed via the UI before
// hitting this; no body needed.
//
// Path is under /api/distance-overrides (not /api/invoices/distance-
// overrides) to avoid Express matching it against the earlier-defined
// app.delete('/api/invoices/:id') route, which treats 'distance-
// overrides' as an :id and returns 'Invoice not found'.
app.delete('/api/distance-overrides', requireExport, (req, res) => {
  try {
    const r = getDb().run(
      `UPDATE invoices SET distance_km = NULL
       WHERE distance_km IS NOT NULL`
    );
    res.json({ ok: true, cleared: r.changes });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// Per-invoice override (kept in the API even though no UI button wires
// to it directly — useful for one-off exceptions where a single invoice
// needs a different distance than the route. Set null to clear.)
app.put('/api/invoices/:id/distance', requireExport, (req, res) => {
  const id = Number(req.params.id);
  const { distance_km } = req.body || {};
  let v = null;
  if (distance_km !== '' && distance_km != null) {
    v = Math.round(Number(distance_km));
    if (!isFinite(v) || v < 0 || v > 5000) {
      return res.status(400).json({ error: 'distance_km must be between 0 and 5000 km, or null to clear' });
    }
  }
  try {
    const r = getDb().run('UPDATE invoices SET distance_km = ? WHERE id = ?', [v, id]);
    if (r.changes === 0) return res.status(404).json({ error: 'Invoice not found' });
    res.json({ ok: true, id, distance_km: v });
  } catch (e) { res.status(500).json({ error: e.message }); }
});


// Party listing for an auction — used by the single-party picker UI.
// Returns every distinct party (buyer/RD/URD) with the kind it would be
// exported under so the frontend can group and filter.
app.get('/api/tally/parties/:auctionId', requireExport, (req, res) => {
  const { auctionId } = req.params;
  try {
    const db = getDb();
    const parties = listAuctionParties(db, auctionId);
    const byKind = parties.reduce((acc, p) => {
      acc[p.kind] = (acc[p.kind] || 0) + 1;
      return acc;
    }, {});
    res.json({ auctionId, total: parties.length, byKind, parties });
  } catch (e) {
    console.error('tally parties error:', e);
    res.status(500).json({ error: e.message });
  }
});

// Single-party ledger XML — emits exactly one ledger for the named party.
// kind: 'sales'|'rd_purchase'|'urd_purchase'  (matches party's source)
// Sales parties go into the ISP Tally company; RD/URD parties go into ASP.
app.get('/api/tally/party-ledger/:kind/:auctionId', requireExport, (req, res) => {
  const { kind, auctionId } = req.params;
  const partyName = req.query.name;
  if (!partyName) return res.status(400).json({ error: 'Missing ?name=<party name>' });
  const partyDef = PARTY_LEDGER_BUILDERS[kind];
  if (!partyDef) return res.status(400).json({ error: 'Unknown party kind', available: Object.keys(PARTY_LEDGER_BUILDERS) });
  try {
    const db = getDb();
    const cfg = getSettingsFlat(db);
    const rows = partyDef.builder(db, auctionId, cfg, { partyName });
    if (rows.length === 0) {
      return res.status(404).json({ error: `Party "${partyName}" not found in ${kind} for auction ${auctionId}` });
    }
    const xml = generLedgerXML(rows, cfg, { companyName: resolveTallyCompanyName(cfg, partyDef.company) });
    const safeName = String(partyName).replace(/[^a-zA-Z0-9]/g, '_').slice(0, 40);
    const filename = `Tally_PartyLedger_${kind}_${safeName}_${anoForFile(db, auctionId)}.xml`;
    res.setHeader('Content-Type', 'application/xml; charset=utf-8');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.send(xml);
  } catch (e) {
    console.error('tally party-ledger error:', e);
    res.status(500).json({ error: e.message });
  }
});

// Single-party voucher XML — emits exactly one voucher for the named
// party, for one of the voucher types: sales_isp, sales_asp, rd_purchase,
// urd_purchase, debit_note. Useful when a single buyer/dealer needs a
// voucher import in isolation (e.g. you missed one and don't want to
// re-import the whole auction).
//
// We reuse the existing TALLY_EXPORTS builders (which produce rows for
// the entire auction) and filter the rows by party name in-memory. The
// "party name" field varies per voucher type:
//   sales_isp     → row.partyName     (= invoices.buyer1, the buyer)
//   sales_asp     → row.buyerName     (= the downstream ISP-side buyer)
//   isp_purchase  → row.buyerName     (= same source as sales_asp; the
//                                       voucher's "party" is always ASP,
//                                       so we filter by the downstream
//                                       buyer to let users pick a single
//                                       transfer voucher)
//   rd_purchase   → row.name          (= purchases.name, the dealer)
//   urd_purchase  → row.name          (= bills.name, the agriculturist)
//   debit_note    → row.partyName     (= the discount-paying supplier)
const VOUCHER_PARTY_KEY = {
  sales_isp:    (r) => r.partyName || '',
  sales_asp:    (r) => r.buyerName || r.buyer || '',
  isp_purchase: (r) => r.buyerName || r.buyer || '',
  rd_purchase:  (r) => r.name || '',
  urd_purchase: (r) => r.name || '',
  debit_note:   (r) => r.partyName || r.name || '',
  // Legacy alias still works
  sales:        (r) => r.partyName || '',
};

app.get('/api/tally/party-voucher/:type/:auctionId', requireExport, (req, res) => {
  const { type, auctionId } = req.params;
  const partyName = req.query.name;
  if (!partyName) return res.status(400).json({ error: 'Missing ?name=<party name>' });
  const def = TALLY_EXPORTS[type];
  const keyFn = VOUCHER_PARTY_KEY[type];
  if (!def || !keyFn || def.isLedger) {
    return res.status(400).json({
      error: 'Unknown or unsupported voucher type for single-party export',
      supported: Object.keys(VOUCHER_PARTY_KEY),
    });
  }
  try {
    const db = getDb();
    const cfg = getSettingsFlat(db);
    const allRows = def.builder(db, auctionId, cfg);
    // Case-insensitive exact match on the party name; fall back to
    // contains-match if no exact hit (handles minor whitespace/case
    // differences between the picker label and the underlying data).
    const target = String(partyName).trim().toUpperCase();
    let rows = allRows.filter(r => String(keyFn(r) || '').trim().toUpperCase() === target);
    if (rows.length === 0) {
      rows = allRows.filter(r => String(keyFn(r) || '').toUpperCase().includes(target));
    }
    if (rows.length === 0) {
      return res.status(404).json({
        error: `No ${def.label} found for "${partyName}" in auction ${auctionId}`,
        availableParties: [...new Set(allRows.map(keyFn).filter(Boolean))].slice(0, 20),
      });
    }
    const xml = def.generator(rows, cfg, { companyName: resolveTallyCompanyName(cfg, def.company) });
    const safeName = String(partyName).replace(/[^a-zA-Z0-9]/g, '_').slice(0, 40);
    const filename = `${def.name}_${safeName}_${anoForFile(db, auctionId)}.xml`;
    res.setHeader('Content-Type', 'application/xml; charset=utf-8');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.send(xml);
  } catch (e) {
    console.error('tally party-voucher error:', e);
    res.status(500).json({ error: e.message });
  }
});

// ── Single-invoice voucher ─────────────────────────────────────────
// Like the single-party voucher, but keyed to ONE invoice/voucher number
// rather than a whole party (a party may have several invoices). Covers
// the three invoice-bearing voucher types; debit notes are per-supplier
// aggregates with no per-invoice identity, so they're excluded.
//
// `ref` is a stable per-voucher key derived from the invoice number (and
// sale-type letter for sales, where L/I/E reuse the same number sequence).
// The list endpoint emits these refs; the download endpoint re-runs the
// builder and filters to the matching row(s).
const _stripPurchaseSuffix = (n) => String(n || '').replace(/\s*-\s*PURCHASE$/i, '').trim();
const SINGLE_INVOICE_VOUCHERS = {
  sales_isp: {
    kind:  'sales',
    ref:   (r) => `${String(r.sale || '').trim().toUpperCase()}#${String(r.invo || '').trim()}`,
    num:   (r) => (r.sale ? `${String(r.sale).trim()}-` : '') + String(r.invo || '').trim(),
    party: (r) => r.partyName || r.buyer || '',
    total: (r) => Number(r.totalRounded != null ? r.totalRounded : r.total) || 0,
  },
  rd_purchase: {
    kind:  'rd_purchase',
    ref:   (r) => `RD#${String(r.voucherNum || '').trim()}`,
    num:   (r) => String(r.voucherNum || '').trim(),
    party: (r) => _stripPurchaseSuffix(r.name),
    total: (r) => Number(r.totalRounded != null ? r.totalRounded : r.total) || 0,
  },
  urd_purchase: {
    kind:  'urd_purchase',
    ref:   (r) => `URD#${String(r.voucherNum || '').trim()}`,
    num:   (r) => String(r.voucherNum || '').trim(),
    party: (r) => _stripPurchaseSuffix(r.name),
    total: (r) => Number(r.totalRounded != null ? r.totalRounded : r.total) || 0,
  },
};

// List individual vouchers (one per invoice) for the picker.
app.get('/api/tally/vouchers/:auctionId', requireExport, (req, res) => {
  const { auctionId } = req.params;
  try {
    const db = getDb();
    const cfg = getSettingsFlat(db);
    const vouchers = [];
    for (const [type, vdef] of Object.entries(SINGLE_INVOICE_VOUCHERS)) {
      const def = TALLY_EXPORTS[type];
      let rows = [];
      try { rows = def.builder(db, auctionId, cfg) || []; } catch (e) { rows = []; }
      for (const r of rows) {
        vouchers.push({
          type, kind: vdef.kind,
          ref: vdef.ref(r), num: vdef.num(r),
          party: vdef.party(r), total: vdef.total(r),
        });
      }
    }
    const byKind = vouchers.reduce((acc, v) => { acc[v.kind] = (acc[v.kind] || 0) + 1; return acc; }, {});
    res.json({ auctionId, total: vouchers.length, byKind, vouchers });
  } catch (e) {
    console.error('tally vouchers error:', e);
    res.status(500).json({ error: e.message });
  }
});

// Single-invoice voucher XML — emits exactly the voucher(s) matching ?ref=.
app.get('/api/tally/invoice-voucher/:type/:auctionId', requireExport, (req, res) => {
  const { type, auctionId } = req.params;
  const ref = req.query.ref;
  if (!ref) return res.status(400).json({ error: 'Missing ?ref=<voucher ref>' });
  const vdef = SINGLE_INVOICE_VOUCHERS[type];
  const def = TALLY_EXPORTS[type];
  if (!vdef || !def) {
    return res.status(400).json({
      error: 'Unsupported voucher type for single-invoice export',
      supported: Object.keys(SINGLE_INVOICE_VOUCHERS),
    });
  }
  try {
    const db = getDb();
    const cfg = getSettingsFlat(db);
    const allRows = def.builder(db, auctionId, cfg) || [];
    const want = String(ref).trim();
    const rows = allRows.filter((r) => vdef.ref(r) === want);
    if (rows.length === 0) {
      return res.status(404).json({
        error: `No ${def.label} matching "${ref}" in auction ${auctionId}`,
        availableRefs: allRows.map(vdef.ref).slice(0, 30),
      });
    }
    const xml = def.generator(rows, cfg, { companyName: resolveTallyCompanyName(cfg, def.company) });
    const safe = want.replace(/[^a-zA-Z0-9]/g, '_').slice(0, 40);
    const filename = `${def.name}_Invoice_${safe}_${anoForFile(db, auctionId)}.xml`;
    res.setHeader('Content-Type', 'application/xml; charset=utf-8');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.send(xml);
  } catch (e) {
    console.error('tally invoice-voucher error:', e);
    res.status(500).json({ error: e.message });
  }
});

// XML download endpoint — the main thing.
// Lorry / vehicle no on sales vouchers comes straight from invoices.lorry_no
// (set via the Sales tab's "Set Lorry No" bulk action). The Tally row
// builder reads it and emits it into <VEHICLENUMBER>/<BASICSHIPVESSELNO>.
app.get('/api/tally/export/:type/:auctionId', requireExport, (req, res) => {
  const { type, auctionId } = req.params;
  const def = TALLY_EXPORTS[type];
  if (!def) return res.status(400).json({ error: 'Unknown Tally export', available: Object.keys(TALLY_EXPORTS) });
  try {
    const db = getDb();
    const cfg = getSettingsFlat(db);
    const rows = def.builder(db, auctionId, cfg);
    if (rows.length === 0) {
      const what = def.isLedger ? def.label.toLowerCase() : `${def.label.toLowerCase()}`;
      return res.status(404).json({ error: `No ${what} found for auction ${auctionId}` });
    }
    const xml = def.generator(rows, cfg, { companyName: resolveTallyCompanyName(cfg, def.company) });
    const filename = `${def.name}_${anoForFile(db, auctionId)}.xml`;
    res.setHeader('Content-Type', 'application/xml; charset=utf-8');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.send(xml);
  } catch (e) {
    console.error('tally export error:', e);
    res.status(500).json({ error: e.message });
  }
});

// ── Crop Receipt PDF ─────────────────────────────────────────
app.get('/api/receipt/:lotId', requireView, async (req, res) => {
  const db = getDb(); const cfg = getSettingsFlat(db);
  const lot = db.get('SELECT l.*, a.ano FROM lots l JOIN auctions a ON a.id=l.auction_id WHERE l.id=?', [req.params.lotId]);
  if (!lot) return res.status(404).json({ error: 'Lot not found' });
  const pdf = await generateCropReceiptPDF(lot, cfg);
  res.setHeader('Content-Type', 'application/pdf');
  res.setHeader('Content-Disposition', `inline; filename="Receipt_${lot.lot_no}.pdf"`);
  res.send(pdf);
});

// ── Lot-entry receipt PDF (for WhatsApp attachment) ──────────
// Pure renderer of the SAME slip the Lot Entry print modal shows. The
// client already builds the receipt payload (co / groups / dateStr /
// ano / printedAt / traderInfo) for the HTML print path; it POSTs that
// same payload here to get a PDF blob it can attach to a WhatsApp
// document message (the browser print flow can't yield a PDF blob).
// Available to lot-entry users since they own this screen.
app.post('/api/lot-receipt/pdf', requireViewOrLotEntry, async (req, res) => {
  try {
    const payload = req.body || {};
    if (!Array.isArray(payload.groups) || !payload.groups.length) {
      return res.status(400).json({ error: 'No lots to render' });
    }
    // Sensitive-field masking is decided server-side from company_settings
    // so a tampered/old client can't bypass it. The renderer masks the
    // seller account no. + IFSC in the slip's meta block accordingly.
    const _mCfg = getSettingsFlat(getDb());
    payload.maskCfg = {
      acct: _mCfg.mask_acct || 'none',
      ifsc: _mCfg.mask_ifsc || 'none',
      phone: _mCfg.mask_phone || 'none',
    };
    // Thermal paper width (Settings → Lot Entry Defaults). Decided
    // server-side so the WhatsApp/PDF slip matches the configured printer
    // roll regardless of what an old client sends. Blank/0 → renderer's
    // built-in default page size.
    const _wmm = Number(_mCfg.lot_receipt_width_mm || 0);
    if (_wmm > 0) payload.widthMm = _wmm;
    // Optional Sample Wt / Gross Wt columns on the detailed slip — decided
    // server-side from company_settings so an old/tampered client can't
    // override them. getSettingsFlat returns booleans for boolean fields.
    payload.showSample = _mCfg.lot_receipt_show_sample === true || _mCfg.lot_receipt_show_sample === 'true';
    payload.showGross  = _mCfg.lot_receipt_show_gross  === true || _mCfg.lot_receipt_show_gross  === 'true';
    const pdf = await generateLotReceiptPDF(payload);
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', 'inline; filename="LotReceipt.pdf"');
    res.send(pdf);
  } catch (e) {
    res.status(500).json({ error: e.message || 'Receipt PDF failed' });
  }
});

// ══════════════════════════════════════════════════════════════
// SUMMARY STATS
// ══════════════════════════════════════════════════════════════
app.get('/api/stats', requireView, (req, res) => {
  // Branch tiles + per-trade breakdown depend on data the user can edit
  // (Settings → Branches, lots, invoices). Disable HTTP caching outright
  // so the dashboard always reflects the current DB state — without
  // this header, browsers can serve a stale /api/stats response after
  // the user edits Settings → Branches & Contacts and the tiles
  // continue to show the previous list.
  res.set('Cache-Control', 'no-store');
  const db = getDb();

  // Counts
  const counts = {
    traders:    (db.get('SELECT COUNT(*) as c FROM traders') || {}).c || 0,
    buyers:     (db.get('SELECT COUNT(*) as c FROM buyers') || {}).c || 0,
    auctions:   (db.get('SELECT COUNT(*) as c FROM auctions') || {}).c || 0,
    lots:       (db.get('SELECT COUNT(*) as c FROM lots') || {}).c || 0,
    invoices:   (db.get('SELECT COUNT(*) as c FROM invoices') || {}).c || 0,
    purchases:  (db.get('SELECT COUNT(*) as c FROM purchases') || {}).c || 0,
    bills:      (db.get('SELECT COUNT(*) as c FROM bills') || {}).c || 0,
    debit_notes:(db.get('SELECT COUNT(*) as c FROM debit_notes') || {}).c || 0,
  };

  // All auctions (for the dashboard picker)
  const allAuctions = db.all(
    `SELECT id, ano, date, crop_type FROM auctions ORDER BY id DESC LIMIT 50`
  );

  // ── Cumulative totals across ALL trades (lifetime) ──
  // Aggregates over every lot in every auction. sold_qty / wd_qty slice
  // the same total — sold = lots with code present and not 'WD',
  // withdrawn = code = 'WD'. Lots with empty code (unsold) are in
  // neither slice but still counted in the qty total.
  // Min/Max/Avg derive from SOLD lots only (price > 0, amount > 0):
  //   - Min/Max use the bounding sold-lot prices
  //   - Avg is weighted: Σ amount ÷ Σ sold_qty
  // Withdrawn lots have no transacted price so their inclusion would
  // skew bounds and averages.
  const cumRow = db.get(
    `SELECT COALESCE(SUM(qty),0) as qty,
            COALESCE(SUM(amount),0) as amount,
            COUNT(*) as lots,
            COALESCE(SUM(CASE WHEN UPPER(TRIM(COALESCE(code,''))) NOT IN ('','WD') THEN qty ELSE 0 END),0) as sold_qty,
            COALESCE(SUM(CASE WHEN UPPER(TRIM(COALESCE(code,'')))  =  'WD'        THEN qty ELSE 0 END),0) as wd_qty,
            COALESCE(SUM(CASE WHEN UPPER(TRIM(COALESCE(code,''))) NOT IN ('','WD') AND amount > 0 THEN amount ELSE 0 END),0) as sold_amount,
            (SELECT COALESCE(MIN(price),0) FROM lots WHERE price > 0 AND amount > 0
              AND UPPER(TRIM(COALESCE(code,''))) NOT IN ('','WD')) as min_price,
            (SELECT COALESCE(MAX(price),0) FROM lots WHERE price > 0 AND amount > 0
              AND UPPER(TRIM(COALESCE(code,''))) NOT IN ('','WD')) as max_price
     FROM lots`
  ) || {};
  const cumSoldQty = Number(cumRow.sold_qty) || 0;
  const cumSoldAmt = Number(cumRow.sold_amount) || 0;
  const cumulative = {
    qty:       cumRow.qty       || 0,
    amount:    cumRow.amount    || 0,
    lots:      cumRow.lots      || 0,
    auctions:  counts.auctions,
    sold_qty:  cumSoldQty,
    wd_qty:    cumRow.wd_qty    || 0,
    min_price: Number(cumRow.min_price) || 0,
    max_price: Number(cumRow.max_price) || 0,
    avg_price: cumSoldQty > 0 ? round2(cumSoldAmt / cumSoldQty) : 0,
  };

  // ── Per-trade breakdown (one row per auction, newest first) ──
  // One query with a LEFT JOIN so auctions with zero lots still appear.
  // sold_qty / wd_qty slice the qty total by lot code:
  //   sold      = code present AND not 'WD'
  //   withdrawn = code = 'WD'
  // Min/Max/Avg derive from SOLD lots only (excludes WD + unsold) —
  // withdrawn lots have no transacted price, including them would
  // distort the bounds and the weighted average. Avg is computed in
  // JS after the SQL aggregate so the divisor (sold_qty) is the same
  // value the row reports — keeps the math auditable.
  const perTradeBreakdown = db.all(
    `SELECT a.id, a.ano, a.date, a.crop_type,
            COUNT(l.id) as lots,
            COALESCE(SUM(l.qty),0) as qty,
            COALESCE(SUM(CASE WHEN UPPER(TRIM(COALESCE(l.code,''))) NOT IN ('','WD') THEN l.qty ELSE 0 END),0) as sold_qty,
            COALESCE(SUM(CASE WHEN UPPER(TRIM(COALESCE(l.code,'')))  =  'WD'        THEN l.qty ELSE 0 END),0) as wd_qty,
            COALESCE(SUM(l.amount),0) as amount,
            COALESCE(SUM(CASE WHEN UPPER(TRIM(COALESCE(l.code,''))) NOT IN ('','WD') AND l.amount > 0 THEN l.amount ELSE 0 END),0) as sold_amount,
            COALESCE(MIN(CASE WHEN l.price > 0 AND l.amount > 0
                               AND UPPER(TRIM(COALESCE(l.code,''))) NOT IN ('','WD')
                              THEN l.price END),0) as min_price,
            COALESCE(MAX(CASE WHEN l.price > 0 AND l.amount > 0
                               AND UPPER(TRIM(COALESCE(l.code,''))) NOT IN ('','WD')
                              THEN l.price END),0) as max_price,
            COALESCE(SUM(CASE WHEN l.amount > 0 THEN 1 ELSE 0 END),0) as priced,
            COALESCE(SUM(CASE WHEN l.invo IS NOT NULL AND l.invo != '' THEN 1 ELSE 0 END),0) as invoiced
     FROM auctions a
     LEFT JOIN lots l ON l.auction_id = a.id
     GROUP BY a.id, a.ano, a.date, a.crop_type
     ORDER BY a.date DESC, a.id DESC
     LIMIT 50`
  ).map(r => {
    const soldQty = Number(r.sold_qty) || 0;
    const soldAmt = Number(r.sold_amount) || 0;
    return { ...r, avg_price: soldQty > 0 ? round2(soldAmt / soldQty) : 0 };
  });

  // Pick: ?auction_id=N if provided
  //   - "all" (or no param) => dashboard shows cumulative view, no individual auction highlighted
  //   - specific id         => dashboard drills into that one auction
  let currentAuction = null;
  const rawAuctionId = req.query.auction_id;
  const isAllMode = (rawAuctionId === 'all' || rawAuctionId === '' || rawAuctionId === undefined);
  if (!isAllMode) {
    const requestedId = parseInt(rawAuctionId);
    if (requestedId) {
      currentAuction = db.get('SELECT * FROM auctions WHERE id = ?', [requestedId]);
    }
  }

  let auctionStats = null;
  if (currentAuction) {
    const totalLots  = (db.get('SELECT COUNT(*) as c FROM lots WHERE auction_id = ?', [currentAuction.id]) || {}).c || 0;
    const priced     = (db.get('SELECT COUNT(*) as c FROM lots WHERE auction_id = ? AND amount > 0', [currentAuction.id]) || {}).c || 0;
    const invoiced   = (db.get(`SELECT COUNT(*) as c FROM lots WHERE auction_id = ? AND invo IS NOT NULL AND invo != ''`, [currentAuction.id]) || {}).c || 0;
    const totalQty   = (db.get('SELECT COALESCE(SUM(qty),0) as s FROM lots WHERE auction_id = ?', [currentAuction.id]) || {}).s || 0;
    const totalAmt   = (db.get('SELECT COALESCE(SUM(amount),0) as s FROM lots WHERE auction_id = ?', [currentAuction.id]) || {}).s || 0;
    // Document-generation progress for the dashboard hero's mini bars.
    // Each workflow targets a different population, so the denominators
    // differ — they mirror the same registered/URD split and
    // remaining-party predicates as _hasRemainingParties() so the bars
    // agree with the generate-lock logic:
    //   • payments  — every priced (non-WD) lot must be paid (lots.paid set)
    //   • purchases — one purchase invoice per REGISTERED seller (GSTIN cr)
    //   • bills     — one bill of supply per UNREGISTERED / agri seller
    const REG_CR = `(UPPER(l.cr) LIKE 'GSTIN%' OR (l.cr GLOB '[0-9][0-9]*' AND LENGTH(l.cr) >= 15))`;
    const URD_CR = `(l.cr IS NULL OR l.cr = '' OR (UPPER(l.cr) NOT LIKE 'GSTIN%' AND l.cr NOT GLOB '[0-9][0-9]*'))`;
    const NOT_WD = `UPPER(TRIM(COALESCE(l.code,''))) != 'WD'`;
    const cnt1 = (sql) => (db.get(sql, [currentAuction.id]) || {}).c || 0;
    const paymentsTotal  = cnt1(`SELECT COUNT(*) AS c FROM lots l WHERE l.auction_id = ? AND l.amount > 0 AND ${NOT_WD}`);
    const paymentsDone   = cnt1(`SELECT COUNT(*) AS c FROM lots l WHERE l.auction_id = ? AND l.amount > 0 AND ${NOT_WD} AND l.paid IS NOT NULL AND l.paid != ''`);
    const purchasesTotal = cnt1(`SELECT COUNT(DISTINCT l.name) AS c FROM lots l WHERE l.auction_id = ? AND l.amount > 0 AND ${NOT_WD} AND l.name IS NOT NULL AND l.name != '' AND ${REG_CR}`);
    const purchasesDone  = cnt1(`SELECT COUNT(DISTINCT l.name) AS c FROM lots l WHERE l.auction_id = ? AND l.amount > 0 AND ${NOT_WD} AND l.name IS NOT NULL AND l.name != '' AND ${REG_CR} AND EXISTS (SELECT 1 FROM purchases p WHERE p.auction_id = l.auction_id AND p.name = l.name)`);
    const billsTotal     = cnt1(`SELECT COUNT(DISTINCT l.name) AS c FROM lots l WHERE l.auction_id = ? AND l.amount > 0 AND ${NOT_WD} AND l.name IS NOT NULL AND l.name != '' AND ${URD_CR}`);
    const billsDone      = cnt1(`SELECT COUNT(DISTINCT l.name) AS c FROM lots l WHERE l.auction_id = ? AND l.amount > 0 AND ${NOT_WD} AND l.name IS NOT NULL AND l.name != '' AND ${URD_CR} AND EXISTS (SELECT 1 FROM bills b WHERE b.auction_id = l.auction_id AND b.name = l.name)`);
    auctionStats = { ...currentAuction, totalLots, priced, invoiced, totalQty, totalAmt,
      paymentsDone, paymentsTotal, purchasesDone, purchasesTotal, billsDone, billsTotal };
  }

  // Top sellers (this week — by total amount in auctions dated within last 7 days)
  const topSellers = db.all(
    `SELECT l.name as name, COUNT(*) as lots, COALESCE(SUM(l.qty),0) as qty, COALESCE(SUM(l.amount),0) as amount
     FROM lots l JOIN auctions a ON a.id = l.auction_id
     WHERE a.date >= date('now','-7 days') AND l.name IS NOT NULL AND l.name != ''
     GROUP BY l.name
     ORDER BY amount DESC
     LIMIT 5`
  );

  // Recent invoices (last 5)
  const recentInvoices = db.all(
    `SELECT i.id, i.sale, i.invo, i.buyer, i.buyer1, i.tot, i.date,
            i.place
     FROM invoices i
     ORDER BY i.id DESC LIMIT 5`
  );

  // Today's trade totals (active auction lots)
  const todayQty = auctionStats ? auctionStats.totalQty : 0;
  const todayAmt = auctionStats ? auctionStats.totalAmt : 0;

  // Revenue this month (sum of invoice totals in current month)
  const monthTot = (db.get(
    `SELECT COALESCE(SUM(tot),0) as s FROM invoices
     WHERE date >= date('now','start of month')`
  ) || {}).s || 0;
  // Revenue last month (for comparison)
  const lastMonthTot = (db.get(
    `SELECT COALESCE(SUM(tot),0) as s FROM invoices
     WHERE date >= date('now','start of month','-1 month')
       AND date <  date('now','start of month')`
  ) || {}).s || 0;

  // Pending invoices:
  //   - Drilled into an auction: un-invoiced priced lots in that auction
  //   - Cumulative mode: un-invoiced priced lots across ALL auctions
  let pendingInvoices = 0;
  if (currentAuction) {
    pendingInvoices = (db.get(
      `SELECT COUNT(DISTINCT buyer) as c FROM lots
       WHERE auction_id = ? AND amount > 0 AND buyer IS NOT NULL AND buyer != ''
         AND (invo IS NULL OR invo = '')`, [currentAuction.id]
    ) || {}).c || 0;
  } else {
    pendingInvoices = (db.get(
      `SELECT COUNT(DISTINCT buyer || '|' || auction_id) as c FROM lots
       WHERE amount > 0 AND buyer IS NOT NULL AND buyer != ''
         AND (invo IS NULL OR invo = '')`
    ) || {}).c || 0;
  }

  // ── Per-branch breakdown ──
  // The branch list is anchored to the user's configured Settings →
  // Branches & Contacts (br1..br9), so the dashboard tiles are stable
  // and reflect the user's organizational structure — not whatever
  // string appeared in `lots.branch` for a given auction.
  //
  // Lots are aggregated per configured branch (case-insensitive match)
  // across either ALL trades (cumulative mode) or just the currently-
  // selected trade. Configured branches with zero matching lots still
  // appear as 0-tiles so the user sees their full org laid out.
  // Lots whose `branch` value doesn't match any configured branch get
  // bucketed under "(unspecified)" — surfaced only when present, so it
  // doesn't clutter the dashboard for clean datasets.
  const cfgBranches = [];
  for (let i = 1; i <= 9; i++) {
    const r = db.get('SELECT value FROM company_settings WHERE key = ?', [`br${i}`]);
    const v = r && r.value ? String(r.value).trim() : '';
    if (v) cfgBranches.push(v);
  }
  const aggSql = currentAuction
    ? `SELECT COALESCE(TRIM(branch), '') AS branch,
              COUNT(*) AS lots,
              COALESCE(SUM(qty), 0) AS qty,
              COALESCE(SUM(amount), 0) AS amount
         FROM lots
        WHERE auction_id = ?
        GROUP BY UPPER(COALESCE(TRIM(branch), ''))`
    : `SELECT COALESCE(TRIM(branch), '') AS branch,
              COUNT(*) AS lots,
              COALESCE(SUM(qty), 0) AS qty,
              COALESCE(SUM(amount), 0) AS amount
         FROM lots
        GROUP BY UPPER(COALESCE(TRIM(branch), ''))`;
  const rawAgg = currentAuction ? db.all(aggSql, [currentAuction.id]) : db.all(aggSql);
  // Index aggregates by uppercased branch name for case-insensitive match.
  const aggIdx = {};
  for (const a of rawAgg) {
    const k = String(a.branch || '').toUpperCase();
    aggIdx[k] = a;
  }
  // Walk configured branches first (preserves Settings order so tiles
  // line up with how the user thinks about their business). Each config
  // branch always emits a tile, even when 0 lots — empty tiles signal
  // "no activity here" which is itself useful information for the user.
  const branchTotals = cfgBranches.map(name => {
    const hit = aggIdx[name.toUpperCase()];
    return {
      branch: name,
      lots:   Number((hit && hit.lots)   || 0),
      qty:    Number((hit && hit.qty)    || 0),
      amount: Number((hit && hit.amount) || 0),
      configured: true,
    };
  });
  // Strays — lots whose `branch` value doesn't match any configured
  // branch from Settings — are intentionally NOT surfaced on the
  // dashboard per spec. The dashboard reflects ONLY the configured
  // organization (Settings → Branches & Contacts). Stray lots still
  // exist in the DB and remain visible in the Lots tab; users clean
  // them up by editing those lots or by adding the missing branch to
  // Settings, which then makes the lots count toward that tile.
  // (Earlier we surfaced strays as orange "stray" tiles + an
  // "(unspecified)" bucket — that confused users into thinking the
  // dashboard was broken when really their lot data had typos.)

  res.json({
    counts,
    cumulative,
    perTradeBreakdown,
    branchTotals,
    currentAuction: auctionStats,
    allAuctions,
    topSellers,
    recentInvoices,
    kpi: {
      todayQty, todayAmt,
      activeLots: auctionStats ? auctionStats.totalLots : 0,
      pendingInvoices,
      monthRevenue: monthTot,
      lastMonthRevenue: lastMonthTot,
    }
  });
});

// ──────────────────────────────────────────────────────────────
// Dashboard: revenue trend (line chart data)
// ──────────────────────────────────────────────────────────────
// Returns daily invoice totals for the last N days (default 7).
// Used by the dashboard line chart card. Days with zero invoices
// still appear in the result with `total: 0` so the X-axis is
// continuous — easier to read than a jagged sparse line. Counted
// from invoices.tot which already includes GST + round; the chart
// shows what the customer actually billed each day.
//
// The cache-busting header matches /api/stats — the dashboard
// reloads this on every visit and must reflect freshly-saved
// invoices without browser-cached responses.
app.get('/api/stats/revenue-trend', requireView, (req, res) => {
  res.set('Cache-Control', 'no-store');
  // Clamp days to 1–90; the chart space only fits ~30 cleanly,
  // but 90 is allowed for power users who want a longer view.
  // `parseInt` returns NaN for non-numeric → use 7 as the default;
  // a valid 0 should clamp UP to 1, so Number.isFinite check first.
  const raw = parseInt(req.query.days, 10);
  const days = Math.max(1, Math.min(90, Number.isFinite(raw) ? raw : 7));
  const db = getDb();
  // SQLite: aggregate by date(date) so a stray time-component
  // doesn't break the group. COALESCE handles NULL totals.
  const rows = db.all(
    `SELECT date(date) as day, COALESCE(SUM(tot), 0) as total, COUNT(*) as count
       FROM invoices
      WHERE date IS NOT NULL AND date != ''
        AND date(date) >= date('now', '-' || ? || ' days')
        AND date(date) <= date('now')
      GROUP BY date(date)
      ORDER BY day ASC`,
    [days - 1]
  );
  // Build a complete day series (zero-fill missing days). The DB
  // result is sparse — if no invoices on a day, it's missing. We
  // emit a flat array of exactly `days` entries so the client
  // doesn't have to do date math.
  const byDay = new Map(rows.map(r => [r.day, r]));
  const series = [];
  const today = new Date();
  for (let i = days - 1; i >= 0; i--) {
    const d = new Date(today);
    d.setDate(d.getDate() - i);
    const iso = d.toISOString().slice(0, 10);
    const hit = byDay.get(iso);
    series.push({
      date: iso,
      total: hit ? Number(hit.total) || 0 : 0,
      count: hit ? Number(hit.count) || 0 : 0,
    });
  }
  res.json({ days, series });
});

// ══════════════════════════════════════════════════════════════
// INSIGHTS — per-trade × per-branch analytics
// ══════════════════════════════════════════════════════════════
//
// Powers the Insights tab. Single endpoint that returns everything
// the page renders so the UI is one fetch + one render instead of
// N coordinated calls. Date range defaults to the current calendar
// month; clamped to a sane window so a runaway query string can't
// pull a multi-year scan.
//
// Returns:
//   range            — {from, to} actually used (after clamping)
//   totals           — overall KPIs across the range
//   perTrade         — one row per auction in range, with branch-level
//                      sub-breakdown baked in (saves a second round
//                      trip when the user expands a trade row)
//   perBranch        — branch-level rollup across all trades in range,
//                      including SUM(lots.balance) as payable_to_sellers
//   branchStacked    — pre-shaped {labels, datasets} for a stacked
//                      Chart.js bar (each trade is a bar, each branch
//                      a stack segment).
//   outstandingByBuyer — sales totals per buyer in the range. Without
//                      a payment-tracking table the totals double as
//                      "outstanding" (every invoice is treated as
//                      unpaid); the column label in the UI makes that
//                      assumption explicit.
//   buyerActivity    — top buyers by value across the range with lots
//                      / qty / value (mirrors the existing top-sellers
//                      leaderboard on the regular dashboard, but on
//                      the buy-side).
//
// All SUM/aggregation happens in SQL where it can (one pass per scope)
// and the final pivot for the stacked bar happens in JS so the
// branches-on-X-axis logic stays readable.
app.get('/api/insights', requireView, (req, res) => {
  res.set('Cache-Control', 'no-store');
  const db = getDb();

  // Two modes:
  //   1. ?auction_id=N   → filter to a single trade. Date range is
  //                        inferred from that auction's `date` so every
  //                        downstream query continues to use the same
  //                        date predicate without special-casing.
  //   2. ?from=... &to=. → calendar window (default: current month).
  // Used by the Insights tab (date range) AND the Dashboard's headline
  // metrics tiles (auction_id, when the operator picks a specific trade).
  const today = new Date();
  const pad = n => String(n).padStart(2, '0');
  const isoMonthStart = `${today.getFullYear()}-${pad(today.getMonth() + 1)}-01`;
  const isoToday      = `${today.getFullYear()}-${pad(today.getMonth() + 1)}-${pad(today.getDate())}`;
  const dateRe = /^\d{4}-\d{2}-\d{2}$/;
  let from, to;
  let singleAuctionId = null;
  const rawAid = String(req.query.auction_id || '').trim();
  // Lifetime mode — the Dashboard's headline card asks for "All trades"
  // via auction_id=all. Span the widest possible date range so every
  // trade is included; the per-query BETWEEN predicates then aggregate
  // across all auctions. (The Insights tab never uses this — it always
  // passes explicit from/to.)
  const allTrades = rawAid === 'all';
  if (rawAid && !allTrades) {
    const n = parseInt(rawAid, 10);
    if (Number.isFinite(n) && n > 0) {
      const a = db.get('SELECT id, date FROM auctions WHERE id = ?', [n]);
      if (a && a.date) {
        singleAuctionId = a.id;
        from = String(a.date).slice(0, 10);
        to   = from;
      }
    }
  }
  if (from == null) {
    if (allTrades) {
      from = '0000-01-01';
      to   = '9999-12-31';
    } else {
      from = dateRe.test(String(req.query.from || '')) ? String(req.query.from) : isoMonthStart;
      to   = dateRe.test(String(req.query.to   || '')) ? String(req.query.to)   : isoToday;
    }
  }

  // Helper SQL fragment — "is this lot actually sold?" Code blank
  // means no buyer was assigned; "WD" means withdrawn. Anything else
  // is a real hammer transaction and counts towards min/max/avg price.
  const SOLD = `(UPPER(COALESCE(l.code,'')) <> '' AND UPPER(COALESCE(l.code,'')) <> 'WD')`;
  const WD   = `(UPPER(COALESCE(l.code,'')) = 'WD')`;

  // Auction filter — appended to every WHERE that references `a` or `l`.
  // Two trades can fall on the same calendar date, so the date predicate
  // alone is not enough when the caller asks for a specific auction.
  const aidA = singleAuctionId ? ` AND a.id = ${Number(singleAuctionId)}` : '';
  const aidL = singleAuctionId ? ` AND l.auction_id = ${Number(singleAuctionId)}` : '';

  // ── Per-trade rollup ──────────────────────────────────────
  const tradeRows = db.all(
    `SELECT
       a.id AS auction_id, a.ano, a.date, COALESCE(a.state, '') AS state,
       COUNT(l.id) AS lots,
       SUM(CASE WHEN ${SOLD} THEN 1 ELSE 0 END) AS sold,
       SUM(CASE WHEN ${WD}   THEN 1 ELSE 0 END) AS withdrawn,
       SUM(CASE WHEN COALESCE(l.code,'') = '' THEN 1 ELSE 0 END) AS unsold,
       COALESCE(SUM(l.qty), 0) AS qty,
       COALESCE(SUM(l.bags), 0) AS bags,
       MIN(CASE WHEN l.price > 0 AND ${SOLD} THEN l.price END) AS min_price,
       MAX(CASE WHEN l.price > 0 AND ${SOLD} THEN l.price END) AS max_price,
       COALESCE(SUM(CASE WHEN ${SOLD} THEN l.amount ELSE 0 END), 0) AS sold_value,
       COALESCE(SUM(CASE WHEN ${SOLD} THEN l.qty    ELSE 0 END), 0) AS sold_qty,
       COALESCE(SUM(CASE WHEN ${SOLD} THEN l.bags   ELSE 0 END), 0) AS sold_bags,
       COALESCE(SUM(CASE WHEN ${WD}   THEN l.amount ELSE 0 END), 0) AS wd_value,
       COALESCE(SUM(CASE WHEN ${WD}   THEN l.qty    ELSE 0 END), 0) AS wd_qty,
       COALESCE(SUM(CASE WHEN ${WD}   THEN l.bags   ELSE 0 END), 0) AS wd_bags,
       COALESCE(SUM(l.amount), 0) AS value
     FROM auctions a
     LEFT JOIN lots l ON l.auction_id = a.id
     WHERE date(a.date) BETWEEN date(?) AND date(?)${aidA}
     GROUP BY a.id, a.ano, a.date, a.state
     ORDER BY a.date DESC, CAST(a.ano AS INTEGER) DESC, a.ano DESC`,
    [from, to]
  );

  // ── Per-trade × per-branch breakdown — fetched in one pass and
  // bucketed onto the parent trade so the UI can expand a trade
  // row to see its branch contributions without another fetch.
  const tradeBranchRows = db.all(
    `SELECT
       l.auction_id,
       COALESCE(NULLIF(TRIM(l.branch),''), '—') AS branch,
       COUNT(l.id) AS lots,
       COALESCE(SUM(l.qty), 0) AS qty,
       COALESCE(SUM(l.amount), 0) AS value
     FROM lots l
     JOIN auctions a ON a.id = l.auction_id
     WHERE date(a.date) BETWEEN date(?) AND date(?)${aidA}
     GROUP BY l.auction_id, COALESCE(NULLIF(TRIM(l.branch),''), '—')`,
    [from, to]
  );
  const branchesByTrade = new Map();
  for (const r of tradeBranchRows) {
    if (!branchesByTrade.has(r.auction_id)) branchesByTrade.set(r.auction_id, []);
    branchesByTrade.get(r.auction_id).push({
      branch: r.branch,
      lots:   Number(r.lots) || 0,
      qty:    Number(r.qty)  || 0,
      value:  Number(r.value)|| 0,
    });
  }
  const perTrade = tradeRows.map(t => ({
    auction_id: t.auction_id,
    ano:   t.ano,
    date:  t.date,
    state: t.state,
    lots:      Number(t.lots)      || 0,
    sold:      Number(t.sold)      || 0,
    withdrawn: Number(t.withdrawn) || 0,
    unsold:    Number(t.unsold)    || 0,
    qty:       Number(t.qty)       || 0,
    bags:      Number(t.bags)      || 0,
    min_price: t.min_price == null ? null : Number(t.min_price),
    max_price: t.max_price == null ? null : Number(t.max_price),
    // sold_* / wd_* cover SOLD vs WITHDRAWN lots only; kept on the row so
    // the window-wide headline can aggregate them (avg, bags/qty/amount
    // breakdowns per category). sold_value/sold_qty also drive avg_price.
    sold_value: Number(t.sold_value) || 0,
    sold_qty:   Number(t.sold_qty)   || 0,
    sold_bags:  Number(t.sold_bags)  || 0,
    wd_value:   Number(t.wd_value)   || 0,
    wd_qty:     Number(t.wd_qty)     || 0,
    wd_bags:    Number(t.wd_bags)    || 0,
    avg_price: Number(t.sold_qty) > 0 ? Number(t.sold_value) / Number(t.sold_qty) : 0,
    value:     Number(t.value)     || 0,
    branches: (branchesByTrade.get(t.auction_id) || []).sort((a, b) => b.value - a.value),
  }));

  // ── Per-branch rollup across the entire window ────────────
  const branchRows = db.all(
    `SELECT
       COALESCE(NULLIF(TRIM(l.branch),''), '—') AS branch,
       COUNT(l.id) AS lots,
       SUM(CASE WHEN ${SOLD} THEN 1 ELSE 0 END) AS sold,
       SUM(CASE WHEN ${WD}   THEN 1 ELSE 0 END) AS withdrawn,
       SUM(CASE WHEN COALESCE(l.code,'') = '' THEN 1 ELSE 0 END) AS unsold,
       COALESCE(SUM(l.qty), 0) AS qty,
       MIN(CASE WHEN l.price > 0 AND ${SOLD} THEN l.price END) AS min_price,
       MAX(CASE WHEN l.price > 0 AND ${SOLD} THEN l.price END) AS max_price,
       COALESCE(SUM(CASE WHEN ${SOLD} THEN l.amount ELSE 0 END), 0) AS sold_value,
       COALESCE(SUM(CASE WHEN ${SOLD} THEN l.qty    ELSE 0 END), 0) AS sold_qty,
       COALESCE(SUM(l.amount), 0) AS value,
       COALESCE(SUM(l.balance), 0) AS payable_to_sellers
     FROM lots l
     JOIN auctions a ON a.id = l.auction_id
     WHERE date(a.date) BETWEEN date(?) AND date(?)${aidA}
     GROUP BY COALESCE(NULLIF(TRIM(l.branch),''), '—')
     ORDER BY value DESC`,
    [from, to]
  );
  const perBranch = branchRows.map(r => ({
    branch:    r.branch,
    lots:      Number(r.lots)      || 0,
    sold:      Number(r.sold)      || 0,
    withdrawn: Number(r.withdrawn) || 0,
    unsold:    Number(r.unsold)    || 0,
    qty:       Number(r.qty)       || 0,
    min_price: r.min_price == null ? null : Number(r.min_price),
    max_price: r.max_price == null ? null : Number(r.max_price),
    avg_price: Number(r.sold_qty) > 0 ? Number(r.sold_value) / Number(r.sold_qty) : 0,
    value:     Number(r.value)     || 0,
    payable_to_sellers: Number(r.payable_to_sellers) || 0,
  }));

  // ── Stacked bar series — pivot perTrade × perBranch into a
  // Chart.js-ready {labels, datasets} shape. Cap to the top 20
  // trades by date so the chart stays readable; the user can
  // narrow the date window to see more.
  const stackedTrades = perTrade.slice(0, 20);
  const branchUniverse = Array.from(new Set(
    stackedTrades.flatMap(t => t.branches.map(b => b.branch))
  )).sort();
  const branchStacked = {
    labels: stackedTrades.map(t => `${t.ano} · ${String(t.date || '').slice(5)}`),
    branches: branchUniverse,
    datasets: branchUniverse.map(name => ({
      label: name,
      data: stackedTrades.map(t => {
        const hit = t.branches.find(b => b.branch === name);
        return hit ? hit.value : 0;
      }),
    })),
  };

  // ── Outstanding by buyer — invoices in the window, summed per
  // buyer code. Without a payments table every invoice total is
  // treated as outstanding; the UI labels the column accordingly.
  const outstandingByBuyer = db.all(
    `SELECT
       COALESCE(NULLIF(TRIM(i.buyer1), ''), TRIM(i.buyer)) AS buyer_code,
       COALESCE(NULLIF(TRIM(i.buyer1), ''), TRIM(i.buyer)) AS buyer_name,
       COUNT(*) AS invoices,
       COALESCE(SUM(i.tot), 0) AS value
     FROM invoices i
     WHERE date(i.date) BETWEEN date(?) AND date(?)${singleAuctionId ? ` AND i.auction_id = ${Number(singleAuctionId)}` : ''}
     GROUP BY COALESCE(NULLIF(TRIM(i.buyer1), ''), TRIM(i.buyer))
     HAVING buyer_code IS NOT NULL AND buyer_code <> ''
     ORDER BY value DESC
     LIMIT 50`,
    [from, to]
  ).map(r => ({
    buyer_code: r.buyer_code,
    buyer_name: r.buyer_name,
    invoices: Number(r.invoices) || 0,
    value: Number(r.value) || 0,
  }));

  // ── Buyer activity leaderboard — same source as outstanding but
  // pulled from lots (one row per lot the buyer hammered) so we
  // can show lots + qty + value, not just invoice totals.
  const buyerActivity = db.all(
    `SELECT
       COALESCE(NULLIF(TRIM(l.buyer1), ''), TRIM(l.buyer)) AS buyer_name,
       COALESCE(NULLIF(TRIM(l.code),  ''), TRIM(l.buyer))  AS buyer_code,
       COUNT(l.id) AS lots,
       COALESCE(SUM(l.qty), 0) AS qty,
       COALESCE(SUM(l.amount), 0) AS value
     FROM lots l
     JOIN auctions a ON a.id = l.auction_id
     WHERE date(a.date) BETWEEN date(?) AND date(?)${aidA}
       AND ${SOLD}
     GROUP BY COALESCE(NULLIF(TRIM(l.buyer1), ''), TRIM(l.buyer))
     HAVING buyer_name IS NOT NULL AND buyer_name <> ''
     ORDER BY value DESC
     LIMIT 10`,
    [from, to]
  ).map(r => ({
    buyer_code: r.buyer_code,
    buyer_name: r.buyer_name,
    lots: Number(r.lots) || 0,
    qty:  Number(r.qty)  || 0,
    value: Number(r.value) || 0,
  }));

  // ── Overall KPIs — derived from perTrade so totals always agree
  // with the per-trade table the user is looking at.
  const totals = perTrade.reduce((a, t) => ({
    trades:     a.trades + 1,
    lots:       a.lots       + t.lots,
    sold:       a.sold       + t.sold,
    withdrawn:  a.withdrawn  + t.withdrawn,
    qty:        a.qty        + t.qty,
    bags:       a.bags       + t.bags,
    value:      a.value      + t.value,
    sold_value: a.sold_value + t.sold_value,
    sold_qty:   a.sold_qty   + t.sold_qty,
    sold_bags:  a.sold_bags  + t.sold_bags,
    wd_value:   a.wd_value   + t.wd_value,
    wd_qty:     a.wd_qty     + t.wd_qty,
    wd_bags:    a.wd_bags    + t.wd_bags,
  }), { trades: 0, lots: 0, sold: 0, withdrawn: 0, qty: 0, bags: 0, value: 0,
        sold_value: 0, sold_qty: 0, sold_bags: 0, wd_value: 0, wd_qty: 0, wd_bags: 0 });
  totals.payable_to_sellers = perBranch.reduce((s, b) => s + b.payable_to_sellers, 0);
  totals.outstanding_by_buyers = outstandingByBuyer.reduce((s, b) => s + b.value, 0);
  // Quantity-weighted avg price across SOLD lots only (sold_value / sold_qty),
  // matching the per-trade and per-branch Avg columns. Using all-lots qty here
  // would dilute the average with withdrawn/unsold kilos that carry no value.
  totals.avg_price = totals.sold_qty > 0 ? (totals.sold_value / totals.sold_qty) : 0;
  // Window-wide min/max sold price = MIN/MAX of each trade's min/max (ignoring
  // trades with no sold lots, where min_price/max_price are null).
  const _mins = perTrade.map(t => t.min_price).filter(v => v != null && v > 0);
  const _maxs = perTrade.map(t => t.max_price).filter(v => v != null && v > 0);
  totals.min_price = _mins.length ? Math.min(..._mins) : null;
  totals.max_price = _maxs.length ? Math.max(..._maxs) : null;

  res.json({
    range: { from, to },
    auction_id: singleAuctionId,
    totals,
    perTrade,
    perBranch,
    branchStacked,
    outstandingByBuyer,
    buyerActivity,
  });
});

// ══════════════════════════════════════════════════════════════
// START
// ══════════════════════════════════════════════════════════════
function repairBadDates(db) {
  // Fix rows where date is an Excel serial number stored as string, or Date-object-toString garbage
  const tables = ['auctions', 'bills', 'debit_notes', 'invoices', 'purchases', 'lots'];
  let totalFixed = 0;
  for (const tbl of tables) {
    try {
      // Only tables that have a `date` column
      const hasDate = db.all(`PRAGMA table_info(${tbl})`).some(c => c.name === 'date');
      if (!hasDate) continue;
      const rows = db.all(`SELECT rowid, date FROM ${tbl} WHERE date IS NOT NULL AND date != ''`);
      let fixed = 0;
      for (const r of rows) {
        const current = String(r.date);
        // Skip if already ISO yyyy-mm-dd
        if (/^\d{4}-\d{2}-\d{2}$/.test(current)) continue;
        const iso = normalizeDate(r.date);
        if (iso && /^\d{4}-\d{2}-\d{2}$/.test(iso) && iso !== current) {
          db.run(`UPDATE ${tbl} SET date = ? WHERE rowid = ?`, [iso, r.rowid]);
          fixed++;
        }
      }
      if (fixed > 0) console.log(`  Date repair: ${tbl} — fixed ${fixed} row(s)`);
      totalFixed += fixed;
    } catch (_) { /* table may not exist yet */ }
  }
  if (totalFixed > 0) console.log(`  Date repair: ${totalFixed} total row(s) normalized to yyyy-mm-dd`);
}

// ══════════════════════════════════════════════════════════════
// SYSTEM: DB backup & restore (admin-only)
// ══════════════════════════════════════════════════════════════
// Backup: streams the live config.db file as a download. We force a
// flush first so any pending in-memory writes (sql.js debounces saves
// 200ms) are on disk before we read the file.
// List the on-disk auto-backup snapshots (admin-only). Returns name +
// size + mtime so the Settings UI can show "Last backup: …" and a
// rolling list. Empty array when auto-backup has never fired.
app.get('/api/system/backups', requireAdmin, (req, res) => {
  try {
    const bkDir = path.join(path.dirname(DB_PATH), 'backups');
    if (!fs.existsSync(bkDir)) return res.json({ backups: [] });
    const out = fs.readdirSync(bkDir)
      .filter(f => f.endsWith('.db'))
      .map(f => {
        const st = fs.statSync(path.join(bkDir, f));
        return { name: f, size: st.size, mtime: st.mtimeMs };
      })
      .sort((a, b) => b.mtime - a.mtime);
    res.json({ backups: out });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// Trigger a backup snapshot now (admin-only). Mirrors the auto-ticker
// path; useful for "test the schedule" or one-off snapshots.
app.post('/api/system/backup-now', requireAdmin, (req, res) => {
  try {
    runBackupTickerOnce.__forceForOneCall = true;  // unused — keep for future
    const bkDir = path.join(path.dirname(DB_PATH), 'backups');
    if (!fs.existsSync(bkDir)) fs.mkdirSync(bkDir, { recursive: true });
    require('./db').flushSave();
    const stamp = new Date().toISOString().slice(0, 19).replace(/[:T]/g, '-');
    const out = path.join(bkDir, `manual-${stamp}.db`);
    fs.copyFileSync(DB_PATH, out);
    res.json({ success: true, file: path.basename(out) });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.get('/api/system/backup', requireAdmin, (req, res) => {
  try {
    // Force-flush pending writes so the on-disk file reflects every
    // committed transaction up to now.
    require('./db').flushSave();
    if (!fs.existsSync(DB_PATH)) {
      return res.status(500).json({ error: 'Database file not found at ' + DB_PATH });
    }
    const stamp = new Date().toISOString().slice(0, 19).replace(/[:T]/g, '-');
    const filename = `spice-etrade-backup-${stamp}.db`;
    res.setHeader('Content-Type', 'application/octet-stream');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    fs.createReadStream(DB_PATH).pipe(res);
  } catch (e) {
    console.error('[backup] failed:', e);
    res.status(500).json({ error: e.message });
  }
});

// Restore: accept a multipart upload of a .db file, validate, swap.
// We use a separate multer instance with a 200MB limit (the global
// `upload` is capped at 20MB which would block legitimate backups
// once the DB grows). The uploaded file lands in `uploadDir` and is
// removed by replaceFromBuffer's flow regardless of success/failure.
const restoreUpload = multer({ dest: uploadDir, limits: { fileSize: 200 * 1024 * 1024 } });
app.post('/api/system/restore', requireAdmin, restoreUpload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded (field name: file)' });
  const tmpPath = req.file.path;
  try {
    const buf = fs.readFileSync(tmpPath);
    const r = await replaceFromBuffer(buf);
    res.json({ ok: true, restoredBytes: r.size });
  } catch (e) {
    console.error('[restore] failed:', e);
    res.status(400).json({ error: e.message });
  } finally {
    try { fs.unlinkSync(tmpPath); } catch(_) {}
  }
});

// ══════════════════════════════════════════════════════════════
// DEV DATA EDITOR — in-app DB Browser (admin-only, dev-gated UI)
// ══════════════════════════════════════════════════════════════
// Lets a developer view/edit/insert/delete rows of ANY table and run
// raw SQL straight from the app, replacing the download-backup →
// DB-Browser → restore round-trip. The UI is revealed only by the
// `spiceDev(true)` console flag (see public/index.html), but these
// endpoints additionally require an admin session — so the data can't
// be touched without admin auth even if the flag is discovered.
//
// SAFETY: every mutating call (insert/update/delete/non-SELECT SQL)
// first takes an automatic `predev-*.db` snapshot into the backups
// folder, so any mistake is one restore away.

// Take a pre-edit snapshot of the live DB. Mirrors /api/system/backup-now.
function _devSnapshot(tag) {
  const bkDir = path.join(path.dirname(DB_PATH), 'backups');
  if (!fs.existsSync(bkDir)) fs.mkdirSync(bkDir, { recursive: true });
  require('./db').flushSave();
  const stamp = new Date().toISOString().slice(0, 19).replace(/[:T]/g, '-');
  const out = path.join(bkDir, `predev-${tag}-${stamp}.db`);
  fs.copyFileSync(DB_PATH, out);
  return path.basename(out);
}

// Guard: confirm `name` is a real, plainly-named base table (not a view,
// not sqlite internal, no funny characters). Returns true/false. We both
// check sqlite_master AND enforce a safe identifier charset because the
// name is interpolated into SQL (parameters can't bind identifiers).
function _devIsTable(db, name) {
  if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(name || '')) return false;
  return !!db.get(
    "SELECT 1 FROM sqlite_master WHERE type='table' AND name = ?", [name]);
}

// List every base table with its row count. Drives the left-hand picker.
app.get('/api/dev/db/tables', requireAdmin, (req, res) => {
  try {
    const db = getDb();
    const tables = db.all(
      "SELECT name FROM sqlite_master WHERE type='table' " +
      "AND name NOT LIKE 'sqlite_%' ORDER BY name");
    const out = tables.map(t => {
      let count = null;
      try { count = db.get(`SELECT COUNT(*) c FROM "${t.name}"`).c; } catch (_) {}
      return { name: t.name, count };
    });
    res.json({ tables: out });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// Read one table: column metadata + a page of rows. Every row carries a
// stable `_rowid_` (SQLite's implicit rowid) so edits/deletes target an
// exact row even when the table has no obvious primary key. `q` does a
// substring match across all columns (cast to text).
app.get('/api/dev/db/table/:name', requireAdmin, (req, res) => {
  try {
    const db = getDb();
    const name = req.params.name;
    if (!_devIsTable(db, name)) return res.status(404).json({ error: 'Unknown table' });
    const cols = db.all(`PRAGMA table_info("${name}")`);
    const limit = Math.min(parseInt(req.query.limit, 10) || 100, 1000);
    const offset = parseInt(req.query.offset, 10) || 0;
    const q = (req.query.q || '').trim();
    let where = '', params = [];
    if (q) {
      where = 'WHERE ' + cols.map(c => `CAST("${c.name}" AS TEXT) LIKE ?`).join(' OR ');
      params = cols.map(() => `%${q}%`);
    }
    const total = db.get(`SELECT COUNT(*) c FROM "${name}" ${where}`, params).c;
    const rows = db.all(
      `SELECT rowid AS _rowid_, * FROM "${name}" ${where} ORDER BY rowid LIMIT ? OFFSET ?`,
      [...params, limit, offset]);
    res.json({ name, columns: cols, rows, total, limit, offset });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// Insert a row. Body: { values: { col: val, … } }. Unknown columns are
// ignored; '' / null pass through as given.
app.post('/api/dev/db/table/:name', requireAdmin, (req, res) => {
  try {
    const db = getDb();
    const name = req.params.name;
    if (!_devIsTable(db, name)) return res.status(404).json({ error: 'Unknown table' });
    const valid = new Set(db.all(`PRAGMA table_info("${name}")`).map(c => c.name));
    const vals = req.body.values || {};
    const keys = Object.keys(vals).filter(k => valid.has(k));
    if (!keys.length) return res.status(400).json({ error: 'No valid columns supplied' });
    _devSnapshot('insert');
    const info = db.run(
      `INSERT INTO "${name}" (${keys.map(k => `"${k}"`).join(',')}) ` +
      `VALUES (${keys.map(() => '?').join(',')})`,
      keys.map(k => vals[k]));
    res.json({ success: true, rowid: info.lastInsertRowid });
  } catch (e) { res.status(400).json({ error: e.message }); }
});

// Update one row by rowid. Body: { values: { col: val, … } }. Used for
// both single-cell saves and full-row edits.
app.put('/api/dev/db/table/:name/:rowid', requireAdmin, (req, res) => {
  try {
    const db = getDb();
    const name = req.params.name;
    if (!_devIsTable(db, name)) return res.status(404).json({ error: 'Unknown table' });
    const valid = new Set(db.all(`PRAGMA table_info("${name}")`).map(c => c.name));
    const vals = req.body.values || {};
    const keys = Object.keys(vals).filter(k => valid.has(k));
    if (!keys.length) return res.status(400).json({ error: 'No valid columns supplied' });
    _devSnapshot('update');
    const info = db.run(
      `UPDATE "${name}" SET ${keys.map(k => `"${k}"=?`).join(',')} WHERE rowid=?`,
      [...keys.map(k => vals[k]), parseInt(req.params.rowid, 10)]);
    res.json({ success: true, changes: info.changes });
  } catch (e) { res.status(400).json({ error: e.message }); }
});

// Delete one row by rowid.
app.delete('/api/dev/db/table/:name/:rowid', requireAdmin, (req, res) => {
  try {
    const db = getDb();
    const name = req.params.name;
    if (!_devIsTable(db, name)) return res.status(404).json({ error: 'Unknown table' });
    _devSnapshot('delete');
    const info = db.run(`DELETE FROM "${name}" WHERE rowid=?`, [parseInt(req.params.rowid, 10)]);
    res.json({ success: true, changes: info.changes });
  } catch (e) { res.status(400).json({ error: e.message }); }
});

// Raw SQL console. SELECT/PRAGMA/EXPLAIN/WITH return { columns, rows };
// anything else is treated as a mutation (snapshot first) and returns
// the affected-row count. Supports multi-statement scripts via exec().
app.post('/api/dev/db/sql', requireAdmin, (req, res) => {
  try {
    const db = getDb();
    const sql = String(req.body.sql || '').trim();
    if (!sql) return res.status(400).json({ error: 'Empty SQL' });
    const isRead = /^(select|pragma|explain|with)\b/i.test(sql);
    if (isRead) {
      const rows = db.all(sql);
      const columns = rows.length ? Object.keys(rows[0]) : [];
      return res.json({ select: true, columns, rows, total: rows.length });
    }
    _devSnapshot('sql');
    try {
      const info = db.run(sql);          // single statement → real change count
      res.json({ select: false, changes: info.changes });
    } catch (_) {
      db.exec(sql);                       // fall back for multi-statement scripts
      res.json({ select: false, changes: null });
    }
  } catch (e) { res.status(400).json({ error: e.message }); }
});

// ══════════════════════════════════════════════════════════════
// FIX DATA — gentle, no-SQL data editor for a non-technical user
// ══════════════════════════════════════════════════════════════
// A safe face over the same engine for the day-to-day end user (dad).
// The UI is dev-revealed (spiceDev flag) and admin-authenticated, but
// the real safety is here on the server: a hard whitelist means these
// endpoints can ONLY ever touch the 8 business tables, never users /
// sessions / license_state / password hashes — and never via raw SQL.
//
// Locked columns (primary key, foreign-key links, machine timestamps)
// can be READ but never WRITTEN, so a stray edit can't re-link a lot to
// the wrong auction or corrupt invoices/payments. Every write snapshots
// a backup first and is recorded in audit_log.

// Friendly key → { label, table }. Anything not listed is unreachable.
const DATA_ENTITIES = {
  sellers:     { label: 'Sellers',         table: 'traders' },
  buyers:      { label: 'Buyers',          table: 'buyers' },
  auctions:    { label: 'Auctions',        table: 'auctions' },
  lots:        { label: 'Lots',            table: 'lots' },
  invoices:    { label: 'Sales Invoices',  table: 'invoices' },
  purchases:   { label: 'Purchases',       table: 'purchases' },
  bills:       { label: 'Bills of Supply', table: 'bills' },
  debit_notes: { label: 'Debit Notes',     table: 'debit_notes' },
};

// Columns the simple editor must never let the user change.
function _dataLocked(col) {
  return col === 'id'
    || /_id$/.test(col)              // foreign-key links (auction_id, trader_id…)
    || col === 'created_at'
    || col === 'locked_at'
    || col === 'price_checked_at'
    || /_hash$/.test(col);
}

// Prettier column labels; falls back to Title Case of the raw name.
const DATA_COL_LABELS = {
  // Sellers (traders)
  name:'Name', cr:'Code', pan:'PAN', tel:'Phone', aadhar:'Aadhaar',
  padd:'Address', ppla:'Place', pin:'PIN', pstate:'State', pst_code:'State Code',
  ifsc:'IFSC', acctnum:'Account No', holder_name:'Account Holder',
  whatsapp:'WhatsApp', email:'Email',
  // Buyers
  buyer:'Buyer Code', buyer1:'Buyer Name', code:'Code', sbl:'SBL Code',
  add1:'Address 1', add2:'Address 2', pla:'Place', state:'State',
  st_code:'State Code', gstin:'GSTIN', ti:'TIN', sale:'Sale Type', tdsq:'TDS',
  cbuyer1:'Consignee Name', cadd1:'Consignee Addr 1', cadd2:'Consignee Addr 2',
  cpla:'Consignee Place', cpin:'Consignee PIN', cstate:'Consignee State',
  cst_code:'Consignee State Code', cgstin:'Consignee GSTIN',
  // Trades / invoices / lots / bills
  place:'Place', ano:'Trade No', invo:'Invoice No', date:'Date',
  qty:'Quantity', amount:'Amount', cgst:'CGST', sgst:'SGST', igst:'IGST',
  lorry_no:'Lorry No', distance_km:'Distance (km)', lot_no:'Lot No',
  crop:'Crop', grade:'Grade', price:'Price', net:'Net', cost:'Cost',
  note_no:'Note No', total:'Total', tds:'TDS',
};
function _dataColLabel(c) {
  return DATA_COL_LABELS[c] ||
    c.replace(/_/g, ' ').replace(/\b\w/g, m => m.toUpperCase());
}
function _dataColType(sqliteType, name) {
  if (/date/i.test(name)) return 'date';
  return /INT|REAL|NUM|DEC|FLOA|DOUB/i.test(sqliteType || '') ? 'number' : 'text';
}

// Catalog: the friendly menu of editable entities + row counts.
app.get('/api/data/catalog', requireAdmin, (req, res) => {
  try {
    const db = getDb();
    const out = Object.entries(DATA_ENTITIES).map(([key, def]) => {
      let count = null;
      try { count = db.get(`SELECT COUNT(*) c FROM "${def.table}"`).c; } catch (_) {}
      return { key, label: def.label, count };
    });
    res.json({ entities: out });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// Read one entity: friendly column metadata + a page of rows.
app.get('/api/data/:entity', requireAdmin, (req, res) => {
  try {
    const def = DATA_ENTITIES[req.params.entity];
    if (!def) return res.status(404).json({ error: 'Unknown data section' });
    const db = getDb();
    const cols = db.all(`PRAGMA table_info("${def.table}")`).map(c => ({
      name: c.name,
      label: _dataColLabel(c.name),
      type: _dataColType(c.type, c.name),
      locked: _dataLocked(c.name),
    }));
    const limit = Math.min(parseInt(req.query.limit, 10) || 100, 1000);
    const offset = parseInt(req.query.offset, 10) || 0;
    const q = (req.query.q || '').trim();
    // Per-column filters: ?filters={"tel":"99","name":"ravi"} — each is a
    // substring match, AND-combined, on top of the global `q` (which is an
    // OR across every column). Unknown columns are ignored.
    let filters = {};
    try { filters = req.query.filters ? JSON.parse(req.query.filters) : {}; } catch (_) {}
    const colNames = new Set(cols.map(c => c.name));
    const clauses = [], params = [];
    if (q) {
      clauses.push('(' + cols.map(c => `CAST("${c.name}" AS TEXT) LIKE ?`).join(' OR ') + ')');
      cols.forEach(() => params.push(`%${q}%`));
    }
    for (const [col, val] of Object.entries(filters)) {
      if (!colNames.has(col) || val == null || val === '') continue;
      clauses.push(`CAST("${col}" AS TEXT) LIKE ?`);
      params.push(`%${val}%`);
    }
    const where = clauses.length ? 'WHERE ' + clauses.join(' AND ') : '';
    const total = db.get(`SELECT COUNT(*) c FROM "${def.table}" ${where}`, params).c;
    const rows = db.all(
      `SELECT rowid AS _rowid_, * FROM "${def.table}" ${where} ORDER BY rowid LIMIT ? OFFSET ?`,
      [...params, limit, offset]);
    res.json({ key: req.params.entity, label: def.label, columns: cols, rows, total, limit, offset });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// Add a new row. Body: { values: { col: val, … } } — only non-locked,
// real columns are accepted; blank fields are simply omitted so the
// table's own defaults apply.
app.post('/api/data/:entity', requireAdmin, (req, res) => {
  try {
    const def = DATA_ENTITIES[req.params.entity];
    if (!def) return res.status(404).json({ error: 'Unknown data section' });
    const db = getDb();
    const real = new Set(db.all(`PRAGMA table_info("${def.table}")`).map(c => c.name));
    const vals = req.body.values || {};
    const keys = Object.keys(vals).filter(k => real.has(k) && !_dataLocked(k));
    if (!keys.length) return res.status(400).json({ error: 'No fields supplied' });
    _devSnapshot('fixdata');
    const info = db.run(
      `INSERT INTO "${def.table}" (${keys.map(k => `"${k}"`).join(',')}) ` +
      `VALUES (${keys.map(() => '?').join(',')})`,
      keys.map(k => vals[k]));
    try {
      db.run('INSERT INTO audit_log (user_id, action, entity, entity_id, details) VALUES (?,?,?,?,?)',
        [req.user.id, 'fixdata_insert', def.table, String(info.lastInsertRowid),
         JSON.stringify(keys.map(k => ({ field: k, to: vals[k] })))]);
    } catch (_) {}
    res.json({ success: true, rowid: info.lastInsertRowid });
  } catch (e) { res.status(400).json({ error: e.message }); }
});

// Update one row by rowid — only non-locked, real columns are honoured.
app.put('/api/data/:entity/:rowid', requireAdmin, (req, res) => {
  try {
    const def = DATA_ENTITIES[req.params.entity];
    if (!def) return res.status(404).json({ error: 'Unknown data section' });
    const db = getDb();
    const real = new Set(db.all(`PRAGMA table_info("${def.table}")`).map(c => c.name));
    const vals = req.body.values || {};
    const keys = Object.keys(vals).filter(k => real.has(k) && !_dataLocked(k));
    const blocked = Object.keys(vals).filter(k => real.has(k) && _dataLocked(k));
    if (blocked.length) return res.status(400).json({ error: `These fields are protected and can't be edited here: ${blocked.join(', ')}` });
    if (!keys.length) return res.status(400).json({ error: 'No editable fields supplied' });
    _devSnapshot('fixdata');
    const rowid = parseInt(req.params.rowid, 10);
    const info = db.run(
      `UPDATE "${def.table}" SET ${keys.map(k => `"${k}"=?`).join(',')} WHERE rowid=?`,
      [...keys.map(k => vals[k]), rowid]);
    try {
      db.run('INSERT INTO audit_log (user_id, action, entity, entity_id, details) VALUES (?,?,?,?,?)',
        [req.user.id, 'fixdata_update', def.table, String(rowid),
         JSON.stringify(keys.map(k => ({ field: k, to: vals[k] })))]);
    } catch (_) {}
    res.json({ success: true, changes: info.changes });
  } catch (e) { res.status(400).json({ error: e.message }); }
});

// Delete one row by rowid.
app.delete('/api/data/:entity/:rowid', requireAdmin, (req, res) => {
  try {
    const def = DATA_ENTITIES[req.params.entity];
    if (!def) return res.status(404).json({ error: 'Unknown data section' });
    const db = getDb();
    _devSnapshot('fixdata');
    const rowid = parseInt(req.params.rowid, 10);
    const info = db.run(`DELETE FROM "${def.table}" WHERE rowid=?`, [rowid]);
    try {
      db.run('INSERT INTO audit_log (user_id, action, entity, entity_id, details) VALUES (?,?,?,?,?)',
        [req.user.id, 'fixdata_delete', def.table, String(rowid), null]);
    } catch (_) {}
    res.json({ success: true, changes: info.changes });
  } catch (e) { res.status(400).json({ error: e.message }); }
});

// ══════════════════════════════════════════════════════════════
// IMPORT OLD DATA (Task 8) — unified upload + preview + run flow
// ══════════════════════════════════════════════════════════════
// Supports SalesInvoice / Purchase / Bills / DebitNotes / Payments /
// Sellers / Buyers. Two endpoints:
//   POST /api/import-old-data/preview   → header detection + first 50 rows
//   POST /api/import-old-data/run       → validation + (optional dryRun) insert
// Each run is recorded in `import_log` for the History panel.
// Position: registered BEFORE the /api 404 catch-all below — otherwise
// the catch-all wins and every request returns "Not Found".
const IMPORT_MODULES = {
  sales_invoice: {
    label: 'Sales Invoices',
    table: 'invoices',
    // Dedup on trade-no + sale-type + invoice-no, mirroring the live
    // numbering rule (a new invoice's number is checked unique within
    // auction_id + sale, see the POST /invoices guard). Invoice numbers
    // restart per trade AND per sale type, so the same `invo` can legitly
    // appear twice in one trade under L vs I vs E — keying on `ano+invo`
    // alone would wrongly skip the second. `state` is deliberately NOT a
    // key: it holds the company's own business_state, identical on every
    // invoice in a trade, so it can't discriminate (and keyCols double as
    // a required-field gate, which would reject rows lacking a state col).
    keyCols: ['ano', 'sale', 'invo'],
    // auction_id is auto-derived from `ano` at import time so the
    // imported rows show up under the matching trade in the Sales tab
    // (the list filters by auction_id, not ano).
    autoFillAuctionId: true,
    fields: ['auction_id','ano','date','state','sale','invo','buyer','buyer1','gstin','place',
             'bag','qty','amount','gunny','pava_hc','ins','cgst','sgst','igst','tcs','rund','tot'],
    aliases: {
      ano: ['ano','auction_no','trade'],
      date: ['date','invoice_date','inv_date'],
      sale: ['sale','sale_type','type'],
      invo: ['invo','invoice','invoice_no','invno'],
      buyer: ['buyer','buyer_code','code'],
      buyer1: ['buyer1','buyer_name','name'],
      gstin: ['gstin','gst','gst_no'],
      place: ['place','city','pla'],
      bag: ['bag','bags','no_of_bags'],
      qty: ['qty','kilos','weight','kgs'],
      amount: ['amount','cardamom','value'],
      tot: ['tot','total','grand_total','invoice_amount'],
    },
  },
  purchase: {
    label: 'Purchase Invoices',
    table: 'purchases',
    // Trade-no + invoice-no — see sales_invoice note. Invoice numbers
    // repeat across trades, so `ano` is needed to avoid false-dup skips.
    keyCols: ['ano', 'invo'],
    autoFillAuctionId: true,
    // Match the actual `purchases` schema. The legacy export's "cr"
    // column maps to `gstin` here; "bag"/"refund" don't exist on the
    // purchases table at all so they're absent from this list.
    fields: ['auction_id','ano','date','state','br','name','add_line','place','gstin','invo',
             'qty','amount','cgst','sgst','igst','rund','total','tds'],
    aliases: {
      ano:     ['ano','auction_no','trade'],
      invo:    ['invo','invoice','invoice_no'],
      name:    ['name','seller','dealer'],
      gstin:   ['gstin','gst','gst_no','cr','registration'],
      place:   ['place','city','pla'],
      add_line:['add_line','address','add','add1','address1'],
      br:      ['br','branch'],
      qty:     ['qty','kilos','weight','kgs'],
      amount:  ['amount','cardamom','value'],
      total:   ['total','grand_total','invoice_amount'],
      rund:    ['rund','round','round_off'],
      tds:     ['tds','tds_amount'],
    },
  },
  bills: {
    label: 'Bills of Supply',
    table: 'bills',
    // Trade-no + bill-no — see sales_invoice note. Bill numbers repeat
    // across trades, so `ano` is needed to avoid false-dup skips.
    keyCols: ['ano', 'bil'],
    // bills.auction_id is the field the Bills tab filters by — without
    // it the imported rows are orphaned and don't show up under any trade.
    autoFillAuctionId: true,
    fields: ['auction_id','ano','date','state','br','crpt','bil','name','add_line','pla',
             'pstate','st_code','crr','pan','qty','cost','igst','net'],
    aliases: {
      ano: ['ano','auction_no','trade'],
      bil: ['bil','bill','bill_no'],
      name: ['name','seller','planter'],
      qty: ['qty','kilos','weight','kgs'],
      cost: ['cost','amount','cardamom'],
      net: ['net','nett','net_amount'],
    },
  },
  debit_notes: {
    label: 'Debit Notes',
    table: 'debit_notes',
    // Already trade-scoped — trade-no + note-no (order normalized to match
    // the other transactional modules; WHERE-clause order is irrelevant).
    keyCols: ['ano', 'note_no'],
    autoFillAuctionId: true,
    fields: ['auction_id','ano','date','state','name','note_no','amount','cgst','sgst','igst','total'],
    aliases: {
      ano: ['ano','auction_no','trade'],
      note_no: ['note_no','note','dn_no'],
      name: ['name','dealer','buyer'],
    },
  },
  // NOTE: a 'payments' module previously lived here but it pointed at
  // the bills table while the Payments tab aggregates from lots — so
  // imports never surfaced under Payments. It also had keyCols=['bil']
  // with `bil` missing from fields, breaking dup-detection. Removed in
  // favor of the dedicated `bills` module above (writes to the same
  // table, with auction_id auto-fill).
  sellers: {
    label: 'Sellers',
    table: 'traders',
    keyCols: ['name','cr'],
    fields: ['name','cr','pan','tel','aadhar','padd','ppla','pin','pstate','pst_code',
             'ifsc','acctnum','holder_name'],
    aliases: {
      name: ['name','seller','planter','trader'],
      cr: ['cr','gstin'],
      padd: ['padd','address','add','add1','address1'],
      ppla: ['ppla','place','pla','city'],
      pin: ['pin','pincode','zip'],
    },
  },
  buyers: {
    label: 'Buyers',
    table: 'buyers',
    keyCols: ['buyer','code'],
    fields: ['buyer','buyer1','code','sbl','add1','add2','pla','pin','state',
             'st_code','gstin','pan','tel','ti','sale','email'],
    aliases: {
      buyer: ['buyer','buyer_code','code'],
      buyer1: ['buyer1','buyer_name','name'],
      pla: ['pla','place','city'],
      pin: ['pin','pincode','zip'],
    },
  },
};
function _importMapHeaders(headers, moduleDef) {
  const norm = s => String(s || '').trim().toLowerCase().replace(/[\s\-/]+/g, '_');
  const out = {};
  for (const field of moduleDef.fields) {
    const aliases = (moduleDef.aliases && moduleDef.aliases[field]) || [field];
    for (const h of headers) {
      if (aliases.includes(norm(h))) { out[field] = h; break; }
    }
  }
  return out;
}

// Per-module GSTIN→master-name lookup config. When a module is in
// this table, the import flow will:
//   1. Build an in-memory map of normalised GSTIN → canonical name
//      from the listed master table (one DB call per import run).
//   2. For each imported row that has a GSTIN, replace the row's name
//      field with the master name when there's a match.
// Solves the legacy-XLS truncation problem: source files often clip
// long seller / buyer names to 30 chars, but the master record
// already holds the full canonical name. The GSTIN is the reliable
// pivot.
const IMPORT_NAME_BY_GSTIN = {
  purchase:      { masterTable: 'traders', masterGstin: 'cr',    masterName: 'name',   rowGstinField: 'gstin', rowNameField: 'name'   },
  bills:         { masterTable: 'traders', masterGstin: 'cr',    masterName: 'name',   rowGstinField: 'crr',   rowNameField: 'name'   },
  sales_invoice: { masterTable: 'buyers',  masterGstin: 'gstin', masterName: 'buyer1', rowGstinField: 'gstin', rowNameField: 'buyer1' },
};

// Normalize a GSTIN string for map keying. Mirrors _normGstin used by
// the purchase same-entity check — strips legacy "GSTIN."/"GSTIN"
// prefix, trims, uppercases. Defined inline here so this section
// stays self-contained.
function _normGstinForLookup(s) {
  let v = String(s == null ? '' : s).trim().toUpperCase();
  if (v.startsWith('GSTIN.')) v = v.slice(6);
  else if (v.startsWith('GSTIN')) v = v.slice(5);
  v = v.trim();
  // Only a genuine 15-char GSTIN is a safe key. Placeholders that
  // sellers without a GSTIN carry — "CR.", case/registration numbers
  // like "CR.A9/1103/19", etc. — are NOT unique, so keying on them
  // collapses thousands of distinct sellers onto one master record
  // (first-wins in _buildGstinNameMap) and silently overwrites every
  // matching row's name. Reject anything that isn't a real GSTIN so it
  // never becomes a map key and never matches on lookup.
  return GSTIN_RE.test(v) ? v : '';
}

// Build the GSTIN → name map for a given master table. Reads every
// non-blank GSTIN row once. First-wins on duplicates (rare, since
// GSTIN is supposed to be unique anyway).
function _buildGstinNameMap(db, table, gstinCol, nameCol) {
  const map = new Map();
  let rows;
  try {
    rows = db.all(`SELECT ${gstinCol} AS g, ${nameCol} AS n FROM ${table} WHERE ${gstinCol} IS NOT NULL AND ${gstinCol} != ''`);
  } catch (e) {
    console.warn('[import] _buildGstinNameMap failed for', table, '-', e.message);
    return map;
  }
  for (const row of rows) {
    const norm = _normGstinForLookup(row.g);
    if (norm && row.n && !map.has(norm)) map.set(norm, String(row.n).trim());
  }
  return map;
}

// Per-module config for "derive round-off when the source file
// doesn't carry one." Mirrors the round-off math that calculations.js
// `buildSalesInvoice` uses at invoice-generation time:
//
//   totalBeforeRound = taxableValue + cgst + sgst + igst
//   subtotalRounded  = round(totalBeforeRound)         // whole rupees
//   rund             = subtotalRounded - totalBeforeRound
//
// `taxableFields` are the source columns that sum to `taxableValue`
// (cardamom amount + gunny + transport + insurance for sales). The
// `targetField` is where the computed value lands. The lookup is
// applied ONLY when the row's existing `targetField` value is blank
// (so source files that DO carry a rund column are honoured).
const IMPORT_DERIVE_RUND = {
  sales_invoice: {
    taxableFields: ['amount', 'gunny', 'pava_hc', 'ins'],
    taxFields:     ['cgst', 'sgst', 'igst'],
    targetField:   'rund',
  },
  // Purchase imports already carry rund in their export; bills don't
  // have GST so there's no round-off to derive. Add modules here if
  // their source file is missing rund.
};

// Excel-compatible integer round (round half away from zero). Mirrors
// calculations.js `round0` so the derived rund matches what the live
// invoice generator would have produced.
function _round0Int(n) {
  const x = Number(n);
  if (!isFinite(x)) return 0;
  if (x === 0) return 0;
  return (x < 0 ? -1 : 1) * Math.round(Math.abs(x));
}

// Given a row of mapped `values` and a derive-rund config, compute the
// round-off the same way invoice generation does. Returns the rund
// value (positive or negative, two decimals). Caller decides whether
// to overwrite — typically only when the row's stored rund is blank
// or zero.
function _deriveRund(values, cfg) {
  let totalBeforeRound = 0;
  for (const f of cfg.taxableFields) totalBeforeRound += Number(values[f] || 0);
  for (const f of cfg.taxFields)     totalBeforeRound += Number(values[f] || 0);
  const subtotalRounded = _round0Int(totalBeforeRound);
  // Match calculations.js round2 — two-decimal precision.
  const rund = subtotalRounded - totalBeforeRound;
  return Math.round(rund * 100) / 100;
}
app.post('/api/import-old-data/preview', requireAdmin, upload.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
  const moduleKey = req.body.module;
  const def = IMPORT_MODULES[moduleKey];
  if (!def) {
    if (req.file) fs.unlink(req.file.path, () => {});
    return res.status(400).json({ error: 'Unknown module', available: Object.keys(IMPORT_MODULES) });
  }
  try {
    const wb = XLSX.readFile(req.file.path);
    const ws = wb.Sheets[wb.SheetNames[0]];
    if (!ws) throw new Error('No worksheet found');
    const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
    const headers = rows.length ? Object.keys(rows[0]) : [];
    const mapping = _importMapHeaders(headers, def);
    res.json({
      module: moduleKey,
      label: def.label,
      total: rows.length,
      headers,
      // Full list of DB fields the client should expose in the mapping
      // editor — without this the UI only shows auto-detected fields and
      // the user can't add a mapping for any column that didn't auto-match.
      fields:  def.fields,
      // Required-for-dup-detection fields the client should highlight.
      keyCols: def.keyCols,
      // Whether auction_id is derived from ano at insert time (so the
      // client can render "(from ano)" instead of a blank cell).
      autoFillAuctionId: !!def.autoFillAuctionId,
      detectedMapping: mapping,
      missingFields: def.fields.filter(f => !mapping[f] && def.keyCols.includes(f)),
      preview: rows.slice(0, 50),
    });
  } catch (e) {
    res.status(400).json({ error: e.message });
  } finally {
    if (req.file) fs.unlink(req.file.path, () => {});
  }
});
// Verify the file against the live database BEFORE writing anything.
// Returns per-row status (new / duplicate / invalid), reasons for invalids,
// the field-by-field diff vs. any existing row, and accurate summary counts
// over the WHOLE file (not just the previewed slice). The detailed row
// list is capped at `sampleLimit` to keep payloads bounded — counts are
// always over the full file.
//
// Differs from /preview (which only does header mapping detection) and
// /run (which actually writes): this endpoint does the same validation
// /run does, against the same mapping the user chose, but without
// touching the DB. Lets the operator catch wrong mappings, missing
// trades, or accidental overwrites before they hit production data.
app.post('/api/import-old-data/verify', requireAdmin, upload.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
  const moduleKey = req.body.module;
  const def = IMPORT_MODULES[moduleKey];
  if (!def) {
    fs.unlink(req.file.path, () => {});
    return res.status(400).json({ error: 'Unknown module' });
  }
  let userMapping = {};
  if (req.body.mapping) {
    try { userMapping = JSON.parse(req.body.mapping) || {}; } catch (_) {}
  }
  // Per-bucket sample cap. Each status (new / invalid / dupChanges) is
  // collected up to this limit independently — that way a file whose
  // first 100 rows happen to all be duplicates still surfaces concrete
  // NEW rows in the UI sample. Counts are always over the WHOLE file.
  const PER_BUCKET_LIMIT = 50;
  const db = getDb();
  try {
    const wb = XLSX.readFile(req.file.path);
    const ws = wb.Sheets[wb.SheetNames[0]];
    if (!ws) throw new Error('No worksheet found');
    const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
    const total = rows.length;
    if (!total) {
      fs.unlink(req.file.path, () => {});
      return res.json({
        module: moduleKey, label: def.label, total: 0,
        fields: def.fields, keyCols: def.keyCols,
        autoFillAuctionId: !!def.autoFillAuctionId,
        sampleLimit: PER_BUCKET_LIMIT,
        counts: { new: 0, duplicate: 0, duplicateChanged: 0, invalidAno: 0, invalidRequired: 0 },
        samples: { new: [], invalid: [], dupChanges: [] },
      });
    }

    const headers = Object.keys(rows[0]);
    // Merge user overrides on top of auto-detected mapping, honouring
    // explicit '' as a skip signal — identical to /run so verify counts
    // reflect what /run would actually do.
    const autoDetected = _importMapHeaders(headers, def);
    const mapping = Object.assign({}, autoDetected, userMapping);
    for (const k of Object.keys(userMapping || {})) {
      const v = userMapping[k];
      if (v === '' || v === null) delete mapping[k];
    }
    const fieldSources = def.fields.map(f => [f, mapping[f] || null]);
    const auctionIdSlot = def.fields.indexOf('auction_id');

    const auctionIdCache = new Map();
    const resolveAuctionId = (ano) => {
      const key = String(ano || '').trim();
      if (!key) return null;
      if (auctionIdCache.has(key)) return auctionIdCache.get(key);
      const row = db.get('SELECT id FROM auctions WHERE ano = ? LIMIT 1', [key]);
      const id  = row ? row.id : null;
      auctionIdCache.set(key, id);
      return id;
    };

    // GSTIN → canonical-name map for this module. Empty when the
    // module doesn't support the lookup. Built once per request.
    const nameLookupCfg = IMPORT_NAME_BY_GSTIN[moduleKey] || null;
    const gstinNameMap = nameLookupCfg
      ? _buildGstinNameMap(db, nameLookupCfg.masterTable, nameLookupCfg.masterGstin, nameLookupCfg.masterName)
      : null;
    let cntNameCorrected = 0;

    // Derive-rund config for this module (null when not applicable).
    // Source files exported from legacy systems often omit the rund
    // column entirely — without this, every imported invoice ends up
    // with rund = '' and Tally exports show 0 in the round-off ledger.
    const deriveRundCfg = IMPORT_DERIVE_RUND[moduleKey] || null;
    let cntRundDerived = 0;

    let cntNew = 0, cntDup = 0, cntDupChanged = 0, cntInvAno = 0, cntInvReq = 0;
    // Four independent sample buckets so the client always has concrete
    // rows from each non-empty status, regardless of where they sit in
    // the file. Each bucket is capped at PER_BUCKET_LIMIT.
    //   • new                — would be inserted
    //   • invalid            — won't be inserted (missing required / no trade)
    //   • dupChanges         — duplicate where the file row differs from DB
    //   • dupIdentical       — duplicate where the file row matches DB exactly
    // Identical dups are surfaced separately because they're the most
    // confusing case for operators: the verify panel says "duplicates: N"
    // but nothing is highlighted, so they can't tell WHICH existing rows
    // are blocking. With this bucket we can show each one's primary key.
    const sampleNew = [];
    const sampleInvalid = [];
    const sampleDupChanges = [];
    const sampleDupIdentical = [];

    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      const values = {};
      for (const [f, src] of fieldSources) {
        values[f] = src ? r[src] : '';
      }

      // GSTIN→master-name lookup: when the row's GSTIN matches a
      // master record, replace the row's name with the master's
      // canonical name. Solves truncated names in legacy XLS dumps.
      if (gstinNameMap && nameLookupCfg) {
        const norm = _normGstinForLookup(values[nameLookupCfg.rowGstinField]);
        if (norm) {
          const masterName = gstinNameMap.get(norm);
          if (masterName) {
            const original = String(values[nameLookupCfg.rowNameField] || '').trim();
            if (masterName !== original) {
              values[nameLookupCfg.rowNameField] = masterName;
              values._nameCorrected = { from: original, to: masterName, gstin: norm };
              cntNameCorrected++;
            }
          }
        }
      }

      // Derive round-off when the source file doesn't carry one.
      // Mirrors invoice generation: rund = round0(taxable+tax) - (taxable+tax).
      // Only fires when the row's stored rund is blank/zero so files
      // that DO carry a rund column are honoured untouched.
      if (deriveRundCfg) {
        const cur = values[deriveRundCfg.targetField];
        const curNum = Number(cur);
        const isBlank = cur === '' || cur == null || (isFinite(curNum) && curNum === 0);
        if (isBlank) {
          const derived = _deriveRund(values, deriveRundCfg);
          if (derived !== 0) {
            values[deriveRundCfg.targetField] = derived;
            values._rundDerived = { value: derived };
            cntRundDerived++;
          }
        }
      }

      const reasons = [];

      let anoResolutionFailed = false;
      if (def.autoFillAuctionId && auctionIdSlot >= 0) {
        const anoSrc = mapping.ano;
        const anoVal = anoSrc ? r[anoSrc] : '';
        const aid = resolveAuctionId(anoVal);
        if (aid == null) {
          anoResolutionFailed = true;
          reasons.push('No trade found for ano="' + String(anoVal || '').trim() + '" — create the auction first or fix the mapping.');
        } else {
          values.auction_id = aid;
        }
      }

      // Required-field check uses the same keyCols /run uses for dup
      // detection. A row missing any of them can't be inserted (and
      // can't be checked for duplicates).
      const missingKeys = [];
      for (const k of def.keyCols) {
        const src = mapping[k];
        const v = src ? r[src] : null;
        if (v == null || String(v).trim() === '') missingKeys.push(k);
      }
      const requiredMissing = missingKeys.length > 0;
      if (requiredMissing) {
        reasons.push('Missing required value(s): ' + missingKeys.join(', '));
      }

      // Duplicate detection — only meaningful when all keyCols resolve
      // to a non-blank value, matching /run's gate.
      // SQLite's loose typing means a string "123" and an integer 123 both
      // match `WHERE invo = ?` even though strict-equality would say no.
      // That's actually what /run does too, so verify mirrors it.
      let existing = null;
      let diff = null;
      if (!requiredMissing) {
        const keyVals = def.keyCols.map(k => r[mapping[k]]);
        const whereSql = def.keyCols.map(k => `${k} = ?`).join(' AND ');
        existing = db.get(`SELECT * FROM ${def.table} WHERE ${whereSql} LIMIT 1`, keyVals);
        if (existing) {
          diff = {};
          for (const f of def.fields) {
            const newVal = values[f];
            const oldVal = existing[f];
            const a = oldVal == null ? '' : String(oldVal);
            const b = newVal == null ? '' : String(newVal);
            if (a !== b) {
              diff[f] = {
                old: oldVal == null ? '' : oldVal,
                new: newVal == null ? '' : newVal,
              };
            }
          }
          if (Object.keys(diff).length === 0) diff = null;
        }
      }

      let status;
      if (requiredMissing) {
        status = 'invalid';
        cntInvReq++;
      } else if (anoResolutionFailed) {
        status = 'invalid';
        cntInvAno++;
      } else if (existing) {
        status = 'duplicate';
        cntDup++;
        if (diff) cntDupChanged++;
      } else {
        status = 'new';
        cntNew++;
      }

      const entry = {
        row: i + 2,
        status,
        reasons,
        values,
        existing: existing || null,
        diff,
      };
      if (status === 'new' && sampleNew.length < PER_BUCKET_LIMIT) {
        sampleNew.push(entry);
      } else if (status === 'invalid' && sampleInvalid.length < PER_BUCKET_LIMIT) {
        sampleInvalid.push(entry);
      } else if (status === 'duplicate' && diff && sampleDupChanges.length < PER_BUCKET_LIMIT) {
        sampleDupChanges.push(entry);
      } else if (status === 'duplicate' && !diff && sampleDupIdentical.length < PER_BUCKET_LIMIT) {
        sampleDupIdentical.push(entry);
      }
    }

    // Total row count + a quick "did the operator actually clear this
    // table?" signal. Without this, when the user deletes invoices under
    // a Sales-tab filter and then re-runs verify, "duplicates: N" is
    // baffling — the UI shows 0 rows but the table still has N rows
    // for other auctions/dates. Surfacing the actual count makes the
    // discrepancy obvious.
    let targetRowCount = 0;
    try {
      const r = db.get(`SELECT COUNT(*) as c FROM ${def.table}`);
      targetRowCount = r ? Number(r.c || 0) : 0;
    } catch (_) { /* table missing — leave at 0 */ }

    fs.unlink(req.file.path, () => {});
    res.json({
      module: moduleKey,
      label: def.label,
      total,
      fields: def.fields,
      keyCols: def.keyCols,
      autoFillAuctionId: !!def.autoFillAuctionId,
      sampleLimit: PER_BUCKET_LIMIT,
      // Pre-import state of the target table — surfaced so the operator
      // can confirm the table really is empty before re-importing.
      targetTable: def.table,
      targetRowCount,
      counts: {
        new: cntNew,
        duplicate: cntDup,
        duplicateChanged: cntDupChanged,
        invalidAno: cntInvAno,
        invalidRequired: cntInvReq,
        // Rows where the imported name was overridden by the matching
        // master record (looked up via GSTIN). Lets the UI tell the
        // operator "20 names were corrected from the master."
        nameCorrected: cntNameCorrected,
        // Rows whose round-off was computed from taxable + tax columns
        // because the source file had no rund column. 0 when the
        // module isn't in IMPORT_DERIVE_RUND.
        rundDerived: cntRundDerived,
      },
      samples: {
        new: sampleNew,
        invalid: sampleInvalid,
        dupChanges: sampleDupChanges,
        dupIdentical: sampleDupIdentical,
      },
    });
  } catch (e) {
    if (req.file) fs.unlink(req.file.path, () => {});
    return res.status(400).json({ error: e.message });
  }
});
app.post('/api/import-old-data/run', requireAdmin, upload.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
  const moduleKey = req.body.module;
  const def = IMPORT_MODULES[moduleKey];
  if (!def) {
    fs.unlink(req.file.path, () => {});
    return res.status(400).json({ error: 'Unknown module' });
  }
  const dryRun  = String(req.body.dryRun || '').toLowerCase() === 'true';
  // Allow the client to override the auto-detected mapping. Shape:
  //   { mapping: { field: 'Source Column Name' } }   (JSON-encoded string)
  let userMapping = {};
  if (req.body.mapping) {
    try { userMapping = JSON.parse(req.body.mapping) || {}; } catch (_) {}
  }

  const db = getDb();
  let imported = 0, skipped = 0, failed = 0;
  // GSTIN→master-name correction counter — hoisted out of the try
  // block so the response builder further down can read it. Always
  // defined; stays 0 when the module isn't in IMPORT_NAME_BY_GSTIN.
  let nameCorrected = 0;
  // Round-off derivation counter — hoisted for the same reason as
  // nameCorrected. Stays 0 when the module isn't in IMPORT_DERIVE_RUND
  // or when every row already had a non-zero rund.
  let rundDerived = 0;
  const errors = [];
  let total = 0;
  // Capture the primary-key id of every row this import inserts so the
  // Undo button on the History panel can roll back this specific file.
  // Declared OUTSIDE the parse/import try so it's still in scope for the
  // audit-log INSERT further down — previously it lived inside the try
  // and the INSERT threw `insertedIds is not defined`, which meant
  // successful imports never made it into the history table at all.
  const insertedIds = [];
  try {
    const wb = XLSX.readFile(req.file.path);
    const ws = wb.Sheets[wb.SheetNames[0]];
    if (!ws) throw new Error('No worksheet found');
    const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
    total = rows.length;
    if (!total) throw new Error('File is empty');

    const headers  = Object.keys(rows[0]);
    // User mapping overrides auto-detected. Two distinct user signals:
    //   • field absent       → keep auto-detected (no opinion)
    //   • field explicit ''  → SKIP this field (user picked "— skip —")
    // The merge below honours the explicit-skip signal by deleting the
    // entry after merging in user values. Without this, picking "— skip —"
    // on an auto-detected field has no effect.
    const autoDetected = _importMapHeaders(headers, def);
    const mapping = Object.assign({}, autoDetected, userMapping);
    for (const k of Object.keys(userMapping || {})) {
      const v = userMapping[k];
      if (v === '' || v === null) delete mapping[k];
    }

    // Field-list with backing source column (or null if no mapping).
    const fieldSources = def.fields.map(f => [f, mapping[f] || null]);
    const valuePlaceholders = def.fields.map(() => '?').join(',');
    const insertSql = `INSERT INTO ${def.table} (${def.fields.join(',')}) VALUES (${valuePlaceholders})`;

    // GSTIN → canonical-name map for this module (same logic as
    // /verify above). Built once, looked up per row. Resolves the
    // row's gstin field position and name field position so we can
    // patch the positional `values` array before INSERT.
    const nameLookupCfg = IMPORT_NAME_BY_GSTIN[moduleKey] || null;
    const gstinNameMap = nameLookupCfg
      ? _buildGstinNameMap(db, nameLookupCfg.masterTable, nameLookupCfg.masterGstin, nameLookupCfg.masterName)
      : null;
    const nameGstinIdx = nameLookupCfg ? def.fields.indexOf(nameLookupCfg.rowGstinField) : -1;
    const nameSlotIdx  = nameLookupCfg ? def.fields.indexOf(nameLookupCfg.rowNameField)  : -1;
    // (nameCorrected is hoisted to the outer scope so the response
    //  builder can read it after this try block.)

    // Round-off derivation: pre-resolve the positional slot for the
    // target field and pre-resolve every taxable/tax field's slot so
    // we can read them out of the positional `values` array without
    // re-mapping per row.
    const deriveRundCfg = IMPORT_DERIVE_RUND[moduleKey] || null;
    const rundSlotIdx   = deriveRundCfg ? def.fields.indexOf(deriveRundCfg.targetField) : -1;
    const rundTaxableSlots = deriveRundCfg
      ? deriveRundCfg.taxableFields.map(f => def.fields.indexOf(f)).filter(i => i >= 0)
      : [];
    const rundTaxSlots = deriveRundCfg
      ? deriveRundCfg.taxFields.map(f => def.fields.indexOf(f)).filter(i => i >= 0)
      : [];

    // Cache `ano → auction_id` lookups for autoFillAuctionId modules.
    // Without this every row of a large import would hit the DB once
    // for the same trade number.
    const auctionIdCache = new Map();
    const resolveAuctionId = (ano) => {
      const key = String(ano || '').trim();
      if (!key) return null;
      if (auctionIdCache.has(key)) return auctionIdCache.get(key);
      const row = db.get('SELECT id FROM auctions WHERE ano = ? LIMIT 1', [key]);
      const id  = row ? row.id : null;
      auctionIdCache.set(key, id);
      return id;
    };
    const auctionIdSlot = def.fields.indexOf('auction_id');

    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      try {
        // Duplicate detection — skip if any keyCol value already exists.
        const keyChecks = def.keyCols.map(k => mapping[k] ? r[mapping[k]] : null).filter(v => v != null && v !== '');
        if (keyChecks.length === def.keyCols.length) {
          const whereSql = def.keyCols.map(k => `${k} = ?`).join(' AND ');
          const dup = db.get(`SELECT 1 FROM ${def.table} WHERE ${whereSql} LIMIT 1`, keyChecks);
          if (dup) { skipped++; continue; }
        }
        // Build positional values from the source mapping. For
        // autoFillAuctionId modules we derive auction_id from `ano` after
        // the row is mapped — the auctions table is the source of truth
        // for ano→id, and without this the imported invoices show up
        // nowhere because the Sales tab filters by auction_id.
        // `date` is normalized to ISO yyyy-mm-dd here so downstream code
        // (Tally XML's strict toTallyDate, the /api/invoices date BETWEEN
        // filter, etc.) sees a canonical value regardless of whether the
        // spreadsheet held an Excel serial, a DD-MM-YYYY string, etc.
        const values = fieldSources.map(([fname, src]) => {
          const v = src ? r[src] : '';
          if (fname === 'date') return normalizeDate(v);
          return v;
        });
        // GSTIN-driven name correction: when the row's GSTIN matches a
        // master record, overwrite the (often truncated) imported name
        // with the master's canonical name. Quiet — counted but not
        // logged per-row.
        if (gstinNameMap && nameGstinIdx >= 0 && nameSlotIdx >= 0) {
          const norm = _normGstinForLookup(values[nameGstinIdx]);
          if (norm) {
            const masterName = gstinNameMap.get(norm);
            if (masterName && masterName !== String(values[nameSlotIdx] || '').trim()) {
              values[nameSlotIdx] = masterName;
              nameCorrected++;
            }
          }
        }
        // Compute round-off when the source file didn't carry one.
        // Only overwrite when the row's stored rund is blank or 0 so
        // files that already include rund are honoured untouched.
        if (deriveRundCfg && rundSlotIdx >= 0) {
          const cur    = values[rundSlotIdx];
          const curNum = Number(cur);
          const isBlank = cur === '' || cur == null || (isFinite(curNum) && curNum === 0);
          if (isBlank) {
            let totalBeforeRound = 0;
            for (const idx of rundTaxableSlots) totalBeforeRound += Number(values[idx] || 0);
            for (const idx of rundTaxSlots)     totalBeforeRound += Number(values[idx] || 0);
            const subtotalRounded = _round0Int(totalBeforeRound);
            const derived = Math.round((subtotalRounded - totalBeforeRound) * 100) / 100;
            if (derived !== 0) {
              values[rundSlotIdx] = derived;
              rundDerived++;
            }
          }
        }
        if (def.autoFillAuctionId && auctionIdSlot >= 0) {
          const anoSrc = mapping.ano;
          const anoVal = anoSrc ? r[anoSrc] : '';
          const aid    = resolveAuctionId(anoVal);
          if (aid == null) {
            failed++;
            if (errors.length < 50) errors.push({
              row: i + 2,
              error: `No trade found for ano="${String(anoVal || '').trim()}" — create the auction first or fix the column.`
            });
            continue;
          }
          values[auctionIdSlot] = aid;
        }
        if (!dryRun) {
          const info = db.run(insertSql, values);
          if (info && info.lastInsertRowid != null) insertedIds.push(Number(info.lastInsertRowid));
        }
        imported++;
      } catch (e) {
        failed++;
        if (errors.length < 50) errors.push({ row: i + 2, error: e.message });
      }
    }
  } catch (e) {
    fs.unlink(req.file.path, () => {});
    return res.status(400).json({ error: e.message });
  }

  // Back-fill auction_id on any pre-existing rows the user imported
  // *before* this fix landed. Idempotent — touches only NULL slots with
  // a matching auctions.ano. Runs even on dryRun so the user gets the
  // immediate fix without a second pass.
  if (def.autoFillAuctionId) {
    try {
      getDb().run(
        `UPDATE ${def.table}
            SET auction_id = (SELECT id FROM auctions WHERE auctions.ano = ${def.table}.ano)
          WHERE auction_id IS NULL
            AND ano IS NOT NULL AND ano != ''
            AND EXISTS (SELECT 1 FROM auctions WHERE auctions.ano = ${def.table}.ano)`
      );
    } catch (_) { /* non-fatal */ }
  }

  // Repair any bad `date` values left behind by earlier imports (before
  // the per-row normalize above was added). Same scan that runs at
  // startup, but on-demand so the user doesn't have to restart the
  // server to clear non-ISO date strings that make Tally XML drop the
  // <DATE>/<REFERENCEDATE> tags. Idempotent and cheap on a fresh DB.
  if (!dryRun) {
    try { repairBadDates(db); } catch (_) { /* non-fatal */ }
  }

  // Log this run regardless of outcome. `inserted_ids` is left blank on
  // dry-runs (no rows were inserted) so the Undo button stays disabled
  // for that entry. Log failures to the server console — silent
  // best-effort hid a class of bugs where the column schema didn't
  // match the INSERT and every run disappeared from History.
  let importLogId = null;
  try {
    const info = db.run(`INSERT INTO import_log
      (module, filename, dry_run, total, imported, skipped, failed, errors, inserted_ids, user_id, username)
      VALUES (?,?,?,?,?,?,?,?,?,?,?)`,
      [moduleKey, req.file.originalname || '', dryRun ? 1 : 0,
       total, imported, skipped, failed, JSON.stringify(errors).slice(0, 4000),
       dryRun ? '' : JSON.stringify(insertedIds),
       (req.user && req.user.id) || null, (req.user && req.user.username) || '']);
    if (info && info.lastInsertRowid != null) importLogId = Number(info.lastInsertRowid);
    console.log('[import-old-data] Logged run id=' + importLogId +
                ' module=' + moduleKey +
                ' file="' + (req.file.originalname || '') + '"' +
                ' imported=' + imported + ' skipped=' + skipped + ' failed=' + failed +
                ' insertedIds=' + insertedIds.length);
  } catch (e) {
    console.error('[import-old-data] Failed to write import_log entry:', e && e.message ? e.message : e);
  }

  fs.unlink(req.file.path, () => {});
  res.json({
    success: true, module: moduleKey, dryRun,
    total, imported, skipped, failed,
    // Number of rows whose seller / buyer name was overridden by the
    // master record (GSTIN match). 0 when the module isn't in
    // IMPORT_NAME_BY_GSTIN — the UI can choose to hide the line.
    nameCorrected,
    // Number of rows whose round-off was computed from taxable + tax
    // columns because the source file had no rund value. 0 when the
    // module isn't in IMPORT_DERIVE_RUND.
    rundDerived,
    errors, importLogId,
  });
});
app.get('/api/import-old-data/history', requireAdmin, (req, res) => {
  // Force every request to hit the DB — without this, browsers
  // heuristically cache the JSON for hours and the "View History" panel
  // stays stuck on yesterday's data even after a fresh import lands.
  res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate');
  res.setHeader('Pragma', 'no-cache');
  const rows = getDb().all(
    `SELECT id, module, filename, dry_run, total, imported, skipped, failed,
            errors, inserted_ids, undone_at, username, created_at
       FROM import_log ORDER BY id DESC LIMIT 200`
  );
  res.json(rows.map(r => {
    const ids = r.inserted_ids ? safeJSON(r.inserted_ids) : [];
    // Don't bloat the wire — the client only needs the count.
    return {
      id: r.id, module: r.module, filename: r.filename,
      dry_run: !!r.dry_run, total: r.total, imported: r.imported,
      skipped: r.skipped, failed: r.failed,
      errors: r.errors ? safeJSON(r.errors) : [],
      undone_at: r.undone_at || '',
      // Undo is available iff this wasn't a dry-run, the import actually
      // inserted rows, and it hasn't already been rolled back.
      undoable: !r.dry_run && Array.isArray(ids) && ids.length > 0 && !r.undone_at,
      inserted_count: Array.isArray(ids) ? ids.length : 0,
      username: r.username, created_at: r.created_at,
    };
  }));
});

// Roll back a specific import: DELETE every row in the target table
// whose primary key was captured in import_log.inserted_ids, after
// snapshotting the live DB so the rollback itself is reversible.
// Marks the log entry as undone with a timestamp so a second click is
// a no-op and the History panel can disable the button.
app.post('/api/import-old-data/undo/:id', requireAdmin, (req, res) => {
  const db = getDb();
  const logId = Number(req.params.id);
  if (!Number.isFinite(logId)) return res.status(400).json({ error: 'Invalid import id' });
  const logRow = db.get('SELECT * FROM import_log WHERE id = ?', [logId]);
  if (!logRow) return res.status(404).json({ error: 'Import not found' });
  if (logRow.undone_at) return res.status(400).json({ error: 'This import has already been undone at ' + logRow.undone_at });
  if (logRow.dry_run)   return res.status(400).json({ error: 'Dry-run imports did not insert any rows — nothing to undo' });

  const def = IMPORT_MODULES[logRow.module];
  if (!def) return res.status(400).json({ error: 'Unknown module on this import — cannot resolve target table' });

  let ids = [];
  try { ids = JSON.parse(logRow.inserted_ids || '[]'); } catch (_) {}
  if (!Array.isArray(ids) || !ids.length) {
    return res.status(400).json({
      error: 'This import did not record its inserted row IDs (was it run before per-import Undo was added?). To clear it, use the Backup tab → Delete All for ' + def.table + '.',
    });
  }

  // Snapshot before we delete — uses the same helper the Delete All
  // routes use, so an Undo misclick is recoverable via Restore.
  let backupPath = '';
  try { backupPath = _snapshotBackupBeforeDelete('undo-import-' + logRow.module + '-' + logId); }
  catch (e) {
    return res.status(500).json({ error: 'Backup snapshot failed; refusing to undo: ' + (e.message || e) });
  }

  // Bulk-delete by id. SQLite limits parameter count per statement
  // (default 999), so chunk for very large imports.
  let deleted = 0;
  const CHUNK = 500;
  try {
    for (let i = 0; i < ids.length; i += CHUNK) {
      const slice = ids.slice(i, i + CHUNK).filter(n => Number.isFinite(Number(n)));
      if (!slice.length) continue;
      const placeholders = slice.map(() => '?').join(',');
      const info = db.run(`DELETE FROM ${def.table} WHERE id IN (${placeholders})`, slice);
      if (info && typeof info.changes === 'number') deleted += info.changes;
    }
    db.run('UPDATE import_log SET undone_at = datetime("now","localtime") WHERE id = ?', [logId]);
  } catch (e) {
    return res.status(500).json({ error: 'Undo failed mid-way; partial deletions may have occurred. Backup at: ' + backupPath + ' — ' + (e.message || e) });
  }

  res.json({
    success: true,
    importLogId: logId,
    module: logRow.module,
    table: def.table,
    requested: ids.length,
    deleted,
    backupPath,
  });
});
function safeJSON(s){ try { return JSON.parse(s); } catch(_) { return []; } }

// ── /api JSON-safety middleware ─────────────────────────────────
// Every /api/* request must return JSON, never an HTML error page.
// Without these handlers, Express's default 404 + 500 produce HTML
// responses ("Cannot POST /api/..." / stack-trace pages) which crash
// the frontend with "Unexpected token '<' is not valid JSON".
//
// Position: AFTER all routes are registered (so this catches anything
// not handled), BEFORE app.listen().

// 404 for unmatched /api routes — JSON body, never HTML
app.use('/api', (req, res) => {
  res.status(404).json({
    error: `Not Found: ${req.method} ${req.originalUrl}`,
    code: 'route_not_found'
  });
});

// Global error handler — also JSON-only. Must have 4 args for Express
// to recognize it as an error handler. Logs full error server-side
// while sending a sanitized message to the client.
app.use((err, req, res, next) => {
  console.error('[server] unhandled error:', err);
  // If the response already started, defer to default handler (rare,
  // mostly relevant for streamed responses).
  if (res.headersSent) return next(err);
  res.status(err.status || 500).json({
    error: err.message || 'Internal server error',
    code: err.code || 'internal_error'
  });
});

// Schema sanity check — runs once at startup. Catches the "no such
// column" class of bug early (loud server-log warning) instead of at
// request time. Currently asserts a few columns we've historically
// gotten wrong; extend the list as more land-mines surface.
function assertSchemaSanity(db) {
  const checks = [
    // [table, expected_column, hint]
    ['lots',         'auction_id', 'lots tracks the trade via auction_id, NOT a denormalised "ano" column'],
    ['lots',         'buyer',      'lots.buyer is the buyer code on the lot'],
    ['lots',         'name',       'lots.name is the seller name'],
    ['auctions',     'ano',        'trade number lives on auctions.ano (string)'],
    ['debit_notes',  'ano',        'DN keeps the trade number denormalised on the row'],
    ['purchases',    'invo',       'purchase invoice number'],
    ['purchases',    'name',       'purchases.name is the seller'],
  ];
  for (const [tbl, col, hint] of checks) {
    try {
      const cols = db.all(`PRAGMA table_info(${tbl})`).map(r => r.name);
      if (!cols.includes(col)) {
        console.warn(`[schema-check] ⚠ table ${tbl} is missing expected column "${col}" — ${hint}`);
      }
    } catch (e) {
      console.warn(`[schema-check] could not introspect ${tbl}:`, e.message);
    }
  }
}

// ── Auto-backup scheduler ─────────────────────────────────────
// Polls the DB every minute to see if a fresh backup is due. We poll
// rather than setInterval-with-the-configured-interval so a settings
// change applies WITHOUT a restart. State (last run timestamp) is held
// in memory; on boot the earliest possible run is `interval` hours
// after the last on-disk backup file's mtime, so a quick restart
// doesn't trigger an immediate snapshot.
function runBackupTickerOnce() {
  try {
    const db = getDb();
    const cfg = getSettingsFlat(db);
    const enabled = String(cfg.backup_auto_enabled || '').toLowerCase() === 'true';
    if (!enabled) return;
    const intervalHrs = Math.max(1, Number(cfg.backup_interval_hours) || 24);
    const keepN       = Math.max(1, Number(cfg.backup_keep_count) || 14);
    const dbDir = path.dirname(DB_PATH);
    const bkDir = path.join(dbDir, 'backups');
    if (!fs.existsSync(bkDir)) fs.mkdirSync(bkDir, { recursive: true });
    // Most recent backup mtime — drives the "is one due?" check.
    const files = fs.readdirSync(bkDir)
      .filter(f => f.endsWith('.db'))
      .map(f => ({ f, m: fs.statSync(path.join(bkDir, f)).mtimeMs }))
      .sort((a, b) => b.m - a.m);
    const latest = files[0] ? files[0].m : 0;
    const dueAt  = latest + intervalHrs * 3600 * 1000;
    if (Date.now() < dueAt) return;
    // Take a snapshot.
    require('./db').flushSave();
    const stamp = new Date().toISOString().slice(0, 19).replace(/[:T]/g, '-');
    const out = path.join(bkDir, `auto-${stamp}.db`);
    fs.copyFileSync(DB_PATH, out);
    console.log('[backup] auto snapshot written:', out);
    // Prune older snapshots beyond keepN.
    const fresh = fs.readdirSync(bkDir)
      .filter(f => f.endsWith('.db'))
      .map(f => ({ f, m: fs.statSync(path.join(bkDir, f)).mtimeMs }))
      .sort((a, b) => b.m - a.m);
    for (const old of fresh.slice(keepN)) {
      try { fs.unlinkSync(path.join(bkDir, old.f)); console.log('[backup] pruned', old.f); }
      catch (_) { /* ignore */ }
    }
  } catch (e) {
    console.error('[backup] ticker failed:', e.message);
  }
}

const PORT = process.env.PORT || 3001;
(async () => {
  const db = await initDb();
  initCompanySettings(db);
  repairBadDates(db);
  assertSchemaSanity(db);
  // Bootstrap the per-install license row on first boot. This generates
  // the install_id, starts the trial window, and logs the current
  // status so the operator can spot expiry-soon at deploy time.
  try {
    const lstatus = license.getStatus(db);
    const tag = lstatus.expired
      ? `EXPIRED (was ${lstatus.expires_at})`
      : `${lstatus.days_remaining} day${lstatus.days_remaining === 1 ? '' : 's'} remaining (expires ${lstatus.expires_at})`;
    console.log(`  License: install ${lstatus.install_id} — ${tag}`);
  } catch (e) {
    console.warn('  License bootstrap failed:', e.message);
  }
  // Auto-backup poller — once a minute. The function itself decides
  // whether a snapshot is actually due based on the configured interval.
  setInterval(runBackupTickerOnce, 60 * 1000);
  app.listen(PORT, '0.0.0.0', () => {
    console.log(`\n  Admin Console running at http://localhost:${PORT}\n`);
  });
})();
