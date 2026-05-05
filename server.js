const express = require('express');
const crypto = require('crypto');
const path = require('path');
const fs = require('fs');
const multer = require('multer');
const ExcelJS = require('exceljs');
const XLSX = require('xlsx');
const { initDb, getDb, DB_PATH, replaceFromBuffer } = require('./db');
const { initCompanySettings, CATEGORIES, getSetting, getAllSettings, updateSettings, getSettingsFlat, getGSTRates } = require('./company-config');
const { calculateLot, buildSalesInvoice, buildPurchaseInvoice, buildAgriBill, buildDebitNote, listAgriSellers, getPaymentSummary, getBankPaymentData, getTDSReturnData, getSalesJournal, getPurchaseJournal } = require('./calculations');
const { generatePurchaseInvoicePDF, generateCropReceiptPDF, generateAgriBillPDF, generateSalesInvoicePDF, generateSalesInvoicesBatchPDF, generatePurchaseInvoicesBatchPDF, generateAgriBillsBatchPDF } = require('./invoice-pdf');
const { EXPORT_TYPES, createExcelBuffer } = require('./exports');
const { getCompanyHeader, writeXlsxCompanyHeader } = require('./report-formatters');
const { exportPdf: exportAnyPdf } = require('./exports-pdf');
const { DBF_EXPORTS } = require('./dbf-exports');
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

const app = express();
app.use(express.json({ limit: '50mb' }));

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

const hash = pw => crypto.createHash('sha256').update(pw).digest('hex');

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

// Display: yyyy-mm-dd → dd/mm/yyyy (handles Excel serials defensively too)
function fmtDate(d) {
  if (!d && d !== 0) return '';
  const iso = normalizeDate(d);
  if (/^\d{4}-\d{2}-\d{2}$/.test(iso)) {
    const [y, m, day] = iso.split('-');
    return `${day}/${m}/${y}`;
  }
  return String(d);
}

function withFmtDate(rows, field = 'date') {
  if (!Array.isArray(rows)) return rows;
  return rows.map(r => ({ ...r, date_fmt: fmtDate(r[field]) }));
}

// Auth middleware: verify a valid session, attach req.user/req.session.
// DOES NOT check role — use this for endpoints that any logged-in user
// (admin OR regular user) should be able to hit (GET list endpoints mostly).
function requireAuth(req, res, next) {
  const token = (req.headers.authorization || '').replace('Bearer ', '');
  if (!token) return res.status(401).json({ error: 'No token' });
  const db = getDb();
  const session = db.get('SELECT * FROM sessions WHERE token = ?', [token]);
  if (!session) return res.status(403).json({ error: 'Session expired — please sign in again' });
  const user = db.get('SELECT * FROM users WHERE id = ?', [session.user_id]);
  if (!user) return res.status(403).json({ error: 'Unauthorized' });
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
app.get('/api/branding', (req, res) => {
  try {
    const cfg = getSettingsFlat(getDb());
    const isKL = String(cfg.business_state || '').toUpperCase().includes('KERALA');
    const branch = (isKL ? cfg.kl_branch : cfg.tn_branch) || cfg.tn_branch || cfg.kl_branch || '';
    const gstin  = (isKL ? cfg.kl_gstin  : cfg.tn_gstin)  || cfg.tn_gstin  || cfg.kl_gstin  || '';
    const logoFile = path.join(__dirname, 'public', 'logo-ispl.png');
    res.json({
      tradeName: cfg.trade_name || cfg.short_name || '',
      shortName: cfg.short_name || cfg.trade_name || '',
      branch,
      gstin,
      logoUrl: fs.existsSync(logoFile) ? '/logo-ispl.png' : null,
    });
  } catch (e) {
    res.json({ tradeName: '', shortName: '', branch: '', gstin: '', logoUrl: null });
  }
});

app.post('/api/login', (req, res) => {
  const { username, password, device_label } = req.body;
  if (!username || !password) return res.status(400).json({ error: 'Missing credentials' });
  const db = getDb();
  const user = db.get('SELECT * FROM users WHERE username = ?', [username]);
  if (!user || user.password_hash !== hash(password)) return res.status(401).json({ error: 'Invalid credentials' });
  const token = crypto.randomBytes(32).toString('hex');
  // Create a new session row WITHOUT deleting any existing sessions —
  // this lets the same user stay logged in on multiple devices simultaneously.
  db.run('INSERT INTO sessions (token, user_id, device_label) VALUES (?, ?, ?)', [token, user.id, device_label || '']);
  // Clean up very old sessions (> 30 days) so the table doesn't grow forever
  db.run(`DELETE FROM sessions WHERE last_used_at < datetime('now','-30 days')`);
  // Return the user's capabilities array so the client can hide buttons
  // they're not allowed to use. Server still validates every request.
  const permissions = Array.from(ROLE_PERMISSIONS[user.role] || ROLE_PERMISSIONS.viewer);
  res.json({ token, role: user.role, username: user.username, permissions });
});
app.post('/api/logout', (req, res) => {
  const t = (req.headers.authorization||'').replace('Bearer ','');
  if (t) getDb().run('DELETE FROM sessions WHERE token = ?', [t]);
  res.json({ success: true });
});
app.get('/api/me', requireAnyPermission('view', 'lot_entry_view', 'self_password'), (req, res) => {
  const permissions = Array.from(ROLE_PERMISSIONS[req.user.role] || ROLE_PERMISSIONS.viewer);
  res.json({ username: req.user.username, role: req.user.role, permissions });
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

app.post('/api/users', requireUserManage, (req, res) => {
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
  db.run(
    'INSERT INTO users (username, password_hash, role) VALUES (?, ?, ?)',
    [username, hash(password), finalRole]
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

app.put('/api/users/:id/password', requireUserManage, (req, res) => {
  const { password } = req.body || {};
  if (!password || password.length < 4) return res.status(400).json({ error: 'Password must be at least 4 characters' });
  const db = getDb();
  const user = db.get('SELECT id, username FROM users WHERE id = ?', [req.params.id]);
  if (!user) return res.status(404).json({ error: 'User not found' });
  db.run('UPDATE users SET password_hash = ? WHERE id = ?', [hash(password), user.id]);
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
app.put('/api/me/password', requirePermission('self_password'), (req, res) => {
  const { current_password, new_password } = req.body || {};
  if (!current_password || !new_password) return res.status(400).json({ error: 'Both current and new password required' });
  if (new_password.length < 4) return res.status(400).json({ error: 'New password must be at least 4 characters' });
  const db = getDb();
  const user = db.get('SELECT * FROM users WHERE id = ?', [req.user.id]);
  if (!user || user.password_hash !== hash(current_password)) return res.status(401).json({ error: 'Current password is incorrect' });
  db.run('UPDATE users SET password_hash = ? WHERE id = ?', [hash(new_password), user.id]);
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
app.put('/api/company-settings', requireSettingsWrite, (req, res) => {
  const count = updateSettings(getDb(), req.body.settings || {});
  res.json({ success: true, updated: count });
});
app.get('/api/company-settings/flat', requireViewOrLotEntry, (req, res) => res.json(getSettingsFlat(getDb())));

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
// e-Trade-only build: a single 'ispl' logo. The 'asp' slot is preserved so
// older client code that still POSTs there gets a clean error rather than
// an unhandled exception.
const LOGO_FILES = {
  ispl: path.join(__dirname, 'public', 'logo-ispl.png'),
  asp:  path.join(__dirname, 'public', 'logo-asp.png'),
};
app.post('/api/company-settings/logo/:which', requireSettingsWrite, upload.single('file'), (req, res) => {
  const which = req.params.which;
  if (!LOGO_FILES[which]) return res.status(400).json({ error: 'Invalid logo type (use ispl or asp)' });
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
  // Only allow image types
  const ext = (req.file.originalname.split('.').pop() || '').toLowerCase();
  if (!['png', 'jpg', 'jpeg'].includes(ext)) {
    fs.unlinkSync(req.file.path);
    return res.status(400).json({ error: 'Only PNG or JPG images allowed' });
  }
  // Always save as .png at the fixed path (PDFKit handles both PNG and JPEG from PNG extension? No — rename to real ext)
  // Simpler: keep .png in the PDF code always pointing to PNG. For JPEG uploads, save as .jpg alongside.
  const target = LOGO_FILES[which];
  fs.copyFileSync(req.file.path, target);
  fs.unlinkSync(req.file.path);
  res.json({ success: true, path: `/logo-${which}.png`, size: fs.statSync(target).size });
});
app.delete('/api/company-settings/logo/:which', requireSettingsWrite, (req, res) => {
  const which = req.params.which;
  if (!LOGO_FILES[which]) return res.status(400).json({ error: 'Invalid logo type' });
  const target = LOGO_FILES[which];
  if (fs.existsSync(target)) fs.unlinkSync(target);
  res.json({ success: true });
});
// Quick probe so the UI knows whether a logo is uploaded
app.get('/api/company-settings/logo/:which', requireView, (req, res) => {
  const which = req.params.which;
  if (!LOGO_FILES[which]) return res.status(400).json({ error: 'Invalid logo type' });
  const target = LOGO_FILES[which];
  if (!fs.existsSync(target)) return res.json({ exists: false });
  const stat = fs.statSync(target);
  res.json({ exists: true, size: stat.size, mtime: stat.mtime });
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
function makeDeleteAll(table) {
  return (req, res) => {
    try {
      const db = getDb();
      const before = db.get(`SELECT COUNT(*) as c FROM ${table}`).c;
      db.run(`DELETE FROM ${table}`);
      try { db.exec(`DELETE FROM sqlite_sequence WHERE name = '${table}'`); } catch(_) {}
      res.json({ success: true, deleted: before });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  };
}
// Traders have a FK from trader_banks.trader_id → traders.id. A plain
// DELETE FROM traders fails with "FOREIGN KEY constraint failed" whenever
// any seller has a row in trader_banks. Wipe children first, then parents.
app.delete('/api/traders/delete-all', requireDeleteAll, (req, res) => {
  try {
    const db = getDb();
    const before = db.get('SELECT COUNT(*) as c FROM traders').c;
    db.run('DELETE FROM trader_banks');
    db.run('DELETE FROM traders');
    try {
      db.exec("DELETE FROM sqlite_sequence WHERE name = 'traders'");
      db.exec("DELETE FROM sqlite_sequence WHERE name = 'trader_banks'");
    } catch(_) {}
    res.json({ success: true, deleted: before });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});
app.delete('/api/buyers/delete-all',      requireDeleteAll, makeDeleteAll('buyers'));
app.delete('/api/invoices/delete-all',    requireDeleteAll, makeDeleteAll('invoices'));
app.delete('/api/purchases/delete-all',   requireDeleteAll, makeDeleteAll('purchases'));
app.delete('/api/bills/delete-all',       requireDeleteAll, makeDeleteAll('bills'));
// Delete All Debit Notes — TRADE-SCOPED.
// Earlier this was wired to the generic makeDeleteAll('debit_notes') which
// wiped EVERY debit note in the database regardless of trade. Per the
// trade-wise model (each trade owns its own DN sequence), Delete All
// must operate within the currently-selected trade only.
//
// Required query param: ?ano=<trade-number>. Without it, return 400 to
// avoid accidental cross-trade wipes.
app.delete('/api/debit-notes/delete-all', requireDeleteAll, (req, res) => {
  try {
    const db = getDb();
    const ano = String(req.query.ano || '').trim();
    if (!ano) {
      return res.status(400).json({
        error: 'Trade number (ano) is required for Delete All. Refusing global wipe.',
      });
    }
    const before = db.get(
      'SELECT COUNT(*) as c FROM debit_notes WHERE ano = ?',
      [ano]
    ).c;
    db.run('DELETE FROM debit_notes WHERE ano = ?', [ano]);
    res.json({ success: true, deleted: before, ano });
  } catch (e) {
    res.status(500).json({ error: 'Delete All failed: ' + (e.message || e) });
  }
});

// Trades (auctions) bulk-delete — cascades through every child table that
// references auction_id. Order matters: clear leaf rows first so foreign-
// key dependencies don't block the parent delete. Wraps in a try/catch
// per-table because some installs may not have every optional table.
app.delete('/api/auctions/delete-all', requireDeleteAll, (req, res) => {
  try {
    const db = getDb();
    const before = db.get('SELECT COUNT(*) as c FROM auctions').c;
    const childTables = ['lots', 'invoices', 'purchases', 'bills', 'debit_notes', 'lot_allocations'];
    for (const t of childTables) {
      try { db.run(`DELETE FROM ${t}`); } catch(_) {}
      try { db.exec(`DELETE FROM sqlite_sequence WHERE name = '${t}'`); } catch(_) {}
    }
    db.run('DELETE FROM auctions');
    try { db.exec("DELETE FROM sqlite_sequence WHERE name = 'auctions'"); } catch(_) {}
    res.json({ success: true, deleted: before });
  } catch (e) {
    res.status(500).json({ error: e.message });
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

  // With API key → attempt live lookup
  try {
    const url = `https://sheet.gstincheck.co.in/check/${apiKey}/${gstin}`;
    const r = await fetch(url);
    if (!r.ok) throw new Error(`HTTP ${r.status}`);
    const body = await r.json();
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
        source:   'live'
      });
    }
    return res.json({
      valid: true, gstin, pan, st_code: stCode, state,
      source: 'structural',
      note: body && body.message ? body.message : 'GST portal returned no data'
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
      `SELECT id, trader_id, bank_name, acctnum, ifsc, holder_name, is_default
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
  if (search) {
    const q = `%${search}%`;
    const rows = db.all(
      `SELECT * FROM traders 
       WHERE name LIKE ? OR tel LIKE ? OR cr LIKE ? OR pan LIKE ? OR ppla LIKE ? OR aadhar LIKE ?
       ORDER BY name LIMIT ?`,
      [q, q, q, q, q, q, parseInt(limit)||50]
    );
    return res.json(hydrateBanks(rows));
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
// Optional automation: when WHATSAPP_TOKEN + WHATSAPP_PHONE_ID env vars
// are set, these endpoints push messages directly via Meta's Graph API.
// When env is unset, return 501 so the client falls back to its manual
// (web.whatsapp.com link / Web Share API) flow without erroring.
//
// Setup: https://developers.facebook.com/docs/whatsapp/cloud-api
//   1. Create a Meta Business account → add a WhatsApp Business phone.
//   2. Generate a permanent access token, capture the phone-number-id.
//   3. Set env vars: WHATSAPP_TOKEN, WHATSAPP_PHONE_ID
//   4. (Outbound to non-customers needs an approved message template;
//      conversation-initiated messages within 24h work without templates.)
function _waConfigured() {
  return Boolean(process.env.WHATSAPP_TOKEN && process.env.WHATSAPP_PHONE_ID);
}
async function _waPost(path, body) {
  const url = `https://graph.facebook.com/v18.0/${process.env.WHATSAPP_PHONE_ID}${path}`;
  const r = await fetch(url, {
    method: 'POST',
    headers: { 'Authorization': 'Bearer ' + process.env.WHATSAPP_TOKEN, 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });
  const j = await r.json().catch(() => ({}));
  if (!r.ok) throw new Error(j.error && j.error.message ? j.error.message : `Cloud API error ${r.status}`);
  return j;
}
function _waNormalizePhone(tel) {
  const d = String(tel || '').replace(/\D/g, '');
  if (!d) return '';
  return d.length === 10 ? '91' + d : d;
}

// Send a plain text message. Body: { phone, message }
app.post('/api/whatsapp/send-text', requireView, async (req, res) => {
  if (!_waConfigured()) return res.status(501).json({ error: 'WhatsApp Cloud API not configured', fallback: true });
  try {
    const phone = _waNormalizePhone(req.body.phone);
    const message = String(req.body.message || '');
    if (!phone || !message) return res.status(400).json({ error: 'phone and message required' });
    const out = await _waPost('/messages', {
      messaging_product: 'whatsapp', to: phone, type: 'text', text: { body: message },
    });
    res.json({ ok: true, id: out.messages && out.messages[0] && out.messages[0].id });
  } catch (e) { res.status(502).json({ error: e.message }); }
});

// Send a document. Body: { phone, caption, doc_url } where doc_url is a
// publicly reachable URL Meta will fetch the PDF from. (For locally
// generated PDFs you'd first POST /media to upload and use the returned
// id — left as a TODO when you wire this up to the actual hosted URLs.)
app.post('/api/whatsapp/send-document', requireView, async (req, res) => {
  if (!_waConfigured()) return res.status(501).json({ error: 'WhatsApp Cloud API not configured', fallback: true });
  try {
    const phone = _waNormalizePhone(req.body.phone);
    const caption = String(req.body.caption || '');
    const docUrl = String(req.body.doc_url || '');
    if (!phone || !docUrl) return res.status(400).json({ error: 'phone and doc_url required' });
    const out = await _waPost('/messages', {
      messaging_product: 'whatsapp', to: phone, type: 'document',
      document: { link: docUrl, caption, filename: req.body.filename || 'document.pdf' },
    });
    res.json({ ok: true, id: out.messages && out.messages[0] && out.messages[0].id });
  } catch (e) { res.status(502).json({ error: e.message }); }
});

app.get('/api/traders/:id', requireViewOrLotEntry, (req, res) => {
  const db = getDb();
  const row = db.get('SELECT * FROM traders WHERE id = ?', [req.params.id]);
  if (!row) return res.status(404).json({ error: 'Not found' });
  // Attach banks array so the edit modal sees all bank accounts.
  row.banks = db.all(
    'SELECT id, trader_id, bank_name, acctnum, ifsc, holder_name, is_default FROM trader_banks WHERE trader_id = ? ORDER BY is_default DESC, id',
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
      'INSERT INTO trader_banks (trader_id, bank_name, acctnum, ifsc, holder_name) VALUES (?,?,?,?,?)',
      [traderId, b.bank_name||'', String(b.acctnum||''), String(b.ifsc||''), b.holder_name||'']
    );
  }
  // Mirror first bank into traders row for legacy compatibility
  const first = arr[0] || {};
  db.run(
    'UPDATE traders SET ifsc=?, acctnum=?, holder_name=? WHERE id=?',
    [first.ifsc||'', first.acctnum||'', first.holder_name||'', traderId]
  );
}

app.post('/api/traders', requireTraderWrite, (req, res) => {
  const t = req.body;
  const db = getDb();
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
  db.run(`UPDATE traders SET name=?,cr=?,pan=?,tel=?,aadhar=?,padd=?,ppla=?,pin=?,pstate=?,pst_code=?,ifsc=?,acctnum=?,holder_name=? WHERE id=?`,
    [t.name,t.cr||'',t.pan||'',t.tel||'',t.aadhar||'',t.padd||'',t.ppla||'',t.pin||'',t.pstate||'',t.pst_code||'',t.ifsc||'',t.acctnum||'',t.holder_name||'',req.params.id]);
  if (Array.isArray(t.banks)) {
    syncTraderBanks(db, parseInt(req.params.id), t.banks);
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
    metaLines: [`Date: ${new Date().toLocaleDateString('en-GB')}`],
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
  if (search) {
    const q = `%${search}%`;
    return res.json(db.all(
      `SELECT * FROM buyers 
       WHERE buyer LIKE ? OR buyer1 LIKE ? OR tel LIKE ? OR gstin LIKE ? OR pan LIKE ? OR pla LIKE ? OR ti LIKE ? OR code LIKE ?
       ORDER BY buyer1 LIMIT 50`,
      [q, q, q, q, q, q, q, q]
    ));
  }
  res.json(db.all('SELECT * FROM buyers ORDER BY buyer1 LIMIT 500'));
});
app.post('/api/buyers', requireBuyerWrite, (req, res) => {
  const b = req.body;
  getDb().run(`INSERT INTO buyers (
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
  getDb().run(`UPDATE buyers SET
      buyer=?, buyer1=?, code=?, sbl=?, add1=?, add2=?, pla=?, pin=?, state=?, st_code=?,
      gstin=?, pan=?, tel=?, ti=?, sale=?, email=?, tdsq=?,
      cbuyer1=?, cadd1=?, cadd2=?, cpla=?, cpin=?, cstate=?, cst_code=?, cgstin=?
    WHERE id=?`,
    [b.buyer, b.buyer1||'', b.code||'', b.sbl||'', b.add1||'', b.add2||'', b.pla||'', b.pin||'', b.state||'', b.st_code||'',
     b.gstin||'', b.pan||'', b.tel||'', b.ti||'', b.sale||'L', b.email||'', b.tdsq||'',
     b.cbuyer1||'', b.cadd1||'', b.cadd2||'', b.cpla||'', b.cpin||'', b.cstate||'', b.cst_code||'', b.cgstin||'',
     req.params.id]);
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
    metaLines: [`Date: ${new Date().toLocaleDateString('en-GB')}`],
  });
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename="buyers-template.xlsx"');
  res.send(Buffer.from(buf));
});

// ══════════════════════════════════════════════════════════════
// AUCTIONS
// ══════════════════════════════════════════════════════════════
app.get('/api/auctions', requireViewOrLotEntry, (req, res) => {
  const rows = getDb().all('SELECT *, (SELECT COUNT(*) FROM lots WHERE auction_id=auctions.id) as lot_count FROM auctions ORDER BY date DESC, ano DESC LIMIT 100');
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
app.put('/api/auctions/:id', requireAuctionWrite, (req, res) => {
  const { ano, date, crop_type, state } = req.body;
  const db = getDb();
  const defaultCrop  = getSetting(db, 'default_crop_type') || 'VST';
  const defaultState = getSetting(db, 'business_state')    || 'TAMIL NADU';
  db.run('UPDATE auctions SET ano=?, date=?, crop_type=?, state=? WHERE id=?',
    [ano, normalizeDate(date), crop_type||defaultCrop, state||defaultState, req.params.id]);
  res.json({ success: true });
});
app.delete('/api/auctions/:id', requireDelete, (req, res) => {
  const db = getDb();
  db.run('DELETE FROM lot_allocations WHERE auction_id = ?', [req.params.id]);
  db.run('DELETE FROM lots WHERE auction_id = ?', [req.params.id]);
  db.run('DELETE FROM auctions WHERE id = ?', [req.params.id]);
  res.json({ success: true });
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

  // Refuse to remove an allocation that still has saved lots in its range
  const existing = db.all('SELECT * FROM lot_allocations WHERE auction_id = ?', [auctionId]);
  for (const ex of existing) {
    const kept = allocations.find(a => a.id === ex.id);
    if (!kept) {
      const lots = db.all('SELECT lot_no FROM lots WHERE auction_id = ?', [auctionId]);
      const lotsInRange = lots.filter(l => isLotInRange(l.lot_no, ex.start_lot, ex.end_lot));
      if (lotsInRange.length > 0) {
        return res.status(400).json({ error: `Cannot remove ${ex.branch} (${ex.start_lot}-${ex.end_lot}): ${lotsInRange.length} lots already entered` });
      }
    }
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

// Allocation stats (used/total per branch + per-lot grid)
// Drives both the admin Allocations modal and the Lot Entry status bar.
app.get('/api/auctions/:id/allocation-stats', requireViewOrLotEntry, (req, res) => {
  const db = getDb();
  const auctionId = parseInt(req.params.id, 10);
  const allocations = db.all(
    'SELECT * FROM lot_allocations WHERE auction_id = ? ORDER BY branch, start_lot',
    [auctionId]
  );
  const lots = db.all('SELECT lot_no, branch FROM lots WHERE auction_id = ?', [auctionId]);
  const usedSet = new Set(lots.map(l => l.lot_no));

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
        lotGrid.push({ lot: lotNo, used: usedSet.has(lotNo) });
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
app.post('/api/auctions/import', requireAuctionWrite, upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
  try {
    const workbook = XLSX.readFile(req.file.path);
    const ws = workbook.Sheets[workbook.SheetNames[0]];
    if (!ws) throw new Error('No worksheet found');
    const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
    if (!rows.length) throw new Error('File is empty');

    const db = getDb();
    const mode = req.body.mode || 'full'; // 'full' = new lots, 'price' = update price/buyer only

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
    const overrideAno = req.body.ano;
    const overrideDate = normalizeDate(req.body.date);
    const cropType = req.body.crop_type || mapCol(rows[0], 'CRPT', 'CROP_TYPE', 'CROPTYPE') || (getSetting(db, 'default_crop_type') || 'VST');
    const state = req.body.state || mapCol(rows[0], 'STATE') || 'TAMIL NADU';

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

    // Build auction breakdown for the response
    const auctionBreakdown = [];
    for (const [key, count] of auctionStats) {
      const [ano, date] = key.split('|');
      const auc = auctionCache.get(key);
      auctionBreakdown.push({ id: auc?.id, ano, date, count });
    }
    auctionBreakdown.sort((a,b) => String(a.ano).localeCompare(String(b.ano), undefined, {numeric:true}));

    fs.unlink(req.file.path, () => {});
    res.json({ 
      success: true, 
      imported, updated, skipped, total: rows.length,
      auctionCount: auctionBreakdown.length,
      auctionBreakdown,
      skipReasons 
    });
  } catch (e) {
    if (req.file) fs.unlink(req.file.path, () => {});
    res.status(400).json({ error: e.message });
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
    metaLines: [`Date: ${new Date().toLocaleDateString('en-GB')}`],
  });
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename="auction-lots-template.xlsx"');
  res.send(Buffer.from(buf));
});

// ══════════════════════════════════════════════════════════════
// LOTS (CPA1.DBF — main data)
// ══════════════════════════════════════════════════════════════
app.get('/api/lots/:auctionId', requireViewOrLotEntry, (req, res) => {
  const { branch, name, buyer, limit, offset, paginated, summary } = req.query;
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

  // Summary mode — returns aggregate counts only (cheap, no row data).
  // Used by the Lot Entry stats badge so it shows true totals even when
  // only a 25-row window of lots is loaded client-side.
  if (summary === '1') {
    const aggSql =
      `SELECT COUNT(*) AS n,
              COALESCE(SUM(CAST(bags AS INTEGER)), 0) AS bags,
              COALESCE(SUM(qty), 0)                  AS qty
         FROM lots
        WHERE lots.auction_id = ?`
      + (branch ? ' AND lots.branch = ?' : '')
      + (name   ? ' AND lots.name LIKE ?' : '')
      + (buyer  ? ' AND lots.buyer = ?' : '');
    const row = db.get(aggSql, p) || { n:0, bags:0, qty:0 };
    return res.json({ n: row.n, bags: row.bags, qty: row.qty });
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

  db.run(`INSERT INTO lots (auction_id,lot_no,crop,grade,crpt,branch,state,trader_id,name,padd,ppla,ppin,pstate,pst_code,cr,pan,tel,aadhar,bags,litre,qty,gross_wt,sample_wt,moisture,user_id)
    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`,
    [auctionId,lotNoStr,l.crop||'',l.grade||'',l.crpt||'',branch,l.state||'TAMIL NADU',l.trader_id||null,l.name||'',l.padd||'',l.ppla||'',l.ppin||'',l.pstate||'',l.pst_code||'',l.cr||'',l.pan||'',l.tel||'',l.aadhar||'',l.bags||0,l.litre||'',l.qty||0,l.gross_wt||0,l.sample_wt||0,l.moisture||'',l.user_id||'']);
  res.json({ success: true });
});

app.put('/api/lots/:id', requireLotWrite, (req, res) => {
  const l = req.body; const sets = []; const vals = [];
  const db = getDb();
  const lotId = parseInt(req.params.id, 10);

  // If lot_no or branch is being changed, re-run the allocation +
  // duplicate validation that POST /api/lots applies. Skipping these
  // on edit was a previous gap that let users move a lot outside its
  // branch's range or collide with another lot's number.
  const current = db.get('SELECT auction_id, lot_no, branch FROM lots WHERE id = ?', [lotId]);
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
    if (k !== 'id' && k !== 'auction_id' && k !== 'created_at') { sets.push(`${k}=?`); vals.push(v); }
  }
  vals.push(lotId);
  db.run(`UPDATE lots SET ${sets.join(',')} WHERE id=?`, vals);
  res.json({ success: true });
});

app.delete('/api/lots/:id', requireDelete, (req, res) => {
  getDb().run('DELETE FROM lots WHERE id = ?', [req.params.id]); res.json({ success: true });
});

// ── Calculate all lots for an auction (GENERATE.PRG) ─────────
app.post('/api/lots/calculate/:auctionId', requireLotWrite, (req, res) => {
  const db = getDb(); const cfg = getSettingsFlat(db);
  const lots = db.all('SELECT * FROM lots WHERE auction_id = ? AND amount > 0', [req.params.auctionId]);
  let count = 0;
  for (const lot of lots) {
    const calc = calculateLot(lot, cfg);
    db.run(`UPDATE lots SET pqty=?,prate=?,puramt=?,com=?,sertax=?,cgst=?,sgst=?,igst=?,advance=?,balance=?,bilamt=?,refund=?,refud=?,isp_pqty=?,isp_prate=?,isp_puramt=?,asp_pqty=?,asp_prate=?,asp_puramt=? WHERE id=?`,
      [calc.pqty,calc.prate,calc.puramt,calc.com,calc.sertax,calc.cgst,calc.sgst,calc.igst,calc.advance,calc.balance,calc.bilamt,calc.refund||0,calc.refud||0,calc.isp_pqty||0,calc.isp_prate||0,calc.isp_puramt||0,calc.asp_pqty||0,calc.asp_prate||0,calc.asp_puramt||0,lot.id]);
    count++;
  }
  res.json({ success: true, calculated: count });
});

// Recalculate every lot in every auction with the CURRENT business
// settings. Used by the client when business_state changes — calculations
// like CGST/SGST/IGST and prate are state-sensitive (intra vs inter), so
// the saved values become stale on a state flip and must be refreshed.
//
// Only touches lots with `amount > 0` (skips empty/auction-floor entries).
// Returns total lots calculated across all auctions.
app.post('/api/lots/calculate-all', requireLotWrite, (req, res) => {
  const db = getDb(); const cfg = getSettingsFlat(db);
  const lots = db.all('SELECT * FROM lots WHERE amount > 0');
  let count = 0;
  for (const lot of lots) {
    const calc = calculateLot(lot, cfg);
    db.run(`UPDATE lots SET pqty=?,prate=?,puramt=?,com=?,sertax=?,cgst=?,sgst=?,igst=?,advance=?,balance=?,bilamt=?,refund=?,refud=?,isp_pqty=?,isp_prate=?,isp_puramt=?,asp_pqty=?,asp_prate=?,asp_puramt=? WHERE id=?`,
      [calc.pqty,calc.prate,calc.puramt,calc.com,calc.sertax,calc.cgst,calc.sgst,calc.igst,calc.advance,calc.balance,calc.bilamt,calc.refund||0,calc.refud||0,calc.isp_pqty||0,calc.isp_prate||0,calc.isp_puramt||0,calc.asp_pqty||0,calc.asp_prate||0,calc.asp_puramt||0,lot.id]);
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

// ══════════════════════════════════════════════════════════════
// INVOICES — Sales (GSTIN.PRG / KGSTIN.PRG)
// ══════════════════════════════════════════════════════════════
app.get('/api/invoices', requireView, (req, res) => {
  const { ano, auction_id, from, to } = req.query;
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
  if (from && to) { q += ' AND date BETWEEN ? AND ?'; p.push(from, to); }
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
      '685553'
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

app.post('/api/invoices/generate/:auctionId', requireInvoiceWrite, (req, res) => {
  const db = getDb(); const cfg = getSettingsFlat(db);
  const { saleType, buyerCode, invoiceNo } = req.body;
  
  if (!saleType || !buyerCode || !invoiceNo) {
    return res.status(400).json({ error: 'saleType, buyerCode, and invoiceNo are required' });
  }
  
  // Auto-calculate lots if puramt is missing (user might not have clicked Calculate)
  const uncalc = db.all(`SELECT * FROM lots WHERE auction_id = ? AND amount > 0 AND (puramt IS NULL OR puramt = 0)`, [req.params.auctionId]);
  for (const lot of uncalc) {
    const c = calculateLot(lot, cfg);
    db.run(`UPDATE lots SET pqty=?,prate=?,puramt=?,com=?,sertax=?,cgst=?,sgst=?,igst=?,advance=?,balance=?,bilamt=?,refund=?,refud=?,isp_pqty=?,isp_prate=?,isp_puramt=?,asp_pqty=?,asp_prate=?,asp_puramt=? WHERE id=?`,
      [c.pqty,c.prate,c.puramt,c.com,c.sertax,c.cgst,c.sgst,c.igst,c.advance,c.balance,c.bilamt,c.refund||0,c.refud||0,c.isp_pqty||0,c.isp_prate||0,c.isp_puramt||0,c.asp_pqty||0,c.asp_prate||0,c.asp_puramt||0,lot.id]);
  }
  
  const invoice = buildSalesInvoice(db, req.params.auctionId, buyerCode, saleType, cfg);
  if (!invoice) return res.status(404).json({ error: `No lots found for buyer "${buyerCode}" in this auction. Make sure lots have this buyer code assigned.` });
  
  const auction = db.get('SELECT * FROM auctions WHERE id = ?', [req.params.auctionId]);
  const s = invoice.summary;
  // Store the BUSINESS context state (TAMIL NADU=ISP, KERALA=ASP), not
  // the auction's physical state. This lets us distinguish ASP invoices
  // from ISP invoices in the same auction, which matters for the sales
  // list cross-reference (ASP Inv# column).
  const invoiceState = cfg.business_state || auction.state || '';
  db.run(`INSERT INTO invoices (auction_id,ano,date,state,sale,invo,buyer,buyer1,gstin,place,bag,qty,amount,gunny,pava_hc,ins,cgst,sgst,igst,tcs,rund,tot,addl_chg,addl_name)
    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`,
    [req.params.auctionId,auction.ano,auction.date,invoiceState,saleType,String(invoiceNo),buyerCode,invoice.buyer.buyer1||'',
     invoice.buyer.gstin||'',invoice.buyer.pla||'',s.totalBags,s.totalQty,s.totalAmount,s.gunnyCost,s.transportCost,s.insuranceCost,
     s.cgst,s.sgst,s.igst,0,s.roundDiff,s.grandTotal,s.addlCharge||0,s.addlChargeName||'']);
  
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
      db.run('UPDATE lots SET sale=?, invo=? WHERE auction_id=? AND lot_no=? AND buyer=?',
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
app.post('/api/invoices/generate-all/:auctionId', requireInvoiceWrite, (req, res) => {
  const db = getDb(); const cfg = getSettingsFlat(db);
  const { startInvoiceNo, saleType } = req.body;
  
  let nextNo = parseInt(startInvoiceNo);
  if (!nextNo || nextNo < 1) return res.status(400).json({ error: 'startInvoiceNo must be a positive integer' });
  
  // Auto-calculate uncalculated lots
  const uncalc = db.all(`SELECT * FROM lots WHERE auction_id = ? AND amount > 0 AND (puramt IS NULL OR puramt = 0)`, [req.params.auctionId]);
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
      const invoice = buildSalesInvoice(db, req.params.auctionId, row.buyer, useSaleType, cfg);
      if (!invoice) { errors.push({ buyer: row.buyer, error: 'No matching lots' }); continue; }
      const s = invoice.summary;
      const invoNo = String(nextNo);
      // Store BUSINESS context state — see single-invoice handler for rationale
      const invoiceState = cfg.business_state || auction.state || '';
      db.run(`INSERT INTO invoices (auction_id,ano,date,state,sale,invo,buyer,buyer1,gstin,place,bag,qty,amount,gunny,pava_hc,ins,cgst,sgst,igst,tcs,rund,tot,addl_chg,addl_name)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`,
        [req.params.auctionId,auction.ano,auction.date,invoiceState,useSaleType,invoNo,row.buyer,invoice.buyer.buyer1||'',
         invoice.buyer.gstin||'',invoice.buyer.pla||'',s.totalBags,s.totalQty,s.totalAmount,s.gunnyCost,s.transportCost,s.insuranceCost,
         s.cgst,s.sgst,s.igst,0,s.roundDiff,s.grandTotal,s.addlCharge||0,s.addlChargeName||'']);
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
            db.run('UPDATE lots SET asp_invo=? WHERE auction_id=? AND lot_no=? AND buyer=?',
              [invoNo, req.params.auctionId, li.lot, row.buyer]);
          } else {
            // Don't set `sale` in ASP context — ISP step decides
            db.run('UPDATE lots SET invo=?, asp_invo=? WHERE auction_id=? AND lot_no=? AND buyer=?',
              [invoNo, invoNo, req.params.auctionId, li.lot, row.buyer]);
          }
        } else {
          db.run('UPDATE lots SET sale=?, invo=? WHERE auction_id=? AND lot_no=? AND buyer=?',
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
    const placeholders = cleanIds.map(() => '?').join(',');
    const r = db.run(
      `UPDATE invoices SET lorry_no = ? WHERE id IN (${placeholders})`,
      [v, ...cleanIds]
    );
    res.json({ ok: true, updated: r.changes, lorry_no: v });
  } catch (e) {
    console.error('[lorry-no] bulk update failed:', e);
    res.status(500).json({ error: e.message });
  }
});

// Update invoice fields (edit)
app.put('/api/invoices/:id', requireInvoiceWrite, (req, res) => {
  const i = req.body;
  const fields = ['ano','date','state','sale','invo','buyer','buyer1','gstin','place',
    'bag','qty','amount','gunny','pava_hc','ins','cgst','sgst','igst','tcs','rund','tot'];
  const sets = []; const vals = [];
  for (const f of fields) {
    if (i[f] !== undefined) { sets.push(`${f}=?`); vals.push(i[f]); }
  }
  if (!sets.length) return res.status(400).json({ error: 'No fields to update' });
  vals.push(req.params.id);
  getDb().run(`UPDATE invoices SET ${sets.join(',')} WHERE id=?`, vals);
  res.json({ success: true });
});

// Delete invoice
app.delete('/api/invoices/:id', requireDelete, (req, res) => {
  const db = getDb();
  const inv = db.get('SELECT * FROM invoices WHERE id=?', [req.params.id]);
  if (!inv) return res.status(404).json({ error: 'Invoice not found' });
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
  let lotsFreed = 0;
  for (const inv of invoices) {
    const n = db.get('SELECT COUNT(*) as c FROM lots WHERE auction_id=? AND sale=? AND invo=? AND buyer=?',
      [inv.auction_id, inv.sale, inv.invo, inv.buyer]).c;
    lotsFreed += n;
    db.run(`UPDATE lots SET sale='', invo='' WHERE auction_id=? AND sale=? AND invo=? AND buyer=?`,
      [inv.auction_id, inv.sale, inv.invo, inv.buyer]);
  }
  db.run('DELETE FROM invoices WHERE auction_id = ?', [aid]);
  // Safety net: clear any orphan invo values from lots in this auction
  const orphan = db.get(
    `SELECT COUNT(*) as c FROM lots WHERE auction_id = ? AND invo IS NOT NULL AND invo != ''`, [aid]
  ).c;
  if (orphan) {
    db.run(`UPDATE lots SET sale='', invo='' WHERE auction_id = ?`, [aid]);
    lotsFreed += orphan;
  }
  res.json({ success: true, invoicesReverted: invoices.length, lotsFreed });
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
      ? buildSalesInvoice(db, stored.auction_id, stored.buyer, stored.sale, cfg)
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
        error: 'Purchase view is only available for ASP invoices. Switch business state to KERALA + e-Trade mode.'
      });
    }
    const stored = db.get('SELECT * FROM invoices WHERE id=?', [req.params.id]);
    if (!stored) return res.status(404).json({ error: 'Invoice not found' });

    // Same enrichment pattern as the sales-invoice endpoint
    let invoice = stored.auction_id
      ? buildSalesInvoice(db, stored.auction_id, stored.buyer, stored.sale, cfg)
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
        ? buildSalesInvoice(db, stored.auction_id, stored.buyer, stored.sale, cfg)
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
        error: 'Purchase view is only available for ASP invoices. Switch business state to KERALA + e-Trade mode.'
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
        ? buildSalesInvoice(db, stored.auction_id, stored.buyer, stored.sale, cfg)
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
  const { auction_id, ano, from, to } = req.query;
  let q = 'SELECT * FROM purchases WHERE 1=1'; const p = [];
  if (auction_id) { q += ' AND auction_id = ?'; p.push(parseInt(auction_id)); }
  if (ano) { q += ' AND ano = ?'; p.push(ano); }
  if (from && to) { q += ' AND date BETWEEN ? AND ?'; p.push(from, to); }
  q += ' ORDER BY date DESC LIMIT 500';
  res.json(getDb().all(q, p));
});

app.post('/api/purchases/generate/:auctionId', requireInvoiceWrite, (req, res) => {
  const db = getDb(); const cfg = getSettingsFlat(db);
  const { sellerName, invoiceNo } = req.body;
  const invoice = buildPurchaseInvoice(db, req.params.auctionId, sellerName, cfg);
  if (!invoice) return res.status(404).json({ error: 'No data for this seller' });
  
  const auction = db.get('SELECT * FROM auctions WHERE id = ?', [req.params.auctionId]);
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
  res.json(getDb().all(
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
  ));
});

// Batch: generate purchase invoice for ALL registered dealers in an auction
app.post('/api/purchases/generate-all/:auctionId', requireInvoiceWrite, (req, res) => {
  const db = getDb(); const cfg = getSettingsFlat(db);
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
  
  for (const row of sellers) {
    try {
      const invoice = buildPurchaseInvoice(db, req.params.auctionId, row.name, cfg);
      if (!invoice) { errors.push({ seller: row.name, error: 'Build failed' }); continue; }
      const s = invoice.summary;
      const invoNo = String(nextNo);
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
  const fields = ['ano','date','state','br','name','add_line','place','gstin','invo',
    'qty','amount','cgst','sgst','igst','rund','total','tds'];
  const sets = []; const vals = [];
  for (const f of fields) {
    if (p[f] !== undefined) { sets.push(`${f}=?`); vals.push(p[f]); }
  }
  if (!sets.length) return res.status(400).json({ error: 'No fields to update' });
  vals.push(req.params.id);
  getDb().run(`UPDATE purchases SET ${sets.join(',')} WHERE id=?`, vals);
  res.json({ success: true });
});

// Delete purchase
app.delete('/api/purchases/:id', requireDelete, (req, res) => {
  getDb().run('DELETE FROM purchases WHERE id=?', [req.params.id]);
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
  // Stamp auction date + e-TRADE no
  if (!invoice.invoiceDate && auctionId) {
    const auction = db.get('SELECT date FROM auctions WHERE id = ?', [auctionId]);
    if (auction && auction.date) {
      const d = new Date(auction.date);
      if (!isNaN(d)) invoice.invoiceDate = d.toLocaleDateString('en-GB');
    }
  }
  if (!invoice.invoiceDate) invoice.invoiceDate = new Date().toLocaleDateString('en-GB');
  if (!invoice.eTradeNo) invoice.eTradeNo = String(auctionId || '');

  // Buyer block — populated from the central company identity (which
  // reads the user's CURRENT settings). Legacy `s_*` / `tn_*` fields are
  // only used as a fallback when identity is blank, NOT as the primary
  // source. Inverting the priority is critical: stale `s_company` left
  // over from the dual-company migration was overriding the user's
  // configured company name and producing "ASP address" on PDFs even
  // after the user updated Settings.
  //
  // The isASP branch is structurally retained so the BILLED/SHIPPED TO
  // address can still pick state-specific address lines (Kerala vs Tamil
  // Nadu) when both are populated.
  const isASP = cfg.business_mode === 'e-Trade' && String(cfg.business_state || '').toUpperCase() === 'KERALA';
  if (!invoice.buyer) {
    const _ident = getCompanyIdentity(cfg);
    invoice.buyer = isASP ? {
      name:    _ident.name    || cfg.s_company  || '',
      address: _ident.address1 || cfg.s_address1 || cfg.kl_address1 || '',
      place:   cfg.s_place || cfg.kl_place || '',
      pin:     cfg.s_pin   || cfg.kl_pin   || '',
      state:   _ident.state    || cfg.s_state || 'Kerala',
      st_code: _ident.stateCode || cfg.s_st_code || '32',
      gstin:   _ident.gstin    || cfg.s_gstin || '',
      pan:     _ident.pan      || cfg.s_pan || cfg.pan || '',
    } : {
      name:    _ident.name    || cfg.short_name || cfg.trade_name || '',
      address: _ident.address1 || cfg.tn_address1 || '',
      place:   cfg.tn_place || '',
      pin:     cfg.tn_pin   || '',
      state:   _ident.state    || cfg.tn_state || 'Tamil Nadu',
      st_code: _ident.stateCode || cfg.tn_st_code || '33',
      gstin:   _ident.gstin    || cfg.tn_gstin || '',
      pan:     _ident.pan      || cfg.pan || '',
    };
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
  const { auction_id, ano, from, to } = req.query;
  let q = 'SELECT * FROM bills WHERE 1=1'; const p = [];
  if (auction_id) { q += ' AND auction_id = ?'; p.push(parseInt(auction_id)); }
  if (ano) { q += ' AND ano = ?'; p.push(ano); }
  if (from && to) { q += ' AND date BETWEEN ? AND ?'; p.push(from, to); }
  q += ' ORDER BY date DESC, bil DESC LIMIT 500';
  res.json(withFmtDate(getDb().all(q, p)));
});

// Generate agri bill for a seller
app.post('/api/bills/generate/:auctionId', requireInvoiceWrite, (req, res) => {
  const db = getDb(); const cfg = getSettingsFlat(db);
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
app.post('/api/bills/generate-all/:auctionId', requireInvoiceWrite, (req, res) => {
  const db = getDb(); const cfg = getSettingsFlat(db);
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
      const bill = buildAgriBill(db, req.params.auctionId, row.name, cfg);
      if (!bill || bill.error) { errors.push({ seller: row.name, error: bill?.error || 'Build failed' }); continue; }
      const s = bill.summary;
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
    const auction = db.get('SELECT date FROM auctions WHERE id = ?', [req.params.auctionId]);
    if (auction && auction.date) {
      const d = new Date(auction.date);
      if (!isNaN(d)) bill.billDate = d.toLocaleDateString('en-GB');
    }
    if (!bill.billDate) bill.billDate = new Date().toLocaleDateString('en-GB');
    bill.eTradeNo = req.query.eTradeNo || req.params.auctionId;

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
      if (stored.auction_id) {
        const auction = db.get('SELECT date FROM auctions WHERE id = ?', [stored.auction_id]);
        if (auction && auction.date) {
          const d = new Date(auction.date);
          if (!isNaN(d)) bill.billDate = d.toLocaleDateString('en-GB');
        }
      }
      if (!bill.billDate) bill.billDate = new Date().toLocaleDateString('en-GB');
      bill.eTradeNo = stored.auction_id || '';
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
  const { auction_id, ano, from, to } = req.query;
  let q = 'SELECT * FROM debit_notes WHERE 1=1'; const p = [];
  if (auction_id) { q += ' AND auction_id = ?'; p.push(parseInt(auction_id)); }
  if (ano) { q += ' AND ano = ?'; p.push(ano); }
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
    const discountPct  = Number(cfg.discount_pct)  || 0;
    const discountDays = Number(cfg.discount_days) || 0;
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
  let dealerStateCode = '';
  {
    let cr = String(purchase.cr || '').trim().toUpperCase();
    if (cr.startsWith('GSTIN.')) cr = cr.slice(6);
    else if (cr.startsWith('GSTIN')) cr = cr.slice(5);
    if (/^\d{2}/.test(cr)) dealerStateCode = cr.slice(0, 2);
  }
  const companyStateCode = String(cfg.tally_state_code
      || (String(cfg.business_state || '').toUpperCase() === 'KERALA' ? '32' : '33'));
  const isInter = !!dealerStateCode && dealerStateCode !== companyStateCode;

  const dnGstRate = Number(cfg.discount_gst) || Number(cfg.gst_service) || 18;
  let cgst = 0, sgst = 0, igst = 0;
  if (Number(purchase.cgst) || Number(purchase.sgst) || Number(purchase.igst)) {
    if (isInter) {
      igst = Math.round(discountAmt * dnGstRate / 100 * 100) / 100;
    } else {
      const half = Math.round(discountAmt * (dnGstRate / 2) / 100 * 100) / 100;
      cgst = half; sgst = half;
    }
  }
  const total = Math.round((discountAmt + cgst + sgst + igst) * 100) / 100;

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
  const discountPct  = Number(cfg.discount_pct)  || 0;
  const discountDays = Number(cfg.discount_days) || 0;
  const dnGstRate    = Number(cfg.discount_gst) || Number(cfg.gst_service) || 18;
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
    const discountAmt = discountDays > 0
      ? Math.round((baseAmt / 1000) * discountDays * discountPct)
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
    let dealerStateCode = '';
    {
      let cr = String(p.cr || '').trim().toUpperCase();
      if (cr.startsWith('GSTIN.')) cr = cr.slice(6);
      else if (cr.startsWith('GSTIN')) cr = cr.slice(5);
      if (/^\d{2}/.test(cr)) dealerStateCode = cr.slice(0, 2);
    }
    const companyStateCode = String(cfg.tally_state_code
        || (String(cfg.business_state || '').toUpperCase() === 'KERALA' ? '32' : '33'));
    const isInter = !!dealerStateCode && dealerStateCode !== companyStateCode;

    // GST is only emitted when the source purchase carried GST
    // (registered dealer); URD/agri purchases produce exempt DNs.
    let cgst = 0, sgst = 0, igst = 0;
    if (Number(p.cgst) || Number(p.sgst) || Number(p.igst)) {
      if (isInter) {
        igst = Math.round(discountAmt * dnGstRate / 100 * 100) / 100;
      } else {
        const half = Math.round(discountAmt * (dnGstRate / 2) / 100 * 100) / 100;
        cgst = half; sgst = half;
      }
    }
    const total = Math.round((discountAmt + cgst + sgst + igst) * 100) / 100;

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
  const rows = db.all(
    `SELECT p.id, p.invo, p.name, p.amount, p.cgst, p.sgst, p.igst, p.total, p.date, p.state
       FROM purchases p
      WHERE p.ano = ?
        AND p.amount > 0
        AND NOT EXISTS (
          SELECT 1 FROM debit_notes dn
           WHERE dn.ano = p.ano AND dn.name = p.name
        )
      ORDER BY p.id`,
    [ano]
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
  getDb().run('DELETE FROM debit_notes WHERE id = ?', [req.params.id]);
  res.json({ success: true });
});

// Update debit note (edit)
app.put('/api/debit-notes/:id', requireInvoiceWrite, (req, res) => {
  const n = req.body;
  const fields = ['ano','date','state','name','note_no','amount','cgst','sgst','igst','total'];
  const sets = []; const vals = [];
  for (const f of fields) {
    if (n[f] !== undefined) { sets.push(`${f}=?`); vals.push(n[f]); }
  }
  if (!sets.length) return res.status(400).json({ error: 'No fields to update' });
  vals.push(req.params.id);
  getDb().run(`UPDATE debit_notes SET ${sets.join(',')} WHERE id=?`, vals);
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
    const fmtDate = (d) => {
      if (!d) return '';
      const m = String(d).match(/^(\d{4})-(\d{2})-(\d{2})/);
      return m ? `${m[3]}/${m[2]}/${m[1]}` : String(d);
    };

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
    const ispAddrLine1  = _ident.address1;
    const ispAddrLine2  = _ident.address2;

    const noteRefSuffix = (cfg.season_short || cfg.tally_season || '26-27').replace(/[^0-9-]/g,'');

    doc.font('Helvetica-Bold').fontSize(10).text(company.toUpperCase(), PAGE_L + 4, y);
    doc.font('Helvetica-Bold').fontSize(10).text(`No.: ${dn.note_no || ''}/${noteRefSuffix}`, PAGE_R - 200, y, { width: 196, align: 'right' });
    y = doc.y + 2;
    // Address line 1 sits directly under the company name. Both lines
    // optional — render whichever are populated.
    if (ispAddrLine1) doc.font('Helvetica').fontSize(9).text(ispAddrLine1, PAGE_L + 4, y);
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
    const showGstCol = !!cfg.flag_disc_gst;
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
          ? Math.round(discount * gstRateFraction * 100) / 100
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

    // GRAND TOTAL row
    const GT_H = 18;
    doc.moveTo(PAGE_L, y).lineTo(PAGE_L, y + GT_H).stroke();
    doc.moveTo(xs[5], y).lineTo(xs[5], y + GT_H).stroke();
    doc.moveTo(xs[7], y).lineTo(xs[7], y + GT_H).stroke();
    doc.moveTo(PAGE_R, y).lineTo(PAGE_R, y + GT_H).stroke();
    doc.font('Helvetica-Bold').fontSize(10).text('GRAND TOTAL', xs[5] + 3, y + 5, { width: cols[5].w + cols[6].w - 6, align: 'right' });
    doc.text(fmtAmt(dn.total || totalTaxable), xs[7] + 3, y + 5, { width: cols[7].w - 6, align: 'right' });
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
// INVOICE PREVIEW (PREINVO.PRG) — dry-run, no save
// ══════════════════════════════════════════════════════════════
app.post('/api/invoices/preview/:auctionId', requireView, (req, res) => {
  const db = getDb(); const cfg = getSettingsFlat(db);
  const { saleType, buyerCode, type } = req.body;
  
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
    invoice = buildSalesInvoice(db, req.params.auctionId, buyerCode, saleType, cfg);
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
function _renderPaymentStatement(doc, db, auctionId, sellerName, cfg) {
  const auction = db.get('SELECT * FROM auctions WHERE id = ?', [auctionId]) || { ano:'', date:'' };
  const lots = db.all(
    `SELECT lot_no, qty, rate, amount, puramt, refund, balance, cgst, sgst, igst
       FROM lots WHERE auction_id = ? AND name = ? AND amount > 0
       ORDER BY lot_no`,
    [auctionId, sellerName]
  ) || [];
  const trader = db.get('SELECT * FROM traders WHERE LOWER(name) = LOWER(?) LIMIT 1', [sellerName]);
  const fmtAmt = (n) => Number(n || 0).toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const fmtQty = (n) => Number(n || 0).toLocaleString('en-IN', { minimumFractionDigits: 3, maximumFractionDigits: 3 });
  const fmtDate = (d) => { if (!d) return ''; const m = String(d).match(/^(\d{4})-(\d{2})-(\d{2})/); return m ? `${m[3]}/${m[2]}/${m[1]}` : String(d); };
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
  doc.text(`Seller: ${sellerName}`, PAGE_L, y); doc.text(`Auction: ${auction.ano}`, PAGE_L + 280, y);
  y += 14;
  doc.text(`Phone: ${trader && trader.tel ? trader.tel : '-'}`, PAGE_L, y);
  doc.text(`Date: ${fmtDate(auction.date)}`, PAGE_L + 280, y);
  y += 18;

  const cols = [
    { k: 'lot_no', label: 'Lot#',     x: PAGE_L,        w: 60,  align: 'left' },
    { k: 'qty',    label: 'Qty',      x: PAGE_L + 60,   w: 70,  align: 'right', fmt: fmtQty },
    { k: 'rate',   label: 'Rate',     x: PAGE_L + 130,  w: 60,  align: 'right', fmt: fmtAmt },
    { k: 'amount', label: 'Amount',   x: PAGE_L + 190,  w: 80,  align: 'right', fmt: fmtAmt },
    { k: 'refund', label: 'Discount', x: PAGE_L + 270,  w: 75,  align: 'right', fmt: fmtAmt },
    { k: 'tax',    label: 'GST',      x: PAGE_L + 345,  w: 70,  align: 'right', fmt: fmtAmt },
    { k: 'balance',label: 'Payable',  x: PAGE_L + 415,  w: 120, align: 'right', fmt: fmtAmt },
  ];
  doc.font('Helvetica-Bold').fontSize(9);
  doc.rect(PAGE_L, y, PAGE_W, 18).fillAndStroke('#f3f4f6', '#999').fillColor('#000');
  for (const c of cols) doc.text(c.label, c.x + 2, y + 5, { width: c.w - 4, align: c.align });
  y += 18;
  doc.font('Helvetica').fontSize(9).fillColor('#000');
  let tQty=0,tAmt=0,tDisc=0,tTax=0,tPay=0;
  for (const l of lots) {
    const tax = (Number(l.cgst)||0)+(Number(l.sgst)||0)+(Number(l.igst)||0);
    const row = { ...l, tax };
    tQty+=Number(l.qty)||0; tAmt+=Number(l.amount)||0; tDisc+=Number(l.refund)||0; tTax+=tax; tPay+=Number(l.balance)||0;
    if (y > 770) { doc.addPage(); y = 40; }
    for (const c of cols) {
      const v = c.fmt ? c.fmt(row[c.k]) : String(row[c.k] ?? '');
      doc.text(v, c.x + 2, y + 4, { width: c.w - 4, align: c.align });
    }
    y += 14;
    doc.moveTo(PAGE_L, y).lineTo(PAGE_R, y).strokeColor('#e5e7eb').lineWidth(0.5).stroke().strokeColor('#000');
  }
  doc.font('Helvetica-Bold').fontSize(10);
  doc.rect(PAGE_L, y, PAGE_W, 20).fillAndStroke('#f3f4f6', '#666').fillColor('#000');
  doc.text('TOTAL', PAGE_L + 2, y + 6);
  doc.text(fmtQty(tQty), PAGE_L + 62, y + 6, { width: 66, align: 'right' });
  doc.text(fmtAmt(tAmt), PAGE_L + 192, y + 6, { width: 76, align: 'right' });
  doc.text(fmtAmt(tDisc),PAGE_L + 272, y + 6, { width: 71, align: 'right' });
  doc.text(fmtAmt(tTax), PAGE_L + 347, y + 6, { width: 66, align: 'right' });
  doc.text(fmtAmt(tPay), PAGE_L + 417, y + 6, { width: 116,align: 'right' });
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
    res.setHeader('Content-Disposition', `inline; filename="Payment_${req.params.sellerName}_${auctionId}.pdf"`);
    doc.pipe(res); piped = true;
    res.on('close', () => { try { doc.destroy(); } catch(_){} });
    _renderPaymentStatement(doc, db, auctionId, sellerName, cfg);
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
      res.setHeader('Content-Disposition', `attachment; filename="${niceName}_${auctionId}.pdf"`);
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
    if (exportDef.needsCfg) {
      const cfg = getSettingsFlat(db);
      // Pass state too so exports that need both (e.g. Praman) can filter
      // by state without losing cfg context. Backward-compatible: existing
      // needsCfg exports that ignore the 4th arg are unaffected.
      buffer = await exportDef.fn(db, auctionId, cfg, req.query.state);
    } else {
      buffer = await exportDef.fn(db, auctionId, req.query.state);
    }
    // Per-export-type content-type/extension override (defaults to xlsx).
    // CSV exports like Praman use ext:'csv', mime:'text/csv'.
    const ext  = exportDef.ext  || 'xlsx';
    const mime = exportDef.mime || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    res.setHeader('Content-Type', mime);
    res.setHeader('Content-Disposition', `attachment; filename="${exportDef.name}_${auctionId}.${ext}"`);
    res.send(Buffer.from(buffer));
  } catch(e) {
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
      res.setHeader('Content-Disposition', `attachment; filename="${def.name}_${auctionId}.pdf"`);
      return res.send(buf);
    }
    const buf = await def.xlsx(db, auctionId);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${def.name}_${auctionId}.xlsx"`);
    return res.send(Buffer.from(buf));
  } catch (e) {
    console.error('lorry-reports error:', e);
    return res.status(500).json({ error: e.message });
  }
});

// ══════════════════════════════════════════════════════════════
// DBF EXPORTS (FoxPro-compatible format)
// ══════════════════════════════════════════════════════════════

// List all available DBF export types with labels
app.get('/api/dbf-exports/list', requireExport, (req, res) => {
  const list = {};
  for (const [key, def] of Object.entries(DBF_EXPORTS)) {
    list[key] = {
      label: def.label,
      name: def.name,
      needsAuction: !!def.needsAuction,
      needsDateRange: !!def.needsDateRange,
    };
  }
  res.json(list);
});

// Generic DBF export endpoint
app.get('/api/dbf-exports/:type', requireExport, async (req, res) => {
  const { type } = req.params;
  const def = DBF_EXPORTS[type];
  if (!def) return res.status(400).json({ error: 'Unknown DBF export type', available: Object.keys(DBF_EXPORTS) });

  try {
    const db = getDb();
    let buffer;

    if (def.needsAuction) {
      const { auctionId } = req.query;
      if (!auctionId) return res.status(400).json({ error: 'auctionId query parameter required' });
      buffer = await def.fn(db, auctionId);
    } else if (def.needsDateRange) {
      const { from, to, ano } = req.query;
      const filters = {};
      if (ano) filters.ano = ano;
      if (from && to) { filters.from = from; filters.to = to; }
      buffer = await def.fn(db, filters);
    } else {
      buffer = await def.fn(db);
    }

    // Build filename: LOTS_1.dbf, INV_2026-04-01_to_2026-04-30.dbf, NAM.dbf
    let filename = def.name;
    if (def.needsAuction && req.query.auctionId) filename += `_${req.query.auctionId}`;
    if (def.needsDateRange && req.query.from) filename += `_${req.query.from}_to_${req.query.to}`;
    filename += '.dbf';

    res.setHeader('Content-Type', 'application/x-dbase');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.send(buffer);
  } catch(e) {
    console.error('DBF export error:', e);
    res.status(500).json({ error: 'DBF export failed: ' + e.message });
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
  sales_isp:           { label: 'Sales Vouchers — ISP',                             name: 'SalesISP',           builder: buildSalesIspRows,         generator: generSalesIspXML,     company: 'isp' },
  sales_asp:           { label: 'Sales Vouchers — ASP',                             name: 'SalesASP',           builder: buildSalesAspRows,         generator: generSalesAspXML,     company: 'isp' },
  sales:               { label: 'Sales Vouchers (legacy alias for ISP)',            name: 'Sales',              builder: buildSalesIspRows,         generator: generSalesIspXML,     company: 'isp' },
  // ISP Purchase = the buyer-side mirror of an ASP→ISP transfer. Each
  // sales_asp row produces one isp_purchase voucher into ISP's books with
  // the same VOUCHERNUMBER (e.g. ASP/I-61/26-27) for cross-reference. We
  // re-use buildSalesAspRows directly since the row shape is identical.
  isp_purchase:        { label: 'ISP Purchase Vouchers (mirror of ASP→ISP)',        name: 'ISPPurchase',        builder: buildSalesAspRows,         generator: generIspPurchaseXML,  company: 'isp' },
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
    const totalLots = rows.reduce((s, r) => s + (Array.isArray(r.lots) ? r.lots.length : 0), 0);
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
    cfg.tally_dispatch_pin || cfg.s_pin || cfg.kl_pin || cfg.tn_pin || '685553'
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
              b.pin AS buyer_pin, b.pla AS buyer_pla,
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

    // Annotate each row with resolved distance + source
    const enriched = rows.map(r => {
      let km = null, source = 'none';
      if (r.distance_km != null) {
        km = r.distance_km;
        source = 'manual';
      } else if (r.buyer_pin && routes[String(r.buyer_pin).trim()] != null) {
        km = routes[String(r.buyer_pin).trim()];
        source = 'route';
      }
      return { ...r, resolved_km: km, distance_source: source };
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
    const filename = `Tally_PartyLedger_${kind}_${safeName}_${auctionId}.xml`;
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
    const filename = `${def.name}_${safeName}_${auctionId}.xml`;
    res.setHeader('Content-Type', 'application/xml; charset=utf-8');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.send(xml);
  } catch (e) {
    console.error('tally party-voucher error:', e);
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
    const filename = `${def.name}_${auctionId}.xml`;
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
  // Aggregates over every lot in every auction, regardless of state.
  const cumRow = db.get(
    `SELECT COALESCE(SUM(qty),0) as qty,
            COALESCE(SUM(amount),0) as amount,
            COUNT(*) as lots
     FROM lots`
  ) || {};
  const cumulative = {
    qty:    cumRow.qty    || 0,
    amount: cumRow.amount || 0,
    lots:   cumRow.lots   || 0,
    auctions: counts.auctions,
  };

  // ── Per-trade breakdown (one row per auction, newest first) ──
  // One query with a LEFT JOIN so auctions with zero lots still appear.
  const perTradeBreakdown = db.all(
    `SELECT a.id, a.ano, a.date, a.crop_type,
            COUNT(l.id) as lots,
            COALESCE(SUM(l.qty),0) as qty,
            COALESCE(SUM(l.amount),0) as amount,
            COALESCE(SUM(CASE WHEN l.amount > 0 THEN 1 ELSE 0 END),0) as priced,
            COALESCE(SUM(CASE WHEN l.invo IS NOT NULL AND l.invo != '' THEN 1 ELSE 0 END),0) as invoiced
     FROM auctions a
     LEFT JOIN lots l ON l.auction_id = a.id
     GROUP BY a.id, a.ano, a.date, a.crop_type
     ORDER BY a.date DESC, a.id DESC
     LIMIT 50`
  );

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
    auctionStats = { ...currentAuction, totalLots, priced, invoiced, totalQty, totalAmt };
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

const PORT = process.env.PORT || 3001;
(async () => {
  const db = await initDb();
  initCompanySettings(db);
  repairBadDates(db);
  assertSchemaSanity(db);
  app.listen(PORT, '0.0.0.0', () => {
    console.log(`\n  Admin Console running at http://localhost:${PORT}\n`);
  });
})();
