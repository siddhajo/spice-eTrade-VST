/**
 * db.js — SQL.js variant for cloud deploy (Railway etc.)
 *
 * Why: better-sqlite3 needs native compilation that's been failing on
 * Railway's build infra. sql.js is pure JavaScript — no native bindings,
 * no architecture issues, no compile step.
 *
 * Trade-off: sql.js holds the entire DB in memory and writes the whole
 * file on every commit. For a single-server Railway deployment this is
 * fine since concurrent writes within one Node process are sequential.
 * For multi-machine deploys, switch back to better-sqlite3.
 *
 * Compatibility: This wrapper preserves the same API server.js,
 * calculations.js, company-config.js, exports.js, etc. already use:
 *
 *   db.run(sql, params)           // INSERT/UPDATE/DELETE (params array or spread)
 *   db.get(sql, params)           // SELECT one row
 *   db.all(sql, params)           // SELECT many rows
 *   db.exec(sql)                  // multi-statement SQL
 *   db.prepare(sql).run(...args)  // prepared INSERT/UPDATE
 *   db.prepare(sql).get(...args)  // prepared SELECT one
 *   db.prepare(sql).all(...args)  // prepared SELECT many
 *   db.transaction(fn)            // returns a wrapped function
 */

const initSqlJs = require('sql.js');

// sql.js loads its WASM via fetch/fs. In dev, the file sits at
// `node_modules/sql.js/dist/sql-wasm.wasm` and the default loader
// finds it. In a packaged Electron app, the source is sealed inside
// the `app.asar` archive and Electron's WASM streaming compile can't
// open archive paths. We tell sql.js where to look explicitly:
//   - process.resourcesPath/sql-wasm.wasm   (when packaged; build
//                                            config copies the file via
//                                            extraResources)
//   - node_modules/sql.js/dist/sql-wasm.wasm (dev fallback)
function _resolveSqlJsOpts() {
  const candidates = [];
  // Packaged Electron builds expose process.resourcesPath; we ship
  // the wasm there via electron-builder's extraResources.
  if (process.resourcesPath) {
    candidates.push(path.join(process.resourcesPath, 'sql-wasm.wasm'));
  }
  candidates.push(path.join(__dirname, 'node_modules', 'sql.js', 'dist', 'sql-wasm.wasm'));
  for (const p of candidates) {
    try { if (fs.existsSync(p)) return { locateFile: () => p }; }
    catch (_) { /* fs.existsSync may throw on weird paths — keep trying */ }
  }
  return undefined;  // fall back to sql.js's default resolver
}
const crypto = require('crypto');
const bcrypt = require('bcryptjs');
const fs = require('fs');
const path = require('path');

// DB path: defaults to ./data/config.db (dev / standalone node).
// Electron packaging sets SPICE_DATA_DIR to %APPDATA%\SpiceConfig so the
// database survives app updates and doesn't sit inside the read-only
// installation folder.
const DB_DIR = process.env.SPICE_DATA_DIR || path.join(__dirname, 'data');
const DB_PATH = path.join(DB_DIR, 'config.db');

let SQL = null;        // sql.js module instance (loaded once)
let rawDb = null;       // sql.js Database instance
let wrapped = null;     // our API wrapper
let pendingSave = null; // debounced fs.writeFile timer

/**
 * Persist the in-memory DB to disk. Debounced 200ms so a burst of writes
 * (e.g. invoice generation) only triggers one write.
 */
function scheduleSave() {
  if (pendingSave) clearTimeout(pendingSave);
  pendingSave = setTimeout(() => {
    pendingSave = null;
    if (!rawDb) return;
    try {
      const buf = Buffer.from(rawDb.export());
      const tmp = DB_PATH + '.tmp';
      fs.writeFileSync(tmp, buf);
      fs.renameSync(tmp, DB_PATH);
    } catch (e) {
      console.error('[db] save failed:', e.message);
    }
  }, 200);
}

/**
 * Force-flush any pending save synchronously. Called on shutdown/close.
 */
function flushSave() {
  if (pendingSave) { clearTimeout(pendingSave); pendingSave = null; }
  if (!rawDb) return;
  try {
    const buf = Buffer.from(rawDb.export());
    const tmp = DB_PATH + '.tmp';
    fs.writeFileSync(tmp, buf);
    fs.renameSync(tmp, DB_PATH);
  } catch (e) {
    console.error('[db] flush failed:', e.message);
  }
}

/**
 * Initialize the database. async/await is necessary because sql.js loads
 * its WASM module asynchronously.
 */
async function initDb() {
  if (wrapped) return wrapped;

  const dir = path.dirname(DB_PATH);
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });

  // Load sql.js wasm runtime once
  if (!SQL) SQL = await initSqlJs(_resolveSqlJsOpts());

  // Open existing DB or create empty one
  if (fs.existsSync(DB_PATH)) {
    const buf = fs.readFileSync(DB_PATH);
    rawDb = new SQL.Database(buf);
  } else {
    rawDb = new SQL.Database();
  }

  // Enable foreign keys
  rawDb.run("PRAGMA foreign_keys = ON;");

  wrapped = makeWrapper();

  // Save on process exit (best-effort)
  const onExit = () => { flushSave(); };
  process.on('SIGINT', onExit);
  process.on('SIGTERM', onExit);
  process.on('beforeExit', onExit);

  // ── LICENSE STATE ──────────────────────────────────────────
  // Single-row table (CHECK id = 1) holding the per-install license
  // state. On first boot, ./license.js inserts a row with a fresh
  // install_id and a 30-day trial expiry. The dev's signed tokens
  // bump expires_at when applied via /api/license/apply.
  //
  // active_token stores the most recently applied token verbatim so
  // an operator can copy it back out if they need to re-apply on a
  // restored backup, and so the dev can audit who has what.
  wrapped.exec(`CREATE TABLE IF NOT EXISTS license_state (
    id INTEGER PRIMARY KEY CHECK (id = 1),
    install_id TEXT NOT NULL,
    first_seen_at TEXT NOT NULL,
    expires_at TEXT NOT NULL,
    active_token TEXT
  )`);

  // ── GST API STATE ─────────────────────────────────────────
  // Single-row table mirroring license_state — tracks the most recent
  // observation of the gstincheck.co.in API quota. Every successful
  // /api/gst-lookup call opportunistically updates this row from
  // credit-related fields the API ships in the response (gstinTotalSearch,
  // gstinAvailableSearch, validUpto, etc.). The Settings → Integrations
  // status card and the topbar warning pill both read from here.
  //
  // last_response_raw stores the trimmed JSON envelope (minus the
  // `data` field, which can be large) so the dev can audit what the
  // API returned the last time it was queried — useful when the API
  // changes its field names and we need to teach the parser.
  wrapped.exec(`CREATE TABLE IF NOT EXISTS gst_api_state (
    id INTEGER PRIMARY KEY CHECK (id = 1),
    credits_remaining INTEGER,
    credits_total INTEGER,
    plan_expires_at TEXT,
    last_checked_at TEXT,
    last_response_raw TEXT
  )`);

  // ── SESSIONS ───────────────────────────────────────────────
  wrapped.exec(`CREATE TABLE IF NOT EXISTS sessions (
    token TEXT PRIMARY KEY,
    user_id INTEGER NOT NULL,
    created_at TEXT DEFAULT (datetime('now','localtime')),
    last_used_at TEXT DEFAULT (datetime('now','localtime')),
    expires_at TEXT,
    device_label TEXT DEFAULT '',
    FOREIGN KEY(user_id) REFERENCES users(id)
  )`);

  // ── USERS ──────────────────────────────────────────────────
  wrapped.exec(`CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT NOT NULL UNIQUE,
    password_hash TEXT NOT NULL,
    role TEXT NOT NULL DEFAULT 'user',
    token TEXT,
    must_change_password INTEGER NOT NULL DEFAULT 0,
    created_at TEXT DEFAULT (datetime('now','localtime'))
  )`);

  // ── TRADERS (NAM.DBF — sellers/poolers) ────────────────────
  wrapped.exec(`CREATE TABLE IF NOT EXISTS traders (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    cr TEXT DEFAULT '',
    pan TEXT DEFAULT '',
    tel TEXT DEFAULT '',
    aadhar TEXT DEFAULT '',
    padd TEXT DEFAULT '',
    ppla TEXT DEFAULT '',
    pin TEXT DEFAULT '',
    pstate TEXT DEFAULT '',
    pst_code TEXT DEFAULT '',
    ifsc TEXT DEFAULT '',
    acctnum TEXT DEFAULT '',
    holder_name TEXT DEFAULT '',
    created_at TEXT DEFAULT (datetime('now','localtime'))
  )`);

  // ── COMPANY LOGOS ──────────────────────────────────────────
  // Uploaded company logos live in the DB as BLOBs so they persist
  // wherever the SQLite file persists — survives Railway/Heroku
  // redeploys without needing a separately-mounted volume for the
  // filesystem upload directory. `key` is the logo slot ('ispl',
  // 'asp'); `mime` is the response Content-Type; `data` is the raw
  // bytes. Falls through to the bundled /public default when no row
  // exists for the slot.
  wrapped.exec(`CREATE TABLE IF NOT EXISTS company_logos (
    key TEXT PRIMARY KEY,
    mime TEXT NOT NULL DEFAULT 'image/png',
    data BLOB NOT NULL,
    updated_at TEXT DEFAULT (datetime('now','localtime'))
  )`);

  // ── TRADER BANKS ───────────────────────────────────────────
  wrapped.exec(`CREATE TABLE IF NOT EXISTS trader_banks (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    trader_id INTEGER NOT NULL,
    bank_name TEXT DEFAULT '',
    branch TEXT DEFAULT '',
    acctnum TEXT NOT NULL,
    ifsc TEXT NOT NULL,
    holder_name TEXT DEFAULT '',
    is_default INTEGER DEFAULT 0,
    FOREIGN KEY (trader_id) REFERENCES traders(id)
  )`);

  // ── BUYERS (SBL.DBF — buyers/dealers/traders) ──────────────
  wrapped.exec(`CREATE TABLE IF NOT EXISTS buyers (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    buyer TEXT NOT NULL,
    buyer1 TEXT DEFAULT '',
    code TEXT DEFAULT '',
    sbl TEXT DEFAULT '',
    add1 TEXT DEFAULT '',
    add2 TEXT DEFAULT '',
    pla TEXT DEFAULT '',
    pin TEXT DEFAULT '',
    state TEXT DEFAULT '',
    st_code TEXT DEFAULT '',
    gstin TEXT DEFAULT '',
    pan TEXT DEFAULT '',
    tel TEXT DEFAULT '',
    ti TEXT DEFAULT '',
    sale TEXT DEFAULT 'L',
    email TEXT DEFAULT '',
    tdsq TEXT DEFAULT '',
    cbuyer1 TEXT DEFAULT '',
    cadd1 TEXT DEFAULT '',
    cadd2 TEXT DEFAULT '',
    cpla TEXT DEFAULT '',
    cpin TEXT DEFAULT '',
    cstate TEXT DEFAULT '',
    cst_code TEXT DEFAULT '',
    cgstin TEXT DEFAULT '',
    created_at TEXT DEFAULT (datetime('now','localtime'))
  )`);

  // ── AUCTIONS (trade sessions) ──────────────────────────────
  wrapped.exec(`CREATE TABLE IF NOT EXISTS auctions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ano TEXT NOT NULL,
    date TEXT NOT NULL,
    crop_type TEXT DEFAULT '',
    state TEXT DEFAULT '',
    start_time TEXT,
    end_time TEXT,
    created_at TEXT DEFAULT (datetime('now','localtime')),
    -- Stamped by /api/price-check/verify when the operator has verified
    -- the auction's lots against an external price sheet AND cleared
    -- every code-level discrepancy. Acts as the green-light gate for
    -- calculate / invoice / purchase / bill / debit-note generation.
    -- Auto-cleared by any endpoint that mutates lot price or code.
    price_checked_at TEXT DEFAULT '',
    -- Set on the FIRST successful verify and never cleared. Lets the
    -- gate distinguish "never checked" (hard 412) from "checked then
    -- edited" (soft warning, allow). Once an auction has been verified
    -- at least once, subsequent lot edits don't re-block transactions.
    price_check_first_passed_at TEXT DEFAULT ''
  )`);

  // ── GENERATION OVERRIDES ──────────────────────────────────────
  // Per-trade, per-doc-type "admin grants one regeneration" rows.
  // Each row authorizes a single subsequent generate call for that
  // (auction_id, doc_type) and is consumed (deleted) when the
  // generation endpoint runs successfully. Once consumed, the
  // generate buttons re-lock until admin grants again. Without a
  // row, generation is allowed by default; the row only exists when
  // there's an active "skip the already-generated lock" allowance.
  wrapped.exec(`CREATE TABLE IF NOT EXISTS generation_overrides (
    auction_id INTEGER NOT NULL,
    doc_type   TEXT NOT NULL,
    granted_at TEXT NOT NULL DEFAULT (datetime('now','localtime')),
    granted_by TEXT NOT NULL DEFAULT '',
    PRIMARY KEY (auction_id, doc_type)
  )`);

  // ── LOTS (CPA1.DBF — main lot data, before + after trade) ─
  wrapped.exec(`CREATE TABLE IF NOT EXISTS lots (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    auction_id INTEGER NOT NULL,
    lot_no TEXT NOT NULL,
    crop TEXT DEFAULT '',
    grade TEXT DEFAULT '',
    crpt TEXT DEFAULT '',
    branch TEXT DEFAULT '',
    state TEXT DEFAULT 'TAMIL NADU',
    trader_id INTEGER,
    name TEXT DEFAULT '',
    padd TEXT DEFAULT '',
    ppla TEXT DEFAULT '',
    ppin TEXT DEFAULT '',
    pstate TEXT DEFAULT '',
    pst_code TEXT DEFAULT '',
    cr TEXT DEFAULT '',
    pan TEXT DEFAULT '',
    tel TEXT DEFAULT '',
    aadhar TEXT DEFAULT '',
    bags INTEGER DEFAULT 0,
    litre TEXT DEFAULT '',
    qty REAL DEFAULT 0,
    gross_wt REAL DEFAULT 0,
    sample_wt REAL DEFAULT 0,
    weight_with_gunny REAL DEFAULT 0,
    moisture TEXT DEFAULT '',
    crop_receipt_no INTEGER DEFAULT NULL,
    reserved_price REAL DEFAULT 0,
    price REAL DEFAULT 0,
    amount REAL DEFAULT 0,
    code TEXT DEFAULT '',
    buyer TEXT DEFAULT '',
    buyer1 TEXT DEFAULT '',
    sale TEXT DEFAULT '',
    invo TEXT DEFAULT '',
    pqty REAL DEFAULT 0,
    prate REAL DEFAULT 0,
    puramt REAL DEFAULT 0,
    com REAL DEFAULT 0,
    sertax REAL DEFAULT 0,
    cgst REAL DEFAULT 0,
    sgst REAL DEFAULT 0,
    igst REAL DEFAULT 0,
    dcgst REAL DEFAULT 0,
    dsgst REAL DEFAULT 0,
    digst REAL DEFAULT 0,
    refud REAL DEFAULT 0,
    refund REAL DEFAULT 0,
    advance REAL DEFAULT 0,
    balance REAL DEFAULT 0,
    bilamt REAL DEFAULT 0,
    paid TEXT DEFAULT '',
    user_id TEXT DEFAULT '',
    -- Record-lock columns (also added via ALTER for older DBs in the
    -- migrations block below). See server.js POST /api/lots/lock for
    -- semantics. Defined here too so fresh DBs get the columns +
    -- idx_lots_locked index on the very first boot.
    locked_at TEXT DEFAULT NULL,
    locked_by TEXT DEFAULT NULL,
    created_at TEXT DEFAULT (datetime('now','localtime')),
    FOREIGN KEY (auction_id) REFERENCES auctions(id),
    FOREIGN KEY (trader_id) REFERENCES traders(id)
  )`);

  // ── INVOICES (INV.DBF — sales invoices) ────────────────────
  wrapped.exec(`CREATE TABLE IF NOT EXISTS invoices (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    auction_id INTEGER,
    ano TEXT NOT NULL,
    date TEXT NOT NULL,
    state TEXT DEFAULT '',
    sale TEXT DEFAULT 'L',
    invo TEXT NOT NULL,
    buyer TEXT DEFAULT '',
    buyer1 TEXT DEFAULT '',
    gstin TEXT DEFAULT '',
    place TEXT DEFAULT '',
    lot TEXT DEFAULT '',
    bag INTEGER DEFAULT 0,
    qty REAL DEFAULT 0,
    price REAL DEFAULT 0,
    amount REAL DEFAULT 0,
    gunny REAL DEFAULT 0,
    pava_hc REAL DEFAULT 0,
    ins REAL DEFAULT 0,
    cgst REAL DEFAULT 0,
    sgst REAL DEFAULT 0,
    igst REAL DEFAULT 0,
    tcs REAL DEFAULT 0,
    rund REAL DEFAULT 0,
    tot REAL DEFAULT 0,
    addl_chg REAL DEFAULT 0,
    addl_name TEXT DEFAULT '',
    created_at TEXT DEFAULT (datetime('now','localtime'))
  )`);

  // ── PURCHASES (PURCHASE.DBF — purchase invoices for registered dealers)
  wrapped.exec(`CREATE TABLE IF NOT EXISTS purchases (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    auction_id INTEGER,
    ano TEXT NOT NULL,
    date TEXT NOT NULL,
    state TEXT DEFAULT '',
    br TEXT DEFAULT '',
    name TEXT DEFAULT '',
    add_line TEXT DEFAULT '',
    place TEXT DEFAULT '',
    gstin TEXT DEFAULT '',
    invo TEXT DEFAULT '',
    qty REAL DEFAULT 0,
    amount REAL DEFAULT 0,
    cgst REAL DEFAULT 0,
    sgst REAL DEFAULT 0,
    igst REAL DEFAULT 0,
    rund REAL DEFAULT 0,
    total REAL DEFAULT 0,
    tds REAL DEFAULT 0,
    created_at TEXT DEFAULT (datetime('now','localtime'))
  )`);

  // ── BILLS (BILL.DBF — bills of supply for unregistered/agriculturist)
  wrapped.exec(`CREATE TABLE IF NOT EXISTS bills (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ano TEXT NOT NULL,
    date TEXT NOT NULL,
    state TEXT DEFAULT '',
    br TEXT DEFAULT '',
    crpt TEXT DEFAULT '',
    bil INTEGER DEFAULT 0,
    name TEXT DEFAULT '',
    add_line TEXT DEFAULT '',
    pla TEXT DEFAULT '',
    pstate TEXT DEFAULT '',
    st_code TEXT DEFAULT '',
    crr TEXT DEFAULT '',
    pan TEXT DEFAULT '',
    qty REAL DEFAULT 0,
    cost REAL DEFAULT 0,
    igst REAL DEFAULT 0,
    net REAL DEFAULT 0,
    created_at TEXT DEFAULT (datetime('now','localtime'))
  )`);

  // ── DEBIT NOTES ────────────────────────────────────────────
  wrapped.exec(`CREATE TABLE IF NOT EXISTS debit_notes (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ano TEXT NOT NULL,
    date TEXT NOT NULL,
    state TEXT DEFAULT '',
    name TEXT DEFAULT '',
    note_no TEXT DEFAULT '',
    amount REAL DEFAULT 0,
    cgst REAL DEFAULT 0,
    sgst REAL DEFAULT 0,
    igst REAL DEFAULT 0,
    total REAL DEFAULT 0,
    created_at TEXT DEFAULT (datetime('now','localtime'))
  )`);

  // ── ROUTE DISTANCES (e-way bill <DISTANCE> field) ──────────
  // Saved per (from_pin, to_pin) pair, normalised so the smaller PIN
  // is always stored first — that way A↔B and B↔A share one row. Used
  // by the To Tally → 🗺️ E-way Bill Distance UI: user looks up the
  // distance once on the NIC portal, types it in, and every invoice
  // between the same two PINs (this auction and all future ones) picks
  // it up automatically.
  wrapped.exec(`CREATE TABLE IF NOT EXISTS route_distances (
    from_pin TEXT NOT NULL,
    to_pin TEXT NOT NULL,
    km INTEGER NOT NULL,
    updated_at TEXT DEFAULT (datetime('now','localtime')),
    PRIMARY KEY (from_pin, to_pin)
  )`);

  // ── LOT ALLOCATIONS (per-trade per-branch lot-number ranges) ──
  // Each row reserves a contiguous range of lot numbers (e.g. 001-080)
  // for one branch within one trade. The Lot Entry workflow validates
  // every saved lot's lot_no against these ranges so two field-staff
  // users in different branches can't collide on the same lot number.
  // Ranges are inclusive on both ends and may have an optional alpha
  // prefix (e.g. A001-A080) — see parseLotNo() in server.js.
  wrapped.exec(`CREATE TABLE IF NOT EXISTS lot_allocations (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    auction_id INTEGER NOT NULL,
    branch TEXT NOT NULL,
    start_lot TEXT NOT NULL,
    end_lot TEXT NOT NULL,
    created_at TEXT DEFAULT (datetime('now','localtime')),
    FOREIGN KEY (auction_id) REFERENCES auctions(id)
  )`);

  // ── AUDIT LOG ──────────────────────────────────────────────
  wrapped.exec(`CREATE TABLE IF NOT EXISTS audit_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id TEXT NOT NULL,
    action TEXT NOT NULL,
    entity TEXT NOT NULL,
    entity_id INTEGER,
    details TEXT,
    created_at TEXT DEFAULT (datetime('now','localtime'))
  )`);

  // Forensic record of every "Delete All" wipe. Captures the operator,
  // the affected resource, how many rows actually went away, where the
  // pre-wipe backup landed, and the client IP — so a misclick can be
  // traced and recovered from the snapshot.
  wrapped.exec(`CREATE TABLE IF NOT EXISTS delete_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    resource TEXT NOT NULL,
    deleted_count INTEGER DEFAULT 0,
    cascade_counts TEXT DEFAULT '',
    backup_path TEXT DEFAULT '',
    user_id INTEGER,
    username TEXT DEFAULT '',
    ip TEXT DEFAULT '',
    created_at TEXT DEFAULT (datetime('now','localtime'))
  )`);

  // Audit trail for the Import Old Data tool. One row per upload (preview
  // or run, dry-run or live). `inserted_ids` / `undone_at` power the
  // per-import Undo button on the History panel.
  wrapped.exec(`CREATE TABLE IF NOT EXISTS import_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    module TEXT NOT NULL,
    filename TEXT DEFAULT '',
    dry_run INTEGER DEFAULT 0,
    total INTEGER DEFAULT 0,
    imported INTEGER DEFAULT 0,
    skipped INTEGER DEFAULT 0,
    failed INTEGER DEFAULT 0,
    errors TEXT DEFAULT '',
    inserted_ids TEXT DEFAULT '',
    undone_at TEXT DEFAULT '',
    user_id INTEGER,
    username TEXT DEFAULT '',
    created_at TEXT DEFAULT (datetime('now','localtime'))
  )`);

  // ── INDEXES ────────────────────────────────────────────────
  const indexes = [
    'CREATE INDEX IF NOT EXISTS idx_traders_name ON traders(name)',
    'CREATE INDEX IF NOT EXISTS idx_lots_auction ON lots(auction_id)',
    'CREATE INDEX IF NOT EXISTS idx_lots_lot ON lots(lot_no)',
    'CREATE INDEX IF NOT EXISTS idx_lots_name ON lots(name)',
    'CREATE INDEX IF NOT EXISTS idx_lots_buyer ON lots(buyer)',
    'CREATE INDEX IF NOT EXISTS idx_lots_sale ON lots(sale)',
    // Cascade lock lookups (lotsLockedFor*) scan by (auction_id, buyer)
    // or (auction_id, name) with locked_at NOT NULL — the auction index
    // above already covers the leading column, so a dedicated partial-
    // ish index isn't necessary. We add this one only so the lots-tab
    // list can sort/filter locked rows cheaply.
    'CREATE INDEX IF NOT EXISTS idx_lots_locked ON lots(locked_at)',
    'CREATE INDEX IF NOT EXISTS idx_invoices_date ON invoices(date)',
    'CREATE INDEX IF NOT EXISTS idx_invoices_sale ON invoices(sale, invo)',
    'CREATE INDEX IF NOT EXISTS idx_purchases_date ON purchases(date)',
    'CREATE INDEX IF NOT EXISTS idx_purchases_name ON purchases(name)',
    'CREATE INDEX IF NOT EXISTS idx_bills_date ON bills(date)',
    'CREATE INDEX IF NOT EXISTS idx_bills_name ON bills(name)',
    'CREATE INDEX IF NOT EXISTS idx_buyers_buyer ON buyers(buyer)',
    'CREATE INDEX IF NOT EXISTS idx_buyers_buyer1 ON buyers(buyer1)',
    'CREATE INDEX IF NOT EXISTS idx_lot_alloc_auction ON lot_allocations(auction_id)',
    // FK child-side index. SQLite auto-indexes the parent side of a FK
    // (traders.id is PK) but not the child column, so every DELETE FROM
    // traders triggers a full scan of trader_banks to check for orphans.
    // Without this, bulk seller deletion is O(N·M) — quadratic.
    'CREATE INDEX IF NOT EXISTS idx_trader_banks_trader ON trader_banks(trader_id)',
  ];
  for (const idx of indexes) { try { wrapped.exec(idx); } catch (e) {} }

  // ── MIGRATIONS (for existing databases created before schema changes) ──
  const migrations = [
    // Per-import undo: existing DBs need the two new columns added so
    // the Undo button on the History panel can find inserted rows and
    // mark the entry as rolled back. CREATE TABLE above already has
    // these; these ALTERs only matter on installs whose import_log was
    // created by an earlier build.
    "ALTER TABLE import_log ADD COLUMN inserted_ids TEXT DEFAULT ''",
    "ALTER TABLE import_log ADD COLUMN undone_at TEXT DEFAULT ''",
    // Price-check gate timestamp — see auctions CREATE TABLE for semantics.
    // Existing DBs without this column need the ALTER; ignored on fresh
    // installs where the column is already present.
    "ALTER TABLE auctions ADD COLUMN price_checked_at TEXT DEFAULT ''",
    // Tri-state price-check gate: existing DBs need this column to
    // distinguish never-checked (hard block) from checked-then-edited
    // (soft warning). See auctions CREATE TABLE for semantics.
    "ALTER TABLE auctions ADD COLUMN price_check_first_passed_at TEXT DEFAULT ''",
    // Backfill: any auction that was already verified BEFORE this column
    // existed gets its first-pass stamp set to the current-verify stamp.
    // Without this, every previously-verified auction would re-enter the
    // 'never' state on upgrade and force a one-off re-verify.
    "UPDATE auctions SET price_check_first_passed_at = price_checked_at WHERE price_checked_at IS NOT NULL AND price_checked_at != '' AND (price_check_first_passed_at IS NULL OR price_check_first_passed_at = '')",
    'ALTER TABLE purchases ADD COLUMN auction_id INTEGER',
    'ALTER TABLE invoices ADD COLUMN auction_id INTEGER',
    'ALTER TABLE bills ADD COLUMN auction_id INTEGER',
    'ALTER TABLE debit_notes ADD COLUMN auction_id INTEGER',
    "ALTER TABLE buyers ADD COLUMN code TEXT DEFAULT ''",
    "ALTER TABLE buyers ADD COLUMN cadd2 TEXT DEFAULT ''",
    "ALTER TABLE buyers ADD COLUMN email TEXT DEFAULT ''",
    "ALTER TABLE buyers ADD COLUMN tdsq TEXT DEFAULT ''",
    "ALTER TABLE buyers ADD COLUMN sbl TEXT DEFAULT ''",
    // Discount GST columns (per-lot, when flag_disc_gst is ON)
    'ALTER TABLE lots ADD COLUMN dcgst REAL DEFAULT 0',
    'ALTER TABLE lots ADD COLUMN dsgst REAL DEFAULT 0',
    'ALTER TABLE lots ADD COLUMN digst REAL DEFAULT 0',
    // ASP invoice traceability — when a lot is first invoiced as an ASP
    // sale (state=Kerala), `lots.invo` gets the ASP invoice number AND a
    // copy is preserved here. Then when the same lot is invoiced as an
    // ISP sale (state=Tamil Nadu) later, `lots.invo` is overwritten with
    // the ISP invoice number, but `lots.asp_invo` keeps the original ASP
    // ref. This lets the sales list show both numbers side-by-side.
    "ALTER TABLE lots ADD COLUMN asp_invo TEXT DEFAULT ''",
    // ── Dual-view planter calculation columns ─────────────────
    // calculateLot() chooses ISP vs ASP rules based on cfg.business_state,
    // then writes the active view into pqty/prate/puramt. The Tally URD
    // voucher needs ISP values regardless of which mode dad is currently in,
    // so we now ALWAYS persist BOTH calculations on every save:
    //   isp_pqty/isp_prate/isp_puramt → planter side as ISP would compute
    //   asp_pqty/asp_prate/asp_puramt → planter side as ASP would compute
    // The legacy pqty/prate/puramt columns continue to mirror whichever
    // matches the current business_state, so the existing UI / reports /
    // exports keep working unchanged. Reports that need a specific view
    // (like the URD Tally voucher) read the prefixed columns directly.
    'ALTER TABLE lots ADD COLUMN isp_pqty REAL DEFAULT 0',
    'ALTER TABLE lots ADD COLUMN isp_prate REAL DEFAULT 0',
    'ALTER TABLE lots ADD COLUMN isp_puramt REAL DEFAULT 0',
    'ALTER TABLE lots ADD COLUMN asp_pqty REAL DEFAULT 0',
    'ALTER TABLE lots ADD COLUMN asp_prate REAL DEFAULT 0',
    'ALTER TABLE lots ADD COLUMN asp_puramt REAL DEFAULT 0',
    // Distance for e-way bill <DISTANCE> field on ISP sales vouchers.
    // Populated manually per-invoice from the To Tally → 🗺️ E-way Bill
    // Distance UI: user looks up the value on NIC's Pin-to-Pin Distance
    // Search page (or Google Maps), pastes it here, clicks Save. Value
    // is then emitted verbatim on the next voucher regen.
    'ALTER TABLE invoices ADD COLUMN distance_km INTEGER',
    // Additional Charge — sum(cardamom) × cfg.addl_charge_value, sits
    // below the Round on/off line. addl_name carries the user-defined
    // ledger label (also used as the Tally ledger name in XML).
    "ALTER TABLE invoices ADD COLUMN addl_chg REAL DEFAULT 0",
    "ALTER TABLE invoices ADD COLUMN addl_name TEXT DEFAULT ''",
    // Per-invoice lorry / truck number. Set from the Invoices tab via a
    // bulk-action button; emitted into the e-way bill <VEHICLENUMBER>
    // (and BASICSHIPVESSELNO) fields when generating sales vouchers
    // so dad doesn't have to type it into Tally manually.
    'ALTER TABLE invoices ADD COLUMN lorry_no TEXT',
    // Drop legacy pincodes/pin_distances tables that supported the old
    // haversine auto-compute path. We replaced that with the manual-
    // override workflow (above), so these tables are now dead weight.
    // IF EXISTS makes this idempotent — fresh DBs have nothing to drop;
    // upgraded DBs shed the orphan tables on next restart.
    'DROP TABLE IF EXISTS pin_distances',
    'DROP TABLE IF EXISTS pincodes',
    // Force-rotate flag for the seeded admin (and any future password reset
    // by an admin). When non-zero, requireAuth blocks every endpoint except
    // the change-password / whoami routes until the user picks a new password.
    'ALTER TABLE users ADD COLUMN must_change_password INTEGER NOT NULL DEFAULT 0',
    // Hard expiry cap for sessions. Existing pre-migration rows get NULL
    // (grandfathered in — relies on the 30-day idle sweep). New rows get
    // a 30-day cap from creation; requireAuth refuses tokens past expires_at
    // even if last_used_at is recent, so a leaked Authorization header has
    // a bounded validity window.
    'ALTER TABLE sessions ADD COLUMN expires_at TEXT',
    // Per-lot record lock. When locked_at is set, the row becomes
    // uneditable for non-admins, both directly (PUT/DELETE /api/lots/:id,
    // bulk lot mutations, calculate) and indirectly (sales invoice /
    // purchase / debit-note edit/delete/revert that would touch the lot).
    // Admins can always edit. Only an admin can clear the lock.
    // locked_by carries the username of whoever set the lock — purely
    // for audit/UI display; permission is decided from req.user.role.
    'ALTER TABLE lots ADD COLUMN locked_at TEXT DEFAULT NULL',
    'ALTER TABLE lots ADD COLUMN locked_by TEXT DEFAULT NULL',
    // Branch name for a seller's bank account. Auto-populated in the
    // Sellers add/edit modal when the user types a valid IFSC (Razorpay
    // public IFSC API). Saved so reports / invoices can show the branch
    // without re-querying the network.
    "ALTER TABLE trader_banks ADD COLUMN branch TEXT DEFAULT ''",
    // Extra lot-entry fields — surfaced in the form behind the
    // `show_extra_lot_fields` setting. `crop_receipt_no` is a per-trade
    // running counter auto-assigned on INSERT; `weight_with_gunny` is
    // the weight including the empty bag and feeds the auto-calc
    // (net = weight_with_gunny - gunny_weight) when the toggle is on;
    // `reserved_price` is the seller's minimum acceptable price and is
    // shown in Recent Entries for reference (not used in calc/invoice
    // gating).
    'ALTER TABLE lots ADD COLUMN weight_with_gunny REAL DEFAULT 0',
    'ALTER TABLE lots ADD COLUMN crop_receipt_no INTEGER DEFAULT NULL',
    'ALTER TABLE lots ADD COLUMN reserved_price REAL DEFAULT 0',
  ];
  for (const m of migrations) {
    try { wrapped.exec(m); console.log('Migration applied:', m); }
    catch (e) { /* column already exists — ignore */ }
  }

  // One-time backfill: assign crop_receipt_no to every lot that doesn't
  // have one yet, per auction, ordered by id (creation order). Without
  // this, existing rows would all read NULL and the new Recent Entries
  // "Receipt" column would be blank for historical data. Idempotent:
  // only rows where crop_receipt_no IS NULL are touched, so re-runs are
  // no-ops once the column is populated.
  try {
    const auctionsNeedingBackfill = wrapped.all(
      `SELECT DISTINCT auction_id FROM lots WHERE crop_receipt_no IS NULL`
    );
    for (const { auction_id } of auctionsNeedingBackfill) {
      if (!auction_id) continue;
      const startRow = wrapped.get(
        `SELECT COALESCE(MAX(crop_receipt_no), 0) AS m FROM lots WHERE auction_id = ?`,
        [auction_id]
      );
      let next = (startRow && startRow.m) || 0;
      const rows = wrapped.all(
        `SELECT id FROM lots WHERE auction_id = ? AND crop_receipt_no IS NULL ORDER BY id`,
        [auction_id]
      );
      for (const r of rows) {
        next += 1;
        wrapped.run(`UPDATE lots SET crop_receipt_no = ? WHERE id = ?`, [next, r.id]);
      }
    }
  } catch (e) { /* fresh DB — no lots yet, nothing to backfill */ }

  // Re-attempt indexes that depend on migration-added columns. The
  // indexes block above runs BEFORE migrations, so on upgrade-paths
  // where the column didn't exist yet the CREATE INDEX silently
  // no-oped. Idempotent — already-created indexes are a no-op here too.
  try { wrapped.exec('CREATE INDEX IF NOT EXISTS idx_lots_locked ON lots(locked_at)'); } catch (_) {}

  // One-time data fix: legacy ASP-only lots (where invo==asp_invo) had their
  // `sale` field set during the old ASP-generation logic. The current logic
  // doesn't set it (so ISP can pick the right sale type per buyer). Clear
  // those legacy rows so they show up in ISP eligibility.
  // Idempotent: subsequent runs do nothing because the rows are already
  // cleared. Safe to run on a fresh DB (no rows match the WHERE).
  try {
    const fix = wrapped.run(
      `UPDATE lots SET sale = ''
       WHERE asp_invo IS NOT NULL AND asp_invo != ''
         AND invo = asp_invo
         AND sale IS NOT NULL AND sale != ''`
    );
    if (fix && fix.changes > 0) {
      console.log(`Migration: cleared sale on ${fix.changes} ASP-only lots so ISP eligibility works`);
    }
  } catch (e) { /* ignore — column may not exist on first run */ }

  // One-time data fix: legacy invoices were stamped with the auction's
  // state (TAMIL NADU/KERALA based on physical auction location), not the
  // business context state. Retag them so the sales list can correctly
  // distinguish ASP rows from ISP rows.
  //
  // Heuristic per invoice:
  //   - If the invoice's `invo` equals `lots.asp_invo` for any of its
  //     buyer/auction lots AND those lots' current `invo` differs from
  //     `asp_invo` → invoice was the ASP step (stamp KERALA).
  //   - Else if any of those lots have `asp_invo == invo == this invoice's
  //     invo` → ASP-only run (stamp KERALA).
  //   - Otherwise → ISP invoice (stamp TAMIL NADU).
  // Safe-by-default: only updates rows we can confidently classify.
  // Idempotent: re-running produces the same labels.
  try {
    const allInvs = wrapped.all('SELECT id, auction_id, buyer, invo FROM invoices');
    let aspCount = 0, ispCount = 0;
    for (const inv of allInvs) {
      const lotMatches = wrapped.all(
        `SELECT invo, asp_invo FROM lots
         WHERE auction_id = ? AND buyer = ?
           AND (invo = ? OR asp_invo = ?)`,
        [inv.auction_id, inv.buyer, inv.invo, inv.invo]
      );
      // Determine state: ASP if this invoice matches asp_invo on any lot,
      // ISP otherwise (lots have a different ISP invo and asp_invo links
      // back to this row).
      let isASP = false;
      for (const l of lotMatches) {
        if (l.asp_invo === inv.invo) { isASP = true; break; }
        // If lot's invo == this inv's invo AND lot's asp_invo is empty,
        // this is most likely an ISP-only invoice — but COULD be an ASP
        // run pre-asp_invo column. Default to ISP context (TN).
      }
      const newState = isASP ? 'KERALA' : 'TAMIL NADU';
      wrapped.run('UPDATE invoices SET state = ? WHERE id = ?', [newState, inv.id]);
      if (isASP) aspCount++; else ispCount++;
    }
    if (aspCount + ispCount > 0) {
      console.log(`Migration: retagged ${aspCount} invoices as KERALA (ASP) and ${ispCount} as TAMIL NADU (ISP) based on lot lineage`);
    }
  } catch (e) { /* table may not exist on fresh DB — ignore */ }

  const row = wrapped.get('SELECT COUNT(*) as cnt FROM users');
  if (!row || row.cnt === 0) {
    // bcrypt cost 12 — single hash at boot adds ~250ms, only on first run.
    const hash = bcrypt.hashSync('admin123', 12);
    wrapped.run(
      'INSERT INTO users (username, password_hash, role, must_change_password) VALUES (?, ?, ?, 1)',
      ['admin', hash, 'admin']
    );
    console.log('Default admin created (admin / admin123) — MUST be changed on first login');
  }

  console.log('Database ready at', DB_PATH, '(better-sqlite3, WAL mode)');
  return wrapped;
}

/**
 * Normalize params so callers can pass either an array or spread arguments.
 * Accepts: fn('sql', [a, b, c])  OR  fn('sql', a, b, c)  OR  fn('sql')
 */
function normalizeParams(args) {
  if (args.length === 0) return [];
  if (args.length === 1 && Array.isArray(args[0])) return args[0];
  return args;
}

function makeWrapper() {
  // sql.js note: prepared statements need .free() to release memory.
  // We create-use-free per call to keep the API simple and match the
  // existing usage patterns (no long-lived prepared statements).

  /**
   * Run a SQL with bound params and return rows as objects.
   * Internal helper used by get/all.
   */
  function execStatement(sql, params) {
    const stmt = rawDb.prepare(sql);
    try {
      stmt.bind(params);
      const rows = [];
      while (stmt.step()) rows.push(stmt.getAsObject());
      return rows;
    } finally {
      stmt.free();
    }
  }

  /**
   * Run an INSERT/UPDATE/DELETE/etc with bound params (no result rows).
   */
  function runStatement(sql, params) {
    const stmt = rawDb.prepare(sql);
    try {
      stmt.run(params);
    } finally {
      stmt.free();
    }
    // sql.js doesn't expose lastInsertRowid/changes per-statement easily.
    // Use the connection-level helpers.
    return {
      lastInsertRowid: rawDb.exec("SELECT last_insert_rowid()")[0]?.values[0]?.[0] ?? 0,
      changes: rawDb.getRowsModified(),
    };
  }

  return {
    /**
     * Execute multi-statement SQL (no params, no return).
     */
    exec(sql) {
      rawDb.exec(sql);
      scheduleSave();
    },

    /**
     * Run an INSERT/UPDATE/DELETE. Accepts params as array or spread.
     */
    run(sql, ...rest) {
      const params = normalizeParams(rest);
      const info = runStatement(sql, params);
      scheduleSave();
      return info;
    },

    /**
     * SELECT one row. Returns row object or null.
     */
    get(sql, ...rest) {
      const params = normalizeParams(rest);
      const rows = execStatement(sql, params);
      return rows[0] || null;
    },

    /**
     * SELECT many rows. Returns array (possibly empty).
     */
    all(sql, ...rest) {
      const params = normalizeParams(rest);
      return execStatement(sql, params);
    },

    /**
     * Prepare a statement. sql.js doesn't naturally cache prepared
     * statements across reuses (and freeing too early causes errors),
     * so we re-prepare on each call. Slower than better-sqlite3 but
     * functionally equivalent.
     */
    prepare(sql) {
      return {
        run(...args) {
          const info = runStatement(sql, args);
          scheduleSave();
          return info;
        },
        get(...args) {
          const rows = execStatement(sql, args);
          return rows[0] || null;
        },
        all(...args) {
          return execStatement(sql, args);
        }
      };
    },

    /**
     * Wrap a function in a transaction. Implements via BEGIN/COMMIT/ROLLBACK.
     */
    transaction(fn) {
      return function (...args) {
        rawDb.run("BEGIN");
        try {
          const result = fn(...args);
          rawDb.run("COMMIT");
          scheduleSave();
          return result;
        } catch (e) {
          rawDb.run("ROLLBACK");
          throw e;
        }
      };
    },

    // Escape hatch — only for code that needs the raw sql.js Database.
    get raw() { return rawDb; }
  };
}

function getDb() {
  if (!wrapped) throw new Error('Call initDb() first');
  return wrapped;
}

function closeDb() {
  flushSave();
  if (rawDb) {
    rawDb.close();
    rawDb = null;
    wrapped = null;
  }
}

/**
 * Replace the entire database from a buffer (used by /api/system/restore).
 * Validates that the buffer is a SQLite file (header magic), opens it as a
 * fresh sql.js Database, persists to disk, then swaps it in. Throws on any
 * error so the caller can surface a clean message — the existing DB is left
 * untouched on failure.
 */
async function replaceFromBuffer(buf) {
  if (!Buffer.isBuffer(buf)) buf = Buffer.from(buf);
  // SQLite files start with "SQLite format 3\0" (16 bytes).
  const magic = buf.slice(0, 16).toString('utf8').replace(/\0+$/, '');
  if (magic !== 'SQLite format 3') {
    throw new Error('Uploaded file is not a valid SQLite database');
  }
  if (!SQL) SQL = await initSqlJs(_resolveSqlJsOpts());
  // Sanity-check: must open without throwing.
  let test;
  try { test = new SQL.Database(buf); } catch (e) {
    throw new Error('SQLite file is corrupt or unreadable: ' + e.message);
  }
  // Quick integrity check
  try {
    const r = test.exec('PRAGMA integrity_check');
    const ok = r && r[0] && r[0].values && r[0].values[0] && r[0].values[0][0] === 'ok';
    if (!ok) throw new Error('integrity_check did not return "ok"');
  } catch (e) {
    test.close();
    throw new Error('Integrity check failed: ' + e.message);
  }
  // Flush any pending writes from the current DB before replacing it.
  flushSave();
  // Close the live DB and swap.
  if (rawDb) { try { rawDb.close(); } catch(_){} }
  rawDb = test;
  // Persist immediately so a crash right after restore doesn't lose it.
  try {
    const out = Buffer.from(rawDb.export());
    const tmp = DB_PATH + '.tmp';
    fs.writeFileSync(tmp, out);
    fs.renameSync(tmp, DB_PATH);
  } catch (e) {
    throw new Error('Failed to persist restored DB: ' + e.message);
  }
  return { ok: true, size: buf.length };
}

module.exports = { initDb, getDb, closeDb, flushSave, replaceFromBuffer, DB_PATH };
