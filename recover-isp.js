// One-time recovery: restore ISP defaults to the company.* fields in the DB.
// Run with: node recover-isp.js
//
// Use this if the DB got corrupted by saving while in KERALA state (which
// previously mutated company.* fields to ASP values before writing).

const Database = require('better-sqlite3');
const path = require('path');

const dbPath = path.join(__dirname, 'data', 'config.db');
const db = new Database(dbPath);

const ISP_DEFAULTS = {
  logo:       'ISP',
  trade_name: 'IDEAL SPICES',
  legal_name: ' PRIVATE LIMITED',
  short_name: 'IDEAL SPICES PRIVATE LIMITED',
  pan:        'AAICI5415L',
  cin:        'U47211TN2025PTC186657',
  fssai:      '',
  sbl:        '',
};

const upd = db.prepare('UPDATE company_settings SET value = ? WHERE key = ?');
let n = 0;
for (const [k, v] of Object.entries(ISP_DEFAULTS)) {
  const cur = db.prepare('SELECT value FROM company_settings WHERE key = ?').get(k);
  if (!cur) {
    console.log(`  · ${k}: not found in DB, skipping`);
    continue;
  }
  if (cur.value === v) {
    console.log(`  ✓ ${k}: already ${JSON.stringify(v)}`);
    continue;
  }
  upd.run(v, k);
  console.log(`  ↻ ${k}: was ${JSON.stringify(cur.value)} → now ${JSON.stringify(v)}`);
  n++;
}
console.log(`\nRecovery complete. ${n} field(s) restored to ISP defaults.`);
db.close();
