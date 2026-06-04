(async () => {
  const dbm = require('./db.js'); await dbm.initDb(); const db = dbm.getDb();
  // For each duplicated buyer code, show invoice.buyer1 vs the resolved master row's buyer1+pin
  const dups = db.all("SELECT buyer FROM buyers GROUP BY buyer HAVING COUNT(*)>1");
  for (const {buyer} of dups) {
    const invs = db.all("SELECT auction_id, invo, buyer1 FROM invoices WHERE buyer=?",[buyer]);
    const masters = db.all("SELECT id, buyer1, pin, cpin FROM buyers WHERE buyer=? ORDER BY id",[buyer]);
    console.log(`\nbuyer code [${buyer}] masters:`, masters.map(m=>`id${m.id}:${m.buyer1}/pin${m.pin}`).join('  '));
    for (const iv of invs) {
      const resolved = db.get(`SELECT x.id, x.buyer1, x.pin FROM buyers x WHERE x.id = COALESCE(
        (SELECT y.id FROM buyers y WHERE y.buyer=? AND UPPER(TRIM(y.buyer1))=UPPER(TRIM(?)) ORDER BY y.id LIMIT 1),
        (SELECT y.id FROM buyers y WHERE y.buyer=? ORDER BY y.id LIMIT 1))`, [buyer, iv.buyer1, buyer]);
      const ok = resolved && resolved.buyer1 && iv.buyer1 && resolved.buyer1.trim().toUpperCase()===iv.buyer1.trim().toUpperCase();
      console.log(`   inv${iv.invo} buyer1=[${iv.buyer1}] -> resolved id${resolved&&resolved.id} [${resolved&&resolved.buyer1}] ${ok?'MATCH':'(fallback/lowest-id)'}`);
    }
  }
  process.exit(0);
})().catch(e=>{console.error(e);process.exit(1);});
