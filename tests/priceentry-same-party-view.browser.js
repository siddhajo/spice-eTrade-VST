// Price Entry — "⚠ N seller = buyer" is a view, not just a count.
//
// The badge has always told the operator HOW MANY lots carry the same party
// as seller and buyer. On a 400-lot trade that leaves the actual job — find
// those 15 rows — to the eye and the scroll wheel. Clicking the badge now
// filters the grid to exactly those lots, and clicking it again restores it.
//
//   [off]      the badge starts unpressed and the grid shows the usual rows
//   [on]       clicking it leaves ONLY the same-party lots on screen
//   [complete] including ones already priced WITH a buyer — the rows the
//              "hide completed" rule would otherwise swallow, and the very
//              ones most worth seeing
//   [filters]  entering the view clears the column filters, so nothing is
//              still hiding part of the answer
//   [back]     clicking again restores the grid
//   [none]     a trade with none says so, and offers the way back
//
// The grid state is injected directly (_pe.lots + renderPriceEntryGrid) —
// this covers the view logic and the real DOM, not the lot-loading pipeline,
// which is untouched.
//
// Run (puppeteer-core is not a dependency of this build — point NODE_PATH at
// any checkout that has it):
//   NODE_PATH=../spice-config-merged-with-pwa/node_modules \
//     node tests/priceentry-same-party-view.browser.js
const os = require('os'), path = require('path'), fs = require('fs');
const { spawn } = require('child_process');
const pptr = require('puppeteer-core');

const ROOT = path.join(__dirname, '..');
const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'pesp-'));
const PORT = 48185;
const B = `http://127.0.0.1:${PORT}`;

let pass = 0, fail = 0;
const check = (n, c, d) => {
  if (c) { pass++; console.log('  ok   ' + n); }
  else { fail++; console.log('  FAIL ' + n + (d !== undefined ? '\n         ' + JSON.stringify(d).slice(0, 300) : '')); }
};

const srv = spawn('node', [path.join(ROOT, 'server.js')], {
  cwd: ROOT,
  env: Object.assign({}, process.env, { SPICE_DATA_DIR: TMP, PORT: String(PORT), NODE_ENV: 'test' }),
  stdio: ['ignore', 'pipe', 'pipe'],
});
let browser = null;
const cleanup = () => {
  try { if (browser) browser.close(); } catch (_) {}
  try { srv.kill('SIGKILL'); } catch (_) {}
  try { fs.rmSync(TMP, { recursive: true, force: true }); } catch (_) {}
};

// Six lots: three same-party (one of them already complete — priced AND with
// a buyer saved), three ordinary.
const LOTS = [
  { id: 1, lot_no: '001', name: 'A.R. Traders',  branch: 'BODI', bags: 10, qty: 100, price: 0,   code: '',    buyer: 'B042', buyer1: 'AR TRADERS' },
  { id: 2, lot_no: '002', name: 'Kumar & Sons',  branch: 'BODI', bags: 8,  qty: 80,  price: 900, code: 'S',   buyer: 'B011', buyer1: 'Vasan Spices' },
  { id: 3, lot_no: '003', name: 'Selvam Estate', branch: 'CUMB', bags: 5,  qty: 50,  price: 0,   code: '',    buyer: '',     buyer1: '' },
  // Already complete: priced and a buyer saved — "hide completed" would drop
  // this one, which is exactly the row an operator needs to find.
  { id: 4, lot_no: '004', name: 'Selvam Estate', branch: 'CUMB', bags: 6,  qty: 60,  price: 850, code: 'S',   buyer: 'B077', buyer1: 'SELVAM ESTATE' },
  { id: 5, lot_no: '005', name: 'Murugan Co',    branch: 'BODI', bags: 4,  qty: 40,  price: 0,   code: '',    buyer: '',     buyer1: '' },
  // Same party via the CODE column, which some imported datasets carry a
  // name in rather than a code.
  { id: 6, lot_no: '006', name: 'Devi Traders',  branch: 'CUMB', bags: 7,  qty: 70,  price: 0,   code: '',    buyer: 'DEVI TRADERS', buyer1: '' },
];
const SAME = ['001', '004', '006'];

(async () => {
  for (let i = 0; i < 200; i++) {
    try { const r = await fetch(B + '/api/health'); if (r.status < 500) break; } catch (_) {}
    await new Promise(r => setTimeout(r, 250));
  }
  // This build refuses every write until the seeded admin's forced password
  // change is cleared; the grid here is injected, but signing in still has
  // to get past it.
  const signIn = async (password) => (await (await fetch(B + '/api/login', { method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ username: 'admin', password }) })).json());
  const first = await signIn('admin123');
  if (first.must_change_password) {
    await fetch(B + '/api/me/password', { method: 'PUT',
      headers: { 'Content-Type': 'application/json', Authorization: 'Bearer ' + first.token },
      body: JSON.stringify({ current_password: 'admin123', new_password: 'peTest1234' }) });
  }
  const T = (first.must_change_password ? await signIn('peTest1234') : first).token;

  let chrome = null;
  for (const p of [process.env.PUPPETEER_EXECUTABLE_PATH,
    '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
    '/Applications/Chromium.app/Contents/MacOS/Chromium'].filter(Boolean)) {
    try { if (fs.existsSync(p)) { chrome = p; break; } } catch (_) {}
  }
  if (!chrome) {
    console.log('  skip no Chrome available');
    console.log(`\n${pass} passed, ${fail} failed\n`); cleanup(); process.exit(0);
  }
  browser = await pptr.launch({ executablePath: chrome, args: ['--no-sandbox'], headless: true });
  const page = await browser.newPage();
  page.on('pageerror', e => { fail++; console.log('  FAIL page error: ' + e.message); });
  await page.setViewport({ width: 1500, height: 900 });

  await page.goto(B + '/', { waitUntil: 'domcontentloaded' });
  await page.evaluate(t => localStorage.setItem('t', t), T);
  await page.goto(B + '/', { waitUntil: 'networkidle2' });
  await new Promise(r => setTimeout(r, 1800));

  // Open the real Price Entry screen and let its own loader finish — it
  // fetches the trade list and repaints the grid a beat after go() returns,
  // which would otherwise wipe the lots injected below.
  await page.evaluate(() => go('priceentry'));
  await new Promise(r => setTimeout(r, 2500));
  const mounted = await page.evaluate((lots) => {
    _pe.aid = 1;
    _pe.lots = lots;
    _pe.byId = new Map(lots.map(l => [l.id, l]));
    _pe.front = []; _pe.dirty = new Set(); _pe.sel = new Set();
    _pe.showAll = false; _pe.sameOnly = false;
    _pe.filter = { lot: '', branch: '', seller: '' };
    renderPriceEntryGrid();
    return { hasGrid: !!document.getElementById('pe-grid'),
             hasSummary: !!document.getElementById('pe-summary') };
  }, LOTS);
  check('the Price Entry grid is on screen', mounted.hasGrid && mounted.hasSummary, mounted);

  // What the grid is showing, and what the badge looks like.
  const state = () => page.evaluate(() => {
    const rows = Array.from(document.querySelectorAll('#pe-grid tbody tr[data-lotid]'));
    const chip = document.querySelector('.pe-same-chip');
    return {
      lotNos: rows.map(r => r.children[1].textContent.trim()),
      flagged: rows.filter(r => r.classList.contains('pe-same-party'))
                   .map(r => r.children[1].textContent.trim()),
      chipText: chip ? chip.textContent.replace(/\s+/g, ' ').trim() : null,
      chipOn: chip ? chip.classList.contains('on') : null,
      pressed: chip ? chip.getAttribute('aria-pressed') : null,
      isButton: chip ? chip.tagName : null,
      summary: (document.getElementById('pe-summary') || {}).textContent || '',
      emptyText: (document.querySelector('#pe-grid tbody tr:not([data-lotid]) td') || {}).textContent || '',
      sameOnly: _pe.sameOnly,
    };
  });

  console.log('[off] the badge starts unpressed, counting across the trade');
  let s = await state();
  // Lot 004 is complete, so it is hidden here — and still counted.
  check('it counts every same-party lot, hidden ones included',
        /⚠ 3 seller = buyer/.test(s.chipText || ''), s.chipText);
  check('it is a button, not just text', s.isButton === 'BUTTON', s.isButton);
  check('…and starts unpressed', s.chipOn === false && s.pressed === 'false', s);
  // 002 and 004 are complete (priced + buyer), so both are hidden here.
  check('the grid is showing the ordinary working set', s.lotNos.join(',') === '001,003,005,006', s.lotNos);

  console.log('\n[on] clicking it leaves only those lots');
  await page.click('.pe-same-chip');
  await new Promise(r => setTimeout(r, 200));
  s = await state();
  check('only the same-party lots are on screen', s.lotNos.join(',') === SAME.join(','), s.lotNos);
  check('…and every row on screen is flagged', s.flagged.join(',') === SAME.join(','), s.flagged);
  check('the badge reads as pressed', s.chipOn === true && s.pressed === 'true', s);
  check('…and offers the way out', /✕/.test(s.chipText || ''), s.chipText);
  check('the summary says the grid is narrowed', /showing these only/.test(s.summary), s.summary);
  // Lot 004 is finished, so a footer reading "3 incomplete lots" would be a
  // lie — this view hides nothing on completeness grounds.
  const foot = await page.evaluate(() => (document.getElementById('pe-tot-lots') || {}).textContent || '');
  check('…and the footer does not call a finished lot incomplete',
        /Total · 3 lots/.test(foot), foot);

  console.log('\n[complete] a finished same-party lot is NOT hidden');
  // Lot 004 is priced AND has a buyer — "complete", so the default view
  // hides it. It is the one most worth seeing here, so this view overrides
  // that rule rather than inheriting it.
  check('the already-priced same-party lot is shown', s.lotNos.includes('004'), s.lotNos);
  const showAllUntouched = await page.evaluate(() => ({
    showAll: _pe.showAll,
    boxTicked: document.getElementById('pe-show-all').checked,
  }));
  check('…without having to tick "Show all"',
        showAllUntouched.showAll === false && showAllUntouched.boxTicked === false, showAllUntouched);

  console.log('\n[filters] entering the view clears what was hiding rows');
  await page.evaluate(() => {
    peToggleSameOnly(false);
    const lf = document.getElementById('pe-filter-lot');
    lf.value = '002'; peSetFilter('lot', '002');
    peSetFilter('branch', 'BODI');
  });
  await new Promise(r => setTimeout(r, 150));
  const narrowed = await state();
  check('a column filter is narrowing the grid first', narrowed.lotNos.length <= 1, narrowed.lotNos);
  await page.click('.pe-same-chip');
  await new Promise(r => setTimeout(r, 200));
  s = await state();
  check('the view still shows all three', s.lotNos.join(',') === SAME.join(','), s.lotNos);
  const boxes = await page.evaluate(() => ({
    lot: document.getElementById('pe-filter-lot').value,
    branch: document.getElementById('pe-filter-branch').value,
    seller: document.getElementById('pe-filter-seller').value,
    f: JSON.parse(JSON.stringify(_pe.filter)),
  }));
  check('…and the filter boxes were emptied, not just ignored',
        boxes.lot === '' && boxes.branch === '' && boxes.seller === ''
        && boxes.f.lot === '' && boxes.f.branch === '' && boxes.f.seller === '', boxes);

  console.log('\n[back] clicking again restores the grid');
  await page.click('.pe-same-chip');
  await new Promise(r => setTimeout(r, 200));
  s = await state();
  check('the ordinary working set is back', s.lotNos.join(',') === '001,003,005,006', s.lotNos);
  check('the badge is unpressed again', s.chipOn === false && s.sameOnly === false, s);
  // "Clear filters" is the other way out, and it must not leave the operator
  // stuck in a narrowed grid.
  await page.evaluate(() => { peToggleSameOnly(true); peClearFilters(); });
  await new Promise(r => setTimeout(r, 150));
  s = await state();
  check('"Clear filters" also leaves the view', s.sameOnly === false
        && s.lotNos.join(',') === '001,003,005,006', s.lotNos);

  console.log('\n[none] a clean trade says so, and offers the way back');
  const clean = await page.evaluate(() => {
    _pe.lots = _pe.lots.filter(l => !['001', '004', '006'].includes(l.lot_no));
    _pe.byId = new Map(_pe.lots.map(l => [l.id, l]));
    renderPriceEntryGrid();
    const before = !!document.querySelector('.pe-same-chip');
    // With none left the badge is gone — but if the operator is already
    // INSIDE the view it has to stay, or there is no way out but hunting
    // for "Clear filters".
    _pe.sameOnly = true;
    renderPriceEntryGrid();
    const chip = document.querySelector('.pe-same-chip');
    return { badgeGoneWhenClean: before === false,
             chipKeptInView: !!chip,
             chipText: chip ? chip.textContent.replace(/\s+/g, ' ').trim() : null,
             empty: (document.querySelector('#pe-grid tbody tr:not([data-lotid]) td') || {}).textContent || '' };
  });
  check('no same-party lots → no badge', clean.badgeGoneWhenClean, clean);
  check('…but inside the view the badge stays, so it can be switched off',
        clean.chipKeptInView && /0 seller = buyer/.test(clean.chipText || ''), clean);
  check('the empty grid explains itself', /same party as seller and buyer/.test(clean.empty), clean.empty);
  check('…and offers a way back', /Back to all lots/.test(clean.empty), clean.empty);
  await page.evaluate(() => {
    const a = Array.from(document.querySelectorAll('#pe-grid a'))
      .find(x => /Back to all lots/.test(x.textContent));
    if (a) a.click();
  });
  await new Promise(r => setTimeout(r, 200));
  check('the "Back to all lots" link works',
        (await page.evaluate(() => _pe.sameOnly)) === false);

  if (process.env.SHOT_DIR) {
    await page.evaluate((lots) => {
      _pe.lots = lots;
      _pe.byId = new Map(lots.map(l => [l.id, l]));
      peToggleSameOnly(true);
    }, LOTS);
    await new Promise(r => setTimeout(r, 400));
    await page.screenshot({ path: process.env.SHOT_DIR + '/pe-same-party-on.png' });
    await page.evaluate(() => peToggleSameOnly(false));
    await new Promise(r => setTimeout(r, 400));
    await page.screenshot({ path: process.env.SHOT_DIR + '/pe-same-party-off.png' });
  }
  console.log(`\n${pass} passed, ${fail} failed\n`);
  cleanup();
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error('ERR', e); cleanup(); process.exit(1); });
