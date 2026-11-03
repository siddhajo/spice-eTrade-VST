// Sidebar groups — one open at a time.
//
// Seven groups and some thirty items: with every group open the menu ran
// past the fold on a laptop, so the screen you wanted was a scroll away.
// Exactly one group is open now — the one holding the screen you are on.
//
//   [boot]     only the group holding the current screen is open
//   [open]     clicking another group's label closes the rest
//   [select]   choosing an item inside it leaves it open, rest shut
//   [switch]   navigating to another group's screen moves the open group
//   [shut]     clicking the open group's own label closes everything
//   [rail]     the icon-only rail keeps every icon — no group is collapsed
//              there, because there would be no label left to click
//
// Run (puppeteer-core is not a dependency of this build — point NODE_PATH at
// any checkout that has it):
//   NODE_PATH=../spice-config-merged-with-pwa/node_modules \
//     node tests/sidebar-accordion.browser.js
const os = require('os'), path = require('path'), fs = require('fs');
const { spawn } = require('child_process');
const pptr = require('puppeteer-core');

const ROOT = path.join(__dirname, '..');
const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'sbacc-'));
const PORT = 48186;
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

(async () => {
  for (let i = 0; i < 200; i++) {
    try { const r = await fetch(B + '/api/health'); if (r.status < 500) break; } catch (_) {}
    await new Promise(r => setTimeout(r, 250));
  }
  const signIn = async (password) => (await (await fetch(B + '/api/login', { method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ username: 'admin', password }) })).json());
  const first = await signIn('admin123');
  if (first.must_change_password) {
    await fetch(B + '/api/me/password', { method: 'PUT',
      headers: { 'Content-Type': 'application/json', Authorization: 'Bearer ' + first.token },
      body: JSON.stringify({ current_password: 'admin123', new_password: 'sbTest1234' }) });
  }
  const T = (first.must_change_password ? await signIn('sbTest1234') : first).token;

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
  // Full-width sidebar: the accordion deliberately stands down on the
  // icon-only rail, which is the shipped default (see [rail] below).
  await page.evaluate(t => { localStorage.setItem('t', t); localStorage.setItem('sidebarCollapsed', '0'); }, T);
  await page.goto(B + '/', { waitUntil: 'networkidle2' });
  await new Promise(r => setTimeout(r, 2200));

  // Every group, whether it is open, and how many of its items are really
  // on screen (offsetParent, not just the class — the whole point is that
  // the items go away).
  const groups = () => page.evaluate(() => Array.from(document.querySelectorAll('.side-group')).map(g => {
    const lbl = g.querySelector('.side-group-lbl');
    const items = Array.from(g.querySelectorAll('.side-item'));
    return {
      name: lbl ? lbl.textContent.trim() : '',
      open: !g.classList.contains('section-collapsed'),
      visibleItems: items.filter(i => i.offsetParent !== null).length,
      items: items.length,
      labelVisible: !!(lbl && lbl.offsetParent !== null),
    };
  }));
  const openNames = (gs) => gs.filter(g => g.open).map(g => g.name);
  const clickLabel = async (name) => {
    await page.evaluate((n) => {
      const lbl = Array.from(document.querySelectorAll('.side-group-lbl'))
        .find(l => l.textContent.trim() === n);
      if (lbl) lbl.click();
    }, name);
    await new Promise(r => setTimeout(r, 250));
  };

  console.log('[boot] only the group holding the current screen is open');
  let gs = await groups();
  check('the sidebar has its seven groups', gs.length === 7, gs.map(g => g.name));
  check('exactly one is open', openNames(gs).length === 1, openNames(gs));
  // The app boots onto the Dashboard, which lives in Overview.
  check('…and it is the one holding the current screen (Dashboard → Overview)',
        openNames(gs)[0] === 'Overview', openNames(gs));
  const shut = gs.filter(g => !g.open);
  check('every other group has nothing on screen under it',
        shut.every(g => g.visibleItems === 0), shut);
  check('…but their labels stay, so they can be opened',
        shut.every(g => g.labelVisible), shut.map(g => [g.name, g.labelVisible]));
  const overview = gs.find(g => g.name === 'Overview');
  check('the open group does show its items', overview.visibleItems > 0, overview);

  console.log('\n[open] clicking another label closes the rest');
  await clickLabel('Reports & Exports');
  gs = await groups();
  check('the clicked group opened', openNames(gs).join(',') === 'Reports & Exports', openNames(gs));
  check('…and the one that was open closed',
        gs.find(g => g.name === 'Overview').visibleItems === 0, gs.find(g => g.name === 'Overview'));

  console.log('\n[select] choosing an item leaves that group open, rest shut');
  // Click a real item inside the open group.
  const picked = await page.evaluate(() => {
    const g = Array.from(document.querySelectorAll('.side-group'))
      .find(x => (x.querySelector('.side-group-lbl') || {}).textContent.trim() === 'Reports & Exports');
    const item = Array.from(g.querySelectorAll('.side-item'))
      .find(i => i.offsetParent !== null && i.dataset.tab);
    if (item) item.click();
    return item ? item.dataset.tab : null;
  });
  await new Promise(r => setTimeout(r, 900));
  gs = await groups();
  check('a sidebar item was clicked', !!picked, picked);
  check('its group is still the only one open',
        openNames(gs).join(',') === 'Reports & Exports', openNames(gs));
  const activeIn = await page.evaluate(() => {
    const a = document.querySelector('.side-item.active');
    const g = a && a.closest('.side-group');
    return { tab: a && a.dataset.tab, group: g ? g.querySelector('.side-group-lbl').textContent.trim() : null,
             activeVisible: !!(a && a.offsetParent !== null) };
  });
  check('…and the screen you landed on is the visible, active one',
        activeIn.group === 'Reports & Exports' && activeIn.activeVisible, activeIn);

  console.log('\n[switch] navigating elsewhere moves the open group');
  // go() from outside the sidebar — a button, a deep link, the Refresh
  // path — has to move the accordion too, or the menu stops describing
  // where you are.
  await page.evaluate(() => go('traders'));
  await new Promise(r => setTimeout(r, 900));
  gs = await groups();
  check('the group holding the new screen opened',
        openNames(gs).join(',') === 'Master Data', openNames(gs));
  check('…and the group we came from is shut',
        gs.find(g => g.name === 'Reports & Exports').visibleItems === 0,
        gs.find(g => g.name === 'Reports & Exports'));
  check('only ever one open', openNames(gs).length === 1, openNames(gs));

  console.log('\n[shut] clicking the open group’s own label closes everything');
  await clickLabel('Master Data');
  gs = await groups();
  check('nothing is open', openNames(gs).length === 0, openNames(gs));
  check('…and no item is on screen anywhere',
        gs.every(g => g.visibleItems === 0), gs.map(g => [g.name, g.visibleItems]));
  check('every label is still there to click',
        gs.every(g => g.labelVisible), gs.map(g => [g.name, g.labelVisible]));
  // …and navigating brings the right one back.
  await page.evaluate(() => go('dash'));
  await new Promise(r => setTimeout(r, 900));
  gs = await groups();
  check('navigating re-opens the group for that screen',
        openNames(gs).join(',') === 'Overview', openNames(gs));

  console.log('\n[rail] the icon-only rail keeps every icon');
  // The rail hides the labels, so a collapsed group there would take its
  // icons away with nothing left to click. The accordion stands down.
  await page.evaluate(() => toggleSidebarCollapse());
  await new Promise(r => setTimeout(r, 400));
  const rail = await page.evaluate(() => {
    const gs = Array.from(document.querySelectorAll('.side-group'));
    return {
      collapsedRail: document.documentElement.classList.contains('sidebar-collapsed'),
      anyGroupCollapsed: gs.some(g => g.classList.contains('section-collapsed')),
      // Every item that is not permission- or role-hidden should still be
      // reachable as an icon.
      visibleItems: gs.reduce((n, g) => n + Array.from(g.querySelectorAll('.side-item'))
                                               .filter(i => i.offsetParent !== null).length, 0),
      labelsHidden: gs.every(g => { const l = g.querySelector('.side-group-lbl'); return !l || l.offsetParent === null; }),
    };
  });
  check('the sidebar is the icon rail', rail.collapsedRail && rail.labelsHidden, rail);
  check('no group is collapsed there', rail.anyGroupCollapsed === false, rail);
  check('…so every icon is still reachable', rail.visibleItems > 10, rail.visibleItems);

  console.log('\n[rail → full] expanding the sidebar restores the accordion');
  await page.evaluate(() => toggleSidebarCollapse());
  await new Promise(r => setTimeout(r, 400));
  gs = await groups();
  check('one group open again', openNames(gs).length === 1, openNames(gs));
  check('…and it is the one holding the current screen',
        openNames(gs)[0] === 'Overview', openNames(gs));

  if (process.env.SHOT_DIR) {
    await page.evaluate(() => go('invoices'));
    await new Promise(r => setTimeout(r, 1200));
    await page.screenshot({ path: process.env.SHOT_DIR + '/sidebar-accordion.png' });
  }
  console.log(`\n${pass} passed, ${fail} failed\n`);
  cleanup();
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error('ERR', e); cleanup(); process.exit(1); });
