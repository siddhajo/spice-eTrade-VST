// Screensaver — the app's own name on an idle screen.
//
// An office machine sits on one screen all day — a priced lot grid, a
// seller's bank details — in a room people walk through. After a few quiet
// minutes this covers it; the first key or click clears it.
//
// The rule it must not break: it is NOT a sign-out and must never stand in
// for one. Nothing about the session is touched — it only covers the screen.
//
//   [off]      0 minutes means never, which is the shipped default
//   [show]     it appears on idle, carrying the app name
//   [clear]    a key or a click clears it
//   [lockup]   name, credit and numbers share one centre axis
//   [fx]       it animates, stays bounded, and stops dead
//   [login]    the product names itself on the sign-in page too
//
// Run (puppeteer-core is not a dependency of this build — point NODE_PATH at
// any checkout that has it):
//   NODE_PATH=../spice-config-merged-with-pwa/node_modules \
//     node tests/screensaver.browser.js
const os = require('os'), path = require('path'), fs = require('fs');
const { spawn } = require('child_process');
const pptr = require('puppeteer-core');

const ROOT = path.join(__dirname, '..');
const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'ss-'));
const PORT = 48184;
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
  // The seeded admin carries must_change_password, and this build refuses
  // EVERY write until that gate is cleared — so the fixture clears it and
  // signs in again, otherwise every settings PUT below silently 403s and
  // the whole run tests the shipped defaults.
  const signIn = async (password) => (await (await fetch(B + '/api/login', { method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ username: 'admin', password }) })).json());
  const first = await signIn('admin123');
  if (first.must_change_password) {
    const r = await fetch(B + '/api/me/password', { method: 'PUT',
      headers: { 'Content-Type': 'application/json', Authorization: 'Bearer ' + first.token },
      body: JSON.stringify({ current_password: 'admin123', new_password: 'ssTest1234' }) });
    if (!r.ok) { console.log('  FAIL could not clear the password gate'); cleanup(); process.exit(1); }
  }
  const T = (first.must_change_password ? await signIn('ssTest1234') : first).token;
  const put = (settings) => fetch(B + '/api/company-settings', {
    method: 'PUT', headers: { 'Content-Type': 'application/json', Authorization: 'Bearer ' + T },
    body: JSON.stringify({ settings }) });

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
  await page.setViewport({ width: 1400, height: 860 });

  const open = async () => {
    await page.goto(B + '/', { waitUntil: 'domcontentloaded' });
    await page.evaluate(t => localStorage.setItem('t', t), T);
    await page.goto(B + '/', { waitUntil: 'networkidle2' });
    await new Promise(r => setTimeout(r, 1800));
  };
  const state = () => page.evaluate(() => {
    const el = document.getElementById('screensaver');
    const cs = getComputedStyle(el);
    return {
      shown: cs.display !== 'none',
      z: Number(cs.zIndex),
      app: (document.getElementById('ss-app') || {}).textContent,
      logo: (() => { const n = document.getElementById('ss-logo');
                     return n && n.offsetParent !== null ? n.getAttribute('src') : null; })(),
      clock: (document.getElementById('ss-clock') || {}).textContent,
      minutes: (typeof _ss !== 'undefined') ? _ss.minutes : null,
      ariaHidden: el.getAttribute('aria-hidden'),
    };
  });

  console.log('[off] 0 minutes means never — the shipped default');
  await put({ screensaver_minutes: '0', trade_name: 'RNS SPICES' });
  await open();
  let s = await state();
  check('it is off by default', s.minutes === 0 && !s.shown, s);
  // Force the timer to fire: even then, 0 must mean never.
  await page.evaluate(() => ssShow());
  s = await state();
  check('…and cannot be shown while off', !s.shown, s);

  console.log('\n[show] it appears on idle, named');
  // A tiny window; ssArm floors the wait at 5s so this is the real path.
  await put({ screensaver_minutes: '0.05', app_name: 'GreenGavel' });
  await open();
  await new Promise(r => setTimeout(r, 6500));
  s = await state();
  check('the screensaver is up', s.shown, s);
  check('it carries the application name', /GreenGavel/.test(s.app || ''), s.app);
  // The APP's mark, never the customer's upload: the company logo belongs on
  // every working screen, not on the one screen that is about the software.
  check('it shows the application logo', /logo-app\.svg/.test(s.logo || ''), s.logo);
  check('…not the customer company logo', !/logo-ispl/.test(s.logo || ''), s.logo);
  check('…and an IST clock', /\d{2}:\d{2}\s*IST/.test(s.clock || ''), s.clock);
  check('it is scenery, not content (aria-hidden)', s.ariaHidden === 'true', s.ariaHidden);
  // Below the 9999 band every banner and warning in this app uses: scenery
  // must never cover something the operator has to read.
  check('it sits below the app’s warning band', s.z === 9000, s.z);

  console.log('\n[clear] a key or a click clears it');
  await page.keyboard.press('a');
  await new Promise(r => setTimeout(r, 400));
  s = await state();
  check('a keypress clears it', !s.shown, s);
  await new Promise(r => setTimeout(r, 6500));
  s = await state();
  check('…and it comes back after the next quiet spell', s.shown, s);
  await page.mouse.click(700, 430);
  await new Promise(r => setTimeout(r, 400));
  s = await state();
  check('a click clears it too', !s.shown, s);

  console.log('\n[session] covering the screen does not sign anybody out');
  const stillIn = await page.evaluate(async () => {
    _ss.minutes = 1; ssShow();
    const r = await fetch('/api/me', { headers: { Authorization: 'Bearer ' + localStorage.getItem('t') } });
    return { ok: r.ok, on: _ss.on };
  });
  check('the session is untouched while the panel is up', stillIn.ok && stillIn.on, stillIn);

  console.log('\n[lockup] name, credit and numbers share one centre axis');
  // The complaint this covers: a long product name overflowed a 640px box
  // and wrapped mid-name, so the block read as crooked. Assert the real
  // geometry — every line centred on the same axis, and the name INSIDE its
  // container — rather than that the CSS string says "center".
  await put({ screensaver_minutes: '1', app_name: 'SpiceFlow eAuction Suite',
              app_credit: 'Powered by KJ & Co', app_contact: '9790745462 / 8248922434' });
  await open();
  const lock = await page.evaluate(() => {
    _ss.minutes = 1; ssShow();
    const box = (id) => {
      const n = document.getElementById(id);
      if (!n || n.offsetParent === null) return null;
      const r = n.getBoundingClientRect();
      return { mid: Math.round(r.left + r.width / 2), w: Math.round(r.width),
               h: Math.round(r.height), text: n.textContent.trim() };
    };
    const col = document.getElementById('ss-app').parentElement.getBoundingClientRect();
    return {
      app: box('ss-app'),
      credit: box('ss-credit'), contact: box('ss-contact'), clock: box('ss-clock'),
      colMid: Math.round(col.left + col.width / 2), colW: Math.round(col.width),
      lineH: parseFloat(getComputedStyle(document.getElementById('ss-app')).fontSize),
    };
  });
  check('the credit line is there', /Powered by KJ & Co/.test((lock.credit || {}).text || ''),
        lock.credit);
  check('…and the support numbers', /9790745462 \/ 8248922434/.test((lock.contact || {}).text || ''),
        lock.contact);
  check('the full name renders unabbreviated',
        /SpiceFlow eAuction Suite/.test((lock.app || {}).text || ''), lock.app);
  // One axis: every line's midpoint within a pixel of the column's.
  const axes = ['app', 'credit', 'contact', 'clock']
    .map(k => ({ k, mid: (lock[k] || {}).mid })).filter(x => x.mid != null);
  const offAxis = axes.filter(x => Math.abs(x.mid - lock.colMid) > 1);
  check('every line sits on the same centre axis', offAxis.length === 0,
        { colMid: lock.colMid, axes, offAxis });
  check('the name fits its column — no overflow', (lock.app || {}).w <= lock.colW,
        { name: (lock.app || {}).w, col: lock.colW });
  // Two lines at this size would be ~2.3x the font size tall. One line is the
  // point: a mid-name wrap is what made it look broken.
  check('…and does not wrap mid-name', (lock.app || {}).h < lock.lineH * 1.8,
        { h: (lock.app || {}).h, fontSize: lock.lineH });

  console.log('\n[unbadged] a blank credit leaves no empty rule');
  await put({ app_credit: '', app_contact: '' });
  await open();
  const bare = await page.evaluate(() => {
    _ss.minutes = 1; ssShow();
    const by = document.getElementById('ss-by');
    return { byShown: !!by && by.offsetParent !== null,
             app: (document.getElementById('ss-app') || {}).textContent };
  });
  check('the whole byline block is gone, hairline and all', bare.byShown === false, bare);
  check('…and the name still shows', /SpiceFlow/.test(bare.app || ''), bare.app);

  console.log('\n[fireworks] it animates, stays bounded, and stops dead');
  await put({ screensaver_minutes: '1', screensaver_fireworks: 'true',
              app_credit: 'Powered by KJ & Co', app_contact: '9790745462 / 8248922434' });
  await open();
  const fx = await page.evaluate(async () => {
    _ss.minutes = 1; ssShow();
    const cv = document.getElementById('ss-fx');
    const ctx = cv.getContext('2d');
    // Sample the canvas rather than trust the state object: this asserts
    // pixels were actually painted, not merely that a loop is running.
    const lit = () => {
      const d = ctx.getImageData(0, 0, cv.width, cv.height).data;
      let n = 0, total = 0;
      for (let i = 3; i < d.length; i += 4 * 13) { total++; if (d[i] > 24) n++; }
      return total ? n / total : 0;      // fraction of the panel carrying light
    };
    // Sample repeatedly and keep the best frame: bursts come and go, so one
    // snapshot can legitimately land in a lull. The figure that matters is
    // the PEAK — whether a burst, when it happens, is actually visible.
    let painted = 0;
    for (let i = 0; i < 24; i++) {
      await new Promise(r => setTimeout(r, 140));
      painted = Math.max(painted, lit());
    }
    const peak = { rockets: _ssFx.rockets.length, sparks: _ssFx.sparks.length };
    // Let it run on and confirm the particle count does not grow with uptime.
    await new Promise(r => setTimeout(r, 2600));
    const later = { rockets: _ssFx.rockets.length, sparks: _ssFx.sparks.length, raf: _ssFx.raf };
    return { painted, peak, later,
             behind: getComputedStyle(cv).pointerEvents,
             canvasW: cv.width };
  });
  // Measured: a real burst lights 1.2–1.6% of the panel at peak. 0.4% sits
  // clear of that and of a near-invisible few specks, and being a fraction
  // it holds at any viewport or DPR.
  check('a burst is actually visible, not a few specks', fx.painted > 0.004,
        (fx.painted * 100).toFixed(2) + '% of panel lit at peak');
  check('rockets stay under the cap', fx.peak.rockets <= 3 && fx.later.rockets <= 3, fx);
  check('…and sparks do not grow without bound', fx.later.sparks < 900, fx.later);
  check('it cannot swallow the dismissing click', fx.behind === 'none', fx.behind);
  check('the canvas is backed by real pixels', fx.canvasW > 0, fx.canvasW);

  // The point of the whole exercise: an unattended machine must not keep a
  // render loop alive behind a dismissed screensaver.
  const stopped = await page.evaluate(async () => {
    ssHide();
    await new Promise(r => setTimeout(r, 300));
    const cv = document.getElementById('ss-fx');
    const d = cv.getContext('2d').getImageData(0, 0, cv.width, cv.height).data;
    let n = 0; for (let i = 3; i < d.length; i += 4 * 13) if (d[i] > 12) n++;
    return { raf: _ssFx.raf, sparks: _ssFx.sparks.length, rockets: _ssFx.rockets.length,
             flashes: _ssFx.flashes.length, lit: n };
  });
  check('hiding cancels the animation frame', stopped.raf === 0, stopped);
  check('…drops every particle', stopped.sparks === 0 && stopped.rockets === 0
        && stopped.flashes === 0, stopped);
  check('…and wipes the canvas', stopped.lit === 0, stopped);

  console.log('\n[fireworks-off] the setting turns it off completely');
  await put({ screensaver_fireworks: 'false' });
  await open();
  const off = await page.evaluate(async () => {
    _ss.minutes = 1; ssShow();
    await new Promise(r => setTimeout(r, 1500));
    const cv = document.getElementById('ss-fx');
    const d = cv.getContext('2d').getImageData(0, 0, cv.width, cv.height).data;
    let n = 0; for (let i = 3; i < d.length; i += 4 * 13) if (d[i] > 12) n++;
    return { raf: _ssFx.raf, lit: n, shown: _ss.on,
             app: (document.getElementById('ss-app') || {}).textContent };
  });
  check('no animation frame is scheduled', off.raf === 0, off);
  check('…nothing is painted', off.lit === 0, off);
  check('…but the screensaver itself still works', off.shown === true && /SpiceFlow/.test(off.app || ''), off);

  console.log('\n[login] the product names itself on the sign-in page too');
  // Two properties:
  //   • it is PRE-AUTH, so it rides /api/branding beside the trade name and
  //     the logo rather than company-settings, which needs a token;
  //   • it does NOT displace the customer's own name — that keeps the line
  //     under the chip, the sign-in card and the top bar. This chip is the
  //     one place the software says what IT is.
  await put({ app_name: 'GreenGavel', trade_name: 'RNS SPICES' });
  const pub = await (await fetch(B + '/api/branding')).json();   // no Authorization header
  check('the public branding payload carries the app name', pub.appName === 'GreenGavel',
        JSON.stringify({ appName: pub.appName, tradeName: pub.tradeName }));
  check('…alongside the customer company, not instead of it', pub.tradeName === 'RNS SPICES',
        pub.tradeName);

  await page.evaluate(() => { try { localStorage.removeItem('t'); } catch (_) {} });
  await page.goto(B + '/', { waitUntil: 'networkidle2' });
  await new Promise(r => setTimeout(r, 1200));
  const login = await page.evaluate(() => {
    const app = document.getElementById('login-hero-app');
    return {
      loginShown: getComputedStyle(document.getElementById('login')).display !== 'none',
      text: app ? app.textContent.trim() : '(absent)',
      // It is the PRODUCT CHIP at the top of the hero, not a line of small
      // print in the footer, which is where it would be missed.
      inChip: !!(app && app.classList.contains('login-hero-eyebrow')),
      visible: !!(app && getComputedStyle(app).display !== 'none'),
      footText: (document.querySelector('#login .login-hero-foot') || {}).textContent || '',
      subText: (document.getElementById('login-hero-sub') || {}).textContent || '',
    };
  });
  check('the login page is what is on screen', login.loginShown === true, JSON.stringify(login));
  check('the app name prints on the login page', login.text === 'GreenGavel', login.text);
  check('…in the product chip at the top, where it cannot be missed',
        login.inChip === true && login.visible === true, JSON.stringify(login));
  check('…and not buried in the footer small print',
        !/GreenGavel/.test(login.footText), login.footText);
  check('the copyright line is still there', /Powered by/.test(login.footText), login.footText);
  check('the customer company is still named, right under the chip',
        /RNS SPICES/.test(login.subText), login.subText);

  // An install that has not named itself keeps the built-in strapline — an
  // empty chip would be worse than the generic one.
  await put({ app_name: '' });
  await page.goto(B + '/', { waitUntil: 'networkidle2' });
  await new Promise(r => setTimeout(r, 1200));
  const blank = await page.evaluate(() => {
    const app = document.getElementById('login-hero-app');
    return { text: app.textContent.trim(),
             chipShown: getComputedStyle(app).display !== 'none' };
  });
  check('an unnamed install falls back to the built-in strapline',
        blank.text === 'Spice e-Trade', blank.text);
  check('…rather than showing an empty chip', blank.chipShown === true, JSON.stringify(blank));
  await put({ app_name: 'SpiceFlow eAuction Suite' });

  if (process.env.SHOT_DIR) {
    await page.evaluate(t => localStorage.setItem('t', t), T);
    await page.goto(B + '/', { waitUntil: 'networkidle2' });
    await new Promise(r => setTimeout(r, 1500));
    await page.evaluate(() => { _ss.minutes = 1; ssShow(); });
    await new Promise(r => setTimeout(r, 700));
    await page.screenshot({ path: process.env.SHOT_DIR + '/screensaver.png' });
    await page.evaluate(() => { ssHide(); localStorage.removeItem('t'); });
    await page.goto(B + '/', { waitUntil: 'networkidle2' });
    await new Promise(r => setTimeout(r, 1200));
    await page.screenshot({ path: process.env.SHOT_DIR + '/login-appname.png' });
  }
  console.log(`\n${pass} passed, ${fail} failed\n`);
  cleanup();
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error('ERR', e); cleanup(); process.exit(1); });
