#!/usr/bin/env node
/**
 * apply-client.js — pre-build hook
 *
 * Reads the CLIENT env var, resolves clients/<id>/, and:
 *   1. Copies any clients/<id>/assets/* into public/  (logos, icons)
 *   2. Writes electron-builder.json that deep-merges
 *      package.json#build with the client's profile.build overrides.
 *
 * electron-builder auto-picks electron-builder.json when present and
 * uses it INSTEAD of package.json#build, so this gives us per-client
 * appId / productName / icon paths without mutating package.json.
 *
 * Usage:
 *   CLIENT=spice-etrade-desktop node scripts/apply-client.js
 *   (called by the build:* npm scripts before electron-builder runs)
 */
const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
// Default to the desktop client when unset, since this script only runs
// during desktop-build npm scripts. The runtime loader (client-profile.js)
// uses a different default — online mode — so a bare `npm start` keeps
// the existing dev workflow.
const CLIENT_ID = process.env.CLIENT || 'spice-etrade-desktop';
const clientDir = path.join(ROOT, 'clients', CLIENT_ID);

if (!fs.existsSync(clientDir)) {
  console.error(`[apply-client] FATAL: clients/${CLIENT_ID} does not exist`);
  process.exit(1);
}

let profile;
try {
  profile = JSON.parse(fs.readFileSync(path.join(clientDir, 'profile.json'), 'utf8'));
} catch (e) {
  console.error(`[apply-client] FATAL: cannot read profile.json: ${e.message}`);
  process.exit(1);
}

console.log(`[apply-client] activating: ${profile.displayName || CLIENT_ID} (mode: ${profile.mode})`);

// 1. Copy clients/<id>/assets/* → public/  (overwrites; non-fatal if dir missing)
const assetsDir = path.join(clientDir, 'assets');
if (fs.existsSync(assetsDir)) {
  const publicDir = path.join(ROOT, 'public');
  if (!fs.existsSync(publicDir)) fs.mkdirSync(publicDir, { recursive: true });
  for (const fname of fs.readdirSync(assetsDir)) {
    const src = path.join(assetsDir, fname);
    if (!fs.statSync(src).isFile()) continue;
    const dst = path.join(publicDir, fname);
    fs.copyFileSync(src, dst);
    console.log(`[apply-client] copied   ${fname} → public/`);
  }
} else {
  console.log(`[apply-client] (no assets/ dir — skipping asset copy)`);
}

// 2. Write merged electron-builder.json
const pkg = JSON.parse(fs.readFileSync(path.join(ROOT, 'package.json'), 'utf8'));
const baseBuild = pkg.build || {};
const overrides = profile.build || {};
const merged = deepMerge(baseBuild, overrides);
const out = path.join(ROOT, 'electron-builder.json');
fs.writeFileSync(out, JSON.stringify(merged, null, 2) + '\n');
console.log(`[apply-client] wrote     ${path.relative(ROOT, out)}`);

/**
 * Deep-merge: objects merge key-by-key recursively, arrays REPLACE
 * (electron-builder array fields like `files` / `extraResources` are
 * complete lists, not partial overlays).
 */
function deepMerge(a, b) {
  if (Array.isArray(a) && Array.isArray(b)) return b;
  if (a && typeof a === 'object' && !Array.isArray(a) && b && typeof b === 'object' && !Array.isArray(b)) {
    const out = { ...a };
    for (const k of Object.keys(b)) out[k] = deepMerge(a[k], b[k]);
    return out;
  }
  return b !== undefined ? b : a;
}
