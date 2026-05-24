// logo-paths.js — resolve where uploaded company logos live so they
// survive across deploys on platforms with ephemeral filesystems
// (Railway, Heroku, Fly, etc.). User uploads go to SPICE_DATA_DIR/logos
// — the same persistent volume the SQLite DB uses — while bundled
// defaults stay in /public. Every PDF/report/route that needs a logo
// reads through resolveLogoPath() so persistent uploads always win
// over stale bundled files of the same name.
//
// On installs without SPICE_DATA_DIR (local dev, packaged Electron),
// uploads continue to go to /public — preserving the historical
// behavior for non-cloud deployments.

const fs = require('fs');
const path = require('path');

const PUBLIC_DIR = path.join(__dirname, 'public');
const PERSISTENT_DIR = process.env.SPICE_DATA_DIR
  ? path.join(process.env.SPICE_DATA_DIR, 'logos')
  : null;

if (PERSISTENT_DIR && !fs.existsSync(PERSISTENT_DIR)) {
  try { fs.mkdirSync(PERSISTENT_DIR, { recursive: true }); }
  catch (_) { /* readers fall back to PUBLIC_DIR if creation fails */ }
}

// Where the upload endpoint should WRITE this logo file. Persistent
// path when SPICE_DATA_DIR is set, else PUBLIC_DIR (legacy/local).
function writePathFor(name) {
  return path.join(PERSISTENT_DIR || PUBLIC_DIR, name);
}

// Where readers should LOOK for this logo file. Returns an existing
// absolute path or null. Persistent location wins so a user upload
// always beats a stale bundled default of the same name.
function resolveLogoPath(name) {
  if (PERSISTENT_DIR) {
    const persistent = path.join(PERSISTENT_DIR, name);
    if (fs.existsSync(persistent)) return persistent;
  }
  const bundled = path.join(PUBLIC_DIR, name);
  if (fs.existsSync(bundled)) return bundled;
  return null;
}

module.exports = { resolveLogoPath, writePathFor, PERSISTENT_DIR, PUBLIC_DIR };
