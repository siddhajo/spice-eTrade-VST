/**
 * client-profile.js — runtime per-client config loader.
 *
 * Reads process.env.CLIENT once at boot, loads clients/<id>/profile.json,
 * and caches the result. When CLIENT is unset OR the named client doesn't
 * exist, falls back to a safe online-mode default — so a plain `npm start`
 * keeps the existing dev/prod boot path with full security on.
 *
 * The shape of a profile is documented in clients/README.md. The two
 * fields the rest of the codebase reads are:
 *   - mode: 'desktop' | 'online'  (security toggle)
 *   - defaults: object             (cfg seed values; not yet wired)
 */
const fs = require('fs');
const path = require('path');

const SAFE_DEFAULT = Object.freeze({
  id: 'default-online',
  mode: 'online',
  displayName: 'Spice Admin',
  defaults: {},
});

let cached = null;

function loadClientProfile() {
  if (cached) return cached;
  const id = process.env.CLIENT;
  if (!id) {
    cached = SAFE_DEFAULT;
    console.log('[client-profile] CLIENT unset — using safe default (mode: online)');
    return cached;
  }
  const p = path.join(__dirname, 'clients', id, 'profile.json');
  try {
    const parsed = JSON.parse(fs.readFileSync(p, 'utf8'));
    cached = { ...SAFE_DEFAULT, ...parsed };
  } catch (e) {
    console.warn(`[client-profile] could not load ${p}: ${e.message} — falling back to online mode`);
    cached = SAFE_DEFAULT;
  }
  console.log(`[client-profile] active: ${cached.id} (mode: ${cached.mode})`);
  return cached;
}

function isOnline()  { return loadClientProfile().mode !== 'desktop'; }
function isDesktop() { return loadClientProfile().mode === 'desktop'; }

/**
 * Load and register optional client-specific overrides.
 *
 * Looks for `clients/<active-client>/overrides/index.js`. If present and it
 * exports a `register(app, ctx)` function, that function is invoked here
 * with the live Express app and a context bag containing the common
 * middleware + helpers a client extension is most likely to need.
 *
 * Called from server.js once at boot, AFTER all built-in routes have been
 * registered and BEFORE `app.listen`. Overrides are therefore *additive* —
 * they can mount new endpoints, schedule jobs, attach middleware, etc.,
 * but cannot intercept an existing route (Express picks the first matching
 * handler in registration order). To override a built-in route, the
 * extension must register a path-specific middleware that short-circuits
 * with `res.send(...)` — Express still walks the middleware chain for
 * those even though it ignores later same-path route handlers.
 *
 * Errors thrown during registration propagate to the caller — a broken
 * override should fail loud at boot, not produce a half-working server.
 */
function loadClientOverrides(app, ctx) {
  const profile = loadClientProfile();
  // Safe-default profiles aren't tied to a clients/ directory — skip silently.
  if (!profile || !profile.id || profile.id === 'default-online') {
    return { loaded: false, reason: 'no client id (using safe default)' };
  }
  const overridesPath = path.join(__dirname, 'clients', profile.id, 'overrides', 'index.js');
  if (!fs.existsSync(overridesPath)) {
    return { loaded: false, reason: 'no overrides/index.js' };
  }
  const mod = require(overridesPath);
  if (typeof mod.register !== 'function') {
    console.warn(`[client-overrides] ${profile.id}/overrides/index.js missing register() export — skipping`);
    return { loaded: false, reason: 'no register() export' };
  }
  mod.register(app, ctx);
  console.log(`[client-overrides] registered: ${profile.id}`);
  return { loaded: true };
}

module.exports = { loadClientProfile, isOnline, isDesktop, loadClientOverrides };
