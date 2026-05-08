/**
 * spice-online — client-specific extensions
 *
 * Loaded by client-profile.js#loadClientOverrides() once at boot, AFTER
 * all built-in routes have registered.
 *
 * The `ctx` object exposes the common bits a route handler typically
 * needs:
 *   - getDb()                — the wrapped DB handle (sql.js / better-sqlite3)
 *   - getSettingsFlat(db)    — flat company-config map (rates, addresses, flags)
 *   - requireAuth            — middleware: validates Bearer token
 *   - requirePermission(cap) — middleware: enforces a role capability
 *   - requireAnyPermission(...) — middleware: any-of capability gate
 *   - IS_ONLINE              — boolean (true here since this is the online client)
 *   - CLIENT_PROFILE         — the parsed profile.json
 *
 * Add new routes / middleware / scheduled jobs below. Keep this file
 * focused on cross-cutting concerns specific to this deployment — for
 * code that should run on every client, edit the shared modules instead.
 */
function register(app, ctx) {
  const { requireAuth, CLIENT_PROFILE, IS_ONLINE } = ctx;

  // GET /api/client-info
  // Auth-gated metadata about the active client. Lets the frontend
  // detect which deployment it's talking to without leaking sensitive
  // config — only the public-safe fields are returned.
  app.get('/api/client-info', requireAuth, (req, res) => {
    res.json({
      id:          CLIENT_PROFILE.id,
      mode:        CLIENT_PROFILE.mode,
      displayName: CLIENT_PROFILE.displayName,
      online:      IS_ONLINE,
    });
  });
}

module.exports = { register };
