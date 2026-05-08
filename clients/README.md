# Client Profiles

Per-deployment overrides for the same shared codebase.

## Layout

```
clients/<id>/
  profile.json    # required — mode + branding + electron-builder overrides
  assets/         # optional — files copied into public/ at build time (logos)
  overrides/      # optional — client-specific routes / UI / migrations (not yet wired)
```

## Activating a client

Set `CLIENT=<id>` before running the server or building the desktop app:

```sh
# Run desktop client locally for testing
CLIENT=spice-etrade-desktop npm start

# Build desktop installer (Windows)
CLIENT=spice-etrade-desktop npm run build:win

# Run the online client (or just leave CLIENT unset — same effect)
CLIENT=spice-online npm start
```

`CLIENT` unset → the server boots with a safe default of `online` mode. Existing
dev workflows keep working unchanged.

## profile.json fields

| Field         | Purpose |
|---------------|---------|
| `id`          | Must match the directory name |
| `mode`        | `desktop` (no rate-limit, sessions never expire, no must-change-password gate, binds 127.0.0.1, 50 MB body limit) OR `online` (full hardening: rate-limit, 30-day session cap, must-change-password gate, 0.0.0.0 bind, 2 MB body limit) |
| `displayName` | Human label shown in startup logs |
| `build`       | electron-builder overrides (deep-merged onto package.json#build at build time) |
| `defaults`    | Optional cfg seed values — reserved for later wiring |

## How it works

1. **Runtime:** `client-profile.js` reads `process.env.CLIENT` once at boot
   and caches the loaded profile. `server.js` and `db.js` consult this for
   the security-mode toggle.
2. **Build time:** `npm run build:*` runs `scripts/apply-client.js` first,
   which copies any `clients/<id>/assets/*` into `public/` and writes a
   merged `electron-builder.json` with the client's `build` overrides.
   electron-builder picks that up automatically (it overrides
   `package.json#build` when present).
3. **Overrides:** if `clients/<id>/overrides/index.js` exists, it's loaded
   once at boot AFTER built-in routes. See below.

## Client-specific extensions (`overrides/index.js`)

Drop a file at `clients/<id>/overrides/index.js` exporting a `register`
function. It runs once at boot, on the way to `app.listen`, with the live
Express app and a context bag of common helpers.

```js
// clients/<id>/overrides/index.js
function register(app, ctx) {
  const { requireAuth, getDb, getSettingsFlat,
          requirePermission, requireAnyPermission,
          IS_ONLINE, CLIENT_PROFILE } = ctx;

  app.get('/api/client-info', requireAuth, (req, res) => {
    res.json({ id: CLIENT_PROFILE.id, mode: CLIENT_PROFILE.mode });
  });

  // any combination of new routes, attached middleware, setInterval
  // jobs, etc. Throwing here aborts boot — broken overrides should
  // fail loud rather than ship a half-working server.
}
module.exports = { register };
```

A working example lives at
[`spice-online/overrides/index.js`](spice-online/overrides/index.js).

**Limitations:**
- Overrides are *additive*. Express picks the first matching route, so
  `app.get('/api/login', ...)` registered here will NOT shadow the
  built-in `/api/login` (which already registered earlier in boot).
  To genuinely intercept an existing endpoint, register a path-specific
  middleware that short-circuits with `res.send(...)` — Express still
  walks middleware in order even when later route handlers are skipped.
- For things that should run on EVERY client, edit the shared modules
  instead; don't duplicate the same override under each `clients/<id>/`.

## Mode comparison

| Behavior                     | `desktop` | `online` |
|------------------------------|-----------|----------|
| bcrypt password hashing      | yes       | yes      |
| `must_change_password` gate  | off       | on       |
| Login rate-limit             | off       | on (10 / 15 min / IP) |
| Session `expires_at`         | NULL      | 30-day cap |
| `app.listen` bind address    | 127.0.0.1 | 0.0.0.0 |
| `express.json` body limit    | 50 MB     | 2 MB    |

The toggle lives in **four spots only**: body-limit middleware, login route
(rate-limit + session insert), `requireAuth` (expiry + must-change gate),
and `app.listen` bind. The DB schema is identical regardless of mode.
