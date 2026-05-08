# sql.js → better-sqlite3 Migration Guide

## What changed

**Two files changed, zero behavioural changes for your other code.**

1. `db.js` — rewritten to use `better-sqlite3` with WAL mode. The wrapper
   API (`run`, `get`, `all`, `exec`, `prepare`, `transaction`) is preserved
   exactly, so your `server.js`, `calculations.js`, `company-config.js`,
   `exports.js`, `invoice-pdf.js`, `dbf-exports.js`, and `exports-pdf.js`
   don't need any changes.

2. `package.json` — replaced `"sql.js"` with `"better-sqlite3"`.

## Why it matters

- **No more lost writes under concurrent use.** sql.js writes the full DB
  to disk on every operation — two branches saving at the same moment
  could race and overwrite each other. better-sqlite3 + WAL mode handles
  this safely.
- **~10x faster** on inserts and queries (native C vs WASM).
- **Crash-safe** — the file on disk is always consistent, even if the
  process is killed mid-write.

## Steps to apply on your Mac

### 1. Back up your current database first (IMPORTANT)

```bash
cd /path/to/spice-config-complete
cp data/config.db data/config.db.backup-$(date +%Y%m%d)
```

### 2. Replace the two files

Extract the provided zip over your existing project. Only `db.js` and
`package.json` are new — all other files are byte-for-byte the same as
what you uploaded.

### 3. Reinstall dependencies

```bash
rm -rf node_modules package-lock.json
npm install
```

This will compile `better-sqlite3` natively. It needs Python and a C++
compiler, which macOS has by default — but if you see any build errors,
run once:

```bash
xcode-select --install
```

### 4. Start the server

```bash
node server.js
```

You should see:
```
Database ready at .../data/config.db (better-sqlite3, WAL mode)
Company settings ready (109 defaults)
Spice Config running at http://localhost:3001
```

**Your existing data is preserved.** SQLite's file format is the same
across sql.js and better-sqlite3 — better-sqlite3 just opens the existing
file and immediately starts using it natively.

### 5. Verify

- Login with your existing admin account — should work unchanged.
- Click through Sellers, Buyers, Auctions, Lots — all data should be
  there.
- Try creating a new trader and refresh. If the record persists, you're
  done.

## If something goes wrong

Roll back:
```bash
cp data/config.db.backup-YYYYMMDD data/config.db
# restore old db.js and package.json, then:
rm -rf node_modules package-lock.json
npm install
```

And send me the error. The most common issue is missing build tools for
`better-sqlite3`'s native compile — fixed by `xcode-select --install` on
Mac.

## New files created alongside config.db

After first run you'll see these extra files in `data/`:
- `config.db-wal` — the write-ahead log (where new writes go before being
  committed to the main file). Normal.
- `config.db-shm` — shared memory index for WAL. Normal.

Both get cleaned up automatically when the process exits cleanly.
**Important:** when backing up, copy the main `config.db` file — WAL
checkpoints on startup, so the main file will have everything after the
next clean restart.

## DigitalOcean deployment note

When you deploy to your Bangalore droplet, Ubuntu needs the same build
tools. You already installed them in the earlier deployment steps
(`sudo apt install build-essential python3`), so `npm install` will just
work there too.
