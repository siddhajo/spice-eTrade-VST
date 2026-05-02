// distance.js — PIN-to-PIN road-distance estimator for the e-way bill
// <DISTANCE> field in ISP sales vouchers.
//
// Strategy: haversine straight-line distance × road multiplier, cached
// in the pin_distances table. PIN coordinates come from the local
// `pincodes` table (seeded on first DB init; growable via the Settings
// → PIN Distances admin panel).
//
// Why this exists:
//   - Tally ISP sales vouchers ship a <DISTANCE> field that the e-way
//     bill portal cross-checks at upload time. Leaving it blank works
//     but means manual entry on every voucher. Pre-computing gets the
//     number on the XML automatically.
//   - The ideal source is the e-way bill portal's own get-distance API,
//     but that requires GSP API access. This module gives us a working
//     90% solution for free until/unless we wire up a GSP later.
//
// Accuracy: haversine × 1.3 is within ~10% of true road distance for
// most India hill-and-plain routes. Actual e-way bill portal numbers
// can differ. To match the portal exactly, override the per-invoice
// distance via `invoices.distance_km` (write any non-null integer and
// the generator will use it as-is).

const DEFAULT_ROAD_MULTIPLIER = 1.50;
const EARTH_RADIUS_KM = 6371;

/**
 * Great-circle distance between two (lat, lon) points, in kilometres.
 * Plain haversine — no curvature corrections (overkill at the ~100 km
 * scale we operate at).
 */
function haversineKm(lat1, lon1, lat2, lon2) {
  const toRad = (d) => (d * Math.PI) / 180;
  const dLat = toRad(lat2 - lat1);
  const dLon = toRad(lon2 - lon1);
  const a =
    Math.sin(dLat / 2) ** 2 +
    Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) * Math.sin(dLon / 2) ** 2;
  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  return EARTH_RADIUS_KM * c;
}

/** Look up a PIN's coordinates in the `pincodes` table; returns null if missing. */
function lookupPin(db, pin) {
  if (!pin) return null;
  const row = db.get('SELECT pin, lat, lon, place, state FROM pincodes WHERE pin = ?', [String(pin).trim()]);
  return row || null;
}

/**
 * Compute the road-km estimate between two PINs. Returns:
 *   { km, source, error? }
 * - km      : integer kilometres (rounded), or null if can't compute
 * - source  : 'cache' | 'haversine' | 'manual' | null
 * - error   : human-readable reason when km is null (e.g. 'unknown pin: 600043')
 *
 * The (from, to) pair is cached in `pin_distances` after the first
 * computation. The cache is symmetric — we always store the
 * lexicographically-smaller PIN as `from_pin` so A→B and B→A share
 * a single cache row.
 */
function getDistance(db, fromPin, toPin, opts = {}) {
  const multiplier = Number(opts.multiplier || DEFAULT_ROAD_MULTIPLIER);
  const a = String(fromPin || '').trim();
  const b = String(toPin || '').trim();
  if (!a || !b) return { km: null, source: null, error: 'missing pin' };
  if (a === b) return { km: 0, source: 'identity' };

  // Normalise the cache key so A↔B share a row
  const [k1, k2] = a < b ? [a, b] : [b, a];

  // Cache hit?
  const cached = db.get(
    'SELECT road_km, source FROM pin_distances WHERE from_pin = ? AND to_pin = ?',
    [k1, k2]
  );
  if (cached) return { km: cached.road_km, source: cached.source || 'cache' };

  // Miss → compute via lookups + haversine
  const p1 = lookupPin(db, a);
  const p2 = lookupPin(db, b);
  if (!p1) return { km: null, source: null, error: `unknown pin: ${a}` };
  if (!p2) return { km: null, source: null, error: `unknown pin: ${b}` };

  const straight = haversineKm(p1.lat, p1.lon, p2.lat, p2.lon);
  const road = Math.round(straight * multiplier);

  // Cache it. Source 'haversine' tells the admin UI this is an
  // estimate (vs 'manual' or 'eway_portal' which would be authoritative).
  try {
    db.run(
      `INSERT OR REPLACE INTO pin_distances (from_pin, to_pin, road_km, source, updated_at)
       VALUES (?, ?, ?, 'haversine', datetime('now','localtime'))`,
      [k1, k2, road]
    );
  } catch (e) { /* cache write is best-effort */ }

  return { km: road, source: 'haversine' };
}

/**
 * List PINs that appear in the data (buyers + traders + invoices) but
 * are NOT in the `pincodes` table — the admin UI uses this to surface
 * PINs that need lat/lon entered.
 *
 * Optionally scoped to one auction so users can clean up just the PINs
 * relevant to vouchers they're about to export.
 */
function listMissingPins(db, auctionId = null) {
  // Pull DISTINCT PINs from buyers (via invoices) + traders (via lots),
  // optionally scoped to a single auction. Then anti-join against
  // `pincodes` to surface only the unknown ones.
  const buyerSql = `
    SELECT DISTINCT b.pin AS pin, COALESCE(b.pla, '') AS place, COALESCE(b.state, '') AS state, 'buyer' AS kind
    FROM invoices i
    LEFT JOIN buyers b ON b.buyer = i.buyer
    WHERE b.pin IS NOT NULL AND b.pin != ''
    ${auctionId ? 'AND i.auction_id = ?' : ''}
  `;
  const traderSql = `
    SELECT DISTINCT l.ppin AS pin, COALESCE(l.ppla, '') AS place, COALESCE(l.pstate, '') AS state, 'trader' AS kind
    FROM lots l
    WHERE l.ppin IS NOT NULL AND l.ppin != ''
    ${auctionId ? 'AND l.auction_id = ?' : ''}
  `;
  const params = auctionId ? [auctionId, auctionId] : [];
  const sql = `
    WITH used_pins AS (
      ${buyerSql}
      UNION
      ${traderSql}
    )
    SELECT u.pin, u.place, u.state, u.kind
    FROM used_pins u
    LEFT JOIN pincodes p ON p.pin = u.pin
    WHERE p.pin IS NULL
    ORDER BY u.pin
  `;
  try {
    return db.all(sql, params);
  } catch (e) {
    return [];
  }
}

module.exports = {
  haversineKm,
  lookupPin,
  getDistance,
  listMissingPins,
  DEFAULT_ROAD_MULTIPLIER,
};
