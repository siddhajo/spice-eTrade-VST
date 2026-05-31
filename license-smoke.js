const lic = require('./license');

// 1) Sign / verify roundtrip
const token = lic.signToken({
  install_id: 'test-uuid-1234',
  expires_at: '2026-12-31',
  issued_at: new Date().toISOString(),
});
console.log('token len:', token.length, 'prefix:', token.slice(0, 40) + '...');
const v = lic.verifyToken(token);
console.log('verify ok =>', v.ok, '/ install:', v.payload && v.payload.install_id);

// 2) Tamper detection
const tampered = token.slice(0, -1) + (token.slice(-1) === 'a' ? 'b' : 'a');
const v2 = lic.verifyToken(tampered);
console.log('tampered ok =>', v2.ok, '/ err:', v2.error);

// 3) Wrong-format strings
console.log('garbage:', JSON.stringify(lic.verifyToken('garbage.signature')));
console.log('empty:',   JSON.stringify(lic.verifyToken('')));
console.log('no dot:',  JSON.stringify(lic.verifyToken('abc')));

// 4) applyToken via in-memory mock DB
let row = null;
const mockDb = {
  get: () => row,
  run: (sql, params) => {
    if (/INSERT INTO license_state/.test(sql)) {
      row = { install_id: params[0], first_seen_at: params[1], expires_at: params[2], active_token: null };
    } else if (/UPDATE license_state/.test(sql)) {
      row.expires_at = params[0]; row.active_token = params[1];
    }
  },
};
const status0 = lic.getStatus(mockDb);
console.log('bootstrap install_id:', status0.install_id.slice(0, 8) + '..', 'expires:', status0.expires_at, 'days_remaining:', status0.days_remaining);

// Mint a future token for THIS install and apply
const realToken = lic.signToken({ install_id: status0.install_id, expires_at: '2030-01-01', issued_at: new Date().toISOString() });
const ar = lic.applyToken(mockDb, realToken);
console.log('applyToken ok =>', ar.ok, '/ new expires:', ar.status && ar.status.expires_at, '/ days:', ar.status && ar.status.days_remaining);

// 5) Reject wrong-install token
const wrongToken = lic.signToken({ install_id: 'someone-else', expires_at: '2030-01-01', issued_at: new Date().toISOString() });
const ar2 = lic.applyToken(mockDb, wrongToken);
console.log('wrong-install ok? =>', ar2.ok, '/ err:', ar2.error);

// 6) Reject past-expiry token
const oldToken = lic.signToken({ install_id: status0.install_id, expires_at: '2000-01-01', issued_at: new Date().toISOString() });
const ar3 = lic.applyToken(mockDb, oldToken);
console.log('past-expiry ok? =>', ar3.ok, '/ err:', ar3.error);

// 7) Reject shorter-than-current token
const shortToken = lic.signToken({ install_id: status0.install_id, expires_at: '2027-01-01', issued_at: new Date().toISOString() });
const ar4 = lic.applyToken(mockDb, shortToken);
console.log('shorter ok? =>', ar4.ok, '/ err:', ar4.error);
