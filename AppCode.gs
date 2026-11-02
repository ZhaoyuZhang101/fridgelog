// FridgeLog app backend: paste into Extensions > Apps Script of a NEW Google Sheet ("FridgeLog app").
// It stores accounts and each person's fridge list in the users, sessions and fridge tabs.
// Study data goes to your other sheet and its own script; nothing from the study is stored here.

const APP_SHEETS = {
  users: ['username', 'salt', 'hash', 'createdAt'],
  sessions: ['token', 'username', 'createdAt'],
  fridge: ['id', 'username', 'name', 'category', 'date', 'addedBy', 'addedAt', 'status', 'closedAt'],
};

function doPost(e) {
  return json_(handleApp_(JSON.parse(e.postData.contents)));
}

// Open the web app URL in a browser to check it is running.
function doGet() {
  return ContentService.createTextOutput('FridgeLog app backend is running.');
}

/* ---------------- everyday app ---------------- */

function handleApp_(b) {
  try {
    switch (b.action) {
      case 'register': return withLock_(function () { return register_(b); });
      case 'login': return login_(b);
      case 'logout': return withLock_(function () { return logout_(b); });
      case 'list': return list_(b);
      case 'add': return withLock_(function () { return add_(b); });
      case 'close': return withLock_(function () { return close_(b); });
    }
    return { ok: false, error: 'unknown_action' };
  } catch (err) {
    return { ok: false, error: 'server', message: String(err) };
  }
}

function register_(b) {
  const username = String(b.username || '').trim().toLowerCase();
  const password = String(b.password || '');
  if (!/^[a-z0-9_.-]{3,30}$/.test(username)) return { ok: false, error: 'bad_username' };
  if (password.length < 6) return { ok: false, error: 'weak_password' };
  if (rows_('users').some(function (u) { return u.username === username; })) return { ok: false, error: 'taken' };
  const salt = Utilities.getUuid();
  append_('users', { username: username, salt: salt, hash: hash_(salt, password), createdAt: new Date().toISOString() });
  return { ok: true, username: username, token: newSession_(username) };
}

function login_(b) {
  const username = String(b.username || '').trim().toLowerCase();
  const user = rows_('users').filter(function (u) { return u.username === username; })[0];
  if (!user || hash_(user.salt, String(b.password || '')) !== user.hash) return { ok: false, error: 'bad_login' };
  return { ok: true, username: username, token: withLock_(function () { return newSession_(username); }) };
}

function logout_(b) {
  const sh = sheet_('sessions');
  const s = rows_('sessions').filter(function (r) { return r.token === b.token; })[0];
  if (s) sh.deleteRow(s._row);
  return { ok: true };
}

function list_(b) {
  const username = userFor_(b.token);
  if (!username) return { ok: false, error: 'auth' };
  const items = rows_('fridge')
    .filter(function (r) { return r.username === username && r.status === 'in'; })
    .map(function (r) { return { id: r.id, name: r.name, category: r.category, date: r.date, addedBy: r.addedBy, addedAt: r.addedAt }; });
  return { ok: true, items: items };
}

function add_(b) {
  const username = userFor_(b.token);
  if (!username) return { ok: false, error: 'auth' };
  const it = b.item || {};
  const name = String(it.name || '').trim().slice(0, 60);
  const date = String(it.date || '');
  if (!name || !/^\d{4}-\d{2}-\d{2}$/.test(date)) return { ok: false, error: 'bad_item' };
  const item = {
    id: Utilities.getUuid(), username: username, name: name, category: String(it.category || 'Other').slice(0, 30),
    date: date, addedBy: String(it.addedBy || ''), addedAt: new Date().toISOString(), status: 'in', closedAt: '',
  };
  append_('fridge', item);
  delete item.username; delete item.status; delete item.closedAt;
  return { ok: true, item: item };
}

function close_(b) {
  const username = userFor_(b.token);
  if (!username) return { ok: false, error: 'auth' };
  if (['eaten', 'binned'].indexOf(b.outcome) === -1) return { ok: false, error: 'bad_outcome' };
  const r = rows_('fridge').filter(function (x) { return x.id === b.id && x.username === username; })[0];
  if (!r) return { ok: false, error: 'not_found' };
  const col = APP_SHEETS.fridge.indexOf('status') + 1;
  sheet_('fridge').getRange(r._row, col, 1, 2).setNumberFormat('@').setValues([[b.outcome, new Date().toISOString()]]);
  return { ok: true };
}

/* ---------------- helpers ---------------- */

function userFor_(token) {
  if (!token) return null;
  const s = rows_('sessions').filter(function (r) { return r.token === token; })[0];
  return s ? s.username : null;
}

function newSession_(username) {
  const token = Utilities.getUuid() + Utilities.getUuid();
  append_('sessions', { token: token, username: username, createdAt: new Date().toISOString() });
  return token;
}

// Salted SHA-256, repeated so a leaked sheet is slow to brute-force. Fine for a student prototype.
function hash_(salt, password) {
  let h = salt + ':' + password;
  for (let i = 0; i < 300; i++) {
    h = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h, Utilities.Charset.UTF_8));
  }
  return h;
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function sheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, APP_SHEETS[name].length).setValues([APP_SHEETS[name]]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function rows_(name) {
  const sh = sheet_(name);
  const n = sh.getLastRow();
  const h = APP_SHEETS[name];
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, h.length).getDisplayValues().map(function (r, i) {
    const o = { _row: i + 2 };
    h.forEach(function (k, j) { o[k] = r[j]; });
    return o;
  });
}

function append_(name, obj) {
  const sh = sheet_(name);
  const values = APP_SHEETS[name].map(function (k) { return obj[k] === undefined ? '' : String(obj[k]); });
  sh.getRange(sh.getLastRow() + 1, 1, 1, values.length).setNumberFormat('@').setValues([values]);
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
