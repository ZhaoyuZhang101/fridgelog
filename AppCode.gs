// Fridgey app backend: paste into Extensions > Apps Script of the "Fridgey app" Google Sheet.
// It stores accounts and each person's fridge list in the users, sessions and fridge tabs.
// Built-in accounts: "admin" can manage users and "study" can open the Study tab.
// Both are created automatically with DEFAULT_PASSWORD. Sign in and change the password straight away.
// Another account can get the same powers by typing admin or study in the "role" column of the users tab.
// Phone notifications: in Project Settings > Script properties, add ONESIGNAL_APP_ID and ONESIGNAL_REST_KEY
// (from OneSignal > Settings > Keys & IDs). The daily reminder then also goes to each person's phone.
// Expiry emails: run setupReminders once from the editor (Run button) to send a daily email at 8am Sydney time.
// Study data goes to your other sheet and its own script; nothing from the study is stored here.

// Password an admin reset gives an account. The person is emailed and asked to change it.
const DEFAULT_PASSWORD = '12345678';
// Address people see when the app emails them.
const APP_LINK = 'https://zhaoyuzhang101.github.io/fridgelog/';

const APP_SHEETS = {
  users: ['username', 'salt', 'hash', 'createdAt', 'email', 'role'],
  sessions: ['token', 'username', 'createdAt'],
  fridge: ['id', 'username', 'name', 'category', 'date', 'addedBy', 'addedAt', 'status', 'closedAt'],
};

function doPost(e) {
  return json_(handleApp_(JSON.parse(e.postData.contents)));
}

// Open the web app URL in a browser to check it is running.
function doGet() {
  return ContentService.createTextOutput('Fridgey app backend is running.');
}

/* ---------------- everyday app ---------------- */

const BUILT_IN = ['admin', 'study'];

function handleApp_(b) {
  try {
    if (b.action === 'login' || b.action === 'register') withLock_(ensureBuiltIns_);
    switch (b.action) {
      case 'register': return withLock_(function () { return register_(b); });
      case 'login': return login_(b);
      case 'logout': return withLock_(function () { return logout_(b); });
      case 'list': return list_(b);
      case 'add': return withLock_(function () { return add_(b); });
      case 'close': return withLock_(function () { return close_(b); });
      case 'account': return account_(b);
      case 'setEmail': return withLock_(function () { return setEmail_(b); });
      case 'changePassword': return withLock_(function () { return changePassword_(b); });
      case 'testPush': return testPush_(b);
      case 'adminUsers': return adminUsers_(b);
      case 'adminReset': return withLock_(function () { return adminReset_(b); });
      case 'adminDelete': return withLock_(function () { return adminDelete_(b); });
    }
    return { ok: false, error: 'unknown_action' };
  } catch (err) {
    return { ok: false, error: 'server', message: String(err) };
  }
}

function register_(b) {
  const username = String(b.username || '').trim().toLowerCase();
  const password = String(b.password || '');
  if (!/^[a-z0-9_.@+-]{3,60}$/.test(username)) return { ok: false, error: 'bad_username' };
  if (password.length < 6) return { ok: false, error: 'weak_password' };
  const email = String(b.email || '').trim().toLowerCase();
  if (email && !isEmail_(email)) return { ok: false, error: 'bad_email' };
  if (rows_('users').some(function (u) { return u.username === username; })) return { ok: false, error: 'taken' };
  const salt = Utilities.getUuid();
  append_('users', { username: username, salt: salt, hash: hash_(salt, password), createdAt: new Date().toISOString(), email: email });
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
  if (['eaten', 'binned', 'deleted'].indexOf(b.outcome) === -1) return { ok: false, error: 'bad_outcome' };
  const r = rows_('fridge').filter(function (x) { return x.id === b.id && x.username === username; })[0];
  if (!r) return { ok: false, error: 'not_found' };
  const col = APP_SHEETS.fridge.indexOf('status') + 1;
  sheet_('fridge').getRange(r._row, col, 1, 2).setNumberFormat('@').setValues([[b.outcome, new Date().toISOString()]]);
  return { ok: true };
}

function account_(b) {
  const username = userFor_(b.token);
  if (!username) return { ok: false, error: 'auth' };
  const u = rows_('users').filter(function (r) { return r.username === username; })[0];
  return { ok: true, username: username, email: u ? emailFor_(u) : '', role: u ? roles_(u).join(', ') : '' };
}

function setEmail_(b) {
  const username = userFor_(b.token);
  if (!username) return { ok: false, error: 'auth' };
  const email = String(b.email || '').trim().toLowerCase();
  if (email && !isEmail_(email)) return { ok: false, error: 'bad_email' };
  const u = rows_('users').filter(function (r) { return r.username === username; })[0];
  if (!u) return { ok: false, error: 'auth' };
  // "-" means the person turned reminders off, even if their username is an email address.
  sheet_('users').getRange(u._row, APP_SHEETS.users.indexOf('email') + 1, 1, 1).setNumberFormat('@').setValues([[email || '-']]);
  return { ok: true, email: email };
}

function changePassword_(b) {
  const username = userFor_(b.token);
  if (!username) return { ok: false, error: 'auth' };
  const u = rows_('users').filter(function (r) { return r.username === username; })[0];
  if (!u || hash_(u.salt, String(b.oldPassword || '')) !== u.hash) return { ok: false, error: 'bad_login' };
  const pw = String(b.newPassword || '');
  if (pw.length < 6) return { ok: false, error: 'weak_password' };
  setPassword_(u, pw);
  return { ok: true };
}

/* ---------------- admin ---------------- */

function adminFor_(token) {
  const username = userFor_(token);
  if (!username) return null;
  const u = rows_('users').filter(function (r) { return r.username === username; })[0];
  return u && roles_(u).indexOf('admin') !== -1 ? username : null;
}

function adminUsers_(b) {
  if (!adminFor_(b.token)) return { ok: false, error: 'forbidden' };
  const fridge = rows_('fridge');
  return {
    ok: true,
    users: rows_('users').map(function (u) {
      return {
        username: u.username, email: emailFor_(u), createdAt: u.createdAt,
        role: roles_(u).filter(function (r, i, a) { return a.indexOf(r) === i; }).join(', '),
        items: fridge.filter(function (f) { return f.username === u.username && f.status === 'in'; }).length,
      };
    }),
  };
}

function adminReset_(b) {
  if (!adminFor_(b.token)) return { ok: false, error: 'forbidden' };
  const u = rows_('users').filter(function (r) { return r.username === b.username; })[0];
  if (!u) return { ok: false, error: 'not_found' };
  setPassword_(u, DEFAULT_PASSWORD);
  endSessions_(u.username);
  const emailed = mail_(emailFor_(u), 'Your Fridgey password was reset',
    'Hi ' + u.username + ',\n\nAn administrator reset your Fridgey password. Sign in with:\n\n' +
    'Username: ' + u.username + '\nPassword: ' + DEFAULT_PASSWORD + '\n\n' +
    'Then open your account (the round button at the top) and choose Change password.\n\n' + APP_LINK);
  return { ok: true, emailed: emailed };
}

function adminDelete_(b) {
  const me = adminFor_(b.token);
  if (!me) return { ok: false, error: 'forbidden' };
  if (b.username === me) return { ok: false, error: 'self' };
  if (BUILT_IN.indexOf(b.username) !== -1) return { ok: false, error: 'built_in' };
  const u = rows_('users').filter(function (r) { return r.username === b.username; })[0];
  if (!u) return { ok: false, error: 'not_found' };
  const to = emailFor_(u);
  deleteRows_('fridge', function (r) { return r.username === u.username; });
  endSessions_(u.username);
  sheet_('users').deleteRow(u._row);
  const emailed = mail_(to, 'Your Fridgey account was deleted',
    'Hi ' + u.username + ',\n\nAn administrator deleted your Fridgey account and its fridge list.\n' +
    'You can create a new account at any time:\n\n' + APP_LINK);
  return { ok: true, emailed: emailed };
}

/* ---------------- expiry emails ---------------- */

// Run once from the editor: select setupReminders, press Run, and allow Gmail access.
function setupReminders() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sendReminders') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendReminders').timeBased().everyDays(1).atHour(8).inTimezone('Australia/Sydney').create();
  return 'Daily reminders set for 8am Sydney time.';
}

// Emails each person whose fridge has food that is expired or due to be used today or tomorrow.
function sendReminders() {
  const tz = 'Australia/Sydney';
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const days = function (iso) { return Math.round((Date.parse(iso + 'T00:00:00Z') - Date.parse(today + 'T00:00:00Z')) / 864e5); };
  const byUser = {};
  rows_('fridge').forEach(function (r) {
    if (r.status !== 'in' || !/^\d{4}-\d{2}-\d{2}$/.test(r.date)) return;
    const d = days(r.date);
    if (d <= 1) (byUser[r.username] = byUser[r.username] || []).push({ name: r.name, d: d, date: r.date });
  });
  let sent = 0;
  rows_('users').forEach(function (u) {
    const list = byUser[u.username];
    if (!list) return;
    list.sort(function (a, b) { return a.d - b.d; });
    const label = function (x) { return x.d < 0 ? 'expired ' + (-x.d) + (x.d === -1 ? ' day' : ' days') + ' ago' : x.d === 0 ? 'use today' : 'use tomorrow'; };
    const expired = list.filter(function (x) { return x.d < 0; }).length;
    const subject = expired
      ? 'Fridgey: ' + expired + ' item' + (expired > 1 ? 's' : '') + ' expired, ' + (list.length - expired) + ' to use soon'
      : 'Fridgey: ' + list.length + ' item' + (list.length > 1 ? 's' : '') + ' to use soon';
    const lines = list.map(function (x) { return x.name + ': ' + label(x) + ' (' + x.date + ')'; });
    // Phone notification, for anyone who turned them on in the app.
    try { push_(u.username, subject.replace('Fridgey: ', ''), list.map(function (x) { return x.name + ' (' + label(x) + ')'; }).join(', ')); } catch (e) { /* keep going */ }
    const to = emailFor_(u);
    if (!to || MailApp.getRemainingDailyQuota() < 1) return;
    MailApp.sendEmail({
      to: to,
      subject: subject,
      body: 'In your fridge:\n\n' + lines.join('\n') + '\n\nMark items as Eaten or Thrown out in Fridgey to stop these reminders.',
      htmlBody: '<p>In your fridge:</p><ul>' + lines.map(function (l) { return '<li>' + escapeHtml_(l) + '</li>'; }).join('') +
        '</ul><p>Mark items as <b>Eaten</b> or <b>Thrown out</b> in Fridgey to stop these reminders.</p>',
      name: 'Fridgey',
    });
    sent++;
  });
  return sent;
}

/* ---------------- phone notifications (OneSignal) ---------------- */

function testPush_(b) {
  const username = userFor_(b.token);
  if (!username) return { ok: false, error: 'auth' };
  return { ok: true, result: push_(username, 'Fridgey test', 'Phone notifications are working.') };
}

// Sends one notification to every phone where this username turned notifications on.
// Returns sent, no_device (they have not turned them on), not_set (keys missing) or bad_key.
function push_(username, title, body) {
  const p = PropertiesService.getScriptProperties();
  const app = p.getProperty('ONESIGNAL_APP_ID'), key = p.getProperty('ONESIGNAL_REST_KEY');
  if (!app || !key) return 'not_set';
  const payload = JSON.stringify({
    app_id: app, target_channel: 'push', include_aliases: { external_id: [username] },
    headings: { en: title }, contents: { en: body }, url: APP_LINK,
  });
  const schemes = ['Key ', 'Basic '];   // newer OneSignal keys use "Key", older ones "Basic"
  for (let i = 0; i < schemes.length; i++) {
    const r = UrlFetchApp.fetch('https://api.onesignal.com/notifications?c=push', {
      method: 'post', contentType: 'application/json', payload: payload, muteHttpExceptions: true,
      headers: { Authorization: schemes[i] + key },
    });
    const code = r.getResponseCode();
    if (code === 401 || code === 403) continue;
    let j = {};
    try { j = JSON.parse(r.getContentText() || '{}'); } catch (e) { /* not JSON */ }
    return j.id ? 'sent' : 'no_device';
  }
  return 'bad_key';
}

/* ---------------- helpers ---------------- */

function roles_(u) {
  const r = String(u.role || '').toLowerCase().split(/[\s,]+/).filter(String);
  if (BUILT_IN.indexOf(u.username) !== -1) r.push(u.username);
  return r.filter(function (x, i) { return r.indexOf(x) === i; });
}

// Creates the admin and study accounts the first time anyone signs in or registers.
function ensureBuiltIns_() {
  const have = rows_('users').map(function (u) { return u.username; });
  BUILT_IN.forEach(function (name) {
    if (have.indexOf(name) !== -1) return;
    const salt = Utilities.getUuid();
    append_('users', { username: name, salt: salt, hash: hash_(salt, DEFAULT_PASSWORD), createdAt: new Date().toISOString(), email: '', role: name });
  });
}

function setPassword_(u, password) {
  const salt = Utilities.getUuid();
  sheet_('users').getRange(u._row, APP_SHEETS.users.indexOf('salt') + 1, 1, 2).setNumberFormat('@').setValues([[salt, hash_(salt, password)]]);
}

function endSessions_(username) {
  deleteRows_('sessions', function (r) { return r.username === username; });
}

// Deletes from the bottom up so row numbers stay valid.
function deleteRows_(name, test) {
  const sh = sheet_(name);
  rows_(name).filter(test).map(function (r) { return r._row; }).sort(function (a, b) { return b - a; })
    .forEach(function (row) { sh.deleteRow(row); });
}

function mail_(to, subject, body) {
  if (!to || MailApp.getRemainingDailyQuota() < 1) return false;
  MailApp.sendEmail({ to: to, subject: subject, body: body, name: 'Fridgey' });
  return true;
}

function isEmail_(s) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(s); }

// The email column wins; otherwise a username that is an email address is used.
function emailFor_(u) {
  if (u.email === '-') return '';
  if (u.email) return u.email;
  return isEmail_(u.username) ? u.username : '';
}

function escapeHtml_(s) {
  return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; });
}


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
  } else if (sh.getLastColumn() < APP_SHEETS[name].length) {
    sh.getRange(1, 1, 1, APP_SHEETS[name].length).setValues([APP_SHEETS[name]]);
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
