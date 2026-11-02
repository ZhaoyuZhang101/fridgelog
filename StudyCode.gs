// FridgeLog study collector: paste into Extensions > Apps Script of your "FridgeLog data" sheet.
// Each message from the web app becomes one row in the tab named by "sheet"
// (participants, responses, items). A row with the same key is updated, not duplicated.

function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (body.action === 'claim') return json_(claim_(body));
    writeRow_(body);
    return ContentService.createTextOutput('ok');
  } finally {
    lock.releaseLock();
  }
}

// Gives a participant the next free number (P1, P2, ...). The same uid always gets the same number,
// so reopening the link does not create a new participant. Delete a row in "participants" to free its number.
function claim_(b) {
  const uid = String(b.uid || '').trim().toLowerCase();
  const name = String(b.name || '').trim();
  if (!/^u\d{7}$/.test(uid)) return { ok: false, error: 'bad_uid' };
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('participants');
  let max = 0;
  if (sh && sh.getLastRow() > 1) {
    const vals = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
    const pc = vals[0].indexOf('participant'), uc = vals[0].indexOf('uid');
    for (let i = 1; i < vals.length; i++) {
      const num = parseInt(String(vals[i][pc]).replace(/\D/g, ''), 10) || 0;
      if (String(vals[i][uc]).toLowerCase() === uid) return { ok: true, participant: num, existing: true };
      max = Math.max(max, num);
    }
  }
  const n = max + 1;
  writeRow_({ sheet: 'participants', key: uid, row: { participant: 'P' + n, uid: uid, name: name, addedAt: new Date().toISOString() } });
  return { ok: true, participant: n };
}

function writeRow_(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(body.sheet) || ss.insertSheet(body.sheet);
  const row = Object.assign({ _key: body.key, _receivedAt: new Date().toISOString() }, body.row);

  let headers = sh.getLastColumn() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] : [];
  const extra = Object.keys(row).filter(function (k) { return headers.indexOf(k) === -1; });
  if (extra.length) {
    headers = headers.concat(extra);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  const values = headers.map(function (h) { return row[h] === undefined || row[h] === null ? '' : String(row[h]); });

  const last = sh.getLastRow();
  let target = last + 1;
  if (last > 1) {
    const keys = sh.getRange(2, 1, last - 1, 1).getValues();
    for (let i = 0; i < keys.length; i++) {
      if (keys[i][0] === body.key) { target = i + 2; break; }
    }
  }
  sh.getRange(target, 1, 1, headers.length).setNumberFormat('@').setValues([values]);
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// Open the web app URL in a browser to check it is running.
function doGet() {
  return ContentService.createTextOutput('FridgeLog collector is running.');
}
