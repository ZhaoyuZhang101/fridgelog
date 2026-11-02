// FridgeLog study collector: this is the script already in your "FridgeLog data" sheet. No change needed.
// Each message from the web app becomes one row in the tab named by "sheet"
// (participants, responses, items). A row with the same key is updated, not duplicated.

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const body = JSON.parse(e.postData.contents);
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
    return ContentService.createTextOutput('ok');
  } finally {
    lock.releaseLock();
  }
}

// Open the web app URL in a browser to check it is running.
function doGet() {
  return ContentService.createTextOutput('FridgeLog collector is running.');
}
