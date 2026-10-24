/**
 * Halloweenjakten – statistik
 *
 * Klistra in det här i Google Sheet-filen: Tillägg → Apps Script.
 * Det räknar bara siffror (t.ex. "view:home" eller "found:k4ws7n").
 * Inga namn, telefonnummer eller annan personlig information sparas.
 * Siffrorna hamnar på en egen flik som heter "Statistik".
 */
const TAB = 'Statistik';
const KEY_OK = /^[a-z0-9:_-]{1,40}$/;  // bara korta, enkla nycklar
const MAX_ROWS = 400;                  // skydd mot skräp

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(TAB);
  if (!sh) {
    // Läggs sist, så att belöningshus-fliken förblir först.
    sh = ss.insertSheet(TAB, ss.getNumSheets());
    sh.appendRow(['Nyckel', 'Antal', 'Senast']);
    sh.setFrozenRows(1);
  }
  return sh;
}

// Appen skickar "+1" hit.
function doPost(e) {
  const raw = (e && e.postData && e.postData.contents) || (e && e.parameter && e.parameter.k) || '';
  const keys = String(raw).split(',')
    .map(function (k) { return k.trim().toLowerCase(); })
    .filter(function (k) { return KEY_OK.test(k); })
    .slice(0, 10);
  if (!keys.length) return json_({ ok: false });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000); // flera mobiler samtidigt: en i taget
  try {
    const sh = sheet_();
    const last = sh.getLastRow();
    const rows = last > 1 ? sh.getRange(2, 1, last - 1, 2).getValues() : [];
    const now = new Date();
    keys.forEach(function (k) {
      const i = rows.findIndex(function (r) { return r[0] === k; });
      if (i >= 0) {
        rows[i][1] = (Number(rows[i][1]) || 0) + 1;
        sh.getRange(i + 2, 2, 1, 2).setValues([[rows[i][1], now]]);
      } else if (rows.length < MAX_ROWS) {
        rows.push([k, 1]);
        sh.appendRow([k, 1, now]);
      }
    });
  } finally {
    lock.releaseLock();
  }
  return json_({ ok: true });
}

// Arrangörssidan hämtar siffrorna här.
function doGet() {
  const sh = sheet_();
  const last = sh.getLastRow();
  const counts = {};
  if (last > 1) {
    sh.getRange(2, 1, last - 1, 2).getValues().forEach(function (r) {
      if (r[0]) counts[r[0]] = Number(r[1]) || 0;
    });
  }
  return json_({ counts: counts, at: new Date().toISOString() });
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
