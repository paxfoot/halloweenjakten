/**
 * Halloweenjakten – ändra status för belöningshusen
 *
 * Klistra in det här i BELÖNINGSHUS-arket: Tillägg → Apps Script.
 * (Inte i statistik-arket – det har sitt eget skript.)
 *
 * Skriptet gör två saker:
 *  1. Lägger till menyn "🎃 Halloweenjakten" i arket, där du hittar länken
 *     som du skickar till belöningshusen.
 *  2. Tar emot ändringar från länkens sida i appen och skriver Öppet / Slut / Stängt
 *     i kolumn B för rätt hus.
 *
 * Den hemliga koden i länken sparas bara i skriptet (inte i arket och inte på GitHub).
 */
const STATUSAR = ['Öppet', 'Slut', 'Stängt'];
const APP = 'https://paxfoot.github.io/halloweenjakten/';

// Menyn i arket (syns när arket öppnas på nytt)
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🎃 Halloweenjakten')
    .addItem('Visa länk för belöningshus', 'visaLank')
    .addItem('Byt till en ny länk (den gamla slutar fungera)', 'nyLank')
    .addToUi();
}

function kod_() {
  const p = PropertiesService.getScriptProperties();
  let k = p.getProperty('KOD');
  if (!k) {
    k = Utilities.getUuid().replace(/-/g, '').slice(0, 12);
    p.setProperty('KOD', k);
  }
  return k;
}

function visaLank() {
  const url = APP + '#/hus/' + kod_();
  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:sans-serif;font-size:14px">' +
    '<p>Skicka den här länken till alla belöningshus:</p>' +
    '<p style="font-size:15px;word-break:break-all"><b><a href="' + url + '" target="_blank">' + url + '</a></b></p>' +
    '<p style="color:#666">Alla som har länken kan ändra status för husen, så dela den bara med belöningshusen.</p>' +
    '</div>').setWidth(560).setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(html, 'Länk för belöningshus');
}

function nyLank() {
  PropertiesService.getScriptProperties().deleteProperty('KOD');
  visaLank();
}

function giltig_(kod) {
  const k = PropertiesService.getScriptProperties().getProperty('KOD');
  return !!k && !!kod && String(kod) === k;
}

// Appen kollar om länken fungerar
function doGet(e) {
  const kod = e && e.parameter ? e.parameter.kod : '';
  return json_({ ok: giltig_(kod) });
}

// Appen ändrar status för ett hus: { kod, hus, status }
function doPost(e) {
  let d = {};
  try { d = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) {}
  if (!giltig_(d.kod)) return json_({ ok: false, fel: 'kod' });
  if (STATUSAR.indexOf(d.status) < 0) return json_({ ok: false, fel: 'status' });
  const hus = String(d.hus || '').trim();
  if (!hus) return json_({ ok: false, fel: 'hus' });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]; // första fliken = listan med hus
    const last = sh.getLastRow();
    if (last < 2) return json_({ ok: false, fel: 'hus' });
    const namn = sh.getRange(2, 1, last - 1, 1).getValues();
    const i = namn.findIndex(function (r) { return String(r[0]).trim() === hus; });
    if (i < 0) return json_({ ok: false, fel: 'hus' });
    sh.getRange(i + 2, 2).setValue(d.status);
    return json_({ ok: true, hus: hus, status: d.status });
  } finally {
    lock.releaseLock();
  }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
