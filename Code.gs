/**
 * Anonyme Mitarbeiter-Feedback-Umfrage – Backend
 * An ein Google Sheet gebunden (Erweiterungen → Apps Script).
 * Bereitstellen als Web-App: Ausführen als "Ich", Zugriff "Jeder".
 *
 * Es werden KEINE personenbezogenen Daten gespeichert
 * (keine IP, keine E-Mail, kein User-Agent).
 */

var SHEET_NAME = 'Antworten';
var MAX_LEN = 1000;
var HEADERS = ['Zeitstempel', 'Datum', 'Wochentag', 'Bewertung', 'Stimmung', 'Was lief gut', 'Was lief nicht gut'];
var LABELS = { 1: 'Schlecht', 2: 'Mäßig', 3: 'Okay', 4: 'Gut', 5: 'Super' };
var WOCHENTAGE = ['Sonntag', 'Montag', 'Dienstag', 'Mittwoch', 'Donnerstag', 'Freitag', 'Samstag'];

function doGet() {
  return json_({ ok: true });
}

function doPost(e) {
  var p = (e && e.parameter) || {};

  // Honeypot: Bots füllen das versteckte Feld aus → stillschweigend ignorieren
  if (p.website) return json_({ ok: true });

  var bewertung = parseInt(p.bewertung, 10);
  if (!(bewertung >= 1 && bewertung <= 5) || String(bewertung) !== String(p.bewertung).trim()) {
    return json_({ ok: false, error: 'Ungültige Bewertung' });
  }

  var gut = clean_(p.gut);
  var schlecht = clean_(p.schlecht);

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (err) {
    return json_({ ok: false, error: 'Server ausgelastet, bitte erneut versuchen' });
  }

  try {
    var sheet = getSheet_();
    var now = new Date();
    var tz = Session.getScriptTimeZone();
    sheet.appendRow([
      now,
      Utilities.formatDate(now, tz, 'yyyy-MM-dd'),
      WOCHENTAGE[parseInt(Utilities.formatDate(now, tz, 'u'), 10) % 7],
      bewertung,
      LABELS[bewertung],
      gut,
      schlecht
    ]);
    return json_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

/** Kürzt auf MAX_LEN Zeichen und verhindert Formel-Injection. */
function clean_(value) {
  var s = String(value == null ? '' : value).trim().slice(0, MAX_LEN);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.getRange('A:A').setNumberFormat('yyyy-MM-dd HH:mm:ss');
    sheet.setColumnWidths(6, 2, 400);
    sheet.getRange('F:G').setWrap(true);
  }
  return sheet;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
