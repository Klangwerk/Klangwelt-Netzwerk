/**
 * Klangwelt Netzwerk — Automatische Übernahme von Formular-Anmeldungen
 * ins Google Sheet (inkl. Foto-Upload nach Google Drive, jetzt auch
 * zwei Fotos) UND von gemeldeten "Erfolgen" (inkl. Zitat) in einen
 * eigenen Tab.
 *
 * NEU nötig, damit alles funktioniert:
 *
 * A) Haupt-Tab ("Tabelle1"): fünf neue Spalten ganz rechts anlegen,
 *    in dieser Reihenfolge, direkt nach der bestehenden "Foto"-Spalte:
 *    Foto2 | Setlänge | Equipment | Verifiziert | Verfuegbarkeit
 *    ("Verifiziert" trägst du selbst "ja" ein, wenn du einen Eintrag
 *    geprüft hast — die Webseite zeigt dann ein "✓ verifiziert"-Zeichen.
 *    "Verfuegbarkeit" wird automatisch mit den im Kalender-Widget
 *    angeklickten Terminen befüllt, kommagetrennt als YYYY-MM-DD.)
 *
 * B) "Erfolge"-Tab: eine neue Spalte "Zitat" ganz rechts anlegen, nach
 *    "Beschreibung". Kopfzeile also:
 *    Titel | Datum | Wer | Aufgabe1 | MitWem | Aufgabe2 | Beschreibung | Zitat
 *
 * C) Diesen kompletten Code hier in den Apps-Script-Editor kopieren
 *    (alten Code komplett ersetzen: Strg+A, Entf, dann einfügen).
 *
 * D) Speichern (Strg+S), dann neu bereitstellen (Bereitstellen >
 *    Bereitstellungen verwalten > Stift-Symbol > Neue Version >
 *    Bereitstellen).
 */

var SHEET_ID = '1WqlhxGwOcHN4wTa1RqY4dEERZr7tEbSO1xE5QQFp5EE';
var SHEET_NAME = 'Tabelle1';
var ERFOLGE_SHEET_NAME = 'Erfolge';
var BENACHRICHTIGUNG_AN = 'torsten@gut-steimke.de';
var FOTO_ORDNER_ID = '1MXm04zHy6bft5hh4rPRcHgYcex-ZOTwi';

function doPost(e) {
  if (e.parameter['Formular'] === 'Erfolg') {
    return speichereErfolg(e);
  }
  return speichereAnmeldung(e);
}

function speichereAnmeldung(e) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    return ContentService.createTextOutput('Tabellenblatt "' + SHEET_NAME + '" nicht gefunden.');
  }

  var sparteArt = [e.parameter['Sparte'], e.parameter['Art']].filter(function(v){ return v; }).join(' / ');
  var fotoErgebnis = speichereFoto(e, 'FotoBase64', 'FotoTyp', 'FotoName');
  var foto2Ergebnis = speichereFoto(e, 'Foto2Base64', 'Foto2Typ', 'Foto2Name');

  var zeile = [
    e.parameter['Kategorie'] || '',
    e.parameter['Name'] || '',
    e.parameter['PLZ-Bereich'] || '',
    sparteArt,
    e.parameter['Kontakt'] || '',
    e.parameter['Website'] || '',
    e.parameter['Beschreibung'] || '',
    fotoErgebnis.url,
    foto2Ergebnis.url,
    e.parameter['Setlänge'] || '',
    e.parameter['Equipment'] || '',
    '', // Verifiziert — bleibt leer, wird von Hand im Sheet gepflegt
    e.parameter['Verfuegbarkeit'] || ''
  ];

  sheet.appendRow(zeile);

  if (BENACHRICHTIGUNG_AN) {
    var betreff = 'Neuer Eintrag im Klangwelt Netzwerk: ' + (e.parameter['Name'] || '');
    var text = 'Kategorie: ' + (e.parameter['Kategorie'] || '-') + '\n' +
      'Name: ' + (e.parameter['Name'] || '-') + '\n' +
      'PLZ-Bereich: ' + (e.parameter['PLZ-Bereich'] || '-') + '\n' +
      'Sparte / Art: ' + (sparteArt || '-') + '\n' +
      'Kontakt: ' + (e.parameter['Kontakt'] || '-') + '\n' +
      'Website: ' + (e.parameter['Website'] || '-') + '\n' +
      'Setlänge: ' + (e.parameter['Setlänge'] || '-') + '\n' +
      'Equipment: ' + (e.parameter['Equipment'] || '-') + '\n' +
      'Verfügbare Termine: ' + (e.parameter['Verfuegbarkeit'] || '-') + '\n' +
      'Beschreibung: ' + (e.parameter['Beschreibung'] || '-') + '\n' +
      '\n--- Foto-Diagnose ---\n' +
      'Foto 1: ' + (fotoErgebnis.url || '(keine)') + ' — ' + fotoErgebnis.debug + '\n' +
      'Foto 2: ' + (foto2Ergebnis.url || '(keine)') + ' — ' + foto2Ergebnis.debug;
    MailApp.sendEmail(BENACHRICHTIGUNG_AN, betreff, text);
  }

  return ContentService.createTextOutput('OK');
}

function speichereErfolg(e) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(ERFOLGE_SHEET_NAME);
  if (!sheet) {
    return ContentService.createTextOutput('Tabellenblatt "' + ERFOLGE_SHEET_NAME + '" nicht gefunden. Bitte anlegen (siehe Anleitung oben im Code).');
  }

  var zeile = [
    e.parameter['Titel'] || '',
    e.parameter['Datum'] || '',
    e.parameter['Wer'] || '',
    e.parameter['Aufgabe1'] || '',
    e.parameter['MitWem'] || '',
    e.parameter['Aufgabe2'] || '',
    e.parameter['Beschreibung'] || '',
    e.parameter['Zitat'] || ''
  ];

  sheet.appendRow(zeile);

  if (BENACHRICHTIGUNG_AN) {
    var betreff = 'Neuer Erfolg gemeldet: ' + (e.parameter['Titel'] || '');
    var text = 'Titel: ' + (e.parameter['Titel'] || '-') + '\n' +
      'Datum: ' + (e.parameter['Datum'] || '-') + '\n' +
      'Wer: ' + (e.parameter['Wer'] || '-') + ' (' + (e.parameter['Aufgabe1'] || 'Sonstiges') + ')\n' +
      'Mit wem: ' + (e.parameter['MitWem'] || '-') + ' (' + (e.parameter['Aufgabe2'] || 'Sonstiges') + ')\n' +
      'Beschreibung: ' + (e.parameter['Beschreibung'] || '-') + '\n' +
      'Zitat: ' + (e.parameter['Zitat'] || '-');
    MailApp.sendEmail(BENACHRICHTIGUNG_AN, betreff, text);
  }

  return ContentService.createTextOutput('OK');
}

// Speichert ein per Formular mitgeschicktes Foto (Base64) in Google Drive.
// base64Feld/typFeld/nameFeld erlauben es, dieselbe Funktion für Foto 1
// und Foto 2 zu benutzen (unterschiedliche Parameter-Namen).
function speichereFoto(e, base64Feld, typFeld, nameFeld) {
  var base64 = e.parameter[base64Feld];
  if (!base64) {
    return { url: '', debug: 'Kein ' + base64Feld + '-Parameter angekommen (kein Foto ausgewählt oder Übertragung fehlgeschlagen).' };
  }
  if (!FOTO_ORDNER_ID) {
    return { url: '', debug: 'FOTO_ORDNER_ID ist im Code leer.' };
  }
  try {
    var typ = e.parameter[typFeld] || 'image/jpeg';
    var name = e.parameter[nameFeld] || ('foto_' + new Date().getTime() + '.jpg');
    var laenge = base64.length;
    var bytes = Utilities.base64Decode(base64);
    var blob = Utilities.newBlob(bytes, typ, name);
    var ordner = DriveApp.getFolderById(FOTO_ORDNER_ID);
    var datei = ordner.createFile(blob);
    datei.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var url = 'https://drive.google.com/uc?export=view&id=' + datei.getId();
    return { url: url, debug: 'Erfolgreich gespeichert. Base64-Länge: ' + laenge + ' Zeichen. Datei-ID: ' + datei.getId() };
  } catch (fehler) {
    return { url: '', debug: 'FEHLER beim Speichern: ' + fehler.toString() };
  }
}

function testDoPost() {
  var fakeEvent = {
    parameter: {
      'Kategorie': 'Künstler',
      'Name': 'Test Eintrag',
      'PLZ-Bereich': '371',
      'Sparte': 'Musik — Band',
      'Art': '',
      'Kontakt': 'test@test.de',
      'Website': '',
      'Beschreibung': 'Testeintrag',
      'Setlänge': '2 x 45 Minuten',
      'Equipment': 'Eigenes Equipment vorhanden'
    }
  };
  var ergebnis = doPost(fakeEvent);
  Logger.log(ergebnis.getContent());
}

function testDriveZugriff() {
  var ordner = DriveApp.getFolderById(FOTO_ORDNER_ID);
  Logger.log('Ordner gefunden: ' + ordner.getName());
}
