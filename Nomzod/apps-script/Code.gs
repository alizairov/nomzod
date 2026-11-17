/**
 * Nomzod — ariza qabul qilish (Google Apps Script web app).
 * Google Sheet ichidan Extensions > Apps Script orqali oching (bound script).
 * Script Properties (ixtiyoriy): TELEGRAM_TOKEN, CHAT_ID
 */
var SHEET_NAME = 'Arizalar';
var HEADERS = ['Sana', 'Ism', 'Telefon', 'Kompaniya', 'Xizmat', 'Izoh', 'Sahifa'];

function doPost(e) {
  try {
    var p = (e && e.parameter) || {};
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if (sh.getLastRow() === 0) {
      sh.appendRow(HEADERS);
      sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    var row = [p.date || new Date().toISOString(), p.name || '', p.phone || '', p.company || '', p.service || '', p.comment || '', p.page || ''];
    // Telefon "+" bilan boshlangani uchun formula deb o'qilmasligi uchun matn sifatida yozamiz
    sh.appendRow(row);
    sh.getRange(sh.getLastRow(), 3).setNumberFormat('@').setValue(row[2]);
    notifyTelegram_(row);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doGet() {
  return json_({ ok: true, service: 'Nomzod forma' });
}

function notifyTelegram_(r) {
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty('TELEGRAM_TOKEN');
  var chat = props.getProperty('CHAT_ID');
  if (!token || !chat) return;
  var text = 'Yangi ariza (Nomzod)\n' +
    'Ism: ' + r[1] + '\nTelefon: ' + r[2] + '\nKompaniya: ' + r[3] +
    '\nXizmat: ' + r[4] + '\nIzoh: ' + r[5];
  try {
    UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ chat_id: chat, text: text }),
      muteHttpExceptions: true
    });
  } catch (e) {}
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
