const SHEET_NAME = 'Responses';

const HEADERS = [
  'Timestamp',
  'Session ID',
  'Meeting Topic',
  'Language',
  'Name',
  'Age',
  'Role',
  'Experience',
  'Company',
  'Site',
  'Score',
  'Total',
  'Answers',
  'Started At',
  'Duration Seconds',
  'Duration'
];

function doGet(e) {
  const params = e.parameter || {};
  const action = params.action || 'ping';

  try {
    if (action === 'submit') return handleSubmit_(params);
    if (action === 'list') return output_({ ok: true, results: getResults_(params.sessionId || '') }, params.callback);
    return output_({ ok: true, message: 'Safety quiz API is running.' }, params.callback);
  } catch (err) {
    return output_({ ok: false, error: String(err && err.message ? err.message : err) }, params.callback);
  }
}

function handleSubmit_(params) {
  const sheet = getSheet_();
  ensureHeaders_(sheet);

  const durationSeconds = Number(params.durationSec || 0);
  const rowObject = {
    'Timestamp': new Date(),
    'Session ID': params.sessionId || '',
    'Meeting Topic': params.topic || '',
    'Language': params.lang || '',
    'Name': params.name || '',
    'Age': params.age || '',
    'Role': params.role || '',
    'Experience': params.exp || '',
    'Company': params.company || '',
    'Site': params.site || '',
    'Score': Number(params.score || 0),
    'Total': Number(params.total || 0),
    'Answers': params.answers || '',
    'Started At': params.startedAt || '',
    'Duration Seconds': durationSeconds || '',
    'Duration': formatDuration_(durationSeconds)
  };

  appendObjectRow_(sheet, rowObject);
  return output_({ ok: true }, params.callback);
}

function getResults_(sessionId) {
  const sheet = getSheet_();
  ensureHeaders_(sheet);

  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];

  const headers = values[0].map(String);
  const idx = name => headers.indexOf(name);
  const get = (row, name) => idx(name) >= 0 ? row[idx(name)] : '';

  return values.slice(1)
    .filter(row => !sessionId || String(get(row, 'Session ID')) === String(sessionId))
    .map(row => {
      const durationSec = Number(get(row, 'Duration Seconds') || 0);
      return {
        timestamp: formatTimestamp_(get(row, 'Timestamp')),
        sessionId: String(get(row, 'Session ID') || ''),
        topic: String(get(row, 'Meeting Topic') || ''),
        lang: String(get(row, 'Language') || ''),
        name: String(get(row, 'Name') || ''),
        age: String(get(row, 'Age') || ''),
        role: String(get(row, 'Role') || ''),
        exp: String(get(row, 'Experience') || ''),
        company: String(get(row, 'Company') || ''),
        site: String(get(row, 'Site') || ''),
        score: Number(get(row, 'Score') || 0),
        total: Number(get(row, 'Total') || 0),
        answers: String(get(row, 'Answers') || ''),
        startedAt: String(get(row, 'Started At') || ''),
        durationSec: durationSec,
        duration: String(get(row, 'Duration') || formatDuration_(durationSec))
      };
    })
    .sort((a, b) => {
      if (b.score !== a.score) return b.score - a.score;
      const da = a.durationSec || 999999;
      const db = b.durationSec || 999999;
      if (da !== db) return da - db;
      return String(a.timestamp).localeCompare(String(b.timestamp));
    });
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  return sheet;
}

function ensureHeaders_(sheet) {
  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const existingRow = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const existing = existingRow.map(v => String(v || '').trim()).filter(Boolean);

  if (existing.length === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
    return;
  }

  const missing = HEADERS.filter(header => existing.indexOf(header) === -1);
  if (missing.length > 0) {
    sheet.getRange(1, existing.length + 1, 1, missing.length).setValues([missing]);
  }

  sheet.setFrozenRows(1);
}

function appendObjectRow_(sheet, obj) {
  ensureHeaders_(sheet);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  const row = headers.map(header => obj[header] !== undefined ? obj[header] : '');
  sheet.appendRow(row);
}

function output_(obj, callback) {
  const json = JSON.stringify(obj);
  if (callback) return ContentService.createTextOutput(callback + '(' + json + ');').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function formatTimestamp_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  }
  return String(value);
}

function formatDuration_(seconds) {
  const n = Number(seconds || 0);
  if (!n) return '';
  const minutes = Math.floor(n / 60);
  const sec = n % 60;
  return minutes ? minutes + 'm ' + String(sec).padStart(2, '0') + 's' : sec + 's';
}
