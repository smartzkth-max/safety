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
  'Answers'
];

function doGet(e) {
  const params = e.parameter || {};
  const action = params.action || 'ping';

  try {
    if (action === 'submit') return handleSubmit_(params);
    if (action === 'list') {
      return output_({
        ok: true,
        results: getResults_(params.sessionId || '')
      }, params.callback);
    }
    return output_({ ok: true, message: 'Safety quiz API is running.' }, params.callback);
  } catch (err) {
    return output_({ ok: false, error: String(err && err.message ? err.message : err) }, params.callback);
  }
}

function handleSubmit_(params) {
  const sheet = getSheet_();
  ensureHeaders_(sheet);

  sheet.appendRow([
    new Date(),
    params.sessionId || '',
    params.topic || '',
    params.lang || '',
    params.name || '',
    params.age || '',
    params.role || '',
    params.exp || '',
    params.company || '',
    params.site || '',
    Number(params.score || 0),
    Number(params.total || 0),
    params.answers || ''
  ]);

  return output_({ ok: true }, params.callback);
}

function getResults_(sessionId) {
  const sheet = getSheet_();
  ensureHeaders_(sheet);

  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];

  return values.slice(1)
    .filter(row => !sessionId || String(row[1]) === String(sessionId))
    .map(row => ({
      timestamp: formatTimestamp_(row[0]),
      sessionId: String(row[1] || ''),
      topic: String(row[2] || ''),
      lang: String(row[3] || ''),
      name: String(row[4] || ''),
      age: String(row[5] || ''),
      role: String(row[6] || ''),
      exp: String(row[7] || ''),
      company: String(row[8] || ''),
      site: String(row[9] || ''),
      score: Number(row[10] || 0),
      total: Number(row[11] || 0),
      answers: String(row[12] || '')
    }));
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  return sheet;
}

function ensureHeaders_(sheet) {
  const firstRow = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  const hasHeaders = firstRow.some(value => String(value || '').trim() !== '');
  if (!hasHeaders) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
  }
}

function output_(obj, callback) {
  const json = JSON.stringify(obj);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function formatTimestamp_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  }
  return String(value);
}
