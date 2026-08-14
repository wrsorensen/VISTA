const VISTA_SPREADSHEET_ID = '1hy0Bb7-GVVLFGnbtTiry99uSl3iaReKX2jRSCoL0K7k';
const VISTA_API_VERSION = 'v1';

const MILEAGE_HEADERS = [
  'Timestamp',
  'Mileage ID',
  'Date',
  'User',
  'Vehicle',
  'Work Area',
  'Route / Property',
  'Start Odometer',
  'End Odometer',
  'Miles',
  'Stops JSON',
  'Purpose',
  'Notes',
  'Source',
  'Save Status',
  'Deleted',
  'Deleted At',
  'Delete Reason',
  'Updated At',
  'Updated Source'
];

const NOTES_HEADERS = [
  'Timestamp',
  'Note ID',
  'Date',
  'User',
  'Type',
  'Property',
  'Related Module',
  'Note',
  'Follow Up?',
  'Status',
  'Source',
  'Deleted',
  'Deleted At',
  'Delete Reason',
  'Updated At',
  'Updated Source'
];

function doGet(e) {
  const params = (e && e.parameter) || {};
  const callback = params.callback || '';
  let response;

  try {
    const action = params.action || 'ping';
    if (action === 'ping') {
      response = { ok: true, apiVersion: VISTA_API_VERSION, app: 'VISTA' };
    } else if (action === 'setup') {
      response = setupVistaSheets_();
    } else if (action === 'saveMileage') {
      response = saveMileage_(parsePayload_(params.payload));
    } else if (action === 'saveNote') {
      response = saveNote_(parsePayload_(params.payload));
    } else if (action === 'recentMileage') {
      response = recentMileage_(Number(params.limit || 20));
    } else {
      response = { ok: false, error: 'Unknown action: ' + action };
    }
  } catch (err) {
    response = { ok: false, error: String(err && err.message ? err.message : err) };
  }

  return jsonp_(callback, response);
}

function parsePayload_(raw) {
  if (!raw) throw new Error('Missing payload');
  return JSON.parse(raw);
}

function jsonp_(callback, payload) {
  const body = callback
    ? callback + '(' + JSON.stringify(payload) + ');'
    : JSON.stringify(payload);
  return ContentService
    .createTextOutput(body)
    .setMimeType(callback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
}

function ss_() {
  return SpreadsheetApp.openById(VISTA_SPREADSHEET_ID);
}

function setupVistaSheets_() {
  const ss = ss_();
  ensureSheet_(ss, 'Mileage', MILEAGE_HEADERS);
  ensureSheet_(ss, 'Notes', NOTES_HEADERS);
  return { ok: true, setup: true };
}

function ensureSheet_(ss, name, headers) {
  const sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  const current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const needsHeaders = headers.some((header, idx) => current[idx] !== header);
  if (needsHeaders) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#0B3557')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  sheet.autoResizeColumns(1, headers.length);
  return sheet;
}

function saveMileage_(payload) {
  if (!payload) throw new Error('Missing mileage payload');
  const ss = ss_();
  const sheet = ensureSheet_(ss, 'Mileage', MILEAGE_HEADERS);
  const miles = calculateMiles_(payload.startOdometer, payload.endOdometer, payload.totalMiles);
  if (miles === '') throw new Error('Invalid mileage: end odometer must be greater than or equal to start');

  const now = new Date();
  const row = [
    now,
    payload.id || 'mileage-' + now.getTime(),
    payload.date || '',
    payload.user || 'Alysha',
    payload.vehicle || '',
    payload.workArea || '',
    payload.route || '',
    numberOrBlank_(payload.startOdometer),
    numberOrBlank_(payload.endOdometer),
    miles,
    JSON.stringify(payload.stops || []),
    payload.purpose || 'Work Travel',
    payload.notes || '',
    payload.source || 'VISTA',
    'saved',
    false,
    '',
    '',
    now,
    payload.source || 'VISTA'
  ];
  sheet.appendRow(row);
  return { ok: true, type: 'mileage', id: row[1], row: sheet.getLastRow(), miles: miles };
}

function saveNote_(payload) {
  if (!payload) throw new Error('Missing note payload');
  const ss = ss_();
  const sheet = ensureSheet_(ss, 'Notes', NOTES_HEADERS);
  const now = new Date();
  const row = [
    now,
    payload.id || 'note-' + now.getTime(),
    payload.date || Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    payload.user || 'Alysha',
    payload.type || '',
    payload.property || '',
    payload.relatedModule || '',
    payload.note || '',
    Boolean(payload.followUp),
    payload.status || 'Open',
    payload.source || 'VISTA',
    false,
    '',
    '',
    now,
    payload.source || 'VISTA'
  ];
  sheet.appendRow(row);
  return { ok: true, type: 'note', id: row[1], row: sheet.getLastRow() };
}

function recentMileage_(limit) {
  const sheet = ss_().getSheetByName('Mileage');
  if (!sheet || sheet.getLastRow() < 2) return { ok: true, rows: [] };
  const safeLimit = Math.max(1, Math.min(limit || 20, 100));
  const lastRow = sheet.getLastRow();
  const startRow = Math.max(2, lastRow - safeLimit + 1);
  const values = sheet.getRange(startRow, 1, lastRow - startRow + 1, MILEAGE_HEADERS.length).getValues();
  const rows = values.map(rowToMileage_).reverse();
  return { ok: true, rows: rows };
}

function rowToMileage_(row) {
  const obj = {};
  MILEAGE_HEADERS.forEach((header, idx) => {
    obj[header] = row[idx];
  });
  return obj;
}

function calculateMiles_(start, end, suppliedMiles) {
  const supplied = numberOrBlank_(suppliedMiles);
  if (supplied !== '') return supplied;
  const startNum = numberOrBlank_(start);
  const endNum = numberOrBlank_(end);
  if (startNum === '' || endNum === '') return '';
  const miles = endNum - startNum;
  return miles >= 0 ? Math.round(miles * 10) / 10 : '';
}

function numberOrBlank_(value) {
  if (value === null || value === undefined || value === '') return '';
  const num = Number(value);
  return Number.isFinite(num) ? num : '';
}
