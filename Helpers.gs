/* ===================== HELPERS ===================== */

function sheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function ensureHeaders_(sh, headers) {
  const width = Math.max(sh.getLastColumn(), headers.length, 1);
  const row1 = sh.getRange(1, 1, 1, width).getValues()[0];
  const map = {};

  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    headers.forEach((h, i) => (map[h.toLowerCase()] = i + 1));
    sh.setFrozenRows(1);
    return map;
  }

  const existing = row1.map((h) => str_(h).toLowerCase());
  let col = sh.getLastColumn();

  headers.forEach((h) => {
    const key = h.toLowerCase();
    if (!existing.includes(key)) {
      col += 1;
      sh.getRange(1, col).setValue(h);
      existing.push(key);
    }
  });

  const refreshed = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];
  refreshed.forEach((h, i) => {
    const key = str_(h).toLowerCase();
    if (key) map[key] = i + 1;
  });

  sh.setFrozenRows(1);
  return map;
}

function readData_(sh, startRow) {
  const lastRow = sh.getLastRow();
  if (lastRow < startRow) return [];
  const width = sh.getLastColumn();
  return sh.getRange(startRow, 1, lastRow - startRow + 1, width).getValues();
}

function appendByMap_(sh, map, obj) {
  const width = Math.max(sh.getLastColumn(), Object.keys(map).length);
  const row = new Array(width).fill('');

  Object.keys(obj).forEach((k) => {
    const col = map[k.toLowerCase()];
    if (col) row[col - 1] = obj[k];
  });

  sh.appendRow(row);
}

function deleteByField_(sh, map, field, value, insensitive) {
  const col = map[field.toLowerCase()];
  if (!col) return false;

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return false;

  const target = str_(value);
  const vals = sh.getRange(2, col, lastRow - 1, 1).getValues();
  const rows = [];

  for (let i = 0; i < vals.length; i++) {
    const cell = str_(vals[i][0]);
    const match = insensitive ? cell.toUpperCase() === target.toUpperCase() : cell === target;
    if (match) rows.push(i + 2);
  }

  if (!rows.length) return false;
  rows.sort((a, b) => b - a).forEach((r) => sh.deleteRow(r));
  return true;
}

function formatRegistroFecha_(v) {
  if (v == null || v === '') return '';
  if (v instanceof Date && !isNaN(v.getTime())) {
    return Utilities.formatDate(v, 'America/Mexico_City', "yyyy-MM-dd'T'HH:mm:ss");
  }
  return str_(v);
}

function cell_(row, map, field) {
  const col = map[field.toLowerCase()];
  if (!col) return '';
  const val = row[col - 1];
  if (field.toLowerCase() === 'fecha') return formatRegistroFecha_(val);
  return str_(val);
}

function str_(v) {
  return String(v == null ? '' : v).trim();
}

function toNum_(v) {
  if (typeof v === 'number') return v;
  const n = parseFloat(str_(v).replace(/,/g, ''));
  return isNaN(n) ? NaN : n;
}

function uuid_(prefix) {
  return prefix + '_' + Utilities.getUuid();
}

function formatHora_(date) {
  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone() || 'America/Mexico_City',
    'h:mm:ss a'
  );
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON
  );
}
