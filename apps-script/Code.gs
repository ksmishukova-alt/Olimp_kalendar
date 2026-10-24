/**
 * Олимпиадный календарь — серверная часть на Google Apps Script.
 * Данные лежат в Google Таблице (лист «data»: col | id | json | updatedAt).
 * Страница на GitHub Pages обращается сюда POST-запросом {fn, args}.
 *
 * Настройка (один раз):
 *  1. Вставьте этот код в новый проект на script.google.com.
 *  2. Впишите свой PIN в функцию setPin ниже, выберите её в списке и нажмите «Выполнить».
 *  3. «Начать развёртывание» → «Новое развёртывание» → тип «Веб-приложение»,
 *     «Запуск от имени»: Я, «У кого есть доступ»: Все. Скопируйте адрес …/exec.
 */

// Откуда взять начальные данные (62 олимпиады и дети), когда таблица ещё пустая.
const SEED_URL = 'https://ksmishukova-alt.github.io/Olimp_kalendar/seed.json';
const COLS = ['olympiads', 'kids', 'plans'];
const ID_RE = /^[A-Za-z0-9_.:~-]{1,120}$/;

function setPin() {
  PropertiesService.getScriptProperties().setProperty('EDIT_PIN', '1234'); // ← замените 1234 на свой PIN
}

function doGet() {
  return out_({ ok: true, result: getState() });
}

function doPost(e) {
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const fns = { getState: getState, save: save, checkPin: checkPin };
    const f = fns[req.fn];
    if (!f) throw new Error('Неизвестная команда');
    return out_({ ok: true, result: f.apply(null, req.args || []) });
  } catch (err) {
    return out_({ ok: false, error: String((err && err.message) || err) });
  }
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function sheet_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SHEET_ID');
  let ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) ss = SpreadsheetApp.create('Олимпиадный календарь — данные');
  if (!id) props.setProperty('SHEET_ID', ss.getId());
  let sh = ss.getSheetByName('data');
  if (!sh) {
    sh = ss.insertSheet('data');
    sh.appendRow(['col', 'id', 'json', 'updatedAt']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function rows_(sh) {
  const n = sh.getLastRow();
  return n < 2 ? [] : sh.getRange(2, 1, n - 1, 4).getValues();
}

function seed_(sh) {
  const data = JSON.parse(UrlFetchApp.fetch(SEED_URL).getContentText());
  const now = new Date().toISOString();
  const out = [];
  COLS.forEach(function (c) {
    Object.keys(data[c] || {}).forEach(function (id) { out.push([c, id, JSON.stringify(data[c][id]), now]); });
  });
  if (out.length) sh.getRange(2, 1, out.length, 4).setValues(out);
}

function getState() {
  const sh = sheet_();
  let rows = rows_(sh);
  if (!rows.length) {
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      if (!rows_(sh).length) seed_(sh);
      rows = rows_(sh);
    } finally { lock.releaseLock(); }
  }
  const st = { olympiads: {}, kids: {}, plans: {} };
  rows.forEach(function (r) {
    if (!st[r[0]] || !r[1]) return;
    try { st[r[0]][r[1]] = JSON.parse(r[2]); } catch (e) {}
  });
  return st;
}

function checkPin(pin) {
  const real = PropertiesService.getScriptProperties().getProperty('EDIT_PIN');
  if (!real) throw new Error('PIN не задан: выполните функцию setPin в редакторе скрипта');
  return String(pin || '') === real;
}

function merge_(a, b) {
  const out = {};
  Object.keys(a || {}).forEach(function (k) { out[k] = a[k]; });
  Object.keys(b || {}).forEach(function (k) {
    const v = b[k];
    out[k] = v && typeof v === 'object' && !Array.isArray(v) && out[k] && typeof out[k] === 'object' && !Array.isArray(out[k]) ? merge_(out[k], v) : v;
  });
  return out;
}

function save(pin, writes) {
  if (!checkPin(pin)) throw new Error('pin');
  if (!Array.isArray(writes) || !writes.length || writes.length > 60) throw new Error('Неверные данные');
  writes.forEach(function (w) {
    if (!w || ['set', 'update', 'delete'].indexOf(w.op) < 0 || COLS.indexOf(w.col) < 0 || !ID_RE.test(String(w.id || ''))) throw new Error('Неверные данные');
    if (w.op !== 'delete' && (!w.data || typeof w.data !== 'object' || Array.isArray(w.data))) throw new Error('Неверные данные');
  });
  if (JSON.stringify(writes).length > 200000) throw new Error('Слишком большой запрос');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_();
    if (!rows_(sh).length) seed_(sh);
    const rows = rows_(sh);
    const index = {};
    rows.forEach(function (r, i) { index[r[0] + '|' + r[1]] = i + 2; });
    const now = new Date().toISOString();
    const toDelete = [];
    writes.forEach(function (w) {
      const key = w.col + '|' + w.id;
      const rowNum = index[key];
      if (w.op === 'delete') { if (rowNum) { toDelete.push(rowNum); delete index[key]; } return; }
      let data = w.data;
      if (w.op === 'update' && rowNum) {
        let old = {};
        try { old = JSON.parse(sh.getRange(rowNum, 3).getValue()); } catch (e) {}
        data = merge_(old, w.data);
      }
      const values = [[w.col, String(w.id), JSON.stringify(data), now]];
      if (rowNum) sh.getRange(rowNum, 1, 1, 4).setValues(values);
      else { sh.appendRow(values[0]); index[key] = sh.getLastRow(); }
    });
    toDelete.sort(function (a, b) { return b - a; }).forEach(function (r) { sh.deleteRow(r); });
  } finally { lock.releaseLock(); }
  return true;
}
