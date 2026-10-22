/**
 * Бэкенд LQA-трекера (Google Apps Script, привязан к Google-таблице).
 *
 * Как обновить: Расширения → Apps Script → вставить этот код →
 * «Начать развёртывание» → «Управление развёртываниями» → карандаш →
 * Версия: «Новая версия» → «Развернуть». Ссылка /exec при этом не меняется.
 * Обновлять бэкенд нужно РАНЬШЕ, чем index.html: новый фронт использует action 'patch'.
 */

const SHEET_NAME = 'LQA';
const SHOT_FOLDER_NAME = 'LQA Screenshots';
// Порядок колонок в листе. Новые поля дописываются В КОНЕЦ (после превью),
// чтобы уже заполненные строки не съезжали.
const LAYOUT = ['id','num','createdAt','updatedAt','language','category','subcategory',
                'location','wrong','right','assignee','priority','status','comment','screenshotUrl',
                'screenshotPreview','section','url','managerComment'];
const PREVIEW = 'screenshotPreview';                 // формула IMAGE(), не данные
const COLUMNS = LAYOUT.filter(c => c !== PREVIEW);    // поля бага
function col(name) { return LAYOUT.indexOf(name) + 1; }
// Поля, которые можно менять точечно прямо из таблицы на сайте
const PATCHABLE = ['assignee','status','priority'];

function doGet(e) {
  try {
    const sheet = getSheet();
    const rows = readAllRows(sheet);
    return jsonOut({ ok: true, bugs: rows });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);

    if (body.action === 'screenshot') {
      // загрузка на Диск долгая — не держим блокировку таблицы
      const url = saveScreenshot(body.id, body.dataUrl);
      return jsonOut({ ok: true, url: url });
    }

    // Все записи в таблицу — под блокировкой, чтобы два человека
    // не получили одинаковый номер и не затёрли строки друг друга.
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      const sheet = getSheet();
      if (body.action === 'upsert') return jsonOut(upsertRow(sheet, body.bug, body.baseUpdatedAt));
      if (body.action === 'patch') return jsonOut(patchRow(sheet, body.id, body.fields));
      if (body.action === 'delete') {
        deleteRow(sheet, body.id);
        return jsonOut({ ok: true });
      }
      return jsonOut({ ok: false, error: 'unknown action' });
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(LAYOUT);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(col(PREVIEW), 160);
  } else {
    // миграция: таблица могла быть создана до появления превью, раздела или ссылки — дописываем заголовки
    if (sheet.getMaxColumns() < LAYOUT.length) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), LAYOUT.length - sheet.getMaxColumns());
    }
    const headerRange = sheet.getRange(1, 1, 1, LAYOUT.length);
    const header = headerRange.getValues()[0];
    if (LAYOUT.some((c, i) => header[i] !== c)) {
      headerRange.setValues([LAYOUT]);
      sheet.setColumnWidth(col(PREVIEW), 160);
    }
  }
  return sheet;
}

function rowToObj(row) {
  const obj = {};
  COLUMNS.forEach(c => {
    const v = row[col(c) - 1];
    obj[c] = v instanceof Date ? v.toISOString() : v;
  });
  return obj;
}

function readAllRows(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, LAYOUT.length).getValues();
  return values
    .filter(row => row[0] !== '') // skip empty rows
    .map(rowToObj);
}

function readRow(sheet, rowIdx) {
  return rowToObj(sheet.getRange(rowIdx, 1, 1, LAYOUT.length).getValues()[0]);
}

function findRowIndexById(sheet, id) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  const idx = ids.indexOf(id);
  return idx === -1 ? -1 : idx + 2; // +2: header row + 1-based index
}

function nextNum(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 1;
  const nums = sheet.getRange(2, col('num'), lastRow - 1, 1).getValues().flat();
  return nums.reduce((m, n) => Math.max(m, Number(n) || 0), 0) + 1;
}

// Текст, начинающийся с «=», Sheets выполнит как формулу, а ISO-дату может
// пересчитать в часовой пояс таблицы. Апостроф заставляет хранить значение как текст.
function safeCell(v) {
  if (v === undefined || v === null) return '';
  if (typeof v === 'string' && (/^[=+\-@]/.test(v) || /^\d{4}-\d\d-\d\dT/.test(v))) return "'" + v;
  return v;
}

function sameTime(a, b) {
  const ta = new Date(a).getTime(), tb = new Date(b).getTime();
  if (isNaN(ta) || isNaN(tb)) return String(a || '') === String(b || '');
  return Math.abs(ta - tb) < 1000; // Sheets может хранить время с округлением
}

function upsertRow(sheet, bug, baseUpdatedAt) {
  const now = new Date().toISOString();
  let rowIdx = findRowIndexById(sheet, bug.id);

  if (rowIdx === -1) {
    bug.num = nextNum(sheet);
    bug.createdAt = bug.createdAt || now;
  } else {
    const current = readRow(sheet, rowIdx);
    // Кто-то сохранил этот баг после того, как пользователь открыл форму
    if (baseUpdatedAt !== undefined && !sameTime(current.updatedAt, baseUpdatedAt)) {
      return { ok: false, conflict: true, bug: current, error: 'conflict' };
    }
    bug.num = current.num;
    bug.createdAt = current.createdAt;
    // скриншот не трогаем, только если клиент вообще не прислал поле (пустая строка = «убрать»)
    if (bug.screenshotUrl === undefined) bug.screenshotUrl = current.screenshotUrl;
    if (bug.screenshotUrl !== current.screenshotUrl) trashScreenshot(current.screenshotUrl);
  }
  bug.updatedAt = now;

  const rowValues = LAYOUT.map(c => c === PREVIEW ? '' : safeCell(bug[c])); // формулу превью ставит setPreviewFormula
  if (rowIdx === -1) {
    sheet.appendRow(rowValues);
    rowIdx = sheet.getLastRow();
  } else {
    sheet.getRange(rowIdx, 1, 1, rowValues.length).setValues([rowValues]);
  }
  setPreviewFormula(sheet, rowIdx);
  return { ok: true, bug: readRow(sheet, rowIdx) };
}

// Меняет только переданные поля — не затирает то, что параллельно поправил коллега
function patchRow(sheet, id, fields) {
  const rowIdx = findRowIndexById(sheet, id);
  if (rowIdx === -1) return { ok: false, error: 'not found' };
  Object.keys(fields || {}).forEach(k => {
    if (PATCHABLE.indexOf(k) === -1) return;
    sheet.getRange(rowIdx, col(k)).setValue(safeCell(fields[k]));
  });
  sheet.getRange(rowIdx, col('updatedAt')).setValue(safeCell(new Date().toISOString()));
  return { ok: true, bug: readRow(sheet, rowIdx) };
}

function setPreviewFormula(sheet, rowIdx) {
  const urlCol = columnToLetter(col('screenshotUrl'));
  const previewCol = col(PREVIEW);
  // точка с запятой вместо запятой — таблица в русской локали ждёт именно такой разделитель
  sheet.getRange(rowIdx, previewCol).setFormula(
    '=IF(' + urlCol + rowIdx + '="";"";IMAGE(' + urlCol + rowIdx + '))'
  );
  sheet.setRowHeight(rowIdx, 90); // чтобы миниатюра была видна, а не обрезана
}

function columnToLetter(col) {
  let letter = '';
  while (col > 0) {
    const rem = (col - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}

function deleteRow(sheet, id) {
  const rowIdx = findRowIndexById(sheet, id);
  if (rowIdx === -1) return;
  trashScreenshot(readRow(sheet, rowIdx).screenshotUrl);
  sheet.deleteRow(rowIdx);
}

// Убирает старый скриншот в корзину Диска (восстановим оттуда в течение 30 дней)
function trashScreenshot(url) {
  const match = typeof url === 'string' && url.match(/[?&]id=([^&]+)/);
  if (!match) return;
  try { DriveApp.getFileById(match[1]).setTrashed(true); } catch (e) { /* файла уже нет — не страшно */ }
}

function saveScreenshot(id, dataUrl) {
  const match = String(dataUrl || '').match(/^data:(image\/\w+);base64,(.*)$/);
  if (!match) throw new Error('bad dataUrl');
  if (match[2].length > 8 * 1024 * 1024) throw new Error('screenshot too large');
  const contentType = match[1];
  const bytes = Utilities.base64Decode(match[2]);
  const blob = Utilities.newBlob(bytes, contentType, id + '.jpg');

  const folder = getFolder();
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  // Ссылку в строку пишет upsert: так старый скриншот корректно уходит в корзину
  return 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1000';
}

function getFolder() {
  const folders = DriveApp.getFoldersByName(SHOT_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(SHOT_FOLDER_NAME);
}

/**
 * Разовая починка: конвертирует старые ссылки вида drive.google.com/uc?id=...
 * (которые Sheets IMAGE() не всегда может загрузить) в формат .../thumbnail?id=...&sz=w1000.
 * Запускать вручную из редактора Apps Script один раз, если в колонке
 * screenshotPreview видите #ERROR! у старых записей.
 */
function migrateScreenshotUrls() {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const urlCol = col('screenshotUrl');
  const range = sheet.getRange(2, urlCol, lastRow - 1, 1);
  const values = range.getValues();
  let changed = false;
  for (let i = 0; i < values.length; i++) {
    const url = values[i][0];
    const match = typeof url === 'string' && url.match(/[?&]id=([^&]+)/);
    if (match && url.indexOf('/thumbnail?') === -1) {
      values[i][0] = 'https://drive.google.com/thumbnail?id=' + match[1] + '&sz=w1000';
      changed = true;
    }
  }
  if (!changed) return;
  range.setValues(values);
  for (let i = 0; i < values.length; i++) {
    if (values[i][0]) setPreviewFormula(sheet, i + 2);
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
