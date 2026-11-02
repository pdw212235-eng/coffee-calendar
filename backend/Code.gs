// 커피 달력 백엔드. 구글 시트에 묶인 Apps Script로 배포한다. 절차는 README.md 참고.

const SHEET_NAME = 'records';
const HEADERS = ['id', 'date', 'name', 'cups', 'updatedAt'];
const TEXT_FORMAT = '@';
const LOCK_TIMEOUT_MS = 10000;
const MAX_NAME_LENGTH = 30;
const MAX_CUPS = 999;
const ID_PATTERN = /^[^\s=+\-@]\S{0,63}$/;
const FORMULA_START = /^[=+\-@]/;

// 시트 편집기에서 직접 실행한다. 여러 번 실행해도 기존 행은 그대로다.
function setup() {
  withLock_(function () {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(SHEET_NAME) || spreadsheet.insertSheet(SHEET_NAME);
    // 구글 시트가 "2026-11-02"를 Date로 바꾸지 못하게 값을 쓰기 전에 서식부터 정한다.
    sheet.getRange(1, 1, sheet.getMaxRows(), HEADERS.length).setNumberFormat(TEXT_FORMAT);
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    }
  });
}

function doGet(e) {
  if (!e || !e.parameter || e.parameter.action !== 'list') {
    return fail_('bad_request');
  }
  return json_(listRecords_());
}

function doPost(e) {
  const body = parseBody_(e);
  if (!body) {
    return fail_('bad_request');
  }

  const expected = PropertiesService.getScriptProperties().getProperty('EDIT_PASSWORD');
  if (!expected) {
    return fail_('password_not_set');
  }
  if (body.password !== expected) {
    return fail_('bad_password');
  }

  if (body.action === 'save') {
    const record = parseRecord_(body.record);
    if (!record) {
      return fail_('invalid_record');
    }
    withLock_(function () {
      upsertRecord_(record);
    });
    return json_({ ok: true });
  }

  if (body.action === 'delete') {
    if (typeof body.id !== 'string' || body.id === '') {
      return fail_('bad_request');
    }
    withLock_(function () {
      deleteRecord_(body.id);
    });
    return json_({ ok: true });
  }

  return fail_('bad_request');
}

function listRecords_() {
  const rows = recordsSheet_().getDataRange().getValues().slice(1);
  const records = [];
  let skipped = 0;
  rows.forEach(function (row) {
    // 내용을 통째로 지운 행은 잘못된 행이 아니라 빈 줄이다.
    if (row.every(function (cell) { return cell === ''; })) {
      return;
    }
    const record = parseRecord_(recordFromRow_(row));
    if (record) {
      records.push(record);
    } else {
      skipped += 1;
    }
  });
  records.sort(function (a, b) {
    return a.date < b.date ? -1 : a.date > b.date ? 1 : 0;
  });
  return { ok: true, records: records, skipped: skipped };
}

function upsertRecord_(record) {
  const sheet = recordsSheet_();
  const existing = rowNumbersOf_(sheet, record.id);
  const rowNumber = existing.length > 0 ? existing[0] : nextEmptyRow_(sheet);
  const updatedAt = Utilities.formatDate(new Date(), sheetTimeZone_(), 'yyyy-MM-dd HH:mm:ss');
  const range = sheet.getRange(rowNumber, 1, 1, HEADERS.length);
  // 사람이 서식을 바꾼 행이나 setup 이후 늘어난 행도 문자열로 저장되도록 행마다 서식을 다시 건다.
  range.setNumberFormat(TEXT_FORMAT);
  range.setValues([[record.id, record.date, record.name, record.cups, updatedAt]]);
}

function deleteRecord_(id) {
  const sheet = recordsSheet_();
  // 아래 행부터 지워야 위쪽 행 번호가 밀리지 않는다.
  rowNumbersOf_(sheet, id).reverse().forEach(function (rowNumber) {
    sheet.deleteRow(rowNumber);
  });
}

function rowNumbersOf_(sheet, id) {
  const numbers = [];
  sheet.getDataRange().getValues().forEach(function (row, index) {
    if (index > 0 && String(row[0]) === id) {
      numbers.push(index + 1);
    }
  });
  return numbers;
}

function nextEmptyRow_(sheet) {
  const rowNumber = sheet.getLastRow() + 1;
  if (rowNumber > sheet.getMaxRows()) {
    sheet.insertRowAfter(sheet.getMaxRows());
  }
  return rowNumber;
}

// 탭이 없을 때 빈 목록을 돌려주면 "기록 없음"처럼 보이므로 실패로 드러낸다.
function recordsSheet_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) {
    throw new Error('records 탭이 없습니다. setup을 먼저 실행하세요.');
  }
  return sheet;
}

// 읽기 경계: 사람이 시트를 직접 고쳐 생긴 셀 형태를 레코드 형태로 맞춘다.
function recordFromRow_(row) {
  return {
    id: String(row[0]),
    date: dateFromCell_(row[1]),
    name: String(row[2]),
    cups: cupsFromCell_(row[3]),
  };
}

function dateFromCell_(cell) {
  if (Object.prototype.toString.call(cell) !== '[object Date]') {
    return String(cell);
  }
  if (isNaN(cell.getTime())) {
    return '';
  }
  return Utilities.formatDate(cell, sheetTimeZone_(), 'yyyy-MM-dd');
}

// 스크립트 시간대는 시트와 다를 수 있다. Date 셀은 시트 시간대의 자정이다.
function sheetTimeZone_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
}

function cupsFromCell_(cell) {
  return typeof cell === 'string' && /^\s*\d+\s*$/.test(cell) ? Number(cell) : cell;
}

// 쓰기·읽기 공통 경계 검증. 통과하면 정리된 레코드, 아니면 null.
function parseRecord_(input) {
  if (!input || typeof input !== 'object') {
    return null;
  }
  if (typeof input.id !== 'string' || !ID_PATTERN.test(input.id)) {
    return null;
  }
  if (!isRealDate_(input.date)) {
    return null;
  }
  if (typeof input.name !== 'string') {
    return null;
  }
  const name = input.name.trim();
  if (name.length < 1 || name.length > MAX_NAME_LENGTH || FORMULA_START.test(name)) {
    return null;
  }
  if (!Number.isInteger(input.cups) || input.cups < 1 || input.cups > MAX_CUPS) {
    return null;
  }
  return { id: input.id, date: input.date, name: name, cups: input.cups };
}

function isRealDate_(value) {
  if (typeof value !== 'string') {
    return false;
  }
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value);
  if (!match) {
    return false;
  }
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const date = new Date(Date.UTC(year, month - 1, day));
  return date.getUTCFullYear() === year && date.getUTCMonth() === month - 1 && date.getUTCDate() === day;
}

function parseBody_(e) {
  if (!e || !e.postData || typeof e.postData.contents !== 'string') {
    return null;
  }
  try {
    const body = JSON.parse(e.postData.contents);
    return body && typeof body === 'object' && !Array.isArray(body) ? body : null;
  } catch (err) {
    return null;
  }
}

function withLock_(work) {
  const lock = LockService.getScriptLock();
  lock.waitLock(LOCK_TIMEOUT_MS);
  try {
    work();
    // 잠금을 풀기 전에 반영해야 다음 요청이 이 쓰기를 본다.
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
}

function fail_(error) {
  return json_({ ok: false, error: error });
}

function json_(value) {
  return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);
}
