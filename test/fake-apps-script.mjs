// 실제 backend/Code.gs를 node:vm에서 실행하는 가짜 Apps Script 환경.
// 시트는 진짜처럼 값을 바꾼다: 일반 서식 칸의 날짜 문자열은 Date, 숫자 문자열은 숫자가 된다.

import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import vm from 'node:vm';

const CODE_PATH = fileURLToPath(new URL('../backend/Code.gs', import.meta.url));
const TIME_ZONE = 'Asia/Seoul';
// 새 Apps Script 프로젝트의 시간대는 시트와 다를 수 있다. 일부러 어긋나게 둬서 잘못된 쪽을 쓰면 날짜가 밀리게 한다.
const SCRIPT_TIME_ZONE = 'America/Los_Angeles';
const UTC_OFFSET = '+09:00';
const TEXT_FORMAT = '@';

function isDate(value) {
  return Object.prototype.toString.call(value) === '[object Date]';
}

function parseSheetDate(text) {
  const match = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(text);
  if (!match) return null;
  const [year, month, day] = match.slice(1).map(Number);
  const utc = new Date(Date.UTC(year, month - 1, day));
  if (utc.getUTCMonth() !== month - 1 || utc.getUTCDate() !== day) return null;
  const iso = `${match[1]}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
  return new Date(`${iso}T00:00:00${UTC_OFFSET}`);
}

function convertOnWrite(value, format) {
  if (value === null || value === undefined || value === '') return '';
  if (format === TEXT_FORMAT) return String(value);
  if (typeof value === 'string') {
    const date = parseSheetDate(value);
    if (date) return date;
    if (value.trim() !== '' && !Number.isNaN(Number(value))) return Number(value);
  }
  return value;
}

function formatDate(date, timeZone, pattern) {
  if (!isDate(date)) throw new Error('formatDate: Date가 아닙니다');
  const parts = {};
  const formatter = new Intl.DateTimeFormat('en-CA', {
    timeZone,
    hourCycle: 'h23',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
  });
  for (const part of formatter.formatToParts(date)) parts[part.type] = part.value;
  const tokens = {
    yyyy: parts.year,
    MM: parts.month,
    dd: parts.day,
    HH: parts.hour,
    mm: parts.minute,
    ss: parts.second,
  };
  return pattern.replace(/yyyy|MM|dd|HH|mm|ss/g, (token) => tokens[token]);
}

class FakeSheet {
  constructor(name, env, maxRows = 1000, maxColumns = 26) {
    this.name = name;
    this.env = env;
    this.maxRows = maxRows;
    this.maxColumns = maxColumns;
    this.values = [];
    this.formats = [];
    for (let row = 0; row < maxRows; row += 1) {
      this.values.push(new Array(maxColumns).fill(''));
      this.formats.push(new Array(maxColumns).fill(''));
    }
  }

  getName() {
    return this.name;
  }

  getMaxRows() {
    return this.maxRows;
  }

  getMaxColumns() {
    return this.maxColumns;
  }

  getLastRow() {
    for (let row = this.maxRows; row >= 1; row -= 1) {
      if (this.values[row - 1].some((cell) => cell !== '')) return row;
    }
    return 0;
  }

  getLastColumn() {
    let last = 0;
    for (const row of this.values) {
      for (let column = row.length; column > last; column -= 1) {
        if (row[column - 1] !== '') {
          last = column;
          break;
        }
      }
    }
    return last;
  }

  getRange(row, column, numRows = 1, numColumns = 1) {
    if (row < 1 || column < 1 || row + numRows - 1 > this.maxRows || column + numColumns - 1 > this.maxColumns) {
      throw new Error('범위의 좌표가 시트 크기를 벗어납니다.');
    }
    return new FakeRange(this, row, column, numRows, numColumns);
  }

  getDataRange() {
    return this.getRange(1, 1, Math.max(this.getLastRow(), 1), Math.max(this.getLastColumn(), 1));
  }

  // 진짜 시트처럼 새 행은 위 행의 서식을 물려받는다.
  insertRowAfter(row) {
    this.env.noteWrite();
    this.values.splice(row, 0, new Array(this.maxColumns).fill(''));
    this.formats.splice(row, 0, [...this.formats[row - 1]]);
    this.maxRows += 1;
  }

  deleteRow(row) {
    this.env.noteWrite();
    if (this.maxRows === 1) throw new Error('시트의 모든 행을 삭제할 수는 없습니다.');
    this.getRange(row, 1);
    this.values.splice(row - 1, 1);
    this.formats.splice(row - 1, 1);
    this.maxRows -= 1;
  }

  // 아래는 테스트 전용: 사람이 시트를 직접 고치는 상황을 흉내 낸다. 잠금 기록에 남지 않는다.
  rawCell(row, column) {
    return this.values[row - 1][column - 1];
  }

  rawRow(row, columns = 5) {
    return this.values[row - 1].slice(0, columns);
  }

  formatAt(row, column) {
    return this.formats[row - 1][column - 1];
  }

  humanType(row, column, value, { clearFormat = false } = {}) {
    if (clearFormat) this.formats[row - 1][column - 1] = '';
    this.values[row - 1][column - 1] = convertOnWrite(value, this.formats[row - 1][column - 1]);
  }

  humanRow(row, values, options) {
    values.forEach((value, index) => this.humanType(row, index + 1, value, options));
  }
}

class FakeRange {
  constructor(sheet, row, column, numRows, numColumns) {
    Object.assign(this, { sheet, row, column, numRows, numColumns });
  }

  getValues() {
    const rows = [];
    for (let r = 0; r < this.numRows; r += 1) {
      rows.push(this.sheet.values[this.row - 1 + r].slice(this.column - 1, this.column - 1 + this.numColumns));
    }
    return rows;
  }

  setValues(values) {
    this.sheet.env.noteWrite();
    if (values.length !== this.numRows || values.some((row) => row.length !== this.numColumns)) {
      throw new Error('데이터 크기가 범위 크기와 다릅니다.');
    }
    values.forEach((row, r) => {
      row.forEach((value, c) => {
        const rowIndex = this.row - 1 + r;
        const columnIndex = this.column - 1 + c;
        this.sheet.values[rowIndex][columnIndex] = convertOnWrite(value, this.sheet.formats[rowIndex][columnIndex]);
      });
    });
    return this;
  }

  setNumberFormat(format) {
    this.sheet.env.noteWrite();
    for (let r = 0; r < this.numRows; r += 1) {
      for (let c = 0; c < this.numColumns; c += 1) {
        this.sheet.formats[this.row - 1 + r][this.column - 1 + c] = format;
      }
    }
    return this;
  }
}

class FakeSpreadsheet {
  constructor(env) {
    this.env = env;
    this.sheets = [new FakeSheet('시트1', env)];
  }

  getSheetByName(name) {
    return this.sheets.find((sheet) => sheet.getName() === name) ?? null;
  }

  insertSheet(name) {
    this.env.noteWrite();
    if (this.getSheetByName(name)) throw new Error(`이름이 "${name}"인 시트가 이미 있습니다.`);
    const sheet = new FakeSheet(name, this.env);
    this.sheets.push(sheet);
    return sheet;
  }
}

// password를 생략하면 EDIT_PASSWORD가 설정되지 않은 상태다.
export function createFakeAppsScript({ password } = {}) {
  const lock = { held: false, waits: [], failNextWait: false };
  const env = {
    unlockedWrites: 0,
    noteWrite() {
      if (!lock.held) this.unlockedWrites += 1;
    },
  };
  const spreadsheet = new FakeSpreadsheet(env);
  spreadsheet.getSpreadsheetTimeZone = () => TIME_ZONE;
  const properties = new Map();
  if (password !== undefined) properties.set('EDIT_PASSWORD', password);

  let clockOffsetMs = 0;
  class FakeDate extends Date {
    static now() {
      return Date.now() + clockOffsetMs;
    }
  }

  const services = {
    Date: FakeDate,
    SpreadsheetApp: {
      getActiveSpreadsheet: () => spreadsheet,
      flush: () => {},
    },
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: (key) => (properties.has(key) ? properties.get(key) : null),
        setProperty(key, value) {
          properties.set(key, String(value));
          return this;
        },
        deleteProperty(key) {
          properties.delete(key);
          return this;
        },
      }),
    },
    LockService: {
      getScriptLock: () => ({
        waitLock(timeoutMs) {
          lock.waits.push(timeoutMs);
          if (lock.failNextWait) {
            lock.failNextWait = false;
            throw new Error('잠금을 기다리다 시간이 초과되었습니다.');
          }
          lock.held = true;
        },
        releaseLock() {
          lock.held = false;
        },
      }),
    },
    ContentService: {
      MimeType: { JSON: 'application/json' },
      createTextOutput(text) {
        return {
          mimeType: 'text/plain',
          setMimeType(mimeType) {
            this.mimeType = mimeType;
            return this;
          },
          getContent: () => text,
          getMimeType() {
            return this.mimeType;
          },
        };
      },
    },
    Utilities: { formatDate },
    Session: { getScriptTimeZone: () => SCRIPT_TIME_ZONE },
  };
  const load = () => {
    const fresh = vm.createContext({ ...services });
    vm.runInContext(readFileSync(CODE_PATH, 'utf8'), fresh, { filename: 'Code.gs' });
    return fresh;
  };
  let context = load();

  return {
    spreadsheet,
    lock,
    env,
    // 실제 Apps Script는 요청마다 새로 실행되어 전역 변수가 남지 않는다.
    restart() {
      context = load();
    },
    advanceClock(ms) {
      clockOffsetMs += ms;
    },
    setup: () => context.setup(),
    doGet: (parameter = {}) => context.doGet({ parameter }),
    doPost: (contents) => context.doPost({ parameter: {}, postData: { contents, type: 'text/plain' } }),
    // 테스트 편의: JSON을 호스트 쪽에서 파싱해 돌려준다.
    list: () => JSON.parse(context.doGet({ parameter: { action: 'list' } }).getContent()),
    post: (body) =>
      JSON.parse(
        context.doPost({ parameter: {}, postData: { contents: JSON.stringify(body), type: 'text/plain' } }).getContent(),
      ),
    sheet: () => spreadsheet.getSheetByName('records'),
  };
}
