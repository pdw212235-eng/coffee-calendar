import assert from 'node:assert/strict';
import { test } from 'node:test';
import { createFakeAppsScript } from './fake-apps-script.mjs';

const PASSWORD = 'test-password';
const ID = 'a0000000-0000-4000-8000-000000000001';
const record = (overrides = {}) => ({ id: ID, date: '2026-11-02', name: '가나다', cups: 2, ...overrides });

function ready() {
  const script = createFakeAppsScript({ password: PASSWORD });
  script.setup();
  return script;
}

const save = (script, rec = record(), password = PASSWORD) => script.post({ password, action: 'save', record: rec });
const dataRowCount = (script) => script.sheet().getLastRow() - 1;

test('가짜 시트 자체 확인: 일반 서식 칸의 날짜 문자열은 Date가 되고 텍스트 칸은 그대로다', () => {
  const script = createFakeAppsScript({ password: PASSWORD });
  const sheet = script.spreadsheet.getSheetByName('시트1');
  sheet.getRange(1, 1, 1, 3).setValues([['2026-11-02', '2026-02-30', '3']]);
  assert.equal(Object.prototype.toString.call(sheet.rawCell(1, 1)), '[object Date]');
  assert.equal(sheet.rawCell(1, 2), '2026-02-30');
  assert.equal(sheet.rawCell(1, 3), 3);

  sheet.getRange(2, 1).setNumberFormat('@').setValues([['2026-11-02']]);
  assert.equal(sheet.rawCell(2, 1), '2026-11-02');
});

test('setup: records 탭과 헤더를 만들고 전체 열을 텍스트 서식으로 둔다', () => {
  const script = ready();
  const sheet = script.sheet();
  assert.deepEqual(sheet.rawRow(1), ['id', 'date', 'name', 'cups', 'updatedAt']);
  for (const column of [1, 2, 3, 4, 5]) {
    assert.equal(sheet.formatAt(1, column), '@');
    assert.equal(sheet.formatAt(sheet.getMaxRows(), column), '@');
  }
});

test('setup을 두 번 실행해도 탭·헤더가 늘지 않고 기록이 남는다', () => {
  const script = ready();
  save(script);
  script.setup();
  assert.equal(script.spreadsheet.sheets.filter((sheet) => sheet.getName() === 'records').length, 1);
  assert.equal(dataRowCount(script), 1);
  assert.deepEqual(script.sheet().rawRow(1), ['id', 'date', 'name', 'cups', 'updatedAt']);
  assert.deepEqual(script.list().records, [record()]);
});

test('저장하면 목록에 그대로 나오고 updatedAt이 시트에 남는다', () => {
  const script = ready();
  assert.deepEqual(save(script), { ok: true });
  assert.deepEqual(script.list(), { ok: true, records: [record()], skipped: 0, note: '' });
  assert.match(script.sheet().rawCell(2, 5), /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/);
});

test('이름은 앞뒤 공백을 지워 저장한다', () => {
  const script = ready();
  save(script, record({ name: '  가나다  ' }));
  assert.equal(script.sheet().rawCell(2, 3), '가나다');
  assert.equal(script.list().records[0].name, '가나다');
});

test('같은 save를 두 번 보내도 행은 하나다', () => {
  const script = ready();
  save(script);
  save(script);
  assert.equal(dataRowCount(script), 1);
  assert.equal(script.list().records.length, 1);
});

test('같은 id로 저장하면 수정이고, 같은 날 다른 id는 별도 기록이다', () => {
  const script = ready();
  const other = 'a0000000-0000-4000-8000-000000000002';
  save(script);
  save(script, record({ id: other, name: '라마바', cups: 1 }));
  save(script, record({ name: '사아자', cups: 9 }));
  assert.equal(dataRowCount(script), 2);
  const records = script.list().records;
  assert.deepEqual(records.find((r) => r.id === ID), record({ name: '사아자', cups: 9 }));
  assert.deepEqual(records.find((r) => r.id === other), record({ id: other, name: '라마바', cups: 1 }));
});

test('삭제하면 해당 행만 없어지고, 없는 id를 지워도 ok다', () => {
  const script = ready();
  const other = 'a0000000-0000-4000-8000-000000000002';
  save(script);
  save(script, record({ id: other }));
  assert.deepEqual(script.post({ password: PASSWORD, action: 'delete', id: ID }), { ok: true });
  assert.deepEqual(script.list().records.map((r) => r.id), [other]);
  assert.deepEqual(script.post({ password: PASSWORD, action: 'delete', id: ID }), { ok: true });
  assert.deepEqual(script.post({ password: PASSWORD, action: 'delete', id: 'no-such-id' }), { ok: true });
  assert.equal(dataRowCount(script), 1);
});

test('사람이 행을 복사해 id가 겹치면 삭제가 겹친 행을 모두 지우고 다른 행은 남긴다', () => {
  const script = ready();
  const keep = record({ id: 'keep-id', name: '라마바' });
  save(script);
  save(script, keep);
  save(script, record({ id: 'second-keep', name: '사아자' }));
  script.sheet().humanRow(5, [ID, '2026-11-02', '가나다', '2', '']);
  script.sheet().humanRow(6, [ID, '2026-11-02', '가나다', '2', '']);
  assert.deepEqual(script.post({ password: PASSWORD, action: 'delete', id: ID }), { ok: true });
  assert.deepEqual(script.list().records.map((r) => r.id), ['keep-id', 'second-keep']);
});

test('틀린 비밀번호는 저장·삭제 모두 거부하고 아무것도 쓰지 않는다', () => {
  const script = ready();
  save(script);
  const denied = { ok: false, error: 'bad_password' };
  assert.deepEqual(save(script, record({ id: 'other-id' }), '틀린값'), denied);
  assert.deepEqual(script.post({ password: '틀린값', action: 'delete', id: ID }), denied);
  assert.deepEqual(script.post({ action: 'delete', id: ID }), denied);
  assert.deepEqual(script.post({ password: '', action: 'delete', id: ID }), denied);
  assert.deepEqual(script.list().records, [record()]);
});

test('비밀번호가 설정되지 않으면 어떤 비밀번호로도 쓸 수 없고 읽기는 된다', () => {
  for (const password of [undefined, '']) {
    const script = createFakeAppsScript({ password });
    script.setup();
    const refused = { ok: false, error: 'password_not_set' };
    assert.deepEqual(save(script, record(), 'anything'), refused);
    assert.deepEqual(save(script, record(), ''), refused);
    assert.deepEqual(script.post({ action: 'save', record: record() }), refused);
    assert.deepEqual(script.post({ password: 'x', action: 'delete', id: ID }), refused);
    assert.deepEqual(script.list(), { ok: true, records: [], skipped: 0, note: '' });
  }
});

const TEN_MINUTES = 10 * 60 * 1000;
const WRONG = { ok: false, error: 'bad_password' };
const LOCKED = { ok: false, error: 'locked' };

// 요청마다 새로 실행되는 실제 환경처럼 매번 restart한다.
function failPassword(script, times) {
  const results = [];
  for (let i = 0; i < times; i += 1) {
    script.restart();
    results.push(save(script, record({ id: `wrong-${i}` }), '틀린값'));
  }
  return results;
}

test('연속 10번 틀리면 잠기고, 잠금 중에는 맞는 비밀번호도 쓸 수 없으며 읽기는 된다', () => {
  const script = ready();
  const results = failPassword(script, 10);
  assert.deepEqual(results.slice(0, 9), Array(9).fill(WRONG));
  assert.deepEqual(results[9], LOCKED);

  script.restart();
  assert.deepEqual(save(script), LOCKED);
  assert.deepEqual(script.post({ password: PASSWORD, action: 'delete', id: ID }), LOCKED);
  assert.deepEqual(script.list(), { ok: true, records: [], skipped: 0, note: '' });
});

test('잠금은 10분 뒤 풀리고, 풀린 뒤에는 다시 10번까지 틀릴 수 있다', () => {
  const script = ready();
  failPassword(script, 10);

  script.advanceClock(TEN_MINUTES - 1000);
  assert.deepEqual(save(script), LOCKED);

  script.advanceClock(1000);
  assert.deepEqual(failPassword(script, 9), Array(9).fill(WRONG));
  assert.deepEqual(save(script), { ok: true });
  assert.deepEqual(script.list().records, [record()]);
});

test('잠금 중 시도는 잠금 시간을 늘리지 않는다', () => {
  const script = ready();
  failPassword(script, 10);
  script.advanceClock(TEN_MINUTES - 1000);
  assert.deepEqual(failPassword(script, 1), [LOCKED]);
  script.advanceClock(1000);
  assert.deepEqual(save(script), { ok: true });
});

test('맞는 비밀번호를 쓰면 틀린 횟수가 처음부터 다시 시작한다', () => {
  const script = ready();
  assert.deepEqual(failPassword(script, 9), Array(9).fill(WRONG));
  assert.deepEqual(save(script), { ok: true });
  assert.deepEqual(failPassword(script, 9), Array(9).fill(WRONG));
  assert.deepEqual(failPassword(script, 1), [LOCKED]);
});

const saveNote = (script, note, password = PASSWORD) => script.post({ password, action: 'saveNote', note });

test('메모를 저장·수정하면 비밀번호 없이 읽히고, 요청이 새로 실행돼도 남는다', () => {
  const script = ready();
  assert.equal(script.list().note, '');
  assert.deepEqual(saveNote(script, '첫 메모'), { ok: true });
  assert.deepEqual(saveNote(script, '고친 메모\n둘째 줄'), { ok: true });
  script.restart();
  assert.equal(script.list().note, '고친 메모\n둘째 줄');
});

test('메모는 앞뒤 공백만 지우고 =, - 로 시작해도 글자 그대로 저장한다', () => {
  const script = ready();
  saveNote(script, '  =SUM(A1)\n- 항목  ');
  assert.equal(script.list().note, '=SUM(A1)\n- 항목');
});

test('메모는 300자까지 저장하고 넘으면 거부하며 기존 내용을 유지한다', () => {
  const script = ready();
  const full = '가'.repeat(300);
  assert.deepEqual(saveNote(script, full), { ok: true });
  const denied = { ok: false, error: 'invalid_note' };
  assert.deepEqual(saveNote(script, '가'.repeat(301)), denied);
  assert.deepEqual(saveNote(script, 123), denied);
  assert.deepEqual(script.post({ password: PASSWORD, action: 'saveNote' }), denied);
  assert.equal(script.list().note, full);
});

test('빈 내용으로 저장하면 메모가 지워진다', () => {
  const script = ready();
  saveNote(script, '지울 메모');
  assert.deepEqual(saveNote(script, '   '), { ok: true });
  assert.equal(script.list().note, '');
});

test('메모 저장도 비밀번호와 잠금 규칙을 똑같이 따른다', () => {
  const script = ready();
  saveNote(script, '원래 메모');
  assert.deepEqual(saveNote(script, '바꾼 메모', '틀린값'), WRONG);

  const unset = createFakeAppsScript({});
  unset.setup();
  assert.deepEqual(saveNote(unset, '메모', 'anything'), { ok: false, error: 'password_not_set' });

  failPassword(script, 10);
  assert.deepEqual(saveNote(script, '바꾼 메모'), LOCKED);
  assert.equal(script.list().note, '원래 메모');
});

test('메모 저장은 기록을 건드리지 않는다', () => {
  const script = ready();
  save(script);
  saveNote(script, '메모');
  assert.deepEqual(script.list().records, [record()]);
  assert.equal(dataRowCount(script), 1);
});

test('잘못된 날짜를 거부한다', () => {
  const script = ready();
  for (const date of ['2026-02-30', '2026-13-01', '2026-00-10', '2027-02-29', '2026-2-3', '20261102', '', null, 20261102]) {
    assert.deepEqual(save(script, record({ date })), { ok: false, error: 'invalid_record' }, `날짜 ${date}`);
  }
  assert.equal(dataRowCount(script), 0);
  assert.deepEqual(save(script, record({ date: '2028-02-29' })), { ok: true });
});

test('잘못된 이름을 거부한다: 빈값, 31자, 수식 시작 문자', () => {
  const script = ready();
  const names = ['', '   ', '가'.repeat(31), '=SUM(A1)', '+1', '-1', '@가나다', '  =SUM(A1)', null, 3];
  for (const name of names) {
    assert.deepEqual(save(script, record({ name })), { ok: false, error: 'invalid_record' }, `이름 ${name}`);
  }
  assert.equal(dataRowCount(script), 0);
  assert.deepEqual(save(script, record({ name: '가'.repeat(30) })), { ok: true });
  assert.deepEqual(save(script, record({ name: `  ${'가'.repeat(30)}  ` })), { ok: true });
});

test('잘못된 잔수를 거부하고 경계값 1과 999는 받는다', () => {
  const script = ready();
  for (const cups of [0, 1000, -1, 1.5, '3', null, undefined]) {
    assert.deepEqual(save(script, record({ cups })), { ok: false, error: 'invalid_record' }, `잔수 ${cups}`);
  }
  assert.equal(dataRowCount(script), 0);
  assert.deepEqual(save(script, record({ cups: 1 })), { ok: true });
  assert.deepEqual(save(script, record({ cups: 999 })), { ok: true });
  assert.equal(script.list().records[0].cups, 999);
});

test('잘못된 id와 빠진 record를 거부한다', () => {
  const script = ready();
  for (const id of ['', '=1+1', 'a b', null, 3, 'x'.repeat(65)]) {
    assert.deepEqual(save(script, record({ id })), { ok: false, error: 'invalid_record' }, `id ${id}`);
  }
  assert.deepEqual(script.post({ password: PASSWORD, action: 'save' }), { ok: false, error: 'invalid_record' });
  assert.equal(dataRowCount(script), 0);
});

test('형식이 틀린 요청은 bad_request다', () => {
  const script = ready();
  const badRequest = { ok: false, error: 'bad_request' };
  const parse = (output) => JSON.parse(output.getContent());
  assert.deepEqual(parse(script.doPost('이건 JSON이 아니다')), badRequest);
  assert.deepEqual(parse(script.doPost('[1,2]')), badRequest);
  assert.deepEqual(parse(script.doPost('null')), badRequest);
  assert.deepEqual(script.post({ password: PASSWORD, action: 'rename' }), badRequest);
  assert.deepEqual(script.post({ password: PASSWORD }), badRequest);
  assert.deepEqual(script.post({ password: PASSWORD, action: 'delete' }), badRequest);
  assert.deepEqual(parse(script.doGet({})), badRequest);
  assert.deepEqual(parse(script.doGet({ action: 'drop' })), badRequest);
});

test('응답은 JSON MIME이다', () => {
  const script = ready();
  assert.equal(script.doGet({ action: 'list' }).getMimeType(), 'application/json');
});

test('날짜는 Date로 바뀌지 않고 문자열로 왕복한다', () => {
  const script = ready();
  save(script);
  assert.equal(script.sheet().rawCell(2, 2), '2026-11-02');
  assert.equal(script.list().records[0].date, '2026-11-02');
});

test('setup 때의 범위를 넘어 시트 행이 늘어나도 날짜가 문자열로 저장된다', () => {
  const script = ready();
  const sheet = script.sheet();
  const before = sheet.getMaxRows();
  for (let row = 2; row <= before; row += 1) {
    sheet.humanRow(row, [`old-${row}`, '2026-11-01', '가나다', '1', '']);
  }
  assert.deepEqual(save(script), { ok: true });
  assert.equal(sheet.getMaxRows(), before + 1);
  assert.equal(sheet.rawCell(before + 1, 2), '2026-11-02');
});

test('시트에서 직접 고쳐 Date가 된 셀은 시트 시간대 기준 날짜로 읽는다', () => {
  const script = ready();
  save(script);
  script.sheet().humanType(2, 2, '2026-11-05', { clearFormat: true });
  assert.equal(Object.prototype.toString.call(script.sheet().rawCell(2, 2)), '[object Date]');
  assert.deepEqual(script.list(), { ok: true, records: [record({ date: '2026-11-05' })], skipped: 0, note: '' });
});

test('Date가 된 행을 화면에서 수정하면 다시 문자열로 저장된다', () => {
  const script = ready();
  save(script);
  script.sheet().humanType(2, 2, '2026-11-05', { clearFormat: true });
  save(script, record({ date: '2026-11-05', cups: 7 }));
  assert.equal(script.sheet().rawCell(2, 2), '2026-11-05');
  assert.equal(script.list().records[0].cups, 7);
});

test('문자열 잔수는 숫자로 읽는다', () => {
  const script = ready();
  script.sheet().humanRow(2, [ID, '2026-11-02', '가나다', ' 3 ', '']);
  assert.equal(script.sheet().rawCell(2, 4), ' 3 ');
  assert.deepEqual(script.list().records, [record({ cups: 3 })]);
});

test('형식에 안 맞는 행은 건너뛴 수로 돌려주고 나머지는 그대로 읽는다', () => {
  const script = ready();
  save(script);
  const badRows = [
    ['bad-date', '2026-02-30', '가나다', '1', ''],
    ['bad-text', '내일', '가나다', '1', ''],
    ['bad-cups', '2026-11-03', '가나다', '많이', ''],
    ['cups-zero', '2026-11-03', '가나다', '0', ''],
    ['cups-float', '2026-11-03', '가나다', '1.5', ''],
    ['bad-name', '2026-11-03', '=SUM(A1)', '1', ''],
    ['no-name', '2026-11-03', '', '1', ''],
    ['', '2026-11-03', '가나다', '1', ''],
  ];
  badRows.forEach((row, index) => script.sheet().humanRow(3 + index, row));
  const result = script.list();
  assert.equal(result.skipped, badRows.length);
  assert.deepEqual(result.records, [record()]);
});

test('통째로 비운 행은 건너뛴 행으로 세지 않는다', () => {
  const script = ready();
  save(script);
  save(script, record({ id: 'second-id' }));
  script.sheet().humanRow(2, ['', '', '', '', '']);
  assert.deepEqual(script.list(), { ok: true, records: [record({ id: 'second-id' })], skipped: 0, note: '' });
});

test('목록은 날짜 오름차순이다', () => {
  const script = ready();
  save(script, record({ id: 'c', date: '2026-12-01' }));
  save(script, record({ id: 'a', date: '2026-01-15' }));
  save(script, record({ id: 'b', date: '2026-11-02' }));
  assert.deepEqual(script.list().records.map((r) => r.date), ['2026-01-15', '2026-11-02', '2026-12-01']);
});

test('records 탭이 없으면 빈 목록이 아니라 실패한다', () => {
  const script = createFakeAppsScript({ password: PASSWORD });
  assert.throws(() => script.list(), /records 탭이 없습니다/);
  assert.throws(() => save(script), /records 탭이 없습니다/);
});

test('쓰기는 10초 잠금 안에서 일어나고 끝나면 풀린다', () => {
  const script = ready();
  const setupWaits = script.lock.waits.length;
  save(script);
  script.post({ password: PASSWORD, action: 'delete', id: ID });
  const waits = script.lock.waits.slice(setupWaits);
  assert.ok(waits.length >= 2);
  assert.ok(waits.every((timeoutMs) => timeoutMs === 10000));
  assert.equal(script.env.unlockedWrites, 0);
  assert.equal(script.lock.held, false);
});

test('잠금을 못 얻으면 쓰지 않는다', () => {
  const script = ready();
  script.lock.failNextWait = true;
  assert.throws(() => save(script), /시간이 초과/);
  assert.equal(dataRowCount(script), 0);
  assert.equal(script.lock.held, false);
});

test('쓰기 중 오류가 나도 잠금이 풀린다', () => {
  const script = createFakeAppsScript({ password: PASSWORD });
  assert.throws(() => save(script));
  assert.equal(script.lock.held, false);
});
