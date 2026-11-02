// KST 자정 직후 날짜 밀림을 잡으려면 시간대를 고정해야 한다. 어떤 Date도 만들기 전에 정한다.
process.env.TZ = 'Asia/Seoul';

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { dateLabel, monthGrid, nameList, recordsOn, shiftMonth, summarize, toDateKey } from '../logic.js';

const dates = (week) => week.map((cell) => cell.date);

test('2026년 11월 격자: 11/1이 일요일이라 첫 칸이고 5주다', () => {
  const weeks = monthGrid(2026, 11);
  assert.equal(weeks.length, 5);
  assert.ok(weeks.every((week) => week.length === 7));
  assert.deepEqual(weeks[0][0], { date: '2026-11-01', day: 1, inMonth: true });
  assert.deepEqual(dates(weeks[0]).slice(0, 2), ['2026-11-01', '2026-11-02']);
  assert.deepEqual(dates(weeks[4]), [
    '2026-11-29',
    '2026-11-30',
    '2026-12-01',
    '2026-12-02',
    '2026-12-03',
    '2026-12-04',
    '2026-12-05',
  ]);
  assert.deepEqual(
    weeks[4].map((cell) => cell.inMonth),
    [true, true, false, false, false, false, false],
  );
});

test('2028년 2월 윤년: 29일까지 있고 앞뒤 달 칸이 채워진다', () => {
  const cells = monthGrid(2028, 2).flat();
  const inMonth = cells.filter((cell) => cell.inMonth);
  assert.equal(inMonth.length, 29);
  assert.equal(inMonth.at(-1).date, '2028-02-29');
  assert.equal(cells[0].date, '2028-01-30');
  assert.equal(cells.at(-1).date, '2028-03-04');
  assert.equal(cells.length, 35);
});

test('평년 2월은 28일이고, 일요일로 시작하면 딱 4주다', () => {
  const weeks = monthGrid(2026, 2);
  assert.equal(weeks.length, 4);
  assert.equal(weeks[0][0].date, '2026-02-01');
  assert.equal(weeks[3][6].date, '2026-02-28');
});

test('6주가 필요한 달과 연말 경계', () => {
  const august = monthGrid(2026, 8);
  assert.equal(august.length, 6);
  assert.equal(august[0][0].date, '2026-07-26');
  assert.equal(august[5][6].date, '2026-09-05');

  const december = monthGrid(2026, 12);
  assert.equal(december[0][0].date, '2026-11-29');
  assert.equal(december.at(-1)[6].date, '2027-01-02');
});

test('날짜 키는 로컬 구성 요소로 만든다: KST 00:30이 전날로 밀리지 않는다', () => {
  assert.equal(toDateKey(new Date(2026, 10, 2, 0, 30)), '2026-11-02');
  assert.equal(toDateKey(new Date(2026, 10, 2, 23, 59)), '2026-11-02');
  assert.equal(toDateKey(new Date(2027, 0, 1, 0, 0)), '2027-01-01');
});

test('달 이동은 연도를 넘긴다', () => {
  assert.deepEqual(shiftMonth({ year: 2026, month: 12 }, 1), { year: 2027, month: 1 });
  assert.deepEqual(shiftMonth({ year: 2026, month: 1 }, -1), { year: 2025, month: 12 });
  assert.deepEqual(shiftMonth({ year: 2026, month: 11 }, 1), { year: 2026, month: 12 });
});

const records = [
  { id: '1', date: '2026-10-31', name: '가나다', cups: 5 },
  { id: '2', date: '2026-11-01', name: '가나다', cups: 2 },
  { id: '3', date: '2026-11-30', name: '라마바', cups: 4 },
  { id: '4', date: '2026-12-01', name: '라마바', cups: 10 },
  { id: '5', date: '2026-11-02', name: '사아자', cups: 3 },
];

test('합계: 보고 있는 달과 전체를 따로 세고 전체 내림차순이다', () => {
  assert.deepEqual(summarize(records, 2026, 11), [
    { name: '라마바', monthCups: 4, totalCups: 14 },
    { name: '가나다', monthCups: 2, totalCups: 7 },
    { name: '사아자', monthCups: 3, totalCups: 3 },
  ]);
});

test('합계: 그 달 기록이 없는 사람도 0잔으로 남는다', () => {
  const rows = summarize(records, 2026, 10);
  assert.deepEqual(rows.find((row) => row.name === '라마바'), { name: '라마바', monthCups: 0, totalCups: 14 });
});

test('합계: 달 비교는 2026-1과 2026-11을 섞지 않는다', () => {
  const rows = summarize([{ id: '1', date: '2026-11-01', name: '가나다', cups: 1 }], 2026, 1);
  assert.equal(rows[0].monthCups, 0);
});

test('합계: 전체 잔수가 같으면 이름순이다', () => {
  const rows = summarize(
    [
      { id: '1', date: '2026-11-01', name: '라마바', cups: 2 },
      { id: '2', date: '2026-11-02', name: '가나다', cups: 2 },
      { id: '3', date: '2026-11-03', name: '다라마', cups: 2 },
    ],
    2026,
    11,
  );
  assert.deepEqual(
    rows.map((row) => row.name),
    ['가나다', '다라마', '라마바'],
  );
});

test('이름 목록: 중복 없이 이름순이고 코드값 순서가 아니다', () => {
  assert.deepEqual(nameList(records), ['가나다', '라마바', '사아자']);
  assert.deepEqual(
    nameList([
      { id: '1', date: '2026-11-01', name: 'B', cups: 1 },
      { id: '2', date: '2026-11-01', name: 'a', cups: 1 },
    ]),
    ['a', 'B'],
  );
  assert.deepEqual(nameList([]), []);
});

test('특정 날짜의 기록만 고른다', () => {
  assert.deepEqual(
    recordsOn(records, '2026-11-02').map((record) => record.id),
    ['5'],
  );
  assert.deepEqual(recordsOn(records, '2026-11-03'), []);
});

test('날짜 표시는 요일까지 맞다', () => {
  assert.equal(dateLabel('2026-11-02'), '2026년 11월 2일 (월)');
  assert.equal(dateLabel('2026-11-01'), '2026년 11월 1일 (일)');
  assert.equal(dateLabel('2028-02-29'), '2028년 2월 29일 (화)');
});
