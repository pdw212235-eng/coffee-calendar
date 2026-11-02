// 화면과 무관한 순수 함수. 브라우저와 Node에서 같이 쓴다.
// 날짜는 로컬 시간 구성 요소로만 다룬다(toISOString은 KST 자정 근처에 하루 밀린다).

const WEEKDAYS = ['일', '월', '화', '수', '목', '금', '토'];

function pad(number) {
  return String(number).padStart(2, '0');
}

export function toDateKey(date) {
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;
}

// month는 1~12. 일요일 시작 주 단위 배열을 돌려준다.
export function monthGrid(year, month) {
  const first = new Date(year, month - 1, 1);
  const last = new Date(year, month, 0);
  const cellCount = Math.ceil((first.getDay() + last.getDate()) / 7) * 7;

  const weeks = [];
  for (let index = 0; index < cellCount; index += 1) {
    const date = new Date(year, month - 1, 1 - first.getDay() + index);
    if (index % 7 === 0) {
      weeks.push([]);
    }
    weeks[weeks.length - 1].push({
      date: toDateKey(date),
      day: date.getDate(),
      inMonth: date.getMonth() === month - 1,
    });
  }
  return weeks;
}

export function shiftMonth({ year, month }, delta) {
  const date = new Date(year, month - 1 + delta, 1);
  return { year: date.getFullYear(), month: date.getMonth() + 1 };
}

// 전체 잔수 내림차순, 같으면 이름순.
export function summarize(records, year, month) {
  const monthPrefix = `${year}-${pad(month)}-`;
  const byName = new Map();
  for (const record of records) {
    const row = byName.get(record.name) ?? { name: record.name, monthCups: 0, totalCups: 0 };
    row.totalCups += record.cups;
    if (record.date.startsWith(monthPrefix)) {
      row.monthCups += record.cups;
    }
    byName.set(record.name, row);
  }
  return [...byName.values()].sort(
    (a, b) => b.totalCups - a.totalCups || a.name.localeCompare(b.name, 'ko'),
  );
}

export function nameList(records) {
  return [...new Set(records.map((record) => record.name))].sort((a, b) => a.localeCompare(b, 'ko'));
}

export function recordsOn(records, dateKey) {
  return records.filter((record) => record.date === dateKey);
}

export function dateLabel(dateKey) {
  const [year, month, day] = dateKey.split('-').map(Number);
  const weekday = WEEKDAYS[new Date(year, month - 1, day).getDay()];
  return `${year}년 ${month}월 ${day}일 (${weekday})`;
}
