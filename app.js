import { apiUrl } from './config.js';
import { dateLabel, monthGrid, nameList, recordsOn, shiftMonth, summarize, toDateKey } from './logic.js';

const PASSWORD_KEY = 'coffee-calendar.password';

const WRITE_ERRORS = {
  password_not_set: '서버에 비밀번호가 아직 설정되지 않았습니다. README.md의 설정 절차를 확인해 주세요.',
  invalid_record: '입력값을 확인해 주세요. 이름은 1~30자이고 =, +, -, @로 시작할 수 없으며, 잔수는 1~999입니다.',
  invalid_note: '메모는 300자까지 저장할 수 있습니다.',
  bad_request: '요청 형식이 잘못되었습니다.',
  locked: '비밀번호를 여러 번 틀려 저장이 잠겼습니다. 최대 10분 뒤 다시 시도해 주세요.',
};
const NETWORK_ERROR = '서버 응답을 받지 못했습니다. 반영되지 않았을 수 있으니 다시 시도해 주세요.';

const $ = (id) => document.getElementById(id);

function currentMonth() {
  const now = new Date();
  return { year: now.getFullYear(), month: now.getMonth() + 1 };
}

const state = {
  status: 'loading', // loading | ready | error | unconfigured
  records: [],
  skipped: 0,
  note: '',
  noteDirty: false,
  view: currentMonth(),
  // 대화상자가 열려 있을 때만 값이 있다: { date, newId, editingId, confirmingDeleteId, busy }
  dialog: null,
};

// ---- 비밀번호 저장소: 저장소가 막힌 브라우저에서는 매번 묻는 것으로 끝난다 ----

function storedPassword() {
  try {
    return localStorage.getItem(PASSWORD_KEY);
  } catch {
    return null;
  }
}

function storePassword(password) {
  try {
    localStorage.setItem(PASSWORD_KEY, password);
  } catch {
    // 저장 못 해도 이번 요청은 이미 성공했다.
  }
}

function forgetPassword() {
  try {
    localStorage.removeItem(PASSWORD_KEY);
  } catch {
    // 지울 값을 읽을 수도 없는 상태다.
  }
}

function askPassword(message) {
  const dialog = $('password-dialog');
  $('password-message').textContent = message;
  $('input-password').value = '';
  dialog.returnValue = '';
  return new Promise((resolve) => {
    dialog.addEventListener(
      'close',
      () => resolve(dialog.returnValue === 'ok' ? $('input-password').value : null),
      { once: true },
    );
    dialog.showModal();
  });
}

// ---- 서버 통신 ----

async function fetchRecords() {
  const response = await fetch(`${apiUrl}?action=list`, { cache: 'no-store' });
  if (!response.ok) {
    throw new Error(`HTTP ${response.status}`);
  }
  const data = await response.json();
  if (!data.ok || !Array.isArray(data.records)) {
    throw new Error(data.error ?? '응답 형식 오류');
  }
  return data;
}

async function postJson(body) {
  // text/plain이어야 Apps Script가 처리 못 하는 CORS preflight가 생기지 않는다.
  const response = await fetch(apiUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify(body),
  });
  if (!response.ok) {
    throw new Error(`HTTP ${response.status}`);
  }
  return response.json();
}

// 비밀번호를 붙여 보낸다. 입력을 취소하면 false, 서버가 거부하면 코드를 담아 throw.
async function sendWrite(payload) {
  let message = '기록을 고치려면 비밀번호를 입력해 주세요.';
  for (;;) {
    const password = storedPassword() ?? (await askPassword(message));
    if (password === null) {
      return false;
    }
    const result = await postJson({ ...payload, password });
    if (result.ok) {
      storePassword(password);
      return true;
    }
    if (result.error !== 'bad_password') {
      throw new Error(WRITE_ERRORS[result.error] ?? NETWORK_ERROR, { cause: 'server' });
    }
    forgetPassword();
    message = '비밀번호가 틀렸습니다. 다시 입력해 주세요.';
  }
}

async function loadRecords() {
  if (state.status !== 'ready') {
    state.status = 'loading';
    render();
  }
  try {
    const data = await fetchRecords();
    state.records = data.records;
    state.skipped = Number(data.skipped) || 0;
    state.note = data.note ?? '';
    state.status = 'ready';
    $('status-error-saved').hidden = true;
  } catch {
    // 옛 데이터나 빈 달력을 남기지 않는다. 실패 안내와 재시도 버튼만 보인다.
    state.records = [];
    state.skipped = 0;
    state.status = 'error';
    if (state.dialog) {
      $('record-dialog').close();
    }
  }
  render();
}

// ---- 화면 그리기: 사용자 입력은 textContent로만 넣는다 ----

function element(tag, className, text) {
  const node = document.createElement(tag);
  if (className) node.className = className;
  if (text !== undefined) node.textContent = text;
  return node;
}

function render() {
  $('status-loading').hidden = state.status !== 'loading';
  $('status-unconfigured').hidden = state.status !== 'unconfigured';
  $('status-error').hidden = state.status !== 'error';
  $('content').hidden = state.status !== 'ready';
  if (state.status !== 'ready') {
    return;
  }

  const warning = $('skipped-warning');
  warning.hidden = state.skipped === 0;
  warning.textContent = `형식이 맞지 않아 건너뛴 시트 행 ${state.skipped}개: 달력과 합계에 포함되지 않았습니다. 시트를 확인해 주세요.`;

  renderCalendar();
  renderNote();
  renderSummary();
  renderNameOptions();
  if (state.dialog) {
    renderDayRecords();
  }
}

function renderCalendar() {
  const { year, month } = state.view;
  $('month-title').textContent = `${year}년 ${month}월`;

  const today = toDateKey(new Date());
  const cells = monthGrid(year, month).flat().map((cell) => {
    const entries = recordsOn(state.records, cell.date);
    const button = element('button', 'day');
    button.type = 'button';
    button.dataset.date = cell.date;
    button.classList.toggle('day-outside', !cell.inMonth);
    if (cell.date === today) {
      button.classList.add('day-today');
      button.setAttribute('aria-current', 'date');
    }
    const spoken = entries.map((entry) => `${entry.name} ${entry.cups}잔`).join(', ') || '기록 없음';
    button.setAttribute('aria-label', `${dateLabel(cell.date)}, ${spoken}`);

    button.append(element('span', 'day-number', String(cell.day)));
    for (const entry of entries) {
      const line = element('span', 'entry');
      line.append(element('span', 'entry-name', entry.name), ' ', element('span', 'entry-cups', `${entry.cups}잔`));
      button.append(line);
    }
    button.addEventListener('click', () => openDialog(cell.date));
    return button;
  });
  $('calendar').replaceChildren(...cells);
}

function renderNote() {
  // 기록을 저장해 목록을 다시 읽어도 쓰고 있던 메모를 덮어쓰면 안 된다.
  if (!state.noteDirty) {
    $('note-input').value = state.note;
  }
  $('note-count').textContent = `${$('note-input').value.length}/300`;
}

function renderSummary() {
  const { year, month } = state.view;
  const rows = summarize(state.records, year, month).map((row) => {
    const tr = element('tr');
    tr.append(
      element('td', '', row.name),
      element('td', 'number', String(row.monthCups)),
      element('td', 'number', String(row.totalCups)),
    );
    return tr;
  });
  $('summary-month').textContent = `${month}월`;
  $('summary-body').replaceChildren(...rows);
  $('summary-table').hidden = rows.length === 0;
  $('summary-empty').hidden = rows.length > 0;
}

function renderNameOptions() {
  const options = nameList(state.records).map((name) => {
    const option = element('option');
    option.value = name;
    return option;
  });
  $('name-options').replaceChildren(...options);
}

function renderDayRecords() {
  const dialog = state.dialog;
  const entries = recordsOn(state.records, dialog.date);
  const items = entries.map((entry) => {
    const item = element('li', 'day-record');
    const buttons = element('div', 'day-record-buttons');
    if (dialog.confirmingDeleteId === entry.id) {
      item.append(element('span', 'day-record-text', `${entry.name} ${entry.cups}잔을 삭제할까요?`));
      buttons.append(
        actionButton('삭제 확인', 'button button-primary', () => deleteRecord(entry.id)),
        actionButton('취소', 'button', () => setConfirmingDelete(null)),
      );
    } else {
      item.append(element('span', 'day-record-text', `${entry.name} ${entry.cups}잔`));
      buttons.append(
        actionButton('수정', 'button', () => startEdit(entry)),
        actionButton('삭제', 'button', () => setConfirmingDelete(entry.id)),
      );
    }
    item.append(buttons);
    return item;
  });
  $('day-records').replaceChildren(...items);
  $('day-empty').hidden = items.length > 0;
}

function actionButton(label, className, onClick) {
  const button = element('button', className, label);
  button.type = 'button';
  button.addEventListener('click', onClick);
  return button;
}

// ---- 대화상자 ----

function openDialog(date) {
  state.dialog = {
    date,
    // 폼을 열 때 정해 두어, 응답을 못 받고 다시 눌러도 같은 행에 저장된다.
    newId: crypto.randomUUID(),
    editingId: null,
    confirmingDeleteId: null,
    busy: false,
  };
  $('dialog-title').textContent = dateLabel(date);
  showError('');
  fillForm(null);
  renderDayRecords();
  $('record-dialog').showModal();
}

function fillForm(entry) {
  state.dialog.editingId = entry ? entry.id : null;
  $('form-title').textContent = entry ? '기록 수정' : '새 기록';
  $('input-name').value = entry ? entry.name : '';
  $('input-cups').value = String(entry ? entry.cups : 1);
  $('form-cancel-edit').hidden = !entry;
}

function startEdit(entry) {
  state.dialog.confirmingDeleteId = null;
  showError('');
  fillForm(entry);
  renderDayRecords();
  $('input-name').focus();
}

function setConfirmingDelete(id) {
  state.dialog.confirmingDeleteId = id;
  showError('');
  renderDayRecords();
}

function showError(message) {
  $('dialog-error').textContent = message;
  $('dialog-error').hidden = message === '';
}

function setBusy(busy, saving) {
  // 목록 재조회 실패로 대화상자가 먼저 닫혀도 다음에 열 때 잠겨 있으면 안 된다.
  if (state.dialog) {
    state.dialog.busy = busy;
  }
  $('dialog-fields').disabled = busy;
  $('form-save').textContent = busy && saving ? '저장 중…' : '저장';
}

// 쓰기 → 서버 응답 → 목록 다시 읽기. 끝까지 성공하면 true.
async function runWrite(payload) {
  setBusy(true, payload.action === 'save');
  showError('');
  try {
    if (!(await sendWrite(payload))) {
      return false;
    }
    await reloadAfterWrite();
    return true;
  } catch (error) {
    showError(error.cause === 'server' ? error.message : NETWORK_ERROR);
    return false;
  } finally {
    setBusy(false);
  }
}

async function reloadAfterWrite() {
  await loadRecords();
  // 쓰기는 끝났는데 재조회만 실패한 경우다. 같은 내용을 또 입력하면 중복 기록이 된다.
  $('status-error-saved').hidden = state.status !== 'error';
}

async function saveNote(event) {
  event.preventDefault();
  $('note-fields').disabled = true;
  $('note-save').textContent = '저장 중…';
  $('note-error').hidden = true;
  $('note-status').textContent = '';
  try {
    if (await sendWrite({ action: 'saveNote', note: $('note-input').value })) {
      state.noteDirty = false;
      await reloadAfterWrite();
      $('note-status').textContent = '저장했습니다.';
    }
  } catch (error) {
    $('note-error').textContent = error.cause === 'server' ? error.message : NETWORK_ERROR;
    $('note-error').hidden = false;
  } finally {
    $('note-fields').disabled = false;
    $('note-save').textContent = '저장';
  }
}

function onNoteInput() {
  state.noteDirty = $('note-input').value !== state.note;
  $('note-status').textContent = '';
  renderNote();
}

async function saveRecord(event) {
  event.preventDefault();
  const dialog = state.dialog;
  if (!dialog || dialog.busy) {
    return;
  }
  const saved = await runWrite({
    action: 'save',
    record: {
      id: dialog.editingId ?? dialog.newId,
      date: dialog.date,
      name: $('input-name').value.trim(),
      cups: Number($('input-cups').value),
    },
  });
  if (saved && state.dialog === dialog) {
    $('record-dialog').close();
  }
}

async function deleteRecord(id) {
  const dialog = state.dialog;
  if (!dialog || dialog.busy) {
    return;
  }
  const deleted = await runWrite({ action: 'delete', id });
  if (deleted && state.dialog === dialog) {
    dialog.confirmingDeleteId = null;
    if (dialog.editingId === id) {
      fillForm(null);
    }
    renderDayRecords();
  }
}

function onDialogClosed() {
  // close 이벤트는 비동기다. 그사이 다시 열렸다면 새 대화상자의 상태를 지우면 안 된다.
  if ($('record-dialog').open) {
    return;
  }
  const date = state.dialog?.date;
  state.dialog = null;
  // 달력을 다시 그리면 원래 버튼이 사라져 포커스가 끊기므로 같은 날짜로 돌려놓는다.
  document.querySelector(`[data-date="${date}"]`)?.focus();
}

// ---- 시작 ----

function moveView(view) {
  state.view = view;
  render();
}

function clearSavedPassword() {
  const hadPassword = storedPassword() !== null;
  forgetPassword();
  $('password-status').textContent = hadPassword
    ? '이 기기에 저장된 비밀번호를 지웠습니다.'
    : '이 기기에 저장된 비밀번호가 없습니다.';
}

$('retry').addEventListener('click', loadRecords);
$('prev-month').addEventListener('click', () => moveView(shiftMonth(state.view, -1)));
$('next-month').addEventListener('click', () => moveView(shiftMonth(state.view, 1)));
$('this-month').addEventListener('click', () => moveView(currentMonth()));
$('clear-password').addEventListener('click', clearSavedPassword);
$('record-form').addEventListener('submit', saveRecord);
$('note-form').addEventListener('submit', saveNote);
$('note-input').addEventListener('input', onNoteInput);
$('form-cancel-edit').addEventListener('click', () => {
  showError('');
  fillForm(null);
});
$('dialog-close').addEventListener('click', () => $('record-dialog').close());
$('password-cancel').addEventListener('click', () => $('password-dialog').close());
$('record-dialog').addEventListener('close', onDialogClosed);
$('record-dialog').addEventListener('cancel', (event) => {
  if (state.dialog?.busy) {
    event.preventDefault();
  }
});

if (apiUrl) {
  loadRecords();
} else {
  state.status = 'unconfigured';
  render();
}
