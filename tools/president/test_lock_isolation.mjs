import assert from 'node:assert/strict';
import fs from 'node:fs';
import test from 'node:test';
import vm from 'node:vm';

const root = new URL('../../', import.meta.url);
const gasSource = fs.readFileSync(new URL('gas.js', root), 'utf8');

const PRES_HEADERS = [
  '登録日時', 'タイトル', '開始日', '開始時刻', '終了日', '終了時刻',
  '場所', 'メモ', 'カテゴリ', '色', 'ID', '更新者',
];

const EXISTING_ROW = [
  '2026-08-12 09:00:00', '既存予定', '2026-08-12', '10:00',
  '2026-08-12', '11:00', '会議室', '確認用', '会議', '#1D9E75',
  'P_EXISTING', 'Ryo',
];

const SECOND_ROW = [
  '2026-08-12 12:00:00', '後続予定', '2026-08-13', '13:00',
  '2026-08-13', '14:00', '応接室', '行ずれ確認用', '会議', '#36c',
  'P_SECOND', 'Ryo',
];

function cloneRows(rows) {
  return rows.map(row => [...row]);
}

class FakeRange {
  constructor(sheet, row, column, numRows = 1, numColumns = 1) {
    this.sheet = sheet;
    this.row = row;
    this.column = column;
    this.numRows = numRows;
    this.numColumns = numColumns;
  }

  getValues() {
    const values = [];
    for (let r = 0; r < this.numRows; r++) {
      const source = this.sheet.rows[this.row - 1 + r] || [];
      const row = [];
      for (let c = 0; c < this.numColumns; c++) {
        row.push(source[this.column - 1 + c] ?? '');
      }
      values.push(row);
    }
    return values;
  }

  setValue(value) {
    return this.setValues([[value]]);
  }

  setValues(values) {
    for (let r = 0; r < this.numRows; r++) {
      while (this.sheet.rows.length < this.row + r) this.sheet.rows.push([]);
      const target = this.sheet.rows[this.row - 1 + r];
      for (let c = 0; c < this.numColumns; c++) {
        target[this.column - 1 + c] = values[r][c];
      }
    }
    return this;
  }

  clearContent() {
    for (let r = 0; r < this.numRows; r++) {
      while (this.sheet.rows.length < this.row + r) this.sheet.rows.push([]);
      const target = this.sheet.rows[this.row - 1 + r];
      for (let c = 0; c < this.numColumns; c++) {
        target[this.column - 1 + c] = '';
      }
    }
    return this;
  }
}

class FakeSheet {
  constructor(name, rows, metrics) {
    this.name = name;
    this.rows = cloneRows(rows);
    this.metrics = metrics;
    this.hidden = false;
  }

  getDataRange() {
    if (this.name === '日報データ') this.metrics.dailyDataReads++;
    const width = Math.max(1, ...this.rows.map(row => row.length));
    return new FakeRange(this, 1, 1, Math.max(1, this.rows.length), width);
  }

  getMaxColumns() {
    return Math.max(1, ...this.rows.map(row => row.length));
  }

  insertColumnsAfter(_after, howMany) {
    this.rows.forEach(row => {
      while (howMany-- > 0) row.push('');
    });
  }

  getRange(row, column, numRows = 1, numColumns = 1) {
    return new FakeRange(this, row, column, numRows, numColumns);
  }

  appendRow(row) {
    this.rows.push([...row]);
    return this;
  }

  deleteRow(row) {
    this.rows.splice(row - 1, 1);
  }

  hideSheet() {
    this.hidden = true;
  }
}

function createLock(metrics, prefix, available) {
  return {
    tryLock() {
      metrics[`${prefix}Try`]++;
      return available;
    },
    releaseLock() {
      metrics[`${prefix}Release`]++;
    },
  };
}

function loadGas({
  scriptLockAvailable = true,
  userLockAvailable = true,
  presidentSheetExists = true,
} = {}) {
  const metrics = {
    scriptTry: 0,
    scriptRelease: 0,
    userTry: 0,
    userRelease: 0,
    dailyDataReads: 0,
    insertedSheets: [],
  };

  const dailySheet = new FakeSheet('日報データ', [
    ['登録日時', '作業日'],
  ], metrics);
  const sheets = new Map([['日報データ', dailySheet]]);
  if (presidentSheetExists) {
    sheets.set('社長予定', new FakeSheet(
      '社長予定',
      [PRES_HEADERS, EXISTING_ROW, SECOND_ROW],
      metrics,
    ));
  }

  const spreadsheet = {
    getSheetByName(name) {
      return sheets.get(name) || null;
    },
    insertSheet(name) {
      metrics.insertedSheets.push(name);
      const sheet = new FakeSheet(name, [], metrics);
      sheets.set(name, sheet);
      return sheet;
    },
  };

  const scriptLock = createLock(metrics, 'script', scriptLockAvailable);
  const userLock = createLock(metrics, 'user', userLockAvailable);
  const scriptProperties = new Map([['CAL_REQUIRE_TOKEN', '0']]);

  const ContentService = {
    MimeType: { JSON: 'application/json' },
    createTextOutput(text) {
      return {
        text,
        mimeType: '',
        setMimeType(mimeType) {
          this.mimeType = mimeType;
          return this;
        },
      };
    },
  };

  const context = vm.createContext({
    console,
    ContentService,
    LockService: {
      getScriptLock: () => scriptLock,
      getUserLock: () => userLock,
      getDocumentLock: () => null,
    },
    SpreadsheetApp: { getActiveSpreadsheet: () => spreadsheet },
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: key => scriptProperties.get(key) || null,
        setProperty: (key, value) => scriptProperties.set(key, String(value)),
      }),
    },
    Session: { getScriptTimeZone: () => 'Asia/Tokyo' },
    Utilities: {
      formatDate(value, _timezone, format) {
        const iso = new Date(value).toISOString();
        return format === 'HH:mm' ? iso.slice(11, 16) : iso.slice(0, 10);
      },
      getUuid: () => '00000000-0000-4000-8000-000000000000',
    },
    Logger: { log() {} },
  });
  vm.runInContext(gasSource, context, { filename: 'gas.js' });

  return { context, metrics, sheets };
}

function post(app, payload) {
  const output = app.context.doPost({
    postData: { contents: JSON.stringify(payload) },
  });
  return JSON.parse(output.text);
}

function payloadFor(action) {
  if (action === 'pres_delete') {
    return { action, pin: '1203', id: 'P_EXISTING', updatedBy: 'test' };
  }
  return {
    action,
    pin: '1203',
    updatedBy: 'test',
    event: {
      id: action === 'pres_update' ? 'P_EXISTING' : undefined,
      title: 'テスト予定',
      startDate: '2099-12-31',
      startTime: '10:00',
      endDate: '2099-12-31',
      endTime: '11:00',
      location: '会議室',
      memo: 'ロック分離テスト',
      category: '会議',
      color: '#1D9E75',
    },
  };
}

test('pres_list bypasses a busy daily-report script lock', () => {
  const app = loadGas({ scriptLockAvailable: false });
  const body = post(app, { action: 'pres_list', pin: '1203' });
  assert.equal(body.status, 'ok');
  assert.equal(app.metrics.scriptTry, 0);
  assert.equal(app.metrics.userTry, 0);
});

test('pres_list skips daily-sheet initialization and preserves the 12-column contract', () => {
  const app = loadGas();
  const body = post(app, { action: 'pres_list', pin: '1203' });
  assert.equal(body.status, 'ok');
  assert.deepEqual(Object.keys(body.rows[0]), PRES_HEADERS);
  assert.equal(app.metrics.dailyDataReads, 0);
});

test('pres_list does not create a missing president sheet', () => {
  const app = loadGas({ presidentSheetExists: false });
  const body = post(app, { action: 'pres_list', pin: '1203' });
  assert.deepEqual(body, { status: 'ok', rows: [] });
  assert.deepEqual(app.metrics.insertedSheets, []);
});

for (const action of ['pres_add', 'pres_update', 'pres_delete']) {
  test(`${action} uses the president write lock instead of the daily-report lock`, () => {
    const app = loadGas({ scriptLockAvailable: false });
    assert.equal(post(app, payloadFor(action)).status, 'ok');
    assert.equal(app.metrics.scriptTry, 0);
    assert.equal(app.metrics.userTry, 1);
    assert.equal(app.metrics.userRelease, 1);
  });
}

test('pres_update appends a new snapshot and pres_list returns only the latest one', () => {
  const app = loadGas();
  const sheet = app.sheets.get('社長予定');

  assert.equal(post(app, payloadFor('pres_update')).status, 'ok');
  assert.equal(sheet.rows.length, 4);

  const listed = post(app, { action: 'pres_list', pin: '1203' });
  const matches = listed.rows.filter(row => row.ID === 'P_EXISTING');
  assert.equal(matches.length, 1);
  assert.equal(matches[0].タイトル, 'テスト予定');
  assert.equal(matches[0].登録日時, EXISTING_ROW[0]);
});

test('pres_add uses a UUID-based opaque ID', () => {
  const app = loadGas();
  const body = post(app, payloadFor('pres_add'));
  assert.equal(body.status, 'ok');
  assert.equal(body.id, 'P00000000000040008000000000000000');
});

test('a busy president write lock returns the existing retryable error', () => {
  const app = loadGas({ userLockAvailable: false });
  const body = post(app, payloadFor('pres_add'));
  assert.equal(body.status, 'error');
  assert.match(body.message, /数秒待って/);
  assert.equal(app.metrics.scriptTry, 0);
  assert.equal(app.metrics.userTry, 1);
});

test('pres_delete leaves a tombstone so another president row cannot shift under a concurrent update', () => {
  const app = loadGas();
  const body = post(app, payloadFor('pres_delete'));
  const sheet = app.sheets.get('社長予定');

  assert.equal(body.status, 'ok');
  assert.equal(sheet.rows.length, 4);
  assert.equal(sheet.rows[3][PRES_HEADERS.indexOf('ID')], 'P_EXISTING');
  assert.equal(sheet.rows[3][PRES_HEADERS.indexOf('カテゴリ')], '__PRES_DELETED__');
  assert.equal(sheet.rows[2][PRES_HEADERS.indexOf('ID')], 'P_SECOND');

  const listed = post(app, { action: 'pres_list', pin: '1203' });
  assert.deepEqual(listed.rows.map(row => row.ID), ['P_SECOND']);
  assert.equal(post(app, payloadFor('pres_update')).status, 'error');
});

test('a delete tombstone wins over a stale same-ID update that commits later', () => {
  const app = loadGas();
  const sheet = app.sheets.get('社長予定');
  const staleUpdate = [...EXISTING_ROW];
  staleUpdate[PRES_HEADERS.indexOf('タイトル')] = '削除と競合した古い更新';

  assert.equal(post(app, payloadFor('pres_delete')).status, 'ok');
  // 別ユーザーの更新が削除前に読み取った内容を、削除後に書く競合を再現する。
  sheet.appendRow(staleUpdate);

  const listed = post(app, { action: 'pres_list', pin: '1203' });
  assert.deepEqual(listed.rows.map(row => row.ID), ['P_SECOND']);
});

test('employee schedule mutations bypass a busy admin-operation user lock', () => {
  const app = loadGas({ userLockAvailable: false });
  const body = post(app, { action: 'delete', ids: [] });
  assert.equal(body.status, 'ok');
  assert.equal(app.metrics.scriptTry, 1);
  assert.equal(app.metrics.userTry, 0);
});
