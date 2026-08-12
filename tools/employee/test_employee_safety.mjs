import assert from 'node:assert/strict';
import fs from 'node:fs';
import test from 'node:test';
import vm from 'node:vm';

const root = new URL('../../', import.meta.url);
const gasSource = fs.readFileSync(new URL('gas.js', root), 'utf8');

const DAILY_HEADERS = [
  '登録日時', '作業日', '元請名', '現場名', '氏名', '役割', '出勤',
  '退勤', '人工', 'メモ', '夜勤', '会社', 'ID', '更新者', '色',
  '事業部', '工番', '作業区分', '車両',
];

const EXISTING_DAILY_ROW = [
  '2026-08-12 09:00:00', '2026-08-12', '既存元請', '既存現場',
  '田中', '代表', '08:00', '17:00', 1, '', '', 'ラーテル',
  'OLD-ID', 'Ryo', '', '', '', '現場作業', '',
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
    return Array.from({ length: this.numRows }, (_, r) => {
      const source = this.sheet.rows[this.row - 1 + r] || [];
      return Array.from(
        { length: this.numColumns },
        (_, c) => source[this.column - 1 + c] ?? '',
      );
    });
  }

  setValue(value) {
    return this.setValues([[value]]);
  }

  setValues(values) {
    this.sheet.metrics.events.push(`setValues:${this.sheet.name}`);
    for (let r = 0; r < this.numRows; r++) {
      while (this.sheet.rows.length < this.row + r) this.sheet.rows.push([]);
      const target = this.sheet.rows[this.row - 1 + r];
      for (let c = 0; c < this.numColumns; c++) {
        target[this.column - 1 + c] = values[r][c];
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
  }

  getDataRange() {
    const width = Math.max(1, ...this.rows.map(row => row.length));
    return new FakeRange(this, 1, 1, Math.max(1, this.rows.length), width);
  }

  getMaxColumns() {
    return Math.max(1, ...this.rows.map(row => row.length));
  }

  insertColumnsAfter(after, howMany) {
    for (const row of this.rows) {
      while (row.length < after + howMany) row.push('');
    }
  }

  getRange(row, column, numRows = 1, numColumns = 1) {
    return new FakeRange(this, row, column, numRows, numColumns);
  }

  getLastRow() {
    return this.rows.length;
  }

  appendRow(row) {
    this.metrics.events.push(`appendRow:${this.name}`);
    this.rows.push([...row]);
    return this;
  }

  deleteRow(row) {
    this.metrics.events.push(`deleteRow:${this.name}`);
    this.rows.splice(row - 1, 1);
  }
}

function loadGas({
  adminLockAvailable = true,
  dailyDataLockAvailable = true,
  withCompanyFixtures = false,
} = {}) {
  const metrics = {
    adminTry: 0,
    adminRelease: 0,
    dailyDataTry: 0,
    dailyDataRelease: 0,
    events: [],
  };
  const sheets = new Map([
    ['日報データ', new FakeSheet(
      '日報データ', [DAILY_HEADERS, EXISTING_DAILY_ROW], metrics,
    )],
  ]);
  if (withCompanyFixtures) {
    sheets.get('日報データ').rows.push([
      '2026-08-12 10:00:00', '2026-08-12', '他社元請', '他社現場',
      '佐藤', '代表', '08:00', '17:00', 1, '', '', '和信カインド',
      'OTHER-ID', 'Ryo', '', '', '', '現場作業', '',
    ]);
    sheets.set('職人マスタ', new FakeSheet('職人マスタ', [
      ['氏名', '会社', '事業部', '単価'],
      ['田中', 'ラーテル', 'INF', 20000],
      ['佐藤', '和信カインド', 'MSC', 30000],
    ], metrics));
    sheets.set('元請マスタ', new FakeSheet('元請マスタ', [
      ['元請名', '会社', '読み'],
      ['既存元請', 'ラーテル', 'きそん'],
      ['他社元請', '和信カインド', 'たしゃ'],
      ['共通元請', '', 'きょうつう'],
    ], metrics));
    sheets.set('現場マスタ', new FakeSheet('現場マスタ', [
      ['元請名', '現場名', '工番', '事業部', '年度', '連番', '売上', '読み', '完了', '請求方式'],
      ['既存元請', '既存現場', 'K-1', 'INF', 2026, 1, 0, '', '', '応援'],
      ['他社元請', '他社現場', 'K-2', 'MSC', 2026, 2, 0, '', '', '応援'],
      ['共通元請', '共通現場', '', '', 2026, 3, 0, '', '', '応援'],
    ], metrics));
  }
  const spreadsheet = {
    getSheetByName(name) {
      return sheets.get(name) || null;
    },
    insertSheet(name) {
      const sheet = new FakeSheet(name, [], metrics);
      sheets.set(name, sheet);
      return sheet;
    },
    toast() {},
  };
  const scriptProperties = new Map([['CAL_REQUIRE_TOKEN', '0']]);
  const adminLock = {
    tryLock() {
      metrics.adminTry++;
      return adminLockAvailable;
    },
    releaseLock() {
      metrics.adminRelease++;
    },
  };
  const dailyDataLock = {
    tryLock() {
      metrics.dailyDataTry++;
      return dailyDataLockAvailable;
    },
    releaseLock() {
      metrics.dailyDataRelease++;
    },
  };
  const ContentService = {
    MimeType: { JSON: 'application/json' },
    createTextOutput(text) {
      return {
        text,
        setMimeType() { return this; },
      };
    },
  };
  const context = vm.createContext({
    console,
    ContentService,
    LockService: {
      // GAS Web apps do not have a containing document, so DocumentLock is null.
      getDocumentLock: () => null,
      getScriptLock: () => dailyDataLock,
      getUserLock: () => adminLock,
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
        if (format === 'HH:mm') return iso.slice(11, 16);
        if (format === 'yyyy-MM-dd HH:mm:ss') return iso.slice(0, 19).replace('T', ' ');
        return iso.slice(0, 10);
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

function get(app, parameters = {}) {
  const output = app.context.doGet({ parameter: parameters });
  return JSON.parse(output.text);
}

function replacementRow() {
  return {
    date: '2026-08-13',
    genba: '新元請',
    loc: '新現場',
    name: '田中',
    role: '代表',
    start: '08:00',
    end: '17:00',
    kosu: 1,
    memo: '',
    company: 'ラーテル',
    id: 'NEW-ID',
    updatedBy: 'test',
    workType: '現場作業',
  };
}

test('get_sheet uses the all-user ScriptLock even when DocumentLock is null', () => {
  const app = loadGas({ dailyDataLockAvailable: false });

  const body = post(app, { action: 'get_sheet', sheet: '日報データ' });

  assert.equal(body.status, 'error');
  assert.match(body.message, /更新中/);
  assert.equal(app.metrics.dailyDataTry, 1);
  assert.equal(app.metrics.adminTry, 0);
});

test('get_sheet waits for an employee update instead of reading duplicate rows', () => {
  const app = loadGas({ dailyDataLockAvailable: false });

  const body = post(app, { action: 'get_sheet', sheet: '日報データ' });

  assert.equal(body.status, 'error');
  assert.match(body.message, /更新中/);
  assert.equal(app.metrics.dailyDataTry, 1);
  assert.equal(app.metrics.adminTry, 0);
});

test('invalid update rows never delete the existing schedule', () => {
  const app = loadGas();

  const body = post(app, { action: 'update', ids: ['OLD-ID'] });

  assert.equal(body.status, 'error');
  assert.equal(app.sheets.get('日報データ').rows.length, 2);
  assert.equal(app.metrics.events.includes('deleteRow:日報データ'), false);
});

test('employee update bypasses a busy admin-operation user lock', () => {
  const app = loadGas({ adminLockAvailable: false });

  const body = post(app, {
    action: 'update',
    ids: ['OLD-ID'],
    rows: [replacementRow()],
  });

  assert.equal(body.status, 'ok');
  assert.equal(app.metrics.adminTry, 0);
  assert.equal(app.metrics.dailyDataTry, 1);
  assert.equal(app.metrics.dailyDataRelease, 1);
});

test('admin actions that mutate schedules also acquire the daily-data lock', () => {
  const app = loadGas({ dailyDataLockAvailable: false });

  const body = post(app, { action: 'archive', months: 3 });

  assert.equal(body.status, 'error');
  assert.match(body.message, /予定を更新中/);
  assert.equal(app.metrics.adminTry, 1);
  assert.equal(app.metrics.dailyDataTry, 1);
});

test('automatic archive participates in the same all-user daily-data lock', () => {
  const app = loadGas({ dailyDataLockAvailable: false });
  let archiveCalls = 0;
  app.context.archiveOldData_ = () => { archiveCalls++; return 0; };

  app.context.autoArchive();

  assert.equal(archiveCalls, 0);
  assert.equal(app.metrics.dailyDataTry, 1);
  assert.equal(app.metrics.adminTry, 0);
});

test('job-number backfill participates in the same all-user daily-data lock', () => {
  const app = loadGas({ dailyDataLockAvailable: false });
  let backfillCalls = 0;
  app.context.backfillJobNosForSheet_ = () => {
    backfillCalls++;
    return { assigned: 0, skippedNoSite: 0, skippedNoDivision: 0 };
  };

  app.context.backfillJobNos();

  assert.equal(backfillCalls, 0);
  assert.equal(app.metrics.dailyDataTry, 1);
  assert.equal(app.metrics.adminTry, 0);
});

test('long yomi backfill uses the admin lock and never occupies the daily-data lock', () => {
  const app = loadGas({ adminLockAvailable: false });
  let yomiCalls = 0;
  app.context._backfillYomiInSheet_ = () => {
    yomiCalls++;
    return { filled: 0, target: 0 };
  };

  app.context.backfillAllYomi();

  assert.equal(yomiCalls, 0);
  assert.equal(app.metrics.adminTry, 1);
  assert.equal(app.metrics.dailyDataTry, 0);
});

test('valid update stores the replacement before deleting the old schedule', () => {
  const app = loadGas();

  const body = post(app, {
    action: 'update',
    ids: ['OLD-ID'],
    rows: [replacementRow()],
  });

  assert.equal(body.status, 'ok');
  const writeIndex = app.metrics.events.indexOf('setValues:日報データ');
  const deleteIndex = app.metrics.events.indexOf('deleteRow:日報データ');
  assert.notEqual(writeIndex, -1);
  assert.notEqual(deleteIndex, -1);
  assert.ok(writeIndex < deleteIndex);
  assert.equal(app.sheets.get('日報データ').rows[1][12], 'NEW-ID');
  assert.equal(app.metrics.dailyDataTry, 1);
  assert.equal(app.metrics.dailyDataRelease, 1);
});

test('unknown actions return an error instead of a silent success', () => {
  const app = loadGas();

  const body = post(app, { action: 'typo_action' });

  assert.equal(body.status, 'error');
  assert.match(body.message, /typo_action/);
});

test('company-filtered GET never returns another company schedule or rate', () => {
  const app = loadGas({ withCompanyFixtures: true });

  const body = get(app, { company: 'ラーテル' });

  assert.equal(body.status, 'ok');
  assert.deepEqual(body.rows.map(row => row['会社']), ['ラーテル']);
  assert.deepEqual(body.members, [
    { name: '田中', company: 'ラーテル', division: 'INF', rate: 20000 },
  ]);
  assert.deepEqual(body.genbaMaster, [
    { name: '既存元請', company: 'ラーテル' },
    { name: '共通元請', company: '' },
  ]);
  assert.deepEqual(body.jobsites.map(site => site.genba), ['既存元請', '共通元請']);
});

test('GET without a company keeps the backwards-compatible all-company response', () => {
  const app = loadGas({ withCompanyFixtures: true });

  const body = get(app);

  assert.equal(body.status, 'ok');
  assert.equal(body.rows.length, 2);
  assert.equal(body.members.length, 2);
  assert.equal(body.genbaMaster.length, 3);
  assert.equal(body.jobsites.length, 3);
});

test('automatic summary never deletes a jobsite as an orphan', () => {
  const app = loadGas();
  let cleanupCalls = 0;
  app.context.generateCompanySummary_ = () => {
    // 長い集計書込みを始める前に、日報スナップショット用ロックを解放する。
    assert.equal(app.metrics.dailyDataRelease, 1);
  };
  app.context.generateMonthSummary_ = () => {};
  app.context.generateBillingSummary_ = () => {};
  app.context.generateBillingFilterSheet_ = () => {};
  app.context.generateKakuninTable_ = () => {};
  app.context.generateDivisionAllocation_ = () => {};
  app.context.cleanupOrphanSites_ = () => {
    cleanupCalls++;
    return 1;
  };

  app.context.generateSummary_();

  assert.equal(cleanupCalls, 0);
  assert.equal(app.metrics.dailyDataTry, 1);
  assert.equal(app.metrics.dailyDataRelease, 1);
});
