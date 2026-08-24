// ★3回目レビューで繰り返し踏んだ落とし穴（progress.md参照）:
//   「モックDBを使ったテストは全部PASSしていたのに、実物のD1では
//    NOT NULL constraint failed で落ちた」（新スキーマにmembers_countを渡さない
//    古いWorkerが動いていただけ、という原因だった）。
// これは「手で書いたモックが本物のSQL意味論を正確に再現できているとは限らない」
// という構造的な弱点で、sync.test.js/index.test.js/read.test.js のモックは
// すべてこの弱点を抱えている。
//
// このファイルは、手書きモックの代わりに **本物のSQLite**（Node組み込みの
// node:sqlite。D1もSQLite互換）へ実際にschema.sqlを適用し、その上で
// sync.js/read.js/index.jsの本物のコードを動かす。これにより:
//   - schema.sqlの列名・NOT NULL制約と、sync.js/read.jsが発行するSQLが
//     本当に噛み合っているか（列名の綴り間違い・追加し忘れ等）
//   - `INSERT ... ON CONFLICT ... WHERE` の同着(`>`)判定が、手書きモックの
//     JS再現ではなく本物のSQLite上で意図どおりに動くか（3回目レビュー修正4）
//   - 急減ガードの自己回復（payload_hashで同一内容を判定する新しいSQL、
//     3回目レビュー修正3）が本物のSQLite上で意図どおりに動くか
// を、モックのバグに関係なく確認できる。
import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
// ★VitestのESM解決(Vite)は 'node:sqlite' を組み込みモジュールとして認識せず解決に
// 失敗するため、CommonJSのrequireを使って読み込む（Node本体では動作確認済み）。
const { DatabaseSync } = createRequire(import.meta.url)('node:sqlite');
import { syncAll } from '../src/sync.js';
import { readSchedule } from '../src/read.js';
import worker from '../src/index.js';

const SCHEMA_SQL = readFileSync(new URL('../schema.sql', import.meta.url), 'utf8');
const HEADERS = ['登録日時','作業日','元請名','現場名','氏名','役割','出勤','退勤','人工',
                 'メモ','夜勤','会社','ID','更新者','色','事業部','工番','作業区分','車両'];

function makeRow(fields) {
  const row = new Array(19).fill('');
  for (const [h, v] of Object.entries(fields)) row[HEADERS.indexOf(h)] = v;
  return row;
}
function makeRows(n) {
  return Array.from({ length: n }, (_, i) => makeRow({ ID: 'id-' + i, 氏名: '作業員' + i, 人工: 1, 会社: 'グローライズ', 作業日: '2026-05-02' }));
}
function makeCompactPayload(overrides = {}) {
  return { status: 'ok', compact: 1, headers: HEADERS, rows: [], members: [], genbaMaster: [], jobsites: [], ...overrides };
}
function mockFetchOk(payload) {
  global.fetch = vi.fn(async () => ({ ok: true, status: 200, json: async () => payload }));
}

// ★D1のprepare/bind/run/all()互換の薄いアダプタ。本物のSQLを本物のSQLiteに
// そのまま流すだけで、SQLの意味を再現するロジックはここには一切書かない
// （手書きモックとの決定的な違い＝「振る舞いを真似る」のではなく「本物を動かす」）。
function makeRealD1(sqliteDb) {
  return {
    prepare(sql) {
      const stmt = sqliteDb.prepare(sql);
      const exec = (args) => ({
        async run() {
          const info = stmt.run(...args);
          return { success: true, meta: { changes: Number(info.changes) } };
        },
        async all() {
          return { results: stmt.all(...args) };
        }
      });
      return {
        bind: (...args) => exec(args),
        run: () => exec([]).run(),
        all: () => exec([]).all()
      };
    }
  };
}

let sqliteDb;
let env;

beforeEach(() => {
  sqliteDb = new DatabaseSync(':memory:');
  sqliteDb.exec(SCHEMA_SQL);
  env = { DB: makeRealD1(sqliteDb), GAS_URL: 'https://example.test/exec' };
});

afterEach(() => {
  sqliteDb.close();
  vi.restoreAllMocks();
  vi.useRealTimers();
});

describe('本物のSQLite上でのsyncAll（schema.sqlとsync.jsのSQLが噛み合っているか）', () => {
  it('初回同期がNOT NULL制約に違反せず成功する（過去に実際のD1で踏んだ落とし穴の再現防止）', async () => {
    mockFetchOk(makeCompactPayload({ rows: makeRows(5), members: [{ name: '森', company: 'G', division: '' }] }));
    const out = await syncAll(env);
    expect(out.ok).toBe(true);

    const snap = sqliteDb.prepare('SELECT * FROM snapshot WHERE id = 1').get();
    expect(snap.rows).toBe(5);
    expect(snap.members_count).toBe(1);
    const log = sqliteDb.prepare('SELECT * FROM sync_log ORDER BY at DESC LIMIT 1').get();
    expect(Number(log.ok)).toBe(1);
    expect(log.payload_hash).toBeTruthy(); // ★修正3: payload_hash列が実際に書き込まれている
  });

  it('3回目レビュー修正4: 同一のfetch_started_atでの2回目の書き込みは本物のSQLite上でも無条件で弾かれる（同着で旧内容が新内容を上書きしない）', async () => {
    const nowSpy = vi.spyOn(Date, 'now').mockReturnValue(9_000_000); // 常に同じミリ秒を返す

    try {
      mockFetchOk(makeCompactPayload({ rows: makeRows(3).map((r, i) => { r[12] = 'NEW-' + i; return r; }) }));
      const first = await syncAll(env);
      expect(first.ok).toBe(true);

      mockFetchOk(makeCompactPayload({ rows: makeRows(3).map((r, i) => { r[12] = 'OLD-' + i; return r; }) }));
      const second = await syncAll(env);
      expect(second.ok).toBe(true);
      expect(second.skipped).toBe(true);

      const snap = sqliteDb.prepare('SELECT payload FROM snapshot WHERE id = 1').get();
      expect(snap.payload).toContain('NEW-0');
      expect(snap.payload).not.toContain('OLD-0');
    } finally {
      nowSpy.mockRestore();
    }
  });

  it('3回目レビュー修正3: 同一内容の急減拒否が本物のSQLite上でも30分続けば自動的に受け入れる', async () => {
    vi.useFakeTimers();
    let t = 1_700_000_000_000;
    vi.setSystemTime(t);
    try {
      mockFetchOk(makeCompactPayload({ rows: makeRows(600) }));
      await syncAll(env);

      mockFetchOk(makeCompactPayload({ rows: makeRows(100) })); // 半分未満・以後同じ内容
      for (let i = 0; i < 6; i++) {
        t += 5 * 60 * 1000; vi.setSystemTime(t);
        const r = await syncAll(env);
        expect(r.ok).toBe(false);
      }
      t += 5 * 60 * 1000; vi.setSystemTime(t); // 最初の拒否からちょうど30分後
      const accepted = await syncAll(env);
      expect(accepted.ok).toBe(true);
      expect(accepted.message).toMatch(/自動的に受け入れ/);

      const snap = sqliteDb.prepare('SELECT rows FROM snapshot WHERE id = 1').get();
      expect(snap.rows).toBe(100);
    } finally {
      vi.useRealTimers();
    }
  });

  it('3回目レビュー修正3: 毎回異なる内容の欠損では本物のSQLite上でも自動受入されない', async () => {
    const members = (n) => Array.from({ length: n }, (_, i) => ({ name: 'm' + i, company: 'G', division: '' }));
    mockFetchOk(makeCompactPayload({ rows: makeRows(100), members: members(10) }));
    await syncAll(env);

    let last;
    for (const attempt of [
      makeCompactPayload({ rows: makeRows(10), members: members(10) }),      // 日報だけ半減
      makeCompactPayload({ rows: makeRows(100), members: members(1) })       // 職人だけ半減（内容が違う＝別ハッシュ）
    ]) {
      mockFetchOk(attempt);
      last = await syncAll(env);
      expect(last.ok).toBe(false);
    }
    const snap = sqliteDb.prepare('SELECT rows FROM snapshot WHERE id = 1').get();
    expect(snap.rows).toBe(100); // 拒否され続け、最初のまま
  });
});

describe('本物のSQLite上でのreadSchedule（read.jsの鮮度ガードが噛み合っているか）', () => {
  it('同期成功直後はstatus:okで読める', async () => {
    mockFetchOk(makeCompactPayload({ rows: makeRows(3) }));
    await syncAll(env);
    const out = await readSchedule(env, '');
    expect(out.status).toBe('ok');
    expect(out.rows).toHaveLength(3);
  });

  it('直近の成功が15分より古ければstatus:errorを返す', async () => {
    vi.useFakeTimers();
    try {
      vi.setSystemTime(1_800_000_000_000);
      mockFetchOk(makeCompactPayload({ rows: makeRows(3) }));
      await syncAll(env);

      vi.setSystemTime(1_800_000_000_000 + 20 * 60 * 1000); // 20分後
      const out = await readSchedule(env, '');
      expect(out.status).toBe('error');
    } finally {
      vi.useRealTimers();
    }
  });
});

describe('本物のSQLite上でのworker.fetch（index.jsの配線が噛み合っているか）', () => {
  it('POST /api/sync → GET /api/scheduleが一連で動く（Origin検証・レート制限を含む実配線の確認）', async () => {
    mockFetchOk(makeCompactPayload({ rows: makeRows(4) }));
    const syncRes = await worker.fetch(
      new Request('https://worker.test/api/sync', { method: 'POST', headers: { Origin: 'https://akirasp4-lgtm.github.io' } }),
      env, {}
    );
    expect(syncRes.status).toBe(200);
    const syncBody = await syncRes.json();
    expect(syncBody.status).toBe('ok');

    const readRes = await worker.fetch(new Request('https://worker.test/api/schedule'), env, {});
    const readBody = await readRes.json();
    expect(readBody.status).toBe('ok');
    expect(readBody.rows).toHaveLength(4);
  });

  it('Originが無いPOST /api/syncは本物のDBに一切触れず403で終わる', async () => {
    const res = await worker.fetch(new Request('https://worker.test/api/sync', { method: 'POST' }), env, {});
    expect(res.status).toBe(403);
    const snap = sqliteDb.prepare('SELECT * FROM snapshot WHERE id = 1').get();
    expect(snap).toBeUndefined();
  });
});
