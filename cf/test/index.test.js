import { describe, it, expect, vi } from 'vitest';
import worker from '../src/index.js';

const HEADERS = ['登録日時','作業日','元請名','現場名','氏名','役割','出勤','退勤','人工',
                 'メモ','夜勤','会社','ID','更新者','色','事業部','工番','作業区分','車両'];

function makeRow(fields) {
  const row = new Array(19).fill('');
  for (const [h, v] of Object.entries(fields)) row[HEADERS.indexOf(h)] = v;
  return row;
}

// D1のprepare/bind/run/all()を模した簡易・状態保持モック。
// snapshot / sync_lock / sync_log の3テーブルをカバーする（index.js経由で
// read.js・sync.jsの両方が発行するクエリをすべて処理できるようにする）。
//
// ★再レビュー対応でsync.js/read.jsのSQLが変わった点を反映している
// （詳細はtest/sync.test.js・test/read.test.jsのコメント参照）:
//   - ロック取得・snapshot書き込みは単一の「INSERT ... ON CONFLICT ... WHERE」文になり、
//     D1のmeta.changesで成否を返す。
//   - read.jsはsnapshotに加えてsync_log(ok=1の直近1件)も見て鮮度を判定する。
function makeMockDB({ snapshot = null, syncLog = null } = {}) {
  const state = { snapshot, lockedAt: null, syncLog: syncLog || [] };

  function respond(sql, args) {
    return {
      async all() {
        if (/SELECT rows, hash, members_count, genba_count, jobsites_count FROM snapshot/.test(sql)) {
          return {
            results: state.snapshot
              ? [{
                  rows: state.snapshot.rows, hash: state.snapshot.hash,
                  members_count: state.snapshot.membersCount || 0, genba_count: state.snapshot.genbaCount || 0,
                  jobsites_count: state.snapshot.jobsitesCount || 0
                }]
              : []
          };
        }
        if (/SELECT payload FROM snapshot/.test(sql)) {
          return { results: state.snapshot ? [{ payload: state.snapshot.payload }] : [] };
        }
        if (/SELECT rows, bytes, at FROM snapshot/.test(sql)) {
          return { results: state.snapshot ? [{ rows: state.snapshot.rows, bytes: state.snapshot.bytes, at: state.snapshot.at }] : [] };
        }
        if (/SELECT at FROM sync_log WHERE ok = 1/.test(sql)) {
          const oks = state.syncLog.filter(l => Number(l.ok) === 1).sort((a, b) => b.at.localeCompare(a.at));
          return { results: oks.length ? [{ at: oks[0].at }] : [] };
        }
        if (/SELECT ok, message FROM sync_log/.test(sql)) {
          return { results: [...state.syncLog].sort((a, b) => b.at.localeCompare(a.at)) };
        }
        if (/FROM sync_log/.test(sql)) {
          return { results: [...state.syncLog].sort((a, b) => b.at.localeCompare(a.at)) };
        }
        return { results: [] };
      },
      async run() {
        if (/VALUES \(1, NULL\)/.test(sql) && /sync_lock/.test(sql)) { state.lockedAt = null; return { success: true, meta: { changes: 1 } }; }
        if (/INSERT INTO sync_lock/.test(sql) && /ON CONFLICT/.test(sql)) {
          const [newLockedAt, staleCutoff] = args;
          const isFree = state.lockedAt == null || Number(state.lockedAt) < Number(staleCutoff);
          if (!isFree) return { success: true, meta: { changes: 0 } };
          state.lockedAt = newLockedAt;
          return { success: true, meta: { changes: 1 } };
        }
        if (/INSERT INTO snapshot/.test(sql) && /ON CONFLICT/.test(sql)) {
          const [payload, hash, rows, membersCount, genbaCount, jobsitesCount, bytes, fetchStartedAt, at] = args;
          const isNewer = !state.snapshot || Number(fetchStartedAt) >= Number(state.snapshot.fetchStartedAt || 0);
          if (!isNewer) return { success: true, meta: { changes: 0 } };
          state.snapshot = { payload, hash, rows, membersCount, genbaCount, jobsitesCount, bytes, fetchStartedAt, at };
          return { success: true, meta: { changes: 1 } };
        }
        if (/INSERT OR REPLACE INTO sync_log/.test(sql)) {
          const [at, rows, ok, message] = args;
          const idx = state.syncLog.findIndex(l => l.at === at);
          const entry = { at, rows, ok, message };
          if (idx >= 0) state.syncLog[idx] = entry; else state.syncLog.push(entry);
          return { success: true };
        }
        return { success: true };
      }
    };
  }

  const db = {
    prepare(sql) {
      return { bind: (...args) => respond(sql, args), all: () => respond(sql, []).all(), run: () => respond(sql, []).run() };
    }
  };
  return { db, state };
}

function makeSnapshot(rows, members = [], overrides = {}) {
  const payload = JSON.stringify({ compact: 1, headers: HEADERS, rows, members, genbaMaster: [], jobsites: [] });
  return {
    payload, hash: 'fixed-hash-for-test', rows: rows.length,
    membersCount: members.length, genbaCount: 0, jobsitesCount: 0,
    bytes: payload.length, fetchStartedAt: 1000, at: '2026-08-24T00:00:00.000Z', ...overrides
  };
}

// ★修正1（鮮度ガード）: /api/scheduleがstatus:'ok'を返すには、snapshotに加えて
// sync_logに「直近15分以内のok=1」が必要になった。/api/scheduleの正常系テストは
// snapshotを直接注入する（syncAllを経由しない）ため、鮮度ログも合わせて用意する。
function freshSyncLog(at = new Date().toISOString()) {
  return [{ at, rows: 0, ok: 1, message: '' }];
}

describe('GET /api/schedule のcompanyパラメータのtrim（レビュー指摘: 会社名の前後空白）', () => {
  it('companyパラメータの前後に空白があっても、trimしてからD1へ問い合わせる', async () => {
    const snapshot = makeSnapshot(
      [makeRow({ ID: 'a-1', 会社: 'グローライズ' })],
      [{ name: '森', company: 'グローライズ', division: 'ICT' }]
    );
    const { db } = makeMockDB({ snapshot, syncLog: freshSyncLog() });
    const env = { DB: db };

    const req = new Request('https://worker.test/api/schedule?' +
      'company=' + encodeURIComponent('  グローライズ　'));
    const res = await worker.fetch(req, env, {});
    const body = await res.json();

    expect(res.status).toBe(200);
    expect(body.status).toBe('ok');
    expect(body.rows).toHaveLength(1);
    expect(body.members).toHaveLength(1);
  });

  it('companyパラメータに空白が無い通常ケースでも従来どおり動く（回帰確認）', async () => {
    const snapshot = makeSnapshot(
      [makeRow({ ID: 'a-1', 会社: 'グローライズ' })],
      [{ name: '森', company: 'グローライズ', division: 'ICT' }]
    );
    const { db } = makeMockDB({ snapshot, syncLog: freshSyncLog() });
    const env = { DB: db };

    const req = new Request('https://worker.test/api/schedule?company=' + encodeURIComponent('グローライズ'));
    const res = await worker.fetch(req, env, {});
    const body = await res.json();

    expect(body.status).toBe('ok');
    expect(body.rows).toHaveLength(1);
  });

  it('snapshotが無ければstatus:errorを返す（画面はGASへ自動フォールバックする）', async () => {
    const { db } = makeMockDB({ snapshot: null });
    const env = { DB: db };
    const res = await worker.fetch(new Request('https://worker.test/api/schedule'), env, {});
    const body = await res.json();
    expect(body.status).toBe('error');
  });
});

describe('POST /api/sync（修正2: 共有秘密による簡易認証）', () => {
  it('SYNC_KEYが未設定なら、認証ヘッダ無しでも実行される', async () => {
    global.fetch = vi.fn(async () => ({
      ok: true, status: 200,
      json: async () => ({ status: 'ok', compact: 1, headers: HEADERS, rows: [], members: [], genbaMaster: [], jobsites: [] })
    }));
    const { db } = makeMockDB();
    const env = { DB: db, GAS_URL: 'https://example.test/exec' }; // SYNC_KEY未設定

    const res = await worker.fetch(new Request('https://worker.test/api/sync', { method: 'POST' }), env, {});
    const body = await res.json();

    expect(res.status).toBe(200);
    expect(body.status).toBe('ok');
  });

  it('SYNC_KEYが設定されているのにヘッダが無いと403で拒否し、同期は実行されない', async () => {
    const fetchMock = vi.fn();
    global.fetch = fetchMock;
    const { db, state } = makeMockDB();
    const env = { DB: db, GAS_URL: 'https://example.test/exec', SYNC_KEY: 'himitsu-123' };

    const res = await worker.fetch(new Request('https://worker.test/api/sync', { method: 'POST' }), env, {});
    const body = await res.json();

    expect(res.status).toBe(403);
    expect(body.status).toBe('error');
    expect(fetchMock).not.toHaveBeenCalled();
    expect(state.snapshot).toBeNull();
  });

  it('SYNC_KEYが設定されていて、ヘッダの値が一致しないと403で拒否する', async () => {
    global.fetch = vi.fn();
    const { db } = makeMockDB();
    const env = { DB: db, GAS_URL: 'https://example.test/exec', SYNC_KEY: 'himitsu-123' };

    const req = new Request('https://worker.test/api/sync', { method: 'POST', headers: { 'X-Sync-Key': 'chigau-key' } });
    const res = await worker.fetch(req, env, {});
    expect(res.status).toBe(403);
  });

  it('SYNC_KEYが設定されていて、ヘッダの値が一致すれば実行される', async () => {
    global.fetch = vi.fn(async () => ({
      ok: true, status: 200,
      json: async () => ({ status: 'ok', compact: 1, headers: HEADERS, rows: [], members: [], genbaMaster: [], jobsites: [] })
    }));
    const { db, state } = makeMockDB();
    const env = { DB: db, GAS_URL: 'https://example.test/exec', SYNC_KEY: 'himitsu-123' };

    const req = new Request('https://worker.test/api/sync', { method: 'POST', headers: { 'X-Sync-Key': 'himitsu-123' } });
    const res = await worker.fetch(req, env, {});
    const body = await res.json();

    expect(res.status).toBe(200);
    expect(body.status).toBe('ok');
    expect(state.snapshot).toBeTruthy();
  });
});

describe('POST /api/sync（修正2: 同時実行の抑止）', () => {
  it('直近に同期が進行中なら、実行せず「進行中」を伝えて200で返す', async () => {
    const fetchMock = vi.fn();
    global.fetch = fetchMock;
    // ロック中の状態を直接注入する（500ミリ秒前に取得＝進行中とみなす範囲内）
    const locked = makeMockDB();
    locked.state.lockedAt = String(Date.now() - 500);
    const env = { DB: locked.db, GAS_URL: 'https://example.test/exec' };

    const res = await worker.fetch(new Request('https://worker.test/api/sync', { method: 'POST' }), env, {});
    const body = await res.json();

    expect(res.status).toBe(200);
    expect(body.status).toBe('ok');
    expect(body.skipped).toBe(true);
    expect(fetchMock).not.toHaveBeenCalled();
  });
});

describe('GET /api/health', () => {
  it('snapshotとsync_logの最新状態を返す', async () => {
    const snapshot = makeSnapshot([makeRow({ ID: 'a-1' })]);
    const { db, state } = makeMockDB({ snapshot });
    state.syncLog.push({ at: '2026-08-24T00:00:00.000Z', rows: 1, ok: 1, message: '' });
    const env = { DB: db };

    const res = await worker.fetch(new Request('https://worker.test/api/health'), env, {});
    const body = await res.json();

    expect(body.status).toBe('ok');
    expect(body.rows).toBe(1);
    expect(body.lastSync).toBeTruthy();
    expect(body.lastSync.ok).toBe(1);
  });

  it('DBアクセスが失敗してもJSONでエラーを返す（素の500にしない）', async () => {
    const env = { DB: { prepare() { return { all: async () => { throw new Error('mock db down'); } }; } } };
    const res = await worker.fetch(new Request('https://worker.test/api/health'), env, {});
    const body = await res.json();
    expect(body.status).toBe('error');
  });
});

describe('GET /api/schedule（修正1・再レビュー: 鮮度ガードの結線）', () => {
  it('sync_logに直近の成功が無ければ、snapshotがあってもstatus:errorを返す（画面はGASへフォールバックする）', async () => {
    const snapshot = makeSnapshot([makeRow({ ID: 'old' })]);
    const { db } = makeMockDB({
      snapshot,
      syncLog: [{ at: new Date(Date.now() - 20 * 60 * 1000).toISOString(), rows: 1, ok: 1, message: '' }] // 20分前＝古い
    });
    const env = { DB: db };
    const res = await worker.fetch(new Request('https://worker.test/api/schedule'), env, {});
    const body = await res.json();
    expect(body.status).toBe('error');
  });
});

describe('POST /api/sync（修正7: force=1の結線）', () => {
  it('?force=1を付けると、急減ガードで拒否されるはずの内容でも受け入れる', async () => {
    // 既存snapshot=600行、今回のGAS応答=100行（半分未満）という急減状況を用意する。
    const rows600 = Array.from({ length: 600 }, (_, i) => makeRow({ ID: 'r' + i, 作業日: '2026-05-01' }));
    const existing = {
      payload: JSON.stringify({ compact: 1, headers: HEADERS, rows: rows600, members: [], genbaMaster: [], jobsites: [] }),
      hash: 'old-hash', rows: 600, membersCount: 0, genbaCount: 0, jobsitesCount: 0,
      bytes: 100, fetchStartedAt: 1000, at: '2026-08-24T00:00:00.000Z'
    };
    const { db, state } = makeMockDB({ snapshot: existing });
    const rows100 = Array.from({ length: 100 }, (_, i) => makeRow({ ID: 'new' + i, 作業日: '2026-05-01' }));
    global.fetch = vi.fn(async () => ({
      ok: true, status: 200,
      json: async () => ({ status: 'ok', compact: 1, headers: HEADERS, rows: rows100, members: [], genbaMaster: [], jobsites: [] })
    }));
    const env = { DB: db, GAS_URL: 'https://example.test/exec' };

    // force無しなら拒否されるはず、という前提の確認は sync.test.js 側で担保済み。
    // ここではforce=1の指定がindex.js→syncAllへ実際に配線されていることを確認する。
    const res = await worker.fetch(new Request('https://worker.test/api/sync?force=1', { method: 'POST' }), env, {});
    const body = await res.json();

    expect(res.status).toBe(200);
    expect(body.status).toBe('ok');
    expect(state.snapshot.rows).toBe(100);
  });
});
