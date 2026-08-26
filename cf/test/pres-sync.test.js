import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { syncPresident, PRES_SHRINK_AUTO_ACCEPT_MS } from '../src/pres-sync.js';

function makeEvent(i) {
  return {
    '登録日時': '2026-08-01T00:00:00.000Z',
    'タイトル': '予定' + i,
    '開始日': '2026-08-20', '開始時刻': '10:00',
    '終了日': '2026-08-20', '終了時刻': '11:00',
    '場所': '', 'メモ': '', 'カテゴリ': '', '色': '',
    'ID': 'P00000000' + i, '更新者': ''
  };
}
const events = n => Array.from({ length: n }, (_, i) => makeEvent(i));

/**
 * env.DB の模擬。
 * - pres_snapshot は1行だけ保持し、社員用と同じ「fetch_started_at が厳密に新しいときだけ
 *   上書き」の条件を実際に評価する（世代ガードが本当に効くかをテストするため）。
 * - 社員用のテーブル(snapshot / sync_log)に触れたら例外を投げる。
 */
function makeMockDB({ snapshot = null, log = [] } = {}) {
  const state = { snapshot, log: [...log] };
  const db = {
    state,
    prepare(sql) {
      const bound = [];
      const api = {
        bind(...args) { bound.push(...args); return api; },
        all: async () => {
          if (/FROM sync_log|FROM snapshot\b|INTO snapshot\b|INTO sync_log/.test(sql)) {
            throw new Error('社長用の取り込みが社員用のテーブルに触れた: ' + sql);
          }
          if (/SELECT .* FROM pres_snapshot/.test(sql)) {
            return { results: state.snapshot ? [state.snapshot] : [] };
          }
          if (/FROM pres_sync_log/.test(sql)) {
            const sorted = [...state.log].sort((a, b) => b.at.localeCompare(a.at));
            return { results: sorted };
          }
          return { results: [] };
        },
        run: async () => {
          if (/INTO snapshot\b|INTO sync_log|sync_lock/.test(sql)) {
            throw new Error('社長用の取り込みが社員用のテーブルに書いた: ' + sql);
          }
          if (/INSERT INTO pres_sync_log/.test(sql)) {
            const [at, rows, ok, message, payload_hash] = bound;
            state.log.push({ at, rows, ok, message, payload_hash });
            return { meta: { changes: 1 } };
          }
          if (/INSERT INTO pres_snapshot/.test(sql)) {
            const [payload, hash, rows, bytes, fetchStartedAt, at] = bound;
            const incoming = Number(fetchStartedAt);
            if (state.snapshot && !(incoming > Number(state.snapshot.fetch_started_at))) {
              return { meta: { changes: 0 } };   // 世代ガードが弾いた
            }
            state.snapshot = { payload, hash, rows, bytes, fetch_started_at: fetchStartedAt, at };
            return { meta: { changes: 1 } };
          }
          if (/DELETE FROM pres_sync_log/.test(sql)) return { meta: { changes: 0 } };
          return { meta: { changes: 0 } };
        }
      };
      return api;
    }
  };
  return db;
}

const baseEnv = (db, over = {}) => ({
  DB: db,
  GAS_URL: 'https://example.test/exec',
  PRES_PIN: '1203',
  ...over
});

function mockGas(responder) {
  globalThis.fetch = vi.fn(async (url, init) => {
    const body = init && init.body ? JSON.parse(init.body) : {};
    const out = await responder(body, url);
    if (out instanceof Response) return out;
    return new Response(JSON.stringify(out), { status: 200 });
  });
}

let realFetch;
beforeEach(() => { realFetch = globalThis.fetch; });
afterEach(() => { globalThis.fetch = realFetch; vi.restoreAllMocks(); });

describe('syncPresident（GASへの問い合わせ）', () => {
  it('pres_list をPINつきでPOSTする（PINはURLに載せない）', async () => {
    let seenUrl = '', seenBody = null;
    mockGas((body, url) => { seenUrl = url; seenBody = body; return { status: 'ok', rows: events(3) }; });
    const db = makeMockDB();
    const r = await syncPresident(baseEnv(db));
    expect(r.ok).toBe(true);
    expect(seenBody.action).toBe('pres_list');
    expect(seenBody.pin).toBe('1203');
    expect(seenUrl).not.toContain('1203');       // PINがURLに漏れていない
    expect(db.state.snapshot.rows).toBe(3);
  });

  it('PRES_PINが未設定なら、GASを叩かず失敗として記録する（何も壊さない）', async () => {
    mockGas(() => { throw new Error('呼ばれてはいけない'); });
    const db = makeMockDB();
    const r = await syncPresident(baseEnv(db, { PRES_PIN: '' }));
    expect(r.ok).toBe(false);
    expect(db.state.snapshot).toBe(null);
    expect(db.state.log.at(-1).ok).toBe(0);
  });

  it('GASが認証失敗を返したら書き込まず失敗として記録する', async () => {
    mockGas(() => ({ status: 'error', message: '認証に失敗しました' }));
    const db = makeMockDB();
    const r = await syncPresident(baseEnv(db));
    expect(r.ok).toBe(false);
    expect(db.state.snapshot).toBe(null);
  });

  it('rowsが配列でない壊れた応答は書き込まない', async () => {
    mockGas(() => ({ status: 'ok', rows: { と: 'おかしい' } }));
    const db = makeMockDB();
    const r = await syncPresident(baseEnv(db));
    expect(r.ok).toBe(false);
    expect(db.state.snapshot).toBe(null);
  });

  it('GASが落ちていても例外を投げず {ok:false} を返す', async () => {
    globalThis.fetch = vi.fn(async () => new Response('<html>エラー</html>', { status: 500 }));
    const db = makeMockDB();
    const r = await syncPresident(baseEnv(db));
    expect(r.ok).toBe(false);
    expect(db.state.snapshot).toBe(null);
  });
});

describe('syncPresident（変更なしスキップ）', () => {
  it('同じ内容なら書き込まないが ok=1 で記録する（鮮度ガードを誤発火させないため）', async () => {
    mockGas(() => ({ status: 'ok', rows: events(2) }));
    const db = makeMockDB();
    await syncPresident(baseEnv(db));
    const firstAt = db.state.snapshot.at;

    const r2 = await syncPresident(baseEnv(db));
    expect(r2.ok).toBe(true);
    expect(r2.skipped).toBe(true);
    expect(r2.skipReason).toBe('unchanged');
    expect(db.state.snapshot.at).toBe(firstAt);          // 書き換わっていない
    expect(db.state.log.at(-1).ok).toBe(1);              // それでも成功として記録
  });
});

describe('syncPresident（世代の逆転を防ぐ）', () => {
  it('保存済みより古い取得開始時刻の結果では上書きしない', async () => {
    mockGas(() => ({ status: 'ok', rows: events(5) }));
    const db = makeMockDB();
    await syncPresident(baseEnv(db));
    expect(db.state.snapshot.rows).toBe(5);

    // 「先に始まったが遅れて完了した」古い取得を模す
    const old = Number(db.state.snapshot.fetch_started_at) - 10_000;
    mockGas(() => ({ status: 'ok', rows: events(4) }));
    const r = await syncPresident(baseEnv(db), { fetchStartedAtOverride: old });
    expect(r.ok).toBe(true);
    expect(r.skipped).toBe(true);
    expect(db.state.snapshot.rows).toBe(5);              // 古い内容で潰れていない
  });
});

describe('syncPresident（急減ガード）', () => {
  it('件数が保存済みの半分未満なら拒否して書き込まない', async () => {
    mockGas(() => ({ status: 'ok', rows: events(10) }));
    const db = makeMockDB();
    await syncPresident(baseEnv(db));

    mockGas(() => ({ status: 'ok', rows: events(2) }));   // 10 → 2（半分未満）
    const r = await syncPresident(baseEnv(db));
    expect(r.ok).toBe(false);
    expect(db.state.snapshot.rows).toBe(10);
    expect(db.state.log.at(-1).payload_hash).toBeTruthy(); // 内容のハッシュを記録している
  });

  it('半分ちょうどは拒否しない（境界）', async () => {
    // ★取得開始時刻を明示的にずらす。実装は「厳密に新しいときだけ上書き」なので、
    //   同一ミリ秒に2回走ると2回目は世代ガードで弾かれる（設計どおりの正しい動作）。
    //   ここで見たいのは急減ガードの境界なので、時刻の衝突を排除して分離する。
    const t0 = Date.now();
    mockGas(() => ({ status: 'ok', rows: events(10) }));
    const db = makeMockDB();
    await syncPresident(baseEnv(db), { fetchStartedAtOverride: t0 });
    mockGas(() => ({ status: 'ok', rows: events(5) }));      // 10 → 5 はちょうど半分
    const r = await syncPresident(baseEnv(db), { fetchStartedAtOverride: t0 + 1000 });
    expect(r.ok).toBe(true);
    expect(db.state.snapshot.rows).toBe(5);
  });

  it('同じ内容の拒否が30分続いたら受け入れる（正当な一括削除で永久に詰まらない）', async () => {
    const t0 = Date.now();
    mockGas(() => ({ status: 'ok', rows: events(10) }));
    const db = makeMockDB();
    await syncPresident(baseEnv(db), { fetchStartedAtOverride: t0 });

    // ★実際に起きる順番を再現する: 成功が一番古く、同じ内容の拒否がそのあと積み上がる。
    //   （成功より古い時刻に拒否を置くと現実には起こらない並びになり、
    //     「新しい順に遡って連続を数える」判定が成功記録で止まってしまう）
    db.state.log.at(-1).at = new Date(t0 - 90 * 60_000).toISOString();

    mockGas(() => ({ status: 'ok', rows: events(2) }));
    const r1 = await syncPresident(baseEnv(db), { fetchStartedAtOverride: t0 + 1000 });
    expect(r1.ok).toBe(false);
    // その最初の拒否は31分前だったことにする（＝同じ内容の拒否が31分続いている状態）
    db.state.log.at(-1).at = new Date(Date.now() - PRES_SHRINK_AUTO_ACCEPT_MS - 60_000).toISOString();

    const r2 = await syncPresident(baseEnv(db), { fetchStartedAtOverride: t0 + 2000 });
    expect(r2.ok).toBe(true);
    expect(db.state.snapshot.rows).toBe(2);
  });

  it('同じミリ秒に始まった2回目は上書きしない（世代ガードの境界・設計どおり）', async () => {
    const t0 = Date.now();
    mockGas(() => ({ status: 'ok', rows: events(10) }));
    const db = makeMockDB();
    await syncPresident(baseEnv(db), { fetchStartedAtOverride: t0 });
    mockGas(() => ({ status: 'ok', rows: events(9) }));
    const r = await syncPresident(baseEnv(db), { fetchStartedAtOverride: t0 });
    expect(r.ok).toBe(true);            // 取得自体は正常なので失敗ではない
    expect(r.skipReason).toBe('stale-generation');
    expect(db.state.snapshot.rows).toBe(10);
  });

  it('拒否のたびに内容が違う場合は自動受入しない（回数だけで通さない）', async () => {
    mockGas(() => ({ status: 'ok', rows: events(10) }));
    const db = makeMockDB();
    await syncPresident(baseEnv(db));

    for (let i = 0; i < 4; i++) {
      mockGas(() => ({ status: 'ok', rows: [makeEvent(100 + i)] }));  // 毎回別の内容
      const r = await syncPresident(baseEnv(db));
      expect(r.ok).toBe(false);
      // 古い時刻にしても「別の内容」なので通ってはいけない
      db.state.log.at(-1).at = new Date(Date.now() - PRES_SHRINK_AUTO_ACCEPT_MS - 60_000).toISOString();
    }
    expect(db.state.snapshot.rows).toBe(10);
  });

  it('初回（保存済みが無い）は急減ガードを飛ばす', async () => {
    mockGas(() => ({ status: 'ok', rows: events(1) }));
    const db = makeMockDB();
    const r = await syncPresident(baseEnv(db));
    expect(r.ok).toBe(true);
    expect(db.state.snapshot.rows).toBe(1);
  });
});

describe('syncPresident（サイズ上限）', () => {
  it('上限を超える大きさは書き込まない', async () => {
    const huge = [{ ...makeEvent(0), 'メモ': 'あ'.repeat(600_000) }];
    mockGas(() => ({ status: 'ok', rows: huge }));
    const db = makeMockDB();
    const r = await syncPresident(baseEnv(db));
    expect(r.ok).toBe(false);
    expect(db.state.snapshot).toBe(null);
  });
});

describe('syncPresident（社員用と混ざっていないこと）', () => {
  it('社員用の snapshot / sync_log / sync_lock に一切触れない', async () => {
    mockGas(() => ({ status: 'ok', rows: events(3) }));
    const db = makeMockDB();   // 社員用テーブルに触れると例外を投げる
    const r = await syncPresident(baseEnv(db));
    expect(r.ok).toBe(true);
  });
});

// ★★2026-08-26 本番障害の再発防止（最重要）
//   社長のカレンダーに予定が出なくなった。原因は、GASへのPOSTがときどき doGet に
//   届いてしまい、その応答（社員の日報データ）が {status:'ok', rows:[...]} という
//   同じ形をしているため validate を通過し、社長予定として保存されていたこと。
//   さらに、その後の正しい202件が急減ガードに「2652→202」と判定されて拒否され続け、
//   間違ったデータが居座り続けた。
describe('syncPresident（社員の日報データを社長予定として保存しない）', () => {
  // GASの doGet が返す形（compact指定なし）。status:'ok' と rows:[] を持つので
  // 形だけ見ると pres_list と区別が付かない。
  const doGetLike = {
    status: 'ok',
    rows: [{ '登録日時': '2026-05-01', '作業日': '2026-05-25', '元請名': 'カワセツ',
             '現場名': '京都縦貫道', '氏名': '東', '役割': '代表', '会社': 'グローライズ',
             'ID': '56b2-299', '拠点': '本社' }],
    members: [{ name: '東', company: 'グローライズ' }],
    genbaMaster: [{ name: 'カワセツ' }],
    jobsites: [{ genba: 'カワセツ', loc: '京都縦貫道' }]
  };

  it('★doGetの応答（社員の日報データ）は絶対に保存しない', async () => {
    mockGas(() => doGetLike);
    const db = makeMockDB();
    const r = await syncPresident(baseEnv(db));
    expect(r.ok).toBe(false);
    expect(db.state.snapshot).toBe(null);
  });

  it('★既に社長予定が入っているとき、doGetの応答で上書きしない', async () => {
    mockGas(() => ({ status: 'ok', rows: events(202) }));
    const db = makeMockDB();
    const t0 = Date.now();
    await syncPresident(baseEnv(db), { fetchStartedAtOverride: t0 });
    expect(db.state.snapshot.rows).toBe(202);

    mockGas(() => doGetLike);
    const r = await syncPresident(baseEnv(db), { fetchStartedAtOverride: t0 + 1000 });
    expect(r.ok).toBe(false);
    expect(db.state.snapshot.rows).toBe(202);     // 社長予定のまま
  });

  it('★doGetの目印（members/genbaMaster/jobsites）が付いていたら拒否する', async () => {
    mockGas(() => ({ status: 'ok', rows: events(3), members: [], genbaMaster: [], jobsites: [] }));
    const db = makeMockDB();
    const r = await syncPresident(baseEnv(db));
    expect(r.ok).toBe(false);
    expect(db.state.snapshot).toBe(null);
  });

  it('★社員用の列（作業日・氏名）を持つ行が混ざっていたら拒否する', async () => {
    mockGas(() => ({ status: 'ok', rows: [{ 'ID': 'P1', 'タイトル': 'x', '作業日': '2026-08-01', '氏名': '東' }] }));
    const db = makeMockDB();
    const r = await syncPresident(baseEnv(db));
    expect(r.ok).toBe(false);
  });

  it('本物の社長予定（0件でも）は通る', async () => {
    mockGas(() => ({ status: 'ok', rows: [] }));
    const db = makeMockDB();
    const r = await syncPresident(baseEnv(db));
    expect(r.ok).toBe(true);
    expect(db.state.snapshot.rows).toBe(0);
  });
});
