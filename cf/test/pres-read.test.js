import { describe, it, expect } from 'vitest';
import { readPresident, PRES_FRESHNESS_THRESHOLD_MS } from '../src/pres-read.js';

// pres_list（gas.js の serializePresidentRows_）が返す1件の形
function makeEvent(overrides = {}) {
  return {
    '登録日時': '2026-08-01T00:00:00.000Z',
    'タイトル': '打合せ',
    '開始日': '2026-08-20',
    '開始時刻': '10:00',
    '終了日': '2026-08-20',
    '終了時刻': '11:00',
    '場所': '本社',
    'メモ': '',
    'カテゴリ': '',
    '色': '',
    'ID': 'P' + Math.random().toString(36).slice(2, 12),
    '更新者': '',
    ...overrides
  };
}

function makeMockDB({ payload = null, log = null, freshSuccess = true } = {}) {
  const entries = log != null
    ? log
    : (freshSuccess ? [{ at: new Date().toISOString(), ok: 1 }] : []);
  return {
    prepare(sql) {
      return {
        all: async () => {
          if (/SELECT payload FROM pres_snapshot/.test(sql)) {
            return { results: payload != null ? [{ payload }] : [] };
          }
          if (/FROM pres_sync_log WHERE ok = 1/.test(sql)) {
            const oks = entries.filter(l => Number(l.ok) === 1)
              .sort((a, b) => b.at.localeCompare(a.at));
            return { results: oks.length ? [{ at: oks[0].at }] : [] };
          }
          // ★社員用のテーブルを読んだら即座に分かるようにする
          if (/FROM sync_log|FROM snapshot\b/.test(sql)) {
            throw new Error('社長用の読み取りが社員用のテーブルを見ている: ' + sql);
          }
          return { results: [] };
        }
      };
    }
  };
}

describe('readPresident（未取り込み・壊れている場合の安全装置）', () => {
  it('pres_snapshotが1行も無いときはエラーを返す（空を「予定ゼロ件」として返さない）', async () => {
    const out = await readPresident({ DB: makeMockDB({ payload: null }) });
    expect(out.status).toBe('error');
    expect(out.rows).toBeUndefined();
  });

  it('保存済みJSONが壊れていても例外を投げずエラーを返す（画面はGASへ落ちる）', async () => {
    const out = await readPresident({ DB: makeMockDB({ payload: '{壊れ' }) });
    expect(out.status).toBe('error');
  });
});

describe('readPresident（鮮度ガード）', () => {
  it('同期の成功記録が1件も無ければエラー', async () => {
    const out = await readPresident({ DB: makeMockDB({
      payload: JSON.stringify([makeEvent()]), log: []
    }) });
    expect(out.status).toBe('error');
  });

  it('直近の成功がしきい値より古ければエラー（古い写しを正常として返さない）', async () => {
    const old = new Date(Date.now() - PRES_FRESHNESS_THRESHOLD_MS - 60_000).toISOString();
    const out = await readPresident({ DB: makeMockDB({
      payload: JSON.stringify([makeEvent()]), log: [{ at: old, ok: 1 }]
    }) });
    expect(out.status).toBe('error');
  });

  it('直近の成功がしきい値以内なら正常に返る', async () => {
    const ev = makeEvent({ タイトル: '銀行' });
    const out = await readPresident({ DB: makeMockDB({ payload: JSON.stringify([ev]) }) });
    expect(out.status).toBe('ok');
    expect(out.rows).toHaveLength(1);
    expect(out.rows[0]['タイトル']).toBe('銀行');
  });

  it('失敗記録しか無ければ（ok=0だけ）エラー扱い', async () => {
    const out = await readPresident({ DB: makeMockDB({
      payload: JSON.stringify([makeEvent()]),
      log: [{ at: new Date().toISOString(), ok: 0 }]
    }) });
    expect(out.status).toBe('error');
  });
});

describe('readPresident（GASのpres_listと同じ形で返す）', () => {
  it('{status:"ok", rows:[...]} の形で、rowsは保存されたものをそのまま返す', async () => {
    const evs = [makeEvent({ タイトル: 'A' }), makeEvent({ タイトル: 'B' })];
    const out = await readPresident({ DB: makeMockDB({ payload: JSON.stringify(evs) }) });
    expect(Object.keys(out).sort()).toEqual(['rows', 'status']);
    expect(out.rows.map(r => r['タイトル'])).toEqual(['A', 'B']);
  });

  it('保存済みが空配列なら「予定0件」として正常に返す（取り込み済みで本当に0件の場合）', async () => {
    const out = await readPresident({ DB: makeMockDB({ payload: JSON.stringify([]) }) });
    expect(out.status).toBe('ok');
    expect(out.rows).toEqual([]);
  });

  it('保存済みが配列でない（形が違う）ときはエラー', async () => {
    const out = await readPresident({ DB: makeMockDB({ payload: JSON.stringify({ rows: [] }) }) });
    expect(out.status).toBe('error');
  });
});

describe('readPresident（社員用と混ざっていないこと）', () => {
  it('社員用の snapshot / sync_log を一切読まない', async () => {
    // makeMockDB は社員用テーブルへの問い合わせで例外を投げる
    const out = await readPresident({ DB: makeMockDB({ payload: JSON.stringify([makeEvent()]) }) });
    expect(out.status).toBe('ok');
  });
});
