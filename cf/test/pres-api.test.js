import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import worker from '../src/index.js';

function makeEvent(i) {
  return {
    '登録日時': '2026-08-01T00:00:00.000Z', 'タイトル': '予定' + i,
    '開始日': '2026-08-20', '開始時刻': '10:00',
    '終了日': '2026-08-20', '終了時刻': '11:00',
    '場所': '', 'メモ': '', 'カテゴリ': '', '色': '',
    'ID': 'P00000000' + i, '更新者': ''
  };
}

function makeDB({ presPayload = null, presLog = null, touched = [] } = {}) {
  const log = presLog != null ? presLog : [{ at: new Date().toISOString(), ok: 1 }];
  return {
    touched,
    prepare(sql) {
      touched.push(sql);
      const api = {
        bind: () => api,
        all: async () => {
          if (/SELECT payload FROM pres_snapshot/.test(sql)) {
            return { results: presPayload != null ? [{ payload: presPayload }] : [] };
          }
          if (/FROM pres_sync_log WHERE ok = 1/.test(sql)) {
            const oks = log.filter(l => Number(l.ok) === 1).sort((a, b) => b.at.localeCompare(a.at));
            return { results: oks.length ? [{ at: oks[0].at }] : [] };
          }
          if (/COUNT\(\*\) AS c FROM pres_sync_log/.test(sql)) {
            return { results: [{ c: 0 }] };
          }
          return { results: [] };
        },
        run: async () => ({ meta: { changes: 1 } })
      };
      return api;
    }
  };
}

const post = (path, body) => new Request('https://w.test' + path, {
  method: 'POST', headers: { 'Content-Type': 'application/json' },
  body: JSON.stringify(body)
});
const ctx = { waitUntil() {} };

let realFetch;
beforeEach(() => { realFetch = globalThis.fetch; });
afterEach(() => { globalThis.fetch = realFetch; vi.restoreAllMocks(); });

describe('POST /api/president（PIN照合）', () => {
  it('正しいPINなら予定を返す', async () => {
    const db = makeDB({ presPayload: JSON.stringify([makeEvent(1)]) });
    const res = await worker.fetch(post('/api/president', { pin: '1203' }), { DB: db, PRES_PIN: '1203' }, ctx);
    const body = await res.json();
    expect(res.status).toBe(200);
    expect(body.status).toBe('ok');
    expect(body.rows).toHaveLength(1);
  });

  it('PINが違えば403で、予定を1件も返さない', async () => {
    const db = makeDB({ presPayload: JSON.stringify([makeEvent(1)]) });
    const res = await worker.fetch(post('/api/president', { pin: '9999' }), { DB: db, PRES_PIN: '1203' }, ctx);
    const body = await res.json();
    expect(res.status).toBe(403);
    expect(body.status).toBe('error');
    expect(body.rows).toBeUndefined();
  });

  it('★PINが違うときはD1に一切問い合わせない（無料枠を消費させない）', async () => {
    const touched = [];
    const db = makeDB({ presPayload: JSON.stringify([makeEvent(1)]), touched });
    await worker.fetch(post('/api/president', { pin: '9999' }), { DB: db, PRES_PIN: '1203' }, ctx);
    expect(touched).toEqual([]);
  });

  it('PINが空・未指定でも403', async () => {
    const db = makeDB({ presPayload: JSON.stringify([makeEvent(1)]) });
    for (const body of [{}, { pin: '' }, { pin: null }]) {
      const res = await worker.fetch(post('/api/president', body), { DB: db, PRES_PIN: '1203' }, ctx);
      expect(res.status).toBe(403);
    }
  });

  it('★PRES_PINが未設定のときは誰も通さない（空文字どうしの一致で素通りしない）', async () => {
    const db = makeDB({ presPayload: JSON.stringify([makeEvent(1)]) });
    for (const body of [{ pin: '' }, {}, { pin: '1203' }]) {
      const res = await worker.fetch(post('/api/president', body), { DB: db, PRES_PIN: '' }, ctx);
      expect(res.status).toBe(503);
      const j = await res.json();
      expect(j.rows).toBeUndefined();
    }
  });

  it('本文がJSONでなくても落ちずに403', async () => {
    const db = makeDB({ presPayload: JSON.stringify([makeEvent(1)]) });
    const req = new Request('https://w.test/api/president', {
      method: 'POST', headers: { 'Content-Type': 'application/json' }, body: 'これはJSONではない'
    });
    const res = await worker.fetch(req, { DB: db, PRES_PIN: '1203' }, ctx);
    expect(res.status).toBe(403);
  });

  it('GETでは応答しない（PINをURLに載せる経路を作らない）', async () => {
    const db = makeDB({ presPayload: JSON.stringify([makeEvent(1)]) });
    const res = await worker.fetch(new Request('https://w.test/api/president?pin=1203'), { DB: db, PRES_PIN: '1203' }, ctx);
    expect(res.status).toBe(404);
  });

  it('未取り込みならstatus:errorを返す（画面はGASへ落ちる）', async () => {
    const db = makeDB({ presPayload: null });
    const res = await worker.fetch(post('/api/president', { pin: '1203' }), { DB: db, PRES_PIN: '1203' }, ctx);
    const body = await res.json();
    expect(body.status).toBe('error');
  });
});

describe('POST /api/pres-sync', () => {
  it('PINが違えば403で、GASを一度も叩かない', async () => {
    globalThis.fetch = vi.fn(async () => { throw new Error('叩いてはいけない'); });
    const db = makeDB();
    const res = await worker.fetch(post('/api/pres-sync', { pin: '9999' }), { DB: db, PRES_PIN: '1203' }, ctx);
    expect(res.status).toBe(403);
    expect(globalThis.fetch).not.toHaveBeenCalled();
  });

  it('正しいPINなら取り込んで結果を返す', async () => {
    globalThis.fetch = vi.fn(async () => new Response(
      JSON.stringify({ status: 'ok', rows: [makeEvent(1), makeEvent(2)] }), { status: 200 }));
    const db = makeDB();
    const res = await worker.fetch(post('/api/pres-sync', { pin: '1203' }),
      { DB: db, PRES_PIN: '1203', GAS_URL: 'https://gas.test/exec' }, ctx);
    const body = await res.json();
    expect(body.status).toBe('ok');
    expect(body.rows).toBe(2);
  });
});

describe('社員用のAPIを壊していないこと', () => {
  it('/api/schedule は社長用の変更後も従来どおり動く', async () => {
    const payload = JSON.stringify({
      compact: 1,
      headers: ['登録日時','作業日','元請名','現場名','氏名','役割','出勤','退勤','人工','メモ','夜勤','会社','ID','更新者','色','事業部','工番','作業区分','車両'],
      rows: [], members: [], genbaMaster: [], jobsites: []
    });
    const db = {
      prepare(sql) {
        return {
          bind: function () { return this; },
          all: async () => {
            if (/SELECT payload FROM snapshot/.test(sql)) return { results: [{ payload }] };
            if (/SELECT at FROM sync_log WHERE ok = 1/.test(sql)) return { results: [{ at: new Date().toISOString() }] };
            return { results: [] };
          },
          run: async () => ({ meta: { changes: 1 } })
        };
      }
    };
    const res = await worker.fetch(new Request('https://w.test/api/schedule?company='), { DB: db }, ctx);
    const body = await res.json();
    expect(body.status).toBe('ok');
    expect(body.compact).toBe(1);
  });
});
