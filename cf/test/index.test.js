import { describe, it, expect } from 'vitest';
import worker from '../src/index.js';

// D1のprepare/bind/all()を模した簡易モック。bind()に渡された引数を記録し、
// company パラメータが .trim() された値でクエリに渡っているかを検証する。
function makeMockDB({ syncLogRows = [], nippoRows = [], memberRows = [], genbaRows = [], jobsiteRows = [] } = {}) {
  const bindCalls = []; // { sql, args } を記録

  function runQuery(sql, args) {
    if (/FROM sync_log/.test(sql)) return Promise.resolve({ results: syncLogRows });
    if (/FROM nippo/.test(sql)) {
      const rows = /WHERE kaisha/.test(sql) ? nippoRows.filter(r => r.kaisha === args[0]) : nippoRows;
      return Promise.resolve({ results: rows });
    }
    if (/FROM members/.test(sql)) {
      const rows = /WHERE company/.test(sql) ? memberRows.filter(r => r.company === args[0]) : memberRows;
      return Promise.resolve({ results: rows });
    }
    if (/FROM genba/.test(sql)) return Promise.resolve({ results: genbaRows });
    if (/FROM jobsites/.test(sql)) return Promise.resolve({ results: jobsiteRows });
    return Promise.resolve({ results: [] });
  }

  const db = {
    prepare(sql) {
      return {
        bind(...args) {
          bindCalls.push({ sql, args });
          return { all: () => runQuery(sql, args) };
        },
        all: () => runQuery(sql, [])
      };
    }
  };

  return { db, bindCalls };
}

describe('GET /api/schedule のcompanyパラメータのtrim（レビュー指摘: 会社名の前後空白）', () => {
  it('companyパラメータの前後に空白があっても、trimしてからD1へ問い合わせる', async () => {
    const { db, bindCalls } = makeMockDB({
      syncLogRows: [{ at: 'x', rows: 1, ok: 1, message: '' }],
      nippoRows: [{ id: 'a-1', sagyoubi: '2026-05-02', shimei: '森', kaisha: 'グローライズ' }],
      memberRows: [{ name: '森', company: 'グローライズ', division: 'ICT' }]
    });
    const env = { DB: db };

    // URLエンコードされた空白（%20と全角の　）を前後に含むcompanyパラメータ
    const req = new Request('https://worker.test/api/schedule?' +
      'company=' + encodeURIComponent('  グローライズ　'));
    const res = await worker.fetch(req, env, {});
    const body = await res.json();

    expect(res.status).toBe(200);
    expect(body.status).toBe('ok');
    // trimされた会社名で絞り込まれ、該当行が正しく返っていること
    expect(body.rows).toHaveLength(1);
    expect(body.members).toHaveLength(1);

    // D1へ渡された実際のbind引数もtrim済みであることを確認
    const nippoCall = bindCalls.find(c => /FROM nippo WHERE kaisha/.test(c.sql));
    const memberCall = bindCalls.find(c => /FROM members WHERE company/.test(c.sql));
    expect(nippoCall.args[0]).toBe('グローライズ');
    expect(memberCall.args[0]).toBe('グローライズ');
  });

  it('companyパラメータに空白が無い通常ケースでも従来どおり動く（回帰確認）', async () => {
    const { db } = makeMockDB({
      syncLogRows: [{ at: 'x', rows: 1, ok: 1, message: '' }],
      nippoRows: [{ id: 'a-1', sagyoubi: '2026-05-02', shimei: '森', kaisha: 'グローライズ' }],
      memberRows: [{ name: '森', company: 'グローライズ', division: 'ICT' }]
    });
    const env = { DB: db };

    const req = new Request('https://worker.test/api/schedule?company=' + encodeURIComponent('グローライズ'));
    const res = await worker.fetch(req, env, {});
    const body = await res.json();

    expect(body.status).toBe('ok');
    expect(body.rows).toHaveLength(1);
  });
});
