import { describe, it, expect } from 'vitest';
import { buildResponse, HEADERS, readSchedule } from '../src/read.js';

describe('buildResponse', () => {
  const row = {
    touroku:'2026-05-01T04:23:04.000Z', sagyoubi:'2026-05-02', motoukr:'NGS', genba:'大阪',
    shimei:'川端（達）', yakuwari:'代表', shukkin:'09:00', taikin:'18:00', kosu:1, memo:'',
    yakin:'', kaisha:'グローライズ', id:'abc-1', koushinsha:'森', iro:'#1D9E75',
    jigyoubu:'ICT', kouban:'INF-26-041', sagyou_kubun:'現場作業', sharyou:''
  };

  it('GASと同じ19個のヘッダを同じ順で返す', () => {
    expect(HEADERS).toEqual(['登録日時','作業日','元請名','現場名','氏名','役割','出勤','退勤',
      '人工','メモ','夜勤','会社','ID','更新者','色','事業部','工番','作業区分','車両']);
  });

  it('rowsはヘッダの順に並んだ値の配列になる', () => {
    const out = buildResponse([row], [], [], []);
    expect(out.status).toBe('ok');
    expect(out.compact).toBe(1);
    expect(out.rows[0][HEADERS.indexOf('作業日')]).toBe('2026-05-02');
    expect(out.rows[0][HEADERS.indexOf('ID')]).toBe('abc-1');
    expect(out.rows[0][HEADERS.indexOf('工番')]).toBe('INF-26-041');
    expect(out.rows[0]).toHaveLength(19);
  });

  it('職人マスタには単価を含めない', () => {
    const out = buildResponse([], [{name:'森',company:'GRHD',division:'ICT'}], [], []);
    expect(out.members[0]).toEqual({name:'森',company:'GRHD',division:'ICT'});
  });

  it('現場マスタのcompletedは真偽値に戻す（画面が真偽値で判定するため）', () => {
    const out = buildResponse([], [], [], [{genba:'A',loc:'B',jobNo:'',completed:1,billingMethod:'応援'}]);
    expect(out.jobsites[0].completed).toBe(true);
  });
});

// --- ここから追加（計画からの変更点1：sync_logの最新行を見て取り込み失敗/未取り込みをエラー扱いにする） ---

// D1のprepare/bind/all()を模した簡易モック。SQL文の中身で対象テーブルを見分け、
// あらかじめ渡した固定データを返す。readScheduleが投げるクエリの種類
// （sync_log / nippo(絞込あり・なし) / members(絞込あり・なし) / genba / jobsites）
// をすべてカバーする。
function makeMockDB({
  syncLogRows = [], nippoRows = [], memberRows = [], genbaRows = [], jobsiteRows = []
} = {}) {
  function runQuery(sql, args) {
    if (/FROM sync_log/.test(sql)) {
      return Promise.resolve({ results: syncLogRows });
    }
    if (/FROM nippo/.test(sql)) {
      const rows = /WHERE kaisha/.test(sql)
        ? nippoRows.filter(r => r.kaisha === args[0])
        : nippoRows;
      return Promise.resolve({ results: rows });
    }
    if (/FROM members/.test(sql)) {
      const rows = /WHERE company/.test(sql)
        ? memberRows.filter(r => r.company === args[0])
        : memberRows;
      return Promise.resolve({ results: rows });
    }
    if (/FROM genba/.test(sql)) {
      return Promise.resolve({ results: genbaRows });
    }
    if (/FROM jobsites/.test(sql)) {
      return Promise.resolve({ results: jobsiteRows });
    }
    return Promise.resolve({ results: [] });
  }

  return {
    prepare(sql) {
      return {
        bind(...args) {
          return { all: () => runQuery(sql, args) };
        },
        all: () => runQuery(sql, [])
      };
    }
  };
}

describe('readSchedule（sync_logの状態による安全装置）', () => {
  it('sync_logの最新行がok=0のとき、通常応答ではなくエラーを返す（一部だけ入った中途半端なデータを見せない）', async () => {
    const db = makeMockDB({
      syncLogRows: [{ at: '2026-08-22T00:00:00.000Z', rows: 0, ok: 0, message: 'mock batch failure' }],
      // 中途半端に入ったデータがあっても、それが返らないことを確認するために用意しておく
      nippoRows: [{ id: 'x-1', sagyoubi: '2026-05-02', shimei: '森', kaisha: 'グローライズ' }]
    });
    const env = { DB: db };

    const out = await readSchedule(env, '');

    expect(out.status).toBe('error');
    expect(out.rows).toBeUndefined();
    expect(out.compact).toBeUndefined();
  });

  it('sync_logが空（まだ一度も取り込んでいない）のときもエラーを返す（空のD1を「予定ゼロ件」として返さない）', async () => {
    const db = makeMockDB({ syncLogRows: [] });
    const env = { DB: db };

    const out = await readSchedule(env, '');

    expect(out.status).toBe('error');
    expect(out.rows).toBeUndefined();
  });

  it('sync_logの最新行がok=1のときは通常どおりGASと同じ形の応答を返す', async () => {
    const db = makeMockDB({
      syncLogRows: [{ at: '2026-08-22T00:00:00.000Z', rows: 1, ok: 1, message: '' }],
      nippoRows: [{
        touroku: '2026-05-01T04:23:04.000Z', sagyoubi: '2026-05-02', motoukr: 'NGS', genba: '大阪',
        shimei: '川端（達）', yakuwari: '代表', shukkin: '09:00', taikin: '18:00', kosu: 1, memo: '',
        yakin: '', kaisha: 'グローライズ', id: 'abc-1', koushinsha: '森', iro: '#1D9E75',
        jigyoubu: 'ICT', kouban: 'INF-26-041', sagyou_kubun: '現場作業', sharyou: ''
      }],
      memberRows: [{ name: '森', company: 'グローライズ', division: 'ICT' }],
      genbaRows: [{ name: '大阪', company: '' }],
      jobsiteRows: [{ genba: '大阪', loc: '本社', jobNo: '', completed: 0, billingMethod: '応援' }]
    });
    const env = { DB: db };

    const out = await readSchedule(env, '');

    expect(out.status).toBe('ok');
    expect(out.compact).toBe(1);
    expect(out.rows).toHaveLength(1);
    expect(out.rows[0][HEADERS.indexOf('ID')]).toBe('abc-1');
    expect(out.members).toEqual([{ name: '森', company: 'グローライズ', division: 'ICT' }]);
    expect(out.jobsites[0].completed).toBe(false);
  });

  it('sync_logがok=1でも、会社を指定すればnippoとmembersだけその会社で絞り込まれる', async () => {
    const db = makeMockDB({
      syncLogRows: [{ at: '2026-08-22T00:00:00.000Z', rows: 2, ok: 1, message: '' }],
      nippoRows: [
        { id: 'a-1', sagyoubi: '2026-05-02', shimei: '森', kaisha: 'グローライズ' },
        { id: 'b-1', sagyoubi: '2026-05-02', shimei: '田中', kaisha: '和信カインド' }
      ],
      memberRows: [
        { name: '森', company: 'グローライズ', division: 'ICT' },
        { name: '田中', company: '和信カインド', division: '' }
      ],
      genbaRows: [{ name: '大阪', company: 'グローライズ' }],
      jobsiteRows: [{ genba: '大阪', loc: '本社', jobNo: '', completed: 0, billingMethod: '応援' }]
    });
    const env = { DB: db };

    const out = await readSchedule(env, 'グローライズ');

    expect(out.status).toBe('ok');
    expect(out.rows).toHaveLength(1);
    expect(out.members).toHaveLength(1);
    expect(out.members[0].name).toBe('森');
  });
});
