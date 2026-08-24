import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';

describe('schema.sql', () => {
  const sql = readFileSync(new URL('../schema.sql', import.meta.url), 'utf8');

  it('日報テーブルに19列そろっている', () => {
    const cols = ['id','touroku','sagyoubi','motoukr','genba','shimei','yakuwari',
                  'shukkin','taikin','kosu','memo','yakin','kaisha','koushinsha',
                  'iro','jigyoubu','kouban','sagyou_kubun','sharyou'];
    for (const c of cols) expect(sql).toContain(c);
  });

  it('作業日に索引がある（期間検索を速くするため）', () => {
    expect(sql).toMatch(/CREATE INDEX .*nippo.*sagyoubi/i);
  });

  it('単価(rate)は保存しない（給料情報をD1へ持ち込まない）', () => {
    expect(sql).not.toMatch(/\brate\b/i);
  });

  // --- ここから追加（2026-08-24 設計変更の回帰防止：本番データ突き合わせで
  // 14行欠落＝車検期限リマインダー行の消失が発覚。原因は複合主キー
  // (id, sagyoubi, shimei) が氏名の空を許さなかったこと。D1はGAS応答の
  // 忠実な写しにする方針へ変更し、4テーブルすべてを連番seq主キーへ変えた）

  it('nippo/members/genba/jobsitesは連番seqをAUTOINCREMENTの主キーにしている', () => {
    // ★table名を直接カラム名(例: nippoの中のgenba列)と取り違えないよう、
    // "CREATE TABLE IF NOT EXISTS <table>" に直接続く定義だけを切り出す。
    for (const table of ['nippo', 'members', 'genba', 'jobsites']) {
      const m = sql.match(new RegExp(
        'CREATE TABLE IF NOT EXISTS\\s+' + table + '\\s*\\([\\s\\S]*?\\n\\);', 'i'));
      expect(m, table + 'のCREATE TABLE文が見つからない').toBeTruthy();
      expect(m[0]).toMatch(/seq\s+INTEGER\s+PRIMARY\s+KEY\s+AUTOINCREMENT/i);
    }
  });

  it('nippo/members/genba/jobsitesは複合主キー(PRIMARY KEY (...))を持たない（重複排除しないため）', () => {
    expect(sql).not.toMatch(/PRIMARY KEY\s*\(/i);
  });

  it('氏名(shimei)・作業日(sagyoubi)・IDにNOT NULL制約が無い（空が正当なデータのため。車検期限リマインダー行は氏名が空）', () => {
    for (const col of ['shimei', 'sagyoubi', 'id']) {
      const m = sql.match(new RegExp('^\\s*' + col + '\\s+TEXT.*$', 'im'));
      expect(m, col + '列の定義行が見つからない').toBeTruthy();
      expect(m[0]).not.toMatch(/NOT NULL/i);
    }
  });
});
