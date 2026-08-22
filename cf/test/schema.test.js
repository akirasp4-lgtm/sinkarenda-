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
});
