// 集計を「本社／関東支店」で分ける。
//
// ★利用者指示（2026-08-28）:
//   「カレンダーは一つにして、関東支店と本社の売上とか割を分けれるようにしておいてくれたらよくて」
//
// ★なぜ必要か（2026-08-28 実測）:
//   集計を作る元データ（sheetToRecords）に拠点が入っておらず、
//   集計シート5枚すべてが拠点を1回も読んでいなかった。
//   今 本社と関東を分けられているのは「GRミツマが別の会社として残っているから」だけで、
//   GRミツマを消した瞬間に分けられなくなる。
import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import vm from 'node:vm';

const here = dirname(fileURLToPath(import.meta.url));
const CODE = readFileSync(join(here, '..', '..', 'gas.js'), 'utf8');

const EXPORT = `
;globalThis.__gas = { summaryGroupKey_, hasKyotenAxis_, defaultKyotenForCompany_ };
`;

function load() {
  const sandbox = vm.createContext({
    SpreadsheetApp: {}, Session: { getScriptTimeZone: () => 'Asia/Tokyo' },
    LockService: {}, Utilities: {}, ContentService: {}, UrlFetchApp: {},
    PropertiesService: {}, Logger: { log() {} }, console
  });
  vm.runInContext(CODE + EXPORT, sandbox, { filename: 'gas.js' });
  return sandbox.__gas;
}

const g = load();

describe('集計の見出し（会社＋拠点）', () => {
  it('★グローライズは 本社／関東支店 に分かれる', () => {
    expect(g.summaryGroupKey_({ company: 'グローライズ', kyoten: '本社' })).toBe('グローライズ（本社）');
    expect(g.summaryGroupKey_({ company: 'グローライズ', kyoten: '関東支店' })).toBe('グローライズ（関東支店）');
  });

  it('★GRミツマも拠点で分かれる（統合の前後どちらでも同じ数字が取れる）', () => {
    expect(g.summaryGroupKey_({ company: 'GRミツマ', kyoten: '関東支店' })).toBe('GRミツマ（関東支店）');
  });

  it('★和信カインド・ラーテル・GRHD は拠点の軸が無いのでそのまま', () => {
    ['和信カインド', 'ラーテル', 'GRHD'].forEach(co => {
      expect(g.summaryGroupKey_({ company: co, kyoten: '' }), co).toBe(co);
      // 拠点欄に何か入っていても、拠点の軸を持たない会社は分けない
      expect(g.summaryGroupKey_({ company: co, kyoten: '関東支店' }), co).toBe(co);
    });
  });

  it('拠点が空のグローライズ行は 本社 とみなす（画面の既定と揃える）', () => {
    expect(g.summaryGroupKey_({ company: 'グローライズ', kyoten: '' })).toBe('グローライズ（本社）');
  });

  it('拠点が空のGRミツマ行は 関東支店 とみなす', () => {
    expect(g.summaryGroupKey_({ company: 'GRミツマ', kyoten: '' })).toBe('GRミツマ（関東支店）');
  });

  it('「両方」の予定はその名前のまま出す（勝手に片方へ寄せない）', () => {
    expect(g.summaryGroupKey_({ company: 'グローライズ', kyoten: '両方' })).toBe('グローライズ（両方）');
  });

  it('前後の空白は無視する', () => {
    expect(g.summaryGroupKey_({ company: ' グローライズ ', kyoten: ' 関東支店 ' })).toBe('グローライズ（関東支店）');
  });

  it('会社が空でも落ちない', () => {
    expect(g.summaryGroupKey_({})).toBe('');
    expect(g.summaryGroupKey_(null)).toBe('');
  });
});

describe('集計の元データに拠点が入っていること', () => {
  it('★sheetToRecords が拠点を読んでいる（これが無いと分けられない）', () => {
    expect(CODE).toContain("kyoten: String(row[colIdx['拠点']] || '')");
  });

  it('★会社別集計が summaryGroupKey_ を使っている', () => {
    const m = CODE.match(/function generateCompanySummary_\(ss, records\) \{[\s\S]*?\n\}/);
    expect(m).toBeTruthy();
    expect(m[0]).toContain('summaryGroupKey_');
    // 会社だけで分けていた古い書き方が残っていないこと
    expect(m[0]).not.toContain("records.map(r => r.company)");
  });
});
