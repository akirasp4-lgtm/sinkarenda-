// 資格マスタを画面へ届ける（フェーズ3の土台）。
//
// ★利用者指示（2026-08-28）:
//   「資格書は今からお前に教えるから、それで割り振ってくれる？」
//
// ★設計書 docs/superpowers/specs/2026-08-27-zensha-jinin-haichi-design.md では
//   「職人マスタに資格列を足して71人分を人が入力する」が前提だったが、
//   NASの資格者証一覧から303件を取り込んで『資格マスタ』シートに入れたので、
//   人の入力作業は不要になった。あとは画面へ届けるだけ。
//
// ★個人情報の扱い（ここが一番大事）:
//   資格マスタには 免許番号・正式氏名・取得日・出典 が入っている。
//   現場画面(index.html)はPINが無く全社員が使い、内容はD1と端末のlocalStorageにも残る。
//   そのため **GASを出る時点で 氏名/会社/資格名/区分/有効期限 だけに削る**。
//   （会社は「和信カインドの画面にグローライズの資格を出さない」ための絞り込みに要る）
//   単価(rate)をWorkerで落としているのと同じ考え方だが、資格は
//   「そもそもGASから出さない」＝もっと手前で止める。
import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import vm from 'node:vm';

const here = dirname(fileURLToPath(import.meta.url));
const CODE = readFileSync(join(here, '..', '..', 'gas.js'), 'utf8');

const EXPORT = `
;globalThis.__gas = { normalizeQualDate_, projectQualifications_, QUAL_SHEET };
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

const HEAD = ['氏名', '会社', '正式氏名', '資格名', '区分', '免許番号', '取得日', '有効期限', '出典'];
function row(name, co, seishiki, qual, kind, no, got, exp, src) {
  return [name, co, seishiki, qual, kind, no, got, exp, src];
}

describe('資格の有効期限の読み取り', () => {
  it('YYYY-MM-DD はそのまま通る', () => {
    expect(g.normalizeQualDate_('2029-01-17')).toBe('2029-01-17');
  });
  it('YYYY/M/D は0埋めして揃える', () => {
    expect(g.normalizeQualDate_('2029/1/7')).toBe('2029-01-07');
    expect(g.normalizeQualDate_('2029/11/17')).toBe('2029-11-17');
  });
  it('★「-」や空欄は「期限なし」として空文字にする（資格証の実データにある）', () => {
    expect(g.normalizeQualDate_('-')).toBe('');
    expect(g.normalizeQualDate_('')).toBe('');
    expect(g.normalizeQualDate_(null)).toBe('');
    expect(g.normalizeQualDate_(undefined)).toBe('');
    expect(g.normalizeQualDate_('   ')).toBe('');
  });
  it('★日付として入っていたセル（Dateオブジェクト）も読める', () => {
    expect(g.normalizeQualDate_(new Date(2029, 0, 17))).toBe('2029-01-17');
  });
  it("★[P1] 読めない文字は '?' にする。空欄（期限なし）と混ぜない", () => {
    // ★Codexレビュー[P1]（2026-08-28）: 以前は '' にしていた。
    //   '' は画面側で「期限なし＝切れない」になるため、読めない資格が
    //   一生有効な資格に化けていた。
    expect(g.normalizeQualDate_('平成31年')).toBe('?');
    expect(g.normalizeQualDate_('20290117')).toBe('?');
    expect(g.normalizeQualDate_('2029-1-7-')).toBe('?');
  });
  it('★存在しない日も通さない（Dateに任せると 2/31 が 3/3 になる）', () => {
    expect(g.normalizeQualDate_('2026-02-31')).toBe('?');
    expect(g.normalizeQualDate_('2026/2/31')).toBe('?');
    expect(g.normalizeQualDate_('2026-13-01')).toBe('?');
    expect(g.normalizeQualDate_('2027/2/29')).toBe('?');   // 平年
    expect(g.normalizeQualDate_('2028/2/29')).toBe('2028-02-29');   // うるう年は通す
  });
  it('★空欄と「-」だけが「期限なし」', () => {
    expect(g.normalizeQualDate_('')).toBe('');
    expect(g.normalizeQualDate_('-')).toBe('');
    expect(g.normalizeQualDate_('   ')).toBe('');
  });
});

describe('資格マスタの投影（画面へ出す列だけに削る）', () => {
  const data = [
    HEAD,
    row('真柄', 'グローライズ', '真柄　静志', '高所作業車運転技能講習', '技能講習', '第19-00674号', '2019-06-26', '', 'x.xlsx／真柄'),
    row('河原', 'グローライズ', '河原　将司', '第一種電気工事士', '国家資格', '03569', '1991-01-24', '2029-01-17', 'x.xlsx／河原')
  ];

  it('★免許番号・正式氏名・取得日・出典は1文字も出さない', () => {
    const out = g.projectQualifications_(data, false, '');
    const json = JSON.stringify(out);
    expect(json).not.toContain('第19-00674号');
    expect(json).not.toContain('03569');
    expect(json).not.toContain('真柄　静志');
    expect(json).not.toContain('河原　将司');
    expect(json).not.toContain('x.xlsx');
    expect(json).not.toContain('1991-01-24');
  });

  it('出すのは 氏名/会社/資格名/区分/有効期限/取得場所 だけ（免許番号は出さない）', () => {
    // ★2026-08-29 取得場所を足した。「第一種工事検査員って何？」が
    //   取得場所（富士通ネットワークソリューションズ）で一発で解けたため。
    const out = g.projectQualifications_(data, false, '');
    expect(out).toHaveLength(2);
    expect(Object.keys(out[0]).sort()).toEqual(['company', 'expires', 'kind', 'name', 'place', 'qual']);
    expect(out[0]).toEqual({ name: '真柄', company: 'グローライズ', qual: '高所作業車運転技能講習',
      kind: '技能講習', expires: '', place: '' });
    expect(out[1]).toEqual({ name: '河原', company: 'グローライズ', qual: '第一種電気工事士',
      kind: '国家資格', expires: '2029-01-17', place: '' });
  });

  it('会社で絞れる（和信カインドの画面にグローライズの資格を出さない）', () => {
    const d = [HEAD,
      row('真柄', 'グローライズ', '', '玉掛け', '技能講習', '', '', '', ''),
      row('誰か', '和信カインド', '', 'フォークリフト', '技能講習', '', '', '', '')];
    const out = g.projectQualifications_(d, true, '和信カインド');
    expect(out).toEqual([{ name: '誰か', company: '和信カインド', qual: 'フォークリフト', kind: '技能講習', expires: '', place: '' }]);
  });

  it('氏名か資格名が空の行は捨てる', () => {
    const d = [HEAD,
      row('', 'グローライズ', '', '玉掛け', '技能講習', '', '', '', ''),
      row('真柄', 'グローライズ', '', '', '技能講習', '', '', '', ''),
      row('真柄', 'グローライズ', '', '玉掛け', '技能講習', '', '', '', '')];
    expect(g.projectQualifications_(d, false, '')).toEqual(
      [{ name: '真柄', company: 'グローライズ', qual: '玉掛け', kind: '技能講習', expires: '', place: '' }]);
  });

  it('★見出しの並びが変わっても名前で探す（列を足されても壊れない）', () => {
    const d = [
      ['会社', '資格名', '氏名', '有効期限', '区分', 'メモ'],
      ['グローライズ', '玉掛け', '真柄', '2030-03-31', '技能講習', 'なにか']
    ];
    expect(g.projectQualifications_(d, false, '')).toEqual(
      [{ name: '真柄', company: 'グローライズ', qual: '玉掛け', kind: '技能講習', expires: '2030-03-31', place: '' }]);
  });

  it('★必要な見出しが無ければ空で返す（doGet全体を巻き込んで落とさない）', () => {
    expect(g.projectQualifications_([['あ', 'い'], ['1', '2']], false, '')).toEqual([]);
    expect(g.projectQualifications_([], false, '')).toEqual([]);
    expect(g.projectQualifications_([HEAD], false, '')).toEqual([]);
  });

  it('シート名の定数は「資格マスタ」', () => {
    expect(g.QUAL_SHEET).toBe('資格マスタ');
  });
});

// ============================================================
// 実データに合わせた読み取り（2026-08-28・本番の303件を見て判明）
// ============================================================
describe('資格者証一覧の実際の書き方に対応する', () => {
  it("★'2029/3/31(合格した年から5年後の3/31)' は日付として読む（実データに6件ある）", () => {
    expect(g.normalizeQualDate_('2029/3/31(合格した年から5年後の3/31)')).toBe('2029-03-31');
    expect(g.normalizeQualDate_('2029/3/31（全角カッコ）')).toBe('2029-03-31');
    expect(g.normalizeQualDate_('2029-03-31 (メモ)')).toBe('2029-03-31');
  });
  it("★'発行日より3年' は読めないまま（推測で日付を作らない。実データに11件ある）", () => {
    expect(g.normalizeQualDate_('発行日より3年')).toBe('?');
  });
  it('★期間の書き方は読めないままにする（どちらの日か決められない）', () => {
    expect(g.normalizeQualDate_('2029/3/31〜2030/3/31')).toBe('?');
    expect(g.normalizeQualDate_('2029/3/31-2030/3/31')).toBe('?');
  });
  it('カッコを外した結果が日付でなければ読めない扱い', () => {
    expect(g.normalizeQualDate_('未定(要確認)')).toBe('?');
  });
});
