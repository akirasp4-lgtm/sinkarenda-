// 候補者を出す（依頼文の要件5）。2026-08-29。
//
// ★依頼文: 「案件に必要な人数・資格・経験を入力すると…候補者を提案する。
//   ただしAIが勝手に予定確定しない。最終決定は管理者が行う。」
//
// ★欲しいのは候補者のリスト。「空き × 資格 × その元請の経験」で出せるので
//   AIは使っていない＝0円（利用者判断 2026-08-29）。
import { describe, it, expect, beforeAll } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import vm from 'node:vm';

const here = dirname(fileURLToPath(import.meta.url));
const read = (f) => readFileSync(join(here, '..', '..', f), 'utf8');

const BEGIN = '// ===== PHASE5-PICK-RULE:BEGIN =====';
const END = '// ===== PHASE5-PICK-RULE:END =====';
function extract(file) {
  const src = read(file);
  const i = src.indexOf(BEGIN), j = src.indexOf(END);
  if (i < 0 || j < 0) throw new Error(file + ' に候補者のルールブロックが無い');
  return src.slice(i + BEGIN.length, j);
}

const EXPORT = ';globalThis.__p5 = { experienceDays, experienceGenbaChoices, rankCandidates };';

let P;
beforeAll(() => {
  const sandbox = vm.createContext({ console, String, Object });
  sandbox.globalThis = sandbox;
  vm.runInContext(extract('index.html') + EXPORT, sandbox, { filename: 'index.html' });
  P = sandbox.__p5;
});

const n = (o) => Object.assign({
  date: '2026-08-01', name: 'A', genba: 'きんでん東', loc: 'X',
  yasumi: '', yotei: '', isGhost: false
}, o);

describe('2つの画面で1文字も違わないこと', () => {
  it('index.html と admin.html が完全に同じ', () => {
    expect(extract('admin.html')).toBe(extract('index.html'));
  });
});

describe('その元請での経験（日数）', () => {
  it('働いた日数を数える', () => {
    expect(P.experienceDays([
      n({ date: '2026-08-01' }), n({ date: '2026-08-02' })], 'A', 'きんでん東')).toBe(2);
  });
  it('★同じ日に何件あっても1日と数える', () => {
    expect(P.experienceDays([
      n({ date: '2026-08-01', loc: 'X' }), n({ date: '2026-08-01', loc: 'Y' })], 'A', 'きんでん東')).toBe(1);
  });
  it('別の元請は数えない', () => {
    expect(P.experienceDays([n({ genba: 'ハイテックス' })], 'A', 'きんでん東')).toBe(0);
  });
  it('★休み・📌予定・ゴーストは経験に数えない', () => {
    expect(P.experienceDays([
      n({ date: '2026-08-01', yasumi: '○' }),
      n({ date: '2026-08-02', yotei: '○' }),
      n({ date: '2026-08-03', isGhost: true })], 'A', 'きんでん東')).toBe(0);
  });
  it('氏名か元請が空なら0', () => {
    expect(P.experienceDays([n({})], '', 'きんでん東')).toBe(0);
    expect(P.experienceDays([n({})], 'A', '')).toBe(0);
    expect(P.experienceDays(null, 'A', 'きんでん東')).toBe(0);
  });
});

describe('候補者の並べ替え', () => {
  const rows = [
    n({ name: 'A', date: '2026-08-01' }), n({ name: 'A', date: '2026-08-02' }),
    n({ name: 'A', date: '2026-08-03' }),
    n({ name: 'B', date: '2026-08-01' }),
    n({ name: 'C', genba: 'ハイテックス', date: '2026-08-01' })
  ];
  it('★経験が多い順に並ぶ', () => {
    const r = P.rankCandidates(['C', 'B', 'A'], rows, 'きんでん東');
    expect(r.map(x => x.name)).toEqual(['A', 'B', 'C']);
    expect(r.map(x => x.days)).toEqual([3, 1, 0]);
  });
  it('★同じ日数なら元の並び（名簿の順）を保つ', () => {
    const r = P.rankCandidates(['C', 'B'], [n({ name: 'B' }), n({ name: 'C' })], 'きんでん東');
    expect(r.map(x => x.name)).toEqual(['C', 'B']);
  });
  it('★元請を選んでいないときは並べ替えない（勝手に順番を変えない）', () => {
    const r = P.rankCandidates(['C', 'B', 'A'], rows, '');
    expect(r.map(x => x.name)).toEqual(['C', 'B', 'A']);
  });
  it('空でも落ちない', () => {
    expect(P.rankCandidates([], [], 'きんでん東')).toEqual([]);
    expect(P.rankCandidates(null, null, 'きんでん東')).toEqual([]);
  });
});

describe('元請のプルダウン', () => {
  it('★名簿に載っている人の実績だけ数える', () => {
    const list = P.experienceGenbaChoices([
      n({ name: 'A', genba: 'きんでん東' }),
      n({ name: '辞めた人', genba: 'ハイテックス' })], ['A']);
    expect(list).toEqual(['きんでん東']);
  });
  it('件数が多い順、同数なら文字順', () => {
    const list = P.experienceGenbaChoices([
      n({ name: 'A', genba: 'あ' }), n({ name: 'A', genba: 'あ' }),
      n({ name: 'A', genba: 'い' }), n({ name: 'A', genba: 'う' })], ['A']);
    expect(list).toEqual(['あ', 'い', 'う']);
  });
  it('{name,company} の形の名簿でも動く', () => {
    expect(P.experienceGenbaChoices([n({ name: 'A' })],
      [{ name: 'A', company: 'グローライズ' }])).toEqual(['きんでん東']);
  });
  it('空でも落ちない', () => {
    expect(P.experienceGenbaChoices([], [])).toEqual([]);
    expect(P.experienceGenbaChoices(null, null)).toEqual([]);
  });
});
