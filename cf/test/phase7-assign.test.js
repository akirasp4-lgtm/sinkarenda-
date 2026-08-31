// Phase 3: 現場の条件で候補者を絞る（2026-08-31）
//
// 社長指示 §7:
//   「AI人員配置提案は、プログラム側で条件判定を行い、
//     AIには順位付けと理由説明だけをさせる。氏名の記号化は維持する」
//
// ★このファイルが守る一番大事なこと:
//   資格がまだ登録されていないだけの人を、候補から消さないこと。
//   資格マスタに1行でも載っているのは62人中22人（2026-08-31 実測）。
//   消すと40人が永久に候補に出てこなくなる。
import { describe, it, expect, beforeAll } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import vm from 'node:vm';

const here = dirname(fileURLToPath(import.meta.url));
const read = (f) => readFileSync(join(here, '..', '..', f), 'utf8');

const cut = (src, name) => {
  const B = `// ===== ${name}:BEGIN =====`;
  const E = `// ===== ${name}:END =====`;
  const i = src.indexOf(B), j = src.indexOf(E);
  if (i < 0 || j < 0) throw new Error(name + ' が無い');
  return src.slice(i + B.length, j);
};

let S;
beforeAll(() => {
  const src = read('index.html');
  const box = vm.createContext({
    console, Map, Set, String, Object, Array, Number, Boolean, Date, RegExp, isFinite, Math
  });
  box.globalThis = box;
  // 資格の判定を使うので、そのブロックも一緒に読み込む
  vm.runInContext(
    cut(src, 'PHASE3-QUAL-RULE')
    + '\n;\n' + cut(src, 'PHASE7-ASSIGN-RULE')
    + '\n;\nglobalThis.__a = { siteNeedOf, screenByQuals, assignProgress, hasAnyQualRecord, qualIndexBy };',
    box, { filename: 'index.html' });
  S = box.__a;
});

describe('★2つの画面で1文字も違わない', () => {
  it('PHASE7-ASSIGN-RULE が index.html と admin.html で同じ', () => {
    expect(cut(read('admin.html'), 'PHASE7-ASSIGN-RULE'))
      .toBe(cut(read('index.html'), 'PHASE7-ASSIGN-RULE'));
  });
});

// ================================================================ siteNeedOf

const site = (o) => Object.assign({
  genba: 'きんでん東', loc: 'A現場', status: '施工中',
  needCount: null, needQuals: [], needExp: '', address: '', startAt: '', endAt: ''
}, o || {});

describe('現場の条件を引く（siteNeedOf）', () => {
  it('元請名と現場名の両方で引く', () => {
    const r = S.siteNeedOf([site({ needCount: 4 }), site({ loc: 'B現場', needCount: 9 })],
      'きんでん東', 'B現場');
    expect(r.needCount).toBe(9);
  });

  it('★見つからなければ null（空の条件を返さない）', () => {
    expect(S.siteNeedOf([site()], 'エクシオ', 'A現場')).toBe(null);
    expect(S.siteNeedOf([site()], 'きんでん東', 'Z現場')).toBe(null);
    expect(S.siteNeedOf([], 'きんでん東', 'A現場')).toBe(null);
    expect(S.siteNeedOf(null, 'きんでん東', 'A現場')).toBe(null);
  });

  it('元請名が空なら引かない', () => {
    expect(S.siteNeedOf([site()], '', 'A現場')).toBe(null);
  });

  it('現場名が空の現場も引ける（元請だけの行がある）', () => {
    const r = S.siteNeedOf([site({ loc: '', needCount: 2 })], 'きんでん東', '');
    expect(r.needCount).toBe(2);
  });

  it('★必要人数の 0・マイナス・文字は「未登録」にする', () => {
    [0, -3, 'あ', '', null, undefined].forEach((v) => {
      expect(S.siteNeedOf([site({ needCount: v })], 'きんでん東', 'A現場').needCount,
        '値=' + JSON.stringify(v)).toBe(null);
    });
    expect(S.siteNeedOf([site({ needCount: '4' })], 'きんでん東', 'A現場').needCount).toBe(4);
    expect(S.siteNeedOf([site({ needCount: 4.7 })], 'きんでん東', 'A現場').needCount).toBe(4);
  });

  it('必要資格は配列でも「、」区切りの文字でも読める', () => {
    expect(S.siteNeedOf([site({ needQuals: ['玉掛け', '高所'] })], 'きんでん東', 'A現場').needQuals)
      .toEqual(['玉掛け', '高所']);
    expect(S.siteNeedOf([site({ needQuals: '玉掛け、高所' })], 'きんでん東', 'A現場').needQuals)
      .toEqual(['玉掛け', '高所']);
    expect(S.siteNeedOf([site({ needQuals: '' })], 'きんでん東', 'A現場').needQuals).toEqual([]);
  });

  it('必要資格の空欄は落とす', () => {
    expect(S.siteNeedOf([site({ needQuals: ['玉掛け', '', '  '] })], 'きんでん東', 'A現場').needQuals)
      .toEqual(['玉掛け']);
  });

  it('前後の空白を無視して引ける', () => {
    const r = S.siteNeedOf([site({ needCount: 3 })], '  きんでん東  ', ' A現場 ');
    expect(r.needCount).toBe(3);
  });

  it('住所・時間・必要経験もそのまま返る', () => {
    const r = S.siteNeedOf([site({
      needExp: '楽天案件経験', address: '大阪市北区1-1', startAt: '08:00', endAt: '17:00'
    })], 'きんでん東', 'A現場');
    expect(r.needExp).toBe('楽天案件経験');
    expect(r.address).toBe('大阪市北区1-1');
    expect(r.startAt).toBe('08:00');
    expect(r.endAt).toBe('17:00');
  });
});

// ================================================================ screenByQuals

const q = (name, qual, expires, company) => ({
  name, qual, expires: expires === undefined ? '' : expires, company: company || 'グローライズ'
});
const M = (name, company) => ({ name, company: company || 'グローライズ' });
const TODAY = '2026-09-09';

describe('必要資格で名簿を3つに分ける（screenByQuals）', () => {
  it('必要資格が空なら全員が「満たす」', () => {
    const r = S.screenByQuals([M('江頭'), M('河原')], {}, [], TODAY);
    expect(r.ok.map((x) => x.name)).toEqual(['江頭', '河原']);
    expect(r.ng).toEqual([]);
    expect(r.unknown).toEqual([]);
  });

  it('必要資格を全部持っていれば「満たす」', () => {
    const idx = S.qualIndexBy([q('江頭', '玉掛け'), q('江頭', '高所作業車')]);
    const r = S.screenByQuals([M('江頭')], idx, ['玉掛け', '高所作業車'], TODAY);
    expect(r.ok.map((x) => x.name)).toEqual(['江頭']);
  });

  it('1つでも足りなければ「足りない」。何が足りないかを返す', () => {
    const idx = S.qualIndexBy([q('江頭', '玉掛け')]);
    const r = S.screenByQuals([M('江頭')], idx, ['玉掛け', '高所作業車'], TODAY);
    expect(r.ok).toEqual([]);
    expect(r.ng).toHaveLength(1);
    expect(r.ng[0].missing).toEqual(['高所作業車']);
  });

  it('★★資格が1件も登録されていない人は「判定できない」（候補から消さない）', () => {
    // ここを間違えると、資格を入力していない40人が永久に候補から消える
    const idx = S.qualIndexBy([q('江頭', '玉掛け')]);
    const r = S.screenByQuals([M('江頭'), M('河原')], idx, ['玉掛け'], TODAY);
    expect(r.ok.map((x) => x.name)).toEqual(['江頭']);
    expect(r.ng).toEqual([]);
    expect(r.unknown.map((x) => x.name), '★資格未登録の人が消えた').toEqual(['河原']);
  });

  it('★誰も落とさない（3つの合計が名簿と同じ）', () => {
    const idx = S.qualIndexBy([q('江頭', '玉掛け'), q('前﨑', '高所作業車')]);
    const members = [M('江頭'), M('河原'), M('前﨑'), M('真柄')];
    const r = S.screenByQuals(members, idx, ['玉掛け'], TODAY);
    expect(r.ok.length + r.ng.length + r.unknown.length).toBe(members.length);
  });

  it('期限が切れた資格は「持っている」に数えない', () => {
    const idx = S.qualIndexBy([q('江頭', '玉掛け', '2024-01-01')]);
    const r = S.screenByQuals([M('江頭')], idx, ['玉掛け'], TODAY);
    expect(r.ng).toHaveLength(1);
    expect(r.ng[0].missing).toEqual(['玉掛け']);
  });

  it('期限が近いだけならまだ持っている（現場には出られる）', () => {
    const idx = S.qualIndexBy([q('江頭', '玉掛け', '2026-10-01')]);
    expect(S.screenByQuals([M('江頭')], idx, ['玉掛け'], TODAY).ok).toHaveLength(1);
  });

  it('期限のない資格（技能講習など）は持っている', () => {
    const idx = S.qualIndexBy([q('江頭', '玉掛け', '')]);
    expect(S.screenByQuals([M('江頭')], idx, ['玉掛け'], TODAY).ok).toHaveLength(1);
  });

  it('★有効期限が読めない資格は「持っている」に数えない（安全側）', () => {
    const idx = S.qualIndexBy([q('江頭', '玉掛け', '未定')]);
    const r = S.screenByQuals([M('江頭')], idx, ['玉掛け'], TODAY);
    // 資格の行はあるので unknown ではなく ng（足りない）に落ちる
    expect(r.unknown).toEqual([]);
    expect(r.ng[0].missing).toEqual(['玉掛け']);
  });

  it('★他社の同姓の資格を使わない（奥田さんは2社に実在する）', () => {
    const idx = S.qualIndexBy([q('奥田', '玉掛け', '', 'GRHD')]);
    const r = S.screenByQuals([M('奥田', 'グローライズ')], idx, ['玉掛け'], TODAY);
    expect(r.ok).toEqual([]);
    expect(r.unknown.map((x) => x.name), '他社の資格で有資格になった').toEqual(['奥田']);
  });

  it('グローライズとGRミツマは同じ束で見る（統合前の行が残っている）', () => {
    const idx = S.qualIndexBy([q('江頭', '玉掛け', '', 'GRミツマ')]);
    expect(S.screenByQuals([M('江頭', 'グローライズ')], idx, ['玉掛け'], TODAY).ok)
      .toHaveLength(1);
  });

  it('氏名が空の行は無視する', () => {
    const r = S.screenByQuals([M(''), M('  '), M('江頭')], {}, [], TODAY);
    expect(r.ok).toHaveLength(1);
  });

  it('名簿が空でも落ちない', () => {
    expect(S.screenByQuals([], {}, ['玉掛け'], TODAY))
      .toEqual({ ok: [], ng: [], unknown: [] });
    expect(S.screenByQuals(null, null, null, TODAY))
      .toEqual({ ok: [], ng: [], unknown: [] });
  });
});

// ================================================================ assignProgress

describe('必要人数に対して何人選べたか（assignProgress）', () => {
  it('足りないときは「あと何人」を出す', () => {
    const r = S.assignProgress(4, 2);
    expect(r).toMatchObject({ need: 4, picked: 2, rest: 2, filled: false });
    expect(r.label).toBe('必要4人 / 選択2人 → あと2人');
  });

  it('ちょうどなら足りている', () => {
    expect(S.assignProgress(4, 4)).toMatchObject({ rest: 0, filled: true });
  });

  it('多く選んでも「あと-1人」にしない', () => {
    const r = S.assignProgress(4, 6);
    expect(r.rest).toBe(0);
    expect(r.filled).toBe(true);
    expect(r.label).toContain('足りています');
  });

  it('★必要人数が未登録なら「未登録」と言う（0人必要と混ぜない）', () => {
    [null, '', undefined, 0, -1, 'あ'].forEach((v) => {
      const r = S.assignProgress(v, 3);
      expect(r.need, '値=' + JSON.stringify(v)).toBe(null);
      expect(r.rest).toBe(null);
      expect(r.filled).toBe(false);
      expect(r.label).toBe('必要人数は未登録');
    });
  });

  it('選択0人でも落ちない', () => {
    expect(S.assignProgress(3, 0).label).toBe('必要3人 / 選択0人 → あと3人');
    expect(S.assignProgress(3, null).picked).toBe(0);
  });
});
