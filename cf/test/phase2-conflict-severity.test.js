// Phase 2: 重複チェックの3段階（2026-08-31）
//
// 社長指示 §6:
//   「250件を検知しない仕様にはしないでください。
//     ただし全部を同じ強さで警告する必要はありません。
//     高優先 / 要確認 / 参考 に分けてください。
//     重要なのは、検知を消すのではなく、検知した上で通知レベルを変えること。
//     『47件が本物』と断定する実装・表示にはしないでください。」
//
// ★守りたいこと:
//   ① 検知は広く（現場系でない予定も拾う）
//   ② 既存の画面の警告は今までどおり「高優先」だけ（既定値 high）
//   ③ 画面とWorkerで強さの判定が1つも食い違わない
import { describe, it, expect, beforeAll } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import vm from 'node:vm';
import * as W from '../src/alerts.js';

const here = dirname(fileURLToPath(import.meta.url));
const read = (f) => readFileSync(join(here, '..', '..', f), 'utf8');

const BEGIN = '// ===== PHASE2-CONFLICT-RULE:BEGIN =====';
const END = '// ===== PHASE2-CONFLICT-RULE:END =====';
function extract(file) {
  const src = read(file);
  const i = src.indexOf(BEGIN), j = src.indexOf(END);
  if (i < 0 || j < 0) throw new Error(file + ' にブロックが無い');
  return src.slice(i + BEGIN.length, j);
}

let S;
beforeAll(() => {
  const box = vm.createContext({ console, Map, Set, String, Object, Array, Number, Boolean });
  box.globalThis = box;
  vm.runInContext(
    extract('index.html')
    + ';globalThis.__c = { findConflicts, conflictSeverity, countsForOverlap, countsForConflict };',
    box, { filename: 'index.html' });
  S = box.__c;
});

// 予定1行を作る近道
const n = (o) => Object.assign({
  id: 'x', date: '2026-09-10', genba: 'きんでん東', loc: 'A現場', name: '江頭',
  company: 'グローライズ', workType: '現場作業', start: '08:00', end: '17:00',
  yakin: false, yotei: '', yasumi: '', isGhost: false, butai: ''
}, o || {});

const all = (rows) => S.findConflicts(rows, { minSeverity: 'info' });

describe('3段階に分かれる', () => {
  it('★高優先：同じ日に別々の現場作業（物理的に両立しない）', () => {
    const r = all([n({ id: '1' }), n({ id: '2', genba: 'エクシオ', loc: 'B現場' })]);
    expect(r.length).toBe(1);
    expect(r[0].severity).toBe('high');
  });

  it('★要確認：現場＋別業務（判断材料が足りない）', () => {
    const r = all([n({ id: '1' }), n({ id: '2', genba: '自社', loc: '事務所', workType: '現調' })]);
    expect(r.length).toBe(1);
    expect(r[0].severity).toBe('check');
  });

  it('★要確認：時刻が入っていない（両立するか決められない）', () => {
    const r = all([
      n({ id: '1', workType: 'その他', start: '', end: '' }),
      n({ id: '2', genba: '自社', loc: '会議', workType: 'その他', start: '', end: '' })
    ]);
    expect(r.length).toBe(1);
    expect(r[0].severity).toBe('check');
  });

  it('★参考：現場作業を含まない組み合わせ（会議＋社内予定など）', () => {
    const r = all([
      n({ id: '1', workType: 'その他', genba: '自社', loc: '安全会議' }),
      n({ id: '2', workType: 'その他', genba: '自社', loc: '社内打合せ' })
    ]);
    expect(r.length).toBe(1);
    expect(r[0].severity).toBe('info');
  });

  it('同じ現場が2行あるだけなら重複ではない', () => {
    expect(all([n({ id: '1' }), n({ id: '2' })])).toEqual([]);
  });
});

describe('★検知を消していない（社長指示 §6 の肝）', () => {
  const rows = [
    // 高優先になる組
    n({ name: '河原', id: 'a1' }), n({ name: '河原', id: 'a2', genba: 'エクシオ', loc: 'B現場' }),
    // 参考どまりの組
    n({ name: '前﨑', id: 'b1', workType: 'その他', genba: '自社', loc: '安全会議' }),
    n({ name: '前﨑', id: 'b2', workType: 'その他', genba: '自社', loc: '社内打合せ' })
  ];

  it('全部ほしいと言えば、参考どまりの組も返ってくる', () => {
    expect(all(rows).length).toBe(2);
  });

  it('★既定は「高優先」だけ＝既存の画面の出方は変わらない', () => {
    const r = S.findConflicts(rows);
    expect(r.length).toBe(1);
    expect(r[0].severity).toBe('high');
  });

  it('要確認まで下げると、その分だけ増える', () => {
    const r = S.findConflicts(rows, { minSeverity: 'check' });
    expect(r.every((x) => x.severity !== 'info')).toBe(true);
  });
});

describe('拾わないものは今までどおり拾わない', () => {
  it('休み・予定・ゴーストは重複に数えない', () => {
    expect(S.countsForOverlap(n({ yasumi: '休み' }))).toBe(false);
    expect(S.countsForOverlap(n({ yotei: '予定' }))).toBe(false);
    expect(S.countsForOverlap(n({ isGhost: true }))).toBe(false);
  });

  it('★現場系でない予定も「拾いはする」（強さで区別する）', () => {
    // ここが今回の変更点。以前は countsForConflict で捨てていた。
    expect(S.countsForOverlap(n({ workType: '事務所' }))).toBe(true);
    expect(S.countsForConflict(n({ workType: '事務所' }))).toBe(false);
  });

  it('昼と夜勤は別枠のまま（夜勤明けに昼の現場は普通にある）', () => {
    expect(all([n({ id: '1' }), n({ id: '2', genba: 'エクシオ', loc: 'B現場', yakin: true })]))
      .toEqual([]);
  });

  it('会社をまたいで混ぜない（和信カインドの「元」とラーテルの「元」は別人）', () => {
    expect(all([
      n({ id: '1', name: '元', company: '和信カインド' }),
      n({ id: '2', name: '元', company: 'ラーテル', genba: 'エクシオ', loc: 'B現場' })
    ])).toEqual([]);
  });
});

describe('★画面とWorkerで強さの判定が食い違わない', () => {
  const cases = [
    ['別々の現場作業', [n({ id: '1' }), n({ id: '2', genba: 'エクシオ', loc: 'B現場' })]],
    ['現場＋現調', [n({ id: '1' }), n({ id: '2', genba: '自社', loc: '事務所', workType: '現調' })]],
    ['会議どうし', [n({ id: '1', workType: 'その他', loc: '安全会議' }),
                    n({ id: '2', workType: 'その他', loc: '打合せ' })]],
    ['時刻なし', [n({ id: '1', workType: 'その他', loc: 'X', start: '', end: '' }),
                  n({ id: '2', workType: 'その他', loc: 'Y', start: '', end: '' })]]
  ];

  cases.forEach(([label, rows]) => {
    it(label + ' で画面とWorkerの結果が同じ', () => {
      const a = JSON.stringify(S.findConflicts(rows, { minSeverity: 'info' }));
      const b = JSON.stringify(W.findConflicts(rows, { minSeverity: 'info' }));
      expect(b, label + ' で食い違っている').toBe(a);
    });
  });

  it('既定（高優先だけ）でも同じ', () => {
    cases.forEach(([label, rows]) => {
      expect(JSON.stringify(W.findConflicts(rows)), label)
        .toBe(JSON.stringify(S.findConflicts(rows)));
    });
  });
});

describe('★「これが本物」と断定していない', () => {
  it('「これだけが本物」と決めつける書き方をしていない（社長指示 §6）', () => {
    // ★「本物」という語そのものは、社長の言葉の引用と戒めの中に出てくるので
    //   禁止できない。禁止したいのは「これだけが本物」という決めつけの方。
    const src = extract('index.html');
    expect(src).not.toContain('これだけが本物');
    expect(src).toContain('断定しないこと');
    expect(src).toContain('まず見てほしい順');
  });
});

// ================================================================
// ★Codexレビュー#1【P1】（2026-08-31）
//
// 同じ現場の行が複数あるとき、**最初の1行だけ**を見て
// 「現場作業かどうか」を決めていた。だから
//     ① 移動     / 元請A・現場A
//     ② 現場作業 / 元請A・現場A
//     ③ 現場作業 / 元請B・現場B
// の順で並んでいると、現場Aが「現場作業ではない」ままになり、
// 別々の2現場の重なりが高優先から落ちて、画面の警告から消えていた。
// **変更前は出ていた警告**なので、これは後退だった。
// ================================================================
describe('★Codexレビュー#1 同じ現場に別区分の行が先にあっても見落とさない', () => {
  const move = (o) => n(Object.assign({ workType: '移動' }, o || {}));

  it('移動の行が先にあっても、後の現場作業を拾って高優先にする', () => {
    const rows = [
      move({ id: '1', genba: '元請A', loc: '現場A' }),
      n({ id: '2', genba: '元請A', loc: '現場A' }),
      n({ id: '3', genba: '元請B', loc: '現場B' })
    ];
    const r = S.findConflicts(rows);          // 既定＝高優先だけ
    expect(r, '★変更前は出ていた警告が消えた').toHaveLength(1);
    expect(r[0].severity).toBe('high');
  });

  it('Workerでも同じ（画面と食い違わない）', () => {
    const rows = [
      move({ id: '1', genba: '元請A', loc: '現場A' }),
      n({ id: '2', genba: '元請A', loc: '現場A' }),
      n({ id: '3', genba: '元請B', loc: '現場B' })
    ];
    expect(JSON.stringify(W.findConflicts(rows)))
      .toBe(JSON.stringify(S.findConflicts(rows)));
  });

  it('現場作業が先でも結果は同じ（並び順に左右されない）', () => {
    const a = S.findConflicts([
      move({ id: '1', genba: '元請A', loc: '現場A' }),
      n({ id: '2', genba: '元請A', loc: '現場A' }),
      n({ id: '3', genba: '元請B', loc: '現場B' })
    ]);
    const b = S.findConflicts([
      n({ id: '2', genba: '元請A', loc: '現場A' }),
      move({ id: '1', genba: '元請A', loc: '現場A' }),
      n({ id: '3', genba: '元請B', loc: '現場B' })
    ]);
    expect(a.length).toBe(b.length);
    expect(a[0].severity).toBe(b[0].severity);
  });

  it('移動だけの現場は、現場作業に格上げしない', () => {
    const r = S.findConflicts([
      move({ id: '1', genba: '元請A', loc: '現場A' }),
      move({ id: '2', genba: '元請B', loc: '現場B' })
    ], { minSeverity: 'info' });
    expect(r).toHaveLength(1);
    expect(r[0].severity, '移動どうしを高優先にしてはいけない').not.toBe('high');
  });
});
