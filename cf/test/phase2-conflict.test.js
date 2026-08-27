// 人員の重複判定を「実際に動かして」確かめる。
//
// ★なぜvmで動かすか（2026-08-27）:
//   画面のコードを正規表現で見張るだけだと「書いてあるが動かない」を通してしまう。
//   重複の件数は人が毎日見る数字なので、実際に動かして数を確かめる。
import { describe, it, expect, beforeAll } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import vm from 'node:vm';

const here = dirname(fileURLToPath(import.meta.url));
const read = (f) => readFileSync(join(here, '..', '..', f), 'utf8');

const BEGIN = '// ===== PHASE2-CONFLICT-RULE:BEGIN =====';
const END = '// ===== PHASE2-CONFLICT-RULE:END =====';

function extract(file) {
  const src = read(file);
  const i = src.indexOf(BEGIN), j = src.indexOf(END);
  if (i < 0 || j < 0) throw new Error(file + ' に判定ルールのブロックが無い');
  return src.slice(i + BEGIN.length, j);
}

// const は vm のコンテキストのプロパティにならないので、同じ塊の末尾で外へ出す
// （gas-phase1.test.js と同じやり方。2026-08-27 に実測して確かめた性質）
const EXPORT = `
;globalThis.__p2 = {
  GENBA_WORKTYPES, isGenbaWork, conflictBucket, jobKey, countsForConflict,
  findConflicts, conflictsIfAdded
};
`;

let P;
beforeAll(() => {
  const sandbox = vm.createContext({ console });
  sandbox.globalThis = sandbox;
  vm.runInContext(extract('index.html') + EXPORT, sandbox, { filename: 'index.html' });
  P = sandbox.__p2;
});

const row = (o) => Object.assign({
  date: '2026-09-01', name: '中島', company: 'グローライズ',
  genba: 'きんでん西', loc: 'A現場', workType: '現場作業',
  yakin: false, yasumi: false, yotei: false, souko: false, isGhost: false, id: 'x'
}, o);

describe('作業区分の判定', () => {
  it('現場系の4つだけを現場作業とみなす', () => {
    expect(P.GENBA_WORKTYPES).toEqual(['現場作業', '置局', '着打ち', '撤去品返却']);
    ['現場作業', '置局', '着打ち', '撤去品返却'].forEach(w =>
      expect(P.isGenbaWork(row({ workType: w })), w).toBe(true));
    ['現調', '事務所', '移動', 'カギ借用', '材料引取・検品', '倉庫作業', '休み', 'その他', '前乗り', '']
      .forEach(w => expect(P.isGenbaWork(row({ workType: w })), w).toBe(false));
  });

  it('前後の空白を落として判定する', () => {
    expect(P.isGenbaWork(row({ workType: ' 現場作業 ' }))).toBe(true);
  });

  it('workTypeが無い行でも落ちない', () => {
    expect(P.isGenbaWork({})).toBe(false);
    expect(P.isGenbaWork(null)).toBe(false);
  });
});

describe('重複の判定（設計書§1.1のルール）', () => {
  it('★同じ人・同じ日・別の現場（現場系）は重複', () => {
    const c = P.findConflicts([row({ loc: 'A現場' }), row({ loc: 'B現場' })]);
    expect(c.length).toBe(1);
    expect(c[0].name).toBe('中島');
    expect(c[0].date).toBe('2026-09-01');
    expect(c[0].jobs.length).toBe(2);
  });

  it('同じ現場が2行（責任者と班員）は重複ではない', () => {
    expect(P.findConflicts([row({ name: '中島' }), row({ name: '中島' })]).length).toBe(0);
  });

  it('元請だけ違えば別の現場として数える', () => {
    expect(P.findConflicts([
      row({ genba: 'きんでん西', loc: 'A現場' }),
      row({ genba: 'ナンジョウ', loc: 'A現場' })
    ]).length).toBe(1);
  });

  it('★元請と現場名の区切りが無いと別物が同じに見える（境界）', () => {
    // 「きんでん」+「西A現場」と「きんでん西」+「A現場」は別の現場。
    // 区切り無しでつなぐと同じ文字列になり、重複を見逃す
    expect(P.findConflicts([
      row({ genba: 'きんでん', loc: '西A現場' }),
      row({ genba: 'きんでん西', loc: 'A現場' })
    ]).length).toBe(1);
  });

  it('★現場作業＋事務所 は重複ではない（同じ日に両立する）', () => {
    expect(P.findConflicts([
      row({ loc: 'A現場', workType: '現場作業' }),
      row({ loc: '本社', workType: '事務所' })
    ]).length).toBe(0);
  });

  it('★昼と夜勤は別枠。重ならない', () => {
    expect(P.findConflicts([
      row({ loc: 'A現場', yakin: false }),
      row({ loc: 'B現場', yakin: true })
    ]).length).toBe(0);
  });

  it('夜勤どうしが別現場なら重複', () => {
    expect(P.findConflicts([
      row({ loc: 'A現場', yakin: true }),
      row({ loc: 'B現場', yakin: true })
    ]).length).toBe(1);
  });

  it('★「予定」「休み」の行は数えない', () => {
    expect(P.findConflicts([row({ loc: 'A現場' }), row({ loc: 'B現場', yotei: true })]).length).toBe(0);
    expect(P.findConflicts([row({ loc: 'A現場' }), row({ loc: 'B現場', yasumi: true })]).length).toBe(0);
  });

  it('ゴースト行（夜勤の翌日ぶんの影）は数えない', () => {
    expect(P.findConflicts([row({ loc: 'A現場' }), row({ loc: 'B現場', isGhost: true })]).length).toBe(0);
  });

  it('日が違えば重複ではない', () => {
    expect(P.findConflicts([
      row({ date: '2026-09-01', loc: 'A現場' }),
      row({ date: '2026-09-02', loc: 'B現場' })
    ]).length).toBe(0);
  });

  it('人が違えば重複ではない', () => {
    expect(P.findConflicts([row({ name: '中島' }), row({ name: '東', loc: 'B現場' })]).length).toBe(0);
  });

  it('★会社が違う同姓同名は別人として扱う', () => {
    expect(P.findConflicts([
      row({ name: '元', company: '和信カインド', loc: 'A現場' }),
      row({ name: '元', company: 'ラーテル', loc: 'B現場' })
    ]).length).toBe(0);
  });

  it('氏名の前後の空白は無視して同じ人とみなす', () => {
    expect(P.findConflicts([row({ name: '中島', loc: 'A' }), row({ name: ' 中島 ', loc: 'B' })]).length).toBe(1);
  });

  it('3つ重なったら1件にまとめて jobs が3つ', () => {
    const c = P.findConflicts([row({ loc: 'A' }), row({ loc: 'B' }), row({ loc: 'C' })]);
    expect(c.length).toBe(1);
    expect(c[0].jobs.length).toBe(3);
  });

  it('jobs に元請・現場名・そのときのIDが入る', () => {
    const c = P.findConflicts([
      row({ loc: 'A', id: 'i1' }), row({ loc: 'A', id: 'i2' }), row({ loc: 'B', id: 'i3' })
    ]);
    expect(c[0].jobs.length).toBe(2);
    const a = c[0].jobs.find(j => j.loc === 'A');
    expect(a.genba).toBe('きんでん西');
    expect(a.ids.sort()).toEqual(['i1', 'i2']);
  });

  it('opts.from より前の日は返さない', () => {
    const rows = [
      row({ date: '2026-06-29', loc: 'A' }), row({ date: '2026-06-29', loc: 'B' }),
      row({ date: '2026-09-01', loc: 'A' }), row({ date: '2026-09-01', loc: 'B' })
    ];
    expect(P.findConflicts(rows).length).toBe(2);
    expect(P.findConflicts(rows, { from: '2026-08-27' }).length).toBe(1);
    expect(P.findConflicts(rows, { from: '2026-08-27' })[0].date).toBe('2026-09-01');
  });

  it('from と同じ日は含む（境界）', () => {
    const rows = [row({ date: '2026-08-27', loc: 'A' }), row({ date: '2026-08-27', loc: 'B' })];
    expect(P.findConflicts(rows, { from: '2026-08-27' }).length).toBe(1);
  });

  it('日付順→氏名順に並ぶ', () => {
    const c = P.findConflicts([
      row({ date: '2026-09-02', name: '東', loc: 'A' }), row({ date: '2026-09-02', name: '東', loc: 'B' }),
      row({ date: '2026-09-01', name: '中島', loc: 'A' }), row({ date: '2026-09-01', name: '中島', loc: 'B' }),
      row({ date: '2026-09-01', name: '鈴木', loc: 'A' }), row({ date: '2026-09-01', name: '鈴木', loc: 'B' })
    ]);
    expect(c.map(x => x.date + '/' + x.name))
      .toEqual(['2026-09-01/中島', '2026-09-01/鈴木', '2026-09-02/東']);
  });

  it('空でも null でも落ちない', () => {
    expect(P.findConflicts([])).toEqual([]);
    expect(P.findConflicts(null)).toEqual([]);
  });
});

describe('保存しようとしている予定が重複を生むか', () => {
  it('★事務所の予定がある日に現場を入れても警告しない（今までは警告していた）', () => {
    expect(P.conflictsIfAdded(
      [row({ loc: '本社', workType: '事務所' })],
      [row({ loc: 'A現場', id: '' })]
    ).length).toBe(0);
  });

  it('★別の現場が既にある日に現場を入れたら警告する', () => {
    const c = P.conflictsIfAdded([row({ loc: 'A現場' })], [row({ loc: 'B現場', id: '' })]);
    expect(c.length).toBe(1);
    expect(c[0].name).toBe('中島');
    expect(c[0].jobs.map(j => j.loc).sort()).toEqual(['A現場', 'B現場']);
  });

  it('同じ現場に班員として足すだけなら警告しない', () => {
    expect(P.conflictsIfAdded([row({ loc: 'A現場' })], [row({ loc: 'A現場', id: '' })]).length).toBe(0);
  });

  it('★元から重なっていた分は「今回のせい」ではないので出さない', () => {
    expect(P.conflictsIfAdded(
      [row({ loc: 'A現場' }), row({ loc: 'B現場' })],
      [row({ date: '2026-09-05', loc: 'Z現場', id: '' })]
    ).length).toBe(0);
  });

  it('★元から重なっている日に、さらに3つ目を足しても新しい警告にはしない', () => {
    // その日その人は既に警告済み。同じ警告を二度見せない
    expect(P.conflictsIfAdded(
      [row({ loc: 'A現場' }), row({ loc: 'B現場' })],
      [row({ loc: 'C現場', id: '' })]
    ).length).toBe(0);
  });

  it('候補が空・null なら何も出ない', () => {
    expect(P.conflictsIfAdded([row({ loc: 'A現場' })], []).length).toBe(0);
    expect(P.conflictsIfAdded([row({ loc: 'A現場' })], null).length).toBe(0);
  });

  it('既存が空でも、候補どうしが重なれば出す', () => {
    expect(P.conflictsIfAdded([], [
      row({ loc: 'A現場', id: '' }), row({ loc: 'B現場', id: '' })
    ]).length).toBe(1);
  });

  it('候補が「予定」なら警告しない', () => {
    expect(P.conflictsIfAdded(
      [row({ loc: 'A現場' })],
      [row({ loc: 'B現場', yotei: true, id: '' })]
    ).length).toBe(0);
  });
});

describe('画面2つで判定ルールが1文字も違わないこと', () => {
  it('index.html と admin.html のブロックが同一', () => {
    expect(extract('admin.html')).toBe(extract('index.html'));
  });
});
