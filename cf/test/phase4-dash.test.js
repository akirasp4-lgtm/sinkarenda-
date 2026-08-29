// 管理画面（経営者が最初に確認する画面）の数え方を、実際に動かして確かめる。2026-08-29。
//
// ★依頼文（原文）10番:
//   「経営者が最初に確認する画面では、
//     『今日』… 稼働人数／空き人数／現場数／重複警告／未確定案件
//     『今週』… 人員稼働率／空き予定／案件予定 が一目で確認できるようにする。」
//
// ★一番大事な検査は「空き確認の画面と数字が食い違わないこと」。
//   同じ「空き24人」が画面ごとに違うと、誰も数字を信じなくなる。
import { describe, it, expect, beforeAll } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import vm from 'node:vm';

const here = dirname(fileURLToPath(import.meta.url));
const read = (f) => readFileSync(join(here, '..', '..', f), 'utf8');

function block(src, name) {
  const B = '// ===== ' + name + ':BEGIN =====', E = '// ===== ' + name + ':END =====';
  const i = src.indexOf(B), j = src.indexOf(E);
  if (i < 0 || j < 0) throw new Error(name + ' のブロックが無い');
  return src.slice(i + B.length, j);
}
// ★経営の画面は「その日の状態」を空き確認と同じ関数（PHASE2の dayStateByName）で
//   出す。だから2つのブロックをまとめて動かす。
function extract(file) {
  const src = read(file);
  return block(src, 'PHASE2-CONFLICT-RULE') + '\n' + block(src, 'PHASE4-DASH-RULE');
}

const EXPORT = `
;globalThis.__p4 = {
  dashDayState, dashIsSite, dashSiteCount, dashToday, dashWeekStart, dashWeekDays,
  dashWeek, dashUtilization, dashUnconfirmed, dashWeekSites, dayStateByName, DASH_NOT_SITE
};
`;

let D;
beforeAll(() => {
  const sandbox = vm.createContext({ console, Date, String, Object, Math, isNaN });
  sandbox.globalThis = sandbox;
  vm.runInContext(extract('admin.html') + EXPORT, sandbox, { filename: 'admin.html' });
  D = sandbox.__p4;
});

const TODAY = '2026-08-28';   // 金曜
const n = (o) => Object.assign({
  date: TODAY, name: 'A', genba: 'きんでん東', loc: 'A現場', workType: '現場作業',
  yasumi: '', yotei: '', souko: '', isGhost: false
}, o);

describe('その日の人の状態', () => {
  it('休み・倉庫・出勤を分ける', () => {
    const s = D.dashDayState([
      n({ name: 'A' }),
      n({ name: 'B', yasumi: '○' }),
      n({ name: 'C', souko: '○' })
    ], TODAY);
    expect(s).toEqual({ A: 'busy', B: 'yasumi', C: 'souko' });
  });
  it('★同じ日に休みと出勤が両方あるときは「出勤」（空き確認と同じ扱い）', () => {
    expect(D.dashDayState([n({ name: 'A', yasumi: '○' }), n({ name: 'A' })], TODAY).A).toBe('busy');
    expect(D.dashDayState([n({ name: 'A' }), n({ name: 'A', yasumi: '○' })], TODAY).A).toBe('busy');
  });
  it('ゴースト行（前日から）は数えない', () => {
    expect(D.dashDayState([n({ name: 'A', isGhost: true })], TODAY)).toEqual({});
  });
  it('別の日は数えない', () => {
    expect(D.dashDayState([n({ date: '2026-08-27' })], TODAY)).toEqual({});
  });
});

describe('今日の数字', () => {
  const roster = ['A', 'B', 'C', 'D', 'E'];
  it('稼働・空き・休みを足すと名簿の人数になる（数が合う）', () => {
    const t = D.dashToday([
      n({ name: 'A' }), n({ name: 'B' }),
      n({ name: 'C', souko: '○' }),
      n({ name: 'D', yasumi: '○' })
    ], roster, TODAY);
    expect(t.working).toBe(3);   // A,B + 倉庫C
    expect(t.yasumi).toBe(1);
    expect(t.free).toBe(1);      // E
    expect(t.working + t.yasumi + t.free).toBe(t.roster);
  });
  it('★倉庫作業も「稼働」に入れる（仕事をしているため）', () => {
    const t = D.dashToday([n({ name: 'A', souko: '○' })], ['A'], TODAY);
    expect(t.working).toBe(1);
    expect(t.souko).toBe(1);
    expect(t.genba).toBe(0);
    expect(t.free).toBe(0);
  });
  it('★名簿に載っていない人が予定に出ても、人数には数えない', () => {
    const t = D.dashToday([n({ name: '知らない人' })], ['A'], TODAY);
    expect(t.working).toBe(0);
    expect(t.free).toBe(1);
  });
  it('予定が1件も無ければ全員が空き', () => {
    const t = D.dashToday([], roster, TODAY);
    expect(t.free).toBe(5);
    expect(t.working).toBe(0);
  });
});

describe('現場の数', () => {
  it('元請＋現場名の種類を数える（同じ現場に何人いても1）', () => {
    expect(D.dashSiteCount([
      n({ name: 'A', loc: 'A現場' }), n({ name: 'B', loc: 'A現場' }),
      n({ name: 'C', loc: 'B現場' })
    ], TODAY)).toBe(2);
  });
  it('★事務所・倉庫作業・移動は現場として数えない（現場名が入っていても）', () => {
    // Codexが実際に動かして見つけた: workType:'事務所', loc:'本社' が1現場になっていた
    expect(D.dashSiteCount([
      n({ workType: '事務所', loc: '本社' }),
      n({ workType: '倉庫作業', loc: '倉庫' }),
      n({ workType: '移動', loc: '埼玉へ移動' })
    ], TODAY)).toBe(0);
    expect(D.DASH_NOT_SITE).toEqual(['事務所', '倉庫作業', '移動', '休み']);
  });
  it('現調・置局・その他などは現場として数える', () => {
    expect(D.dashSiteCount([
      n({ workType: '現調', loc: 'A' }), n({ workType: '置局', loc: 'B' }),
      n({ workType: 'その他', loc: 'C' })
    ], TODAY)).toBe(3);
  });
  it('★元請が違えば同じ現場名でも別に数える', () => {
    expect(D.dashSiteCount([
      n({ genba: 'きんでん東', loc: '同じ名前' }),
      n({ genba: 'ハイテックス', loc: '同じ名前' })
    ], TODAY)).toBe(2);
  });
  it('休み・📌予定・ゴースト・現場名が空の行は数えない', () => {
    expect(D.dashSiteCount([
      n({ loc: 'A現場', yasumi: '○' }),
      n({ loc: 'B現場', yotei: '○' }),
      n({ loc: 'C現場', isGhost: true }),
      n({ loc: '' })
    ], TODAY)).toBe(0);
  });
});

describe('週の区切り', () => {
  it('月曜始まり', () => {
    expect(D.dashWeekStart('2026-08-28')).toBe('2026-08-24');   // 金 → その週の月曜
    expect(D.dashWeekStart('2026-08-24')).toBe('2026-08-24');   // 月 → そのまま
  });
  it('★日曜は「その週の終わり」として前の月曜に寄せる', () => {
    expect(D.dashWeekStart('2026-08-30')).toBe('2026-08-24');
  });
  it('月をまたいでも正しい', () => {
    expect(D.dashWeekStart('2026-09-01')).toBe('2026-08-31');
  });
  it('7日ぶん出る', () => {
    const d = D.dashWeekDays('2026-08-28');
    expect(d).toHaveLength(7);
    expect(d[0]).toBe('2026-08-24');
    expect(d[6]).toBe('2026-08-30');
  });
  it('おかしな日付でも落ちない', () => {
    expect(D.dashWeekStart('へんな文字')).toBe('');
    expect(D.dashWeekDays('')).toEqual([]);
  });
});

describe('人員稼働率', () => {
  it('稼働 ÷ （名簿×日数）', () => {
    const u = D.dashUtilization([
      { working: 5, yasumi: 0, roster: 10 },
      { working: 5, yasumi: 0, roster: 10 }
    ]);
    expect(u.percent).toBe(50);
    expect(u.days).toBe(2);
  });
  it('★予定が1件も無い日（日曜など）は分母に入れない', () => {
    // これを入れると実態より低い数字が出て、率の意味が無くなる
    const u = D.dashUtilization([
      { working: 10, yasumi: 0, roster: 10 },
      { working: 0, yasumi: 0, roster: 10 }   // まるごと予定なし
    ]);
    expect(u.percent).toBe(100);
    expect(u.days).toBe(1);
  });
  it('★休みの人は分母から外す（有休が多い週ほど低く出るのを防ぐ）', () => {
    // 10人中5人が休み、動ける5人のうち5人が稼働 → 100%
    const u = D.dashUtilization([{ working: 5, yasumi: 5, roster: 10 }]);
    expect(u.percent).toBe(100);
    expect(u.cap).toBe(5);
  });
  it('★全員休みの日は「計算できない」（0%と言い切らない）', () => {
    const u = D.dashUtilization([{ working: 0, yasumi: 10, roster: 10 }]);
    expect(u.percent).toBe(null);
    expect(u.days).toBe(0);
  });
  it('1日も無ければ null（0%と言い切らない）', () => {
    expect(D.dashUtilization([]).percent).toBe(null);
    expect(D.dashUtilization(null).percent).toBe(null);
  });
});

describe('未確定案件', () => {
  it('「見積中」だけ数える', () => {
    expect(D.dashUnconfirmed([
      { status: '見積中' }, { status: '見積中' },
      { status: '受注' }, { status: '施工中' }, { status: '完工' }, { status: '中止' }
    ])).toBe(2);
  });
  it('ステータスが空でも落ちない', () => {
    expect(D.dashUnconfirmed([{}, { status: '' }, null])).toBe(0);
    expect(D.dashUnconfirmed(null)).toBe(0);
  });
});

describe('今週の案件数', () => {
  it('週の中に出てくる現場の種類を数える', () => {
    const rows = [
      n({ date: '2026-08-24', loc: 'A現場' }),
      n({ date: '2026-08-25', loc: 'A現場' }),
      n({ date: '2026-08-26', loc: 'B現場' }),
      n({ date: '2026-08-31', loc: 'C現場' })   // 来週なので入らない
    ];
    expect(D.dashWeekSites(rows, '2026-08-28')).toBe(2);
  });
});

describe('今週の表', () => {
  it('7日分そろい、日付が入る', () => {
    const w = D.dashWeek([n({ name: 'A', date: '2026-08-25' })], ['A', 'B'], '2026-08-28');
    expect(w).toHaveLength(7);
    expect(w[0].date).toBe('2026-08-24');
    expect(w[1].working).toBe(1);
    expect(w[1].free).toBe(1);
    expect(w[0].working).toBe(0);
  });
});
