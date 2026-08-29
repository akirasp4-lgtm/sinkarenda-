// 資格の判定を「実際に動かして」確かめる（フェーズ3の土台・2026-08-28）。
//
// ★なぜvmで動かすか:
//   phase2-conflict.test.js と同じ理由。画面のコードを正規表現で見張るだけだと
//   「書いてあるが動かない」を通してしまう。資格の期限は現場の安全に直結するので、
//   実際に動かして結果を確かめる。
//
// ★Codexレビュー[P1]（2026-08-28）で直した3点をここで押さえている:
//   1. 読めない日付が「期限なし＝一生有効」に化けていた
//   2. 氏名だけの索引で、他社の同姓同名の資格が混ざっていた（奥田さんが実在）
//   3. 選んでいた資格が消えると黙って絞り込みなしに戻っていた（updateQualSelect側）
import { describe, it, expect, beforeAll } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import vm from 'node:vm';

const here = dirname(fileURLToPath(import.meta.url));
const read = (f) => readFileSync(join(here, '..', '..', f), 'utf8');

const BEGIN = '// ===== PHASE3-QUAL-RULE:BEGIN =====';
const END = '// ===== PHASE3-QUAL-RULE:END =====';

function extract(file) {
  const src = read(file);
  const i = src.indexOf(BEGIN), j = src.indexOf(END);
  if (i < 0 || j < 0) throw new Error(file + ' に資格ルールのブロックが無い');
  return src.slice(i + BEGIN.length, j);
}

const EXPORT = `
;globalThis.__p3 = {
  QUAL_SOON_DAYS, QUAL_UNKNOWN, qualCompanyKey, qualKey, qualValidYmd,
  qualIndexBy, qualDaysUntil, qualStatus, qualHeldBy, qualChoices, qualAlerts, qualSafe
};
`;

let Q;
beforeAll(() => {
  const sandbox = vm.createContext({ console });
  sandbox.globalThis = sandbox;
  vm.runInContext(extract('index.html') + EXPORT, sandbox, { filename: 'index.html' });
  Q = sandbox.__p3;
});

const TODAY = '2026-08-28';
const GLO = 'グローライズ';
const q = (name, qual, expires, company) =>
  ({ name, qual, expires: expires === undefined ? '' : expires, kind: '技能講習', company: company || GLO });
const m = (name, company) => ({ name, company: company || GLO });

describe('2つの画面で1文字も違わないこと', () => {
  it('index.html と admin.html の資格ルールが完全に同じ', () => {
    expect(extract('admin.html')).toBe(extract('index.html'));
  });
  it('admin.html 側だけでも実際に動く', () => {
    const sandbox = vm.createContext({ console });
    sandbox.globalThis = sandbox;
    vm.runInContext(extract('admin.html') + EXPORT, sandbox, { filename: 'admin.html' });
    expect(sandbox.__p3.qualStatus('2026-08-27', TODAY)).toBe('expired');
  });
});

describe('実在する日かどうか', () => {
  it('普通の日付は通る', () => {
    expect(Q.qualValidYmd('2026-08-28')).toBe(true);
    expect(Q.qualValidYmd('2028-02-29')).toBe(true);   // うるう年
  });
  it('★存在しない日を弾く（Dateに任せると 2/31 が 3/3 になって通ってしまう）', () => {
    expect(Q.qualValidYmd('2026-02-31')).toBe(false);
    expect(Q.qualValidYmd('2026-13-01')).toBe(false);
    expect(Q.qualValidYmd('2026-00-10')).toBe(false);
    expect(Q.qualValidYmd('2027-02-29')).toBe(false);  // 平年
  });
  it('形が違う物を弾く', () => {
    expect(Q.qualValidYmd('2026/08/28')).toBe(false);
    expect(Q.qualValidYmd('20260828')).toBe(false);
    expect(Q.qualValidYmd('')).toBe(false);
  });
});

describe('期限の判定', () => {
  it('★空欄は「期限なし」＝警告しない（技能講習の多くがこれ）', () => {
    expect(Q.qualStatus('', TODAY)).toBe('none');
    expect(Q.qualStatus(null, TODAY)).toBe('none');
  });
  it('★[P1] 読めない値は「期限なし」ではなく unknown（一生有効な資格に化けさせない）', () => {
    expect(Q.qualStatus('?', TODAY)).toBe('unknown');
    expect(Q.qualStatus('平成31年', TODAY)).toBe('unknown');
    expect(Q.qualStatus('2026-02-31', TODAY)).toBe('unknown');
    expect(Q.qualStatus('20290117', TODAY)).toBe('unknown');
  });
  it('★当日はまだ有効。切れるのは翌日から', () => {
    expect(Q.qualStatus('2026-08-28', TODAY)).toBe('soon');
    expect(Q.qualStatus('2026-08-27', TODAY)).toBe('expired');
  });
  it('60日以内は soon、61日以上先は ok', () => {
    expect(Q.QUAL_SOON_DAYS).toBe(60);
    expect(Q.qualStatus('2026-10-27', TODAY)).toBe('soon');   // 60日後
    expect(Q.qualStatus('2026-10-28', TODAY)).toBe('ok');     // 61日後
  });
  it('残り日数を数えられる', () => {
    expect(Q.qualDaysUntil('2026-08-28', TODAY)).toBe(0);
    expect(Q.qualDaysUntil('2026-08-29', TODAY)).toBe(1);
    expect(Q.qualDaysUntil('2026-08-01', TODAY)).toBe(-27);
    expect(Q.qualDaysUntil('', TODAY)).toBe(null);
    expect(Q.qualDaysUntil('2026-02-31', TODAY)).toBe(null);
  });
});

describe('会社込みで引く（★[P1] 奥田さんはグローライズとGRHDの両方に実在する）', () => {
  it('同じ名字でも会社が違えば別人として扱う', () => {
    const idx = Q.qualIndexBy([q('奥田', '玉掛け', '', GLO)]);
    expect(Q.qualHeldBy(idx, GLO, '奥田', '玉掛け', TODAY)).toBe(true);
    expect(Q.qualHeldBy(idx, 'GRHD', '奥田', '玉掛け', TODAY)).toBe(false);
  });
  it('★川端さんも同じ（グローライズとラーテルの両方にいる）', () => {
    const idx = Q.qualIndexBy([q('川端（達）', '高所作業車', '', GLO)]);
    expect(Q.qualHeldBy(idx, 'ラーテル', '川端（達）', '高所作業車', TODAY)).toBe(false);
  });
  it('★グローライズとGRミツマは1つの名簿として扱う（統合前に取り込んだ26行が消えない）', () => {
    expect(Q.qualCompanyKey('GRミツマ')).toBe(Q.qualCompanyKey('グローライズ'));
    const idx = Q.qualIndexBy([q('江頭', '玉掛け', '', 'GRミツマ')]);
    expect(Q.qualHeldBy(idx, GLO, '江頭', '玉掛け', TODAY)).toBe(true);
  });
  it('他事業（和信カインド）は束ねない', () => {
    expect(Q.qualCompanyKey('和信カインド')).toBe('和信カインド');
    expect(Q.qualCompanyKey('GRHD')).toBe('GRHD');
  });
});

describe('氏名で引けるようにする', () => {
  it('人ごとにまとまる', () => {
    const idx = Q.qualIndexBy([q('真柄', '玉掛け'), q('真柄', '高所作業車'), q('河原', '玉掛け')]);
    expect(idx[Q.qualKey(GLO, '真柄')].map(x => x.qual).sort()).toEqual(['玉掛け', '高所作業車']);
    expect(idx[Q.qualKey(GLO, '河原')]).toHaveLength(1);
  });
  it('★同じ資格が2件あれば期限が先の方を残す（更新して取り直した資格で切れ判定にしない）', () => {
    const idx = Q.qualIndexBy([q('真柄', '玉掛け', '2025-01-01'), q('真柄', '玉掛け', '2030-01-01')]);
    const list = idx[Q.qualKey(GLO, '真柄')];
    expect(list).toHaveLength(1);
    expect(list[0].expires).toBe('2030-01-01');
  });
  it('★期限なしが一番強い（順番が逆でも勝つ）', () => {
    const a = Q.qualIndexBy([q('A', '玉掛け', ''), q('A', '玉掛け', '2025-01-01')]);
    const b = Q.qualIndexBy([q('A', '玉掛け', '2025-01-01'), q('A', '玉掛け', '')]);
    expect(a[Q.qualKey(GLO, 'A')][0].expires).toBe('');
    expect(b[Q.qualKey(GLO, 'A')][0].expires).toBe('');
  });
  it('★読める日付は、読めない物より優先する', () => {
    const a = Q.qualIndexBy([q('A', '玉掛け', '?'), q('A', '玉掛け', '2030-01-01')]);
    const b = Q.qualIndexBy([q('A', '玉掛け', '2030-01-01'), q('A', '玉掛け', '?')]);
    expect(a[Q.qualKey(GLO, 'A')][0].expires).toBe('2030-01-01');
    expect(b[Q.qualKey(GLO, 'A')][0].expires).toBe('2030-01-01');
  });
  it('★索引を作る時点でも、おかしな日付は unknown に倒す', () => {
    const idx = Q.qualIndexBy([q('A', '玉掛け', '2026-02-31')]);
    expect(idx[Q.qualKey(GLO, 'A')][0].expires).toBe('?');
  });
  it('氏名か資格名が空の行は入れない', () => {
    expect(Q.qualIndexBy([q('', '玉掛け'), q('A', ''), {}])).toEqual({});
  });
  it('空・未定義でも落ちない', () => {
    expect(Q.qualIndexBy(null)).toEqual({});
    expect(Q.qualIndexBy([])).toEqual({});
  });
});

describe('資格で人を選ぶ', () => {
  it('持っていれば true', () => {
    const i = Q.qualIndexBy([q('真柄', '玉掛け')]);
    expect(Q.qualHeldBy(i, GLO, '真柄', '玉掛け', TODAY)).toBe(true);
    expect(Q.qualHeldBy(i, GLO, '真柄', '高所作業車', TODAY)).toBe(false);
    expect(Q.qualHeldBy(i, GLO, '知らない人', '玉掛け', TODAY)).toBe(false);
  });
  it('★期限切れの資格では「持っている」と数えない（ここが安全の肝）', () => {
    const i = Q.qualIndexBy([q('真柄', '高所作業車', '2026-08-27')]);
    expect(Q.qualHeldBy(i, GLO, '真柄', '高所作業車', TODAY)).toBe(false);
  });
  it('★[P1] 期限が読めない資格でも「持っている」と数えない（安全側に倒す）', () => {
    const i = Q.qualIndexBy([q('真柄', '高所作業車', 'へんな文字')]);
    expect(Q.qualHeldBy(i, GLO, '真柄', '高所作業車', TODAY)).toBe(false);
  });
  it('期限が近いだけならまだ使える', () => {
    const i = Q.qualIndexBy([q('真柄', '高所作業車', '2026-09-01')]);
    expect(Q.qualHeldBy(i, GLO, '真柄', '高所作業車', TODAY)).toBe(true);
  });
});

describe('プルダウンの中身', () => {
  it('★名簿に載っている人の資格だけ出す（辞めた人の資格を選ばせない）', () => {
    const i = Q.qualIndexBy([q('真柄', '玉掛け'), q('辞めた人', 'フォークリフト')]);
    expect(Q.qualChoices(i, [m('真柄')], TODAY)).toEqual(['玉掛け']);
  });
  it('持っている人が多い順、同数なら文字順', () => {
    const i = Q.qualIndexBy([
      q('A', '玉掛け'), q('B', '玉掛け'), q('C', '玉掛け'),
      q('A', '高所作業車'), q('B', '高所作業車'),
      q('A', 'あ資格'), q('A', 'い資格')
    ]);
    expect(Q.qualChoices(i, [m('A'), m('B'), m('C')], TODAY)).toEqual(['玉掛け', '高所作業車', 'あ資格', 'い資格']);
  });
  it('★[P3] その日に使える人が1人もいない資格は出さない', () => {
    const i = Q.qualIndexBy([q('A', '切れてる', '2026-08-01'), q('A', '生きてる', '')]);
    expect(Q.qualChoices(i, [m('A')], TODAY)).toEqual(['生きてる']);
  });
  it('空でも落ちない', () => {
    expect(Q.qualChoices({}, [], TODAY)).toEqual([]);
    expect(Q.qualChoices(null, null, TODAY)).toEqual([]);
  });
});

describe('期限のお知らせ', () => {
  const src = [
    q('A', '切れてる', '2026-08-01'),
    q('A', 'もうすぐ', '2026-09-10'),
    q('A', '期限なし', ''),
    q('A', 'まだ先', '2030-01-01'),
    q('A', '読めない', 'へんな文字'),
    q('辞めた人', '切れてる', '2026-08-01')
  ];
  const idx = () => Q.qualIndexBy(src);

  it('★切れた物・もうすぐ切れる物・読めない物だけ出す', () => {
    const out = Q.qualAlerts(idx(), [m('A')], TODAY);
    expect(out.map(o => o.qual)).toEqual(['切れてる', 'もうすぐ', '読めない']);
    expect(out[0].status).toBe('expired');
    expect(out[1].status).toBe('soon');
    expect(out[2].status).toBe('unknown');
    expect(out[0].days).toBe(-27);
    expect(out[2].days).toBe(null);
  });
  it('★期限なしと、まだ先の物は出さない（毎日出る警告は誰も読まない）', () => {
    const names = Q.qualAlerts(idx(), [m('A')], TODAY).map(o => o.qual);
    expect(names).not.toContain('期限なし');
    expect(names).not.toContain('まだ先');
  });
  it('★名簿に載っていない人は出さない', () => {
    expect(Q.qualAlerts(idx(), [m('A')], TODAY).map(o => o.name)).not.toContain('辞めた人');
  });
  it('★会社が違えば出さない', () => {
    expect(Q.qualAlerts(idx(), [m('A', 'GRHD')], TODAY)).toEqual([]);
  });
  it('空でも落ちない', () => {
    expect(Q.qualAlerts({}, [], TODAY)).toEqual([]);
    expect(Q.qualAlerts(null, null, TODAY)).toEqual([]);
  });
});

describe('画面側の歯止め（qualSafe）', () => {
  it('★免許番号などが混ざっていても、決めた5項目だけにする', () => {
    const out = Q.qualSafe([{
      name: '河原', company: GLO, qual: '第一種電気工事士', kind: '国家資格', expires: '',
      免許番号: '03569', 正式氏名: '河原　将司', 取得日: '1991-01-24', 出典: 'x.xlsx'
    }]);
    // ★2026-08-29 取得場所(place)を足した。免許番号は引き続き出さない
    expect(Object.keys(out[0]).sort()).toEqual(['company', 'expires', 'kind', 'name', 'place', 'qual']);
    const j = JSON.stringify(out);
    expect(j).not.toContain('03569');
    expect(j).not.toContain('将司');
    expect(j).not.toContain('1991-01-24');
    expect(j).not.toContain('x.xlsx');
  });
  it('空でも落ちない', () => {
    expect(Q.qualSafe(null)).toEqual([]);
    expect(Q.qualSafe([])).toEqual([]);
  });
});
