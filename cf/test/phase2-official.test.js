// Phase 2: 人員不足(§3)・資格不足(§4)の正式判定（2026-08-31）
//
// 社長指示 §3:
//   「第1優先: 案件ごとの必要人数と、実際に配置されている人数を比較する。
//     第2優先: 過去実績との比較。ただし『参考値』であり、断定しない」
// 社長指示 §4:
//   「資格不足 / 期限切れ / 期限間近 / 資格情報が未登録 を別状態として扱う」
// 社長指示 §9:
//   「正式判定・参考判定・判定不可 を明確に分ける」
//
// ★このファイルが守っている一番大事なこと:
//   資格情報をまだ入力していないだけの人を「資格不足」と断定しないこと。
//   資格マスタに1行でも載っているのは62人中22人しかいない（2026-08-31 実測）。
//   ここを間違えると、40人が毎朝「資格不足」と名指しされる。
import { describe, it, expect } from 'vitest';
import {
  buildAlerts, formatAlertsText, siteNeedIndex, siteQualCheck, qualsByPerson, qualPersonKey
} from '../src/alerts.js';

const H = ['ID', '作業日', '元請名', '現場名', '氏名', '役割', '会社', '拠点',
  '作業区分', '出勤', '退勤', '夜勤', '部隊'];

let seq = 0;
const row = (o) => {
  const d = Object.assign({
    ID: 'r' + (++seq), 作業日: '2026-09-10', 元請名: 'きんでん東', 現場名: 'A現場',
    氏名: '江頭', 役割: '', 会社: 'グローライズ', 拠点: '本社',
    作業区分: '現場作業', 出勤: '08:00', 退勤: '17:00', 夜勤: '', 部隊: ''
  }, o || {});
  return H.map((h) => d[h]);
};

const payload = (rows, extra) => Object.assign({
  headers: H, rows,
  members: [
    { name: '江頭', company: 'グローライズ', active: true },
    { name: '河原', company: 'グローライズ', active: true },
    { name: '前﨑', company: 'グローライズ', active: true },
    { name: '真柄', company: 'グローライズ', active: true }
  ],
  genbaMaster: [], jobsites: [], qualifications: []
}, extra || {});

const opt = { date: '2026-09-10', today: '2026-09-09', company: 'グローライズ' };

const site = (o) => Object.assign({
  genba: 'きんでん東', loc: 'A現場', status: '受注',
  needCount: null, needQuals: [], needExp: ''
}, o || {});

// ================================================================ §3 人員不足

describe('§3 人員不足 — 必要人数との比較（正式判定）', () => {
  it('必要4人のところに2人なら「2人不足」と言い切る', () => {
    const a = buildAlerts(payload(
      [row({ 氏名: '江頭' }), row({ 氏名: '河原' })],
      { jobsites: [site({ needCount: 4 })] }
    ), opt);
    expect(a.shortOfficial).toHaveLength(1);
    expect(a.shortOfficial[0]).toMatchObject({ loc: 'A現場', need: 4, count: 2, gap: 2 });
  });

  it('必要人数どおりなら知らせない', () => {
    const a = buildAlerts(payload(
      [row({ 氏名: '江頭' }), row({ 氏名: '河原' })],
      { jobsites: [site({ needCount: 2 })] }
    ), opt);
    expect(a.shortOfficial).toEqual([]);
  });

  it('必要人数より多い日は知らせない', () => {
    const a = buildAlerts(payload(
      [row({ 氏名: '江頭' }), row({ 氏名: '河原' }), row({ 氏名: '前﨑' })],
      { jobsites: [site({ needCount: 2 })] }
    ), opt);
    expect(a.shortOfficial).toEqual([]);
  });

  it('★必要人数が未登録なら正式判定はしない（勝手に推測しない・§0）', () => {
    const a = buildAlerts(payload(
      [row({ 氏名: '江頭' })],
      { jobsites: [site({ needCount: null })] }
    ), opt);
    expect(a.shortOfficial).toEqual([]);
  });

  it('★同じ人が2行入っていても1人と数える', () => {
    const a = buildAlerts(payload(
      [row({ 氏名: '江頭' }), row({ 氏名: '江頭', 作業区分: '現場作業' })],
      { jobsites: [site({ needCount: 2 })] }
    ), opt);
    expect(a.shortOfficial[0]).toMatchObject({ need: 2, count: 1, gap: 1 });
  });

  it('★昼と夜を分けない（必要人数の欄は現場に1つしかない）', () => {
    // 分けると「昼1＋夜1で必要2」の現場が両方とも不足と誤報する
    const a = buildAlerts(payload(
      [row({ 氏名: '江頭' }), row({ 氏名: '河原', 夜勤: '夜勤' })],
      { jobsites: [site({ needCount: 2 })] }
    ), opt);
    expect(a.shortOfficial).toEqual([]);
  });

  it('★1人も入っていない現場は判定しない（動いていない日と区別できない）', () => {
    const a = buildAlerts(payload(
      [row({ 現場名: '別現場' })],
      { jobsites: [site({ needCount: 5 }), site({ loc: '別現場' })] }
    ), opt);
    expect(a.shortOfficial).toEqual([]);
  });

  it('不足が大きい順に並ぶ', () => {
    const a = buildAlerts(payload(
      [row({ 氏名: '江頭' }), row({ 氏名: '河原', 現場名: 'B現場' })],
      { jobsites: [site({ needCount: 3 }), site({ loc: 'B現場', needCount: 9 })] }
    ), opt);
    expect(a.shortOfficial.map((s) => s.loc)).toEqual(['B現場', 'A現場']);
  });

  it('★事務所・倉庫は現場として数えない', () => {
    const a = buildAlerts(payload(
      [row({ 氏名: '江頭', 現場名: '事務所', 作業区分: 'その他' })],
      { jobsites: [site({ loc: '事務所', needCount: 5 })] }
    ), opt);
    expect(a.shortOfficial).toEqual([]);
  });

  it('★他社の人を頭数に入れない（§0 会社境界）', () => {
    const a = buildAlerts(payload(
      [row({ 氏名: '江頭' }), row({ 氏名: '別人', 会社: '和信カインド' })],
      { jobsites: [site({ needCount: 2 })] }
    ), opt);
    expect(a.shortOfficial[0]).toMatchObject({ count: 1, gap: 1 });
  });
});

describe('§3 正式判定と参考判定を二重に出さない', () => {
  // いつも4人の実績を作り、その日2人にする
  const past = [];
  ['2026-09-01', '2026-09-02', '2026-09-03', '2026-09-04', '2026-09-05'].forEach((d) => {
    ['江頭', '河原', '前﨑', '真柄'].forEach((n) => past.push(row({ 作業日: d, 氏名: n })));
  });
  const dayRows = [row({ 氏名: '江頭' }), row({ 氏名: '河原' })];

  it('必要人数が未登録なら、今までどおり参考判定（実績ベース）で出る', () => {
    const a = buildAlerts(payload(past.concat(dayRows)), opt);
    expect(a.shortStaff).toHaveLength(1);
    expect(a.shortOfficial).toEqual([]);
  });

  it('★必要人数を登録したら、参考判定からは消えて正式判定だけになる', () => {
    const a = buildAlerts(payload(past.concat(dayRows),
      { jobsites: [site({ needCount: 4 })] }), opt);
    expect(a.shortStaff).toEqual([]);
    expect(a.shortOfficial).toHaveLength(1);
  });

  it('★文面でも正式と参考が別の見出しになっている（§9）', () => {
    const t = formatAlertsText(buildAlerts(payload(past.concat(dayRows),
      { jobsites: [site({ needCount: 4 })] }), opt));
    expect(t).toContain('必要人数に足りていません');
    expect(t).toContain('必要4人 → 2人（2人不足）');
    expect(t).not.toContain('いつもより人が少ない可能性');
  });

  it('★参考判定の文面は断定していない（§3）', () => {
    const t = formatAlertsText(buildAlerts(payload(past.concat(dayRows)), opt));
    expect(t).toContain('可能性があります');
    expect(t).toContain('（参考）');
  });
});

// ================================================================ §4 資格不足

const q = (name, qual, expires, company) => ({
  name, qual, expires, company: company || 'グローライズ'
});

describe('§4 資格不足 — 4つの状態を混ぜない', () => {
  const jobs = { jobsites: [site({ needCount: 1, needQuals: ['高所作業車'] })] };
  const one = [row({ 氏名: '江頭' })];

  it('有効な資格を持っていれば何も出さない', () => {
    const a = buildAlerts(payload(one, Object.assign({
      qualifications: [q('江頭', '高所作業車', '2030-01-01')]
    }, jobs)), opt);
    expect(a.qualShort).toEqual([]);
  });

  it('期限が切れていれば「期限切れ」', () => {
    const a = buildAlerts(payload(one, Object.assign({
      qualifications: [q('江頭', '高所作業車', '2024-05-31')]
    }, jobs)), opt);
    expect(a.qualShort).toHaveLength(1);
    expect(a.qualShort[0]).toMatchObject({ qual: '高所作業車', status: 'expired' });
    expect(a.qualShort[0].expired).toEqual(['江頭']);
  });

  it('期限が近ければ「期限間近」', () => {
    const a = buildAlerts(payload(one, Object.assign({
      qualifications: [q('江頭', '高所作業車', '2026-10-01')]
    }, jobs)), opt);
    expect(a.qualShort[0]).toMatchObject({ status: 'soon' });
  });

  it('★資格が登録されている人が誰も持っていなければ「誰も持っていない」', () => {
    const a = buildAlerts(payload(one, Object.assign({
      qualifications: [q('江頭', '玉掛け', '2030-01-01')]   // 別の資格は持っている
    }, jobs)), opt);
    expect(a.qualShort[0]).toMatchObject({ status: 'missing' });
  });

  it('★★資格情報が1行も無い人は「資格不足」と断定しない（§0・一番大事）', () => {
    // 資格マスタに載っているのは62人中22人。ここを間違えると40人が毎朝名指しされる。
    const a = buildAlerts(payload(one, Object.assign({
      qualifications: []
    }, jobs)), opt);
    expect(a.qualShort[0].status).toBe('unknown');
    expect(a.qualShort[0].status).not.toBe('missing');
    expect(a.qualShort[0].noRecord).toEqual(['江頭']);
  });

  it('★1人でも資格情報が無ければ、他の人が持っていなくても断定しない', () => {
    const a = buildAlerts(payload(
      [row({ 氏名: '江頭' }), row({ 氏名: '河原' })],
      Object.assign({
        qualifications: [q('江頭', '玉掛け', '2030-01-01')]  // 河原は1行も無い
      }, { jobsites: [site({ needCount: 2, needQuals: ['高所作業車'] })] })
    ), opt);
    expect(a.qualShort[0].status).toBe('unknown');
    expect(a.qualShort[0].noRecord).toEqual(['河原']);
  });

  it('有効期限が読めない資格は「判定できない」（勝手に切れている扱いにしない）', () => {
    const a = buildAlerts(payload(one, Object.assign({
      qualifications: [q('江頭', '高所作業車', '未定')]
    }, jobs)), opt);
    expect(a.qualShort[0].status).toBe('unknown');
    expect(a.qualShort[0].why).toContain('有効期限');
  });

  it('★1人でも有効なら足りている（他の人が切れていても現場は回る）', () => {
    const a = buildAlerts(payload(
      [row({ 氏名: '江頭' }), row({ 氏名: '河原' })],
      Object.assign({
        qualifications: [q('江頭', '高所作業車', '2024-01-01'),
          q('河原', '高所作業車', '2030-01-01')]
      }, { jobsites: [site({ needCount: 2, needQuals: ['高所作業車'] })] })
    ), opt);
    expect(a.qualShort).toEqual([]);
  });

  it('★必要資格が未登録の現場は何も判定しない（§0）', () => {
    const a = buildAlerts(payload(one, {
      jobsites: [site({ needCount: 1, needQuals: [] })],
      qualifications: []
    }), opt);
    expect(a.qualShort).toEqual([]);
  });

  it('★他社の人の資格を混ぜない（§0 会社境界）', () => {
    // 和信カインドの同姓「江頭」が持っていても、グローライズの通知では足りていない
    const a = buildAlerts(payload(one, Object.assign({
      qualifications: [q('江頭', '高所作業車', '2030-01-01', '和信カインド')]
    }, jobs)), opt);
    expect(a.qualShort).toHaveLength(1);
    expect(a.qualShort[0].status).toBe('unknown');   // 資格情報が無い扱い＝断定しない
  });

  it('困る順に並ぶ（誰も持っていない → 切れている → 期限間近 → 判定不可）', () => {
    const a = buildAlerts(payload(
      [row({ 氏名: '江頭' }), row({ 氏名: '河原', 現場名: 'B現場' }),
        row({ 氏名: '前﨑', 現場名: 'C現場' })],
      {
        jobsites: [
          site({ needCount: 1, needQuals: ['資格X'] }),                      // 情報無し→unknown
          site({ loc: 'B現場', needCount: 1, needQuals: ['資格Y'] }),         // 切れている
          site({ loc: 'C現場', needCount: 1, needQuals: ['資格Z'] })          // 誰も持っていない
        ],
        qualifications: [q('河原', '資格Y', '2024-01-01'), q('前﨑', '別資格', '2030-01-01')]
      }
    ), opt);
    expect(a.qualShort.map((x) => x.status)).toEqual(['missing', 'expired', 'unknown']);
  });

  it('文面に状態が日本語で出る', () => {
    const t = formatAlertsText(buildAlerts(payload(one, Object.assign({
      qualifications: [q('江頭', '高所作業車', '2024-05-31')]
    }, jobs)), opt));
    expect(t).toContain('現場に必要な資格が足りていません');
    expect(t).toContain('高所作業車 … 期限が切れている（江頭さん）');
  });

  it('★★正式判定と「判定できない」を同じ見出しに混ぜない（§9）', () => {
    // 江頭は資格が登録されていない → 判定できない
    // 河原は資格が登録されていて、必要な資格を持っていない → 正式に不足
    const a = buildAlerts(payload(
      [row({ 氏名: '江頭' }), row({ 氏名: '河原', 現場名: 'B現場' })],
      {
        jobsites: [site({ needCount: 1, needQuals: ['高所作業車'] }),
          site({ loc: 'B現場', needCount: 1, needQuals: ['高所作業車'] })],
        qualifications: [q('河原', '玉掛け', '2030-01-01')]
      }
    ), opt);
    const t = formatAlertsText(a);
    expect(t).toContain('現場に必要な資格が足りていません 1件');
    expect(t).toContain('資格を確かめられませんでした 1件（判定できません）');
    // 判定できない方が、足りていない方の件数に混ざっていないこと
    expect(t).not.toContain('資格が足りていません 2件');
  });

  it('★「判定できない」側は不足と書かない（断定しない）', () => {
    const t = formatAlertsText(buildAlerts(payload(one, Object.assign({
      qualifications: []
    }, jobs)), opt));
    expect(t).toContain('判定できません');
    expect(t).toContain('資格情報が未登録の人がいます');
    expect(t).not.toContain('資格が足りていません');
  });
});

// ================================================================ 部品ごと

describe('部品：siteNeedIndex', () => {
  it('0人・マイナス・文字は「未登録」にする（0人必要とは別物）', () => {
    const ix = siteNeedIndex([
      site({ loc: 'A', needCount: 0 }), site({ loc: 'B', needCount: -3 }),
      site({ loc: 'C', needCount: 'あ' }), site({ loc: 'D', needCount: 4 })
    ]);
    expect(ix['きんでん東A'].needCount).toBe(null);
    expect(ix['きんでん東B'].needCount).toBe(null);
    expect(ix['きんでん東C'].needCount).toBe(null);
    expect(ix['きんでん東D'].needCount).toBe(4);
  });

  it('必要資格の空欄は落とす', () => {
    const ix = siteNeedIndex([site({ needQuals: ['玉掛け', '', '  '] })]);
    expect(ix['きんでん東A現場'].needQuals).toEqual(['玉掛け']);
  });
});

describe('部品：qualsByPerson', () => {
  const key = (name, company) => qualPersonKey({ name, company });

  it('会社で絞る', () => {
    const ix = qualsByPerson(
      [q('江頭', '玉掛け', '2030-01-01'), q('江頭', '高所', '2030-01-01', '和信カインド')],
      'グローライズ');
    expect(ix[key('江頭', 'グローライズ')]).toHaveLength(1);
    expect(ix[key('江頭', 'グローライズ')][0].qual).toBe('玉掛け');
  });

  it('★全社でも、鍵が会社込みなので他社と混ざらない（Codexレビュー#8）', () => {
    // /api/alerts は company を省くと「全社」で動く。
    // 昔は氏名だけの鍵だったので、和信カインドの江頭さんの資格が
    // グローライズの江頭さんの物として使えていた。
    const ix = qualsByPerson(
      [q('江頭', '玉掛け', '2030-01-01'), q('江頭', '高所', '2030-01-01', '和信カインド')],
      '全社');
    expect(ix[key('江頭', 'グローライズ')]).toHaveLength(1);
    expect(ix[key('江頭', '和信カインド')]).toHaveLength(1);
    expect(ix[key('江頭', 'グローライズ')][0].qual).toBe('玉掛け');
    expect(ix[key('江頭', '和信カインド')][0].qual).toBe('高所');
  });

  it('★会社が空欄の行を落とさない（Codexレビュー#2）', () => {
    // 会社の列が無い／セルが空の行は「どの会社か分からない」であって
    // 「他社の行」ではない。落とすと期限のお知らせが黙って消える。
    const ix = qualsByPerson([{ name: '江頭', qual: '玉掛け', expires: '', company: '' }],
      'グローライズ');
    expect(Object.keys(ix)).toHaveLength(1);
    expect(ix[key('江頭', '')]).toHaveLength(1);
  });
});

describe('部品：siteQualCheck', () => {
  const byP = { 江頭: [q('江頭', '玉掛け', '2030-01-01')] };

  it('必要資格が空なら何も返さない', () => {
    expect(siteQualCheck([], ['江頭'], byP, '2026-09-09')).toEqual([]);
  });

  it('空文字の必要資格は無視する', () => {
    expect(siteQualCheck(['', '  '], ['江頭'], byP, '2026-09-09')).toEqual([]);
  });

  it('必要資格ごとに1件ずつ返す', () => {
    const r = siteQualCheck(['玉掛け', '高所作業車'], ['江頭'], byP, '2026-09-09');
    expect(r).toHaveLength(2);
    expect(r[0].status).toBe('ok');
    expect(r[1].status).toBe('missing');
  });
});

// ================================================================
// Codexレビューで見つかった穴（2026-08-31）。同じ穴を二度と空けない。
// ================================================================

describe('★Codexレビュー#7【重大】期限のない資格を「読めない」にしない', () => {
  // 資格284行のうち245行（86%）が期限なし＝ゴンドラ特別教育・足場の組立て等・
  // 職長など、そもそも切れない資格。これを全部「判定できない」にしていた。
  const jobs = { jobsites: [site({ needCount: 1, needQuals: ['職長・安全衛生責任者'] })] };
  const one = [row({ 氏名: '江頭' })];

  it('有効期限が空欄なら「持っている」（切れない資格）', () => {
    const a = buildAlerts(payload(one, Object.assign({
      qualifications: [q('江頭', '職長・安全衛生責任者', '')]
    }, jobs)), opt);
    expect(a.qualShort, '期限なしの資格を持っているのに何か言われた').toEqual([]);
  });

  it('期限の欄が無い（undefined）でも「持っている」', () => {
    const a = buildAlerts(payload(one, Object.assign({
      qualifications: [{ name: '江頭', qual: '職長・安全衛生責任者', company: 'グローライズ' }]
    }, jobs)), opt);
    expect(a.qualShort).toEqual([]);
  });

  it('部品でも同じ（siteQualCheck）', () => {
    const ix = { '江頭': [{ name: '江頭', qual: '玉掛け', expires: '' }] };
    expect(siteQualCheck(['玉掛け'], ['江頭'], ix, '2026-09-09')[0].status).toBe('ok');
  });

  it('★読めない値は今までどおり「持っている」に数えない（安全側）', () => {
    const ix = { '江頭': [{ name: '江頭', qual: '玉掛け', expires: '未定' }] };
    expect(siteQualCheck(['玉掛け'], ['江頭'], ix, '2026-09-09')[0].status).toBe('unknown');
  });
});

describe('★Codexレビュー#4 必要資格だけでも資格の判定が動く', () => {
  it('必要人数が未登録でも、必要資格が入っていれば照らす', () => {
    const a = buildAlerts(payload(
      [row({ 氏名: '江頭' })],
      {
        jobsites: [site({ needCount: null, needQuals: ['高所作業車'] })],
        qualifications: [q('江頭', '玉掛け', '')]
      }
    ), opt);
    expect(a.qualShort, '必要資格だけの現場が素通りした').toHaveLength(1);
    expect(a.qualShort[0].status).toBe('missing');
  });

  it('その場合でも人数の判定はしない（人数は未登録なのだから）', () => {
    const a = buildAlerts(payload(
      [row({ 氏名: '江頭' })],
      {
        jobsites: [site({ needCount: null, needQuals: ['高所作業車'] })],
        qualifications: [q('江頭', '玉掛け', '')]
      }
    ), opt);
    expect(a.shortOfficial).toEqual([]);
  });
});

describe('★Codexレビュー#5 資格情報が無い人がいたら、期限切れでも断定しない', () => {
  it('Aが期限切れ・Bが資格情報なし → 「判定できない」', () => {
    const a = buildAlerts(payload(
      [row({ 氏名: '江頭' }), row({ 氏名: '河原' })],
      {
        jobsites: [site({ needCount: 2, needQuals: ['高所作業車'] })],
        qualifications: [q('江頭', '高所作業車', '2024-01-01')]   // 河原は1行も無い
      }
    ), opt);
    expect(a.qualShort[0].status, '期限切れで断定してしまった').toBe('unknown');
    expect(a.qualShort[0].why).toContain('期限切れ');
    expect(a.qualShort[0].noRecord).toEqual(['河原']);
  });

  it('全員ぶん登録があって誰も持っていなければ、今までどおり断定する', () => {
    const a = buildAlerts(payload(
      [row({ 氏名: '江頭' }), row({ 氏名: '河原' })],
      {
        jobsites: [site({ needCount: 2, needQuals: ['高所作業車'] })],
        qualifications: [q('江頭', '高所作業車', '2024-01-01'), q('河原', '玉掛け', '')]
      }
    ), opt);
    expect(a.qualShort[0].status).toBe('expired');
  });
});

describe('★Codexレビュー#8 全社で見ても他社の同姓の資格を使わない', () => {
  it('全社（company未指定）でも混ざらない', () => {
    // /api/alerts は company を省くと「全社」で動く＝既定でこの穴が開いていた
    const a = buildAlerts({
      headers: H,
      rows: [row({ 氏名: '江頭', 会社: 'グローライズ' })],
      members: [{ name: '江頭', company: 'グローライズ', active: true },
        { name: '江頭', company: '和信カインド', active: true }],
      genbaMaster: [],
      jobsites: [site({ needCount: 1, needQuals: ['高所作業車'] })],
      qualifications: [q('江頭', '高所作業車', '', '和信カインド')]
    }, { date: '2026-09-10', today: '2026-09-09', company: '全社' });
    expect(a.qualShort, '他社の資格で足りている事にされた').toHaveLength(1);
    expect(a.qualShort[0].status).toBe('unknown');
  });
});
