// Phase 1: 案件ごとの「必要人員条件」を持てるようにする（2026-08-31）
//
// 社長指示 §2:
//   「案件・現場ごとに 必要人数 / 必要資格 / 必要経験 / 現場住所 /
//     開始時間 / 終了時間 を持てるようにする」
//   「既存223現場を一括で無理に入力する必要はない。空欄を許容し、
//     空欄の場合は『条件未登録』として扱い、勝手な推測値で正式判定しないこと」
//
// ★この段階は「入れ物」だけ。判定（人員不足・資格不足・移動時間）は次の段階。
//
// ★ここで一番守りたいのは2つ:
//   ① 既存12列が1ミリも動いていないこと（動かすと全データの意味がずれる）
//   ② 窓口の出力に6項目が書かれていること
//      （書かないと列を作っても画面に永久に届かない。過去3回ハマっている）
import { describe, it, expect, beforeAll } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import vm from 'node:vm';

const here = dirname(fileURLToPath(import.meta.url));
const CODE = readFileSync(join(here, '..', '..', 'gas.js'), 'utf8');

let G;
beforeAll(() => {
  const box = vm.createContext({ console, String, Number, Object, Array, Math, isFinite, JSON, Date, RegExp });
  box.globalThis = box;
  vm.runInContext(
    CODE + ';globalThis.__p1 = { normalizeNeedCount_, normalizeNeedQuals_, normalizeNeedTime_, SITE_NEED_SEP };',
    box, { filename: 'gas.js' });
  G = box.__p1;
});

// ---------------------------------------------------------------- 台帳の形
describe('現場マスタの列', () => {
  // 見出しを定義している所（新規作成時）を切り出す
  const headerLine = () => {
    const i = CODE.indexOf("function getOrCreateJobSiteSheet_");
    const j = CODE.indexOf('return sheet;', i);
    return CODE.slice(i, j);
  };

  it('★既存12列の名前と順番が1つも変わっていない', () => {
    const src = headerLine();
    const want = ['元請名', '現場名', '工番', '事業部', '年度', '連番',
      '売上', '読み', '完了', '請求方式', '拠点', 'ステータス'];
    const m = src.match(/appendRow\(\[([\s\S]*?)\]\)/);
    expect(m, '見出しの定義が見つからない').toBeTruthy();
    const got = [...m[1].matchAll(/'([^']+)'/g)].map((x) => x[1]);
    expect(got.slice(0, 12), '既存12列が動いている').toEqual(want);
  });

  it('★末尾に6列（必要人数〜終了時間）が足されている', () => {
    const src = headerLine();
    const m = src.match(/appendRow\(\[([\s\S]*?)\]\)/);
    const got = [...m[1].matchAll(/'([^']+)'/g)].map((x) => x[1]);
    expect(got.length, '18列になっていない').toBe(18);
    expect(got.slice(12)).toEqual(
      ['必要人数', '必要資格', '必要経験', '現場住所', '開始時間', '終了時間']);
  });

  it('既にあるシートも18列へ広げる', () => {
    expect(headerLine()).toContain('ensureColumns_(sheet, 18)');
  });
});

// ---------------------------------------------------------------- 窓口の出力
describe('窓口（doGet）が6項目を返す', () => {
  // ★同じ文字列が別の関数にもあるので、必ず doGet の中から探す
  const jobsitesOut = () => {
    const g = CODE.indexOf('function doGet(e)');
    const i = CODE.indexOf('const jobSiteSheet = getOrCreateJobSiteSheet_(ss);', g);
    const j = CODE.indexOf('.filter(', i);
    return CODE.slice(i, j);
  };

  it('★6項目すべてが書かれている（書かないと画面に永久に届かない）', () => {
    const src = jobsitesOut();
    ['needCount', 'needQuals', 'needExp', 'address', 'startAt', 'endAt'].forEach((k) => {
      expect(src, k + ' が窓口の出力に無い').toMatch(new RegExp('\\b' + k + '\\s*:'));
    });
  });

  it('既存の項目も残っている（消しすぎていない）', () => {
    const src = jobsitesOut();
    ['genba', 'loc', 'jobNo', 'completed', 'billingMethod', 'kyoten'].forEach((k) => {
      expect(src, k + ' が消えている').toMatch(new RegExp('\\b' + k + '\\s*:'));
    });
  });

  it('★必要人数は素の値ではなく整える関数を通す（空欄を0にしないため）', () => {
    expect(jobsitesOut()).toContain('normalizeNeedCount_(r[12])');
  });
});

// ---------------------------------------------------------------- 空欄＝条件未登録
describe('必要人数：空欄は「条件未登録」であって「0人」ではない', () => {
  it('★空欄は null（0ではない）', () => {
    expect(G.normalizeNeedCount_('')).toBeNull();
    expect(G.normalizeNeedCount_(null)).toBeNull();
    expect(G.normalizeNeedCount_(undefined)).toBeNull();
    expect(G.normalizeNeedCount_('   ')).toBeNull();
  });

  it('★0 や マイナス も「条件未登録」として扱う', () => {
    // 0を通すと、次の段階の判定が「常に足りている」と誤って判断する
    expect(G.normalizeNeedCount_(0)).toBeNull();
    expect(G.normalizeNeedCount_('0')).toBeNull();
    expect(G.normalizeNeedCount_(-3)).toBeNull();
  });

  it('数字でないものも「条件未登録」（勝手に推測しない）', () => {
    expect(G.normalizeNeedCount_('たぶん6人')).toBeNull();
    expect(G.normalizeNeedCount_('？')).toBeNull();
  });

  it('ふつうの数字はそのまま通る', () => {
    expect(G.normalizeNeedCount_(6)).toBe(6);
    expect(G.normalizeNeedCount_('6')).toBe(6);
    expect(G.normalizeNeedCount_(' 12 ')).toBe(12);
  });

  it('小数は切り捨てる（人は割れない）', () => {
    expect(G.normalizeNeedCount_('6.7')).toBe(6);
  });
});

// ---------------------------------------------------------------- 必要資格
describe('必要資格：読点区切りの文字列を配列にする', () => {
  it('★空欄は空配列（「資格が要らない」ではなく「未登録」として次の段階で扱う）', () => {
    expect(G.normalizeNeedQuals_('')).toEqual([]);
    expect(G.normalizeNeedQuals_(null)).toEqual([]);
  });

  it('全角の読点で区切れる', () => {
    expect(G.normalizeNeedQuals_('高所作業認定、玉掛け'))
      .toEqual(['高所作業認定', '玉掛け']);
  });

  it('半角カンマ・全角カンマでも拾う（人が打つ欄なので）', () => {
    expect(G.normalizeNeedQuals_('高所作業認定,玉掛け')).toEqual(['高所作業認定', '玉掛け']);
    expect(G.normalizeNeedQuals_('高所作業認定，玉掛け')).toEqual(['高所作業認定', '玉掛け']);
  });

  it('★前後の空白を落とす（資格マスタと1文字も違ってはいけないため）', () => {
    expect(G.normalizeNeedQuals_(' 高所作業認定 、 玉掛け '))
      .toEqual(['高所作業認定', '玉掛け']);
  });

  it('区切りだけ・空の要素は捨てる', () => {
    expect(G.normalizeNeedQuals_('、、')).toEqual([]);
    expect(G.normalizeNeedQuals_('玉掛け、、')).toEqual(['玉掛け']);
  });

  it('1つだけでも配列で返す', () => {
    expect(G.normalizeNeedQuals_('フルハーネス型墜落制止用器具'))
      .toEqual(['フルハーネス型墜落制止用器具']);
  });
});

// ---------------------------------------------------------------- 時刻
describe('開始・終了時間：おかしい値は空欄にする（勝手に補わない）', () => {
  it('HH:MM だけ通す', () => {
    expect(G.normalizeNeedTime_('08:00')).toBe('08:00');
    expect(G.normalizeNeedTime_('8:00')).toBe('8:00');
    expect(G.normalizeNeedTime_('23:59')).toBe('23:59');
  });

  it('★おかしい時刻は空欄（推測で直さない）', () => {
    expect(G.normalizeNeedTime_('24:00')).toBe('');
    expect(G.normalizeNeedTime_('08:60')).toBe('');
    expect(G.normalizeNeedTime_('8時')).toBe('');
    expect(G.normalizeNeedTime_('')).toBe('');
    expect(G.normalizeNeedTime_(null)).toBe('');
  });
});

// ---------------------------------------------------------------- 既存を壊していない
describe('既存の仕組みを壊していない', () => {
  it('日報データの21列は1つも変わっていない', () => {
    const m = CODE.match(/const HEADERS = \[([\s\S]*?)\];/);
    const got = [...m[1].matchAll(/'([^']+)'/g)].map((x) => x[1]);
    expect(got.length).toBe(21);
    expect(got[19]).toBe('拠点');
    expect(got[20]).toBe('部隊');
  });

  it('★案件ステータスの8段階は変わっていない', () => {
    const m = CODE.match(/const SITE_STATUSES = \[([\s\S]*?)\];/);
    const got = [...m[1].matchAll(/'([^']+)'/g)].map((x) => x[1]);
    expect(got).toEqual(['見積中', '受注', '準備中', '施工中', '残工事', '完工', '延期', '中止']);
  });

  it('窓口は日当を出さないまま（前回の対策が生きている）', () => {
    const g = CODE.indexOf('function doGet(e)');
    const i = CODE.indexOf('const memberSheet = getOrCreateMemberSheet_(ss);', g);
    const j = CODE.indexOf('const genbaSheet = getOrCreateGenbaSheet_(ss);', i);
    expect(CODE.slice(i, j)).not.toMatch(/\brate\s*:/);
  });
});
