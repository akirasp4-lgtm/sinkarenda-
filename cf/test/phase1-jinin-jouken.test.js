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
  it('日報データの既存21列は1つも動いていない（末尾に夜勤手当・夜勤請求を追加）', () => {
    const m = CODE.match(/const HEADERS = \[([\s\S]*?)\];/);
    const got = [...m[1].matchAll(/'([^']+)'/g)].map((x) => x[1]);
    // ★2026-09-03 夜勤区分の改修で末尾に2列足した。
    //   このテストの目的は「既存列が動いていないこと」なので、
    //   先頭21列を丸ごと突き合わせる形に強めた。
    expect(got.slice(0, 21)).toEqual([
      '登録日時', '作業日', '元請名', '現場名', '氏名', '役割', '出勤', '退勤',
      '人工', 'メモ', '夜勤', '会社', 'ID', '更新者', '色', '事業部', '工番',
      '作業区分', '車両', '拠点', '部隊'
    ]);
    expect(got.length).toBe(23);
    expect(got[21]).toBe('夜勤手当');
    expect(got[22]).toBe('夜勤請求');
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

// ---------------------------------------------------------------- 入力画面
describe('入力画面（管理画面の現場ごとの詳細）', () => {
  const adm = readFileSync(join(here, '..', '..', 'admin.html'), 'utf8');

  it('★必要資格は打ち込みではなく一覧から選ぶ', () => {
    // 資格名は105種類あり、打ち込むと「高所作業車」と「高所作業認定」のように
    // 割れて永久に一致しない（元請名で起きた表記ゆれと同じ事故）。
    expect(adm).toContain('function qualOptionsHtml()');
    // ★id の存在だけを見ると、<input>（打ち込み欄）に差し替えられても気づけない。
    //   実際に守りを外して試したとき、これで素通りした。必ず <select> であることを見る。
    expect(adm, '資格の欄が打ち込み欄になっている').toContain('<select id="need-qual-pick-');
    // 選択肢は資格マスタの実在する資格名から作る
    const i = adm.indexOf('function qualOptionsHtml()');
    const body = adm.slice(i, adm.indexOf(String.fromCharCode(10) + '}', i));
    expect(body).toContain('allQuals');
    expect(body).toContain('q.qual');
  });

  it('★資格名をHTMLに直接埋めず data-* から読む（引用符事故よけ）', () => {
    expect(adm).toContain('data-quals=');
    expect(adm).toContain('box.dataset.quals');
    // onclick に資格名そのものを埋めていないこと
    expect(adm).not.toMatch(/onclick="removeNeedQual\('/);
  });

  it('★属性へ入れる値は escAttr を通す（esc は引用符を逃がさない）', () => {
    // ★「人員条件」はJS側のコメントにも出てくる。入力欄そのものを起点にする。
    const i = adm.indexOf('id="need-count-');
    expect(i, '入力欄が見つからない').toBeGreaterThan(-1);
    const body = adm.slice(i, i + 3000);
    // ★正規表現にしない。'escAttr(' の丸括弧がグループ開始と解釈されて壊れる。
    //   素直な文字列一致で足りる。
    ['needExp', 'address', 'startAt', 'endAt'].forEach((k) => {
      expect(body, k + ' が escAttr を通っていない').toContain('escAttr(' + k);
    });
  });

  it('6項目すべての入力欄がある', () => {
    ['need-count-', 'need-qual-pick-', 'need-exp-', 'need-addr-', 'need-start-', 'need-end-']
      .forEach((id) => expect(adm, id + ' が無い').toContain('id="' + id));
  });

  it('★未入力は「未登録」と出す（勝手な推測値で埋めない）', () => {
    expect(adm).toContain('placeholder="未登録"');
    expect(adm).toContain("'<span style=\"font-size:11px;color:#999\">未登録</span>'");
  });

  it('★保存は set_site_needs を呼ぶ', () => {
    const i = adm.indexOf('async function saveSiteNeeds');
    const body = adm.slice(i, adm.indexOf(String.fromCharCode(10) + '}', i));
    expect(body).toContain("action:'set_site_needs'");
    expect(body).toContain('updatedBy: getUsername()');
  });

  it('同じ資格を二重に入れない', () => {
    const i = adm.indexOf('function addNeedQual');
    expect(adm.slice(i, adm.indexOf(String.fromCharCode(10) + '}', i))).toContain('indexOf(v)>=0');
  });

  it('★入力を必須にしていない（既存運用を止めない）', () => {
    const i = adm.indexOf('人員条件');
    const body = adm.slice(i, i + 3000);
    expect(body).not.toContain('required');
  });
});
