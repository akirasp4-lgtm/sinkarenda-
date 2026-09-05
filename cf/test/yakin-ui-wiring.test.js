// 夜勤手当・夜勤請求が「画面を一往復して消えない」ことを守るテスト（2026-09-03）
//
// ★なぜこれが要るか（過去3回同じ壊れ方をしている）:
//   日報データに列を足しても、画面の parseRows に書かなければ画面のデータから消える。
//   すると夜勤と無関係な編集（時刻を1つ直しただけ等）で保存し直したときに、
//   事務が入れた「対象外」が黙って消える。拠点（2026-08-26）と部隊（2026-08-27）で
//   実際に起きた壊れ方で、Codexレビュー[P1]#1 でも指摘されている。
//
//   現場画面(index.html)にはUIを付けないが、**読み書きの往復だけは必ず要る**。
//   現場の職人が予定を1つ直しただけで事務の設定が飛ぶのを防ぐため。
import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const here = dirname(fileURLToPath(import.meta.url));
const read = (f) => readFileSync(join(here, '..', '..', f), 'utf8');
const FILES = ['admin.html', 'index.html'];

FILES.forEach((f) => {
  const src = read(f);

  describe(f + ': 夜勤手当・夜勤請求が画面を一往復する', () => {
    it('★parseRows でシートの列を読んでいる（読まないと画面から消える）', () => {
      expect(src).toContain("yakinTeate:String(r['夜勤手当']||'')");
      expect(src).toContain("yakinSeikyu:String(r['夜勤請求']||'')");
    });

    it('★編集の保存で値を送っている（送らないと保存で消える）', () => {
      // admin は作業員ごとに解決した値、index は元の行から引き継いだ値を送る
      expect(src).toMatch(/yakinTeate:_y\.teate|\.\.\.inheritYakinFlags\(editIds,m\.name\)/);
      expect(src).toMatch(/yakinSeikyu:_y\.seikyu|\.\.\.inheritYakinFlags\(editIds,m\.name\)/);
    });

    it('★編集で置き換える前の行から引き継ぐ仕組みがある', () => {
      expect(src).toContain('function inheritYakinFlags(ids,name)');
      expect(src).toContain("src?String(src.yakinTeate||''):''");
      expect(src).toContain("src?String(src.yakinSeikyu||''):''");
    });

    it('グループ化でも作業員ごとに値を持ち回る（依頼書2: 作業員ごとの上書き）', () => {
      expect(src).toContain("yakinSelf:!!n.yakin");
      expect(src).toContain("yakinTeate:String(n.yakinTeate||'')");
      expect(src).toContain("yakinSeikyu:String(n.yakinSeikyu||'')");
    });

    it('送信前の楽観表示にも載せている（送信が終わるまで消えて見えない）', () => {
      expect(src).toContain("yakinTeate:String(r.yakinTeate||''),yakinSeikyu:String(r.yakinSeikyu||'')");
    });
  });
});

describe('admin.html: 作業員ごとの夜勤設定パネル（依頼書2・3）', () => {
  const src = read('admin.html');

  it('パネルの入れ物がある', () => {
    expect(src).toContain('id="e-yakin-panel"');
    expect(src).toContain('id="e-yakin-rows"');
  });

  it('勤務区分・手当・請求の3つを作業員ごとに選べる', () => {
    expect(src).toContain("sel(name,'mode',[['','現場に合わせる'],['夜勤','夜勤'],['日勤','日勤']])");
    expect(src).toContain("sel(name,'teate',[['','自動'],['対象','対象'],['対象外','対象外']])");
    expect(src).toContain("sel(name,'seikyu',[['','自動'],['対象','対象'],['対象外','対象外']])");
  });

  it('★モーダルを開くときに保存済みの値を読み戻す（読み戻さないと開くだけで消える）', () => {
    expect(src).toContain('function loadEditYakinOverrides(g)');
    // 単体編集・一括編集の両方の入口で呼ぶ
    expect(src.split('loadEditYakinOverrides(g);').length - 1).toBe(2);
  });

  // ★2026-09-04 検品（実際に押して目で見た）で出た2件。両方ここで止める。
  it('★氏名を onchange の文字列へ埋め込まない（アポストロフィを含む名前で壊れた）', () => {
    // JSの "\\'" は ' そのものなので、.replace(/'/g,"\\'") は何も escape しない。
    // 名前ではなく並び順の番号で引く形にした。
    expect(src).not.toContain('setEditYakinOverride(\'${String(name)');
    expect(src).toContain('setEditYakinOverrideAt(${i}');
    expect(src).toContain('function setEditYakinOverrideAt(i,field,value)');
  });

  it('★画面に出す氏名とラベルを esc() に通す', () => {
    const i = src.indexOf('function renderEditYakinPanel()');
    const body = src.slice(i, src.indexOf('\n}', i));
    expect(body).toContain('${esc(name)}');
    expect(body).toContain('${esc(label)}');
  });

  it('★選択欄に width:auto を付ける（付けないと1人で3行ぶん縦に伸びる）', () => {
    // このアプリは先頭のCSSで select を全部 width:100% にしている。
    expect(src).toContain("const SELST='width:auto;margin:0;");
  });

  it('責任者を変えても班員を足し引きしてもパネルが追従する', () => {
    expect(src).toContain("onchange=\"refreshButaiField('e');renderEditYakinPanel()\"");
    expect(src).toContain('toggleEditMember);renderEditYakinPanel();');
  });

  // ★2026-09-05 Codexレビュー[P1]#3: 一括編集で、触ってもいない設定が全日へコピーされていた。
  //   9/1は「対象外」・9/2は「自動」の人を2日まとめて時刻だけ直すと、片方が黙って消えた。
  it('★触った欄だけを全日へ適用し、触っていない欄は日ごとの元の値を残す', () => {
    const i = src.indexOf('function resolveMemberYakin(name,groupYakin,date)');
    expect(i).toBeGreaterThan(-1);
    const body = src.slice(i, src.indexOf('\n}', i));
    expect(body).toContain('const t=editYakinTouched[name]||{};');
    expect(body).toContain('t[field] ?');
    // 元の行は「その人・その日」で探す（日をまたいで潰さない）
    expect(body).toContain('n.name===name&&n.date===date');
  });

  it('★🌙夜勤を切り替えたときだけ全員へ適用する（触っていないなら各人の元の区分を保つ）', () => {
    const i = src.indexOf('function resolveMemberYakin(name,groupYakin,date)');
    const body = src.slice(i, src.indexOf('\n}', i));
    expect(body).toContain("ov.mode==='夜勤'?true: ov.mode==='日勤'?false: !!groupYakin");
    expect(body).toContain('const switched=(!!groupYakin)!==(!!editOrigGroupYakin);');
    expect(body).toContain('switched ? !!groupYakin : (src?!!src.yakin:!!groupYakin)');
  });

  it('★開き直したら「触った」印が消える（前回の操作が残らない）', () => {
    expect(src).toContain('editYakinTouched={};');
    expect(src).toContain('editOrigGroupYakin=!!g.yakin;');
  });
});

// ★2026-09-05 Codexレビュー[P1]#2:
//   現場画面はUIが無いぶん、保存時に全員へグループの夜勤フラグを配っていた。
//   「Aさんは日勤・Bさんは夜勤」の予定を現場でメモだけ直すと、全員同じに潰れた。
describe('index.html: 現場画面の編集で作業員ごとの夜勤区分を潰さない', () => {
  const src = read('index.html');

  it('★保存時に作業員ごと・日ごとに解決する', () => {
    expect(src).toContain('function resolveMemberYakinSite(ids,name,date,groupYakin)');
    expect(src).toContain('const _y=resolveMemberYakinSite(editIds,m.name,date,yakin);');
    expect(src).toContain('yakin:_y.yakin');
  });

  it('★🌙を切り替えたときだけ全員へ適用する', () => {
    const i = src.indexOf('function resolveMemberYakinSite(ids,name,date,groupYakin)');
    const body = src.slice(i, src.indexOf('\n}', i));
    expect(body).toContain('const switched = (!!groupYakin)!==(!!editOrigYakin);');
    expect(body).toContain('switched ? !!groupYakin : (src?!!src.yakin:!!groupYakin)');
  });

  it('★元の行は「その人・その日」で探す（一括編集で他の日を潰さない）', () => {
    const i = src.indexOf('function resolveMemberYakinSite(ids,name,date,groupYakin)');
    const body = src.slice(i, src.indexOf('\n}', i));
    expect(body).toContain('n.name===name&&n.date===date');
  });

  it('編集画面を開くときに元の夜勤フラグを覚える（単体編集・一括編集の両方）', () => {
    expect(src.split('editOrigYakin=!!g.yakin;').length - 1).toBe(2);
  });
});

describe('index.html: 現場画面には給与に関わるUIを出さない', () => {
  const src = read('index.html');
  it('夜勤手当の設定パネルは現場画面に無い（職人マスタの単価と同じ扱い）', () => {
    expect(src).not.toContain('id="e-yakin-panel"');
  });
});
