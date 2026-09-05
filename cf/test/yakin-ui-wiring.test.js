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

  it('責任者を変えても班員を足し引きしてもパネルが追従する', () => {
    expect(src).toContain("onchange=\"refreshButaiField('e');renderEditYakinPanel()\"");
    expect(src).toContain('toggleEditMember);renderEditYakinPanel();');
  });

  it('★「自動」に戻す操作も保存される（パネルの値が正、引き継ぎで上書きしない）', () => {
    const i = src.indexOf('function resolveMemberYakin(name,groupYakin)');
    expect(i).toBeGreaterThan(-1);
    const body = src.slice(i, src.indexOf('\n}', i));
    // パネルに行があるならパネルの値をそのまま返す（空なら空のまま）
    expect(body).toContain("teate: ov.teate||'', seikyu: ov.seikyu||''");
    // パネルに無い人（編集中に追加した人）だけ引き継ぐ
    expect(body).toContain('inheritYakinFlags(editIds,name)');
  });

  it('作業員ごとに勤務区分を上書きできる（現場全体は夜勤でもAさんだけ日勤）', () => {
    const i = src.indexOf('function resolveMemberYakin(name,groupYakin)');
    const body = src.slice(i, src.indexOf('\n}', i));
    expect(body).toContain("ov.mode==='夜勤'?true:ov.mode==='日勤'?false:!!groupYakin");
  });
});

describe('index.html: 現場画面には給与に関わるUIを出さない', () => {
  const src = read('index.html');
  it('夜勤手当の設定パネルは現場画面に無い（職人マスタの単価と同じ扱い）', () => {
    expect(src).not.toContain('id="e-yakin-panel"');
  });
});
