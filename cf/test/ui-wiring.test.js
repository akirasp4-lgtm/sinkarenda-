// 画面（index.html / admin.html）の「配線もれ」を機械的に見張るテスト。
//
// ★なぜ必要か（2026-08-27 に実際にやらかした）:
//   部隊を足したとき、members 配列には butai を載せたのに、
//   GASへ送る rows.push({...}) の方に載せ忘れていた。
//   その結果「画面では選べるのに1件も保存されない」状態になっていた。
//   拠点のときも parseRows / 編集モーダル / 一括編集で同じ取りこぼしが起きている。
//
//   1つの項目を足すときに触る場所が多すぎて、目視では必ず漏れる。
//   「拠点があるところには部隊もあるはず」という不変条件で見張る。
import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const here = dirname(fileURLToPath(import.meta.url));
const read = (f) => readFileSync(join(here, '..', '..', f), 'utf8');

const FILES = ['index.html', 'admin.html'];

describe('画面の配線もれ（拠点があるところには部隊もあること）', () => {
  FILES.forEach(f => {
    describe(f, () => {
      const src = read(f);

      it('★GASへ送る行に butai が載っている（無いと1件も保存されない）', () => {
        // rows.push({...kyoten:...}) の形を全部拾い、同じ括弧の中に butai があるか見る
        const pushes = src.match(/rows\.push\(\{[^}]*\}\)/g) || [];
        const withKyoten = pushes.filter(x => x.includes('kyoten:'));
        expect(withKyoten.length).toBeGreaterThan(0);
        withKyoten.forEach(x => {
          expect(x, 'rows.push に butai が無い: ' + x.slice(0, 120)).toContain('butai:');
        });
      });

      it('parseRows が部隊を読んでいる', () => {
        expect(src).toContain("butai:String(r['部隊']||'')");
      });

      it('groupNippos のグループと members が部隊を持っている', () => {
        expect(src).toContain("butai:n.butai||''");
        expect(src).toMatch(/members\.push\(\{name:n\.name,role:n\.role,butai:/);
      });

      it('端末キャッシュ（saveSnapshot）が部隊を落としていない', () => {
        expect(src).toMatch(/members:\(json\.members\|\|\[\]\)\.map\(m=>\(\{[^}]*butai:/);
      });

      it('楽観表示の行に拠点と部隊が載っている', () => {
        expect(src).toContain("kyoten:String(r.kyoten||''),butai:String(r.butai||'')");
      });

      it('単一編集で既存の部隊を読み戻している', () => {
        expect(src).toContain('function applyEditButai(');
        expect(src).toContain('applyEditButai(g)');
      });

      it('一括編集でも部隊を初期化している（別入口）', () => {
        expect(src).toContain('function applyBulkEditButai(');
        expect(src).toContain('applyBulkEditButai(targets)');
      });

      it('★一括編集で部隊が混在するとき「そのまま」を使う（空で全消しにしない）', () => {
        // Codexレビュー[P1]#2: 空文字を初期値にすると、時間だけ直したつもりで
        // 1部隊と2部隊がまとめて「部隊なし」に消える。
        expect(src).toContain("const BUTAI_KEEP='__KEEP__'");
        expect(src).toContain("el.value='__KEEP__'");
        expect(src).toContain('function resolveEditButai(');
        // 保存経路は必ず解決関数を通す（生の readButai('e') を直接使わない）
        expect(src).toContain("butai:resolveEditButai(date)");
        expect(src).not.toContain("kyoten:readKyoten('e',company),butai:readButai('e')");
      });

      it('単一編集では「そのまま」を出さない', () => {
        expect(src).toMatch(/function applyEditButai\(g\)\{[\s\S]*?keep\.style\.display='none'/);
      });

      it('責任者を選ぶと既定部隊が入る', () => {
        expect(src).toContain("refreshButaiField('s')");
        expect(src).toContain("refreshButaiField('e')");
      });
    });
  });
});

describe('役割の保存値を壊していないこと', () => {
  it('★代表/同行 の比較・代入は39箇所のまま（2026-08-27 実測の基準値）', () => {
    // 呼称は表示だけ変える方針。保存値を変えると gas.js:477 の代表者判定が壊れ、
    // 事業部が決まらず工番が空で保存される。
    const all = ['gas.js', 'index.html', 'admin.html'].map(read).join('\n');
    const re = /role:'代表'|role:'同行'|role==='代表'|role==='同行'|role === '代表'|=== '代表'|=== '同行'|==='代表'|==='同行'/g;
    expect((all.match(re) || []).length).toBe(39);
  });

  it('画面に「代表者」「同行メンバー」という表示ラベルが残っていない', () => {
    FILES.forEach(f => {
      const src = read(f);
      expect(src, f).not.toContain('<label>代表者');
      expect(src, f).not.toContain('代表者を選択してください');
      expect(src, f).not.toContain('同行メンバー（タップで選択）');
    });
  });
});

describe('案件ステータスの旧経路が残っていないこと', () => {
  it('★職人用に「完了にする」ボタンが無い（管理者の延期を上書きするため）', () => {
    const src = read('index.html');
    expect(src).not.toContain('この現場を完了にする');
    expect(src).not.toContain('toggleSiteCompletion_(');
  });

  it('管理者も8段階の set_site_status に一本化されている', () => {
    const src = read('admin.html');
    expect(src).not.toContain('toggleSiteCompletion_(');
    expect(src).not.toContain('doneToggle(');
    expect(src).toContain("action:'set_site_status'");
  });
});
