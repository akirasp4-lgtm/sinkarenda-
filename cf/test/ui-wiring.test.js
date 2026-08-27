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

describe('二重登録の統合に画面が追随していること（2026-08-27）', () => {
  FILES.forEach(f => {
    it(f + ': 端末に保存された操作者名を読み替える', () => {
      const src = read(f);
      // これが無いと、名前を変えた人の端末だけ古い名前で「更新者」を書き続け、
      // せっかくまとめた名前がまた2つに割れる
      expect(src).toContain('function migrateUsername(');
      expect(src).toContain('currentUsername=migrateUsername(_raw)');
      expect(src).toContain("const UPDATER_MERGE={'高田':'高田（関東）'}");
    });

    it(f + ': 第五部隊の部隊長名が統合後の名前になっている', () => {
      expect(read(f)).toContain("'第五部隊':'高田（関東）'");
    });
  });

  it('★端末側は「更新者用の表」だけを持つ（氏名は会社込みでないと判定できない）', () => {
    FILES.forEach(f => {
      const src = read(f);
      // 端末には会社の情報が無いので、氏名の読み替え表を持たせてはいけない。
      // 持たせると他社の同姓同名を巻き込む（Codexレビュー[P1]#1と同じ穴）。
      expect(src, f).not.toContain("'GRME髙田':");
      expect(src, f).not.toContain('MEMBER_MERGE_BY_COMPANY');
    });
  });

  it('★GAS側は (会社|氏名) の組で判定している', () => {
    const gas = read('gas.js');
    ['GRミツマ|高田', 'グローライズ|GRME髙田', 'GRミツマ|柳澤', 'GRミツマ|栁澤',
     'グローライズ|GRME栁澤', 'GRミツマ|内村', 'グローライズ|GRME内村']
      .forEach(k => expect(gas, 'gas.jsに ' + k + ' が無い').toContain("'" + k + "':"));
    // 実データで確認していない組は載せない
    expect(gas).not.toContain("'GRミツマ|髙田':");
    expect(gas).not.toContain("'グローライズ|GRME高田':");
  });

  it("★保存の入口でも読み替える（開いたままの端末が旧名を復活させない）", () => {
    const gas = read('gas.js');
    expect(gas).toContain('const _name = mergedMemberName_(row.company, row.name)');
    expect(gas).toContain('const _by = mergedUpdaterName_(');
  });

  it('★職人マスタを触る操作が日報と同じロックを使う', () => {
    const gas = read('gas.js');
    expect(gas).toContain('const memberMutation =');
    expect(gas).toContain('(employeeMutation || memberMutation) ? getDailyDataLock_()');
  });

  it('★admin の単価・事業部・削除が「その人の実際の会社」を送る', () => {
    const src = read('admin.html');
    expect(src).toContain("{action:'update_member_rate',name,company:_co,rate}");
    expect(src).toContain("{action:'update_member_division',name,company:_co,division}");
    expect(src).toContain("{action:'remove_member',name,company:_co}");
    expect(src).not.toContain("{action:'update_member_rate',name,company:currentCompany,rate}");
  });
});
