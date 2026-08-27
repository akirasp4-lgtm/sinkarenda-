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

describe('保存時の重複警告が新ルールを使っていること（2026-08-27 フェーズ2）', () => {
  FILES.forEach(f => {
    const src = read(f);

    it(f + ': 古い「同じ日に1件でもあれば警告」が残っていない', () => {
      // 旧ルールは実データ250件が該当し、その大半が「現場＋事務所」等の正常。
      // 毎回出る警告は読まれなくなるので、文言ごと消えていることを固定する
      expect(src).not.toContain('既に予定が入っています');
      expect(src).not.toContain('既に他の予定が入っています');
    });

    it(f + ': 判定ブロックが入っている', () => {
      expect(src).toContain('// ===== PHASE2-CONFLICT-RULE:BEGIN =====');
      expect(src).toContain('// ===== PHASE2-CONFLICT-RULE:END =====');
    });

    it(f + ': ★新規登録と編集の両方が新ルールを通る（片方だけ直すのを防ぐ）', () => {
      expect((src.match(/conflictsIfAdded\(/g) || []).length).toBe(3); // 定義1 + 呼び出し2
    });

    it(f + ': ★重複の母集団は拠点で絞らない', () => {
      // filteredNippos() を使うと、本社ビューのとき関東の現場と重なっても気付けない
      expect(src).toContain('function companyNippos(');
      expect(src).not.toMatch(/conflictsIfAdded\(filteredNippos\(\)/);
    });

    it(f + ': 車両のダブルブッキング判定は今までどおり残っている', () => {
      expect(src).toContain('既に予約があります');
    });
  });
});

describe('重複の知らせ（2026-08-27 フェーズ2 Task3）', () => {
  const src = read('index.html');

  it('★母集団は拠点で絞らない（本社ビューでも関東の重なりに気付ける）', () => {
    expect(src).toMatch(/findConflicts\(companyNippos\(\),\s*\{\s*from:\s*todayYmd\(\)\s*\}\)/);
  });

  it('★今日以降だけを出す（過ぎた日の重複は直しようがない）', () => {
    expect(src).toContain('function todayYmd(');
    expect(src).toContain('function currentConflicts(');
  });

  it('0件のときはバナーを消す（常に出ていると読まれなくなる）', () => {
    expect(src).toMatch(/if\(!list\.length\)\{el\.style\.display='none';return;\}/);
  });

  it('カレンダーの描画から呼ばれている', () => {
    expect(src).toMatch(/try\{renderConflictBanner\(\);\}catch\(e\)\{\}/);
  });

  it('★知らせが落ちてもカレンダー本体は描く', () => {
    // 重複の計算で例外が出ても、予定表そのものが真っ白になってはいけない
    expect(src).toMatch(/try\{renderConflictBanner\(\);\}catch\(e\)\{\}[^\n]*\n\s*renderCalendar\(\);/);
  });

  it('★一覧から編集・削除をさせない（複数の現場と日にまたがるため）', () => {
    const m = src.match(/function openConflictList\(\)\{[\s\S]*?\n\}/);
    expect(m).toBeTruthy();
    ['openEditModal', 'deleteNippo', 'gm-delete-bar', 'gm-edit-btn', 'toggleGmSelect']
      .forEach(bad => expect(m[0], bad).not.toContain(bad));
  });

  it('氏名・現場名をそのままHTMLに入れていない（escを通している）', () => {
    const m = src.match(/function openConflictList\(\)\{[\s\S]*?\n\}/);
    expect(m[0]).toContain('esc(c.name)');
    expect(m[0]).toContain('esc(j.genba)');
  });

  it('★新しいUIに class="tab" を増やしていない（下部ナビの添字がずれる）', () => {
    // switchTab() は querySelectorAll('.tab') と tabs配列を添字で対応させている。
    // .tab が1つ増えるだけで下部ナビの選択表示が全部ずれる
    expect((read('index.html').match(/class="tab"/g) || []).length).toBe(4);
    expect((read('admin.html').match(/class="tab"/g) || []).length).toBe(6);
  });
});

describe('重複判定の名簿のまとまり（2026-08-27 実機で発見した欠陥の再発防止）', () => {
  FILES.forEach(f => {
    it(f + ': ★判定側の名簿一覧が KYOTEN_COMPANIES と同じ内容', () => {
      // ここがずれると「本社⇔関東をまたぐ応援の重複」を丸ごと見逃す。
      // 実機で実測: 会社で分けると39件、1つの名簿として見ると47件（差の8件は全部本物）
      const src = read(f);
      const a = src.match(/const KYOTEN_COMPANIES=\[([^\]]*)\]/);
      const b = src.match(/var CONFLICT_SAME_ROSTER = \[([^\]]*)\]/);
      expect(a, 'KYOTEN_COMPANIES が見つからない').toBeTruthy();
      expect(b, 'CONFLICT_SAME_ROSTER が見つからない').toBeTruthy();
      const norm = (s) => s.split(',').map(x => x.trim().replace(/^'|'$/g, '')).filter(Boolean).sort();
      expect(norm(b[1])).toEqual(norm(a[1]));
    });

    it(f + ': ★会社そのままで突き合わせていない', () => {
      const src = read(f);
      expect(src).toContain('rosterKey(n.company)');
      expect(src).not.toMatch(/var key = \[String\(n\.date\), String\(n\.company/);
    });
  });
});

describe('詳しく探す（2026-08-27 フェーズ2 Task4）', () => {
  const src = read('index.html');

  it('4つ目のモードがある', () => {
    expect(src).toContain("setGmMode('search')");
    expect(src).toContain('id="gm-filter-search"');
    expect(src).toContain('id="gm-mode-search"');
  });

  it('★モードの一覧に search が入っている（入れ忘れるとボタンの色が戻らない）', () => {
    expect(src).toContain("['genba','person','pin','search'].forEach");
    expect(src).toMatch(/gmMode === 'search'/);
  });

  it('★絞り込みの軸がそろっている', () => {
    ['gm-sc-name', 'gm-sc-butai', 'gm-sc-genba', 'gm-sc-worktype', 'gm-sc-loc', 'gm-sc-from', 'gm-sc-to']
      .forEach(id => expect(src, id).toContain('id="' + id + '"'));
  });

  it('★結果から一括編集・一括削除をさせない', () => {
    // 検索結果は複数の現場・複数の日にまたがる。そこで一括削除を押せると
    // 関係のない予定まで消える
    const m = src.match(/function renderSearchResults\(\)\s*\{[\s\S]*?\n\}/);
    expect(m, 'renderSearchResults が見つからない').toBeTruthy();
    ['gm-delete-bar', 'gm-edit-btn', 'toggleGmSelect', 'openBulkEditModal', 'deleteGmChecked']
      .forEach(bad => expect(m[0], bad).not.toContain(bad));
  });

  it('★検索の結果枠と、一括編集付きの一覧を同時に出さない', () => {
    expect(src).toContain("sr.style.display = (mode === 'search') ? '' : 'none'");
    expect(src).toContain('id="gm-search-result"');
  });

  it('氏名・現場名をそのままHTMLに入れていない', () => {
    const m = src.match(/function renderSearchResults\(\)\s*\{[\s\S]*?\n\}/);
    expect(m[0]).toContain('esc(g.genba)');
    expect(m[0]).toContain('esc(m.name)');
  });

  it('部隊の選択肢は BUTAI_VALUES を使う（手打ちで増やさない）', () => {
    expect(src).toContain("fill('gm-sc-butai', BUTAI_VALUES.slice())");
  });

  it('★拠点の絞り込みを二重に置いていない（見出し下の切替が唯一の拠点操作）', () => {
    expect(src).not.toContain('id="gm-sc-kyoten"');
    expect(src).toContain('function searchNippos(');
    expect(src).toMatch(/function searchNippos\(\)\s*\{[\s\S]*?filteredNippos\(\)/);
  });
});

describe('空き人員の名前リスト（2026-08-27 フェーズ2 Task5・要件4と要件8後半）', () => {
  const src = read('index.html');

  it('日付を選ぶ欄と結果の枠がある', () => {
    expect(src).toContain('id="avail-day"');
    expect(src).toContain('id="avail-day-result"');
    expect(src).toContain('function renderAvailDay(');
  });

  it('★母集団は拠点で絞らない（関東の現場に入っている人を「空き」にしない）', () => {
    const m = src.match(/function renderAvailDay\(\)\s*\{[\s\S]*?\n\}/);
    expect(m).toBeTruthy();
    expect(m[0]).toContain('companyNippos()');
    expect(m[0]).not.toContain('filteredNippos()');
  });

  it('★名簿は有効な人だけ（職人マスタで無効にした人を空きに数えない）', () => {
    const m = src.match(/function renderAvailDay\(\)\s*\{[\s\S]*?\n\}/);
    expect(m[0]).toContain('getActiveShokunin()');
  });

  it('★同じ日に休みと出勤が両方あるときは「出勤」とみなす（空きに数えない）', () => {
    const m = src.match(/function renderAvailDay\(\)\s*\{[\s\S]*?\n\}/);
    expect(m[0]).toContain("if (cur === 'busy') return;");
  });

  it('延期・中止になった現場の人を別枠で出す（要件8の後半）', () => {
    expect(src).toContain('function releasedByStatus(');
    expect(src).toMatch(/st === '延期' \|\| st === '中止'/);
  });

  it('★元請と現場名を区切ってから突き合わせる', () => {
    const m = src.match(/function releasedByStatus\(date\)\s*\{[\s\S]*?\n\}/);
    expect(m[0]).toContain("String(j.genba || '') + '|' + String(j.loc || '')");
  });

  it('データを読み直したときにも描き直す', () => {
    expect(src).toContain('try{renderAvailDay();}catch(e){}if(vehicleWeekStart)');
    expect(src).toContain("if(t==='avail'){renderAvailList();try{renderAvailDay();}catch(e){}}");
  });

  it('氏名をそのままHTMLに入れていない', () => {
    const m = src.match(/function renderAvailDay\(\)\s*\{[\s\S]*?\n\}/);
    expect(m[0]).toContain('esc(n)');
  });
});
