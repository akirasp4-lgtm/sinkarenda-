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

  it('★下部ナビのボタンの並びと switchTab の tabs配列が完全に一致する', () => {
    // switchTab() は querySelectorAll('.tab') と tabs配列を**添字で**対応させている。
    // 1つずれるだけで「押したタブと違う画面が開く」。
    // ★2026-08-29: 以前は個数を決め打ちで数えていたが、それだと
    //   「両方を同時に間違えた」場合を通してしまう。実際の並びどうしを突き合わせる。
    ['index.html', 'admin.html'].forEach(f => {
      const src = read(f);
      // 下部ナビのボタン（class="tab" / class="tab active"）を出てくる順に拾う
      const buttons = [...src.matchAll(/<button class="tab(?: active)?"[^>]*onclick="switchTab\('([^']+)'\)"/g)]
        .map(m => m[1]);
      const arr = src.match(/const tabs=\[([^\]]*)\]/);
      expect(arr, f + ' に tabs配列が無い').toBeTruthy();
      const names = arr[1].split(',').map(x => x.trim().replace(/^'|'$/g, '')).filter(Boolean);
      expect(buttons.length, f + ': ボタンが1つも拾えていない').toBeGreaterThan(0);
      expect(names, f + ': タブボタンの並びと tabs配列がずれている').toEqual(buttons);
    });
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
    // ★2026-08-28: 資格を会社込みで引くため、renderAvailDay は
    //   activeRosterMembers() 経由になった。無効の人を外す責任はその中の
    //   getActiveShokunin() が持つので、2段とも確かめる。
    const m = src.match(/function renderAvailDay\(\)\s*\{[\s\S]*?\n\}/);
    expect(m[0]).toContain('activeRosterMembers()');
    const r = src.match(/function activeRosterMembers\(\)\s*\{[\s\S]*?\n\}/);
    expect(r, 'activeRosterMembers が無い').toBeTruthy();
    expect(r[0]).toContain('getActiveShokunin()');
  });

  it('★分類は共通関数 dayStateByName を使う（経営の画面と食い違わせない）', () => {
    // ★2026-08-29: 以前はここで renderAvailDay の中に書かれた分岐を文字で見張っていた。
    //   経営の画面が同じ分類を別に書いたので、1つの関数にまとめた。
    //   優先順位そのもの（休みと出勤が両方なら出勤）は phase4-dash.test.js が
    //   vm で実際に動かして検査している。
    const m = src.match(/function renderAvailDay\(\)\s*\{[\s\S]*?\n\}/);
    expect(m[0]).toContain('dayStateByName(');
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

describe('管理画面にもフェーズ2のUIが入っていること（2026-08-28）', () => {
  // ★2026-08-28 実機で発見: フェーズ2で作った3つが社員用にしか入っていなかった。
  //   重複を直すのも人を配置するのも管理する人の仕事なので、むしろ管理画面に要る。
  const src = read('admin.html');

  it('★重複の知らせ（バナーと一覧）がある', () => {
    expect(src).toContain('id="conflict-bar"');
    expect(src).toContain('id="conflict-modal"');
    expect(src).toContain('function renderConflictBanner(');
    expect(src).toContain('function openConflictList(');
    expect(src).toMatch(/findConflicts\(companyNippos\(\),\s*\{\s*from:\s*todayYmd\(\)\s*\}\)/);
  });

  it('★カレンダーの描画から呼ばれている', () => {
    expect(src).toMatch(/try\{renderConflictBanner\(\);\}catch\(e\)\{\}[^\n]*\n\s*renderCalendar\(\);/);
  });

  it('★空き人員（日付を選ぶと名前が出る）がある', () => {
    expect(src).toContain('id="avail-day"');
    expect(src).toContain('id="avail-day-result"');
    expect(src).toContain('function renderAvailDay(');
    expect(src).toContain('function releasedByStatus(');
    expect(src).toContain('try{renderAvailDay();}catch(e){}if(vehicleWeekStart)');
  });

  it('★詳しく探すがある（既存の完了管理を消していない）', () => {
    expect(src).toContain("setGmMode('search')");
    expect(src).toContain('id="gm-filter-search"');
    expect(src).toContain('id="gm-search-result"');
    expect(src).toContain("setGmMode('done')");      // 元からある完了管理
    expect(src).toContain("['genba','person','pin','done','search'].forEach");
  });

  it('★検索結果に一括編集・一括削除を出さない', () => {
    const m = src.match(/function renderSearchResults\(\)\s*\{[\s\S]*?\n\}/);
    expect(m).toBeTruthy();
    ['gm-delete-bar', 'gm-edit-btn', 'toggleGmSelect', 'openBulkEditModal', 'deleteGmChecked']
      .forEach(bad => expect(m[0], bad).not.toContain(bad));
  });

  it('★判定の中身が社員用と1文字も違わない（片方だけ直すのを防ぐ）', () => {
    // ★正規表現だと入れ子の { } で早く切れるので、括弧を数えて取る
    const pick = (s, name) => {
      const start = s.indexOf('function ' + name + '(');
      if (start < 0) return null;
      let depth = 0;
      for (let i = s.indexOf('{', start); i < s.length; i++) {
        if (s[i] === '{') depth++;
        else if (s[i] === '}') { depth--; if (depth === 0) return s.slice(start, i + 1); }
      }
      return null;
    };
    // ★2026-08-28 フェーズ3: updateQualSelect（資格プルダウン）も同一を強制する
    ['renderConflictBanner', 'currentConflicts', 'renderAvailDay', 'releasedByStatus', 'searchNippos',
     'updateQualSelect', 'activeRosterMembers', 'todayYmd', 'updateGenbaRankSelect']
      .forEach(fn => {
        const a = pick(read('index.html'), fn), b = pick(src, fn);
        expect(a, 'index.html に ' + fn + ' が無い').toBeTruthy();
        expect(b, 'admin.html に ' + fn + ' が無い').toBeTruthy();
        expect(b, fn + ' が2画面で食い違っている').toBe(a);
      });
  });
});

// ============================================================
// 資格（2026-08-28 フェーズ3の土台）
// ★配線の見張り。判定そのものは phase3-qual.test.js が vm で実際に動かす。
// ============================================================
describe('資格の配線', () => {
  FILES.forEach(f => {
    const s = read(f);
    it(f + ': 受け取り・端末キャッシュ・復元の3か所すべてに資格が書いてある', () => {
      // ★1か所でも書き忘れると「初回だけ出る」「2回目から消える」という
      //   再現しにくい不具合になる（既定部隊で実際にやらかした）
      // ★どの入口も qualSafe を通すこと（免許番号が端末へ焼き付くのを防ぐ3枚目の歯止め）
      expect(s, '受け取り').toContain('allQuals=qualSafe(Array.isArray(json.qualifications)');
      expect(s, 'localStorageへ保存').toContain('qualifications:qualSafe(json.qualifications||[])');
      expect(s, 'localStorageから復元').toContain('allQuals=qualSafe(s.qualifications||[])');
    });
    it(f + ': 資格の絞り込みプルダウンが画面にある', () => {
      expect(s).toContain('id="avail-qual"');
      expect(s).toContain('onchange="renderAvailDay()"');
    });
    it(f + ': ★免許番号・正式氏名を画面に出そうとしていない（個人情報）', () => {
      // ★コメント行は除いて見る。説明として名前を書くのは構わないが、
      //   コードとして参照していたら個人情報が画面に出る。
      const code = s.split('\n').filter(L => !/^\s*(\/\/|\*|\/\*)/.test(L)).join('\n');
      ['免許番号', '正式氏名', 'licenseNo', '.license'].forEach(bad => {
        expect(code, bad + ' を参照している').not.toContain(bad);
      });
      // 資格のオブジェクトから読んでよいのは4つだけ
      expect(s).not.toMatch(/q\.(licence|license|number|seishiki)/);
    });
    it(f + ': 資格の判定ブロックがある', () => {
      expect(s).toContain('// ===== PHASE3-QUAL-RULE:BEGIN =====');
      expect(s).toContain('// ===== PHASE3-QUAL-RULE:END =====');
    });
  });
});

// ============================================================
// 「今日」は日本時間（2026-08-28・Codexレビューの指摘で追加）
// ★端末の時計が日本以外だと「今日」が1日ずれ、重複バナーと資格の期限が
//   両方おかしくなる。実際に動かして、日本時間の日付になることを確かめる。
// ============================================================
describe('todayYmd は日本時間', () => {
  FILES.forEach(f => {
    const src = read(f);
    const body = (() => {
      const start = src.indexOf('function todayYmd(');
      let depth = 0;
      for (let i = src.indexOf('{', start); i < src.length; i++) {
        if (src[i] === '{') depth++;
        else if (src[i] === '}') { depth--; if (depth === 0) return src.slice(start, i + 1); }
      }
      return null;
    })();

    it(f + ': ★日本時間で日付を出す（UTCの夕方＝日本の翌日でずれない）', async () => {
      const vm = await import('node:vm');
      const sandbox = vm.createContext({ Intl, Date, String, console });
      sandbox.globalThis = sandbox;
      vm.runInContext(body, sandbox, { filename: f });
      // 2026-08-28 20:00 UTC ＝ 日本時間では 2026-08-29 05:00
      const RealDate = Date;
      sandbox.Date = class extends RealDate {
        constructor(...a) { return a.length ? new RealDate(...a) : new RealDate('2026-08-28T20:00:00Z'); }
      };
      expect(sandbox.todayYmd(), 'UTCのまま数えている').toBe('2026-08-29');
    });

    it(f + ': 形は YYYY-MM-DD', () => {
      expect(body).toContain("timeZone:'Asia/Tokyo'");
      expect(body).toContain('en-CA');   // en-CA が YYYY-MM-DD を返す
    });
  });
});

// ============================================================
// 経営の画面（依頼文10番）2026-08-29
// ★数え方そのものは phase4-dash.test.js が vm で実際に動かす。ここは配線の見張り。
// ============================================================
describe('経営の画面（管理画面のみ）', () => {
  const src = read('admin.html');

  it('画面・タブ・描画の呼び出しがそろっている', () => {
    expect(src, '画面の箱').toContain('id="screen-dash"');
    expect(src, '中身の入れ物').toContain('id="dash-body"');
    expect(src, 'タブボタン').toContain("switchTab('dash')");
    expect(src, 'タブを開いたときの描画').toMatch(/if\(t==='dash'\)renderDash\(\)/);
    expect(src, '判定ブロック').toContain('// ===== PHASE4-DASH-RULE:BEGIN =====');
  });

  it('★データが届いたあとにも数字を出し直す（初回だけ空になるのを防ぐ）', () => {
    const m = src.match(/function rerenderAll\(\)\{[\s\S]{0,400}/);
    expect(m).toBeTruthy();
    expect(m[0], 'rerenderAll から renderDash を呼んでいない').toContain('renderDash()');
  });

  it('★依頼文10番の8項目がすべて画面に出ている', () => {
    const m = src.match(/function renderDash\(\)\{[\s\S]*?\n\}/);
    expect(m).toBeTruthy();
    const f = m[0];
    ['稼働', '空き', '現場の数', '重複警告', '未確定案件',   // 今日の5つ
     '人員稼働率', '案件'].forEach(w => expect(f, w + ' が無い').toContain(w));
    expect(f, '今週の日別（空き予定）').toContain('dashWeek(');
  });

  it('★重複警告はフェーズ2の判定を使い回す（別の数え方をしない）', () => {
    const f = src.match(/function renderDash\(\)\{[\s\S]*?\n\}/)[0];
    expect(f).toContain('currentConflicts()');
  });

  it('★空き人数は空き確認と同じ名簿から出す', () => {
    const f = src.match(/function renderDash\(\)\{[\s\S]*?\n\}/)[0];
    expect(f).toContain('getActiveShokunin()');
  });

  it('氏名・現場名をそのままHTMLに入れていない（escを通している）', () => {
    const f = src.match(/function renderDash\(\)\{[\s\S]*?\n\}/)[0];
    // 画面に出す外部由来の文字は会社名・拠点・日付。すべて esc を通す
    expect(f).toContain('esc(currentCompany)');
    expect(f).toContain('esc(ymd)');
  });

  it('★職人用（index.html）には入れない（経営の数字は現場に出さない）', () => {
    const idx = read('index.html');
    expect(idx).not.toContain('id="screen-dash"');
    expect(idx).not.toContain('PHASE4-DASH-RULE');
  });
});

// ============================================================
// 経営の画面：Codexレビューで見つかった穴の再発防止（2026-08-29）
// ============================================================
describe('経営の画面：Codexレビューの[P1]の再発防止', () => {
  const src = read('admin.html');

  it('★[P1] データが届く経路そのものから renderDash を呼んでいる', () => {
    // ★以前は rerenderAll() の中だけを見張っていたが、loadData の成功経路は
    //   rerenderAll() を通らない。つまり「テストは緑・画面はずっと0」だった。
    //   数字が変わりうる3つの入口すべてを見張る。
    const i = src.indexOf('saveSnapshot(requestCompany,json)');
    expect(i, 'loadData の成功経路が見つからない').toBeGreaterThan(0);
    expect(src.slice(i - 600, i), 'データ取得の直後に renderDash を呼んでいない')
      .toContain('renderDash()');

    const body = (name) => {
      const st = src.indexOf('function ' + name + '(');
      let d = 0;
      for (let k = src.indexOf('{', st); k < src.length; k++) {
        if (src[k] === '{') d++;
        else if (src[k] === '}') { d--; if (d === 0) return src.slice(st, k + 1); }
      }
      return '';
    };
    expect(body('rerenderAll'), 'rerenderAll から呼んでいない').toContain('renderDash()');
    expect(body('switchCompany'), '会社を切り替えた直後に前の会社の数字が残る')
      .toContain('renderDash()');
  });

  it('★[P1] 未確定案件を今の会社の元請だけに絞っている（他社の見積中を混ぜない）', () => {
    const f = src.match(/function renderDash\(\)\{[\s\S]*?\n\}/)[0];
    expect(f, '会社で絞っていない').toContain('getGenbaMasterNames()');
    expect(f, '全件をそのまま数えている').not.toMatch(/dashUnconfirmed\(allJobsites\)/);
  });

  it('★[P1] 拠点で数字は変わらないので、拠点名をラベルに出さない', () => {
    const f = src.match(/function renderDash\(\)\{[\s\S]*?\n\}/)[0];
    expect(f, '拠点名を出すと「拠点で絞った数字」に見える').not.toContain('esc(currentKyoten)');
    expect(f).toContain('本社＋関東支店の合計');
  });

  it('★[P2] 現場として数えない作業区分がはっきり決まっている', () => {
    expect(src).toContain('var DASH_NOT_SITE');
    const f = src.match(/function dashIsSite\([\s\S]*?\n\}/)[0];
    expect(f).toContain('DASH_NOT_SITE');
  });
});

describe('無効の人の判定は会社ごと（2026-08-29 Codexレビュー[P2]）', () => {
  FILES.forEach(f => {
    it(f + ': ★他社の同姓同名が無効でも、こちらの人が消えない', () => {
      // 奥田さんはグローライズとGRHDの両方に実在する
      const src = read(f);
      const m = src.match(/function getActiveShokunin\(\)\{[\s\S]*?\n\}/);
      expect(m).toBeTruthy();
      expect(m[0], '会社を見ずに氏名だけで無効にしている').toContain('hasKyotenAxis');
    });
  });
});

describe('タブと画面の対応（2026-08-29 Codexレビュー[P3]）', () => {
  FILES.forEach(f => {
    it(f + ': ★タブ名に対応する画面が1つずつ存在する', () => {
      const src = read(f);
      const arr = src.match(/const tabs=\[([^\]]*)\]/);
      const names = arr[1].split(',').map(x => x.trim().replace(/^'|'$/g, '')).filter(Boolean);
      names.forEach(t => {
        const hits = (src.match(new RegExp('id="screen-' + t + '"', 'g')) || []).length;
        expect(hits, f + ': screen-' + t + ' が ' + hits + ' 個ある（1個であるべき）').toBe(1);
      });
    });
  });
});

describe('経営の画面では拠点バーを出さない（2026-08-29 Codexレビュー[P1]）', () => {
  it('★押しても数字が変わらないバーを出さない（拠点で絞った数字だと誤解される）', () => {
    const src = read('admin.html');
    const m = src.match(/function renderKyotenBar\(\)\{[\s\S]*?\n\}/);
    expect(m).toBeTruthy();
    expect(m[0], '経営の画面でバーを隠していない').toContain("active.id==='screen-dash'");
    // ★他の画面で選んでいた拠点を勝手に戻さないこと
    const seg = m[0].slice(m[0].indexOf("screen-dash"), m[0].indexOf("screen-dash") + 120);
    expect(seg, '拠点の選択を勝手に戻している').not.toContain('currentKyoten=');
  });
});

// ============================================================
// 候補者を出す（依頼文の要件5）2026-08-29
// ★AIは使わない。空き × 資格 × 元請の経験で出す＝0円。
// ============================================================
describe('候補者の配線', () => {
  FILES.forEach(f => {
    const src = read(f);
    it(f + ': 元請で並べ替えるプルダウンがある', () => {
      expect(src).toContain('id="avail-genba"');
      expect(src).toContain('function updateGenbaRankSelect(');
      expect(src).toContain('// ===== PHASE5-PICK-RULE:BEGIN =====');
    });
    it(f + ': ★空きの一覧を経験順に並べている', () => {
      const m = src.match(/function renderAvailDay\(\)\s*\{[\s\S]*?\n\}/);
      expect(m[0]).toContain('rankCandidates(');
      expect(m[0]).toContain("searchFilterValue('avail-genba')");
    });
    it(f + ': ★候補から予定を勝手に作らない（依頼文「最終決定は管理者が行う」）', () => {
      const m = src.match(/function renderAvailDay\(\)\s*\{[\s\S]*?\n\}/);
      ['saveNippo', 'submitNippo', 'action:\'add\'', 'openEditModal']
        .forEach(bad => expect(m[0], bad + ' を呼んでいる').not.toContain(bad));
    });
  });
});
