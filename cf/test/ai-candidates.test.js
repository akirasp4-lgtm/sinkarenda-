// AIへ渡す候補の作り方（利用者指摘 2026-08-29）
//
// 実機で2つ問題が出た:
//   ①「加藤　秀男は出向行ってるからおらん」「内藤は現場にでない」「田中も出ない」
//     → 現場に出ない人・出向中の人がAIの推薦に上がっていた。
//     出向は予定として入っているが、予定が入っていない日は「空き」になってしまう。
//   ②「ほんとに資格もってるのかわからない人も資格があるからおすすめ人選に上がってた」
//     → AIへ渡す資格の有効期限を一切見ていなかった（作りのバグ）。
//       実データで 切れている7件・期限が読めない10件 が混ざっていた。
//
// ★index.html から prepareAiSuggest の周辺だけを取り出して実際に動かす。
import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import vm from 'node:vm';

const here = dirname(fileURLToPath(import.meta.url));
const read = (f) => readFileSync(join(here, '..', '..', f), 'utf8');

// 画面から必要な部分だけ切り出して動かす小さな舞台を作る
function stage(file) {
  const src = read(file);
  const cut = (b, e) => {
    const i = src.indexOf(b), j = src.indexOf(e);
    if (i < 0 || j < 0) throw new Error('ブロックが見つからない: ' + b + ' in ' + file);
    return src.slice(i + b.length, j);
  };
  const qual = cut('// ===== PHASE3-QUAL-RULE:BEGIN =====', '// ===== PHASE3-QUAL-RULE:END =====');
  const pick = cut('// ===== PHASE5-PICK-RULE:BEGIN =====', '// ===== PHASE5-PICK-RULE:END =====');
  const i = src.indexOf('const AI_NOT_ONSITE');
  const j = src.indexOf('async function askAiSuggest');
  if (i < 0 || j < 0) throw new Error('prepareAiSuggest が見つからない in ' + file);
  const prep = src.slice(i, j);

  const box = {
    console, Map, Set, String, Object, Array, Number, Boolean, Date, JSON, Math, isNaN,
    document: { getElementById: () => null },
    // 画面から借りる部分は最小限のダミーに差し替える
    todayYmd: () => '2026-08-29',
    companyNippos: () => [],
    aiRowVisible: () => {},
    _aiCand: []
  };
  const ctx = vm.createContext(box);
  box.globalThis = box;
  vm.runInContext(qual + ';' + pick + ';' + prep +
    ';globalThis.__prep=prepareAiSuggest;globalThis.__cand=()=>_aiCand;' +
    'globalThis.__notOnsite=AI_NOT_ONSITE;', ctx, { filename: file });
  return box;
}

const FILES = ['index.html', 'admin.html'];

FILES.forEach((f) => {
  describe(f + ' — AIへ渡す候補', () => {
    const s = stage(f);

    const roster = [
      { name: '中村', company: 'グローライズ' },
      { name: '内藤', company: 'グローライズ' },
      { name: '田中（智）', company: 'グローライズ' },
      { name: '加藤　秀男', company: 'グローライズ' },
      { name: '真柄', company: 'グローライズ' }
    ];

    const idxOf = (list) => {
      const o = {};
      list.forEach((q) => {
        const k = s.qualKey(q.company, q.name);
        (o[k] || (o[k] = [])).push(q);
      });
      return o;
    };

    it('★現場に出ない人・出向中の人は候補にしない', () => {
      s.__prep(['中村', '内藤', '田中（智）', '加藤　秀男'], 'きんでん東', {}, roster);
      expect(s.__cand().map((c) => c.name)).toEqual(['中村']);
    });

    it('外す人の一覧に3人が入っている（異動したらここを直す）', () => {
      expect(s.__notOnsite).toEqual(['内藤', '田中（智）', '加藤　秀男']);
    });

    it('外した結果1人も残らなければ候補ゼロ', () => {
      s.__prep(['内藤', '田中（智）'], 'きんでん東', {}, roster);
      expect(s.__cand()).toEqual([]);
    });

    it('★切れた資格をAIへ渡さない', () => {
      const q = idxOf([
        { name: '中村', company: 'グローライズ', qual: '切れた資格', expires: '2024-05-31' },
        { name: '中村', company: 'グローライズ', qual: '生きてる資格', expires: '2030-01-01' }
      ]);
      s.__prep(['中村'], 'きんでん東', q, roster);
      expect(s.__cand()[0].quals).toEqual(['生きてる資格']);
    });

    it('★期限が読めない資格をAIへ渡さない（持っていることにしない）', () => {
      const q = idxOf([
        { name: '中村', company: 'グローライズ', qual: '読めない資格', expires: '?' },
        { name: '中村', company: 'グローライズ', qual: '生きてる資格', expires: '2030-01-01' }
      ]);
      s.__prep(['中村'], 'きんでん東', q, roster);
      expect(s.__cand()[0].quals).toEqual(['生きてる資格']);
    });

    it('期限のない資格（技能講習など）は渡す', () => {
      const q = idxOf([{ name: '中村', company: 'グローライズ', qual: '玉掛け', expires: '' }]);
      s.__prep(['中村'], 'きんでん東', q, roster);
      expect(s.__cand()[0].quals).toEqual(['玉掛け']);
    });

    it('まもなく切れる資格は渡す（まだ使える）', () => {
      const q = idxOf([{ name: '中村', company: 'グローライズ', qual: 'もうすぐ', expires: '2026-09-10' }]);
      s.__prep(['中村'], 'きんでん東', q, roster);
      expect(s.__cand()[0].quals).toEqual(['もうすぐ']);
    });

    it('★他社の同名の人の資格を混ぜない', () => {
      const q = idxOf([
        { name: '中村', company: 'GRHD', qual: '他社の資格', expires: '2030-01-01' }
      ]);
      s.__prep(['中村'], 'きんでん東', q, roster);
      expect(s.__cand()[0].quals).toEqual([]);
    });

    it('AIへ渡すのは id・経験日数・資格だけ（氏名は画面側だけで持つ）', () => {
      const q = idxOf([{ name: '中村', company: 'グローライズ', qual: '玉掛け', expires: '' }]);
      s.__prep(['中村'], 'きんでん東', q, roster);
      const c = s.__cand()[0];
      expect(c.id).toBe('c1');
      expect(Object.keys(c).sort()).toEqual(['days', 'id', 'name', 'quals']);
    });

    it('元請を選んでいなければ候補を作らない', () => {
      s.__prep(['中村'], '', {}, roster);
      expect(s.__cand()).toEqual([]);
    });
  });
});


// ============================================================
// 描く順番（コードレビュー 2026-08-30）
//
// renderAvailDay の中で chip(free,true) は availPicked を読んでチェック印を描く。
// その **後ろ** で availSyncPicked が選択を捨てる／間引くと、
// 「チップには✓が付いているのにバーは0人」という食い違いが画面に出る。
// 実際にそうなっていたので、順番をソースの並びで固定する。
// ============================================================
describe('renderAvailDay の描く順番', () => {
  ['index.html', 'admin.html'].forEach((f) => {
    const src = read(f);

    it(f + ': availSyncPicked は chip(free,true) より前に呼ぶ', () => {
      const sync = src.indexOf('availSyncPicked(freeAll)');
      const chip = src.indexOf('chip(free, true)');
      expect(sync).toBeGreaterThan(-1);
      expect(chip).toBeGreaterThan(-1);
      expect(sync).toBeLessThan(chip);
    });

    it(f + ': availSyncPicked の呼び出しは1か所だけ（二重に同期しない）', () => {
      // ★定義（function availSyncPicked(...)）も数えないよう、呼び出しだけ数える
      const n = src.split('availSyncPicked(freeAll)').length - 1;
      expect(n).toBe(1);
    });

    // ★2026-08-31 Phase 3: 現場の必要資格で候補を絞るようになった。
    //   絞った後の free を渡すと、「資格が分からない」欄から選んだ人が
    //   次の描画で黙って消える。**絞る前の freeAll を渡すこと。**
    it(f + ': availSyncPicked には絞る前の一覧を渡す（選んだ人が黙って消えない）', () => {
      expect(src).toContain('availSyncPicked(freeAll)');
      expect(src, '絞った後を渡している').not.toContain('availSyncPicked(free)');
      expect(src).toContain('const freeAll = roster.filter(name => !state[name]);');
    });

    it(f + ': goToCalendarDate は renderList を自分で呼ばない（switchTabが呼ぶ）', () => {
      const i = src.indexOf('function goToCalendarDate');
      // ★コメントに「renderList()」と書いてあるだけで落ちないよう、
      //   コメント行を落としてから中身を見る。
      // ★2026-09-05: 行を分けるのは /\r?\n/ で。'\n' だけで分けると、
      //   Windowsのチェックアウト（改行がCRLF）では行末に \r が残る。
      //   /\/\/.*$/ の「.」は \r に当たらないのでコメントが1文字も消えず、
      //   コメントの中の「renderList()」を拾ってこのテストが必ず落ちていた。
      const body = src.slice(i, src.indexOf('\n}', i))
        .split(/\r?\n/).map(function (L) { return L.replace(/\/\/.*$/, ''); }).join('\n');
      expect(body).toContain("switchTab('list')");
      expect(body).not.toContain('renderList()');
    });

    it(f + ': 現場管理の日付は data-goday をエスケープして埋める', () => {
      expect(src).toContain('data-goday="${esc(g.date)}"');
      expect(src).not.toContain('data-goday="${g.date}"');
    });
  });
});
