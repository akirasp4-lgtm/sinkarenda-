// 資格の絞り込みプルダウンを「実際に動かして」確かめる（2026-08-28）。
//
// ★Codexレビュー[P1]#3:
//   選んでいた資格が一覧から消えたとき、innerHTML を作り直すと選択が '' に戻り、
//   一覧が黙って「全員表示」に化けていた。資格で選んだつもりの人選が崩れるので、
//   選択は残したまま「この日に使える人はいません」と見せる。
//
// ★文字列を眺めるだけのテストでは、この不具合は捕まらない（実際に innerHTML を
//   入れ替えて value を読まないと分からない）。そこで偽のDOMを作って動かす。
import { describe, it, expect, beforeEach } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import vm from 'node:vm';

const here = dirname(fileURLToPath(import.meta.url));
const read = (f) => readFileSync(join(here, '..', '..', f), 'utf8');

const B3 = '// ===== PHASE3-QUAL-RULE:BEGIN =====';
const E3 = '// ===== PHASE3-QUAL-RULE:END =====';

// 括弧を数えて関数本体を取り出す（ui-wiring.test.js と同じやり方）
function pick(src, name) {
  const start = src.indexOf('function ' + name + '(');
  if (start < 0) return null;
  let depth = 0;
  for (let i = src.indexOf('{', start); i < src.length; i++) {
    if (src[i] === '{') depth++;
    else if (src[i] === '}') { depth--; if (depth === 0) return src.slice(start, i + 1); }
  }
  return null;
}

// 選択肢を持つだけの、とても小さな <select> の代わり
function makeSelect() {
  return {
    value: '',
    dataset: {},
    _html: '',
    get innerHTML() { return this._html; },
    set innerHTML(v) {
      this._html = v;
      // value= の並びを選択肢とみなす。今の値が無くなったら '' に戻る（本物と同じ挙動）
      const opts = [];
      const re = /<option value="([^"]*)"/g;
      let m;
      while ((m = re.exec(v)) !== null) opts.push(m[1]);
      this._options = opts;
      if (opts.indexOf(this.value) < 0) this.value = '';
    },
    _options: []
  };
}

function load(file, { quals, roster, day }) {
  const src = read(file);
  const rule = src.slice(src.indexOf(B3) + B3.length, src.indexOf(E3));
  const fn = pick(src, 'updateQualSelect');
  if (!fn) throw new Error(file + ' に updateQualSelect が無い');
  const sel = makeSelect();
  const sandbox = vm.createContext({
    console,
    document: { getElementById: (id) => (id === 'avail-qual' ? sel : null) },
    searchFilterValue: (id) => (id === 'avail-day' ? day : ''),
    todayYmd: () => '2026-08-28',
    activeRosterMembers: () => roster,
    allQuals: quals,
    esc: (s) => String(s == null ? '' : s)
  });
  sandbox.globalThis = sandbox;
  vm.runInContext(rule + '\n' + fn, sandbox, { filename: file });
  return { sel, run: () => sandbox.updateQualSelect() };
}

const GLO = 'グローライズ';
const q = (name, qual, expires) => ({ name, qual, expires: expires || '', kind: '技能講習', company: GLO });
const ROSTER = [{ name: 'A', company: GLO }, { name: 'B', company: GLO }];

describe.each(['index.html', 'admin.html'])('資格プルダウン（%s）', (file) => {
  it('資格の一覧が選択肢になる', () => {
    const { sel, run } = load(file, {
      quals: [q('A', '玉掛け'), q('B', '玉掛け'), q('A', '高所作業車')],
      roster: ROSTER, day: '2026-08-28'
    });
    run();
    expect(sel._options).toEqual(['', '玉掛け', '高所作業車']);
    expect(sel.value).toBe('');
  });

  it('★[P1] 選んでいた資格がその日に使えなくなっても、選択が黙って外れない', () => {
    // 1回目: 「高所作業車」が選べる日
    const st = load(file, {
      quals: [q('A', '玉掛け'), q('A', '高所作業車', '2026-09-30')],
      roster: ROSTER, day: '2026-08-28'
    });
    st.run();
    st.sel.value = '高所作業車';
    expect(st.sel._options).toContain('高所作業車');

    // 2回目: 期限が切れた後の日を見る → 選択肢からは消えるが、選択は残す
    const st2 = load(file, {
      quals: [q('A', '玉掛け'), q('A', '高所作業車', '2026-09-30')],
      roster: ROSTER, day: '2026-10-31'
    });
    st2.sel.value = '高所作業車';
    st2.run();
    expect(st2.sel.value, '選択が勝手に外れて全員表示に化けている').toBe('高所作業車');
    expect(st2.sel._html).toContain('いません');
  });

  it('★[P3] その日に使える人が0人の資格は、普通の選択肢としては出さない', () => {
    const { sel, run } = load(file, {
      quals: [q('A', '切れてる', '2026-08-01'), q('A', '生きてる')],
      roster: ROSTER, day: '2026-08-28'
    });
    run();
    expect(sel._options).toEqual(['', '生きてる']);
  });

  it('中身が変わらなければ作り直さない（選択がちらつかない）', () => {
    const { sel, run } = load(file, {
      quals: [q('A', '玉掛け')], roster: ROSTER, day: '2026-08-28'
    });
    run();
    const first = sel._html;
    sel._html = '__触った印__';
    run();
    expect(sel._html, '同じ中身なのに作り直している').toBe('__触った印__');
    expect(first).toContain('玉掛け');
  });

  it('資格が1件も無くても落ちない', () => {
    const { sel, run } = load(file, { quals: [], roster: ROSTER, day: '2026-08-28' });
    run();
    expect(sel._options).toEqual(['']);
  });
});
