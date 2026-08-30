// 元請まわりの画面を「実際に描かせて」確かめる（2026-08-31）
//
// ★なぜ必要か:
//   このプロジェクトは引用符の入った文字列をHTMLへ埋めて**4回**壊している。
//   esc() は textContent→innerHTML の仕組みなので & < > しか逃がさない
//   （引用符はそのまま通る）。value="${esc(名前)}" に二重引用符入りの元請名が
//   来ると属性がそこで閉じて、プルダウンから下が丸ごと壊れる。
//   正規表現で見張っても気付けないので、偽のDOMの上で実際に描いて中身を見る。
import { describe, it, expect, beforeAll } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import vm from 'node:vm';

const here = dirname(fileURLToPath(import.meta.url));
const read = (f) => readFileSync(join(here, '..', '..', f), 'utf8');

const cut = (src, begin, end, file) => {
  const i = src.indexOf(begin), j = src.indexOf(end, i);
  if (i < 0 || j < 0) throw new Error(file + ' に ' + begin + ' が無い');
  return src.slice(i, j + end.length);
};

// ブラウザの textContent→innerHTML と同じ逃がし方をする偽のDOM。
// ★& < > だけ。引用符はわざと逃がさない（本物と同じにしないと試験にならない）。
function fakeDoc(store) {
  return {
    createElement: () => ({
      _t: '',
      set textContent(v) { this._t = String(v); },
      get innerHTML() {
        return this._t.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
      }
    }),
    getElementById: (id) => store[id] || null,
    querySelector: () => null,
    addEventListener: () => {}
  };
}

function stage(file) {
  const src = read(file);
  const rule = cut(src, '// ===== PHASE6-GENBA-RULE:BEGIN =====', '// ===== PHASE6-GENBA-RULE:END =====', file);
  const escs = cut(src, 'function esc(', "function escAttr(str){return esc(str).replace(/\"/g,'&quot;').replace(/'/g,'&#39;');}", file);
  const render = cut(src, 'function renderGenbaFix(){', '\n}', file);
  const populate = cut(src, 'function populateGenbaSelect(p){', '\n}', file);

  const store = {
    'gfix-body': { innerHTML: '' },
    's-genba-select': { value: '', innerHTML: '', options: [] },
    's-genba-search': { value: '' }
  };
  // innerHTML を入れたら options を作り直す最小限の select もどき
  Object.defineProperty(store['s-genba-select'], 'innerHTML', {
    get() { return this._h || ''; },
    set(h) {
      this._h = h;
      this.options = [...String(h).matchAll(/<option value="([^"]*)"/g)].map((m) => ({ value: m[1] }));
    }
  });

  const box = {
    console, String, Object, Array, Number, Boolean, Math, JSON,
    document: fakeDoc(store),
    esc: null,
    allNippos: [],
    currentCompany: 'グローライズ',
    getGenbaMasterNames: () => box._master || [],
    gfixGroups: [], gfixUnreg: [], gfixBusy: false,
    _master: [], _store: store
  };
  const ctx = vm.createContext(box);
  box.globalThis = box;
  // ★必ず改行で繋ぐこと。ルールブロックの最後の行は `// ===== …:END =====` という
  //   行コメントなので、`;` で繋ぐと次のコードが丸ごとコメントに飲まれる。
  //   （実際それで「Illegal return statement」が出た）
  vm.runInContext([
    escs,
    rule,
    'let gfixGroups=[],gfixUnreg=[],gfixBusy=false;',
    render,
    populate,
    'globalThis.__render=renderGenbaFix;globalThis.__pop=populateGenbaSelect;',
    'globalThis.__esc=esc;globalThis.__escAttr=escAttr;',
    'globalThis.__peek=()=>({groups:gfixGroups,unreg:gfixUnreg});'
  ].join('\n;\n'), ctx, { filename: file });
  return box;
}

const n = (genba, o) => Object.assign({
  genba, date: '2026-08-01', company: 'グローライズ', isGhost: false
}, o || {});

['index.html', 'admin.html'].forEach((file) => {
  describe(file + ' — 実際に描いてみる', () => {
    let S;
    beforeAll(() => { S = stage(file); });

    it('偽のDOMが本物と同じ振る舞いをしている（試験そのものの確認）', () => {
      // esc は引用符を逃がさない。だから escAttr が要る、という前提を固定する。
      expect(S.__esc('a"b')).toBe('a"b');
      expect(S.__esc('a<b>c')).toBe('a&lt;b&gt;c');
    });

    it('★escAttr は引用符を逃がす（属性が途中で閉じない）', () => {
      expect(S.__escAttr('a"b')).toBe('a&quot;b');
      expect(S.__escAttr("O'Brien")).toBe('O&#39;Brien');
      expect(S.__escAttr('<b>&')).toBe('&lt;b&gt;&amp;');
    });

    it('★二重引用符入りの元請名でもプルダウンが壊れない', () => {
      S._master = [];
      S.allNippos = [n('A" onmouseover="alert(1)'), n('ふつうの元請')];
      S.__pop('s');
      const html = S._store['s-genba-select'].innerHTML;
      // 属性が途中で閉じていないこと＝生の二重引用符が value の中に無いこと
      expect(html).not.toContain('value="A" onmouseover');
      expect(html).toContain('&quot;');
      // option の数が想定どおり（壊れていれば数が狂う）
      expect(S._store['s-genba-select'].options.length).toBe(4);  // 空 + 2件 + 直接入力
    });

    it('★マスタ登録済みの元請でも引用符を逃がす（登録済みと直接入力の両方を守る）', () => {
      // ★わざと守りを外して赤くなるか試したとき、直接入力ぶんしか見ていなくて
      //   マスタ側の抜けに気付けなかった。両方を別々に見張る。
      S._master = ['B" onmouseover="alert(1)'];
      S.allNippos = [];
      S.__pop('s');
      const html = S._store['s-genba-select'].innerHTML;
      expect(html).not.toContain('value="B" onmouseover');
      expect(html).toContain('&quot;');
      expect(S._store['s-genba-select'].options.length).toBe(3);  // 空 + 1件 + 直接入力
    });

    it('★山括弧入りの元請名がタグとして解釈されない', () => {
      S._master = [];
      S.allNippos = [n('<script>bad</script>')];
      S.__pop('s');
      const html = S._store['s-genba-select'].innerHTML;
      expect(html).not.toContain('<script>');
      expect(html).toContain('&lt;script&gt;');
    });

    it('表記ゆれカードにも生のタグが出ない', () => {
      S._master = ['<img src=x>会社'];
      S.allNippos = [n('<img src=x>会社'), n('<img src=x>会社支店')];
      S.__render();
      const html = S._store['gfix-body'].innerHTML;
      expect(html).not.toContain('<img src=x>');
      expect(html).toContain('&lt;img');
    });

    it('★カードのボタンには番号だけを入れる（名前は入れない）', () => {
      S._master = ['グローライズ自社'];
      S.allNippos = [n('グローライズ自社'), n('グローライズ')];
      S.__render();
      const html = S._store['gfix-body'].innerHTML;
      expect(html).toMatch(/data-gfixmerge="\d+"/);
      expect(html).toMatch(/data-gfixidx="\d+"/);
      // data-* に元請名そのものが入っていないこと
      expect(html).not.toMatch(/data-gfixmerge="[^"]*グローライズ/);
    });

    it('★描いた中身と配列の並びが一致する（ボタンの番号がズレない）', () => {
      S._master = ['グローライズ自社'];
      S.allNippos = [n('グローライズ自社'), n('グローライズ'), n('HSJ'), n('株式会社HSJ')];
      S.__render();
      const { groups } = S.__peek();
      const html = S._store['gfix-body'].innerHTML;
      groups.forEach((g, gi) => {
        g.forEach((x, xi) => {
          expect(html).toContain('data-gfixmerge="' + gi + '" data-gfixidx="' + xi + '"');
        });
      });
    });

    it('似た元請が無ければ「見つかりませんでした」と出す', () => {
      S._master = ['きんでん東'];
      S.allNippos = [n('きんでん東')];
      S.__render();
      expect(S._store['gfix-body'].innerHTML).toContain('見つかりませんでした');
    });

    it('データが空でも落ちない', () => {
      S._master = [];
      S.allNippos = [];
      expect(() => S.__render()).not.toThrow();
      expect(() => S.__pop('s')).not.toThrow();
    });

    it('★絞り込み中でも「直接入力する」は必ず残る（打ち直す道を塞がない）', () => {
      S._master = ['きんでん東'];
      S.allNippos = [n('きんでん東')];
      S._store['s-genba-search'].value = 'zzz該当なし';
      S.__pop('s');
      expect(S._store['s-genba-select'].innerHTML).toContain('__manual__');
      S._store['s-genba-search'].value = '';
    });

    it('★「直接入力する」を選んだまま絞り込んでも選択が外れない', () => {
      S._master = ['きんでん東'];
      S.allNippos = [n('きんでん東')];
      S._store['s-genba-select'].value = '__manual__';
      S._store['s-genba-search'].value = 'きん';
      S.__pop('s');
      expect(S._store['s-genba-select'].value).toBe('__manual__');
      S._store['s-genba-search'].value = '';
    });
  });
});
