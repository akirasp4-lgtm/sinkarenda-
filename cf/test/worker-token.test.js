// Cloudflare が GAS を呼ぶとき、パスワードを必ず付けること（2026-08-31）
//
// ★なぜこのテストが要るか（実際に開いていた穴）:
//   gas.js:270-281 の calAuthOk_ は、GASの設定 CAL_REQUIRE_TOKEN が '1' のとき
//   本文/クエリの k を照合する。ところが Cloudflare は k を1文字も送っていなかった。
//   ＝ **設定を入れた瞬間に、5分ごとの取り込みが全滅する**状態だった。
//
//   この穴は「設定を入れるまで誰も気づかない」種類のもので、
//   入れた瞬間に本番が止まる。だから機械で縛る。
//
// ★守ること:
//   ① 秘密があれば必ず付く（日報の取り込み・社長予定の取り込みの両方）
//   ② 秘密が無ければ付けない（設定前に壊さない）
//   ③ ソースからパスワードを付ける処理が消えたら赤くなる
import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { gasKeyParam } from '../src/sync.js';

const here = dirname(fileURLToPath(import.meta.url));
const src = (f) => readFileSync(join(here, '..', 'src', f), 'utf8');

describe('日報の取り込み（sync.js）', () => {
  it('秘密があればURLに付く', () => {
    expect(gasKeyParam({ CAL_TOKEN: 'abc123' })).toBe('&k=abc123');
  });

  it('★秘密が無ければ付けない（設定前に壊さない）', () => {
    expect(gasKeyParam({})).toBe('');
    expect(gasKeyParam({ CAL_TOKEN: '' })).toBe('');
    expect(gasKeyParam({ CAL_TOKEN: '   ' })).toBe('');
    expect(gasKeyParam(null)).toBe('');
    expect(gasKeyParam(undefined)).toBe('');
  });

  it('記号を含む秘密でもURLが壊れない', () => {
    expect(gasKeyParam({ CAL_TOKEN: 'a b&c=d' })).toBe('&k=a%20b%26c%3Dd');
  });

  it('前後の空白は落とす', () => {
    expect(gasKeyParam({ CAL_TOKEN: '  abc  ' })).toBe('&k=abc');
  });

  it('★GASを呼ぶURLに、この処理が入っている', () => {
    // ここが消えると、設定を入れた瞬間に取り込みが全滅する
    const s = src('sync.js');
    expect(s).toContain('gasKeyParam(env)');
    // 呼ぶURLを組み立てている行に付いていること
    const line = s.split('\n').find((l) => l.includes("'?compact=1&company=&t='"));
    expect(line, 'GASを呼ぶ行が見つからない').toBeTruthy();
    const idx = s.indexOf(line);
    const after = s.slice(idx, idx + 300);
    expect(after, '★GASを呼ぶURLにパスワードが付いていない').toContain('gasKeyParam(env)');
  });
});

describe('社長予定の取り込み（pres-sync.js）', () => {
  const s = src('pres-sync.js');

  it('★本文にパスワードを入れる処理が入っている', () => {
    expect(s).toContain('if (k) body.k = k;');
  });

  it('★秘密を受け取って渡している', () => {
    expect(s).toContain('fetchPresList(gasUrl, pin, cacheBuster, calToken)');
    expect(s).toContain('env.CAL_TOKEN');
  });

  it('★秘密が無ければ入れない（設定前に壊さない）', () => {
    // k が空文字なら body.k を作らない、という形になっていること
    expect(s).toContain("const k = String(calToken || '').trim();");
  });

  it('本文は text/plain のまま（GASのdoPostがそのまま読むため）', () => {
    expect(s).toContain("'Content-Type': 'text/plain'");
  });
});

describe('★取り込みの経路を数える（見落とし防止）', () => {
  it('GASを呼んでいるのは、この2つのファイルだけ', () => {
    // 3つ目が増えたら、そこにもパスワードを付ける必要がある
    const files = ['sync.js', 'pres-sync.js', 'index.js', 'read.js',
      'pres-read.js', 'alerts.js', 'suggest.js'];
    const callers = files.filter((f) => {
      const t = src(f);
      return t.includes('env.GAS_URL') || t.includes('GAS_URL +');
    });
    expect(callers.sort(), '★GASを呼ぶ場所が増えている。パスワードを付けたか確認すること')
      .toEqual(['pres-sync.js', 'sync.js']);
  });
});
