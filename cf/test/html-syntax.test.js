// 画面のJavaScriptが文法として通ることを見張る。
//
// ★なぜ必要か（2026-08-27）:
//   index.html / admin.html は1ファイル4,900行超で、置き換えを機械的にやる。
//   括弧を1つ落としただけで画面が丸ごと白くなるが、正規表現の見張りでは絶対に気付けない。
//   （実際、フェーズ2で保存時の判定を差し替えたとき、これで初めて安心して進められた）
import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const here = dirname(fileURLToPath(import.meta.url));
const read = (f) => readFileSync(join(here, '..', '..', f), 'utf8');

// HTMLコメントの中に <script という文字が入っていることがあるので先に落とす
const scripts = (src) => {
  const clean = src.replace(/<!--[\s\S]*?-->/g, '');
  return [...clean.matchAll(/<script(?![^>]*\ssrc=)[^>]*>([\s\S]*?)<\/script>/g)].map(m => m[1]);
};

describe('画面のJavaScriptが文法として通ること', () => {
  ['index.html', 'admin.html'].forEach(f => {
    it(f + ': すべての <script> が構文エラーなし', () => {
      const blocks = scripts(read(f));
      expect(blocks.length).toBeGreaterThan(0);
      // new Function(code) は「構文解析するだけ」で中身は実行しない。
      // 読むのは自分たちのリポジトリのファイルだけで、外から来た文字列は混ぜない。
      blocks.forEach((code, i) => {
        expect(() => new Function(code), `${f} の script#${i + 1}`).not.toThrow();
      });
    });
  });
});
