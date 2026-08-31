// 窓口（doGet）が日当（単価）を外へ出さないことを見張る。
//
// ★なぜ必要か（2026-08-31・実機で確認した事実）:
//   このアプリの画面は GitHub Pages で公開されており、
//   ソースを表示すると台帳の窓口URLが1行そのまま見える。
//   そのURLをブラウザのアドレス欄に貼るだけで、
//   合言葉も暗証番号も無しに 職人マスタ62人ぶんが返り、
//   **うち45人に日当（単価）が入っていた**。
//
//   画面側では2026-06-11に「現場画面のCSVから職人マスタを外す」で塞いであったが、
//   窓口の裏側は塞ぎ忘れていた。免許番号を出さないのと同じ扱いにする。
//   （gas.js の projectQualifications_ に「★免許番号は引き続き出さない（個人情報）」とある）
//
// ★この対策の限界（正直に書いておく）:
//   単価が取れる道はもう1本ある（POST の get_sheet で職人マスタを丸ごと）。
//   そちらは管理画面の単価設定が使っているので消せない。
//   ここで塞げるのは「URLを貼るだけで見える」という一番うっかりしやすい道だけ。
//   完全に塞ぐにはログインの仕組みが要る（Phase 5 で権限とまとめて検討）。
import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const here = dirname(fileURLToPath(import.meta.url));
const CODE = readFileSync(join(here, '..', '..', 'gas.js'), 'utf8');

// doGet の中の「職人マスタを外向きの形に組み立てている所」だけを切り出す。
// ★同じ文字列が別の関数（職人管理の保存側）にもあるので、必ず doGet の中から探すこと。
//   ここを間違えると、直っていないのにテストが緑になる。
function membersProjection() {
  const g = CODE.indexOf('function doGet(e)');
  if (g < 0) throw new Error('doGet が見つからない');
  const i = CODE.indexOf('const memberSheet = getOrCreateMemberSheet_(ss);', g);
  if (i < 0) throw new Error('doGet の職人マスタ組み立てが見つからない');
  const j = CODE.indexOf('const genbaSheet = getOrCreateGenbaSheet_(ss);', i);
  if (j < 0) throw new Error('組み立ての終わりが見つからない');
  return CODE.slice(i, j);
}

describe('窓口（doGet）は日当を外へ出さない', () => {
  const src = membersProjection();

  it('★日当（単価）を組み立てに含めない', () => {
    // 4列目 r[3] が単価。ここを読んでいたら漏れている。
    expect(src, 'doGet の職人マスタに rate が残っている').not.toMatch(/\brate\s*:/);
    expect(src, '単価の列(r[3])を読んでいる').not.toMatch(/r\[3\]/);
  });

  it('必要な項目は今までどおり出す（消しすぎていないこと）', () => {
    ['name', 'company', 'division', 'butai', 'active'].forEach((k) => {
      expect(src, k + ' が消えている').toMatch(new RegExp('\\b' + k + '\\s*:'));
    });
  });

  it('★単価を出さない理由がコードに書いてある（次の人が戻さないように）', () => {
    expect(src).toMatch(/単価|日当/);
  });
});

describe('画面側は窓口の日当に依存していない', () => {
  const idx = readFileSync(join(here, '..', '..', 'index.html'), 'utf8');
  const adm = readFileSync(join(here, '..', '..', 'admin.html'), 'utf8');

  it('職人用の画面は日当を1か所も読まない', () => {
    // allMembers は窓口(doGet)由来。そこから rate を読んでいたら壊れる。
    expect(idx).not.toMatch(/allMembers[^;\n]*\.rate/);
  });

  it('★管理画面の単価設定は別の道（get_sheet）から取る＝壊れない', () => {
    // settingsMembers が単価の供給元。get_sheet で職人マスタを取っている。
    expect(adm).toContain("action:'get_sheet',sheet:'職人マスタ'");
    expect(adm).toContain('settingsMembers');
  });
});
