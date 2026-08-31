// 引っ越しの間、新旧どちらのURLからも使えること（2026-08-31）
//
// 利用者の心配:
//   「今作ってる間に一時的に動けへんようになるとかっていうことが困る」
//
// 止まる可能性があるのは1か所だけだった。受け付けるURLを**1つしか**書けない作りで、
// 新URLに書き換えた瞬間に古いURLが死ぬ。両方を同時に受け付ければ、
// 「切り替えの瞬間」そのものが無くなる。
//
// ★このファイルが守ること:
//   ① 今のURLは、引っ越しの前も後も必ず通る（社員が困らない）
//   ② 一覧に足したURLも通る
//   ③ 一覧に無いURLは通らない（守りを緩めていない）
//   ④ 返事のヘッダが、送ってきた側に合わせて変わる（片方が弾かれない）
import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { matchOrigin, corsFor } from '../src/index.js';

const here = dirname(fileURLToPath(import.meta.url));
const SRC = readFileSync(join(here, '..', 'src', 'index.js'), 'utf8');

const NOW = 'https://akirasp4-lgtm.github.io';
const req = (origin) => ({ headers: { get: (k) => (k === 'Origin' ? origin : null) } });

describe('★今のURLは絶対に通る', () => {
  it('今のURLが一覧に入っている', () => {
    expect(matchOrigin(NOW)).toBe(NOW);
  });

  it('★ソースにも今のURLが残っている（引っ越しで消してしまわない）', () => {
    // 全員が新しいURLへ移り終わるまで、この行を消してはいけない
    expect(SRC).toContain("'https://akirasp4-lgtm.github.io'");
  });

  it('今のURLからのCORSの返事が、今のURL宛になっている', () => {
    expect(corsFor(req(NOW))['Access-Control-Allow-Origin']).toBe(NOW);
  });
});

describe('一覧に無いURLは通さない（守りを緩めていない）', () => {
  it('よその、まったく別のURLは通らない', () => {
    expect(matchOrigin('https://evil.example.com')).toBe('');
  });

  it('★前方一致で騙されない（偽物のドメイン）', () => {
    // 'https://akirasp4-lgtm.github.io.evil.com' のような偽物は通してはいけない
    expect(matchOrigin(NOW + '.evil.com')).toBe('');
    expect(matchOrigin(NOW + '/../evil')).toBe('');
  });

  it('http:// と https:// を取り違えない', () => {
    expect(matchOrigin('http://akirasp4-lgtm.github.io')).toBe('');
  });

  it('Originが無い・空でも落ちない', () => {
    expect(matchOrigin(null)).toBe('');
    expect(matchOrigin(undefined)).toBe('');
    expect(matchOrigin('')).toBe('');
    expect(corsFor(null)['Access-Control-Allow-Origin']).toBe(NOW);
    expect(corsFor({})['Access-Control-Allow-Origin']).toBe(NOW);
  });

  it('一覧に無いURLには、こちらの既定を返す（相手のブラウザが弾く）', () => {
    expect(corsFor(req('https://evil.example.com'))['Access-Control-Allow-Origin']).toBe(NOW);
  });
});

describe('★引っ越しの仕組みが用意できている', () => {
  it('受け付け先が「一覧」になっている（1つ固定ではない）', () => {
    expect(SRC).toContain('const ALLOWED_ORIGINS = [');
  });

  it('返事のヘッダをリクエストごとに作っている', () => {
    expect(SRC).toContain('const cors = corsFor(request);');
  });

  it('★途中の機械にキャッシュさせない（Vary: Origin）', () => {
    // これが無いと、片方のURL宛の返事がもう片方に配られて弾かれることがある
    expect(corsFor(req(NOW)).Vary).toBe('Origin');
  });

  it('★返事の入れ物をリクエストごとに作っている（同時の依頼が混ざらない）', () => {
    // Workerは同じ入れ物で複数の依頼を同時に扱う。
    // 外側の変数に入れ替える書き方だと、別の依頼の返事に混ざる。
    expect(SRC).toContain('const json = (obj, status = 200) => new Response');
    expect(SRC, '外側の変数へ入れ替えている').not.toMatch(/^\s*cors\s*=\s*corsFor/m);
  });

  // ★引っ越し当日、ここに新しいURLを足したら、この2件のコメントアウトを外す。
  //   足す前は「まだ足していない」ことを確認しておく。
  it('（引っ越し前）まだ新しいURLは足していない', () => {
    const list = SRC.slice(SRC.indexOf('const ALLOWED_ORIGINS = ['));
    const block = list.slice(0, list.indexOf(']'));
    const urls = block.match(/'https?:\/\/[^']+'/g) || [];
    expect(urls).toHaveLength(1);
  });
});

describe('★2つ受け付ける形が正しく動く（引っ越し当日の予行演習）', () => {
  // 実際の一覧は変えずに、同じ規則を手元で作って確かめる
  const LIST = [NOW, 'https://yotei.pages.dev'];
  const match2 = (o) => (LIST.indexOf(String(o || '')) >= 0 ? String(o) : '');
  const cors2 = (o) => ({
    'Access-Control-Allow-Origin': match2(o) || LIST[0],
    Vary: 'Origin'
  });

  it('古いURLも新しいURLも、どちらも通る', () => {
    expect(match2(NOW)).toBe(NOW);
    expect(match2('https://yotei.pages.dev')).toBe('https://yotei.pages.dev');
  });

  it('★返事がそれぞれ自分宛になる（片方が弾かれない）', () => {
    expect(cors2(NOW)['Access-Control-Allow-Origin']).toBe(NOW);
    expect(cors2('https://yotei.pages.dev')['Access-Control-Allow-Origin'])
      .toBe('https://yotei.pages.dev');
  });

  it('2つ受け付けても、よそは通らない', () => {
    expect(match2('https://evil.example.com')).toBe('');
  });
});
