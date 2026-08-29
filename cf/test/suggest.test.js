// 候補者の順位付けと理由付け（要件5「AI人員配置提案」・2026-08-29）
//
// ★一番大事な検査（この順に大事）:
//   1. **氏名がOpenAIへ送られない**（個人情報を外へ出さない）
//   2. **課金が青天井にならない**（鍵なし・上限超え・数えられない時は呼ばない）
//   3. AIの返事を鵜呑みにしない（知らないid・重複・壊れたJSONを弾く）
//   4. 外へ実際の通信をしない（全部モックする）
import { describe, it, expect } from 'vitest';
import {
  sanitizeCandidates, buildPrompt, parsePicks, overDailyLimit, callOpenAI,
  SUGGEST_MAX_CANDIDATES, SUGGEST_MAX_TOKENS, SUGGEST_DAILY_LIMIT
} from '../src/suggest.js';

const C = (id, days, quals = [], kyoten = '') => ({ id, days, quals, kyoten });

describe('★氏名を外へ出さない', () => {
  it('idが c1 形式でない候補は落とす（氏名が紛れ込んだら通さない）', () => {
    const out = sanitizeCandidates([
      C('c1', 10), C('真柄', 5), C('c2', 3), C('', 1), C('c999x', 1)
    ]);
    expect(out.map(c => c.id)).toEqual(['c1', 'c2']);
  });

  it('★組み立てた文章に日本人の氏名が1文字も入らない', () => {
    const cand = sanitizeCandidates([C('c1', 84, ['玉掛け']), C('c2', 20, [], '関東支店')]);
    const p = buildPrompt({ genba: 'きんでん東', need: 2, candidates: cand });
    ['真柄', '江頭', '奥田', '高田', '河原', '中島'].forEach(n => {
      expect(p).not.toContain(n);
    });
    expect(p).toContain('c1');
    expect(p).toContain('経験:84日');
  });

  it('資格名は8個・40字までに切る（長文を外へ流さない）', () => {
    const out = sanitizeCandidates([C('c1', 1, new Array(20).fill('あ'.repeat(100)))]);
    expect(out[0].quals).toHaveLength(8);
    expect(out[0].quals[0].length).toBe(40);
  });

  it('拠点は決まった値だけ通す', () => {
    expect(sanitizeCandidates([C('c1', 1, [], '本社')])[0].kyoten).toBe('本社');
    expect(sanitizeCandidates([C('c1', 1, [], '大阪の自宅')])[0].kyoten).toBe('');
  });

  it('候補は40人までしか送らない', () => {
    const many = Array.from({ length: 100 }, (_, i) => C('c' + i, i));
    expect(sanitizeCandidates(many)).toHaveLength(SUGGEST_MAX_CANDIDATES);
  });
});

describe('★課金が青天井にならない', () => {
  const dbWith = (count) => ({
    DB: { prepare: () => ({ bind: () => ({ all: async () => ({ results: [{ c: count }] }) }) }) }
  });

  it('1日の上限に達したら呼ばない', async () => {
    expect(await overDailyLimit(dbWith(SUGGEST_DAILY_LIMIT))).toBe(true);
    expect(await overDailyLimit(dbWith(SUGGEST_DAILY_LIMIT - 1))).toBe(false);
  });

  it('★回数を数えられないときは「呼ばない」側に倒す', async () => {
    const broken = { DB: { prepare: () => { throw new Error('D1 down'); } } };
    expect(await overDailyLimit(broken)).toBe(true);
  });

  it('1回に返させる長さに上限をかけている', async () => {
    let sentBody = null;
    const fake = async (_u, o) => {
      sentBody = JSON.parse(o.body);
      return { ok: true, json: async () => ({ choices: [{ message: { content: '{}' } }] }) };
    };
    await callOpenAI({ OPENAI_API_KEY: 'k' }, 'p', fake);
    expect(sentBody.max_tokens).toBe(SUGGEST_MAX_TOKENS);
    expect(sentBody.temperature).toBe(0);
    expect(sentBody.model).toBe('gpt-4o-mini');
  });

  it('モデルは設定で差し替えられる', async () => {
    let m = null;
    const fake = async (_u, o) => {
      m = JSON.parse(o.body).model;
      return { ok: true, json: async () => ({ choices: [{ message: { content: '{}' } }] }) };
    };
    await callOpenAI({ OPENAI_API_KEY: 'k', OPENAI_MODEL: 'gpt-4.1-mini' }, 'p', fake);
    expect(m).toBe('gpt-4.1-mini');
  });

  it('OpenAIがエラーを返したら例外にする（黙って成功扱いにしない）', async () => {
    const fake = async () => ({ ok: false, status: 429 });
    await expect(callOpenAI({ OPENAI_API_KEY: 'k' }, 'p', fake)).rejects.toThrow('OpenAI 429');
  });
});

describe('★AIの返事を鵜呑みにしない', () => {
  const cand = sanitizeCandidates([C('c1', 10), C('c2', 5), C('c3', 1)]);

  it('候補に無いidは捨てる（幻の人を出さない）', () => {
    const t = '{"picks":[{"id":"c1","reason":"経験が多い"},{"id":"c9","reason":"知らない人"}]}';
    expect(parsePicks(t, cand).map(p => p.id)).toEqual(['c1']);
  });

  it('同じidが2回来たら1回にする', () => {
    const t = '{"picks":[{"id":"c1","reason":"a"},{"id":"c1","reason":"b"}]}';
    expect(parsePicks(t, cand)).toHaveLength(1);
  });

  it('壊れたJSONでも落ちない', () => {
    expect(parsePicks('これはJSONではありません', cand)).toEqual([]);
    expect(parsePicks('', cand)).toEqual([]);
    expect(parsePicks(null, cand)).toEqual([]);
  });

  it('前後に文章が付いていても中のJSONを拾う', () => {
    const t = 'はい、こちらです。\n{"picks":[{"id":"c2","reason":"資格あり"}]}\n以上です。';
    expect(parsePicks(t, cand).map(p => p.id)).toEqual(['c2']);
  });

  it('理由が長すぎたら切る', () => {
    const t = JSON.stringify({ picks: [{ id: 'c1', reason: 'あ'.repeat(500) }] });
    expect(parsePicks(t, cand)[0].reason.length).toBe(60);
  });

  it('picks が配列でなければ空を返す', () => {
    expect(parsePicks('{"picks":"c1"}', cand)).toEqual([]);
    expect(parsePicks('{}', cand)).toEqual([]);
  });
});

describe('文章の中身', () => {
  it('★「決めるのは人間」と必ず書く（依頼文: 最終決定は管理者が行う）', () => {
    const p = buildPrompt({ genba: 'X', need: 2, candidates: sanitizeCandidates([C('c1', 1)]) });
    expect(p).toContain('決めるのは人間です');
  });

  it('材料に無いことを推測しないよう明示する', () => {
    const p = buildPrompt({ genba: 'X', need: 1, candidates: sanitizeCandidates([C('c1', 1)]) });
    expect(p).toContain('ここに無いことを推測しないでください');
  });

  it('必要人数は1〜50に丸める', () => {
    const c = sanitizeCandidates([C('c1', 1)]);
    expect(buildPrompt({ need: 0, candidates: c })).toContain('1人 必要');
    expect(buildPrompt({ need: 999, candidates: c })).toContain('50人 必要');
  });

  it('元請が未指定でも組み立てられる', () => {
    const p = buildPrompt({ candidates: sanitizeCandidates([C('c1', 1)]) });
    expect(p).toContain('（元請未指定）');
  });

  it('資格が無い人は「資格:なし」と書く', () => {
    const p = buildPrompt({ candidates: sanitizeCandidates([C('c1', 3)]) });
    expect(p).toContain('c1 経験:3日 資格:なし');
  });
});
