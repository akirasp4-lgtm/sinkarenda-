// 候補者の順位付けと理由付け（依頼文の要件5「AI人員配置提案」・2026-08-29）
//
// ★依頼文（原文）
//   「5. AI人員配置提案 — 案件に必要な人数・資格・経験を入力すると、
//     予定・資格・経験・過去案件を確認して候補者を提案する。
//     ただしAIが勝手に予定確定しない。最終決定は管理者が行う。」
//
// ★候補者を出すところ（空き×資格×経験）は既に0円で出来ている
//   （index.html の PHASE5-PICK-RULE）。ここが足すのは
//   **順位付けと「なぜこの人か」の一文だけ**。
//
// ★予定は絶対に作らない。ここは文章を返すだけ。
//
// ===== 個人情報を外へ出さない ==========================================
// ★氏名をOpenAIへ送らない。画面が候補を c1,c2,... に置き換えて送り、
//   返ってきた番号を画面が名前に戻す。AIは「経験日数・資格・拠点」だけ見れば
//   順位を付けられるので、氏名は要らない。単価も現場名も送らない。
//   （元請名だけは判断材料として送る。会社名であって個人情報ではない）
//
// ===== 課金の上限 ======================================================
// ★公開URLなので、叩かれたら課金が積み上がる。3重で止める:
//   1) Origin検証（画面からの呼び出しだけ通す）
//   2) 1日の呼び出し回数の上限（D1に記録して数える）
//   3) 1回あたりの候補人数と返す文字数の上限
//   鍵が未設定なら何もせず enabled:false を返す（画面はAI欄を出さない）。

export const SUGGEST_MAX_CANDIDATES = 40;   // 1回に送る候補の上限
export const SUGGEST_MAX_TOKENS = 500;      // 1回に返させる長さの上限
export const SUGGEST_DAILY_LIMIT = 200;     // 1日の呼び出し回数の上限
export const SUGGEST_DEFAULT_MODEL = 'gpt-4o-mini';

// 送ってよい形か検査する。★氏名らしき文字が混ざっていたら弾く（事故防止）。
export function sanitizeCandidates(list) {
  const out = [];
  (list || []).forEach((c, i) => {
    if (!c || out.length >= SUGGEST_MAX_CANDIDATES) return;
    const id = String(c.id || '').trim();
    // idは c1, c2 … の形だけ許す。氏名が入っていたらここで落ちる
    if (!/^c[0-9]{1,3}$/.test(id)) return;
    out.push({
      id,
      days: Math.max(0, Math.min(9999, Number(c.days) || 0)),
      quals: (Array.isArray(c.quals) ? c.quals : []).slice(0, 8)
        .map(q => String(q || '').slice(0, 40)),
      kyoten: ['本社', '関東支店', '両方'].indexOf(String(c.kyoten || '')) >= 0
        ? String(c.kyoten) : ''
    });
  });
  return out;
}

export function buildPrompt({ genba, need, candidates }) {
  const g = String(genba || '').slice(0, 60);
  const n = Math.max(1, Math.min(50, Number(need) || 1));
  const lines = candidates.map(c => {
    const q = c.quals.length ? ' 資格:' + c.quals.join('・') : ' 資格:なし';
    const k = c.kyoten ? ' 拠点:' + c.kyoten : '';
    return `${c.id} 経験:${c.days}日${q}${k}`;
  });
  return [
    'あなたは建設・電気工事会社の配置担当です。',
    `明日「${g || '（元請未指定）'}」の現場に ${n}人 必要です。`,
    '下の候補から、適した順に並べ替えてください。',
    '',
    '判断の材料はこれだけです。ここに無いことを推測しないでください。',
    ...lines,
    '',
    '規則:',
    `- 上位${n}人を選び、それぞれ理由を日本語30字以内で1文。`,
    '- 理由は経験日数・資格・拠点のどれに基づくかを必ず書く。',
    '- 候補に無いidを出さない。同じidを2回出さない。',
    '- 決めるのは人間です。断定せず「候補」として述べる。',
    '',
    'JSONだけを返す。形式:',
    '{"picks":[{"id":"c1","reason":"…"}]}'
  ].join('\n');
}

// AIの返事を検査して、こちらの候補にあるidだけ残す。
export function parsePicks(text, candidates) {
  const known = new Set(candidates.map(c => c.id));
  let obj = null;
  try {
    const m = String(text || '').match(/\{[\s\S]*\}/);
    obj = m ? JSON.parse(m[0]) : null;
  } catch (_e) { obj = null; }
  if (!obj || !Array.isArray(obj.picks)) return [];
  const seen = new Set();
  const out = [];
  obj.picks.forEach(p => {
    const id = String((p && p.id) || '').trim();
    if (!known.has(id) || seen.has(id)) return;   // 知らないid・重複は捨てる
    seen.add(id);
    out.push({ id, reason: String((p && p.reason) || '').slice(0, 60) });
  });
  return out;
}

// 1日の呼び出し回数。D1に無ければ数えない（フェイルオープンにしない＝
// 数えられないときは呼ばない。課金は止める側に倒す）。
export async function overDailyLimit(env) {
  try {
    const today = new Date().toISOString().slice(0, 10);
    const res = await env.DB.prepare(
      'SELECT COUNT(*) AS c FROM ai_log WHERE at LIKE ?').bind(today + '%').all();
    const count = (res.results && res.results[0] && Number(res.results[0].c)) || 0;
    return count >= SUGGEST_DAILY_LIMIT;
  } catch (_e) {
    return true;   // ★数えられないなら呼ばない（課金が青天井になるのを防ぐ）
  }
}

export async function logCall(env, ok) {
  try {
    await env.DB.prepare('INSERT INTO ai_log(at, ok) VALUES(?,?)')
      .bind(new Date().toISOString(), ok ? 1 : 0).run();
  } catch (_e) { /* 記録できなくても本筋は止めない */ }
}

// OpenAIを呼ぶ。fetchを差し替えられるようにしてテストから実際に動かす。
export async function callOpenAI(env, prompt, fetchImpl) {
  const f = fetchImpl || fetch;
  const res = await f('https://api.openai.com/v1/chat/completions', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Authorization: 'Bearer ' + env.OPENAI_API_KEY
    },
    body: JSON.stringify({
      model: env.OPENAI_MODEL || SUGGEST_DEFAULT_MODEL,
      temperature: 0,
      max_tokens: SUGGEST_MAX_TOKENS,
      response_format: { type: 'json_object' },
      messages: [{ role: 'user', content: prompt }]
    })
  });
  if (!res.ok) throw new Error('OpenAI ' + res.status);
  const j = await res.json();
  return (j && j.choices && j.choices[0] && j.choices[0].message
    && j.choices[0].message.content) || '';
}
