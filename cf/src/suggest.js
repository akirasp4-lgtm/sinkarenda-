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
export const SUGGEST_TIMEOUT_MS = 15000;    // OpenAIの応答を待つ上限

// 送ってよい形か検査する。★氏名らしき文字が混ざっていたら弾く（事故防止）。
export function sanitizeCandidates(list) {
  const out = [];
  // ★Codexレビュー[P2]: 40件たまった後も末尾まで走査していた。break で止める
  for (const c of (list || [])) {
    if (out.length >= SUGGEST_MAX_CANDIDATES) break;
    if (!c) continue;
    const id = String(c.id || '').trim();
    // idは c1, c2 … の形だけ許す。氏名が入っていたらここで落ちる
    if (!/^c[0-9]{1,3}$/.test(id)) continue;
    out.push({
      id,
      days: Math.max(0, Math.min(9999, Number(c.days) || 0)),
      quals: (Array.isArray(c.quals) ? c.quals : []).slice(0, 8)
        .map(q => String(q || '').slice(0, 40)),
      kyoten: ['本社', '関東支店', '両方'].indexOf(String(c.kyoten || '')) >= 0
        ? String(c.kyoten) : ''
    });
  }
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
export function parsePicks(text, candidates, need) {
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
  // ★Codexレビュー[P2]: 頼んだ人数より多く返ってきたら切る
  const n = Math.max(1, Math.min(candidates.length, Number(need) || candidates.length));
  return out.slice(0, n);
}

// 1日の呼び出し回数。D1に無ければ数えない（フェイルオープンにしない＝
// 数えられないときは呼ばない。課金は止める側に倒す）。
// ★Codexレビュー[P1]（2026-08-29）: 「数える→OpenAIを呼ぶ→記録する」の順だと、
//   同時に大量に叩かれたとき全部が上限判定を通過してから呼びに行ける（すり抜け）。
//   **先に1件記録してから数える**（席を取ってから数える）。取った席が上限を超えて
//   いたら呼ばずに返す。多少多めに記録されても、課金する側が止まるので安全側。
// ★コードレビュー（2026-08-30）: 予約で1行、成功でもう1行 INSERT していたため、
//   **1日200回の上限が実質100回**になっていた（成功1回につき2行数えていた）。
//   さらに ok 列は予約行が常に0のままで、成功率の判定にも使えなかった。
//   → 予約した行のidを返し、結果はその行を **UPDATE** する。1呼び出し＝1行。
export async function reserveCall(env) {
  try {
    const ins = await env.DB.prepare('INSERT INTO ai_log(at, ok) VALUES(?,?)')
      .bind(new Date().toISOString(), 0).run();
    const id = (ins && ins.meta && ins.meta.last_row_id) || null;
    const today = new Date().toISOString().slice(0, 10);
    const res = await env.DB.prepare(
      'SELECT COUNT(*) AS c FROM ai_log WHERE at LIKE ?').bind(today + '%').all();
    const count = (res.results && res.results[0] && Number(res.results[0].c)) || 0;
    return { ok: count <= SUGGEST_DAILY_LIMIT, id };   // 席が取れたか＋予約した行
  } catch (_e) {
    // ★記録も数えもできないなら呼ばない（課金が青天井になるのを防ぐ）
    return { ok: false, id: null };
  }
}

// 予約した行に結果を書く。★新しい行を足さない（足すと上限が半分になる）。
export async function logCall(env, ok, id) {
  try {
    if (id == null) return;              // 予約できていなければ何もしない
    await env.DB.prepare('UPDATE ai_log SET ok = ? WHERE id = ?')
      .bind(ok ? 1 : 0, id).run();
  } catch (_e) { /* 記録できなくても本筋は止めない */ }
}

// 古い記録を消す。★Cronから呼ぶ。放っておくと無制限に増え、
//   reserveCall が毎回走らせる COUNT が行数に比例して重くなる。
export async function cleanupAiLog(env, keepDays = 30) {
  try {
    const cutoff = new Date(Date.now() - keepDays * 86400000).toISOString();
    await env.DB.prepare('DELETE FROM ai_log WHERE at < ?').bind(cutoff).run();
  } catch (_e) { /* 掃除に失敗しても本筋は止めない */ }
}

// OpenAIを呼ぶ。fetchを差し替えられるようにしてテストから実際に動かす。
export async function callOpenAI(env, prompt, fetchImpl) {
  const f = fetchImpl || fetch;
  // ★Codexレビュー[P2]: タイムアウトが無いと、応答が返らないとき窓口が詰まる
  const ac = typeof AbortController === 'function' ? new AbortController() : null;
  const timer = ac ? setTimeout(() => ac.abort(), SUGGEST_TIMEOUT_MS) : null;
  try {
  const res = await f('https://api.openai.com/v1/chat/completions', {
    signal: ac ? ac.signal : undefined,
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
  const ch = j && j.choices && j.choices[0];
  // ★Codexレビュー[P2]: 長さ切れで途中終了した返事をそのまま使わない
  if (ch && ch.finish_reason && ch.finish_reason !== 'stop') {
    throw new Error('OpenAI 応答が途中で終わりました (' + ch.finish_reason + ')');
  }
  return (ch && ch.message && ch.message.content) || '';
  } finally {
    if (timer) clearTimeout(timer);
  }
}
