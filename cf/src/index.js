import { readSchedule } from './read.js';
import { syncAll } from './sync.js';

const CORS = {
  'Access-Control-Allow-Origin': 'https://akirasp4-lgtm.github.io',
  'Access-Control-Allow-Methods': 'GET,POST,OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type'
};

const json = (obj, status = 200) =>
  new Response(JSON.stringify(obj), {
    status, headers: { 'Content-Type': 'application/json; charset=utf-8', ...CORS }
  });

export default {
  async fetch(request, env, ctx) {
    const url = new URL(request.url);
    if (request.method === 'OPTIONS') return new Response(null, { headers: CORS });

    if (url.pathname === '/api/schedule') {
      try {
        // readScheduleは失敗を投げず {status:'error', message} を返すこともある
        // （sync_logが未取り込み/失敗のとき。計画からの変更1）。
        // どちらの形もそのままjson化して返せばよい。
        // ★companyパラメータは.trim()してから渡す。D1側は完全一致（WHERE kaisha = ?）
        // で絞り込むため、URLにたまたま前後の空白が付いても一致しない事故を防ぐ
        // （gas.jsのdoGetもrequestedCompanyを.trim()してから比較している。レビュー指摘）。
        const company = (url.searchParams.get('company') || '').trim();
        return json(await readSchedule(env, company));
      } catch (e) {
        // 画面側は status!=='ok' を見てGASへ落ちる
        return json({ status: 'error', message: String(e.message || e) }, 500);
      }
    }

    // 書き込み直後に画面から呼ぶ。取り込んでから返す。
    // ★計画からの変更2：syncAllは例外を投げない契約（Task 2で確定）なので、
    // try/catchではなく戻り値のokをそのまま見る。
    if (url.pathname === '/api/sync' && request.method === 'POST') {
      const r = await syncAll(env);
      return json({ status: r.ok ? 'ok' : 'error', rows: r.rows, message: r.message });
    }

    if (url.pathname === '/api/health') {
      // ★計画からの変更3：DBが落ちているときこそhealthを見たいので、
      // ここも失敗したら素の500ではなくJSONでエラーを返す。
      try {
        const last = await env.DB.prepare('SELECT * FROM sync_log ORDER BY at DESC LIMIT 1').all();
        const cnt = await env.DB.prepare('SELECT COUNT(*) AS c FROM nippo').all();
        return json({ status: 'ok', rows: cnt.results[0].c, lastSync: last.results[0] || null });
      } catch (e) {
        return json({ status: 'error', message: String(e.message || e) }, 500);
      }
    }

    return json({ status: 'error', message: 'not found' }, 404);
  },

  async scheduled(event, env, ctx) {
    ctx.waitUntil(syncAll(env));
  }
};
