import { readSchedule } from './read.js';
import { syncAll } from './sync.js';

const CORS = {
  'Access-Control-Allow-Origin': 'https://akirasp4-lgtm.github.io',
  'Access-Control-Allow-Methods': 'GET,POST,OPTIONS',
  // ★修正2: 画面側が /api/sync の簡易認証ヘッダ(X-Sync-Key)を送れるよう許可する。
  'Access-Control-Allow-Headers': 'Content-Type, X-Sync-Key'
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
        // （snapshotが未取り込みのとき）。どちらの形もそのままjson化して返せばよい。
        // ★companyパラメータは.trim()してから渡す。D1側は完全一致で絞り込むため、
        // URLにたまたま前後の空白が付いても一致しない事故を防ぐ
        // （gas.jsのdoGetもrequestedCompanyを.trim()してから比較している）。
        const company = (url.searchParams.get('company') || '').trim();
        return json(await readSchedule(env, company));
      } catch (e) {
        // 画面側は status!=='ok' を見てGASへ落ちる
        return json({ status: 'error', message: String(e.message || e) }, 500);
      }
    }

    // 書き込み直後に画面から呼ぶ。取り込んでから返す。
    if (url.pathname === '/api/sync' && request.method === 'POST') {
      // ★修正2（/api/syncが誰でも叩ける問題）: 共有秘密による簡易認証。
      // Worker URLもbackend.jsonも公開されているため、これは「第三者・連打による
      // 無差別な同期起動」を減らすための簡易な抑止であり、堅牢な秘匿ではない
      // （backend.jsonにも同じ鍵を書く必要があるため、鍵自体は公開情報になる）。
      // SYNC_KEYが未設定の間（移行期間中の初期状態）は認証を要求しない。
      // ただしその場合はログに残す。
      const requiredKey = env.SYNC_KEY || '';
      if (requiredKey) {
        const provided = request.headers.get('X-Sync-Key') || '';
        if (provided !== requiredKey) {
          return json({ status: 'error', message: '認証に失敗しました' }, 403);
        }
      } else {
        console.log('[sync] SYNC_KEY未設定のため、認証なしで /api/sync を許可しています（移行期間中の暫定運用）');
      }

      // ★修正7（急減ガードの自己回復）: ?force=1 が付いていれば、件数急減ガード
      // （sync.js）だけを明示的に無視して受け入れる。連続拒否を待たずに今すぐ
      // 反映したい場合の脱出口。サイズ上限や応答形式の検証はforceでも無条件のまま。
      const force = url.searchParams.get('force') === '1';

      // ★syncAllは例外を投げない契約。同時実行中は{ok:true, skipped:true}で
      // 返る（修正2の同時実行抑止）ので、try/catchではなく戻り値のokを見る。
      const r = await syncAll(env, { force });
      return json({ status: r.ok ? 'ok' : 'error', rows: r.rows, message: r.message, skipped: !!r.skipped });
    }

    if (url.pathname === '/api/health') {
      // DBが落ちているときこそhealthを見たいので、ここも失敗したら
      // 素の500ではなくJSONでエラーを返す。
      try {
        const snap = await env.DB.prepare('SELECT rows, bytes, at FROM snapshot WHERE id = 1').all();
        const last = await env.DB.prepare('SELECT * FROM sync_log ORDER BY at DESC LIMIT 1').all();
        const snapRow = snap.results && snap.results[0];
        return json({
          status: 'ok',
          rows: snapRow ? snapRow.rows : 0,
          bytes: snapRow ? snapRow.bytes : 0,
          snapshotAt: snapRow ? snapRow.at : null,
          lastSync: (last.results && last.results[0]) || null
        });
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
