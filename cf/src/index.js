import { readSchedule } from './read.js';
import { syncAll, cleanupSyncLog } from './sync.js';

// 画面（GitHub Pages）だけが正規の呼び出し元。/api/syncのOrigin検証・CORSの両方で使う。
const ALLOWED_ORIGIN = 'https://akirasp4-lgtm.github.io';

const CORS = {
  'Access-Control-Allow-Origin': ALLOWED_ORIGIN,
  'Access-Control-Allow-Methods': 'GET,POST,OPTIONS',
  // ★修正2: 画面側が /api/sync の簡易認証ヘッダ(X-Sync-Key)を送れるよう許可する。
  'Access-Control-Allow-Headers': 'Content-Type, X-Sync-Key'
};

// ★3回目レビュー修正5（/api/syncの認証）: SYNC_KEYは空でも設定しても、backend.jsonという
// 公開ファイルに同じ値を置く構造上「秘密」にならない（前回レビューで指摘済み・現状維持）。
// そこで、鍵に頼らない緩和策としてOriginヘッダの検証を追加する。
// ★これは「完全な認証」ではない。ブラウザはOriginヘッダを自分で偽装できないため、
// ブラウザ経由（＝この画面以外のWebページから叩く等）の悪用は防げる。しかし
// curl等のHTTPクライアントはOriginヘッダを自由に書き換えられるため、直接リクエストは
// 防げない。「連打による無料枠枯渇の緩和」であって、なりすまし防止ではないことを
// 正直に書いておく。
function isAllowedOrigin(request) {
  return (request.headers.get('Origin') || '') === ALLOWED_ORIGIN;
}

// ★修正5（レート制限）: Origin検証はcurl等の直接リクエストを防げないため、それでも
// 無料枠を守れるよう「直近1分間に実行された同期の回数」に上限を設ける。
// sync_logは「ロック待ちでスキップ」以外の同期試行(成功・スキップ・拒否・失敗)ごとに
// 必ず1行書かれるため、直近の行数は「実際にGAS/D1へ負荷をかけた回数」の実測値に
// なる（リクエスト受信数そのものではなく、実際の消費量を見ている点がポイント）。
// しきい値を超えている間は、syncAll自体を呼ばずに「進行中」と同じ扱いでスキップする
// （GASへの取得もD1への書き込みも一切発生させない）。
const SYNC_RATE_WINDOW_MS = 60_000;
const SYNC_RATE_LIMIT = 12; // 通常はCron5分に1回=1件/分。書き込み後の即時同期や複数端末の同時操作を見込んだ余裕込みの上限。

async function isSyncRateLimited(env) {
  try {
    const cutoff = new Date(Date.now() - SYNC_RATE_WINDOW_MS).toISOString();
    const res = await env.DB.prepare('SELECT COUNT(*) AS c FROM sync_log WHERE at > ?').bind(cutoff).all();
    const count = (res.results && res.results[0] && Number(res.results[0].c)) || 0;
    return count >= SYNC_RATE_LIMIT;
  } catch (_e) {
    // 判定できない場合は同期そのものを止めない（フェイルオープン。sync.jsの
    // tryAcquireLockと同じ方針＝この機構はあくまでbest-effortの緩和策のため）。
    return false;
  }
}

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
      // ★3回目レビュー修正5: Origin検証（上記isAllowedOriginのコメント参照）。
      // ここを通らないリクエストは、GASへもD1へも一切触れさせずに拒否する
      // （force=1もこの後にあるため、Origin不一致では絶対に有効にならない）。
      if (!isAllowedOrigin(request)) {
        return json({ status: 'error', message: '許可されていないOriginからのリクエストです' }, 403);
      }

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
        console.log('[sync] SYNC_KEY未設定のため、認証なしで /api/sync を許可しています（移行期間中の暫定運用。Origin検証・レート制限は別途有効）');
      }

      // ★修正5（レート制限）: 直近1分間に実行済みの同期がしきい値を超えていれば、
      // GASへの取得もD1への書き込みも一切行わず「進行中」と同じ扱いでスキップする。
      if (await isSyncRateLimited(env)) {
        return json({
          status: 'ok', rows: 0, skipped: true,
          message: '直近の同期回数が多いため今回はスキップしました（無料枠保護のためのレート制限）'
        });
      }

      // ★修正7（急減ガードの自己回復）: ?force=1 が付いていれば、件数急減ガード
      // （sync.js）だけを明示的に無視して受け入れる。連続拒否を待たずに今すぐ
      // 反映したい場合の脱出口。サイズ上限や応答形式の検証はforceでも無条件のまま。
      // ★修正5: force=1はOrigin検証を通ったリクエストでしか到達しない（上のreturnで
      // 弾かれるため）。よってここに「Origin検証を通ったときだけ有効」を別途
      // 書く必要はない＝構造的に保証されている。ただし上のisAllowedOriginのコメント
      // どおり、これはブラウザ経由の悪用しか防げない（curl等はOriginを偽装できる）。
      // ★5回目レビュー修正5（高・Codex）: 前回「force は Origin検証を通過した後にしか
      // 到達しない」から「第三者はforceを使えない」と結論づけたのは言い過ぎだった
      // （curl等の直接HTTPクライアントにはOrigin検証は無力なため）。そこでforce
      // そのものの実害を小さくする対応をsync.js側に追加した：(a) forceが即時受理に
      // 効くのは日報(nippo)だけが急減しているときに限り、マスタ（職人・元請・現場）の
      // 急減を含む場合はforceは一切効かない、(b) forceによる即時受理そのものにも
      // 専用の頻度制限（直近30分に1回まで）を課す。詳細はcf/src/sync.jsの
      // FORCE_ACCEPT_MARKER周りのコメント参照。
      const force = url.searchParams.get('force') === '1';

      // ★syncAllは例外を投げない契約。同時実行中は{ok:true, skipped:true}で
      // 返る（修正2の同時実行抑止）ので、try/catchではなく戻り値のokを見る。
      const r = await syncAll(env, { force });
      // ★6回目レビュー修正1: skipReasonをそのまま画面へ返す。'unchanged'（変更なし
      // スキップ＝GASを実際に取得しD1と完全一致することを確認できた）のときだけ、
      // 画面側（sync-guard.js）はこの応答を「確実成功」として扱う。
      return json({ status: r.ok ? 'ok' : 'error', rows: r.rows, message: r.message, skipped: !!r.skipped, skipReason: r.skipReason });
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
    // ★修正8（低・sync_logの掃除）: 同期本体とは独立に、古いsync_log行を掃除する。
    // 失敗しても同期そのものには影響しない（cleanupSyncLogは例外を投げない契約）。
    ctx.waitUntil(cleanupSyncLog(env));
  }
};
