// 社長予定をGASからD1へ取り込む。
//
// ★社員用（cf/src/sync.js）との違いと、その理由
//   1. 触るテーブルは pres_snapshot / pres_sync_log だけ。社員用の snapshot / sync_log /
//      sync_lock には絶対に触れない。混ぜると社員用の鮮度判定（read.jsのgetLastSuccessAt）と
//      レート制限（index.jsのisSyncRateLimited）が壊れる（cf/schema-president.sql参照）。
//   2. ロックテーブルを作らない。社員用の cf/schema.sql 自身が「ロックはbest-effortに
//      過ぎず、正しさの最終防衛は fetch_started_at のWHERE条件が担う」と明記しており、
//      社長予定は件数が2桁と小さく取得も速い。守りの本体である条件付き単一文だけを採用した。
//   3. GASの呼び方が違う。日報は doGet(GET) だが社長予定は doPost(POST) で、
//      毎回PINを body に載せる必要がある（gas.js:205 で照合）。
//      ★PINはURLのクエリに載せない（アクセスログ・Referer・履歴に残るため）。
//   4. マスタ（職人・元請・現場）が無いので、その急減チェックは持たない。

const SIZE_LIMIT_BYTES = 1_500_000;
const FETCH_TIMEOUT_MS = 60_000;
const FETCH_TRIES = 2;

// 急減ガードで拒否した内容が「まったく同じまま」この時間続いたら受け入れる。
// 社長が本当に大量の予定を消した場合に、永久に取り込めなくなるのを防ぐための出口。
// ★「回数」ではなく「同じ内容か」で判定する理由: 社員用のレビューで、毎回別の内容で
// 拒否させれば回数だけの判定は素通りできることが再現されたため。
export const PRES_SHRINK_AUTO_ACCEPT_MS = 30 * 60 * 1000;

const SHRINK_REJECT_MARKER = '件数が急減しました';
const SHRINK_LOG_SCAN_LIMIT = 200;
const SYNC_LOG_RETENTION_MS = 30 * 24 * 60 * 60 * 1000;

async function sha256Hex(text) {
  const bytes = new TextEncoder().encode(text);
  const digest = await crypto.subtle.digest('SHA-256', bytes);
  return Array.from(new Uint8Array(digest)).map(b => b.toString(16).padStart(2, '0')).join('');
}

async function writeLog(env, { rows, ok, message, payloadHash = null }) {
  try {
    await env.DB.prepare(
      'INSERT INTO pres_sync_log (at, rows, ok, message, payload_hash) VALUES (?, ?, ?, ?, ?)'
    ).bind(new Date().toISOString(), rows, ok, String(message || '').slice(0, 500), payloadHash).run();
  } catch (_e) {
    // 記録に失敗しても取り込み本体の成否は変えない（記録は障害調査用の付随情報）。
  }
}

/**
 * GASの pres_list をPOSTで叩く。失敗時は例外を投げる（呼び出し側で捕捉する）。
 * ★PINはbodyに入れる。URLには絶対に載せない。
 */
async function fetchPresList(gasUrl, pin, cacheBuster) {
  let last = null;
  for (let i = 0; i < FETCH_TRIES; i++) {
    try {
      const res = await fetch(gasUrl + '?t=' + cacheBuster, {
        method: 'POST',
        // GASのdoPostは本文をそのまま読むため text/plain。既存の画面側(president.html)と同じ。
        headers: { 'Content-Type': 'text/plain' },
        body: JSON.stringify({ action: 'pres_list', pin }),
        signal: AbortSignal.timeout(FETCH_TIMEOUT_MS)
      });
      if (!res.ok) { last = new Error('HTTP ' + res.status); continue; }
      return await res.json();    // HTMLが返ると例外＝リトライ対象
    } catch (e) { last = e; }
  }
  throw last || new Error('取得に失敗しました');
}

// 社長予定1件が必ず持つキー。doGet（社員の日報データ）の行はこれを持たない。
const PRES_ROW_KEYS = ['タイトル', '開始日'];
// 社員の日報データだけが持つキー。1つでもあれば社長予定ではない。
const NIPPO_ROW_KEYS = ['作業日', '氏名', '元請名', '人工', '会社', '拠点'];

function validate(json) {
  if (!json || typeof json !== 'object') {
    return { ok: false, message: '応答がJSONではありません' };
  }
  if (json.status !== 'ok') {
    return { ok: false, message: 'GASがエラーを返しました: ' + String(json.message || '(理由なし)') };
  }
  if (!Array.isArray(json.rows)) {
    return { ok: false, message: 'rowsが配列ではありません' };
  }

  // ★★2026-08-26 本番障害の再発防止（社長のカレンダーに予定が出なくなった）
  //   GASへのPOSTがときどき doGet に届き、その応答（社員の日報データ2,652件）が
  //   {status:'ok', rows:[...]} という社長予定と同じ形をしているため、ここを
  //   素通りして「社長予定」として保存されていた。さらにその後の正しい202件が
  //   急減ガードに「2652→202」と判定されて拒否され続け、間違ったデータが居座った。
  //   → 「形が合っている」だけでは足りない。中身が本当に社長予定かを確かめる。

  // (1) doGetの応答だけが持つ目印。1つでもあれば社長予定ではない。
  for (const k of ['members', 'genbaMaster', 'jobsites', 'headers', 'compact']) {
    if (json[k] !== undefined) {
      return { ok: false, message: `社長予定ではない応答です（${k} が付いている＝doGetの応答）。保存を中止しました` };
    }
  }
  // (2) 行の中身を確かめる。0件（本当に予定が無い）は正常なので通す。
  for (const r of json.rows) {
    if (!r || typeof r !== 'object') {
      return { ok: false, message: '行がオブジェクトではありません' };
    }
    if (NIPPO_ROW_KEYS.some(k => r[k] !== undefined)) {
      return { ok: false, message: '社員の日報データが混ざっています（作業日・氏名等の列がある）。保存を中止しました' };
    }
    if (!PRES_ROW_KEYS.some(k => r[k] !== undefined)) {
      return { ok: false, message: '社長予定の列（タイトル・開始日）が見当たりません。保存を中止しました' };
    }
  }
  return { ok: true, message: '' };
}

/**
 * 急減ガードの自己回復判定。
 * 「同じ payload_hash での拒否」が最初に記録された時刻から PRES_SHRINK_AUTO_ACCEPT_MS
 * 以上経っていれば受け入れる。間に別の内容の拒否や成功が挟まっていたら連続とみなさない。
 */
async function sameHashRejectedSince(env, hash, now) {
  try {
    const res = await env.DB.prepare(
      'SELECT at, ok, message, payload_hash FROM pres_sync_log ORDER BY at DESC LIMIT ?'
    ).bind(SHRINK_LOG_SCAN_LIMIT).all();
    const rows = (res.results || []);
    let earliest = null;
    for (const r of rows) {           // 新しい順に遡る
      const isShrinkReject = Number(r.ok) === 0 &&
        String(r.message || '').includes(SHRINK_REJECT_MARKER) &&
        r.payload_hash === hash;
      if (!isShrinkReject) break;     // 連続が途切れた時点で終了
      const t = Date.parse(r.at);
      if (Number.isFinite(t)) earliest = earliest === null ? t : Math.min(earliest, t);
    }
    if (earliest === null) return 0;
    return now - earliest;
  } catch (_e) {
    // 判定できないときは自己回復させない（保守的に拒否のまま）。
    return 0;
  }
}

/**
 * 社長予定を取り込む。★例外を投げない契約（社員用のsyncAllと同じ）。
 * @returns {{ok:boolean, rows:number, message:string, skipped?:boolean, skipReason?:string}}
 */
export async function syncPresident(env, opts = {}) {
  try {
    const pin = String(env.PRES_PIN || '');
    if (!pin) {
      // Cloudflareのシークレット未設定。GASを叩いても必ず認証で弾かれるので手前で止める。
      // 何も壊さない: pres_snapshot はそのまま、読み取り側は鮮度ガードで自然にGASへ落ちる。
      const message = 'PRES_PINが未設定です（npx wrangler secret put PRES_PIN で設定してください）';
      await writeLog(env, { rows: 0, ok: 0, message });
      return { ok: false, rows: 0, message };
    }

    // ★世代の判定基準は「取得を開始した時刻」。完了時刻ではない。
    // 後から始まった取得の結果が、常に先に始まった取得の結果より新しいと正しく判定できる。
    const fetchStartedAt = Number.isFinite(opts.fetchStartedAtOverride)
      ? opts.fetchStartedAtOverride
      : Date.now();

    let raw;
    try {
      raw = await fetchPresList(env.GAS_URL, pin, fetchStartedAt);
    } catch (e) {
      const message = 'GASからの取得に失敗しました: ' + String((e && e.message) || e);
      await writeLog(env, { rows: 0, ok: 0, message });
      return { ok: false, rows: 0, message };
    }

    const check = validate(raw);
    if (!check.ok) {
      await writeLog(env, { rows: 0, ok: 0, message: check.message });
      return { ok: false, rows: 0, message: check.message };
    }

    // 保存するのは rows の配列そのもの（pres_listの応答の忠実な写し）。
    const payloadText = JSON.stringify(raw.rows);
    const bytes = new TextEncoder().encode(payloadText).length;
    const rowCount = raw.rows.length;

    if (bytes > SIZE_LIMIT_BYTES) {
      const message = `payloadが上限(${SIZE_LIMIT_BYTES}バイト)を超えました（${bytes}バイト）。書き込みを中止しました`;
      await writeLog(env, { rows: rowCount, ok: 0, message });
      return { ok: false, rows: 0, message };
    }

    const hash = await sha256Hex(payloadText);

    const existingRes = await env.DB.prepare(
      'SELECT rows, hash, fetch_started_at FROM pres_snapshot WHERE id = 1'
    ).all();
    const existing = (existingRes.results && existingRes.results[0]) || null;

    // ── 急減ガード（初回は飛ばす）─────────────────────────────
    let shrinkNote = '';
    if (existing && rowCount < existing.rows / 2) {
      const elapsed = await sameHashRejectedSince(env, hash, Date.now());
      if (elapsed < PRES_SHRINK_AUTO_ACCEPT_MS) {
        const message = `${SHRINK_REJECT_MARKER}（${existing.rows}→${rowCount}）。` +
          `同じ内容が${Math.round(PRES_SHRINK_AUTO_ACCEPT_MS / 60000)}分続けば自動で受け入れます`;
        await writeLog(env, { rows: rowCount, ok: 0, message, payloadHash: hash });
        return { ok: false, rows: 0, message };
      }
      shrinkNote = `（件数が${existing.rows}→${rowCount}へ減っていますが、同じ内容が` +
        `${Math.round(elapsed / 60000)}分続いたため受け入れました）`;
    }

    // ── 変更なしスキップ ───────────────────────────────────
    // ★ok=1で記録する。「変更が無いことを確認できた」のも成功であり、
    //   これが無いと予定を触らない日が続くだけで読み取り側の鮮度ガードが誤発火する。
    if (existing && existing.hash === hash) {
      const message = '内容に変更が無いため書き込みをスキップしました' + shrinkNote;
      await writeLog(env, { rows: rowCount, ok: 1, message, payloadHash: hash });
      return { ok: true, rows: rowCount, message, skipped: true, skipReason: 'unchanged' };
    }

    // ── 書き込み（原子性・世代逆転防止の本体）──────────────────
    // 単一文のINSERT ... ON CONFLICT ... WHERE。より古い取得結果では上書きできない。
    // `>=` ではなく `>` にしてあるのは、同一ミリ秒に始まった2件で「後から完了した方が
    // 内容の新旧を問わず常に勝つ」のを防ぐため（社員用の3回目レビュー修正4と同じ）。
    const at = new Date().toISOString();
    const writeRes = await env.DB.prepare(
      `INSERT INTO pres_snapshot (id, payload, hash, rows, bytes, fetch_started_at, at)
       VALUES (1, ?, ?, ?, ?, ?, ?)
       ON CONFLICT(id) DO UPDATE SET
         payload = excluded.payload,
         hash = excluded.hash,
         rows = excluded.rows,
         bytes = excluded.bytes,
         fetch_started_at = excluded.fetch_started_at,
         at = excluded.at
       WHERE CAST(excluded.fetch_started_at AS INTEGER) > CAST(pres_snapshot.fetch_started_at AS INTEGER)`
    ).bind(payloadText, hash, rowCount, bytes, String(fetchStartedAt), at).run();

    const wrote = !!(writeRes && writeRes.meta && writeRes.meta.changes > 0);
    if (!wrote) {
      // 世代ガードが働いた＝この取得結果と同じか新しいものが既に保存されている。
      // 取得自体は正常なので ok:true。失敗ではないことを skipped で表す。
      const message = 'より新しい（または同じ開始時刻の）取得結果が既に保存されているため、今回の取得内容は書き込みませんでした（正常な動作です）';
      await writeLog(env, { rows: rowCount, ok: 1, message, payloadHash: hash });
      return { ok: true, rows: rowCount, message, skipped: true, skipReason: 'stale-generation' };
    }

    const message = `社長予定を${rowCount}件取り込みました` + shrinkNote;
    await writeLog(env, { rows: rowCount, ok: 1, message, payloadHash: hash });
    return { ok: true, rows: rowCount, message };
  } catch (e) {
    // ここへ来るのは想定外。契約どおり例外は投げず失敗として返す。
    const message = '取り込み中に想定外のエラー: ' + String((e && e.message) || e);
    await writeLog(env, { rows: 0, ok: 0, message });
    return { ok: false, rows: 0, message };
  }
}

/** 30日より古い pres_sync_log を掃除する。例外を投げない契約。 */
export async function cleanupPresSyncLog(env) {
  try {
    const cutoff = new Date(Date.now() - SYNC_LOG_RETENTION_MS).toISOString();
    await env.DB.prepare('DELETE FROM pres_sync_log WHERE at < ?').bind(cutoff).run();
  } catch (_e) { /* 掃除の失敗は同期に影響させない */ }
}
