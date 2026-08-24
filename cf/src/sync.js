// GASの doGet(compact=1) からスプレッドシートの内容を取り込み、D1へ入れる。
// ★ここは「読むだけ」。スプレッドシートには何も書かない。
// ★D1はあくまで派生コピー。壊れても全件取り込み直せば完全に戻る。
//
// ★2026-08-24 最終総合レビュー（Fable 5 / Codex 両者）で「切り替え不可」の判定を受け、
// D1の持ち方を「行ごとのテーブル（nippo/members/genba/jobsites）」から
// 「スナップショット1行（snapshotテーブル）」へ全面的に変更した。理由は schema.sql
// のコメント参照。ここではその実装（取り込み側）を担う。

const EXPECTED_HEADERS = ['登録日時','作業日','元請名','現場名','氏名','役割','出勤','退勤',
  '人工','メモ','夜勤','会社','ID','更新者','色','事業部','工番','作業区分','車両'];

// D1の1行あたりの上限は2,000,000バイト。実測の全社compactは約701,421バイト（35%使用）。
// 上限に近づいたら黙って壊れる前に「失敗」として記録する（修正1のサイズガード）。
const SIZE_LIMIT_BYTES = 1_500_000;

// Workerからのfetchが無応答のまま待ち続けるとGASへフォールバックする機会を失う
// ため、1回あたりのタイムアウトを設ける（修正3）。GASの応答時間は実測で
// 3.9〜56秒とばらつくため、20秒で打ち切ってリトライに回す。
const FETCH_TIMEOUT_MS = 20_000;

// 同時実行の抑止（修正2）。fetchの最大リトライ(3回×20秒)を踏まえても収まるよう
// 余裕を持たせてある。これより古いロックは「前回が例外的に解放されないまま
// 終わった」とみなして上書きする（ロックが永久に固まらないための安全弁。
// 正しさ自体はロックではなく snapshot への単一文書き込みの原子性が担保する）。
const LOCK_STALE_MS = 90_000;

/**
 * GAS応答の妥当性を検証する（修正3）。ここで甘い判定をすると、将来GAS側で
 * 項目名が変わったときに「members が0件のまま保存されて成功扱いになる」等の
 * 事故につながる。おかしければ書き込まず失敗として扱う。
 */
export function validateGasPayload(json) {
  if (!json || json.compact !== 1) {
    return { ok: false, message: 'compact形式の応答ではありません（?compact=1 を付けて取得すること）' };
  }
  if (!Array.isArray(json.headers) || json.headers.length !== EXPECTED_HEADERS.length ||
      !json.headers.every((h, i) => h === EXPECTED_HEADERS[i])) {
    return { ok: false, message: 'headersが現行の19列・並びと一致しません: ' + JSON.stringify(json && json.headers) };
  }
  if (!Array.isArray(json.rows) || !Array.isArray(json.members) ||
      !Array.isArray(json.genbaMaster) || !Array.isArray(json.jobsites)) {
    return { ok: false, message: 'rows/members/genbaMaster/jobsitesのいずれかが配列ではありません' };
  }
  return { ok: true, message: '' };
}

/**
 * D1へ保存する形へ整える。★単価(rate)はここで落とす。給料情報をD1へ持ち込まない
 * （2026-06-11の方針。admin.htmlの職人管理はbackendに関わらず常にGASから直接
 * 取り直す設計に確定済みなので、D1側の応答にrateが無いことが前提になっている）。
 * rows/genbaMaster/jobsitesはGASが返した内容・順序をそのまま保つ
 * （「忠実な写し」の方針。2026-08-24の設計変更。重複行や氏名が空の行も捨てない）。
 */
export function sanitizeForStorage(json) {
  return {
    compact: 1,
    headers: json.headers,
    rows: json.rows,
    members: json.members.map(m => ({
      name: String(m.name || ''), company: String(m.company || ''), division: String(m.division || '')
    })),
    genbaMaster: json.genbaMaster,
    jobsites: json.jobsites
  };
}

export async function fetchWithRetry(url, tries = 3) {
  let last = null;
  for (let i = 0; i < tries; i++) {
    try {
      // ★修正3: 無応答のまま待ち続けるとGASへ移れないため、1回あたりの
      // タイムアウトを必ず付ける。
      const res = await fetch(url, { signal: AbortSignal.timeout(FETCH_TIMEOUT_MS) });
      if (!res.ok) { last = new Error('HTTP ' + res.status); continue; }
      return await res.json();   // HTMLが返ると例外になる＝リトライ対象
    } catch (e) { last = e; }
  }
  throw last || new Error('取得に失敗しました');
}

async function sha256Hex(text) {
  const bytes = new TextEncoder().encode(text);
  const digest = await crypto.subtle.digest('SHA-256', bytes);
  return Array.from(new Uint8Array(digest)).map(b => b.toString(16).padStart(2, '0')).join('');
}

// sync_log に1行残す（取り込みの成功/失敗どちらも記録する。障害調査用）。
async function writeSyncLog(env, { rows, ok, message }) {
  const at = new Date().toISOString();
  await env.DB.prepare('INSERT OR REPLACE INTO sync_log (at,rows,ok,message) VALUES (?,?,?,?)')
    .bind(at, rows, ok, message).run();
}

// writeSyncLog自体が失敗しても、それを理由に本来の失敗原因(message)を
// すり替えたり握りつぶしたりしない。ここで止め、呼び出し元へは常に
// 元のmessageを返す。
async function safeWriteSyncLog(env, entry) {
  try {
    await writeSyncLog(env, entry);
  } catch (_e) {
    // ログ書き込み自体の失敗は無視する。本来の失敗原因はentry.messageのまま
    // 呼び出し元(syncAll)が返すので、ここで追加の対応は不要。
  }
}

// ★修正2（同時実行の抑止）。/api/sync が並行して複数回走ると、無認証で
// 叩かれたときに重大2（書き込み量）を加速させうる。snapshot方式で原子性
// 自体は既に保証されているため、これは「無駄な重複実行を減らす」ための
// best-effort なロックであり、堅牢な分散ロックではない。
async function tryAcquireLock(env) {
  try {
    const row = await env.DB.prepare('SELECT locked_at FROM sync_lock WHERE id = 1').all();
    const lockedAt = row.results && row.results[0] ? row.results[0].locked_at : null;
    if (lockedAt != null) {
      const age = Date.now() - Number(lockedAt);
      if (Number.isFinite(age) && age >= 0 && age < LOCK_STALE_MS) return false; // 進行中とみなしスキップ
    }
    await env.DB.prepare('INSERT OR REPLACE INTO sync_lock (id, locked_at) VALUES (1, ?)')
      .bind(String(Date.now())).run();
    return true;
  } catch (_e) {
    // ロック機構自体が読めない/書けない場合でも同期は続行する（フェイルオープン）。
    // データの正しさは snapshot への単一文書き込みの原子性が担保しているため、
    // ロックが使えないことを理由に同期そのものを止める必要はない。
    return true;
  }
}

async function releaseLock(env) {
  try {
    await env.DB.prepare('INSERT OR REPLACE INTO sync_lock (id, locked_at) VALUES (1, NULL)').run();
  } catch (_e) {
    // 解放に失敗しても LOCK_STALE_MS 経過後は自動的に「進行中とみなさない」
    // 扱いに戻るため、致命的ではない。
  }
}

/**
 * GASから取得してD1のスナップショットを更新する。
 *
 * ★契約：この関数は例外を投げない。失敗は必ず戻り値の `ok === false` で表す。
 * 呼び出し側は try/catch を書かず `result.ok` だけを見ればよい。
 *
 * ★原子性（重大1の解決）：書き込みは `INSERT OR REPLACE INTO snapshot` の
 * 単一文のみ。DELETE+複数INSERTのような「複数の文にまたがる書き込み」が
 * 無いため、途中状態が外部の読み取りに見える瞬間が原理的に存在しない。
 *
 * ★費用（重大2の解決）：1回の同期の書き込みは常に1行（変更が無ければ0行）。
 * Cronが5分ごと(288回/日)に走っても最大288行/日で、無料枠10万行/日の0.3%。
 *
 * 返り値: { ok, rows, message, skipped? }
 *   - skipped: true のときは「進行中のためスキップ」または「変更なしのため
 *     書き込みをスキップ」のどちらか（message で区別できる）。どちらも ok:true。
 */
export async function syncAll(env) {
  let locked = false;
  try {
    locked = await tryAcquireLock(env);
    if (!locked) {
      return { ok: true, rows: 0, message: '前回の同期が進行中のため今回はスキップしました', skipped: true };
    }

    const url = env.GAS_URL + '?compact=1&company=&t=' + Date.now();
    const raw = await fetchWithRetry(url, 3);

    // ★修正3: 応答の妥当性検証。おかしければ書き込まず失敗として記録する
    // （→ 読み取り側は既存のsnapshotをそのまま返し続け、利用者には見えない）。
    const check = validateGasPayload(raw);
    if (!check.ok) {
      await safeWriteSyncLog(env, { rows: 0, ok: 0, message: check.message });
      return { ok: false, rows: 0, message: check.message };
    }

    const sanitized = sanitizeForStorage(raw);
    const payloadText = JSON.stringify(sanitized);
    const bytes = new TextEncoder().encode(payloadText).length;
    const rowCount = sanitized.rows.length;

    // ★修正1（サイズガード）: 黙って上限に当たって壊れるのを防ぐ。
    if (bytes > SIZE_LIMIT_BYTES) {
      const message = `payloadが上限(${SIZE_LIMIT_BYTES}バイト)を超えました（${bytes}バイト）。書き込みを中止しました`;
      await safeWriteSyncLog(env, { rows: rowCount, ok: 0, message });
      return { ok: false, rows: 0, message };
    }

    const existingRes = await env.DB.prepare('SELECT rows, hash FROM snapshot WHERE id = 1').all();
    const existing = (existingRes.results && existingRes.results[0]) || null;

    // ★修正3（急激な件数減少ガード）: GAS側の障害・誤設定で全消えするのを防ぐ。
    // 初回（保存済みが無い）ときはこの検査を飛ばす。
    if (existing && rowCount < existing.rows / 2) {
      const message = `日報の件数が急減しました（前回${existing.rows}件→今回${rowCount}件）。GAS側の異常を疑い、書き込みを中止しました`;
      await safeWriteSyncLog(env, { rows: rowCount, ok: 0, message });
      return { ok: false, rows: 0, message };
    }

    const hash = await sha256Hex(payloadText);

    // ★修正1（変更が無ければ書かない）: 夜間・休日の書き込みが0になる。
    if (existing && existing.hash === hash) {
      await safeWriteSyncLog(env, { rows: rowCount, ok: 1, message: '変更なし（書き込みをスキップしました）' });
      return { ok: true, rows: rowCount, message: '変更なし（書き込みをスキップしました）', skipped: true };
    }

    const at = new Date().toISOString();
    // ★原子性の本体：単一文のINSERT OR REPLACE。この1文が成功するか丸ごと
    // 失敗するかのどちらかであり、「一部だけ入った」中途半端な状態は無い。
    await env.DB.prepare(
      'INSERT OR REPLACE INTO snapshot (id, payload, hash, rows, bytes, at) VALUES (1, ?, ?, ?, ?, ?)'
    ).bind(payloadText, hash, rowCount, bytes, at).run();

    await safeWriteSyncLog(env, { rows: rowCount, ok: 1, message: '' });
    return { ok: true, rows: rowCount, message: '' };
  } catch (e) {
    const message = String((e && e.message) || e);
    await safeWriteSyncLog(env, { rows: 0, ok: 0, message });
    return { ok: false, rows: 0, message };
  } finally {
    if (locked) await releaseLock(env);
  }
}
