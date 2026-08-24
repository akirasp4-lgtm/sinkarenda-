// GASの doGet(compact=1) からスプレッドシートの内容を取り込み、D1へ入れる。
// ★ここは「読むだけ」。スプレッドシートには何も書かない。
// ★D1はあくまで派生コピー。壊れても全件取り込み直せば完全に戻る。
//
// ★2026-08-24 最終総合レビュー（Fable 5 / Codex 両者）で「切り替え不可」の判定を受け、
// D1の持ち方を「行ごとのテーブル（nippo/members/genba/jobsites）」から
// 「スナップショット1行（snapshotテーブル）」へ全面的に変更した。理由は schema.sql
// のコメント参照。ここではその実装（取り込み側）を担う。
//
// ★2026-08-24 再レビュー（同2者）で、スナップショット方式のままなお「切り替え不可」の
// 判定を受け、さらに以下を修正した（詳細は各関数のコメント）：
//   修正2: sync_lockの取得を単一のSQL文にし、かつsnapshotの書き込み自体にも
//          「取得開始時刻が保存済みより新しいときだけ」というWHERE条件を付けた
//          （世代の逆転防止の本体）。
//   修正3: members/genbaMaster/jobsitesにも日報と同じ半減チェックを適用した。
//   修正7: 半減チェックで拒否し続けると自己回復しない失敗ループになるため、
//          連続拒否がしきい値を超えたら自動的に受け入れる／force=1での明示的上書きを設けた。

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
// ★このロックはあくまで「無駄な二重実行を減らす」best-effortであり、
// 正しさそのもの（新しいデータが古いデータに上書きされないこと）は
// snapshot書き込みのWHERE条件（fetch_started_at比較）が担う＝下記syncAll参照）。
const LOCK_STALE_MS = 90_000;

// 修正7（急減ガードの自己回復）。件数半減で拒否したことを示す固定の目印文字列。
// sync_logのmessageにこれが含まれる行を「拒否ログ」として数える。
const SHRINK_REJECT_MARKER = '件数が急減しました';
// 連続でこの回数だけ拒否が続いたら、以後は自動的に受け入れる（自己回復）。
// Cronは5分間隔のため、3回連続＝最大15分は安全側（拒否＝既存の古いデータを保持）に倒し、
// それでも状況が変わらなければ「本当にそういうデータなのだろう」とみなして受け入れる。
const SHRINK_AUTO_ACCEPT_AFTER_N = 3;

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

// ★修正2（同時実行の抑止・再レビュー対応）。以前は「SELECTで確認」→「INSERTで取得」の
// 2文に分かれており、2つの同期が両方SELECTを終えてからINSERTすると両方が取得成功
// できてしまっていた（Codexが再現）。ここでは取得そのものを単一のSQL文にし、
// D1が返す meta.changes で「自分が取得できたか」を判定する（2文の間に割り込む
// 余地を無くす）。
// ★ただしこのロックは「無駄な二重実行を減らす」ためのbest-effortに過ぎない。
// 正しさの最終防衛はsnapshot書き込み自体のWHERE条件（下記syncAll）が担う。
async function tryAcquireLock(env) {
  try {
    const now = Date.now();
    const staleCutoff = now - LOCK_STALE_MS;
    const res = await env.DB.prepare(
      `INSERT INTO sync_lock (id, locked_at) VALUES (1, ?)
       ON CONFLICT(id) DO UPDATE SET locked_at = excluded.locked_at
       WHERE sync_lock.locked_at IS NULL OR CAST(sync_lock.locked_at AS INTEGER) < ?`
    ).bind(String(now), String(staleCutoff)).run();
    // meta.changes > 0 なら「自分がこの1文で更新できた」＝取得成功。
    // 0（WHERE条件が不成立で更新されなかった）なら他が進行中とみなしスキップする。
    return !!(res && res.meta && res.meta.changes > 0);
  } catch (_e) {
    // ロック機構自体が読めない/書けない場合でも同期は続行する（フェイルオープン）。
    // データの正しさは snapshot への単一文・条件付き書き込みが担保しているため、
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

// 修正7（急減ガードの自己回復）。直近のsync_logを新しい順に見て、
// 「件数急減による拒否」が何回連続しているかを数える。拒否以外（成功や別の理由の
// 失敗）に当たった時点で数えるのをやめる＝あくまで「直近ずっと同じ理由で
// 拒否され続けているか」を見る。
async function recentConsecutiveShrinkRejections(env) {
  try {
    const res = await env.DB.prepare(
      'SELECT ok, message FROM sync_log ORDER BY at DESC LIMIT ?'
    ).bind(SHRINK_AUTO_ACCEPT_AFTER_N).all();
    const logs = (res.results || []).slice(0, SHRINK_AUTO_ACCEPT_AFTER_N);
    let count = 0;
    for (const r of logs) {
      if (Number(r.ok) === 0 && String(r.message || '').includes(SHRINK_REJECT_MARKER)) count++;
      else break;
    }
    return count;
  } catch (_e) {
    // 履歴が読めなければ「まだ0回」として扱う＝これまでどおり拒否を継続する安全側。
    return 0;
  }
}

/**
 * GASから取得してD1のスナップショットを更新する。
 *
 * ★契約：この関数は例外を投げない。失敗は必ず戻り値の `ok === false` で表す。
 * 呼び出し側は try/catch を書かず `result.ok` だけを見ればよい。
 *
 * ★原子性（重大1の解決）：書き込みは単一のSQL文（INSERT ... ON CONFLICT ... WHERE）
 * のみ。DELETE+複数INSERTのような「複数の文にまたがる書き込み」が無いため、
 * 途中状態が外部の読み取りに見える瞬間が原理的に存在しない。
 *
 * ★世代の逆転防止（再レビュー修正2）：書き込み文のWHERE条件が「今回の取得開始時刻
 * (fetch_started_at) が保存済みのものと同じか新しいときだけ上書きする」ことを強制する。
 * 同時に2つの同期が走り、遅く始まった方が先に完了しても、先に始まった方（古い内容）が
 * 後から上書きすることはできない。
 *
 * ★費用（重大2の解決）：1回の同期の書き込みは常に1行（変更が無ければ0行）。
 * Cronが5分ごと(288回/日)に走っても最大288行/日で、無料枠10万行/日の0.3%。
 *
 * @param {object} env
 * @param {{force?: boolean}} [opts] force:true のとき、件数急減ガード（修正3/修正7）
 *   のみを無条件で通過させる（他の検証＝応答形式・サイズ上限は無条件のまま維持する）。
 *   利用者が管理画面等から明示的に「今すぐ反映したい」場合の脱出口（修正7）。
 *
 * 返り値: { ok, rows, message, skipped? }
 *   - skipped: true のときは「進行中のためスキップ」「変更なしのためスキップ」
 *     「より新しい取得結果が既に保存されているためスキップ」のいずれか
 *     （message で区別できる）。いずれも ok:true。
 */
export async function syncAll(env, opts = {}) {
  const force = !!(opts && opts.force);
  let locked = false;
  try {
    locked = await tryAcquireLock(env);
    if (!locked) {
      return { ok: true, rows: 0, message: '前回の同期が進行中のため今回はスキップしました', skipped: true };
    }

    // ★修正2（再レビュー対応）: GASへの取得を開始する直前の時刻を記録しておく。
    // これが「どちらの取得結果がより新しいか」の唯一の判定基準になる
    // （完了した時刻ではなく、開始した時刻で比較することで、後から始まった
    // 取得の結果が常に先に始まった取得の結果より「新しい」と正しく判定できる）。
    const fetchStartedAt = Date.now();
    const url = env.GAS_URL + '?compact=1&company=&t=' + fetchStartedAt;
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
    const memberCount = sanitized.members.length;
    const genbaCount = sanitized.genbaMaster.length;
    const jobsitesCount = sanitized.jobsites.length;

    // ★修正1（サイズガード）: 黙って上限に当たって壊れるのを防ぐ。force=1でも
    // これは無条件（D1の実際の上限に関わる安全装置のため、明示的な上書き手段の対象外）。
    if (bytes > SIZE_LIMIT_BYTES) {
      const message = `payloadが上限(${SIZE_LIMIT_BYTES}バイト)を超えました（${bytes}バイト）。書き込みを中止しました`;
      await safeWriteSyncLog(env, { rows: rowCount, ok: 0, message });
      return { ok: false, rows: 0, message };
    }

    const existingRes = await env.DB.prepare(
      'SELECT rows, hash, members_count, genba_count, jobsites_count FROM snapshot WHERE id = 1'
    ).all();
    const existing = (existingRes.results && existingRes.results[0]) || null;

    // ★修正3（急激な件数減少ガード。マスタにも適用）: GAS側の障害・誤設定で
    // 全消え・大幅減少するのを防ぐ。日報(rows)だけでなく職人・元請・現場の
    // 3マスタにも同じ「半分未満なら拒否」を適用する（レビュー指摘: 以前はrowsにしか
    // 無く、members:[]のような全消えを正常として受け入れていた）。
    // 初回（保存済みが無い）ときはこの検査を飛ばす。
    let shrinkNote = '';
    if (existing) {
      const shrunk = [];
      if (rowCount < existing.rows / 2) shrunk.push(`日報(${existing.rows}→${rowCount})`);
      if (memberCount < existing.members_count / 2) shrunk.push(`職人マスタ(${existing.members_count}→${memberCount})`);
      if (genbaCount < existing.genba_count / 2) shrunk.push(`元請マスタ(${existing.genba_count}→${genbaCount})`);
      if (jobsitesCount < existing.jobsites_count / 2) shrunk.push(`現場マスタ(${existing.jobsites_count}→${jobsitesCount})`);

      if (shrunk.length > 0) {
        // ★修正7（急減ガードの自己回復）: アーカイブ等の正当な操作でも件数は
        // 半分以下になりうる。拒否し続けるだけだと、正しい操作の後でも人手での
        // 復旧が必要になる失敗ループになる（レビュー指摘）。
        if (force) {
          shrinkNote = `（★${SHRINK_REJECT_MARKER}：${shrunk.join('、')}／force=1が指定されたため強制的に受け入れました）`;
        } else {
          const consecutive = await recentConsecutiveShrinkRejections(env);
          if (consecutive < SHRINK_AUTO_ACCEPT_AFTER_N) {
            const remain = SHRINK_AUTO_ACCEPT_AFTER_N - consecutive - 1;
            const hint = remain > 0
              ? `あと${remain}回連続で同じ状態が続くと自動的に受け入れます。`
              : `次回も同じ状態が続けば自動的に受け入れます。`;
            const message = `${SHRINK_REJECT_MARKER}：${shrunk.join('、')}。GAS側の異常を疑い、書き込みを中止しました`
              + `（連続${consecutive + 1}回目。${hint}今すぐ反映したい場合は /api/sync?force=1 を使ってください）`;
            await safeWriteSyncLog(env, { rows: rowCount, ok: 0, message });
            return { ok: false, rows: 0, message };
          }
          // 連続拒否がしきい値に達した＝自己回復。今回は受け入れて書き込みへ進む。
          shrinkNote = `（★${SHRINK_REJECT_MARKER}：${shrunk.join('、')}／連続${consecutive}回拒否のため自動的に受け入れました）`;
        }
      }
    }

    const hash = await sha256Hex(payloadText);

    // ★修正1（変更が無ければ書かない）: 夜間・休日の書き込みが0になる。
    if (existing && existing.hash === hash) {
      const message = '変更なし（書き込みをスキップしました）' + shrinkNote;
      await safeWriteSyncLog(env, { rows: rowCount, ok: 1, message });
      return { ok: true, rows: rowCount, message, skipped: true };
    }

    const at = new Date().toISOString();
    // ★原子性・世代逆転防止の本体：単一文のINSERT ... ON CONFLICT ... WHERE。
    // WHERE条件（fetch_started_atの比較）により、より古い取得結果では
    // 保存済みのより新しい結果を上書きできない。この1文が成功するか
    // （条件不成立で）何も変えずに終わるかのどちらかであり、「一部だけ入った」
    // 中途半端な状態は無い。
    const writeRes = await env.DB.prepare(
      `INSERT INTO snapshot (id, payload, hash, rows, members_count, genba_count, jobsites_count, bytes, fetch_started_at, at)
       VALUES (1, ?, ?, ?, ?, ?, ?, ?, ?, ?)
       ON CONFLICT(id) DO UPDATE SET
         payload = excluded.payload,
         hash = excluded.hash,
         rows = excluded.rows,
         members_count = excluded.members_count,
         genba_count = excluded.genba_count,
         jobsites_count = excluded.jobsites_count,
         bytes = excluded.bytes,
         fetch_started_at = excluded.fetch_started_at,
         at = excluded.at
       WHERE CAST(excluded.fetch_started_at AS INTEGER) >= CAST(snapshot.fetch_started_at AS INTEGER)`
    ).bind(payloadText, hash, rowCount, memberCount, genbaCount, jobsitesCount, bytes, String(fetchStartedAt), at).run();

    const wrote = !!(writeRes && writeRes.meta && writeRes.meta.changes > 0);
    if (!wrote) {
      // ★世代の逆転防止が実際に働いた瞬間：この取得結果より新しいものが
      // 既に保存されている。取得自体は正常に終わっているのでok:trueとし、
      // 「書かなかったこと」をskippedで表す（失敗ではない）。
      const message = 'より新しい取得結果が既に保存されているため、今回の取得内容は書き込みませんでした（正常な動作です）' + shrinkNote;
      await safeWriteSyncLog(env, { rows: rowCount, ok: 1, message });
      return { ok: true, rows: rowCount, message, skipped: true };
    }

    const message = shrinkNote || '';
    await safeWriteSyncLog(env, { rows: rowCount, ok: 1, message });
    return { ok: true, rows: rowCount, message };
  } catch (e) {
    const message = String((e && e.message) || e);
    await safeWriteSyncLog(env, { rows: 0, ok: 0, message });
    return { ok: false, rows: 0, message };
  } finally {
    if (locked) await releaseLock(env);
  }
}
