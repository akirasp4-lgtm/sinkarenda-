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
//
// ★2026-08-24 3回目レビュー（Fable 5 / Codex）で、なお2件の重大・高が残っているとの
// 判定を受け、さらに以下を修正した：
//   修正3（急減ガードの自己回復・作り直し）: 旧実装は「直近3件が急減ログか」という
//     “回数”だけで自己回復していた。Codexが「日報→職人→元請→現場と毎回まったく別の
//     欠損を起こしても、4回目が『3回連続』の条件を満たして自動受入されてしまう」ことを
//     実際に再現した（=私が入れた安全装置そのものが攻撃経路になっていた）。
//     回数ではなく「拒否した取得内容のハッシュが直近と同一で、かつ最初の拒否から
//     一定時間（30分）が経過している」ことを条件にする。中身が変わるたびに時計が
//     リセットされるため、異なる欠損を連発しても自動受入には辿り着けない。
//   修正4（世代ガードの同着対策）: WHERE条件を `>=` から `>` に変える。旧条件では
//     ロックがフェイルオープンした状態で2つの同期が同一ミリ秒に開始すると、
//     先に完了した新しい内容を、後から完了した（同じ開始ミリ秒の）古い内容が
//     上書きできてしまうことをCodexが再現した。`>` にすることで「同じ開始時刻の
//     書き込みは最初の1回しか勝てない」に変わり、後続の同着書き込みは無条件で弾かれる
//     （meta.changes=0）。真に同時始まりの2件のどちらが「本当は新しいか」は
//     決められないが、少なくとも「先に保存済みの内容が後から上書きされる」ことは無くなる。

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
// ★3回目レビュー修正3: 「回数」ではなく「同一内容の拒否が続いた時間」で自己回復を
// 判定する。最初に拒否してからこの時間が経過し、かつその間ずっと同じ内容
// （ハッシュ一致）で拒否され続けていた場合のみ、自動的に受け入れる。
// 30分＝Cron(5分間隔)なら約6回、同じ状態が続いて初めて「本当にそういうデータ
// なのだろう」とみなす。回数を基準にしないため、/api/syncを連打しても
// （＝異なる内容を次々送りつけても、同じ内容を高速に送りつけても）早められない。
const SHRINK_AUTO_ACCEPT_MIN_ELAPSED_MS = 30 * 60 * 1000;
// 「直近が同一ハッシュで何件連続拒否されているか」を遡って見るための上限。
// Cron間隔(5分)基準なら30分でも6件だが、書き込み後の即時同期等で呼び出し頻度が
// 上がっても十分に遡れるよう余裕を持たせてある（多く見ても実害は無い＝古い方向に
// 安全側なだけで、自動受入を早めることはない）。
const SHRINK_LOG_SCAN_LIMIT = 500;

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
// ★3回目レビュー修正3: payloadHash（今回取得した内容のハッシュ。取得自体が失敗した
// ときはnull）も一緒に記録する。急減ガードの自己回復が「同一内容の拒否が何分
// 続いているか」を判定するのに使う（下記 sameHashShrinkRejectStreak 参照）。
async function writeSyncLog(env, { rows, ok, message, payloadHash }) {
  const at = new Date().toISOString();
  await env.DB.prepare('INSERT OR REPLACE INTO sync_log (at,rows,ok,message,payload_hash) VALUES (?,?,?,?,?)')
    .bind(at, rows, ok, message, payloadHash || null).run();
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

// ★3回目レビュー修正3（作り直し）: 急減ガードの自己回復を「回数」ではなく
// 「同一内容（ハッシュ一致）の拒否が、最初の拒否からどれだけの時間続いているか」で
// 判定する。直近のsync_logを新しい順に遡り、「件数急減による拒否」かつ「今回と
// 同じハッシュ」である限り数え続け、それ以外（成功・別の理由の失敗・違う内容の
// 拒否）に当たった時点で止める。
//
// 旧実装（回数だけを見る版）は、Codexにより「日報→職人→元請→現場と毎回まったく
// 別の欠損を起こしても、4回目に“3回連続拒否”の条件を満たして自動受入されてしまう」
// ことを再現された。ハッシュを条件に加えることで、内容が変わるたびにこの判定は
// リセットされる＝異なる欠損を連続させても自動受入には辿り着けない。
async function sameHashShrinkRejectStreak(env, currentHash) {
  try {
    const res = await env.DB.prepare(
      'SELECT at, ok, message, payload_hash FROM sync_log ORDER BY at DESC LIMIT ?'
    ).bind(SHRINK_LOG_SCAN_LIMIT).all();
    const logs = res.results || [];
    let count = 0;
    let earliestAt = null; // ハッシュが一致したまま遡れた中で最も古い（＝最初の拒否）at
    for (const r of logs) {
      const isReject = Number(r.ok) === 0 && String(r.message || '').includes(SHRINK_REJECT_MARKER);
      if (!isReject || r.payload_hash !== currentHash) break;
      count++;
      earliestAt = r.at;
    }
    return { count, earliestAt };
  } catch (_e) {
    // 履歴が読めなければ「まだ実績なし」として扱う＝これまでどおり拒否を継続する安全側。
    return { count: 0, earliestAt: null };
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
 * ★世代の逆転防止（再レビュー修正2・3回目レビュー修正4）：書き込み文のWHERE条件が
 * 「今回の取得開始時刻(fetch_started_at) が保存済みのものより厳密に新しいときだけ
 * 上書きする」ことを強制する（`>`。同じ時刻での上書きは不可＝3回目レビュー修正4）。
 * 同時に2つの同期が走り、遅く始まった方が先に完了しても、先に始まった方（古い内容）が
 * 後から上書きすることはできない。同一ミリ秒に始まった2件は「先に書き込めた方」が
 * 保持され、後続は無条件で弾かれる。
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

    // ★3回目レビュー修正3: 急減ガードの自己回復（下記）がハッシュ一致の判定を
    // 必要とするため、ここで先に計算しておく（以前は変更なしスキップの直前だけで
    // 計算していた）。値そのものは変わらない＝安全なリファクタリング。
    const hash = await sha256Hex(payloadText);

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
        //
        // ★3回目レビュー修正3（作り直し）: 「回数」ではなく「同一内容（ハッシュ一致）の
        // 拒否が最初の拒否から何分続いているか」で判定する。Codexが「毎回違う内容の
        // 欠損を連発しても3回で自動受入されてしまう」ことを再現したため、回数だけの
        // 判定は廃止した。詳細は sameHashShrinkRejectStreak のコメント参照。
        if (force) {
          shrinkNote = `（★${SHRINK_REJECT_MARKER}：${shrunk.join('、')}／force=1が指定されたため強制的に受け入れました）`;
        } else {
          const streak = await sameHashShrinkRejectStreak(env, hash);
          const elapsedMs = streak.earliestAt ? Date.now() - Date.parse(streak.earliestAt) : 0;
          const stable = !!streak.earliestAt && Number.isFinite(elapsedMs) && elapsedMs >= SHRINK_AUTO_ACCEPT_MIN_ELAPSED_MS;
          if (!stable) {
            const remainMin = streak.earliestAt
              ? Math.max(1, Math.ceil((SHRINK_AUTO_ACCEPT_MIN_ELAPSED_MS - elapsedMs) / 60000))
              : Math.ceil(SHRINK_AUTO_ACCEPT_MIN_ELAPSED_MS / 60000);
            const hint = streak.earliestAt
              ? `同じ内容の拒否が既に${streak.count}回続いています。このままあと約${remainMin}分、同じ内容が続けば自動的に受け入れます。`
              : `今回が最初の拒否です。同じ内容のまま約${remainMin}分続けば自動的に受け入れます。`;
            const message = `${SHRINK_REJECT_MARKER}：${shrunk.join('、')}。GAS側の異常を疑い、書き込みを中止しました`
              + `（${hint}今すぐ反映したい場合は /api/sync?force=1 を使ってください）`;
            await safeWriteSyncLog(env, { rows: rowCount, ok: 0, message, payloadHash: hash });
            return { ok: false, rows: 0, message };
          }
          // 同一内容の拒否が30分以上続いた＝自己回復。今回は受け入れて書き込みへ進む。
          const elapsedMin = Math.round(elapsedMs / 60000);
          shrinkNote = `（★${SHRINK_REJECT_MARKER}：${shrunk.join('、')}／同一内容の拒否が${elapsedMin}分継続したため自動的に受け入れました）`;
        }
      }
    }

    // ★修正1（変更が無ければ書かない）: 夜間・休日の書き込みが0になる。
    if (existing && existing.hash === hash) {
      const message = '変更なし（書き込みをスキップしました）' + shrinkNote;
      await safeWriteSyncLog(env, { rows: rowCount, ok: 1, message, payloadHash: hash });
      return { ok: true, rows: rowCount, message, skipped: true };
    }

    const at = new Date().toISOString();
    // ★原子性・世代逆転防止の本体：単一文のINSERT ... ON CONFLICT ... WHERE。
    // WHERE条件（fetch_started_atの比較）により、より古い取得結果では
    // 保存済みのより新しい結果を上書きできない。この1文が成功するか
    // （条件不成立で）何も変えずに終わるかのどちらかであり、「一部だけ入った」
    // 中途半端な状態は無い。
    //
    // ★3回目レビュー修正4: 条件を `>=` から `>` に変更。ロックがフェイルオープンした
    // 状態で2つの同期が同一ミリ秒(Date.now())に開始すると、`>=`では「後から完了した
    // 方（内容の新旧を問わない）」が常に上書きできてしまい、Codexが「新しい内容を
    // 保存した後に古い内容が上書きする」ケースを再現した。`>`にすることで、同じ
    // fetch_started_atでの書き込みは最初の1回しか成功しなくなる（2回目以降は
    // meta.changes=0で弾かれる）。真に同一ミリ秒に始まった2件のどちらが本当に
    // 新しいかは決められないが、「既に保存済みの内容が後から上書きされる」ことは
    // 無くなる（＝最初に書き込めた方が保持される。それ以上は分からないので保守的に
    // 「何もしない」側へ倒す）。
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
       WHERE CAST(excluded.fetch_started_at AS INTEGER) > CAST(snapshot.fetch_started_at AS INTEGER)`
    ).bind(payloadText, hash, rowCount, memberCount, genbaCount, jobsitesCount, bytes, String(fetchStartedAt), at).run();

    const wrote = !!(writeRes && writeRes.meta && writeRes.meta.changes > 0);
    if (!wrote) {
      // ★世代の逆転防止が実際に働いた瞬間：この取得結果と同じか新しいものが
      // 既に保存されている。取得自体は正常に終わっているのでok:trueとし、
      // 「書かなかったこと」をskippedで表す（失敗ではない）。
      const message = 'より新しい（または同じ開始時刻の）取得結果が既に保存されているため、今回の取得内容は書き込みませんでした（正常な動作です）' + shrinkNote;
      await safeWriteSyncLog(env, { rows: rowCount, ok: 1, message, payloadHash: hash });
      return { ok: true, rows: rowCount, message, skipped: true };
    }

    const message = shrinkNote || '';
    await safeWriteSyncLog(env, { rows: rowCount, ok: 1, message, payloadHash: hash });
    return { ok: true, rows: rowCount, message };
  } catch (e) {
    const message = String((e && e.message) || e);
    await safeWriteSyncLog(env, { rows: 0, ok: 0, message });
    return { ok: false, rows: 0, message };
  } finally {
    if (locked) await releaseLock(env);
  }
}

// ★修正8（低・sync_logの掃除）: sync_logはCronのたびに1行増える（5分間隔で最大288行/日）。
// 何もしないと無限に増え続けるため、保持期間を過ぎた行をCronのたびに削除する。
// 30日より新しい行は、読み取り側の鮮度ガード（15分）・急減ガードの自己回復（30分）
// どちらの判定にも十分すぎるほど余裕がある。呼び出し元（scheduled()）はsyncAllとは
// 独立にこれを呼ぶため、掃除の成否が同期そのものに影響しないようにする（例外を
// 投げない契約はsyncAllと揃える）。
const SYNC_LOG_RETENTION_MS = 30 * 24 * 60 * 60 * 1000;

export async function cleanupSyncLog(env) {
  try {
    const cutoff = new Date(Date.now() - SYNC_LOG_RETENTION_MS).toISOString();
    await env.DB.prepare('DELETE FROM sync_log WHERE at < ?').bind(cutoff).run();
    return { ok: true };
  } catch (e) {
    // 掃除の失敗は同期本体に影響させない（ログ肥大化は障害調査の材料が増えるだけで、
    // アプリの正しさには関わらないため）。
    return { ok: false, message: String((e && e.message) || e) };
  }
}
