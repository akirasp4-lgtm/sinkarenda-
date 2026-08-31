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
// ★2026-08-27 フェーズ1: 19列より後ろに増えてよい列と、その順番。
//   ここに書いた順番どおりでなければ取り込みを止める。
//   「増えた列が何であっても受け入れる」にすると、GASが別の列を足したとき
//   同期は成功しているのに画面はその列を見つけられず、静かな誤表示になる。
//   列を足すときは必ず GAS の HEADERS と同じ名前・同じ順番でここにも足すこと。
export const OPTIONAL_HEADERS = ['拠点', '部隊'];

const SIZE_LIMIT_BYTES = 1_500_000;

// Workerからのfetchが無応答のまま待ち続けるとGASへフォールバックする機会を失う
// ため、1回あたりのタイムアウトを設ける（修正3）。GASの応答時間は実測で
// 3.9〜56秒とばらつく。
// ★5回目レビュー修正6（中・Codex）: 「GAS読みタイムアウトを60秒に延ばした」という
// 前回の報告は誤りだった。延ばしたのはブラウザ側（index.html/admin.htmlの
// GAS_READ_TIMEOUT_MS）だけで、Worker内のここ（FETCH_TIMEOUT_MS）は20秒のまま
// だった。「3回リトライで合計60秒待てる」ことと「1回56秒の応答を待てる」ことは
// 別物で、GASが毎回20秒を超える混雑状態だと56秒以内に正常応答できてもWorker側の
// 取得は3回とも20秒で打ち切られて全滅してしまう（レビュー指摘）。
// → 1回あたりを実測最大56秒＋安全マージンの60秒へ引き上げる。ただし
// fetchWithRetry(url, tries)のtriesを3のままにすると最悪60秒×3=180秒待つことになり、
// Cloudflare Workersの実行時間上限（プラン・設定に依存。cronのscheduled()は
// wrangler.tomlで明示的に延長しない限りデフォルトの上限に収まる保証がない）に
// 抵触しかねない。「超える設定にするくらいならリトライ回数を減らして1回あたりを
// 長くする」という指摘どおり、下のsyncAll呼び出し側でtriesを3→2に減らす
// （最悪60秒×2=120秒。無応答での待ち時間はCPU時間ではなくI/O待ちだが、念のため
// 総待ち時間を抑える方向にした）。
const FETCH_TIMEOUT_MS = 60_000;

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

// ★4回目レビュー修正5: 上記「経過時間」だけの判定は、「最初の拒否」と「今回」の
// “2点”が同じ内容でありさえすれば、その間の観測が0回（＝Cron停止等で誰も見ていない
// 空白期間）でも成立してしまう（レビュー指摘: 「30分間ずっと同じだった」ではなく
// 「31分離れた2点で同じだった」しか確認できていない）。例: 0分に1回拒否→Cronが
// 何らかの理由で30分近く止まる→31分後に同じ欠損が再発、の2点だけで自動受理される。
// そこで以下の2条件を追加する（両方とも既存の「同一ハッシュ・経過30分」条件に
// 加えて満たす必要がある。既存条件を緩めるものではない）。
//   ①最低観測回数: 直近の拒否ログが最低でもこの件数だけ同一ハッシュで連続している
//     こと。Cron間隔(5分)で30分継続していれば通常6件は観測されるはずなので、
//     実際に継続監視できていたことの傍証にする。
//   ②直近観測の鮮度: 直近の（今回より前の）同一ハッシュ拒否ログが、今回からこの
//     時間以内であること。これが無いと「Cronが長時間止まっていた」ケースを
//     ①の回数条件だけでは弾けない（止まる前に3回以上観測済みなら回数は満たせて
//     しまうため）。
const SHRINK_AUTO_ACCEPT_MIN_COUNT = 3;
const SHRINK_AUTO_ACCEPT_MAX_GAP_MS = 10 * 60 * 1000;

// ★5回目レビュー修正3（高・Codex）: 上記②「直近観測の鮮度」は、実装では
// 「直前の1件と“今回”の間が10分以内か」しか見ていなかった。30分の履歴全体で
// 隣接する観測間隔がすべて10分以内かは確認していなかった（＝途中に長い空白が
// あっても、最後の1区間さえ短ければfreshOkが成立してしまっていた）。
// Codexの再現例: 0分・1分・2分に3回拒否→27分の空白（Cron停止）→29分・30分に
// 再度拒否、という履歴でも「最古から30分・件数5件・直近1分」を満たして自動受理
// されてしまう。
// → sameHashShrinkRejectStreakを「新しい順に遡りながら、隣接する観測（一番新しい
// 観測については“今回”）との間隔がSHRINK_AUTO_ACCEPT_MAX_GAP_MSを超えた時点で
// 遡るのを打ち切る」方式に作り直した。これにより返ってくるcount/earliestAtは
// 常に「今回から遡って、隙間なく続いている観測の連なり」だけを表す＝
// 「隣接するすべての観測間隔が規定内」であることを判定ロジック自身が保証する。
// 空白（例: 27分間Cronが止まる）があれば、その時点で遡るのを止めるため、空白より
// 前の観測実績は今回の判定に一切持ち越されない（＝空白の後、あらためて最低観測
// 回数・経過時間をゼロから積み直す必要がある。空白前の実績を使い回して早期に
// 自動受理されることは無い）。

/**
 * GAS応答の妥当性を検証する（修正3）。ここで甘い判定をすると、将来GAS側で
 * 項目名が変わったときに「members が0件のまま保存されて成功扱いになる」等の
 * 事故につながる。おかしければ書き込まず失敗として扱う。
 */
export function validateGasPayload(json) {
  if (!json || json.compact !== 1) {
    return { ok: false, message: 'compact形式の応答ではありません（?compact=1 を付けて取得すること）' };
  }
  // ★2026-08-26: 「19列ちょうど」から「先頭19列が一致していれば通す」へ緩めた。
  //   理由: 拠点（本社/関東支店）を20列目として足すとき、Workerが19列ちょうどを
  //   要求していると、GASを出した瞬間に取り込みが止まる（＝画面はGASへ落ちて
  //   遅くなるだけで壊れはしないが、無用な障害になる）。先に許容しておけば、
  //   Worker → GAS → 画面 の順で安全に出せる。
  //   ★緩めるのは「後ろに増えること」だけ。先頭19列の中身と並びは従来どおり
  //   完全一致を要求する（並びが変わったら列の意味がずれるので必ず止める）。
  if (!Array.isArray(json.headers) || json.headers.length < EXPECTED_HEADERS.length ||
      !EXPECTED_HEADERS.every((h, i) => json.headers[i] === h)) {
    return { ok: false, message: 'headersの先頭19列が現行と一致しません: ' + JSON.stringify(json && json.headers) };
  }
  // ★Codexレビュー[P2]#12: 増えた列が想定どおりのものであること。
  //   別の列が紛れ込んだまま受け入れると、同期は成功しているのに画面はその列を
  //   見つけられず全件が既定値扱いになる（静かな誤分類）。止めた方がよい。
  // ★2026-08-27 フェーズ1: 20列目だけを見る形から、OPTIONAL_HEADERS の順番を
  //   1つずつ確かめる形へ拡張した（21列目 部隊 を足すため）。
  const extraHeaders = json.headers.slice(EXPECTED_HEADERS.length);
  for (let i = 0; i < extraHeaders.length; i++) {
    const want = OPTIONAL_HEADERS[i];
    const colNo = EXPECTED_HEADERS.length + i + 1;
    if (!want) {
      return { ok: false, message: colNo + '列目は想定外の列です: ' + JSON.stringify(extraHeaders.slice(i)) };
    }
    if (extraHeaders[i] !== want) {
      return { ok: false, message: colNo + '列目が「' + want + '」ではありません: ' + JSON.stringify(extraHeaders.slice(i)) };
    }
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
export function sanitizeForStorage(json, prevQualifications) {
  return {
    compact: 1,
    headers: json.headers,
    rows: json.rows,
    members: json.members.map(m => ({
      name: String(m.name || ''), company: String(m.company || ''), division: String(m.division || ''),
      // ★2026-08-27 フェーズ1: 既定部隊と有効フラグ。ここに書かないと黙って消える
      //   （rate を意図的に落としているのと同じ仕組みのため）。
      //   activeが無い＝まだ列を足していない古いGAS応答 → 全員 有効 とみなす。
      butai: String(m.butai || ''), active: m.active !== false
    })),
    genbaMaster: json.genbaMaster,
    jobsites: json.jobsites,
    // ★2026-08-28 資格。ここに書かないと黙って消える（rate を落としているのと同じ仕組み）。
    //   GAS側で決めた項目に削ってあるので、そのまま持つ。
    // ★Codexレビュー[P2]（2026-08-28）: 資格マスタの読み取りに失敗したとき、GASは
    //   この項目ごと省いてくる。そのときに [] を書くと、D1の303件を空で
    //   上書きしてしまう。項目が無い＝「今回は分からない」なので前回のまま残す。
    //   （古いGASデプロイも項目が無い＝前回のまま。どちらでも安全側に倒れる）
    qualifications: Array.isArray(json.qualifications)
      ? json.qualifications.map(q => ({
          name: String(q.name || ''), company: String(q.company || ''),
          qual: String(q.qual || ''),
          kind: String(q.kind || ''), expires: String(q.expires || ''),
          // ★2026-08-29 取得場所。ここに書かないと黙って消える
          place: String(q.place || '')
        }))
      : (Array.isArray(prevQualifications) ? prevQualifications : [])
  };
}

// ★2026-08-31 CPU上限対策: 解析せず「生の文字列のまま」取ってくる。
//   同じ内容かどうかは生のまま比べれば分かるので、JSONの解析を後回しにできる。
//   HTML（GASのエラーページ）が返ったときは、以前と同じくやり直しの対象にする。
export async function fetchTextWithRetry(url, tries = 3) {
  let last = null;
  for (let i = 0; i < tries; i++) {
    try {
      const res = await fetch(url, { signal: AbortSignal.timeout(FETCH_TIMEOUT_MS) });
      if (!res.ok) { last = new Error('HTTP ' + res.status); continue; }
      const text = await res.text();
      // 先頭だけ見る（全文の解析はしない＝ここがCPUを使わない肝）
      if (!/^\s*\{/.test(text)) { last = new Error('JSONではない応答が返りました'); continue; }
      return text;
    } catch (e) { last = e; }
  }
  throw last || new Error('取得に失敗しました');
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

async function sha256HexFromBytes(bytes) {
  const digest = await crypto.subtle.digest('SHA-256', bytes);
  return Array.from(new Uint8Array(digest)).map(b => b.toString(16).padStart(2, '0')).join('');
}

async function sha256Hex(text) {
  return sha256HexFromBytes(new TextEncoder().encode(text));
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
async function sameHashShrinkRejectStreak(env, currentHash, now) {
  const nowMs = typeof now === 'number' ? now : Date.now();
  try {
    const res = await env.DB.prepare(
      'SELECT at, ok, message, payload_hash FROM sync_log ORDER BY at DESC LIMIT ?'
    ).bind(SHRINK_LOG_SCAN_LIMIT).all();
    const logs = res.results || [];
    let count = 0;
    let earliestAt = null; // 遡れた範囲で最も古い（＝この連続監視の起点）at
    let latestAt = null;   // 遡れた範囲で最も新しい（＝今回の直前の拒否）at
    // ★4回目レビュー修正5: 最低観測回数・直近観測の鮮度も判定できるよう、
    // count・earliestAtに加えてlatestAtも返す（syncAll側の追加条件で使う）。
    // ★5回目レビュー修正3: 新しい順に遡りながら、隣接する観測（最初はnow＝今回）との
    // 間隔がSHRINK_AUTO_ACCEPT_MAX_GAP_MSを超えた時点で遡るのを打ち切る。これにより
    // 返り値は常に「今回から隙間なく続いている観測の連なり」だけを表す。
    let prevAtMs = nowMs;
    for (const r of logs) {
      const isReject = Number(r.ok) === 0 && String(r.message || '').includes(SHRINK_REJECT_MARKER);
      if (!isReject || r.payload_hash !== currentHash) break; // 内容が変わった／拒否でない→打ち切り
      const atMs = Date.parse(r.at);
      if (!Number.isFinite(atMs)) break; // 時刻が読めない行は安全側で打ち切り
      if (prevAtMs - atMs > SHRINK_AUTO_ACCEPT_MAX_GAP_MS) break; // 空白が規定を超えた→ここで打ち切り（それより古い実績は持ち越さない）
      count++;
      if (latestAt === null) latestAt = r.at;
      earliestAt = r.at;
      prevAtMs = atMs;
    }
    return { count, earliestAt, latestAt };
  } catch (_e) {
    // 履歴が読めなければ「まだ実績なし」として扱う＝これまでどおり拒否を継続する安全側。
    return { count: 0, earliestAt: null, latestAt: null };
  }
}

// ★5回目レビュー修正5（高・Codex）: 「Origin検証を通過した後にしかforce=1へ
// 到達しない」はブラウザ相手には正しいが、curl等のHTTPクライアントはOriginヘッダを
// 自由に書き換えられるため、「第三者はforceを使えない」とは言えない（前回の結論は
// 言い過ぎだった。cf/src/index.jsのisAllowedOriginのコメント参照）。
// そこで force そのものの実害を2方向で小さくする：
//   (a) forceが即時受理に効くのは「日報(nippo)だけが急減していて、マスタ
//       （職人・元請・現場）は急減していない」ときに限る。マスタが1つでも
//       急減していればforceは一切効かず、通常の自己回復（同一内容が続けば
//       時間で解除）にのみ委ねる。サイズ上限・応答形式検証はもともと無条件。
//   (b) forceによる即時受理そのものにも専用の頻度制限を課す。一般のレート制限
//       （cf/src/index.jsのisSyncRateLimited・直近1分12回）は「連打」対策であり、
//       GASが一瞬だけ半欠損を返した瞬間を狙った一度きりのforce=1は防げない
//       （＝「最初の1回」を防げていない、という指摘への対応）。直近
//       FORCE_ACCEPT_COOLDOWN_MS以内にforceで受理した実績があれば、今回はforceを
//       無効化し通常の自己回復に委ねる。
// ★正直な限界: これでも「一度も過去にforceで受理されたことがない、初めての
// force=1」自体は止められない（count-basedの仕組み全般に共通する限界）。
// あくまで「実害を小さくする」緩和策であり、force=1を完全に無害化するものではない。
const FORCE_ACCEPT_MARKER = 'force=1が指定されたため強制的に受け入れました';
const FORCE_ACCEPT_COOLDOWN_MS = 30 * 60 * 1000;

async function recentForceAcceptCount(env, now, windowMs) {
  try {
    const res = await env.DB.prepare(
      'SELECT at, ok, message, payload_hash FROM sync_log ORDER BY at DESC LIMIT ?'
    ).bind(SHRINK_LOG_SCAN_LIMIT).all();
    const logs = res.results || [];
    const cutoffMs = now - windowMs;
    let count = 0;
    for (const r of logs) {
      const atMs = Date.parse(r.at);
      if (!Number.isFinite(atMs) || atMs < cutoffMs) break; // 新しい順なので範囲外に出たら終了
      if (String(r.message || '').includes(FORCE_ACCEPT_MARKER)) count++;
    }
    return count;
  } catch (_e) {
    // 判定できなければ「実績なし」として扱う（フェイルオープン。他の安全装置と
    // 方針を揃える。誤判定の実害は「forceが本来より使いやすくなる」程度で、
    // 他の無条件の検証＝サイズ上限・応答形式・マスタ半減には影響しない）。
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
 * @param {{force?: boolean}} [opts] force:true のとき、件数急減ガード（修正3/修正7）を
 *   条件付きで通過させる（他の検証＝応答形式・サイズ上限は常に無条件のまま維持する）。
 *   利用者が管理画面等から明示的に「今すぐ反映したい」場合の脱出口（修正7）。
 *   ★5回目レビュー修正5: forceが実際に効くのは「日報(nippo)だけが急減していて、
 *   マスタ（職人・元請・現場）は急減していない」ときに限る（マスタ急減が1つでも
 *   あればforceは無効。通常の自己回復のみに委ねる）。さらに直近30分以内に一度
 *   forceで受理していれば、今回のforceは無効化される（force連打の頻度制限）。
 *   詳細はsyncAll内のforceEligible周りのコメント参照。
 *
 * 返り値: { ok, rows, message, skipped?, skipReason? }
 *   - skipped: true のときは「進行中のためスキップ」「変更なしのためスキップ」
 *     「より新しい取得結果が既に保存されているためスキップ」のいずれか
 *     （message で区別できる）。いずれも ok:true。
 *   - skipReason: 'unchanged' が付くのは「変更なしのためスキップ」のときだけ
 *     （★6回目レビュー修正1）。GASを実際に取得しD1と完全一致することを確認できた
 *     ことを示す機械可読な合図。呼び出し側（sync-guard.jsのdecideSyncOutcome）は
 *     これを「確実成功」として扱ってよい。「進行中のためスキップ」（GASへ一度も
 *     取得しに行っていない）にはこのフィールドを付けない＝区別できる。
 */
// GASのパスワードをURLの後ろに付ける。未設定なら空文字＝今までどおり。
// ★ここと pres-sync.js の2か所でしか使わないが、
//   「付け忘れると取り込みが全滅する」ので、名前を付けて目立たせておく。
export function gasKeyParam(env) {
  const k = String((env && env.CAL_TOKEN) || '').trim();
  return k ? ('&k=' + encodeURIComponent(k)) : '';
}

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
    // ★2026-08-31 GASのパスワードを付ける。
    //   gas.js の calAuthOk_ は、設定 CAL_REQUIRE_TOKEN が '1' のとき
    //   クエリの k を見る。ここに付けていなかったので、
    //   **設定を入れた瞬間に5分ごとの取り込みが全滅する**状態だった。
    //   ★秘密が未設定なら今までどおり付けない（設定前に壊さないため）。
    const url = env.GAS_URL + '?compact=1&company=&t=' + fetchStartedAt
      + gasKeyParam(env);
    // ★5回目レビュー修正6: FETCH_TIMEOUT_MSを60秒に延ばした分、リトライ回数は
    // 3→2に減らす（上のFETCH_TIMEOUT_MSのコメント参照。最悪でも60秒×2=120秒に収める）。
    const rawText = await fetchTextWithRetry(url, 2);

    // ★2026-08-31 CronのCPU上限対策（本番障害）:
    //   GASの応答は同じ内容なら1バイトも変わらない（実測で確認済み）。
    //   そして実際のログでは、ほぼ毎回「変更なし」。
    //   なので**まず生のまま比べて**、同じなら解析も組み直しも全部やめる。
    //   ここを通ると、この後の JSON.parse / sanitize / JSON.stringify を丸ごと省ける。
    const rawBytes = new TextEncoder().encode(rawText);
    const rawHash = await sha256HexFromBytes(rawBytes);

    const existingRes = await env.DB.prepare(
      'SELECT rows, hash, raw_hash, members_count, genba_count, jobsites_count FROM snapshot WHERE id = 1'
    ).all();
    const existing = (existingRes.results && existingRes.results[0]) || null;

    // ★raw_hash が NULL の行（この機能より前に保存された行）では省かない＝安全側。
    if (existing && existing.raw_hash && existing.raw_hash === rawHash) {
      const message = '変更なし（前回と同じ応答のため取り込みを省きました）';
      await safeWriteSyncLog(env, { rows: existing.rows, ok: 1, message, payloadHash: existing.hash });
      return { ok: true, rows: existing.rows, message, skipped: true, skipReason: 'unchanged' };
    }

    // ここから先は「中身が変わった回」だけが通る。重い処理はここに集めてある。
    let raw;
    try {
      raw = JSON.parse(rawText);
    } catch (e) {
      const message = '応答をJSONとして読めませんでした: ' + String((e && e.message) || e);
      await safeWriteSyncLog(env, { rows: 0, ok: 0, message });
      return { ok: false, rows: 0, message };
    }

    // ★修正3: 応答の妥当性検証。おかしければ書き込まず失敗として記録する
    // （→ 読み取り側は既存のsnapshotをそのまま返し続け、利用者には見えない）。
    const check = validateGasPayload(raw);
    if (!check.ok) {
      await safeWriteSyncLog(env, { rows: 0, ok: 0, message: check.message });
      return { ok: false, rows: 0, message: check.message };
    }

    // ★資格の項目がGAS応答に無いときだけ、前回のD1の値を読んで引き継ぐ。
    //   普段（項目がある）は読まない＝D1の読み取り枠を余計に使わない。
    let prevQuals = null;
    if (!Array.isArray(raw.qualifications)) {
      try {
        const r = await env.DB.prepare('SELECT payload FROM snapshot WHERE id = 1').all();
        const prev = r.results && r.results[0] ? JSON.parse(r.results[0].payload) : null;
        if (prev && Array.isArray(prev.qualifications)) prevQuals = prev.qualifications;
      } catch (e) {
        // ★Codexレビュー（2026-08-28）: ここで握り潰して先へ進むと、前回の資格を
        //   引き継げないまま [] を書き、D1の資格が消える。「正しく作れないものは
        //   書かない」。5分後の次の同期でやり直せばよい（読み取り側は今の
        //   スナップショットを返し続けるので、利用者には何も起きない）。
        const message = '前回の資格を読めなかったため、今回の取り込みは見送りました: ' + String(e && e.message || e);
        await safeWriteSyncLog(env, { rows: 0, ok: 0, message });
        return { ok: false, rows: 0, message };
      }
    }
    const sanitized = sanitizeForStorage(raw, prevQuals);
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

    // ★修正3（急激な件数減少ガード。マスタにも適用）: GAS側の障害・誤設定で
    // 全消え・大幅減少するのを防ぐ。日報(rows)だけでなく職人・元請・現場の
    // 3マスタにも同じ「半分未満なら拒否」を適用する（レビュー指摘: 以前はrowsにしか
    // 無く、members:[]のような全消えを正常として受け入れていた）。
    // 初回（保存済みが無い）ときはこの検査を飛ばす。
    let shrinkNote = '';
    if (existing) {
      const rowsShrunk = rowCount < existing.rows / 2;
      const memberShrunk = memberCount < existing.members_count / 2;
      const genbaShrunk = genbaCount < existing.genba_count / 2;
      const jobsitesShrunk = jobsitesCount < existing.jobsites_count / 2;
      const masterShrunk = memberShrunk || genbaShrunk || jobsitesShrunk;

      const shrunk = [];
      if (rowsShrunk) shrunk.push(`日報(${existing.rows}→${rowCount})`);
      if (memberShrunk) shrunk.push(`職人マスタ(${existing.members_count}→${memberCount})`);
      if (genbaShrunk) shrunk.push(`元請マスタ(${existing.genba_count}→${genbaCount})`);
      if (jobsitesShrunk) shrunk.push(`現場マスタ(${existing.jobsites_count}→${jobsitesCount})`);

      if (shrunk.length > 0) {
        // ★修正7（急減ガードの自己回復）: アーカイブ等の正当な操作でも件数は
        // 半分以下になりうる。拒否し続けるだけだと、正しい操作の後でも人手での
        // 復旧が必要になる失敗ループになる（レビュー指摘）。
        //
        // ★3回目レビュー修正3（作り直し）: 「回数」ではなく「同一内容（ハッシュ一致）の
        // 拒否が最初の拒否から何分続いているか」で判定する。Codexが「毎回違う内容の
        // 欠損を連発しても3回で自動受入されてしまう」ことを再現したため、回数だけの
        // 判定は廃止した。詳細は sameHashShrinkRejectStreak のコメント参照。
        //
        // ★5回目レビュー修正5（force実害の低減その1）: forceが即時受理として効くのは
        // 「日報だけが急減していて、マスタ（職人・元請・現場）は急減していない」ときに
        // 限る。日報のアーカイブ等は正当な運用でありうるが、マスタが半分以下に消える
        // ことは通常ありえず、これをforceで押し切ると職人・元請・現場のデータが
        // 本当に壊れる。マスタ急減を1つでも含む場合は、force=1が指定されていても
        // 下のelse（通常の自己回復のみ）へ進む（force実質無効）。
        const now = Date.now();
        const forceEligible = force && !masterShrunk;

        if (forceEligible) {
          // ★5回目レビュー修正5（force実害の低減その2）: forceによる即時受理そのものに
          // 専用の頻度制限を課す（recentForceAcceptCountのコメント参照）。
          const recentForceAccepts = await recentForceAcceptCount(env, now, FORCE_ACCEPT_COOLDOWN_MS);
          if (recentForceAccepts > 0) {
            const cooldownMin = Math.round(FORCE_ACCEPT_COOLDOWN_MS / 60000);
            const message = `${SHRINK_REJECT_MARKER}：${shrunk.join('、')}。force=1が指定されましたが、直近${cooldownMin}分以内に既にforceで受理済みのため、`
              + `今回は無効化しました（forceの連続使用を防ぐための頻度制限）。GAS側の異常を疑い、書き込みを中止しました`;
            await safeWriteSyncLog(env, { rows: rowCount, ok: 0, message, payloadHash: hash });
            return { ok: false, rows: 0, message };
          }
          shrinkNote = `（★${SHRINK_REJECT_MARKER}：${shrunk.join('、')}／${FORCE_ACCEPT_MARKER}）`;
        } else {
          const streak = await sameHashShrinkRejectStreak(env, hash, now);
          const elapsedMs = streak.earliestAt ? now - Date.parse(streak.earliestAt) : 0;
          const elapsedOk = !!streak.earliestAt && Number.isFinite(elapsedMs) && elapsedMs >= SHRINK_AUTO_ACCEPT_MIN_ELAPSED_MS;
          // ★4回目レビュー修正5: 「同一ハッシュの拒否が最初から30分続いた」だけでは、
          // その30分の間に観測が実質1回（＝最初の拒否と今回の2点）しか無くても成立
          // してしまう（レビュー指摘: 「31分離れた2点で同じだった」で十分成立する）。
          // 最低観測回数（countOk）も同時に満たすことを要求する。
          const countOk = streak.count >= SHRINK_AUTO_ACCEPT_MIN_COUNT;
          // ★5回目レビュー修正3: 「直近の観測が途切れていないか」は、
          // sameHashShrinkRejectStreak自身が「隣接する観測間隔が規定（10分）を超えた
          // 時点で遡るのを打ち切る」ことで保証している（count・earliestAtは常に
          // “今回”から隙間なく続く観測の連なりだけを表す）。そのためelapsedOk・countOk
          // の2条件だけで「30分間、隙間なく監視できていた」ことを表せる（旧実装の
          // freshOkのような別チェックは不要。打ち切りの仕組みそのものがfreshOkを
          // 兼ねている）。
          const stable = elapsedOk && countOk;
          if (!stable) {
            const remainMin = streak.earliestAt
              ? Math.max(1, Math.ceil((SHRINK_AUTO_ACCEPT_MIN_ELAPSED_MS - elapsedMs) / 60000))
              : Math.ceil(SHRINK_AUTO_ACCEPT_MIN_ELAPSED_MS / 60000);
            const reasons = [];
            if (!elapsedOk) reasons.push(`最初の拒否からまだ約${remainMin}分（監視に空白期間があった場合は、空白の後から数え直します）`);
            if (!countOk) reasons.push(`観測回数が${streak.count}回（最低${SHRINK_AUTO_ACCEPT_MIN_COUNT}回必要。監視に空白期間があった場合は空白の後の分だけを数えます）`);
            const hint = streak.earliestAt
              ? `同じ内容の拒否が既に${streak.count}回続いています（${reasons.join('、')}）。条件を満たせば自動的に受け入れます。`
              : `今回が最初の拒否です。同じ内容のまま約${remainMin}分・最低${SHRINK_AUTO_ACCEPT_MIN_COUNT}回続けば自動的に受け入れます。`;
            const forceHint = masterShrunk
              ? '（マスタ（職人・元請・現場）の半減を含むため、/api/sync?force=1 は無効です。同一内容が続けば自動的に受け入れます）'
              : '今すぐ反映したい場合は /api/sync?force=1 を使ってください（直近30分以内に一度forceで受理していない場合のみ有効）';
            const message = `${SHRINK_REJECT_MARKER}：${shrunk.join('、')}。GAS側の異常を疑い、書き込みを中止しました`
              + `（${hint}${forceHint}）`;
            await safeWriteSyncLog(env, { rows: rowCount, ok: 0, message, payloadHash: hash });
            return { ok: false, rows: 0, message };
          }
          // 同一内容の拒否が、最低観測回数・観測の連続性（途中に空白が無いこと）を
          // 満たしたまま30分以上続いた＝自己回復。今回は受け入れて書き込みへ進む。
          const elapsedMin = Math.round(elapsedMs / 60000);
          shrinkNote = `（★${SHRINK_REJECT_MARKER}：${shrunk.join('、')}／同一内容の拒否が${elapsedMin}分継続したため自動的に受け入れました）`;
        }
      }
    }

    // ★修正1（変更が無ければ書かない）: 夜間・休日の書き込みが0になる。
    // ★6回目レビュー修正1（高・両者一致）: skipReason:'unchanged' を付ける。
    // 画面側（sync-guard.jsのdecideSyncOutcome）はこれを「GASを実際に取得し、
    // D1と完全一致することを確認できた＝確実成功」として扱う（ロック競合による
    // skip「前回の同期が進行中のためスキップ」にはこの値を付けないため、
    // GASへ一度も取得しに行っていないskipと区別できる）。
    if (existing && existing.hash === hash) {
      // ★2026-08-31: ここまで来た＝「生の応答は違って見えたが、組み直したら中身は同じ」。
      //   このとき raw_hash を控えておかないと、次回も同じ重い処理を繰り返す。
      //   （raw_hash は「変わった回」の書き込みでしか入らないので、
      //     内容が長く変わらない期間はいつまでも省けないままになる）
      //   payload は今保存されている物と同じだと確認済みなので、印だけ付け替えて安全。
      if (existing.raw_hash !== rawHash) {
        try {
          await env.DB.prepare('UPDATE snapshot SET raw_hash = ? WHERE id = 1 AND hash = ?')
            .bind(rawHash, hash).run();
        } catch (e) {
          // 控えられなくても実害は「次も省けない」だけ。取り込み自体は成功している。
        }
      }
      const message = '変更なし（書き込みをスキップしました）' + shrinkNote;
      await safeWriteSyncLog(env, { rows: rowCount, ok: 1, message, payloadHash: hash });
      return { ok: true, rows: rowCount, message, skipped: true, skipReason: 'unchanged' };
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
      `INSERT INTO snapshot (id, payload, hash, raw_hash, rows, members_count, genba_count, jobsites_count, bytes, fetch_started_at, at)
       VALUES (1, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
       ON CONFLICT(id) DO UPDATE SET
         payload = excluded.payload,
         hash = excluded.hash,
         raw_hash = excluded.raw_hash,
         rows = excluded.rows,
         members_count = excluded.members_count,
         genba_count = excluded.genba_count,
         jobsites_count = excluded.jobsites_count,
         bytes = excluded.bytes,
         fetch_started_at = excluded.fetch_started_at,
         at = excluded.at
       WHERE CAST(excluded.fetch_started_at AS INTEGER) > CAST(snapshot.fetch_started_at AS INTEGER)`
    ).bind(payloadText, hash, rawHash, rowCount, memberCount, genbaCount, jobsitesCount, bytes, String(fetchStartedAt), at).run();

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
