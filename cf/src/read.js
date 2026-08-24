// 現行GASの doGet と "同じ形" を返す。画面側を書き換えずに差し替えるため、
// キー名も順番も1つも変えない。
//
// ★2026-08-24 最終総合レビューでの設計変更：D1は「行ごとのテーブル」ではなく
// snapshotテーブルの1行（GAS compact応答をJSON文字列化した忠実な写し）を持つ。
// 読み取りは1行SELECT→JSON.parseし、会社での絞り込みはWorker内(JS)で行う。
// 絞り込みの条件は現行GASのdoGet（gas.js の1164行あたり）と完全に同じにする
// （特に「全社」と空文字の扱い、genbaMasterとjobsitesの絞り込み条件）。
//
// ★2026-08-24 再レビュー（Fable 5 / Codex 両者）で「切り替え不可」の判定を受け、
// 鮮度ガード（修正1）を追加した：snapshotが「存在するだけ」で正常返却していたため、
// 同期が失敗し続けても最後に成功した内容を永久に「正常」として返し続けてしまっていた
// （Codexが実測で再現）。ここでは直近の同期成功時刻（sync_log）を確認し、
// 一定時間より古ければstatus:'error'を返す（→ 画面側は自動でGASへフォールバックする）。
export const HEADERS = ['登録日時','作業日','元請名','現場名','氏名','役割','出勤','退勤',
  '人工','メモ','夜勤','会社','ID','更新者','色','事業部','工番','作業区分','車両'];

// 修正1（鮮度ガード）: 直近の同期成功からこの時間より古ければ「もう正常データとは
// 言えない」とみなしてエラーを返す。Cronは5分間隔のため、15分＝3回分の猶予を持たせて
// ある（一時的な取得失敗が1〜2回続いても、GASの遅延・混雑程度なら次のCronで自然に
// 回復するため無用にフォールバックしない）。
const FRESHNESS_THRESHOLD_MS = 15 * 60 * 1000;

/**
 * 保存済みスナップショット（sanitizeForStorageで保存した形＝
 * {compact,headers,rows,members,genbaMaster,jobsites}）を、companyで絞り込む。
 *
 * gas.js の doGet と完全に同じ条件にすること（レビュー指摘）:
 *   - filter = company && company !== '全社'（空文字・'全社'は絞り込みなし）
 *   - nippo(rows): 会社(kaisha)セルを比較のたびにtrimしてから完全一致で比較
 *     （GASのdoGetも読み取りのたびに requestedCompany と行の会社セルをtrimして比較。
 *     gas.js:1206「String(row[companyIdx] || '').trim() !== requestedCompany」）
 *   - members: company の完全一致のみ。genbaMasterと違い「会社が空なら通す」例外は無い
 *     （gas.js:1240「!filterByCompany || m.company === requestedCompany」）
 *   - genbaMaster: g.name が真であることを常に要求し、絞り込み時は
 *     「会社が空 or 一致」なら通す＝共通元請の扱い
 *     （gas.js:1244「g.name && (!filterByCompany || !g.company || g.company === requestedCompany)」）
 *   - jobsites: j.genba が真であることを常に要求し、絞り込み時は絞り込み後の
 *     genbaMasterに現場名(genba)が含まれることを要求する
 *     （gas.js:1256「j.genba && (!filterByCompany || allowedGenba.has(j.genba))」）
 */
export function filterSnapshot(payload, company) {
  const filter = !!(company && company !== '全社');
  const headers = payload.headers;
  const kaishaIdx = headers.indexOf('会社');

  const rows = filter
    ? payload.rows.filter(r => String((kaishaIdx >= 0 ? r[kaishaIdx] : '') ?? '').trim() === company)
    : payload.rows;

  const members = filter
    ? payload.members.filter(m => m.company === company)
    : payload.members;

  const genbaMaster = payload.genbaMaster.filter(g =>
    g.name && (!filter || !g.company || g.company === company));
  const allowedGenba = new Set(genbaMaster.map(g => g.name));
  const jobsites = payload.jobsites.filter(j =>
    j.genba && (!filter || allowedGenba.has(j.genba)));

  return {
    status: 'ok',
    compact: 1,
    headers,
    rows,
    members,
    genbaMaster,
    jobsites
  };
}

// 修正1（鮮度ガード）: sync_logのうちok=1（成功）の直近1件のatを取得する。
// ハッシュ一致で書き込みをスキップした場合もsync.jsはok=1で記録するため
// （「変更が無いだけ」を確認できた、という意味での成功）、その場合もここで
// 「直近の成功」として時刻が更新される。これが無いと、変更が無いだけの
// 日・週末が続くと鮮度ガードに誤って引っかかってしまう。
async function getLastSuccessAt(env) {
  const res = await env.DB.prepare(
    'SELECT at FROM sync_log WHERE ok = 1 ORDER BY at DESC LIMIT 1'
  ).all();
  const row = (res.results && res.results[0]) || null;
  return row ? row.at : null;
}

export async function readSchedule(env, company) {
  const res = await env.DB.prepare('SELECT payload FROM snapshot WHERE id = 1').all();
  const row = (res.results && res.results[0]) || null;
  if (!row) {
    // まだ一度も取り込みが成功していない（＝snapshotが書かれたことが無い）。
    // 空のD1を「予定ゼロ件」として返してはいけないため、GASのerror()と
    // 同じ形で返す。画面側は status!=='ok' を見て自動的にGASへ切り替わる
    // ので、利用者にはエラーすら見えない（遅くなるだけで済む）。
    return { status: 'error', message: 'まだ取り込みが行われていません' };
  }

  // ★修正1（鮮度ガード・再レビュー対応）: 以前はここで「snapshotが存在する時点で
  // 常に検証済みの内容」と判断していたが、これは誤りだった。失敗した同期は
  // snapshotを書き換えないぶん一見安全に見えるが、それは裏を返せば「同期が
  // 何日失敗し続けても、最後に成功した古い内容を無条件に正常として返し続ける」
  // ということでもある（Codexが実測で再現）。sync_logの直近の成功時刻を確認し、
  // 一定時間より古ければ画面をGASへフォールバックさせる。
  const lastSuccessAt = await getLastSuccessAt(env);
  if (!lastSuccessAt) {
    return { status: 'error', message: '同期の成功記録がありません' };
  }
  const lastSuccessMs = Date.parse(lastSuccessAt);
  if (!Number.isFinite(lastSuccessMs) || Date.now() - lastSuccessMs > FRESHNESS_THRESHOLD_MS) {
    return {
      status: 'error',
      message: `同期が長時間成功していません（最終成功: ${lastSuccessAt}）。最新のデータではない可能性があるため取得を中止しました`
    };
  }

  let payload;
  try {
    payload = JSON.parse(row.payload);
  } catch (e) {
    // 理論上は起こらない（書き込み前にJSON.stringifyしたものしか入らない）が、
    // 万一壊れていた場合にクラッシュさせず、GASへフォールバックさせる。
    return { status: 'error', message: '保存済みデータの形式が壊れています' };
  }
  return filterSnapshot(payload, company);
}
