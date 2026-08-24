// 現行GASの doGet と "同じ形" を返す。画面側を書き換えずに差し替えるため、
// キー名も順番も1つも変えない。
//
// ★2026-08-24 最終総合レビューでの設計変更：D1は「行ごとのテーブル」ではなく
// snapshotテーブルの1行（GAS compact応答をJSON文字列化した忠実な写し）を持つ。
// 読み取りは1行SELECT→JSON.parseし、会社での絞り込みはWorker内(JS)で行う。
// 絞り込みの条件は現行GASのdoGet（gas.js の1164行あたり）と完全に同じにする
// （特に「全社」と空文字の扱い、genbaMasterとjobsitesの絞り込み条件）。
export const HEADERS = ['登録日時','作業日','元請名','現場名','氏名','役割','出勤','退勤',
  '人工','メモ','夜勤','会社','ID','更新者','色','事業部','工番','作業区分','車両'];

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

export async function readSchedule(env, company) {
  const res = await env.DB.prepare('SELECT payload FROM snapshot WHERE id = 1').all();
  const row = (res.results && res.results[0]) || null;
  if (!row) {
    // まだ一度も取り込みが成功していない（＝snapshotが書かれたことが無い）。
    // 空のD1を「予定ゼロ件」として返してはいけないため、GASのerror()と
    // 同じ形で返す。画面側は status!=='ok' を見て自動的にGASへ切り替わる
    // ので、利用者にはエラーすら見えない（遅くなるだけで済む）。
    //
    // ★スナップショット方式では、失敗した同期はsnapshotを一切書き換えない
    // （修正1の原子性・ガード類）。そのため「直近の同期が失敗したか」を
    // sync_logで確認する必要はない：snapshotが存在する時点で、それは常に
    // 検証済みの内容（最新の成功、または前回成功時点のまま）である。
    return { status: 'error', message: 'まだ取り込みが行われていません' };
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
