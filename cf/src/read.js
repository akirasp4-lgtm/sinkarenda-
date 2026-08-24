// 現行GASの doGet と "同じ形" を返す。画面側を書き換えずに差し替えるため、
// キー名も順番も1つも変えない。
export const HEADERS = ['登録日時','作業日','元請名','現場名','氏名','役割','出勤','退勤',
  '人工','メモ','夜勤','会社','ID','更新者','色','事業部','工番','作業区分','車両'];

const COL = ['touroku','sagyoubi','motoukr','genba','shimei','yakuwari','shukkin','taikin',
             'kosu','memo','yakin','kaisha','id','koushinsha','iro','jigyoubu','kouban',
             'sagyou_kubun','sharyou'];

export function buildResponse(nippo, members, genba, jobsites) {
  return {
    status: 'ok',
    compact: 1,
    headers: HEADERS,
    rows: nippo.map(r => COL.map(c => (r[c] == null ? '' : r[c]))),
    members: members.map(m => ({ name: m.name, company: m.company, division: m.division })),
    genbaMaster: genba.map(g => ({ name: g.name, company: g.company })),
    jobsites: jobsites.map(j => ({
      genba: j.genba, loc: j.loc, jobNo: j.jobNo,
      completed: !!j.completed, billingMethod: j.billingMethod
    }))
  };
}

// ★計画からの裁定（変更1）：取り込み(syncAll)は500文ずつ分割してD1へ投げるため、
// 「全部DELETEしたあと途中のchunkで失敗した」状態が起こりうる。そのときD1には
// 一部だけ入った中途半端なデータが残る。これを正常なデータとして画面へ返すと、
// 利用者には予定が半分消えたように見えてしまう（sync_log自体はTask 2で必ず
// ok=0を記録する設計になっている＝そちらが本体の安全装置、こちらはそれを読む側）。
//
// なので読み取り側は、返す前に必ず sync_log の最新1行を見る：
//   - 1行も無い（＝まだ一度も取り込んでいない） → エラー扱い
//   - 最新行の ok が 0（＝直近の取り込みが失敗） → エラー扱い
// エラーを返せば、画面側は status!=='ok' を見て自動的に既存のGASへ切り替わる
// ので、利用者にはエラーすら見えない。遅くなるだけで済む。
// 空のD1を「予定ゼロ件」として返してはいけない。
async function checkSyncStatus(env) {
  const last = await env.DB.prepare('SELECT * FROM sync_log ORDER BY at DESC LIMIT 1').all();
  const row = (last.results && last.results[0]) || null;
  if (!row) {
    return { ok: false, message: 'まだ取り込みが行われていません' };
  }
  if (!row.ok) {
    return { ok: false, message: row.message || '直近の取り込みに失敗しています' };
  }
  return { ok: true, message: '' };
}

export async function readSchedule(env, company) {
  const syncStatus = await checkSyncStatus(env);
  if (!syncStatus.ok) {
    // 通常応答（buildResponseの形）ではなく、GASのerror()と同じ形で返す。
    return { status: 'error', message: syncStatus.message };
  }

  // ★2026-08-24 設計変更：D1はGAS応答の忠実な写し（重複排除・一意制約なし）
  // にしたため、並び順そのものが情報を持つ。取り込み順（=GASが返した順）を
  // seq の昇順で保つよう、すべてのSELECTに ORDER BY seq を付ける。
  const filter = company && company !== '全社';
  const nippo = filter
    ? await env.DB.prepare('SELECT * FROM nippo WHERE kaisha = ? ORDER BY seq').bind(company).all()
    : await env.DB.prepare('SELECT * FROM nippo ORDER BY seq').all();
  const members = filter
    ? await env.DB.prepare('SELECT * FROM members WHERE company = ? ORDER BY seq').bind(company).all()
    : await env.DB.prepare('SELECT * FROM members ORDER BY seq').all();
  const genba = await env.DB.prepare('SELECT * FROM genba ORDER BY seq').all();
  const jobsites = await env.DB.prepare('SELECT * FROM jobsites ORDER BY seq').all();

  const allowed = new Set(genba.results.filter(g => !filter || !g.company || g.company === company)
                                       .map(g => g.name));
  return buildResponse(
    nippo.results, members.results,
    genba.results.filter(g => !filter || !g.company || g.company === company),
    jobsites.results.filter(j => !filter || allowed.has(j.genba))
  );
}
