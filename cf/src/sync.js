// GASの doGet(compact=1) からスプレッドシートの内容を取り込み、D1へ入れる。
// ★ここは「読むだけ」。スプレッドシートには何も書かない。
// ★D1はあくまで派生コピー。壊れても全件取り込み直せば完全に戻る。

const H = ['登録日時','作業日','元請名','現場名','氏名','役割','出勤','退勤','人工',
           'メモ','夜勤','会社','ID','更新者','色','事業部','工番','作業区分','車両'];

const COL = ['touroku','sagyoubi','motoukr','genba','shimei','yakuwari','shukkin','taikin',
             'kosu','memo','yakin','kaisha','id','koushinsha','iro','jigyoubu','kouban',
             'sagyou_kubun','sharyou'];

// D1のbatch()には1回あたりの文数上限があるため、これより多い文数は分割して投げる。
const BATCH_CHUNK_SIZE = 500;

export function parseGasPayload(json) {
  if (!json || json.compact !== 1 || !Array.isArray(json.headers)) {
    throw new Error('compact形式の応答ではありません（?compact=1 を付けて取得すること）');
  }
  // ヘッダの並びがGAS側で変わっても壊れないよう、名前で位置を引く
  const pos = {};
  json.headers.forEach((h, i) => { pos[h] = i; });

  const nippo = [];
  for (const row of (json.rows || [])) {
    const rec = {};
    H.forEach((h, i) => { rec[COL[i]] = row[pos[h]]; });
    if (!rec.id) continue;                       // 主キーにできない行は捨てる
    if (!rec.sagyoubi || !rec.shimei) continue;  // 同上
    rec.kosu = Number(rec.kosu) || 0;
    for (const k of COL) if (rec[k] == null) rec[k] = '';
    nippo.push(rec);
  }

  // ★単価(rate)は落とす。給料情報をD1へ持ち込まない（2026-06-11の方針）。
  const members = (json.members || []).map(m => ({
    name: String(m.name || ''), company: String(m.company || ''), division: String(m.division || '')
  })).filter(m => m.name);

  const genba = (json.genbaMaster || []).map(g => ({
    name: String(g.name || ''), company: String(g.company || '')
  })).filter(g => g.name);

  const jobsites = (json.jobsites || []).map(j => ({
    genba: String(j.genba || ''), loc: String(j.loc || ''), jobNo: String(j.jobNo || ''),
    completed: j.completed ? 1 : 0, billingMethod: String(j.billingMethod || '')
  })).filter(j => j.genba && j.loc);

  return { nippo, members, genba, jobsites };
}

export async function fetchWithRetry(url, tries = 3) {
  let last = null;
  for (let i = 0; i < tries; i++) {
    try {
      const res = await fetch(url);
      if (!res.ok) { last = new Error('HTTP ' + res.status); continue; }
      return await res.json();   // HTMLが返ると例外になる＝リトライ対象
    } catch (e) { last = e; }
  }
  throw last || new Error('取得に失敗しました');
}

// sync_log に1行残す（取り込みの成功/失敗どちらも記録する。障害調査用）。
async function writeSyncLog(env, { rows, ok, message }) {
  const at = new Date().toISOString();
  await env.DB.prepare('INSERT OR REPLACE INTO sync_log (at,rows,ok,message) VALUES (?,?,?,?)')
    .bind(at, rows, ok, message).run();
}

export async function syncAll(env) {
  const url = env.GAS_URL + '?compact=1&company=&t=' + Date.now();
  let parsed;
  try {
    parsed = parseGasPayload(await fetchWithRetry(url, 3));
  } catch (e) {
    return { ok: false, rows: 0, message: String(e.message || e) };
  }

  const stmts = [];
  // 全件入れ替え。日報は削除も起きるため差分ではなく総入れ替えにする。
  stmts.push(env.DB.prepare('DELETE FROM nippo'));
  const ins = env.DB.prepare(
    `INSERT OR REPLACE INTO nippo (${COL.join(',')}) VALUES (${COL.map(() => '?').join(',')})`
  );
  for (const r of parsed.nippo) stmts.push(ins.bind(...COL.map(c => r[c])));

  stmts.push(env.DB.prepare('DELETE FROM members'));
  const im = env.DB.prepare('INSERT OR REPLACE INTO members (name,company,division) VALUES (?,?,?)');
  for (const m of parsed.members) stmts.push(im.bind(m.name, m.company, m.division));

  stmts.push(env.DB.prepare('DELETE FROM genba'));
  const ig = env.DB.prepare('INSERT OR REPLACE INTO genba (name,company) VALUES (?,?)');
  for (const g of parsed.genba) stmts.push(ig.bind(g.name, g.company));

  stmts.push(env.DB.prepare('DELETE FROM jobsites'));
  const ij = env.DB.prepare(
    'INSERT OR REPLACE INTO jobsites (genba,loc,jobNo,completed,billingMethod) VALUES (?,?,?,?,?)');
  for (const j of parsed.jobsites) stmts.push(ij.bind(j.genba, j.loc, j.jobNo, j.completed, j.billingMethod));

  // ★500文ずつに分割して投げる（D1のbatch上限対策）。
  // 分割するとチャンク単位でしか原子性が無いため、途中で失敗すると
  // 「DELETEは済んだが一部しか入っていない」中途半端な状態が起こりうる。
  // それを正しいデータとして読み手に返すのが最悪の事故なので、
  // 失敗したら必ず sync_log に ok=0 を記録してから例外を投げ、
  // 呼び出し元（読み取り側）が異常だと分かるようにする。
  try {
    for (let i = 0; i < stmts.length; i += BATCH_CHUNK_SIZE) {
      const chunk = stmts.slice(i, i + BATCH_CHUNK_SIZE);
      await env.DB.batch(chunk);
    }
  } catch (e) {
    const message = String((e && e.message) || e);
    await writeSyncLog(env, { rows: 0, ok: 0, message });
    throw e;
  }

  await writeSyncLog(env, { rows: parsed.nippo.length, ok: 1, message: '' });
  return { ok: true, rows: parsed.nippo.length, message: '' };
}
