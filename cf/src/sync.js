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

  // ★2026-08-24 設計変更：行を捨てるフィルタは全部やめる。スプレッドシートは
  // 「行を一意に識別できる列」を持たない、ただの記録の羅列。氏名やIDが空の
  // 行（車検期限リマインダー等、人が出る予定ではなく車両の期限を置いてある
  // だけの行）も正当なデータであり、GASが返した行は1行残らず、返ってきた
  // 順のままD1へ入れる（本番データ突き合わせで14行欠落が発覚した反省）。
  const nippo = [];
  for (const row of (json.rows || [])) {
    const rec = {};
    H.forEach((h, i) => { rec[COL[i]] = row[pos[h]]; });
    // ★2026-08-24 追加修正：以前はここで `rec.kosu = Number(rec.kosu) || 0;`
    // と強制変換していたが、これがスプレッドシート側で「人工」セルが
    // 計算エラー（例: #NUM!）になっている行（車検期限リマインダー行）を
    // 一律 0 に書き換えてしまい、本番データ突き合わせでGASとの値の不一致
    // （14行）として検出された。GASから来た値をそのまま渡す（強制変換しない）。
    // SQLiteの型親和性（type affinity）により、REAL列(kosu)へ数値化できない
    // 文字列（#NUM!等）を入れるとTEXTのまま保持され、通常の数値はそのまま
    // 数値として入るため、schema.sql側の列定義（kosu REAL DEFAULT 0）は
    // 変更不要（下の null→'' 正規化だけは他の列と同じく適用される）。
    // ★D1側の絞り込み（WHERE kaisha = ?）は完全一致のため、会社名セルに
    // 紛れ込んだ前後の空白をここ（格納時）で落としておく。GASのdoGetは
    // 読み取りのたびに .trim() して比較しているが、D1は毎回trimしないので、
    // ここで揃えないと「スプレッドシートでは表示されるのにCF経由だけ
    // その人の行が静かに消える」事故になる（レビュー指摘）。
    rec.kaisha = String(rec.kaisha == null ? '' : rec.kaisha).trim();
    for (const k of COL) if (rec[k] == null) rec[k] = '';
    nippo.push(rec);
  }

  // ★単価(rate)は落とす。給料情報をD1へ持ち込まない（2026-06-11の方針）。
  // ★重複排除・空値フィルタはしない。同名・同会社の重複行（元データの重複）
  // も、GASが返す順のままそのまま保持する。
  const members = (json.members || []).map(m => ({
    name: String(m.name || ''), company: String(m.company || ''), division: String(m.division || '')
  }));

  const genba = (json.genbaMaster || []).map(g => ({
    name: String(g.name || ''), company: String(g.company || '')
  }));

  const jobsites = (json.jobsites || []).map(j => ({
    genba: String(j.genba || ''), loc: String(j.loc || ''), jobNo: String(j.jobNo || ''),
    completed: j.completed ? 1 : 0, billingMethod: String(j.billingMethod || '')
  }));

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

// writeSyncLog自体が失敗しても、それを理由に本来の失敗原因(message)を
// すり替えたり握りつぶしたりしない。ここで止め、呼び出し元へは常に
// 元のmessageを返す（修正2）。
async function safeWriteSyncLog(env, entry) {
  try {
    await writeSyncLog(env, entry);
  } catch (_e) {
    // ログ書き込み自体の失敗は無視する。本来の失敗原因はentry.messageのまま
    // 呼び出し元(syncAll)が返すので、ここで追加の対応は不要。
  }
}

/**
 * GASから取得してD1へ全件入れ替えする。
 *
 * ★契約：この関数は例外を投げない。失敗は必ず戻り値の `ok === false` で表す
 * （fetch失敗・投入失敗のどちらも同じ形で返る）。呼び出し側は try/catch を
 * 書かず `result.ok` だけを見ればよい。
 *
 * 中途半端なデータを「正しいデータ」として画面に返さないための安全装置は
 * sync_log テーブルの ok 列のほうであり、こちらが本体。読み取り側（Task 3の
 * 読み取りAPI）は最新の sync_log の ok を見て、ok=0 なら「古いが一貫した
 * 状態」（前回成功時点のD1、またはGASへの自動フォールバック）を返すこと。
 * syncAll の戻り値の ok は、呼び出し元（cronハンドラ等）への直接の結果通知用。
 */
export async function syncAll(env) {
  // ★関数全体をひとつのtryで包む。fetch/parseだけでなく、文の組み立て
  // （env.DB.prepare()/bind()）もGASの想定外応答（配列やオブジェクトが
  // 混ざる等）で例外を投げうるため、「どのtryにも入っていない行」を
  // 作らない。捕まえた例外は種類を問わず同じ形（sync_logにok=0を記録し
  // てから{ok:false, rows:0, message}を返す）に正規化する。
  try {
    const url = env.GAS_URL + '?compact=1&company=&t=' + Date.now();
    const parsed = parseGasPayload(await fetchWithRetry(url, 3));

    // ★2026-08-24 設計変更：nippo/members/genba/jobsitesは複合主キーを
    // やめ、連番(seq)だけが主キーになった。行同士がぶつかって上書きされる
    // ことは起きなくなったため、"OR REPLACE" は意味を持たない（=ここでの
    // INSERTは常に新規行として入る）。重複排除はしない設計であることを
    // コードの意図として明確にするため、素の INSERT にしている。
    const stmts = [];
    // 全件入れ替え。日報は削除も起きるため差分ではなく総入れ替えにする。
    stmts.push(env.DB.prepare('DELETE FROM nippo'));
    const ins = env.DB.prepare(
      `INSERT INTO nippo (${COL.join(',')}) VALUES (${COL.map(() => '?').join(',')})`
    );
    for (const r of parsed.nippo) stmts.push(ins.bind(...COL.map(c => r[c])));

    stmts.push(env.DB.prepare('DELETE FROM members'));
    const im = env.DB.prepare('INSERT INTO members (name,company,division) VALUES (?,?,?)');
    for (const m of parsed.members) stmts.push(im.bind(m.name, m.company, m.division));

    stmts.push(env.DB.prepare('DELETE FROM genba'));
    const ig = env.DB.prepare('INSERT INTO genba (name,company) VALUES (?,?)');
    for (const g of parsed.genba) stmts.push(ig.bind(g.name, g.company));

    stmts.push(env.DB.prepare('DELETE FROM jobsites'));
    const ij = env.DB.prepare(
      'INSERT INTO jobsites (genba,loc,jobNo,completed,billingMethod) VALUES (?,?,?,?,?)');
    for (const j of parsed.jobsites) stmts.push(ij.bind(j.genba, j.loc, j.jobNo, j.completed, j.billingMethod));

    // ★500文ずつに分割して投げる（D1のbatch上限対策）。
    // 分割するとチャンク単位でしか原子性が無いため、途中で失敗すると
    // 「DELETEは済んだが一部しか入っていない」中途半端な状態が起こりうる。
    // それを正しいデータとして読み手に返すのが最悪の事故なので、
    // 失敗したら（このtry全体のどこで起きても）必ず sync_log に ok=0 を
    // 記録し、例外は投げずに {ok:false, ...} を返す。
    for (let i = 0; i < stmts.length; i += BATCH_CHUNK_SIZE) {
      const chunk = stmts.slice(i, i + BATCH_CHUNK_SIZE);
      await env.DB.batch(chunk);
    }

    await safeWriteSyncLog(env, { rows: parsed.nippo.length, ok: 1, message: '' });
    return { ok: true, rows: parsed.nippo.length, message: '' };
  } catch (e) {
    const message = String((e && e.message) || e);
    await safeWriteSyncLog(env, { rows: 0, ok: 0, message });
    return { ok: false, rows: 0, message };
  }
}
