// 毎朝のアラート（依頼文の要件9）。2026-08-29。
//
// ★依頼文（原文）:
//   「9. AIアラート 毎朝、以下を自動確認する。
//     ・人員重複 ・人員不足 ・空き人員 ・翌日の現場 ・未確定案件
//     ・延期案件 ・資格不足 ・移動時間が現実的でない予定
//     問題があれば管理者へ通知する。」
//
// ★8項目のうち3つは、今のデータでは**額面どおりには出せない**。
//   代わりに何を見ているかを、この場所と通知の文面の両方に必ず書くこと。
//   黙って別の物を出すのが一番たちが悪い。
//
//   | 依頼         | 今のデータで出せるか | 代わりに見るもの |
//   |--------------|---------------------|------------------|
//   | 人員不足     | ❌ 現場に「必要人数」の欄が無い | **責任者（代表）がいない現場** |
//   | 資格不足     | ❌ 現場に「必要な資格」の欄が無い | **その日出る人の、切れた／もうすぐ切れる資格** |
//   | 移動時間     | ❌ 現場に住所が無い（地図APIは有料） | **前後の日で本社↔関東支店をまたぐ人** |
//
// ★「問題があれば通知する」＝問題が無い日は送らない（利用者判断 2026-08-29）。
//   毎日必ず届く通知は読まれなくなる（フェーズ2で実測して学んだこと）。

// ---- 重複判定（画面と1文字も違ってはいけない）--------------------------
// ★index.html / admin.html の PHASE2-CONFLICT-RULE と同じ規則をここへ写している。
//   cf/test/alerts-conflict-parity.test.js が「画面の実装」と「ここ」を
//   同じデータで動かして、結果が完全に一致することを検査する。
//   片方だけ直すと、画面の警告とLINEの通知が食い違う。
export const GENBA_WORKTYPES = ['現場作業', '置局', '着打ち', '撤去品返却'];
export const CONFLICT_SAME_ROSTER = ['グローライズ', 'GRミツマ'];
const SEP = String.fromCharCode(1);

export function rosterKey(company) {
  const c = String(company == null ? '' : company).trim();
  return CONFLICT_SAME_ROSTER.indexOf(c) >= 0 ? 'GR' : c;
}
export function isGenbaWork(n) {
  return GENBA_WORKTYPES.indexOf(String((n && n.workType) || '').trim()) >= 0;
}
export function countsForConflict(n) {
  if (!n || n.isGhost || n.yotei || n.yasumi) return false;
  if (!isGenbaWork(n)) return false;
  return !!(String(n.date || '').trim() && String(n.name || '').trim());
}

// ★2026-08-31 Phase 2: 検知そのものは広く取る（社長指示 §6）。
//   「250件を検知しない仕様にはしないでください。
//     検知を消すのではなく、検知した上で通知レベルを変えること」
export function countsForOverlap(n) {
  if (!n || n.isGhost || n.yotei || n.yasumi) return false;
  return !!(String(n.date || '').trim() && String(n.name || '').trim());
}

function hasWorkTime(j) {
  return !!(j && String(j.start || '').trim() && String(j.end || '').trim());
}

// 重複の強さ（画面と同じ規則。片方だけ直さないこと）
//   high  … 同じ日に別々の「現場作業」。物理的に両立しない
//   check … 現場作業を含むが片方は別業務／時刻が入っていない＝判断材料が足りない
//   info  … 現場作業を含まない（会議＋社内予定など）。運用上両立しうる
// ⚠️「これが本物」と断定しないこと。high は「まず見てほしい順」。
export function conflictSeverity(jobs) {
  const list = jobs || [];
  const siteJobs = list.filter(j => j && j.isSite);
  const sites = {};
  siteJobs.forEach(j => { sites[j.genba + SEP + j.loc] = true; });
  if (siteJobs.length >= 2 && Object.keys(sites).length >= 2) return 'high';
  if (siteJobs.length >= 1) return 'check';
  if (!list.every(hasWorkTime)) return 'check';
  return 'info';
}

export const CONFLICT_SEVERITY_ORDER = { high: 3, check: 2, info: 1 };

function jobKey(n) {
  return String((n && n.genba) || '').trim() + SEP + String((n && n.loc) || '').trim();
}

export function findConflicts(nippos, opts) {
  const from = (opts && opts.from) || '';
  // ★既定は 'high'。既存の呼び出しの出方を変えないため。
  const minSev = (opts && opts.minSeverity) || 'high';
  const minRank = CONFLICT_SEVERITY_ORDER[minSev] || 3;
  const map = new Map();
  (nippos || []).forEach(n => {
    if (!countsForOverlap(n)) return;
    if (from && String(n.date) < from) return;
    const bucket = n.yakin ? 'night' : 'day';
    const key = [String(n.date), rosterKey(n.company), String(n.name).trim(), bucket].join(SEP);
    if (!map.has(key)) map.set(key, new Map());
    const jobs = map.get(key);
    const jk = jobKey(n);
    // ★画面(index.html)の作りと完全に同じ形にする。trim の有無・項目の並びまで合わせること。
    //   cf/test/alerts.test.js が両方を同じデータで動かして JSON を突き合わせる。
    if (!jobs.has(jk)) {
      jobs.set(jk, {
        genba: String(n.genba || ''), loc: String(n.loc || ''),
        butai: String(n.butai || ''), ids: [],
        // ★2026-08-31 Phase 2: 強さを決めるための材料。画面と同じ並び・同じ形。
        workType: String(n.workType || ''),
        isSite: countsForConflict(n),
        start: String(n.start || ''), end: String(n.end || '')
      });
    } else if (countsForConflict(n)) {
        // ★2026-08-31 Codexレビュー#1【P1】: 最初の1行だけで決めていた。
        //   同じ現場に「移動」の行が先にあると、その後の「現場作業」の行を
        //   見ずに「現場作業ではない」と決めつけ、別々の2現場の重なりが
        //   高優先から落ちて画面の警告から消えていた（変更前は出ていた）。
        //   **1行でも現場作業があれば、その現場は現場作業として数える。**
      const _j = jobs.get(jk);
      _j.isSite = true;
      _j.workType = String(n.workType || '');
      if (!_j.start) _j.start = String(n.start || '');
      if (!_j.end) _j.end = String(n.end || '');
    }
    if (n.id) jobs.get(jk).ids.push(n.id);
  });
  const out = [];
  map.forEach((jobs, key) => {
    if (jobs.size < 2) return;
    const p = key.split(SEP);
    const list = Array.from(jobs.values());
    const sev = conflictSeverity(list);
    if ((CONFLICT_SEVERITY_ORDER[sev] || 0) < minRank) return;
    out.push({ date: p[0], roster: p[1], name: p[2], bucket: p[3], severity: sev, jobs: list });
  });
  out.sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1
    : a.name < b.name ? -1 : a.name > b.name ? 1 : 0));
  return out;
}

// ---- 資格の期限（画面と同じ考え方）------------------------------------
const QUAL_SOON_DAYS = 60;
// 資格の期限を知らせる「節目の日」。これ以外の日は出さない（毎朝出るのを防ぐ）。
const QUAL_NOTIFY_DAYS = [60, 30, 14, 7, 3, 1, 0];
// 移動として見る拠点。「両方」は移動ではないので入れない。
const MOVE_KYOTEN = ['本社', '関東支店'];

// ===== 人員不足（実績ベース）=====================================
// ★2026-08-29 利用者判断。依頼書の「人員不足」を、現場マスタに欄を足さずに出す。
//   現場マスタに「必要人数」の列は無い。列を足しても**184現場を手で埋めるまで効かない**。
//   代わりに、その現場の過去の実績から「いつも何人か」を出し、
//   その日だけ大きく少なければ知らせる。**入力ゼロで明日から効く。**
//
//   ⚠️ 限界（正直に書く）: 223現場のうち「いつも」が決められるのは
//      **5日以上の実績がある32現場だけ**。1日で終わる現場は判定しない
//      （実績1日では「いつも」が無く、必ず誤報になる）。
//
//   しきい値は実データ130日で測って決めた（2026-08-29）:
//     50%未満/2人差 → 全期間3日・これから0日（鳴らなすぎ）
//     60%未満/2人差 → 全期間7日・これから1日  ← これを採用
//     60%未満/3人差 → 全期間0日（鳴らなすぎ）
const SHORT_MIN_DAYS = 5;    // これ未満の実績しかない現場は判定しない
const SHORT_RATIO = 0.6;     // いつもの何割を下回ったら
const SHORT_MIN_GAP = 2;     // かつ何人以上少なかったら
// ★Codexレビュー[P2]（2026-08-29）: 全期間の中央値だと現場の工程変化に追従しない
//   （着工期10人・終盤3人の現場で、終盤の3人が正常でも古い10人に引かれて鳴り続ける）。
//   判定日より前の直近この日数だけを見る。値は未来日を除いたバックテストで決めた。
const SHORT_WINDOW_DAYS = 90;
// ===== 人員不足の正式判定（2026-08-31 Phase 2 / 社長指示 §3）===========
// 「第1優先: 案件ごとの必要人数と、実際に配置されている人数を比較する。
//   第2優先: 過去実績との比較。ただし『参考値』であり、断定しない」
//
// ★必要人数が登録されている現場は、正式判定だけを出す。
//   同じ現場を「正式」と「参考」で二重に出さない（どちらを信じるか分からなくなる）。
//
// ⚠️ 正直に書いておく限界（勝手に埋めないこと・社長指示 §0）:
//   ・その日1人も入っていない現場は判定しない。現場マスタに「いつからいつまで
//     動く現場か」の欄が無いので、休みの日と人が付いていない日を区別できない。
//     全部出すと223現場ぶん毎朝鳴って、誰も読まなくなる。
//   ・昼勤と夜勤を分けない。必要人数の欄は現場に1つしかなく、交替制を表せない。
//     分けると「昼2＋夜2で必要4」の現場が、両方とも不足と誤報する。
//   ・必要人数が未登録の現場は、今までどおり実績ベースの参考判定に回る。

// 現場マスタから「必要人数・必要資格」を引く索引。鍵は 元請名＋現場名。
export function siteNeedIndex(jobsites) {
  const out = {};
  (jobsites || []).forEach(j => {
    if (!j) return;
    const k = String(j.genba || '').trim() + SEP + String(j.loc || '').trim();
    const n = Number(j.needCount);
    out[k] = {
      needCount: (j.needCount == null || !isFinite(n) || n <= 0) ? null : Math.floor(n),
      needQuals: Array.isArray(j.needQuals) ? j.needQuals.filter(q => String(q || '').trim() !== '') : [],
      needExp: String(j.needExp || '').trim(),
      status: String(j.status || '').trim()
    };
  });
  return out;
}

// ===== 資格不足の正式判定（2026-08-31 Phase 2 / 社長指示 §4）===========
// 「資格不足 / 期限切れ / 期限間近 / 資格情報が未登録 を別状態として扱う」
//
// ★「誰も持っていない」と「そもそも資格情報が入っていない」を絶対に混ぜないこと。
//   資格マスタに1行でも載っているのは62人中22人（2026-08-31 実測）。
//   混ぜると、資格をまだ入力していないだけの人が「資格不足」と断定される。
//   社長指示 §0「不明な資格・経験を勝手に補完しない」。
export const QUAL_NEED_LABEL = {
  ok: '有効な資格を持つ人がいる',
  soon: '期限が近い',
  expired: '期限が切れている',
  missing: '誰も持っていない',
  unknown: '資格情報が未登録で判定できない'
};

// ★2026-08-31 Codexレビュー#8: 索引の鍵を「会社＋氏名」にする。
//   氏名だけだと、全社で見たとき和信カインドの江頭さんの資格で
//   グローライズの江頭さんが「有資格」になり得た。
//   （/api/alerts は company を省くと '全社' で動く＝既定でこの穴が開いていた）
export function qualPersonKey(who) {
  const name = String((who && who.name) || who || '').trim();
  const co = String((who && who.company) || '').trim();
  return co ? (rosterKey(co) + SEP + name) : name;
}

// 会社＋氏名 -> その人の資格の行。★会社をまたいで混ぜない（同姓が実在する）。
export function qualsByPerson(qualifications, company) {
  const out = {};
  (qualifications || []).forEach(q => {
    if (!q) return;
    const n = String(q.name || '').trim();
    if (!n) return;
    // ★2026-08-31 Codexレビュー#2: 会社が空欄の行まで落としていた。
    //   資格マスタに会社の列が無い／セルが空の行は「どの会社か分からない」であって
    //   「他社の行」ではない。落とすと、その人の期限のお知らせが黙って消える。
    //   （実データでは今0件。だが1行足された瞬間に消えるので、先に塞いでおく）
    const co = String((q && q.company) || '').trim();
    if (co && company && company !== '全社'
        && rosterKey(co) !== rosterKey(company)) return;
    // ★鍵は会社＋氏名（qualPersonKey と同じ作り方でそろえること）
    const key = qualPersonKey({ name: n, company: co });
    (out[key] || (out[key] = [])).push(q);
  });
  return out;
}

// 1つの現場について、必要資格ごとの状態を返す。
//   holders  … 有効な資格を持っている人
//   soon     … 期限が近い人 / expired … 切れている人
//   noRecord … 資格マスタに1行も無い人（この人がいる限り「不足」と断定しない）
export function siteQualCheck(needQuals, names, byPerson, today) {
  const out = [];
  const idx = byPerson || {};          // ★#10: null を渡されても落ちない
  (needQuals || []).forEach(raw => {
    const need = String(raw || '').trim();
    if (!need) return;
    const holders = [], soon = [], expired = [], unreadable = [], noRecord = [];
    (names || []).forEach(who => {
      // who は氏名の文字列でも {name, company} でもよい
      const n = String((who && who.name) || who || '').trim();
      if (!n) return;
      const list = idx[qualPersonKey(who)] || [];
      if (!list.length) { noRecord.push(n); return; }
      list.forEach(x => {
        if (String(x.qual || '').trim() !== need) return;
        const st = qualStatus(x.expires, today);
        // ★2026-08-31 Codexレビュー#7【重大】: 'none'（期限のない資格）を
        //   「読めない」に落としていた。資格284行のうち245行（86%）が期限なし
        //   ＝ゴンドラ特別教育・足場の組立て等・職長など、切れない資格。
        //   ほぼ全部を「判定できない」にしていた＝この機能が働いていなかった。
        if (st === 'none' || st === 'ok') holders.push(n);
        else if (st === 'soon') soon.push(n);   // 期限は近いが、今日は使える
        else if (st === 'expired') expired.push(n);
        else unreadable.push(n);                // 期限欄が日付として読めない
      });
    });
    let status, why = '';
    if (holders.length) status = 'ok';
    else if (soon.length) status = 'soon';      // 今日使える人がいる＝足りてはいる
    // ★2026-08-31 Codexレビュー#5: 期限切れの人が1人いるだけで
    //   「足りていません」と断定していた。資格情報が無い人が同じ現場にいるなら、
    //   その人が持っているかもしれない。**断定できない方を先に見る。**
    else if (noRecord.length) {
      status = 'unknown';
      why = expired.length
        ? '資格情報が未登録の人がいます（期限切れの人もいます）'
        : '資格情報が未登録の人がいます';
    } else if (expired.length) status = 'expired';
    else if (unreadable.length) { status = 'unknown'; why = '有効期限が読めません'; }
    else status = 'missing';               // 全員ぶん登録があり、誰も持っていない
    out.push({
      qual: need, status, why,
      holders: holders, soon: soon, expired: expired, noRecord: noRecord
    });
  });
  return out;
}

export function qualValidYmd(s) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(s || ''));
  if (!m) return false;
  const y = +m[1], mo = +m[2], d = +m[3];
  const dt = new Date(Date.UTC(y, mo - 1, d));
  return dt.getUTCFullYear() === y && dt.getUTCMonth() === mo - 1 && dt.getUTCDate() === d;
}
export function qualStatus(expires, today) {
  const e = String(expires == null ? '' : expires);
  if (e === '') return 'none';
  if (!qualValidYmd(e)) return 'unknown';
  const d = Math.round((Date.parse(e + 'T00:00:00Z') - Date.parse(today + 'T00:00:00Z')) / 86400000);
  if (isNaN(d)) return 'unknown';
  if (d < 0) return 'expired';
  return d <= QUAL_SOON_DAYS ? 'soon' : 'ok';
}

// ---- 日付 --------------------------------------------------------------
export function addDays(ymd, n) {
  const t = Date.parse(String(ymd || '') + 'T00:00:00Z');
  if (isNaN(t)) return '';
  return new Date(t + n * 86400000).toISOString().slice(0, 10);
}

// ---- compact応答をレコードへ ------------------------------------------
// ★画面(parseRows)と同じ意味になるように「夜勤」列を読み分ける。
//   この1列に 夜勤／予定／休み／倉庫 が同居している（昔からの作り）。
export function toRecords(payload) {
  const H = (payload && payload.headers) || [];
  const ix = {};
  H.forEach((h, i) => { ix[String(h).trim()] = i; });
  const get = (r, k) => (ix[k] === undefined ? '' : (r[ix[k]] == null ? '' : r[ix[k]]));
  return ((payload && payload.rows) || []).map(r => {
    // ★Codexレビュー[P2]（2026-08-29）: ここで .trim() してはいけない。
    //   画面(index.html の parseRows)は trim せず完全一致で見ている。
    //   trim すると「休み␣」がWorkerでは休み・画面では通常勤務になり、
    //   同じデータで画面と通知の件数が食い違う。**画面に合わせる。**
    const mode = String(get(r, '夜勤') || '');
    return {
      id: String(get(r, 'ID') || ''),
      date: String(get(r, '作業日') || '').slice(0, 10),
      genba: String(get(r, '元請名') || '').trim(),
      loc: String(get(r, '現場名') || '').trim(),
      name: String(get(r, '氏名') || '').trim(),
      role: String(get(r, '役割') || '').trim(),
      company: String(get(r, '会社') || '').trim(),
      kyoten: String(get(r, '拠点') || '').trim(),
      workType: String(get(r, '作業区分') || '').trim(),
      // ★2026-08-31 Phase 2: 重複の強さを決めるのに時刻が要る。
      //   時刻が入っていない予定は「両立するか判断できない」＝要確認へ落とす。
      //   ここに書かないと、画面は時刻を持っているのにWorkerだけ持たず、
      //   同じデータで判定が食い違う（alerts.test.js が突き合わせている）。
      start: String(get(r, '出勤') || '').trim(),
      end: String(get(r, '退勤') || '').trim(),
      butai: String(get(r, '部隊') || '').trim(),
      yakin: mode === '夜勤',
      yotei: mode === '予定',
      yasumi: mode === '休み',
      souko: mode === '倉庫',
      isGhost: false
    };
  });
}

// ---- 名簿 --------------------------------------------------------------
export function activeRoster(members, company) {
  const inScope = (co) => {
    if (!company || company === '全社') return true;
    const kyoten = CONFLICT_SAME_ROSTER.indexOf(company) >= 0;
    return kyoten ? CONFLICT_SAME_ROSTER.indexOf(String(co || '').trim()) >= 0
      : String(co || '').trim() === company;
  };
  const off = {};
  (members || []).forEach(m => {
    if (m && m.active === false && inScope(m.company)) off[String(m.name || '').trim()] = true;
  });
  const seen = [];
  (members || []).forEach(m => {
    const n = String((m && m.name) || '').trim();
    if (!n || !inScope(m && m.company)) return;
    if (off[n]) return;
    if (seen.indexOf(n) < 0) seen.push(n);
  });
  return seen;
}

// ---- 本体 --------------------------------------------------------------
// date … 見る日（ふつうは「明日」）。today … 実行日（資格の期限の基準）。
// 現場として数えない作業区分。★「休み」を入れているのは、
//   夜勤列が空のまま 作業区分='休み' の行が実在するため（画面の作業区分一覧に「休み」がある）。
//   Codexレビュー[P2]（2026-08-29）の指摘。
const NOT_SITE_WORKTYPE = ['事務所', '倉庫作業', '移動', '休み'];

// その行を「現場の稼働」として数えるか。★当日の集計と過去の集計で必ず同じ条件を使う。
//   ここがズレると「いつも」と「その日」を違う物差しで比べることになり、誤報が出る。
// 現場名そのものが事務所のときも外す。★実データで発見（2026-08-29）:
//   現場名「事務所」の行は464行が 作業区分「その他」で入っており、
//   作業区分だけで判定すると現場として数えてしまう。
//   バックテストで通知15件中5件が「事務所 いつも9人→4人」の誤報だった。
//   事務所に何人居るかは人員不足ではない。**完全一致でだけ外す**
//   （「倉庫材料準備」のような本物の作業を巻き込まないため部分一致にしない）。
const NOT_SITE_LOC = ['事務所'];

function isSiteWork(r) {
  if (!r || !r.loc || !r.name) return false;
  if (r.yotei || r.yasumi) return false;
  if (r.souko) return false;
  if (NOT_SITE_LOC.indexOf(r.loc) >= 0) return false;
  return NOT_SITE_WORKTYPE.indexOf(r.workType) < 0;
}

// 人員不足で使う現場の鍵。★Codexレビュー[P1]（2026-08-29）:
//   元請名+現場名 だけだと、company=全社 のとき別会社の同名現場が合算される
//   （他社の人数で「いつも」が水増しされ、平常なのに不足と誤報する）。
//   **会社（名簿の単位）を鍵に入れる。** グローライズとGRミツマだけは
//   rosterKey が同じ 'GR' になるので、意図どおり1つに統合される。
//   ★昼勤と夜勤も分ける。重複判定は昔から別枠にしているのに、
//   ここだけ合算すると「昼4+夜4=いつも8人」→「昼だけの日は4人＝不足」と誤報する。
function shortKey(r) {
  return rosterKey(r.company) + SEP + r.genba + SEP + r.loc + SEP + (r.yakin ? '夜' : '昼');
}

// 現場ごとの「いつもの人数」＝日ごとの実人数の中央値。
//   平均でなく中央値にするのは、1日だけ大人数を入れた日に引っ張られないため。
export function usualHeadcount(recs, { minDays = SHORT_MIN_DAYS, since = '' } = {}) {
  const byDay = {};                       // 現場鍵+SEP+日 -> Set(氏名)
  recs.forEach(r => {
    if (!isSiteWork(r)) return;
    const k = shortKey(r) + SEP + r.date;
    (byDay[k] || (byDay[k] = new Set())).add(r.name);
  });
  const counts = {};                      // 現場 -> [人数,...]
  Object.keys(byDay).forEach(k => {
    const idx = k.lastIndexOf(SEP);
    const site = k.slice(0, idx), ymd = k.slice(idx + SEP.length);
    if (since && ymd < since) return;     // 窓の外は見ない
    (counts[site] || (counts[site] = [])).push(byDay[k].size);
  });
  const out = {};
  Object.keys(counts).forEach(site => {
    const a = counts[site].slice().sort((x, y) => x - y);
    if (a.length < minDays) return;       // 実績が浅い現場は「いつも」を決めない
    const m = a.length % 2
      ? a[(a.length - 1) / 2]
      : (a[a.length / 2 - 1] + a[a.length / 2]) / 2;
    out[site] = { usual: m, days: a.length };
  });
  return out;
}

export function buildAlerts(payload, { date, today, company } = {}) {
  const recs = toRecords(payload).filter(r => !company || company === '全社'
    || rosterKey(r.company) === rosterKey(company));
  const roster = activeRoster((payload && payload.members) || [], company);
  const rosterSet = new Set(roster);
  const day = recs.filter(r => r.date === date);
  const working = day.filter(r => !r.yotei && !r.yasumi && r.name);

  // ① 人員重複
  const conflicts = findConflicts(recs, { from: date }).filter(c => c.date === date);

  // ② 人員不足 → 責任者（代表）がいない現場
  const sites = {};
  working.forEach(r => {
    if (!r.loc) return;
    // ★Codexレビュー[P1]（2026-08-29）: 倉庫の行を落とし忘れていた。
    //   「倉庫・同行」の1行で 現場1件・責任者なし1件 の誤通知が出た（再現済み）。
    if (r.souko) return;
    if (['事務所', '倉庫作業', '移動'].indexOf(r.workType) >= 0) return;
    const k = r.genba + SEP + r.loc;
    if (!sites[k]) sites[k] = { genba: r.genba, loc: r.loc, people: [], lead: false };
    sites[k].people.push(r.name);
    if (r.role === '代表') sites[k].lead = true;
  });
  const siteList = Object.keys(sites).map(k => sites[k]);
  const noLead = siteList.filter(s => !s.lead);

  // ②-2 人員不足 … その現場のいつもの人数より大きく少ない日
  // ★テストが見つけた不具合（2026-08-29）: 判定する日自身を「いつも」の計算に入れていた。
  //   ①その日が少ないと基準まで一緒に下がって、鳴るべき日が鳴らない
  //   ②実績4日しかない現場が、その日を足して5日になり判定対象になってしまう
  // ★Codexレビュー[P1]（2026-08-29）: さらに **未来の予定まで「いつも」に入っていた**。
  //   毎朝は翌日を判定するので、翌々日以降の予定が基準に混ざる。
  //   実績4日の現場が未来1日で5日になり判定対象にもなる。**判定日より前だけを使う。**
  //   （しきい値もこの修正後に測り直した）
  const past = recs.filter(r => r.date && r.date < date);
  const usual = usualHeadcount(past, { minDays: SHORT_MIN_DAYS, since: addDays(date, -SHORT_WINDOW_DAYS) });
  // 当日側も過去側と**同じ関数**で数える（条件を二度書かない）
  const todayCount = {};
  day.filter(isSiteWork).forEach(r => {
    const k = shortKey(r);
    (todayCount[k] || (todayCount[k] = { genba: r.genba, loc: r.loc, yakin: !!r.yakin, names: new Set() }))
      .names.add(r.name);
  });
  // ★2026-08-31 Phase 2（社長指示 §3）: 必要人数が登録されている現場は
  //   正式判定へ回し、参考判定（実績ベース）からは外す。二重に出さない。
  const needIx = siteNeedIndex((payload && payload.jobsites) || []);
  const hasNeed = k => {
    const t = todayCount[k];
    const nd = t && needIx[t.genba + SEP + t.loc];
    return !!(nd && nd.needCount != null);
  };

  // 正式判定は現場ごと（昼夜を分けない。必要人数の欄が現場に1つしかないため）
  //
  // ⚠️ 2026-08-31 Codexレビュー#3 の限界を正直に書いておく:
  //   **現場マスタには会社の列が無い。** だから「和信カインドの〇〇現場」と
  //   「グローライズの〇〇現場」が同じ名前だと、条件を取り違える。
  //   実データでは、同じ（元請名＋現場名）が2社に出る組は **0件**（2026-08-31 実測）。
  //   現場マスタの鍵の重複も0件。今は起きないが、起きたら誤判定になる。
  //   直すには現場マスタに会社の列を足すしかない（Phase 5 で検討）。
  //   ★ただし人の資格は会社込みで引く（下の {name, company}）ので、
  //     資格の側では他社の同姓の人を使わない。
  const officialSite = {};
  day.filter(isSiteWork).forEach(r => {
    const k = r.genba + SEP + r.loc;
    const nd = needIx[k];
    // ★2026-08-31 Codexレビュー#4: 「必要人数が無ければ何もしない」にしていたので、
    //   必要資格だけ入れた現場が丸ごと素通りしていた。画面では別々に入れられる。
    if (!nd) return;
    if (nd.needCount == null && !(nd.needQuals && nd.needQuals.length)) return;
    const cell = officialSite[k] || (officialSite[k] = {
      genba: r.genba, loc: r.loc, need: nd.needCount, needQuals: nd.needQuals,
      names: new Set(), people: [], seen: {}
    });
    cell.names.add(r.name);
    // ★2026-08-31 Codexレビュー#8: 資格を引くのに会社が要る。
    //   氏名だけだと、全社で見たとき他社の同姓の人の資格で「足りている」になる。
    const pk = rosterKey(r.company) + SEP + r.name;
    if (!cell.seen[pk]) { cell.seen[pk] = 1; cell.people.push({ name: r.name, company: r.company }); }
  });
  const shortOfficial = Object.keys(officialSite).map(k => {
    const s = officialSite[k];
    if (s.need == null) return null;               // 人数は未登録＝人数の判定はしない
    const n = s.names.size;
    if (n >= s.need) return null;
    return { genba: s.genba, loc: s.loc, need: s.need, count: n, gap: s.need - n };
  }).filter(Boolean)
    .sort((a, b) => b.gap - a.gap || (a.loc < b.loc ? -1 : 1));

  // 必要資格の照合（正式判定）。状態は ok/soon/expired/missing/unknown。
  const qualIndex = qualsByPerson((payload && payload.qualifications) || [], company);
  const qualShort = [];
  Object.keys(officialSite).forEach(k => {
    const s = officialSite[k];
    if (!s.needQuals || !s.needQuals.length) return;
    siteQualCheck(s.needQuals, s.people, qualIndex, today || date).forEach(c => {
      if (c.status === 'ok') return;
      qualShort.push({ genba: s.genba, loc: s.loc, ...c });
    });
  });
  // 困る順（誰も持っていない → 切れている → 期限間近 → 判定不可）
  const QS_ORDER = { missing: 0, expired: 1, soon: 2, unknown: 3 };
  qualShort.sort((a, b) => (QS_ORDER[a.status] - QS_ORDER[b.status])
    || (a.loc < b.loc ? -1 : a.loc > b.loc ? 1 : 0));

  const shortStaff = Object.keys(todayCount).map(k => {
    if (hasNeed(k)) return null;                // 正式判定に回した現場は参考から外す
    const u = usual[k];
    if (!u) return null;                        // 実績が浅い現場は判定しない
    const t = todayCount[k];
    const n = t.names.size;
    if (n >= u.usual * SHORT_RATIO) return null;
    if (u.usual - n < SHORT_MIN_GAP) return null;
    return { genba: t.genba, loc: t.loc, yakin: t.yakin, usual: u.usual, count: n, days: u.days };
  }).filter(Boolean)
    // 足りない人数が多い順。★件数を切るとき、どれが落ちるかを決まった順にする
    .sort((a, b) => (b.usual - b.count) - (a.usual - a.count) || (a.loc < b.loc ? -1 : 1));

  // ③ 空き人員
  const busy = new Set(working.filter(r => !r.souko).map(r => r.name));
  const soukoNames = new Set(working.filter(r => r.souko).map(r => r.name));
  const restNames = new Set(day.filter(r => r.yasumi).map(r => r.name));
  const free = roster.filter(n => !busy.has(n) && !soukoNames.has(n) && !restNames.has(n));

  // ⑤ 未確定案件（見積中）／⑥ 延期・中止なのに人が入っている
  const jobsites = (payload && payload.jobsites) || [];
  const unconfirmed = jobsites.filter(j => String((j && j.status) || '').trim() === '見積中');
  const stopped = {};
  jobsites.forEach(j => {
    const st = String((j && j.status) || '').trim();
    if (st === '延期' || st === '中止') stopped[String(j.genba || '').trim() + SEP + String(j.loc || '').trim()] = st;
  });
  const stoppedWithPeople = siteList
    .map(s => ({ ...s, status: stopped[s.genba + SEP + s.loc] }))
    .filter(s => s.status);

  // ★Codexレビュー[P1]（2026-08-29）: 依頼の8項目のうち「未確定案件」「延期案件」が
  //   通知の条件から落ちていた。ただし「見積中◯件」を毎朝出すと、案件が増えるほど
  //   毎日同じ数字が出て読まれなくなる。
  //   → **その日に人が入っているのに見積中のまま**＝明日動くのに受注が決まっていない、
  //     という本当に困る形だけを問題として出す。総数は下のまとめ行に出す。
  const unconfirmedKeys = {};
  unconfirmed.forEach(j => {
    unconfirmedKeys[String(j.genba || '').trim() + SEP + String(j.loc || '').trim()] = 1;
  });
  const unconfirmedWithPeople = siteList.filter(s => unconfirmedKeys[s.genba + SEP + s.loc]);
  const stoppedAll = Object.keys(stopped).length;

  // ⑦ 資格不足 → その日出る人の「**もうすぐ切れる**」資格だけ
  // ★ここを間違えると通知が死ぬ。実データで確かめた話（2026-08-29）:
  //   切れている資格・期限欄が読めない資格を毎朝出すと、真柄さんの
  //   「高所作業認定（2024-05-31 切れ）」が**未来永劫、毎朝出続ける**。
  //   毎日出る警告は誰も読まなくなる（フェーズ2で実測して学んだこと）。
  //   さらに利用者は「その期限切れの資格は一旦ほっといていい（元請さんの物だと思う）」
  //   と判断済み（2026-08-29）。
  //   → 毎朝の通知は「60日以内に切れる」だけ。**新しく起きた事だけを知らせる。**
  //   切れている物・読めない物は、管理画面の「空き確認」に出続けるのでそちらで見る。
  const outNames = new Set(working.map(r => r.name));
  const seenQ = {};
  // ★2026-08-31 会社境界の抜けを塞いだ（社長指示 §0「他社のデータを巻き込まない」）。
  //   氏名だけで引いていたので、和信カインドの同姓の人の資格が
  //   グローライズの通知に混ざり得た。会社で先に絞る。
  // ★2026-08-31 索引の鍵は「会社＋氏名」。その日出る人を会社込みで引く。
  //   （氏名だけで引くと、全社で見たとき他社の同姓の人の資格が混ざる）
  const outPeople = [];
  const seenOut = {};
  working.forEach(r => {
    const pk = qualPersonKey({ name: r.name, company: r.company });
    if (seenOut[pk]) return;
    seenOut[pk] = 1;
    outPeople.push(pk);
  });
  const quals = outPeople
    .reduce((acc, pk) => acc.concat(qualIndex[pk] || []), [])
    .map(q => ({ name: q.name, qual: q.qual, expires: q.expires, status: qualStatus(q.expires, today || date) }))
    // ★Codexレビュー[P2]（2026-08-29）: 「60日以内」だけだと、その人が出る日は
    //   最大60日ぶん毎朝同じ警告が出る。コメントの「新しく起きた事だけ」と食い違う。
    //   節目の日（60/30/14/7/3/1/0日前）にだけ出す。状態を持たずに回数を1/8にできる。
    .filter(q => QUAL_NOTIFY_DAYS.indexOf(
      Math.round((Date.parse(q.expires + 'T00:00:00Z')
                  - Date.parse((today || date) + 'T00:00:00Z')) / 86400000)) >= 0)
    .filter(q => q.status === 'soon')
    .filter(q => {                       // 同じ人の同じ資格が2行あることがある
      const k = q.name + SEP + q.qual;
      if (seenQ[k]) return false;
      seenQ[k] = 1; return true;
    })
    .sort((a, b) => (a.expires < b.expires ? -1 : a.expires > b.expires ? 1 : 0));

  // ⑧ 移動 → 前後の日で本社↔関東支店をまたぐ人
  const byPerson = {};
  recs.forEach(r => {
    if (r.yotei || r.yasumi || !r.name || !r.date) return;
    const k = String(r.kyoten || '').trim() || '本社';
    (byPerson[r.name] = byPerson[r.name] || {});
    (byPerson[r.name][r.date] = byPerson[r.name][r.date] || new Set()).add(k);
  });
  // ★Codexレビュー[P2]（2026-08-29）: 前後どちらも見ると、同じ移動が
  //   「明日の分」と「今日の分」で2朝続けて出る。**その日→翌日だけ**にする。
  //   毎朝 date=明日 で動くので、どの移動もちょうど1回・2日前に知らせることになる。
  //   ★本社↔関東支店に限る（「両方」は移動ではない）。
  const moves = [];
  const seenMove = {};
  const next = addDays(date, 1);
  Object.keys(byPerson).forEach(name => {
    if (!rosterSet.has(name)) return;
    const d = byPerson[name];
    const at = (x) => (d[x] && d[x].size === 1 ? [...d[x]][0] : null);
    const t = at(date);
    if (!t) return;
    const n2 = at(next);
    const real = (a, b) => MOVE_KYOTEN.indexOf(a) >= 0 && MOVE_KYOTEN.indexOf(b) >= 0 && a !== b;
    if (n2 && real(t, n2)) {
      const k = name + SEP + date + SEP + next;
      if (!seenMove[k]) {
        seenMove[k] = 1;
        moves.push({ name, from: date, fromKyoten: t, to: next, toKyoten: n2 });
      }
    }
  });

  return {
    date, today: today || date, company: company || '全社',
    conflicts,
    noLead,
    free, freeCount: free.length, rosterCount: roster.length,
    sites: siteList, siteCount: siteList.length,
    workingCount: new Set(working.map(r => r.name)).size,
    unconfirmed, unconfirmedWithPeople, stoppedWithPeople, stoppedAll, quals, moves,
    shortStaff,
    // ★2026-08-31 Phase 2: 正式判定（社長指示 §3 §4 §9）
    shortOfficial, qualShort
  };
}

// 通知に「問題」として出すものがあるか。翌日の現場・空き人員は"お知らせ"であって問題ではない。
export function hasProblem(a) {
  return !!(a && (a.conflicts.length || a.noLead.length || a.stoppedWithPeople.length
    || a.unconfirmedWithPeople.length || a.quals.length || a.moves.length
    || (a.shortStaff && a.shortStaff.length)
    || (a.shortOfficial && a.shortOfficial.length)
    || (a.qualShort && a.qualShort.length)));
}

const WD = ['日', '月', '火', '水', '木', '金', '土'];
function ymdLabel(ymd) {
  const t = Date.parse(ymd + 'T00:00:00Z');
  if (isNaN(t)) return ymd;
  return ymd.slice(5).replace('-', '/') + '（' + WD[new Date(t).getUTCDay()] + '）';
}

// LINEへ送る文面。★問題が無ければ空文字を返す（＝送らない）。
export function formatAlertsText(a) {
  if (!a || !hasProblem(a)) return '';
  const L = [];
  L.push('【予定管理】' + ymdLabel(a.date) + ' の確認');

  if (a.conflicts.length) {
    L.push('');
    L.push('■ 予定が重なっています ' + a.conflicts.length + '件');
    a.conflicts.slice(0, 10).forEach(c => {
      L.push('・' + c.name + (c.bucket === 'night' ? '（夜勤）' : '') + ' … '
        + c.jobs.map(j => (j.genba ? j.genba + ' ' : '') + j.loc).join(' と '));
    });
    if (a.conflicts.length > 10) L.push('・ほか ' + (a.conflicts.length - 10) + '件');
  }
  // ★2026-08-31 Phase 2（社長指示 §9）: 正式判定を、参考判定より前に・別の見出しで出す。
  //   「必要人数」という決まった数字と比べた結果なので、言い切ってよい。
  if (a.shortOfficial && a.shortOfficial.length) {
    L.push('');
    L.push('■ 必要人数に足りていません ' + a.shortOfficial.length + '件');
    a.shortOfficial.slice(0, 10).forEach(s => L.push('・' + (s.genba ? s.genba + ' ' : '') + s.loc
      + '  必要' + s.need + '人 → ' + s.count + '人（' + s.gap + '人不足）'));
    if (a.shortOfficial.length > 10) L.push('・ほか ' + (a.shortOfficial.length - 10) + '件');
  }
  // ★2026-08-31 Phase 2（社長指示 §9）: 資格を「正式判定」と「判定できない」に分ける。
  //   混ぜると、資格をまだ入力していないだけの人が資格不足に見える。
  //   資格マスタに1行でも載っているのは62人中22人しかいない（実測）。
  const qsOfficial = (a.qualShort || []).filter(q => q.status !== 'unknown');
  const qsUnknown = (a.qualShort || []).filter(q => q.status === 'unknown');
  const qualWho = q => q.status === 'expired' ? '（' + q.expired.join('・') + 'さん）'
    : q.status === 'soon' ? '（' + q.soon.join('・') + 'さん）'
    : q.noRecord.length ? '（' + q.noRecord.join('・') + 'さん）'
    : '';
  if (qsOfficial.length) {
    L.push('');
    L.push('■ 現場に必要な資格が足りていません ' + qsOfficial.length + '件');
    qsOfficial.slice(0, 10).forEach(q => L.push('・' + (q.genba ? q.genba + ' ' : '') + q.loc
      + '  ' + q.qual + ' … ' + QUAL_NEED_LABEL[q.status] + qualWho(q)));
    if (qsOfficial.length > 10) L.push('・ほか ' + (qsOfficial.length - 10) + '件');
  }
  if (qsUnknown.length) {
    L.push('');
    // ★断定しない。足りないのではなく「調べられない」。
    L.push('■ 資格を確かめられませんでした ' + qsUnknown.length + '件（判定できません）');
    qsUnknown.slice(0, 10).forEach(q => L.push('・' + (q.genba ? q.genba + ' ' : '') + q.loc
      + '  ' + q.qual + ' … ' + (q.why || '資格情報が未登録です') + qualWho(q)));
    if (qsUnknown.length > 10) L.push('・ほか ' + (qsUnknown.length - 10) + '件');
  }
  if (a.shortStaff && a.shortStaff.length) {
    L.push('');
    // ★参考判定。断定しない書き方にすること（社長指示 §3）。
    L.push('■ いつもより人が少ない可能性があります ' + a.shortStaff.length + '件（参考）');
    a.shortStaff.slice(0, 10).forEach(s => L.push('・' + (s.genba ? s.genba + ' ' : '') + s.loc
      + (s.yakin ? '（夜勤）' : '') + ' … いつも' + s.usual + '人 → ' + s.count + '人'));
    // ★Codexレビュー[P2]: 11件以上のとき、見出しの件数と本文の行数が合わなくなる
    if (a.shortStaff.length > 10) L.push('・ほか ' + (a.shortStaff.length - 10) + '件');
  }
  if (a.noLead.length) {
    L.push('');
    L.push('■ 責任者がいない現場 ' + a.noLead.length + '件');
    a.noLead.slice(0, 10).forEach(s => L.push('・' + (s.genba ? s.genba + ' ' : '') + s.loc
      + '（' + s.people.length + '人）'));
  }
  if (a.quals.length) {
    L.push('');
    L.push('■ この日に出る人の資格が、まもなく切れます ' + a.quals.length + '件');
    a.quals.slice(0, 10).forEach(q => L.push('・' + q.name + ' ' + q.qual
      + ' … ' + q.expires + ' に切れます'));
  }
  if (a.moves.length) {
    L.push('');
    L.push('■ 拠点をまたぐ移動 ' + a.moves.length + '件');
    a.moves.slice(0, 10).forEach(m => L.push('・' + m.name + ' … '
      + ymdLabel(m.from) + m.fromKyoten + ' → ' + ymdLabel(m.to) + m.toKyoten));
  }
  if (a.unconfirmedWithPeople.length) {
    L.push('');
    L.push('■ 受注が決まっていないのに人が入っています ' + a.unconfirmedWithPeople.length + '件');
    a.unconfirmedWithPeople.slice(0, 10).forEach(s => L.push('・' + (s.genba ? s.genba + ' ' : '')
      + s.loc + '（見積中・' + s.people.length + '人）'));
  }
  if (a.stoppedWithPeople.length) {
    L.push('');
    L.push('■ 延期・中止なのに人が入っています ' + a.stoppedWithPeople.length + '件');
    a.stoppedWithPeople.slice(0, 10).forEach(s => L.push('・' + (s.genba ? s.genba + ' ' : '')
      + s.loc + '（' + s.status + '・' + s.people.length + '人）'));
  }

  L.push('');
  L.push('― この日の予定 ―');
  L.push('現場 ' + a.siteCount + '件 / 出る人 ' + a.workingCount + '人 / 空き ' + a.freeCount + '人');
  if (a.freeCount && a.free.length <= 15) L.push('空き: ' + a.free.join('、'));
  if (a.unconfirmed.length) L.push('見積中の案件 ' + a.unconfirmed.length + '件');
  if (a.stoppedAll) L.push('延期・中止の案件 ' + a.stoppedAll + '件');
  return L.join('\n');
}
