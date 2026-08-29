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
function jobKey(n) {
  return String((n && n.genba) || '').trim() + SEP + String((n && n.loc) || '').trim();
}

export function findConflicts(nippos, opts) {
  const from = (opts && opts.from) || '';
  const map = new Map();
  (nippos || []).forEach(n => {
    if (!countsForConflict(n)) return;
    if (from && String(n.date) < from) return;
    const bucket = n.yakin ? 'night' : 'day';
    const key = [String(n.date), rosterKey(n.company), String(n.name).trim(), bucket].join(SEP);
    if (!map.has(key)) map.set(key, new Map());
    const jobs = map.get(key);
    const jk = jobKey(n);
    // ★画面(index.html)の作りと完全に同じ形にする。trim の有無・項目の並びまで合わせること。
    //   cf/test/alerts.test.js が両方を同じデータで動かして JSON を突き合わせる。
    if (!jobs.has(jk)) jobs.set(jk, {
      genba: String(n.genba || ''), loc: String(n.loc || ''),
      butai: String(n.butai || ''), ids: []
    });
    if (n.id) jobs.get(jk).ids.push(n.id);
  });
  const out = [];
  map.forEach((jobs, key) => {
    if (jobs.size < 2) return;
    const p = key.split(SEP);
    out.push({ date: p[0], roster: p[1], name: p[2], bucket: p[3], jobs: Array.from(jobs.values()) });
  });
  out.sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1
    : a.name < b.name ? -1 : a.name > b.name ? 1 : 0));
  return out;
}

// ---- 資格の期限（画面と同じ考え方）------------------------------------
const QUAL_SOON_DAYS = 60;
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
    const mode = String(get(r, '夜勤') || '').trim();
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
    if (['事務所', '倉庫作業', '移動'].indexOf(r.workType) >= 0) return;
    const k = r.genba + SEP + r.loc;
    if (!sites[k]) sites[k] = { genba: r.genba, loc: r.loc, people: [], lead: false };
    sites[k].people.push(r.name);
    if (r.role === '代表') sites[k].lead = true;
  });
  const siteList = Object.keys(sites).map(k => sites[k]);
  const noLead = siteList.filter(s => !s.lead);

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
  const quals = ((payload && payload.qualifications) || [])
    .filter(q => outNames.has(String((q && q.name) || '').trim()))
    .map(q => ({ name: q.name, qual: q.qual, expires: q.expires, status: qualStatus(q.expires, today || date) }))
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
  const moves = [];
  const seenMove = {};
  const prev = addDays(date, -1), next = addDays(date, 1);
  Object.keys(byPerson).forEach(name => {
    if (!rosterSet.has(name)) return;
    const d = byPerson[name];
    const at = (x) => (d[x] && d[x].size === 1 ? [...d[x]][0] : null);
    const t = at(date);
    if (!t) return;
    const p = at(prev), n2 = at(next);
    const add = (o) => {
      const k = o.name + SEP + o.from + SEP + o.to;
      if (seenMove[k]) return;
      seenMove[k] = 1; moves.push(o);
    };
    if (p && p !== t) add({ name, from: prev, fromKyoten: p, to: date, toKyoten: t });
    if (n2 && n2 !== t) add({ name, from: date, fromKyoten: t, to: next, toKyoten: n2 });
  });

  return {
    date, today: today || date, company: company || '全社',
    conflicts,
    noLead,
    free, freeCount: free.length, rosterCount: roster.length,
    sites: siteList, siteCount: siteList.length,
    workingCount: new Set(working.map(r => r.name)).size,
    unconfirmed, stoppedWithPeople, quals, moves
  };
}

// 通知に「問題」として出すものがあるか。翌日の現場・空き人員は"お知らせ"であって問題ではない。
export function hasProblem(a) {
  return !!(a && (a.conflicts.length || a.noLead.length || a.stoppedWithPeople.length
    || a.quals.length || a.moves.length));
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
  return L.join('\n');
}
