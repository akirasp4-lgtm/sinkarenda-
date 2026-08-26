const SHEET_NAME = '日報データ';
const ARCHIVE_SHEET = 'アーカイブ';
const MEMBER_SHEET = '職人マスタ';
const GENBA_MASTER_SHEET = '元請マスタ';
const JOBSITE_SHEET = '現場マスタ';
const SUMMARY_COMPANY = '会社別集計';
const SUMMARY_MONTH = '月別集計';
const KAKUNIN_SHEET = '月別確認表';
const BILLING_SHEET = '元請別請求集計';
const BILLING_FILTER_SHEET = '元請別請求集計_フィルタ用';
const BILLING_RATE_SHEET = '請求単価マスタ';
const BILLING_CALC_SHEET = '請求計算';
const ALLOCATION_SHEET = '事業部別按分';
const OPLOG_SHEET = '操作ログ';
const HEADERS = ['登録日時','作業日','元請名','現場名','氏名','役割','出勤','退勤','人工','メモ','夜勤','会社','ID','更新者','色','事業部','工番','作業区分','車両','拠点'];
const GROWISE = 'グローライズ';

// ============================================================
// 拠点（本社 / 関東支店）— 2026-08-26 追加
//   依頼書 calendar_request_20260826.md の核心:
//   「会社名で分けるのではなく、本社／関東支店という管理区分を別で持たせる」
//   法人（会社）と拠点は別の軸。GRミツマを関東支店として運用していても、
//   将来GRミツマ法人で本社案件を持つことはありうるため、片方から他方を
//   導出してはいけない。値は必ず行ごとに保存する（読むときに計算し直さない）。
// ============================================================
const KYOTEN_HONSHA = '本社';
const KYOTEN_KANTO  = '関東支店';
const KYOTEN_BOTH   = '両方';           // 本社・関東の両方に関係する予定。1件で両方の画面に出る
const KYOTEN_VALUES = [KYOTEN_HONSHA, KYOTEN_KANTO, KYOTEN_BOTH];

// ★★2026-08-26 利用者指定（最重要）:
//   「関東は、今のカレンダーでいうミツマとグローライズだけの話。
//     ラーテルと和信カインドは混ぜたらあかん」
//   本社／関東支店 はグローライズという組織の中の話。和信カインド・ラーテル・GRHD は
//   別事業なので拠点の軸に入れない（拠点は空欄のままにする）。
const KYOTEN_COMPANIES = [GROWISE, 'GRミツマ'];
function hasKyotenAxis_(company) {
  return KYOTEN_COMPANIES.indexOf(String(company || '').trim()) >= 0;
}

// 会社から拠点の「既定値」を出す。★あくまで初期値を入れるためだけに使う。
// 保存された値を読むときにこれを使ってはいけない（依頼書の要件）。
const KYOTEN_DEFAULT_BY_COMPANY = { 'GRミツマ': KYOTEN_KANTO };
function defaultKyotenForCompany_(company) {
  if (!hasKyotenAxis_(company)) return '';   // 別事業の会社には拠点を入れない
  return KYOTEN_DEFAULT_BY_COMPANY[String(company || '').trim()] || KYOTEN_HONSHA;
}

// 保存する拠点を決める。優先順位:
//   1. 画面から明示的に来た値（利用者がその場で変えた）
//   2. 現場マスタに登録された拠点（現場を選べば自動で入る＝入力を増やさない）
//   3. 会社からの既定値
// ★拠点の軸を持たない会社（和信カインド・ラーテル・GRHD）は常に空欄。
//   画面から値が来ても入れない（取り違えて混ざるのを構造的に防ぐ）。
function resolveKyoten_(explicit, jobsiteKyoten, company) {
  if (!hasKyotenAxis_(company)) return '';
  const e = String(explicit || '').trim();
  if (KYOTEN_VALUES.indexOf(e) >= 0) return e;
  const j = String(jobsiteKyoten || '').trim();
  if (KYOTEN_VALUES.indexOf(j) >= 0) return j;
  return defaultKyotenForCompany_(company);
}

// ==============================================================
// 社長専用カレンダー（極秘）
// シート名は意図的に内部呼称のみ。PIN認証でのみアクセス可能。
// ==============================================================
const PRES_SHEET = '社長予定';
const PRES_HEADERS = ['登録日時','タイトル','開始日','開始時刻','終了日','終了時刻','場所','メモ','カテゴリ','色','ID','更新者'];
const PRES_PIN = '1203';
const PRES_DELETE_MARKER = '__PRES_DELETED__';

// ==============================================================
// 車両予約シート（LINEボット連携用 - GR社内秘書ボットから書き込み）
// 既存カレンダー機能とは独立。トークン認証でのみアクセス可能。
// ==============================================================
const VEHICLE_RES_SHEET = '車両予約';
const VEHICLE_RES_HEADERS = [
  '予約ID','車両名','ナンバー','所有会社','使用者氏名','使用者LINE_ID',
  '開始日時','返却予定日時','実返却日時','行先','状態','備考','登録日時','更新日時'
];
const VEHICLE_RES_TOKEN = '車両予約用トークン1234';

// ==============================================================
// 全体認証（合言葉 k）— 2026-06-10 セキュリティ強化
// 読み書きすべてに合言葉を要求する仕組み。スクリプトプロパティで制御:
//   CAL_TOKEN         = 合言葉の値（コードには書かない＝このファイルは公開リポに上がるため）
//   CAL_REQUIRE_TOKEN = '1' にすると照合が有効になる（未設定の間は従来どおり素通し＝移行期間）
// ロールバック: CAL_REQUIRE_TOKEN プロパティを削除すれば即・従来動作に戻る。
// ==============================================================
function calAuthOk_(provided) {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('CAL_REQUIRE_TOKEN') !== '1') return true;  // 移行期間: 照合OFF
  var t = props.getProperty('CAL_TOKEN') || '';
  if (!t) return false;  // 照合ONなのに合言葉未設定 = 全拒否（fail-closed）
  return String(provided || '') === t;
}
function authError_() {
  return ContentService.createTextOutput(JSON.stringify({
    status: 'error', code: 'auth', message: '認証に失敗しました。合言葉を確認してください。'
  })).setMimeType(ContentService.MimeType.JSON);
}

// ==============================================================
// 日付整形の高速版（2026-08-21 レスポンス改善）
// Utilities.formatDate は1回あたりが重く、日報2,600行×3箇所＝約8,000回
// 呼ぶと数秒単位で効いてくる。スクリプトのタイムゾーンとJSのローカル時刻が
// 一致している場合だけ自前整形に切り替え、ズレていれば従来どおりに戻す。
// ==============================================================
var _TZ_FAST_OK_ = null;
function tzFastOk_() {
  if (_TZ_FAST_OK_ !== null) return _TZ_FAST_OK_;
  try {
    var probe = new Date(2026, 0, 2, 3, 4, 5);
    _TZ_FAST_OK_ = (Utilities.formatDate(probe, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss') === '2026-01-02 03:04:05');
  } catch (err) {
    _TZ_FAST_OK_ = false;
  }
  return _TZ_FAST_OK_;
}
function _p2_(n) { return (n < 10 ? '0' : '') + n; }
function fmtDate_(v, tz) {
  if (!(v instanceof Date)) return String(v || '');
  if (!tzFastOk_()) return Utilities.formatDate(v, tz || Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return v.getFullYear() + '-' + _p2_(v.getMonth() + 1) + '-' + _p2_(v.getDate());
}
function fmtTime_(v, tz) {
  if (!(v instanceof Date)) return String(v || '');
  if (!tzFastOk_()) return Utilities.formatDate(v, tz || Session.getScriptTimeZone(), 'HH:mm');
  return _p2_(v.getHours()) + ':' + _p2_(v.getMinutes());
}
function fmtDateTime_(v, tz) {
  if (!(v instanceof Date)) return String(v || '');
  if (!tzFastOk_()) return Utilities.formatDate(v, tz || Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  return fmtDate_(v, tz) + ' ' + _p2_(v.getHours()) + ':' + _p2_(v.getMinutes()) + ':' + _p2_(v.getSeconds());
}

// ==============================================================
// 読み(フリガナ)自動生成用 - Groq API
// スクリプトプロパティ GROQ_API_KEY が設定されていれば有効
// ==============================================================
const GROQ_MODEL = 'meta-llama/llama-4-scout-17b-16e-instruct';

function needsYomi_(text) {
  return typeof text === 'string' && /[\u3400-\u9FFF]/.test(text);
}

// Groq に一括で読みを問い合わせる。失敗時は空配列を返す。
function fetchYomiFromGroq_(texts) {
  if (!texts || !texts.length) return [];
  const key = PropertiesService.getScriptProperties().getProperty('GROQ_API_KEY');
  if (!key) return [];
  const numbered = texts.map((t, i) => `${i + 1}. ${t}`).join('\n');
  const prompt = '次の日本語名称をそれぞれひらがなの読み(フリガナ)に変換してください。\n'
               + '- 人名・地名・建物名・店名・会社名を想定\n'
               + '- 必ず「ひらがなのみ」で出力(長音符「ー」は使用可)\n'
               + '- 元の文字列の順番を保持\n'
               + '- JSON配列のみで回答(説明不要)\n\n'
               + 'テキスト:\n' + numbered + '\n\n'
               + '出力形式例: ["やまだてい","ひがしおおさかびる",...]';
  try {
    const res = UrlFetchApp.fetch('https://api.groq.com/openai/v1/chat/completions', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + key },
      payload: JSON.stringify({
        model: GROQ_MODEL,
        messages: [{ role: 'user', content: prompt }],
        temperature: 0,
      }),
      muteHttpExceptions: true,
    });
    if (res.getResponseCode() !== 200) return [];
    const data = JSON.parse(res.getContentText());
    const content = (data.choices && data.choices[0] && data.choices[0].message && data.choices[0].message.content) || '';
    const m = content.match(/\[[\s\S]*\]/);
    if (!m) return [];
    const arr = JSON.parse(m[0]);
    return Array.isArray(arr) ? arr : [];
  } catch (e) {
    return [];
  }
}

// 1件の読みを生成(新規追加時に使用)。失敗/不要時は空文字。
function generateYomiSafe_(text) {
  if (!needsYomi_(text)) return '';
  const arr = fetchYomiFromGroq_([text]);
  return String(arr[0] || '').trim();
}

function ensureHeaders_(sheet) {
  ensureColumns_(sheet, HEADERS.length);
  const data = sheet.getDataRange().getValues();
  const currentHeaders = data[0] || [];
  HEADERS.forEach((h, i) => {
    if (String(currentHeaders[i] || '').trim() !== h) sheet.getRange(1, i + 1).setValue(h);
  });
}

function getIdCol_() { return HEADERS.indexOf('ID'); }

function ensureColumns_(sheet, needed) {
  if (sheet.getMaxColumns() < needed) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), needed - sheet.getMaxColumns());
  }
}

function isPresidentAction_(action) {
  return action === 'pres_list'
      || action === 'pres_add'
      || action === 'pres_update'
      || action === 'pres_delete';
}

function serializePresidentRows_(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const tz = Session.getScriptTimeZone();
  const idCol = PRES_HEADERS.indexOf('ID');
  const categoryCol = PRES_HEADERS.indexOf('カテゴリ');
  const rows = data.slice(1);
  const deletedIds = new Set();
  rows.forEach(r => {
    const id = String(r[idCol] || '').trim();
    if (id && String(r[categoryCol] || '') === PRES_DELETE_MARKER) deletedIds.add(id);
  });
  const latestById = new Map();
  rows.forEach(r => {
    const id = String(r[idCol] || '').trim();
    if (!id || deletedIds.has(id)) return;
    latestById.set(id, r);
  });
  return Array.from(latestById.values()).map(r => {
    const obj = {};
    PRES_HEADERS.forEach((h, j) => {
      const v = r[j];
      if (h === '開始日' || h === '終了日') {
        obj[h] = (v instanceof Date) ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : String(v || '');
      } else if (h === '開始時刻' || h === '終了時刻') {
        obj[h] = (v instanceof Date) ? Utilities.formatDate(v, tz, 'HH:mm') : String(v || '');
      } else {
        obj[h] = (v === undefined || v === null) ? '' : v;
      }
    });
    return obj;
  });
}

function handlePresidentAction_(body, action, updatedBy) {
  if (String(body.pin || '') !== PRES_PIN) {
    return error('認証に失敗しました');
  }

  // 一覧取得は読み取り専用。日報処理の長時間ロックとは独立させる。
  if (action === 'pres_list') {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const presSheet = ss.getSheetByName(PRES_SHEET);
      return ok({rows: presSheet ? serializePresidentRows_(presSheet) : []});
    } catch (err) {
      return error(err.toString());
    }
  }

  // 同一Googleユーザーの書き込みを直列化し、日報・集計とは待ち合わせない。
  const lock = LockService.getUserLock();
  if (!lock.tryLock(10000)) {
    return error('現在他の人が更新中です。数秒待ってから再度お試しください。');
  }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const presSheet = getOrCreatePresSheet_(ss);

    if (action === 'pres_add') {
      const ev = body.event || {};
      // ── 2026-08-26: 画面側が作ったIDを受け付ける（社員用の add と同じ形にする）──
      // なぜ必要か: 楽観的保存（送信を待たずに画面へ出す）では、送り直しの前に
      // 「もう入っていないか」をIDで確認する。サーバーが毎回新しいIDを振ると
      // 画面側はそのIDを知りようがなく、確認できないまま送り直して予定が二重になる。
      // 受け入れ条件を英数字のみに絞る（式・区切り文字・全角の混入を防ぐ）。
      // 条件に合わなければ従来どおりサーバーで採番＝古い画面から呼ばれても壊れない。
      const rawId = String(ev.id || '').trim();
      const id = /^P[0-9a-zA-Z]{8,64}$/.test(rawId)
        ? rawId
        : ('P' + Utilities.getUuid().replace(/-/g, ''));
      // ── 二段構え: サーバー側でも二重登録を止める ──────────────────
      // 画面側の「もう入っていないか」確認をすり抜けた再送（確認の通信自体が
      // 失敗した場合など）が来ても、ここで必ず止まる。既にあるなら足さずに
      // 成功を返す＝画面側は送信済みとして扱えるので未送信が残り続けない。
      const presData = presSheet.getDataRange().getValues();
      const presIdCol = PRES_HEADERS.indexOf('ID');
      for (let i = 1; i < presData.length; i++) {
        if (String(presData[i][presIdCol] || '').trim() === id) {
          return ok({id: id, duplicate: true});
        }
      }
      presSheet.appendRow([
        new Date(),
        String(ev.title || ''),
        String(ev.startDate || ''),
        String(ev.startTime || ''),
        String(ev.endDate || ev.startDate || ''),
        String(ev.endTime || ''),
        String(ev.location || ''),
        String(ev.memo || ''),
        String(ev.category || ''),
        String(ev.color || '#1D9E75'),
        id,
        updatedBy
      ]);
      return ok({id});
    }

    if (action === 'pres_update') {
      const ev = body.event || {};
      const id = String(ev.id || '');
      if (!id) return error('IDが指定されていません');
      const data = presSheet.getDataRange().getValues();
      const idCol = PRES_HEADERS.indexOf('ID');
      const categoryCol = PRES_HEADERS.indexOf('カテゴリ');
      let found = false;
      let registeredAt = null;
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][idCol]) === id) {
          if (String(data[i][categoryCol] || '') === PRES_DELETE_MARKER) {
            return error('対象が見つかりませんでした');
          }
          if (!found) registeredAt = data[i][0] || new Date();
          found = true;
        }
      }
      if (!found) return error('対象が見つかりませんでした');
      // 更新を追記履歴にすることで、異なるユーザーの同時更新でも別行を上書きしない。
      presSheet.appendRow([
        registeredAt,
        String(ev.title || ''),
        String(ev.startDate || ''),
        String(ev.startTime || ''),
        String(ev.endDate || ev.startDate || ''),
        String(ev.endTime || ''),
        String(ev.location || ''),
        String(ev.memo || ''),
        String(ev.category || ''),
        String(ev.color || '#1D9E75'),
        id,
        updatedBy
      ]);
      return ok({updated: id});
    }

    if (action === 'pres_delete') {
      const id = String(body.id || '');
      if (!id) return error('IDが指定されていません');
      const data = presSheet.getDataRange().getValues();
      const idCol = PRES_HEADERS.indexOf('ID');
      const categoryCol = PRES_HEADERS.indexOf('カテゴリ');
      let found = false;
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][idCol]) === id) {
          if (String(data[i][categoryCol] || '') === PRES_DELETE_MARKER) {
            return error('対象が見つかりませんでした');
          }
          found = true;
        }
      }
      if (!found) return error('対象が見つかりませんでした');
      // 削除印は同じIDへの遅延更新より常に優先し、削除済み予定の復活を防ぐ。
      presSheet.appendRow([
        new Date(), '', '', '', '', '', '', '', PRES_DELETE_MARKER, '', id, updatedBy
      ]);
      return ok({deleted: id});
    }

    return error('未対応のアクションです');
  } catch (err) {
    return error(err.toString());
  } finally {
    lock.releaseLock();
  }
}

function getDailyDataLock_() {
  // Webアプリでは DocumentLock は null、UserLock は利用者ごとになる。
  // ScriptLock だけが実行ユーザーをまたいで共有されるため、日報データの
  // 短い読書き専用に使う。長時間の集計・帳票処理はこのロックを保持しない。
  return LockService.getScriptLock();
}

function isEmployeeScheduleMutation_(action) {
  return action === 'add' || action === 'update' || action === 'delete';
}

function isAdminDailyMutation_(action) {
  return action === 'archive'
      || action === 'cleanup_orphan_jobnos'
      || action === 'merge_genba'
      || action === 'merge_loc'
      || action === 'reassign_jobno';
}

// 読み取り専用。集計・帳票の管理ロックを待たず、予定更新とだけ直列化する。
function handleGetSheet_(body) {
  const dataLock = getDailyDataLock_();
  if (!dataLock.tryLock(10000)) {
    return error('現在予定を更新中です。数秒待ってから再度お試しください。');
  }
  try {
  const sheetName = body.sheet || '';
  const allowed = [SHEET_NAME, ARCHIVE_SHEET, MEMBER_SHEET, GENBA_MASTER_SHEET, JOBSITE_SHEET, SUMMARY_COMPANY, SUMMARY_MONTH, KAKUNIN_SHEET, BILLING_SHEET, BILLING_FILTER_SHEET, ALLOCATION_SHEET, OPLOG_SHEET];
  if (!allowed.includes(sheetName)) return error('無効なシート名です');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = ss.getSheetByName(sheetName);
  if (!targetSheet) return error('シートが見つかりません: ' + sheetName);
  const data = targetSheet.getDataRange().getValues();
  const tz = Session.getScriptTimeZone();
  // 期間フィルタ（任意）: 日報データ・アーカイブのみ作業日列で絞り込む
  // dateFrom/dateTo は 'YYYY-MM-DD' 形式の文字列、両端含む
  const dateFrom = String(body.dateFrom || '').trim();
  const dateTo = String(body.dateTo || '').trim();
  let filtered = data;
  if ((dateFrom || dateTo) && (sheetName === SHEET_NAME || sheetName === ARCHIVE_SHEET) && data.length > 1) {
    const headers = data[0];
    const dateColIdx = headers.indexOf('作業日');
    if (dateColIdx >= 0) {
      const head = [data[0]];
      const bodyRows = data.slice(1).filter(row => {
        const v = row[dateColIdx];
        const d = v instanceof Date
          ? Utilities.formatDate(v, tz, 'yyyy-MM-dd')
          : String(v || '').slice(0, 10);
        if (dateFrom && d < dateFrom) return false;
        if (dateTo && d > dateTo) return false;
        return true;
      });
      filtered = head.concat(bodyRows);
    }
  }
  const formatted = filtered.map(row => row.map(v => {
    if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd HH:mm:ss');
    return v;
  }));
    return ok({sheetName, data: formatted});
  } catch (err) {
    return error(err.toString());
  } finally {
    dataLock.releaseLock();
  }
}

function requireDailyRows_(body) {
  if (!body || !Array.isArray(body.rows) || body.rows.length === 0) {
    throw new Error('登録する予定データがありません');
  }
  return body.rows;
}

// 現場マスタの「現場名→拠点」を1回だけ読んで辞書にする（行ごとにシートを読まない）。
function getJobsiteKyotenMap_(ss) {
  const map = {};
  try {
    const sheet = getOrCreateJobSiteSheet_(ss);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const loc = String(data[i][1] || '').trim();      // 現場名
      const kyoten = String(data[i][10] || '').trim();  // 拠点（11列目）
      if (loc && kyoten) map[loc] = kyoten;
    }
  } catch (e) { /* 読めなくても既定値で動く。登録を止めない */ }
  return map;
}

function buildDailyValues_(ss, rows, updatedBy) {
  const jobNoCache = {};
  const kyotenMap = getJobsiteKyotenMap_(ss);
  let leaderDivision = null;
  const leaderRow = rows.find(r => r.role === '代表');
  const leaderName = leaderRow ? leaderRow.name : '';
  return rows.map(row => {
    let division = '';
    let jobNo = '';
    // 工番発行は「グローライズ × 倉庫/休み/予定 のいずれでもない」場合のみ
    if (row.company === GROWISE && !row.souko && !row.yotei && !row.yasumi && row.workType === '現場作業') {
      const explicitDiv = String(row.jobNoDivision || '').trim();
      if (explicitDiv) {
        division = explicitDiv;
      } else {
        if (leaderDivision === null) leaderDivision = getMemberDivision_(ss, leaderName);
        division = leaderDivision;
      }
      if (row.genba && row.loc) {
        const cacheKey = row.genba + '|||' + row.loc;
        if (!jobNoCache[cacheKey]) {
          jobNoCache[cacheKey] = getOrGenerateJobNo_(ss, row.genba, row.loc, division);
        }
        jobNo = jobNoCache[cacheKey];
      }
    }
    return [
      new Date().toLocaleString('ja-JP'),
      row.date, row.genba, row.loc, row.name, row.role,
      String(row.start || ''), String(row.end || ''),
      Number(row.kosu), row.memo,
      row.souko ? '倉庫' : row.yotei ? '予定' : row.yasumi ? '休み' : row.yakin ? '夜勤' : '',
      row.company || '',
      row.id || '',
      row.updatedBy || updatedBy || '',
      row.color || '',
      division,
      jobNo,
      row.workType || '',
      row.vehicle || '',
      // ★2026-08-26 拠点。画面が明示した値 > 現場マスタの拠点 > 会社からの既定値。
      //   決まった値をここで必ず保存する（読むときに会社から計算し直さない）。
      resolveKyoten_(row.kyoten, kyotenMap[String(row.loc || '').trim()], row.company)
    ];
  });
}

function appendDailyValues_(sheet, values) {
  if (!values.length) return;
  sheet.getRange(sheet.getLastRow() + 1, 1, values.length, HEADERS.length).setValues(values);
}

function doPost(e) {
  let body;
  let action;
  let updatedBy;
  try {
    body = JSON.parse(e.postData.contents);
    if (!calAuthOk_(body.k)) return authError_();
    action = body.action || 'add';
    updatedBy = String(body.updatedBy || '');
    if (isPresidentAction_(action)) {
      return handlePresidentAction_(body, action, updatedBy);
    }
    if (action === 'get_sheet') {
      return handleGetSheet_(body);
    }
  } catch (err) {
    return error(err.toString());
  }

  const employeeMutation = isEmployeeScheduleMutation_(action);
  // 社員の保存は全利用者共通の日報ロックだけを使う。管理処理は別の
  // UserLock で直列化し、集計・帳票の長時間処理から社員保存を分離する。
  const lock = employeeMutation ? getDailyDataLock_() : LockService.getUserLock();
  if (!lock.tryLock(10000)) {
    return error(employeeMutation
      ? '現在予定を更新中です。数秒待ってから再度お試しください。'
      : '現在他の人が更新中です。数秒待ってから再度お試しください。');
  }
  let dailyDataLock = null;
  try {
    // アーカイブ・マージ等は管理ロックに加え、LINE読取と社員保存も止める。
    if (!employeeMutation && isAdminDailyMutation_(action)) {
      dailyDataLock = getDailyDataLock_();
      if (!dailyDataLock.tryLock(10000)) {
        return error('現在予定を更新中です。数秒待ってから再度お試しください。');
      }
    }
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);

    ensureHeaders_(sheet);
    const idCol = getIdCol_();

    // 照合のON/OFF切替（正しい合言葉を知っている人だけが叩ける管理アクション）
    // body: { action:'cal_set_enforce', k:<合言葉>, on:'1'(ON) | '0'(OFF) }
    if (action === 'cal_set_enforce') {
      var calProps = PropertiesService.getScriptProperties();
      var calT = calProps.getProperty('CAL_TOKEN') || '';
      if (!calT || String(body.k || '') !== calT) return authError_();
      calProps.setProperty('CAL_REQUIRE_TOKEN', String(body.on) === '1' ? '1' : '0');
      return ok({enforce: calProps.getProperty('CAL_REQUIRE_TOKEN') === '1'});
    }

    if (action === 'add') {
      const rows = requireDailyRows_(body);
      const values = buildDailyValues_(ss, rows, updatedBy);
      appendDailyValues_(sheet, values);
      logOperation_(ss, 'add', rows[0].genba + '/' + (rows[0].loc || ''), '行数=' + rows.length, updatedBy);
      return ok({count: rows.length});
    }

    if (action === 'delete') {
      const ids = body.ids || [];
      if (ids.length === 0) return ok({deleted: 0});
      const data = sheet.getDataRange().getValues();
      const rowsToDelete = [];
      for (let i = data.length - 1; i >= 1; i--) {
        const rowId = String(data[i][idCol] || '').trim();
        if (rowId && ids.includes(rowId)) rowsToDelete.push(i + 1);
      }
      rowsToDelete.forEach(rowNum => sheet.deleteRow(rowNum));
      logOperation_(ss, 'delete', 'IDs=' + ids.length + '件', '削除行=' + rowsToDelete.length, updatedBy);
      return ok({deleted: rowsToDelete.length, requested: ids.length});
    }

    if (action === 'update') {
      const rows = requireDailyRows_(body);
      const ids = body.ids || [];
      const rowsToDelete = [];
      if (ids.length > 0) {
        const data = sheet.getDataRange().getValues();
        for (let i = data.length - 1; i >= 1; i--) {
          const rowId = String(data[i][idCol] || '').trim();
          if (rowId && ids.includes(rowId)) rowsToDelete.push(i + 1);
        }
      }
      const values = buildDailyValues_(ss, rows, updatedBy);
      // 新しい予定を先に一括保存する。保存に失敗しても元予定は残る。
      appendDailyValues_(sheet, values);
      rowsToDelete.forEach(rowNum => sheet.deleteRow(rowNum));
      logOperation_(ss, 'update', rows[0].genba + '/' + (rows[0].loc || ''), '行数=' + rows.length + ', 旧ID=' + ids.length, updatedBy);
      return ok({updated: rows.length});
    }

    if (action === 'archive') {
      const months = body.months || 3;
      const archived = archiveOldData_(ss, months);
      logOperation_(ss, 'archive', months + 'ヶ月以前', '件数=' + archived, updatedBy);
      return ok({archived});
    }

    if (action === 'cleanup_orphan_jobnos') {
      const cleaned = cleanupOrphanJobNos_(ss);
      logOperation_(ss, 'cleanup_orphan_jobnos', '休み/倉庫/予定', '清掃=' + cleaned, updatedBy);
      return ok({cleaned: cleaned});
    }

    if (action === 'merge_genba') {
      const from = String(body.from || '').trim();
      const to = String(body.to || '').trim();
      if (!from || !to) return error('from と to を指定してください');
      if (from === to) return error('同じ名前です');
      const result = mergeGenba_(ss, from, to);
      logOperation_(ss, 'merge_genba', from + ' → ' + to, JSON.stringify(result), updatedBy);
      return ok(result);
    }

    // 現場マスタから 1 行削除する。
    // 日報・アーカイブのいずれかに参照があれば削除拒否（マージか完了フラグを促す）。
    // body: { genba, loc, force (任意・売上があっても削除する場合 true), updatedBy }
    if (action === 'delete_site') {
      const genba = String(body.genba || '').trim();
      const loc = String(body.loc || '').trim();
      const force = !!body.force;
      if (!genba) return error('元請名は必須です');
      // 日報・アーカイブの参照チェック
      let nippoRefs = 0;
      let archiveRefs = 0;
      [SHEET_NAME, ARCHIVE_SHEET].forEach((name, idx) => {
        const sh = ss.getSheetByName(name);
        if (!sh) return;
        const d = sh.getDataRange().getValues();
        if (d.length <= 1) return;
        const headers = d[0];
        const gCol = headers.indexOf('元請名');
        const lCol = headers.indexOf('現場名');
        if (gCol < 0 || lCol < 0) return;
        let count = 0;
        for (let i = 1; i < d.length; i++) {
          if (String(d[i][gCol] || '').trim() === genba && String(d[i][lCol] || '').trim() === loc) count++;
        }
        if (idx === 0) nippoRefs = count; else archiveRefs = count;
      });
      if (nippoRefs > 0 || archiveRefs > 0) {
        return error('日報またはアーカイブに参照があるため削除できません（日報:' + nippoRefs + '件 / アーカイブ:' + archiveRefs + '件）。マージ機能か完了フラグを使ってください。');
      }
      // 現場マスタから該当行を削除
      const jobSite = getOrCreateJobSiteSheet_(ss);
      const data = jobSite.getDataRange().getValues();
      let revenueOnRow = 0;
      let targetRow = -1;
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0] || '').trim() === genba && String(data[i][1] || '').trim() === loc) {
          revenueOnRow = Number(data[i][6] || 0);
          targetRow = i;
          break;
        }
      }
      if (targetRow < 0) return error('現場マスタに該当現場が見つかりません');
      if (revenueOnRow > 0 && !force) {
        return error('この現場には売上 ' + revenueOnRow + ' 円 が入力されています。本当に削除する場合は再度実行してください（クライアントで force フラグ）。');
      }
      const oldJobNo = String(data[targetRow][2] || '');
      jobSite.deleteRow(targetRow + 1);
      logOperation_(ss, 'delete_site', genba + '/' + loc, '工番=' + oldJobNo + ' 売上=' + revenueOnRow, updatedBy);
      return ok({deleted: genba + '/' + loc, oldJobNo: oldJobNo, revenue: revenueOnRow});
    }

    // 同じ元請内で「現場名」を統合する。日報・アーカイブ・現場マスタを書き換え。
    // 統合先に工番がある場合は日報側の工番もそちらに統一する。
    // body: { genba, fromLoc, toLoc, updatedBy }
    if (action === 'merge_loc') {
      const genba = String(body.genba || '').trim();
      const fromLoc = String(body.fromLoc || '').trim();
      const toLoc = String(body.toLoc || '').trim();
      if (!genba || !fromLoc || !toLoc) return error('元請、統合元、統合先の現場名は必須です');
      if (fromLoc === toLoc) return error('同じ現場名です');
      const result = mergeLoc_(ss, genba, fromLoc, toLoc);
      logOperation_(ss, 'merge_loc', genba + '/' + fromLoc + ' → ' + toLoc, JSON.stringify(result), updatedBy);
      return ok(result);
    }

    if (action === 'summarize') {
      generateSummary_();
      return ok({message: '集計を更新しました'});
    }

    // Phase 2: 期間指定の月別確認表風データを返す（シートには書かず、直接 CSV 用 2D 配列を返す）
    // body: { dateFrom, dateTo, company (任意、未指定なら全社) }
    // 返却: { rows: [タイトル行, ヘッダ行, データ行×n, 合計行] } と、columns（日付列のラベル）
    if (action === 'period_kakunin') {
      const dateFrom = String(body.dateFrom || '').trim();
      const dateTo = String(body.dateTo || '').trim();
      if (!dateFrom || !dateTo) return error('開始日と終了日を指定してください');
      if (dateFrom > dateTo) return error('開始日が終了日より後です');
      const companyFilter = String(body.company || '').trim();
      const result = generatePeriodKakuninData_(ss, dateFrom, dateTo, companyFilter);
      return ok(result);
    }

    // 月別確認表シートを xlsx 形式（色・罫線・書式そのまま）でエクスポートして base64 で返す
    if (action === 'export_kakunin_xlsx') {
      const kSheet = ss.getSheetByName(KAKUNIN_SHEET);
      if (!kSheet) return error('月別確認表シートが見つかりません。先に集計を更新してください');
      const result = exportSheetAsXlsxBase64_(ss, kSheet);
      return ok({base64: result.base64, filename: '月別確認表.xlsx'});
    }

    // 任意のシートを xlsx でエクスポート（CSV ダウンロードの置き換え）
    // 日報データ／アーカイブで dateFrom/dateTo 指定があれば一時シートを作って絞り込んでからエクスポート
    if (action === 'export_sheet_xlsx') {
      const sheetName = body.sheet || '';
      const allowed = [SHEET_NAME, ARCHIVE_SHEET, MEMBER_SHEET, GENBA_MASTER_SHEET, JOBSITE_SHEET, SUMMARY_COMPANY, SUMMARY_MONTH, KAKUNIN_SHEET, BILLING_SHEET, BILLING_FILTER_SHEET, ALLOCATION_SHEET, OPLOG_SHEET];
      if (!allowed.includes(sheetName)) return error('無効なシート名です');
      const targetSheet = ss.getSheetByName(sheetName);
      if (!targetSheet) return error('シートが見つかりません: ' + sheetName);

      const dateFrom = String(body.dateFrom || '').trim();
      const dateTo = String(body.dateTo || '').trim();
      const canFilter = (sheetName === SHEET_NAME || sheetName === ARCHIVE_SHEET) && (dateFrom || dateTo);

      let tempSheet = null;
      let exportTarget = targetSheet;
      try {
        if (canFilter) {
          const data = targetSheet.getDataRange().getValues();
          if (data.length > 1) {
            const headers = data[0];
            const dateColIdx = headers.indexOf('作業日');
            if (dateColIdx >= 0) {
              const tz = Session.getScriptTimeZone();
              const filtered = [data[0]].concat(data.slice(1).filter(row => {
                const v = row[dateColIdx];
                const d = v instanceof Date
                  ? Utilities.formatDate(v, tz, 'yyyy-MM-dd')
                  : String(v || '').slice(0, 10);
                if (dateFrom && d < dateFrom) return false;
                if (dateTo && d > dateTo) return false;
                return true;
              }));
              tempSheet = ss.insertSheet('_TMP' + sheetName + '_' + (new Date().getTime()));
              tempSheet.getRange(1, 1, filtered.length, filtered[0].length).setValues(filtered);
              SpreadsheetApp.flush();
              exportTarget = tempSheet;
            }
          }
        }
        const result = exportSheetAsXlsxBase64_(ss, exportTarget);
        const suffix = (dateFrom || dateTo) ? '_' + (dateFrom || '始') + '_' + (dateTo || '今') : '';
        return ok({base64: result.base64, filename: sheetName + suffix + '.xlsx'});
      } finally {
        if (tempSheet) { try { ss.deleteSheet(tempSheet); } catch (e) {} }
      }
    }

    // 期間指定の月別確認表（見た目付き）を xlsx でエクスポート。一時シートを作って書式設定→xlsx化→削除
    if (action === 'export_period_kakunin_xlsx') {
      const dateFrom = String(body.dateFrom || '').trim();
      const dateTo = String(body.dateTo || '').trim();
      if (!dateFrom || !dateTo) return error('開始日と終了日を指定してください');
      if (dateFrom > dateTo) return error('開始日が終了日より後です');
      const companyFilter = String(body.company || '').trim();
      const tag = (companyFilter && companyFilter !== '全社') ? '_' + companyFilter : '';
      const filename = '期間集計' + tag + '_' + dateFrom + '_' + dateTo + '.xlsx';
      const result = exportPeriodKakuninAsXlsxBase64_(ss, dateFrom, dateTo, companyFilter);
      return ok({base64: result.base64, filename: filename});
    }

    if (action === 'add_member') {
      const memberSheet = getOrCreateMemberSheet_(ss);
      const name = String(body.name || '').trim();
      const company = String(body.company || '').trim();
      const division = String(body.division || '').trim();
      const rate = Number(body.rate || 0);
      if (!name || !company) return error('氏名と会社は必須です');
      memberSheet.appendRow([name, company, division, rate]);
      logOperation_(ss, 'add_member', name + '/' + company, '事業部=' + division + ', 単価=' + rate, updatedBy);
      return ok({added: name});
    }

    if (action === 'update_member_division') {
      const memberSheet = getOrCreateMemberSheet_(ss);
      const name = String(body.name || '').trim();
      const company = String(body.company || '').trim();
      const division = String(body.division || '').trim();
      const data = memberSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim() === name && String(data[i][1]).trim() === company) {
          memberSheet.getRange(i + 1, 3).setValue(division);
          logOperation_(ss, 'update_member_division', name + '/' + company, '事業部=' + division, updatedBy);
          return ok({updated: name});
        }
      }
      return ok({updated: null});
    }

    if (action === 'update_member_rate') {
      const memberSheet = getOrCreateMemberSheet_(ss);
      const name = String(body.name || '').trim();
      const company = String(body.company || '').trim();
      const rate = Number(body.rate || 0);
      const data = memberSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim() === name && String(data[i][1]).trim() === company) {
          memberSheet.getRange(i + 1, 4).setValue(rate);
          logOperation_(ss, 'update_member_rate', name + '/' + company, '単価=' + rate, updatedBy);
          return ok({updated: name});
        }
      }
      return ok({updated: null});
    }

    if (action === 'update_site_revenue') {
      const jobSiteSheet = getOrCreateJobSiteSheet_(ss);
      const genba = String(body.genba || '').trim();
      const loc = String(body.loc || '').trim();
      const revenue = Number(body.revenue || 0);
      const data = jobSiteSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim() === genba && String(data[i][1]).trim() === loc) {
          jobSiteSheet.getRange(i + 1, 7).setValue(revenue);
          logOperation_(ss, 'update_site_revenue', genba + '/' + loc, '売上=' + revenue, updatedBy);
          return ok({updated: genba, jobNo: String(data[i][2] || '')});
        }
      }
      return error('現場マスタに該当現場が見つかりません');
    }

    if (action === 'get_billing_rates') {
      const sheet = getOrCreateBillingRateSheet_(ss);
      const data = sheet.getDataRange().getValues();
      const rates = data.length > 1 ? data.slice(1).map(r => ({
        genba: String(r[0] || ''),
        loc: String(r[1] || ''),
        rate: Number(r[2] || 0)
      })).filter(x => x.genba) : [];
      return ok({rates: rates});
    }

    if (action === 'save_billing_rate') {
      const sheet = getOrCreateBillingRateSheet_(ss);
      const genba = String(body.genba || '').trim();
      const loc = String(body.loc || '').trim();
      const rate = Number(body.rate || 0);
      if (!genba) return error('元請名は必須です');
      const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim() === genba && String(data[i][1]).trim() === loc) {
          sheet.getRange(i + 1, 3).setValue(rate);
          sheet.getRange(i + 1, 4).setValue(now);
          return ok({updated: true});
        }
      }
      sheet.appendRow([genba, loc, rate, now]);
      return ok({added: true});
    }

    if (action === 'update_site_billing_method') {
      const jobSiteSheet = getOrCreateJobSiteSheet_(ss);
      const genba = String(body.genba || '').trim();
      const loc = String(body.loc || '').trim();
      const method = String(body.method || '応援').trim();
      const data = jobSiteSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim() === genba && String(data[i][1]).trim() === loc) {
          jobSiteSheet.getRange(i + 1, 10).setValue(method);
          logOperation_(ss, 'update_site_billing_method', genba + '/' + loc, '方式=' + method, updatedBy);
          return ok({updated: true});
        }
      }
      // 現場マスタに行が無い現場（非グローライズ等）は新規行を作って方式を保存（upsert）
      if (!genba || !loc) return error('元請名・現場名は必須です');
      jobSiteSheet.appendRow([genba, loc, '', '', '', '', '', '', '', method]);
      logOperation_(ss, 'update_site_billing_method', genba + '/' + loc, '方式=' + method + '(新規行)', updatedBy);
      return ok({added: true});
    }

    if (action === 'generate_billing_calc_xlsx') {
      const genba = String(body.genba || '');
      const month = String(body.month || '');
      const lines = Array.isArray(body.lines) ? body.lines : [];
      let sheet = ss.getSheetByName(BILLING_CALC_SHEET);
      if (sheet) { sheet.clear(); } else { sheet = ss.insertSheet(BILLING_CALC_SHEET); }
      sheet.appendRow([genba + '　' + month + '　請求計算']);
      sheet.appendRow(['現場名', '出面数', '単価', '金額', '経費', '方式']);
      const dataStart = 3; // 1=タイトル 2=ヘッダ 3=先頭データ
      lines.forEach(ln => {
        const r = sheet.getLastRow() + 1;
        const isOuen = String(ln.method || '応援') === '応援';
        if (isOuen) {
          // 金額 = 出面(B) × 単価(C)
          sheet.appendRow([
            String(ln.loc || ''), Number(ln.manDays || 0), Number(ln.rate || 0),
            '=B' + r + '*C' + r, Number(ln.expense || 0), '応援'
          ]);
        } else {
          // 請負：今回請求額を金額(D)に直接。出面/単価/経費は空
          sheet.appendRow([
            String(ln.loc || ''), '', '', Number(ln.amount || 0), 0, '請負'
          ]);
        }
      });
      const dataEnd = sheet.getLastRow();
      if (dataEnd >= dataStart) {
        const totalRow = dataEnd + 1;
        sheet.getRange(totalRow, 1).setValue('合計');
        sheet.getRange(totalRow, 4).setFormula('=SUM(D' + dataStart + ':D' + dataEnd + ')');
        sheet.getRange(totalRow, 5).setFormula('=SUM(E' + dataStart + ':E' + dataEnd + ')');
        sheet.getRange(totalRow + 1, 3).setValue('請求合計');
        sheet.getRange(totalRow + 1, 4).setFormula('=D' + totalRow + '+E' + totalRow);
      }
      SpreadsheetApp.flush();
      const result = exportSheetAsXlsxBase64_(ss, sheet);
      return ok({base64: result.base64, filename: '請求計算_' + genba + '_' + month + '.xlsx'});
    }

    // 現場の「完了」フラグを設定/解除。工番・売上・既存日報には一切影響しない。
    // body: { genba, loc, completed (bool), updatedBy }
    if (action === 'update_site_status') {
      const jobSiteSheet = getOrCreateJobSiteSheet_(ss);
      const genba = String(body.genba || '').trim();
      const loc = String(body.loc || '').trim();
      const completed = !!body.completed;
      if (!genba) return error('元請名は必須です');
      const data = jobSiteSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim() === genba && String(data[i][1]).trim() === loc) {
          jobSiteSheet.getRange(i + 1, 9).setValue(completed ? '✓' : '');
          logOperation_(ss, 'update_site_status', genba + '/' + loc, completed ? '完了' : '進行中に戻す', updatedBy);
          return ok({updated: genba + '/' + loc, completed: completed});
        }
      }
      return error('現場マスタに該当現場が見つかりません');
    }

    if (action === 'remove_member') {
      const memberSheet = getOrCreateMemberSheet_(ss);
      const name = String(body.name || '').trim();
      const company = String(body.company || '').trim();
      const data = memberSheet.getDataRange().getValues();
      for (let i = data.length - 1; i >= 1; i--) {
        if (String(data[i][0]).trim() === name && String(data[i][1]).trim() === company) {
          memberSheet.deleteRow(i + 1);
          return ok({removed: name});
        }
      }
      return ok({removed: null});
    }

    if (action === 'add_genba') {
      const genbaSheet = getOrCreateGenbaSheet_(ss);
      const name = String(body.name || '').trim();
      const company = String(body.company || '').trim();
      if (!name) return error('元請名は必須です');
      const data = genbaSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim() === name && String(data[i][1] || '').trim() === company) return ok({added: name, duplicate: true});
      }
      // 漢字を含む名称なら読みを自動生成(失敗時は空欄)
      const yomi = generateYomiSafe_(name);
      genbaSheet.appendRow([name, company, yomi]);
      return ok({added: name});
    }

    if (action === 'reassign_jobno') {
      const genba = String(body.genba || '').trim();
      const loc = String(body.loc || '').trim();
      const newDivision = String(body.newDivision || '').trim();
      if (!genba || !newDivision) return error('元請名と新事業部は必須です');

      const jobSiteSheet = getOrCreateJobSiteSheet_(ss);
      const jobSiteData = jobSiteSheet.getDataRange().getValues();
      let siteRowIdx = -1;
      let currentJobNo = '';
      let currentDivision = '';
      let fiscalYear = 0;
      for (let i = 1; i < jobSiteData.length; i++) {
        if (String(jobSiteData[i][0]).trim() === genba && String(jobSiteData[i][1]).trim() === loc) {
          siteRowIdx = i;
          currentJobNo = String(jobSiteData[i][2] || '');
          currentDivision = String(jobSiteData[i][3] || '').trim();
          fiscalYear = Number(jobSiteData[i][4]) || 0;
          break;
        }
      }
      if (siteRowIdx === -1) return error('現場マスタに該当現場が見つかりません');
      if (currentDivision === newDivision) return ok({ message: '事業部は変更されていません' });
      if (!fiscalYear) {
        const now = new Date();
        fiscalYear = now.getMonth() >= 3 ? now.getFullYear() : now.getFullYear() - 1;
      }
      const yearStr = String(fiscalYear).slice(-2);

      let maxSerial = 0;
      for (let i = 1; i < jobSiteData.length; i++) {
        if (String(jobSiteData[i][3]).trim() === newDivision && Number(jobSiteData[i][4]) === fiscalYear) {
          const s = Number(jobSiteData[i][5]) || 0;
          if (s > maxSerial) maxSerial = s;
        }
      }
      const newSerial = maxSerial + 1;
      const newJobNo = `${newDivision}-${yearStr}-${String(newSerial).padStart(3, '0')}`;

      jobSiteSheet.getRange(siteRowIdx + 1, 3).setValue(newJobNo);
      jobSiteSheet.getRange(siteRowIdx + 1, 4).setValue(newDivision);
      jobSiteSheet.getRange(siteRowIdx + 1, 6).setValue(newSerial);

      function updateSheetRows_(targetSheet) {
        if (!targetSheet) return 0;
        const data = targetSheet.getDataRange().getValues();
        if (data.length <= 1) return 0;
        const headers = data[0];
        const gCol = headers.indexOf('元請名');
        const lCol = headers.indexOf('現場名');
        const dCol = headers.indexOf('事業部');
        const jCol = headers.indexOf('工番');
        if (gCol < 0 || lCol < 0 || dCol < 0 || jCol < 0) return 0;
        let cnt = 0;
        for (let i = 1; i < data.length; i++) {
          if (String(data[i][gCol]).trim() === genba && String(data[i][lCol]).trim() === loc) {
            targetSheet.getRange(i + 1, dCol + 1).setValue(newDivision);
            targetSheet.getRange(i + 1, jCol + 1).setValue(newJobNo);
            cnt++;
          }
        }
        return cnt;
      }

      const updatedRows = updateSheetRows_(sheet);
      const archivedUpdated = updateSheetRows_(ss.getSheetByName(ARCHIVE_SHEET));

      logOperation_(ss, 'reassign_jobno', genba + '/' + loc, currentJobNo + '→' + newJobNo + '（日報' + updatedRows + '行・アーカイブ' + archivedUpdated + '行）', updatedBy);
      return ok({ oldJobNo: currentJobNo, newJobNo, updatedRows, archivedUpdated });
    }

    // ============================================================
    // 車両予約（LINEボット連携）
    // トークン認証。既存カレンダー機能とは独立した「車両予約」シートを操作。
    // ============================================================
    if (action === 'vehicle_res_add' || action === 'vehicle_res_update' || action === 'vehicle_res_delete' || action === 'vehicle_res_list') {
      if (String(body.token || '') !== VEHICLE_RES_TOKEN) return error('認証失敗');
      const vehicleSheet = getOrCreateVehicleResSheet_(ss);

      if (action === 'vehicle_res_add') {
        const ev = body.event || {};
        const now = new Date();
        vehicleSheet.appendRow([
          String(ev.reservation_id || ''),
          String(ev.vehicle_name || ''),
          String(ev.plate || ''),
          String(ev.company || ''),
          String(ev.user_name || ''),
          String(ev.user_line_id || ''),
          String(ev.start_dt || ''),
          String(ev.end_dt_planned || ''),
          String(ev.end_dt_actual || ''),
          String(ev.destination || ''),
          String(ev.status || '予約'),
          String(ev.memo || ''),
          now,
          now
        ]);
        logOperation_(ss, 'vehicle_res_add', String(ev.reservation_id || ''), String(ev.vehicle_name || '') + '/' + String(ev.user_name || ''), 'linebot');
        return ok({id: String(ev.reservation_id || '')});
      }

      if (action === 'vehicle_res_update') {
        const ev = body.event || {};
        const id = String(ev.reservation_id || '');
        if (!id) return error('予約IDが指定されていません');
        const data = vehicleSheet.getDataRange().getValues();
        const idCol = VEHICLE_RES_HEADERS.indexOf('予約ID');
        const fieldMap = {
          vehicle_name: '車両名',
          plate: 'ナンバー',
          company: '所有会社',
          user_name: '使用者氏名',
          user_line_id: '使用者LINE_ID',
          start_dt: '開始日時',
          end_dt_planned: '返却予定日時',
          end_dt_actual: '実返却日時',
          destination: '行先',
          status: '状態',
          memo: '備考'
        };
        for (let i = 1; i < data.length; i++) {
          if (String(data[i][idCol]) === id) {
            const updates = data[i].slice();
            Object.keys(fieldMap).forEach(key => {
              if (ev[key] !== undefined) {
                const colIdx = VEHICLE_RES_HEADERS.indexOf(fieldMap[key]);
                if (colIdx >= 0) updates[colIdx] = String(ev[key] || '');
              }
            });
            const updColIdx = VEHICLE_RES_HEADERS.indexOf('更新日時');
            if (updColIdx >= 0) updates[updColIdx] = new Date();
            vehicleSheet.getRange(i + 1, 1, 1, VEHICLE_RES_HEADERS.length).setValues([updates]);
            logOperation_(ss, 'vehicle_res_update', id, '状態=' + String(ev.status || ''), 'linebot');
            return ok({updated: id});
          }
        }
        return error('対象が見つかりませんでした');
      }

      if (action === 'vehicle_res_delete') {
        const id = String(body.id || '');
        if (!id) return error('予約IDが指定されていません');
        const data = vehicleSheet.getDataRange().getValues();
        const idCol = VEHICLE_RES_HEADERS.indexOf('予約ID');
        const statusCol = VEHICLE_RES_HEADERS.indexOf('状態');
        const updCol = VEHICLE_RES_HEADERS.indexOf('更新日時');
        for (let i = 1; i < data.length; i++) {
          if (String(data[i][idCol]) === id) {
            vehicleSheet.getRange(i + 1, statusCol + 1).setValue('キャンセル');
            vehicleSheet.getRange(i + 1, updCol + 1).setValue(new Date());
            logOperation_(ss, 'vehicle_res_delete', id, '論理削除', 'linebot');
            return ok({cancelled: id});
          }
        }
        return error('対象が見つかりませんでした');
      }

      if (action === 'vehicle_res_list') {
        const tz = Session.getScriptTimeZone();
        const data = vehicleSheet.getDataRange().getValues();
        let rows = [];
        if (data.length > 1) {
          const headers = data[0];
          rows = data.slice(1).map(r => {
            const obj = {};
            headers.forEach((h, j) => {
              const v = r[j];
              if (v instanceof Date) {
                obj[h] = Utilities.formatDate(v, tz, "yyyy-MM-dd'T'HH:mm:ssXXX");
              } else {
                obj[h] = (v === undefined || v === null) ? '' : String(v);
              }
            });
            return obj;
          });
        }
        return ok({rows});
      }
    }

    // LINEボット連携: 指定日（既定は今日）の倉庫作業者一覧を返す
    // body: { token, date (YYYY-MM-DD, 省略時は今日) }
    if (action === 'warehouse_today') {
      if (String(body.token || '') !== VEHICLE_RES_TOKEN) return error('認証失敗');
      const tz = Session.getScriptTimeZone();
      const date = String(body.date || '').trim() || Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
      const sh = ss.getSheetByName(SHEET_NAME);
      if (!sh) return error('日報シートが見つかりません');
      const data = sh.getDataRange().getValues();
      if (data.length < 2) return ok({date: date, entries: [], grouped: []});
      const headers = data[0];
      const idx = {
        date: headers.indexOf('作業日'),
        name: headers.indexOf('氏名'),
        role: headers.indexOf('役割'),
        loc: headers.indexOf('現場名'),
        memo: headers.indexOf('メモ'),
        yakin: headers.indexOf('夜勤'),
        workType: headers.indexOf('作業区分'),
        company: headers.indexOf('会社'),
        id: headers.indexOf('ID')
      };
      const entries = [];
      for (let i = 1; i < data.length; i++) {
        const r = data[i];
        const d = r[idx.date] instanceof Date
          ? Utilities.formatDate(r[idx.date], tz, 'yyyy-MM-dd')
          : String(r[idx.date] || '').slice(0, 10);
        if (d !== date) continue;
        // 倉庫モードのレコードのみ（夜勤カラムが '倉庫'）
        if (String(r[idx.yakin] || '') !== '倉庫') continue;
        // 倉庫タスクはメモ欄に格納（現行仕様）。旧データで現場名カラムに入って
        // いるケースがあるため、memo が空なら loc にフォールバック。
        const memoVal = String(r[idx.memo] || '');
        const locVal = String(r[idx.loc] || '');
        entries.push({
          name: String(r[idx.name] || ''),
          role: String(r[idx.role] || ''),
          tasks: memoVal || locVal,
          company: String(r[idx.company] || ''),
          id: String(r[idx.id] || '')
        });
      }
      // ID 単位でグループ化（同じグループ＝同じ tasks）
      const groupMap = {};
      entries.forEach(e => {
        const k = e.id || (e.tasks + '|' + e.company);
        if (!groupMap[k]) groupMap[k] = {tasks: e.tasks, company: e.company, leader: '', members: []};
        if (e.role === '代表' && !groupMap[k].leader) groupMap[k].leader = e.name;
        groupMap[k].members.push(e.name);
      });
      const grouped = Object.values(groupMap);
      return ok({date: date, entries: entries, grouped: grouped});
    }

    if (action === 'remove_genba') {
      const genbaSheet = getOrCreateGenbaSheet_(ss);
      const name = String(body.name || '').trim();
      const company = String(body.company || '').trim();
      const data = genbaSheet.getDataRange().getValues();
      for (let i = data.length - 1; i >= 1; i--) {
        if (String(data[i][0]).trim() === name && String(data[i][1] || '').trim() === company) {
          genbaSheet.deleteRow(i + 1);
          return ok({removed: name});
        }
      }
      return ok({removed: null});
    }

    return error('無効な操作です: ' + action);
  } catch(err) {
    return error(err.toString());
  } finally {
    if (dailyDataLock) dailyDataLock.releaseLock();
    lock.releaseLock();
  }
}

function doGet(e) {
  try {
    if (!calAuthOk_(e && e.parameter && e.parameter.k)) return authError_();
    const requestedCompany = String(e && e.parameter && e.parameter.company || '').trim();
    const filterByCompany = requestedCompany && requestedCompany !== '全社';
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    const tz = Session.getScriptTimeZone();
    let data;
    const snapshotLock = getDailyDataLock_();
    if (!snapshotLock.tryLock(10000)) {
      return error('現在予定を更新中です。数秒待ってから再度お試しください。');
    }
    try {
      ensureHeaders_(sheet);
      data = sheet.getDataRange().getValues();
    } finally {
      // 会社絞込みやマスタ読取の前に解放し、社員保存の待ち時間を最小化する。
      snapshotLock.releaseLock();
    }
    // 2026-08-21 レスポンス改善: 従来は全行をオブジェクト化してから会社で
    // 絞っていたため、捨てる行の整形コストまで毎回払っていた。先に絞る。
    // dateFrom / dateTo（'YYYY-MM-DD'・両端含む）を付ければ期間も絞れる。
    // 省略時は従来どおり全件返すので、古い画面もそのまま動く。
    const reqFrom = String(e && e.parameter && e.parameter.dateFrom || '').trim();
    const reqTo = String(e && e.parameter && e.parameter.dateTo || '').trim();
    // 2026-08-21 転送量削減: compact=1 を付けると、1行を19個のキー付き
    // オブジェクトではなく「値だけの配列」で返し、項目名は headers として
    // 先頭に1回だけ送る。2,600行なら約5万回分のキー文字列が消える。
    // ★互換性: compact を付けない従来の呼び出しは今までどおりの形で返す。
    //   そのため古い画面と新しい画面が混在しても壊れない（デプロイ順不同）。
    const wantCompact = String(e && e.parameter && e.parameter.compact || '') === '1';
    let rows = [];
    let outHeaders = [];
    if (data.length > 1) {
      const headers = data[0];
      outHeaders = headers;
      const dateIdx = headers.indexOf('作業日');
      const companyIdx = headers.indexOf('会社');
      const hLen = headers.length;
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (filterByCompany && companyIdx >= 0 && String(row[companyIdx] || '').trim() !== requestedCompany) continue;
        const dStr = dateIdx >= 0 ? fmtDate_(row[dateIdx], tz) : '';
        if (reqFrom && dStr && dStr < reqFrom) continue;
        if (reqTo && dStr && dStr > reqTo) continue;
        if (wantCompact) {
          const arr = new Array(hLen);
          for (let j = 0; j < hLen; j++) {
            const h = headers[j];
            const v = row[j];
            if (h === '作業日') arr[j] = dStr;
            else if (h === '出勤' || h === '退勤') arr[j] = fmtTime_(v, tz);
            else arr[j] = (v === undefined || v === null) ? '' : v;
          }
          rows.push(arr);
        } else {
          const obj = {};
          for (let j = 0; j < hLen; j++) {
            const h = headers[j];
            const v = row[j];
            if (h === '作業日') obj[h] = dStr;
            else if (h === '出勤' || h === '退勤') obj[h] = fmtTime_(v, tz);
            else obj[h] = (v === undefined || v === null) ? '' : v;
          }
          rows.push(obj);
        }
      }
    }
    const memberSheet = getOrCreateMemberSheet_(ss);
    const mData = memberSheet.getDataRange().getValues();
    const members = mData.length > 1 ? mData.slice(1).map(r => ({
      name: String(r[0]||''),
      company: String(r[1]||''),
      division: String(r[2]||''),
      rate: Number(r[3]||0)
    })).filter(m => !filterByCompany || m.company === requestedCompany) : [];

    const genbaSheet = getOrCreateGenbaSheet_(ss);
    const gData = genbaSheet.getDataRange().getValues();
    const genbaMaster = gData.length > 1 ? gData.slice(1).map(r => ({name: String(r[0]||''), company: String(r[1]||'')})).filter(g => g.name && (!filterByCompany || !g.company || g.company === requestedCompany)) : [];
    const allowedGenba = new Set(genbaMaster.map(g => g.name));

    // 現場マスタも返す（完了フラグでプルダウンを絞り込むため）
    const jobSiteSheet = getOrCreateJobSiteSheet_(ss);
    const jData = jobSiteSheet.getDataRange().getValues();
    const jobsites = jData.length > 1 ? jData.slice(1).map(r => ({
      genba: String(r[0] || ''),
      loc: String(r[1] || ''),
      jobNo: String(r[2] || ''),
      completed: String(r[8] || '').trim() !== '',
      billingMethod: String(r[9] || '').trim() || '応援',
      kyoten: String(r[10] || '').trim()    // ★2026-08-26 拠点。空なら画面側は会社の既定値を使う
    })).filter(j => j.genba && (!filterByCompany || allowedGenba.has(j.genba))) : [];

    if (wantCompact) return ok({compact: 1, headers: outHeaders, rows, members, genbaMaster, jobsites});
    return ok({rows, members, genbaMaster, jobsites});
  } catch(err) {
    return error(err.toString());
  }
}

function getOrCreatePresSheet_(ss) {
  let sheet = ss.getSheetByName(PRES_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(PRES_SHEET);
    sheet.appendRow(PRES_HEADERS);
    try { sheet.hideSheet(); } catch (e) {}
  } else {
    ensureColumns_(sheet, PRES_HEADERS.length);
    const headers = sheet.getRange(1, 1, 1, PRES_HEADERS.length).getValues()[0];
    PRES_HEADERS.forEach((h, i) => {
      if (String(headers[i] || '').trim() !== h) sheet.getRange(1, i + 1).setValue(h);
    });
  }
  return sheet;
}

// 車両予約シート（LINEボット連携）。既存カレンダーには影響しない独立シート。
function getOrCreateVehicleResSheet_(ss) {
  let sheet = ss.getSheetByName(VEHICLE_RES_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(VEHICLE_RES_SHEET);
    sheet.appendRow(VEHICLE_RES_HEADERS);
  } else {
    ensureColumns_(sheet, VEHICLE_RES_HEADERS.length);
    const headers = sheet.getRange(1, 1, 1, VEHICLE_RES_HEADERS.length).getValues()[0];
    VEHICLE_RES_HEADERS.forEach((h, i) => {
      if (String(headers[i] || '').trim() !== h) sheet.getRange(1, i + 1).setValue(h);
    });
  }
  return sheet;
}

function getOrCreateMemberSheet_(ss) {
  let sheet = ss.getSheetByName(MEMBER_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(MEMBER_SHEET);
    sheet.appendRow(['氏名', '会社', '事業部', '単価']);
  } else {
    ensureColumns_(sheet, 4);
    const headers = sheet.getRange(1, 1, 1, 4).getValues()[0];
    if (String(headers[2] || '').trim() !== '事業部') sheet.getRange(1, 3).setValue('事業部');
    if (String(headers[3] || '').trim() !== '単価') sheet.getRange(1, 4).setValue('単価');
  }
  return sheet;
}

function getOrCreateGenbaSheet_(ss) {
  let sheet = ss.getSheetByName(GENBA_MASTER_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(GENBA_MASTER_SHEET);
    sheet.appendRow(['元請名', '会社', '読み']);
  } else {
    ensureColumns_(sheet, 3);
    const headers = sheet.getRange(1, 1, 1, 3).getValues()[0];
    if (String(headers[1] || '').trim() !== '会社') sheet.getRange(1, 2).setValue('会社');
    if (String(headers[2] || '').trim() !== '読み') sheet.getRange(1, 3).setValue('読み');
  }
  return sheet;
}

function getOrCreateJobSiteSheet_(ss) {
  let sheet = ss.getSheetByName(JOBSITE_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(JOBSITE_SHEET);
    sheet.appendRow(['元請名', '現場名', '工番', '事業部', '年度', '連番', '売上', '読み', '完了', '請求方式', '拠点']);
  } else {
    ensureColumns_(sheet, 11);
    const headers = sheet.getRange(1, 1, 1, 11).getValues()[0];
    if (String(headers[6] || '').trim() !== '売上') sheet.getRange(1, 7).setValue('売上');
    if (String(headers[7] || '').trim() !== '読み') sheet.getRange(1, 8).setValue('読み');
    if (String(headers[8] || '').trim() !== '完了') sheet.getRange(1, 9).setValue('完了');
    if (String(headers[9] || '').trim() !== '請求方式') sheet.getRange(1, 10).setValue('請求方式');
    // ★2026-08-26: 拠点（本社/関東支店）。現場を選べば予定に自動で入る＝入力を増やさない
    if (String(headers[10] || '').trim() !== '拠点') sheet.getRange(1, 11).setValue('拠点');
  }
  return sheet;
}

function getOrCreateBillingRateSheet_(ss) {
  let sheet = ss.getSheetByName(BILLING_RATE_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(BILLING_RATE_SHEET);
    sheet.appendRow(['元請名', '現場名', '単価', '更新日時']);
    return sheet;
  }
  // 旧5列（元請/現場/職人/単価/更新日時）からの移行：職人列があれば元請×現場へ畳む
  const headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
  if (String(headers[2] || '').trim() === '職人名') {
    const data = sheet.getDataRange().getValues();
    const map = {};
    for (let i = 1; i < data.length; i++) {
      const g = String(data[i][0] || '').trim();
      if (!g) continue;
      const l = String(data[i][1] || '').trim();
      map[g + '|||' + l] = { genba: g, loc: l, rate: Number(data[i][3] || 0), ts: String(data[i][4] || '') };
    }
    sheet.clear();
    sheet.appendRow(['元請名', '現場名', '単価', '更新日時']);
    Object.keys(map).forEach(k => { const x = map[k]; sheet.appendRow([x.genba, x.loc, x.rate, x.ts]); });
  }
  return sheet;
}

// 現場マスタの孤立行を削除（日報データ＋アーカイブのいずれにも参照されず、売上未入力の行）
function cleanupOrphanSites_(ss) {
  const jobSiteSheet = getOrCreateJobSiteSheet_(ss);
  const data = jobSiteSheet.getDataRange().getValues();
  if (data.length <= 1) return 0;
  // 日報データ＋アーカイブから使用中の (元請+現場) と 工番 を収集
  const usedKeys = new Set();
  const usedJobNos = new Set();
  [SHEET_NAME, ARCHIVE_SHEET].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    const sd = sh.getDataRange().getValues();
    if (sd.length <= 1) return;
    const headers = sd[0];
    const gC = headers.indexOf('元請名');
    const lC = headers.indexOf('現場名');
    const jC = headers.indexOf('工番');
    for (let i = 1; i < sd.length; i++) {
      const g = String(sd[i][gC] || '').trim();
      const l = String(sd[i][lC] || '').trim();
      const j = String(sd[i][jC] || '').trim();
      if (g) usedKeys.add(g + '|||' + l);
      if (j) usedJobNos.add(j);
    }
  });
  // 削除候補（後ろから走査して deleteRow しても index がずれないように）
  const rowsToDelete = [];
  for (let i = data.length - 1; i >= 1; i--) {
    const genba = String(data[i][0] || '').trim();
    const loc = String(data[i][1] || '').trim();
    const jobNo = String(data[i][2] || '').trim();
    const revenue = Number(data[i][6] || 0);
    if (revenue > 0) continue; // 売上が入っている行は将来の現場として残す
    const key = genba + '|||' + loc;
    const refByKey = usedKeys.has(key);
    const refByJob = jobNo && usedJobNos.has(jobNo);
    if (!refByKey && !refByJob) {
      rowsToDelete.push(i + 1); // 1-indexed
    }
  }
  rowsToDelete.forEach(rowNum => jobSiteSheet.deleteRow(rowNum));
  return rowsToDelete.length;
}

function getOrCreateOpLogSheet_(ss) {
  let sheet = ss.getSheetByName(OPLOG_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(OPLOG_SHEET);
    sheet.appendRow(['日時', '操作', '対象', '詳細', '実行者']);
  }
  return sheet;
}

function logOperation_(ss, action, target, detail, user) {
  try {
    const sheet = getOrCreateOpLogSheet_(ss);
    sheet.appendRow([new Date().toLocaleString('ja-JP'), action, target, detail, user || '']);
  } catch (e) {}
}

function getMemberDivision_(ss, name) {
  if (!name) return '';
  const sheet = getOrCreateMemberSheet_(ss);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === name) {
      return String(data[i][2] || '').trim();
    }
  }
  return '';
}

function getOrGenerateJobNo_(ss, genba, loc, division) {
  if (!division || !genba) return '';
  const sheet = getOrCreateJobSiteSheet_(ss);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === genba && String(data[i][1]).trim() === loc) {
      return String(data[i][2]);
    }
  }

  const now = new Date();
  const fiscalYear = now.getMonth() >= 3 ? now.getFullYear() : now.getFullYear() - 1;
  const yearStr = String(fiscalYear).slice(-2);

  let maxSerial = 0;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][3]).trim() === division && Number(data[i][4]) === fiscalYear) {
      const serial = Number(data[i][5]) || 0;
      if (serial > maxSerial) maxSerial = serial;
    }
  }
  const newSerial = maxSerial + 1;
  const jobNo = `${division}-${yearStr}-${String(newSerial).padStart(3, '0')}`;

  // 現場名の読みを自動生成(漢字なしなら空)。売上は空欄のまま。
  const yomi = generateYomiSafe_(loc);
  sheet.appendRow([genba, loc, jobNo, division, fiscalYear, newSerial, '', yomi]);
  return jobNo;
}

// ========== 集計機能 ==========

function generateSummary_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = Session.getScriptTimeZone();

  function sheetToRecords(sheet) {
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    const headers = data[0];
    const colIdx = {};
    headers.forEach((h, j) => colIdx[h] = j);
    return data.slice(1).map(row => {
      const dateVal = row[colIdx['作業日']];
      const dateStr = (dateVal instanceof Date) ? Utilities.formatDate(dateVal, tz, 'yyyy-MM-dd') : String(dateVal || '');
      return {
        date: dateStr, month: dateStr.slice(0, 7),
        name: String(row[colIdx['氏名']] || ''),
        kosu: Number(row[colIdx['人工']] || 0),
        company: String(row[colIdx['会社']] || ''),
        genba: String(row[colIdx['元請名']] || ''),
        loc: String(row[colIdx['現場名']] || ''),
        yakin: String(row[colIdx['夜勤']] || '')
      };
    }).filter(r => r.date && r.name);
  }

  let mainRecords;
  let archiveRecords;
  const snapshotLock = getDailyDataLock_();
  if (!snapshotLock.tryLock(10000)) {
    throw new Error('現在予定を更新中です。集計は数秒後に再実行してください。');
  }
  try {
    mainRecords = sheetToRecords(ss.getSheetByName(SHEET_NAME));
    archiveRecords = sheetToRecords(ss.getSheetByName(ARCHIVE_SHEET));
  } finally {
    // 長い集計シート書込みの前に解放し、社員の保存を待たせない。
    snapshotLock.releaseLock();
  }

  generateCompanySummary_(ss, mainRecords);
  generateMonthSummary_(ss, mainRecords);
  generateBillingSummary_(ss, mainRecords);
  generateBillingFilterSheet_(ss, mainRecords);

  const allRecords = [...mainRecords, ...archiveRecords];
  generateKakuninTable_(ss, allRecords);
  generateDivisionAllocation_(ss, allRecords);

  // 孤立現場の削除は保存処理と競合するため、管理画面の手動操作だけで行う。
}

function calcEffective_(records, name) {
  const byDate = {};
  records.filter(r => r.name === name).forEach(r => {
    if (r.yakin === '休み' || r.yakin === '予定') return;
    if (!byDate[r.date]) byDate[r.date] = {day: 0, night: 0, hasDay: false, hasNight: false};
    if (r.yakin === '夜勤') {
      byDate[r.date].night = Math.max(byDate[r.date].night, r.kosu);
      byDate[r.date].hasNight = true;
    } else {
      byDate[r.date].day = Math.max(byDate[r.date].day, r.kosu);
      byDate[r.date].hasDay = true;
    }
  });
  let days = 0, kosu = 0, yakinCount = 0;
  Object.values(byDate).forEach(v => {
    if (v.hasDay) { days++; kosu += v.day; }
    if (v.hasNight) { days++; kosu += v.night; yakinCount++; }
  });
  return {days, kosu, yakinCount, dates: Object.keys(byDate).sort()};
}

// 会社別/月別集計（6列）の書式を一括適用する。
// 以前は formats を1行ずつ setBackground/setFontWeight していて通信が数百回に達し、集計を重くしていた。
// 背景色・太字を行×列グリッドにまとめ、setBackgrounds/setFontWeights 各1回で流し込む（フォントサイズだけ対象が少ないので個別）。
function applyGroupedSummaryFormats_(sheet, numRows, formats, accentType, accentBg) {
  const bgs = [], fws = [];
  for (let i = 0; i < numRows; i++) {
    bgs.push(['#FFFFFF', '#FFFFFF', '#FFFFFF', '#FFFFFF', '#FFFFFF', '#FFFFFF']);
    fws.push(['normal', 'normal', 'normal', 'normal', 'normal', 'normal']);
  }
  const fontSizes = [];
  formats.forEach(f => {
    const ri = f.row - 1;
    if (f.type === 'title') { fws[ri][0] = 'bold'; fontSizes.push({ row: f.row, size: 14 }); }
    else if (f.type === accentType) { for (let c = 0; c < 6; c++) bgs[ri][c] = accentBg; fws[ri][0] = 'bold'; fontSizes.push({ row: f.row, size: 12 }); }
    else if (f.type === 'header') { for (let c = 0; c < 6; c++) { bgs[ri][c] = '#F5F5F5'; fws[ri][c] = 'bold'; } }
    else if (f.type === 'total') { for (let c = 0; c < 6; c++) { bgs[ri][c] = '#FFF9C4'; fws[ri][c] = 'bold'; } }
  });
  sheet.getRange(1, 1, numRows, 6).setBackgrounds(bgs).setFontWeights(fws);
  fontSizes.forEach(fs => sheet.getRange(fs.row, 1).setFontSize(fs.size));
}

function generateCompanySummary_(ss, records) {
  let sheet = ss.getSheetByName(SUMMARY_COMPANY);
  if (sheet) { sheet.clear(); sheet.clearFormats(); } else { sheet = ss.insertSheet(SUMMARY_COMPANY); }
  const companies = [...new Set(records.map(r => r.company))].filter(Boolean).sort();
  const now = new Date();
  const thisMonth = now.getFullYear() + '-' + String(now.getMonth() + 1).padStart(2, '0');
  const allRows = [];
  const formats = [];
  allRows.push(['会社別集計', '', '', '', '更新日時: ' + new Date().toLocaleString('ja-JP'), '']);
  formats.push({row: allRows.length, type: 'title'});
  allRows.push(['', '', '', '', '', '']);
  companies.forEach(company => {
    const cr = records.filter(r => r.company === company);
    const mr = cr.filter(r => r.month === thisMonth);
    allRows.push(['▶ ' + company, '', '', '', '', '']);
    formats.push({row: allRows.length, type: 'company'});
    allRows.push(['氏名', '当月出勤日数', '当月人工', '当月夜勤回数', '全期間出勤日数', '全期間人工']);
    formats.push({row: allRows.length, type: 'header'});
    // 実働(休み/予定以外)のあるメンバーのみ氏名に含める。倉庫は実働扱い。
    const effRecords = cr.filter(r => r.yakin !== '休み' && r.yakin !== '予定');
    const names = [...new Set(effRecords.map(r => r.name))].sort();
    let tMD=0,tMK=0,tMY=0,tAD=0,tAK=0;
    names.forEach(name => {
      const mEff=calcEffective_(mr, name), aEff=calcEffective_(cr, name);
      tMD+=mEff.days;tMK+=mEff.kosu;tMY+=mEff.yakinCount;tAD+=aEff.days;tAK+=aEff.kosu;
      allRows.push([name, mEff.days, mEff.kosu, mEff.yakinCount, aEff.days, aEff.kosu]);
    });
    allRows.push(['合計', tMD, tMK, tMY, tAD, tAK]);
    formats.push({row: allRows.length, type: 'total'});
    allRows.push(['', '', '', '', '', '']);
  });
  if (allRows.length > 0) {
    sheet.getRange(1, 1, allRows.length, 6).setValues(allRows);
    applyGroupedSummaryFormats_(sheet, allRows.length, formats, 'company', '#E8F5E9');
  }
  sheet.setColumnWidth(1, 120);
  for (let c = 2; c <= 6; c++) sheet.setColumnWidth(c, 110);
}

function generateMonthSummary_(ss, records) {
  let sheet = ss.getSheetByName(SUMMARY_MONTH);
  if (sheet) { sheet.clear(); sheet.clearFormats(); } else { sheet = ss.insertSheet(SUMMARY_MONTH); }
  const months = [...new Set(records.map(r => r.month))].filter(Boolean).sort().reverse();
  const allRows = [];
  const formats = [];
  allRows.push(['月別集計', '', '', '', '更新日時: ' + new Date().toLocaleString('ja-JP'), '']);
  formats.push({row: allRows.length, type: 'title'});
  allRows.push(['', '', '', '', '', '']);
  months.forEach(month => {
    const mr = records.filter(r => r.month === month);
    const parts = month.split('-');
    const label = parts[0] + '年' + Number(parts[1]) + '月';
    allRows.push(['▶ ' + label, '', '', '', '', '']);
    formats.push({row: allRows.length, type: 'month'});
    allRows.push(['氏名', '会社', '出勤日数', '人工合計', '夜勤回数', '日別詳細']);
    formats.push({row: allRows.length, type: 'header'});
    // 実働(休み/予定以外)のあるメンバーのみ表示。倉庫は実働扱い。
    const effRecords = mr.filter(r => r.yakin !== '休み' && r.yakin !== '予定');
    const names = [...new Set(effRecords.map(r => r.name))].sort();
    let tD=0,tK=0,tY=0;
    names.forEach(name => {
      const eff=calcEffective_(mr, name);
      const b=mr.filter(r=>r.name===name);
      tD+=eff.days;tK+=eff.kosu;tY+=eff.yakinCount;
      allRows.push([name, b[0].company||'', eff.days, eff.kosu, eff.yakinCount, eff.dates.map(x=>x.slice(5)).join(', ')]);
    });
    allRows.push(['合計', '', tD, tK, tY, '']);
    formats.push({row: allRows.length, type: 'total'});
    allRows.push(['', '', '', '', '', '']);
  });
  if (allRows.length > 0) {
    sheet.getRange(1, 1, allRows.length, 6).setValues(allRows);
    applyGroupedSummaryFormats_(sheet, allRows.length, formats, 'month', '#E3F2FD');
  }
  sheet.setColumnWidth(1, 100); sheet.setColumnWidth(2, 120);
  for (let c = 3; c <= 5; c++) sheet.setColumnWidth(c, 100);
  sheet.setColumnWidth(6, 300);
}

// 期間指定の月別確認表風データを生成（シートには書かない、CSV化用の 2D 配列を返す）
// dateFrom/dateTo は 'YYYY-MM-DD' 両端含む。companyFilter が空なら全社、'全社' も全社扱い。
// 既存の generateKakuninTable_ と同じく、休み/予定レコードは合計から除外、
// 同日 昼+夜勤は別バケットで max を取り合算する。
function generatePeriodKakuninData_(ss, dateFrom, dateTo, companyFilter) {
  const tz = Session.getScriptTimeZone();
  // 日報データとアーカイブ両方からレコードを集める（期間によってはアーカイブ側にしかない可能性）
  const allRecords = [];
  [SHEET_NAME, ARCHIVE_SHEET].forEach(sname => {
    const sh = ss.getSheetByName(sname);
    if (!sh) return;
    const data = sh.getDataRange().getValues();
    if (data.length < 2) return;
    const headers = data[0];
    const idx = {
      date: headers.indexOf('作業日'),
      name: headers.indexOf('氏名'),
      kosu: headers.indexOf('人工'),
      yakin: headers.indexOf('夜勤'),
      company: headers.indexOf('会社')
    };
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const d = row[idx.date] instanceof Date
        ? Utilities.formatDate(row[idx.date], tz, 'yyyy-MM-dd')
        : String(row[idx.date] || '').slice(0, 10);
      if (!d || d < dateFrom || d > dateTo) continue;
      const co = String(row[idx.company] || '');
      if (companyFilter && companyFilter !== '全社' && co !== companyFilter) continue;
      allRecords.push({
        date: d,
        name: String(row[idx.name] || ''),
        kosu: Number(row[idx.kosu]) || 0,
        yakin: String(row[idx.yakin] || ''),
        company: co
      });
    }
  });

  // 期間内の日付リスト
  const days = [];
  const sd = new Date(dateFrom + 'T00:00:00');
  const ed = new Date(dateTo + 'T00:00:00');
  for (let d = new Date(sd); d <= ed; d.setDate(d.getDate() + 1)) {
    days.push(Utilities.formatDate(d, tz, 'yyyy-MM-dd'));
  }

  // 実働(休み/予定以外)のあるメンバーのみ表示
  const effRecords = allRecords.filter(r => r.yakin !== '休み' && r.yakin !== '予定');
  const names = [...new Set(effRecords.map(r => r.name))].filter(Boolean).sort();

  function getKosuForDay(name, dateStr) {
    const dayRecords = allRecords.filter(r => r.name === name && r.date === dateStr);
    const effective = dayRecords.filter(r => r.yakin !== '休み' && r.yakin !== '予定');
    if (effective.length === 0) return 0;
    let dayKosu = 0, nightKosu = 0;
    effective.forEach(r => {
      const k = Number(r.kosu) || 0;
      if (r.yakin === '夜勤') {
        if (k > nightKosu) nightKosu = k;
      } else {
        if (k > dayKosu) dayKosu = k;
      }
    });
    return dayKosu + nightKosu;
  }

  const dayNames = ['日','月','火','水','木','金','土'];
  // ヘッダ: ['名前 ▼', 'M/D(曜)', ..., '合計']
  const header = ['名前 ▼'].concat(days.map(d => {
    const dt = new Date(d + 'T00:00:00');
    return (dt.getMonth() + 1) + '/' + dt.getDate() + '(' + dayNames[dt.getDay()] + ')';
  })).concat(['合計']);

  // タイトル行
  const titleRow = ['期間: ' + dateFrom + ' 〜 ' + dateTo + (companyFilter && companyFilter !== '全社' ? ' / ' + companyFilter : ' / 全社')];

  // データ行
  const dataRows = names.map(name => {
    const row = [name];
    let total = 0;
    days.forEach(d => {
      const k = getKosuForDay(name, d);
      row.push(k > 0 ? k : 0);
      total += k;
    });
    row.push(total);
    return row;
  });

  // 合計行
  const totalRow = ['合計'];
  let grandTotal = 0;
  days.forEach(d => {
    let s = 0;
    names.forEach(n => { s += getKosuForDay(n, d); });
    totalRow.push(s > 0 ? s : 0);
    grandTotal += s;
  });
  totalRow.push(grandTotal);

  return {
    rows: [titleRow, header].concat(dataRows).concat([totalRow]),
    dateFrom: dateFrom,
    dateTo: dateTo,
    days: days.length,
    members: names.length
  };
}

// 列番号 → 列文字（1 → A、27 → AA など）
function colLetter_(n) {
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

// 指定シートを xlsx として書き出し base64 で返す（書式・色・罫線そのまま保持）
// 内部的に Google Sheets の export URL を OAuth トークン付きで叩く方式。
function exportSheetAsXlsxBase64_(ss, sheet) {
  const url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export'
    + '?format=xlsx&gid=' + sheet.getSheetId();
  const resp = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  if (resp.getResponseCode() >= 300) {
    throw new Error('xlsxエクスポート失敗: HTTP ' + resp.getResponseCode());
  }
  const bytes = resp.getBlob().getBytes();
  return { base64: Utilities.base64Encode(bytes) };
}

// 期間指定の月別確認表（見た目付き）を一時シートに描いて xlsx に書き出し、base64 で返す
function exportPeriodKakuninAsXlsxBase64_(ss, dateFrom, dateTo, companyFilter) {
  // 1) データを準備
  const data = generatePeriodKakuninData_(ss, dateFrom, dateTo, companyFilter);
  const rows = data.rows;
  const daysInRange = data.days;
  const namesCount = data.members;
  const totalCols = 1 + daysInRange + 1; // 名前 + 日付×n + 合計

  // 2) 一時シートを作成（重複名対策でタイムスタンプ）
  const tempName = '_TMP期間集計_' + (new Date().getTime());
  const tempSheet = ss.insertSheet(tempName);

  try {
    // 3) 値を一括書き込み（rows は可変長なので 2D 配列に揃える）
    const writeData = rows.map(r => {
      const out = [];
      for (let i = 0; i < totalCols; i++) out.push(i < r.length ? r[i] : '');
      return out;
    });
    tempSheet.getRange(1, 1, writeData.length, totalCols).setValues(writeData);

    // 4) 列幅
    tempSheet.setColumnWidth(1, 100);
    for (let c = 2; c <= 1 + daysInRange; c++) tempSheet.setColumnWidth(c, 28);
    tempSheet.setColumnWidth(totalCols, 50);

    // 5) 書式：タイトル行（黄色背景、結合、太字）
    const titleRow = tempSheet.getRange(1, 1, 1, totalCols);
    titleRow.merge().setHorizontalAlignment('center').setFontSize(13)
      .setFontWeight('bold').setBackground('#F9E400');

    // 6) ヘッダ行：灰色背景、整数書式、土日色、太字
    const headerRow = tempSheet.getRange(2, 1, 1, totalCols);
    headerRow.setFontWeight('bold').setBackground('#CCCCCC').setHorizontalAlignment('center');
    tempSheet.getRange(2, 2, 1, daysInRange).setNumberFormat('@');  // 文字として扱う（M/D(曜) なので数値解釈の心配は元々ないが念のため）
    // 日付ごとに曜日色を適用
    const sd = new Date(dateFrom + 'T00:00:00');
    for (let i = 0; i < daysInRange; i++) {
      const d = new Date(sd.getFullYear(), sd.getMonth(), sd.getDate() + i);
      const dow = d.getDay();
      const cell = tempSheet.getRange(2, 2 + i);
      if (dow === 0) cell.setFontColor('#CC0000');
      else if (dow === 6) cell.setFontColor('#0000CC');
    }

    // 7) データ行：交互背景、ゼロは薄色、土日列の背景
    const dataStartRow = 3;
    const dataEndRow = dataStartRow + namesCount - 1;
    for (let ri = 0; ri < namesCount; ri++) {
      const r = dataStartRow + ri;
      const bg = ri % 2 === 0 ? '#FFFFFF' : '#F0FFF0';
      tempSheet.getRange(r, 1, 1, totalCols).setBackground(bg);
      tempSheet.getRange(r, 1).setFontWeight('bold');
      tempSheet.getRange(r, 2, 1, totalCols - 1).setNumberFormat('0.##').setHorizontalAlignment('center');
      // ゼロセルのフォント色 + 土日列の背景
      for (let i = 0; i < daysInRange; i++) {
        const d = new Date(sd.getFullYear(), sd.getMonth(), sd.getDate() + i);
        const dow = d.getDay();
        const cell = tempSheet.getRange(r, 2 + i);
        const v = rows[2 + ri][1 + i];
        if (v === 0) cell.setFontColor('#CCCCCC');
        if (dow === 0) cell.setBackground('#FFE6E6');
        else if (dow === 6) cell.setBackground('#E6E6FF');
      }
      tempSheet.getRange(r, totalCols).setFontWeight('bold').setHorizontalAlignment('center');
    }

    // 8) 合計行：黄色背景、太字、罫線
    const totalRowNum = dataEndRow + 1;
    tempSheet.getRange(totalRowNum, 1, 1, totalCols)
      .setFontWeight('bold').setBackground('#FFF9C4').setHorizontalAlignment('center');
    tempSheet.getRange(totalRowNum, 2, 1, totalCols - 1).setNumberFormat('0.##');
    // テーブル全体に罫線
    if (namesCount > 0) {
      tempSheet.getRange(2, 1, namesCount + 2, totalCols).setBorder(true, true, true, true, true, true);
    }

    // 8.5) 合計セルを SUM 関数に置き換え（Excel で値を編集すると合計が自動更新される）
    if (namesCount > 0 && daysInRange > 0) {
      const firstDayCol = colLetter_(2);                       // B
      const lastDayCol  = colLetter_(1 + daysInRange);          // 例: AE
      const totalColLet = colLetter_(totalCols);

      // 各データ行の右端「合計」列を =SUM(B行:AE行)
      const dataTotalFormulas = [];
      for (let r = dataStartRow; r <= dataEndRow; r++) {
        dataTotalFormulas.push([`=SUM(${firstDayCol}${r}:${lastDayCol}${r})`]);
      }
      tempSheet.getRange(dataStartRow, totalCols, namesCount, 1).setFormulas(dataTotalFormulas);

      // 合計行：各日列 =SUM(列<dataStart>:列<dataEnd>) ＋ 右端 =SUM(B合計行:AE合計行)
      const totalRowFormulas = [[]];
      for (let i = 0; i < daysInRange; i++) {
        const col = colLetter_(2 + i);
        totalRowFormulas[0].push(`=SUM(${col}${dataStartRow}:${col}${dataEndRow})`);
      }
      totalRowFormulas[0].push(`=SUM(${firstDayCol}${totalRowNum}:${lastDayCol}${totalRowNum})`);
      tempSheet.getRange(totalRowNum, 2, 1, daysInRange + 1).setFormulas(totalRowFormulas);
    }

    // 9) ヘッダ行を固定（行のみ。列を固定するとタイトル行のセル結合と競合してエラーになる）
    tempSheet.setFrozenRows(2);

    // 10) 一時シートに反映してから xlsx エクスポート（少し待つ）
    SpreadsheetApp.flush();

    const result = exportSheetAsXlsxBase64_(ss, tempSheet);
    return result;
  } finally {
    // 11) 一時シートを削除
    try { ss.deleteSheet(tempSheet); } catch (e) { /* 削除失敗は黙殺 */ }
  }
}

function generateKakuninTable_(ss, records) {
  let sheet = ss.getSheetByName(KAKUNIN_SHEET);
  if (sheet) {
    sheet.clear();
    sheet.clearFormats();
  } else {
    sheet = ss.insertSheet(KAKUNIN_SHEET);
  }

  const now = new Date();
  const months = [];
  for (let i = 1; i >= -2; i--) {
    const d = new Date(now.getFullYear(), now.getMonth() + i, 1);
    months.push({ year: d.getFullYear(), month: d.getMonth() });
  }

  const maxCols = 33;
  ensureColumns_(sheet, maxCols);
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidths(2, 31, 28);
  sheet.setColumnWidth(33, 50);

  const outputData = [];
  const formatRules = [];

  months.forEach(({ year, month }) => {
    const monthStr = year + '-' + String(month + 1).padStart(2, '0');
    const daysInMonth = new Date(year, month + 1, 0).getDate();
    const mr = records.filter(r => r.month === monthStr);
    // 実働(休み/予定以外)のあるメンバーのみ表示。倉庫は実働扱い。
    const effRecords = mr.filter(r => r.yakin !== '休み' && r.yakin !== '予定');
    const names = [...new Set(effRecords.map(r => r.name))].filter(Boolean).sort();
    const totalCols = daysInMonth + 2;

    function getKosuForDay(name, day) {
      const dateStr = year + '-' + String(month + 1).padStart(2, '0') + '-' + String(day).padStart(2, '0');
      const dayRecords = mr.filter(r => r.name === name && r.date === dateStr);
      if (dayRecords.length === 0) return 0;
      // 休み・予定の単体レコードは除外（同日に実働があればそちらを採用、calcEffective_と同じ挙動）
      const effective = dayRecords.filter(r => r.yakin !== '休み' && r.yakin !== '予定');
      if (effective.length === 0) return 0;
      // 昼/夜勤は別バケットでmaxを取り、最後に合算（同日 昼+夜勤=2.0）
      let dayKosu = 0, nightKosu = 0;
      effective.forEach(r => {
        const k = Number(r.kosu) || 0;
        if (r.yakin === '夜勤') {
          if (k > nightKosu) nightKosu = k;
        } else {
          if (k > dayKosu) dayKosu = k;
        }
      });
      return dayKosu + nightKosu;
    }

    const titleRow = Array(maxCols).fill('');
    titleRow[0] = year + '年' + (month + 1) + '月';
    outputData.push(titleRow);
    formatRules.push({ type: 'title', row: outputData.length - 1, cols: totalCols });

    const headerRow = Array(maxCols).fill('');
    headerRow[0] = '名前 ▼';
    for (let d = 1; d <= daysInMonth; d++) headerRow[d] = d;
    headerRow[daysInMonth + 1] = '合計';
    outputData.push(headerRow);
    formatRules.push({ type: 'header', row: outputData.length - 1, cols: totalCols, year, month, daysInMonth });

    if (names.length === 0) {
      const emptyRow = Array(maxCols).fill('');
      emptyRow[0] = '（データなし）';
      outputData.push(emptyRow);
      formatRules.push({ type: 'empty_data', row: outputData.length - 1 });
      outputData.push(Array(maxCols).fill(''));
      formatRules.push({ type: 'empty', row: outputData.length - 1 });
      return;
    }

    names.forEach((name, ni) => {
      const row = Array(maxCols).fill('');
      row[0] = name;
      let total = 0;
      for (let d = 1; d <= daysInMonth; d++) {
        const k = getKosuForDay(name, d);
        row[d] = k > 0 ? k : 0;
        total += k;
      }
      row[daysInMonth + 1] = total;
      outputData.push(row);
      formatRules.push({ type: 'data', row: outputData.length - 1, cols: totalCols, index: ni, year, month, daysInMonth });
    });

    const totalRow = Array(maxCols).fill('');
    totalRow[0] = '合計';
    let grandTotal = 0;
    for (let d = 1; d <= daysInMonth; d++) {
      let dayTotal = 0;
      names.forEach(name => { dayTotal += getKosuForDay(name, d); });
      totalRow[d] = dayTotal > 0 ? dayTotal : 0;
      grandTotal += dayTotal;
    }
    totalRow[daysInMonth + 1] = grandTotal;
    outputData.push(totalRow);
    formatRules.push({ type: 'total', row: outputData.length - 1, cols: totalCols, daysInMonth, namesLength: names.length });

    outputData.push(Array(maxCols).fill(''));
    formatRules.push({ type: 'empty', row: outputData.length - 1 });
  });

  if (outputData.length > 0) {
    const numRows = outputData.length;
    sheet.getRange(1, 1, numRows, maxCols).setValues(outputData);

    // === 書式は「表全体を一括設定」する ===
    // 以前は1マスずつ setBackground/setFontColor 等を呼んでいたため、確認表（人×最大31日×4ヶ月）で
    // Spreadsheetサービスへの通信が数千回に達し、集計が3分超→「サービスに接続できなくなりました」で失敗していた。
    // 書式を行×列のグリッドに組み立て、列一括の setBackgrounds/setFontColors 等で流し込む（通信を数千回→数百回に削減）。
    const bgs = [], fcs = [], has = [], fws = [];
    for (let i = 0; i < numRows; i++) {
      bgs.push(new Array(maxCols).fill('#FFFFFF'));
      fcs.push(new Array(maxCols).fill('#000000'));
      has.push(new Array(maxCols).fill('left'));
      fws.push(new Array(maxCols).fill('normal'));
    }
    // 一括にできない書式（結合／罫線／数値書式／フォントサイズ）だけ後でまとめて掛ける
    const merges = [];   // タイトル行
    const borders = [];  // 月ブロックの外枠
    const numFmts = [];  // 日付列の数値書式

    formatRules.forEach(rule => {
      const ri = rule.row;      // 0-based（グリッド添字）
      const r = ri + 1;         // 1-based（シート行）
      if (rule.type === 'title') {
        for (let c = 0; c < rule.cols; c++) { bgs[ri][c] = '#F9E400'; has[ri][c] = 'center'; fws[ri][c] = 'bold'; }
        merges.push({ row: r, cols: rule.cols });
      } else if (rule.type === 'header') {
        for (let c = 0; c < rule.cols; c++) { bgs[ri][c] = '#CCCCCC'; has[ri][c] = 'center'; fws[ri][c] = 'bold'; }
        // 日付の数字(1〜31)が日付シリアルとして解釈されないよう、整数書式を明示
        numFmts.push({ row: r, startCol: 2, numCols: rule.daysInMonth, fmt: '0' });
        for (let d = 1; d <= rule.daysInMonth; d++) {
          const dow = new Date(rule.year, rule.month, d).getDay();
          if (dow === 0) fcs[ri][d] = '#CC0000';
          else if (dow === 6) fcs[ri][d] = '#0000CC';
        }
      } else if (rule.type === 'empty_data') {
        fcs[ri][0] = '#999999';
      } else if (rule.type === 'data') {
        const base = rule.index % 2 === 0 ? '#FFFFFF' : '#F0FFF0';
        for (let c = 0; c < rule.cols; c++) bgs[ri][c] = base;
        fws[ri][0] = 'bold';
        numFmts.push({ row: r, startCol: 2, numCols: rule.cols - 1, fmt: '0.##' });
        for (let d = 1; d <= rule.daysInMonth; d++) {
          const dow = new Date(rule.year, rule.month, d).getDay();
          has[ri][d] = 'center';
          const val = outputData[ri][d];
          if (val === 0) fcs[ri][d] = '#CCCCCC';
          if (dow === 0) bgs[ri][d] = '#FFE6E6';
          else if (dow === 6) bgs[ri][d] = '#E6E6FF';
        }
        has[ri][rule.cols - 1] = 'center';
        fws[ri][rule.cols - 1] = 'bold';
      } else if (rule.type === 'total') {
        for (let c = 0; c < rule.cols; c++) { bgs[ri][c] = '#FFF9C4'; fws[ri][c] = 'bold'; }
        numFmts.push({ row: r, startCol: 2, numCols: rule.cols - 1, fmt: '0.##' });
        for (let d = 1; d <= rule.daysInMonth; d++) has[ri][d] = 'center';
        has[ri][rule.cols - 1] = 'center';
        borders.push({ startRow: r - rule.namesLength - 1, numRows: rule.namesLength + 2, cols: rule.cols });
      }
    });

    // 一括流し込み（それぞれ1回の通信で全行に適用）
    const fullRange = sheet.getRange(1, 1, numRows, maxCols);
    fullRange.setBackgrounds(bgs);
    fullRange.setFontColors(fcs);
    fullRange.setHorizontalAlignments(has);
    fullRange.setFontWeights(fws);
    // 数値書式: 連続行×同一書式はまとめて1回で適用（人数分の個別呼び出しを月ごと1回に圧縮）
    numFmts.sort((a, b) => a.row - b.row || a.startCol - b.startCol);
    let nfRun = null;
    const flushNfRun_ = () => { if (nfRun) sheet.getRange(nfRun.row, nfRun.startCol, nfRun.rows, nfRun.numCols).setNumberFormat(nfRun.fmt); };
    numFmts.forEach(n => {
      if (nfRun && n.row === nfRun.row + nfRun.rows && n.startCol === nfRun.startCol && n.numCols === nfRun.numCols && n.fmt === nfRun.fmt) { nfRun.rows++; }
      else { flushNfRun_(); nfRun = { row: n.row, startCol: n.startCol, rows: 1, numCols: n.numCols, fmt: n.fmt }; }
    });
    flushNfRun_();
    merges.forEach(m => sheet.getRange(m.row, 1, 1, m.cols).merge().setFontSize(13));
    borders.forEach(b => sheet.getRange(b.startRow, 1, b.numRows, b.cols).setBorder(true, true, true, true, true, true));

    // 合計セルを SUM 関数に置換（Excel で値を編集すると合計が自動更新される）
    // ※ data 行の最右列 + total 行の各日列 + total 行の最右列
    formatRules.forEach(rule => {
      if (rule.type !== 'data' && rule.type !== 'total') return;
      const r = rule.row + 1;                              // 1-based
      const days = rule.daysInMonth;
      if (!days) return;
      const firstDayCol = colLetter_(2);                    // B
      const lastDayCol = colLetter_(1 + days);              // 例: AF (31日なら)
      const totalCol = rule.cols;                           // 数値 = days + 2
      if (rule.type === 'data') {
        // 行の右端 = SUM(B行:lastDayCol行)
        sheet.getRange(r, totalCol).setFormula(`=SUM(${firstDayCol}${r}:${lastDayCol}${r})`);
      } else {
        // 合計行：各日列 = SUM(日列<dataStart>:<dataEnd>)、右端 = SUM(行内全日)
        const dataStartRow = r - rule.namesLength;
        const dataEndRow = r - 1;
        const totalFormulas = [[]];
        for (let i = 0; i < days; i++) {
          const col = colLetter_(2 + i);
          totalFormulas[0].push(`=SUM(${col}${dataStartRow}:${col}${dataEndRow})`);
        }
        totalFormulas[0].push(`=SUM(${firstDayCol}${r}:${lastDayCol}${r})`);
        sheet.getRange(r, 2, 1, days + 1).setFormulas(totalFormulas);
      }
    });
  }
}

function generateBillingSummary_(ss, records) {
  let sheet = ss.getSheetByName(BILLING_SHEET);
  if (sheet) { sheet.clear(); sheet.clearFormats(); } else { sheet = ss.insertSheet(BILLING_SHEET); }
  const W = 35; // 3列(会社/現場/名前) + 31日 + 合計
  ensureColumns_(sheet, W);
  // 倉庫は元請に請求しない作業のため除外（旧データで元請名が入っているものも対象外にする）
  const workRecords = records.filter(r => r.yakin !== '休み' && r.yakin !== '予定' && r.yakin !== '倉庫');
  const months = [...new Set(workRecords.map(r => r.month).filter(Boolean))].sort().reverse();
  const genbas = [...new Set(workRecords.map(r => r.genba).filter(Boolean))].sort();
  const DOW = ['日','月','火','水','木','金','土'];

  // 旧実装は1行ごとに setValues＋書式（約5〜8通信/行）で、行数に比例して遅くなり
  // 接続エラーの残存リスクだった。値・書式を全てメモリ上で組み立て、最後に
  // setValues＋一括書式（各1通信）で流し込む方式に統一（generateKakuninTable_ と同型）。
  const rows = [], bgs = [], fcs = [], has = [], fws = [], wraps = [], vas = [];
  const titleMerges = [];  // {row, cols} 月タイトル行の横結合
  const blockMerges = [];  // {row, numRows} 会社名/現場名セルの縦結合
  function addRow_(vals) {
    rows.push(vals.concat(Array(W - vals.length).fill('')));
    bgs.push(new Array(W).fill('#FFFFFF'));
    fcs.push(new Array(W).fill('#000000'));
    has.push(new Array(W).fill('left'));
    fws.push(new Array(W).fill('normal'));
    wraps.push(new Array(W).fill(false));
    vas.push(new Array(W).fill('bottom'));
    return rows.length - 1;
  }

  months.forEach(month => {
    const parts = month.split('-');
    const year = Number(parts[0]);
    const mon = Number(parts[1]);
    const monthLabel = year + '年' + mon + '月';
    const daysInMonth = new Date(year, mon, 0).getDate();
    const totalCols = 3 + daysInMonth + 1;
    const mr = workRecords.filter(r => r.month === month);
    // (氏名, 日付, 昼夜区分) → 行った現場のSet。1日に複数現場行ったら 1/N で按分する
    const sitesByPDN = {};
    mr.forEach(r => {
      const dn = r.yakin === '夜勤' ? 'N' : 'D';
      const k = r.name + '|' + r.date + '|' + dn;
      if (!sitesByPDN[k]) sitesByPDN[k] = new Set();
      sitesByPDN[k].add(r.genba + '|||' + (r.loc || '（現場名なし）'));
    });

    // 月タイトル行
    let ri = addRow_(['▶ ' + monthLabel]);
    for (let c = 0; c < totalCols; c++) { bgs[ri][c] = '#1D9E75'; fcs[ri][c] = '#FFFFFF'; fws[ri][c] = 'bold'; }
    titleMerges.push({ row: ri + 1, cols: totalCols });

    // ヘッダー行（曜日つき日付ラベル・土日は文字色）
    const headerRow = ['会社名', '現場名', '名前'];
    for (let d = 1; d <= daysInMonth; d++) { const dow = new Date(year, mon - 1, d).getDay(); headerRow.push(d + ' ' + DOW[dow]); }
    headerRow.push('合計');
    ri = addRow_(headerRow);
    for (let c = 0; c < totalCols; c++) { bgs[ri][c] = '#CCCCCC'; fws[ri][c] = 'bold'; has[ri][c] = 'center'; wraps[ri][c] = true; }
    for (let d = 1; d <= daysInMonth; d++) {
      const dow = new Date(year, mon - 1, d).getDay();
      if (dow === 0) fcs[ri][2 + d] = '#CC0000';
      else if (dow === 6) fcs[ri][2 + d] = '#0000CC';
    }

    genbas.forEach(genba => {
      const gr = mr.filter(r => r.genba === genba);
      if (gr.length === 0) return;
      const locs = [...new Set(gr.map(r => r.loc || '（現場名なし）'))].sort();
      locs.forEach(loc => {
        const lr = gr.filter(r => (r.loc || '（現場名なし）') === loc);
        const namesInLoc = [...new Set(lr.map(r => r.name))].sort();
        const activeNames = namesInLoc.filter(name => calcEffective_(lr, name).kosu > 0);
        if (activeNames.length === 0) return;
        const blockStartRi = rows.length;
        activeNames.forEach((name, ni) => {
          const row = [ni === 0 ? genba : '', ni === 0 ? loc : '', name];
          let rowTotal = 0;
          for (let d = 1; d <= daysInMonth; d++) {
            const dateStr = year + '-' + String(mon).padStart(2,'0') + '-' + String(d).padStart(2,'0');
            const dayRecs = lr.filter(r => r.name === name && r.date === dateStr);
            // 昼/夜勤の有無を判定し、行った現場数で1人工を按分（昼と夜勤は別カウント）
            const hasDay = dayRecs.some(r => r.yakin !== '夜勤');
            const hasNight = dayRecs.some(r => r.yakin === '夜勤');
            let kosu = 0;
            if (hasDay) {
              const sCnt = (sitesByPDN[name + '|' + dateStr + '|D'] || new Set()).size || 1;
              kosu += 1 / sCnt;
            }
            if (hasNight) {
              const sCnt = (sitesByPDN[name + '|' + dateStr + '|N'] || new Set()).size || 1;
              kosu += 1 / sCnt;
            }
            row.push(kosu > 0 ? kosu : 0);
            rowTotal += kosu;
          }
          row.push(rowTotal);
          ri = addRow_(row);
          const bg = ni % 2 === 0 ? '#FFFFFF' : '#F0FFF0';
          for (let c = 0; c < totalCols; c++) bgs[ri][c] = bg;
          fws[ri][0] = 'bold';
          for (let d = 1; d <= daysInMonth; d++) {
            const dow = new Date(year, mon - 1, d).getDay();
            const gi = 2 + d; // 日付dのシート列は3+d → グリッド添字は2+d
            has[ri][gi] = 'center';
            if (row[2 + d] === 0) fcs[ri][gi] = '#CCCCCC';
            if (dow === 0) bgs[ri][gi] = '#FFE6E6';
            else if (dow === 6) bgs[ri][gi] = '#E6E6FF';
          }
          fws[ri][totalCols - 1] = 'bold';
          has[ri][totalCols - 1] = 'center';
        });
        const totalRow = ['', '', '合計'];
        let grandTotal = 0;
        for (let d = 1; d <= daysInMonth; d++) { const dateStr = year + '-' + String(mon).padStart(2,'0') + '-' + String(d).padStart(2,'0'); let daySum = 0; activeNames.forEach(name => { const dayRecs = lr.filter(r => r.name === name && r.date === dateStr); const hasDay = dayRecs.some(r => r.yakin !== '夜勤'); const hasNight = dayRecs.some(r => r.yakin === '夜勤'); if (hasDay) { const sCnt = (sitesByPDN[name + '|' + dateStr + '|D'] || new Set()).size || 1; daySum += 1 / sCnt; } if (hasNight) { const sCnt = (sitesByPDN[name + '|' + dateStr + '|N'] || new Set()).size || 1; daySum += 1 / sCnt; } }); totalRow.push(daySum > 0 ? daySum : 0); grandTotal += daySum; }
        totalRow.push(grandTotal);
        ri = addRow_(totalRow);
        for (let c = 0; c < totalCols; c++) { bgs[ri][c] = '#FFF9C4'; fws[ri][c] = 'bold'; }
        for (let d = 1; d <= daysInMonth; d++) has[ri][2 + d] = 'center';
        has[ri][totalCols - 1] = 'center';
        // ブロック先頭の会社名・現場名セル: 太字＋縦中央、複数人なら縦結合
        if (activeNames.length > 1) blockMerges.push({ row: blockStartRi + 1, numRows: activeNames.length });
        fws[blockStartRi][0] = 'bold'; vas[blockStartRi][0] = 'middle';
        fws[blockStartRi][1] = 'bold'; vas[blockStartRi][1] = 'middle';
        addRow_([]); // ブロック間の空行
      });
    });
    addRow_([]); // 月間の空行
  });

  if (rows.length > 0) {
    const n = rows.length;
    const rng = sheet.getRange(1, 1, n, W);
    rng.setValues(rows);
    rng.setBackgrounds(bgs);
    rng.setFontColors(fcs);
    rng.setHorizontalAlignments(has);
    rng.setFontWeights(fws);
    rng.setWraps(wraps);
    rng.setVerticalAlignments(vas);
    // 日付＋合計列の数値書式（テキストのヘッダー/タイトル行に掛かっても表示は変わらない）
    sheet.getRange(1, 4, n, W - 3).setNumberFormat('0.##');
    titleMerges.forEach(m => sheet.getRange(m.row, 1, 1, m.cols).merge().setFontSize(12));
    blockMerges.forEach(m => { sheet.getRange(m.row, 1, m.numRows, 1).merge(); sheet.getRange(m.row, 2, m.numRows, 1).merge(); });
    sheet.getRange(1, 1, n, W).setBorder(true, true, true, true, true, true, '#DDDDDD', SpreadsheetApp.BorderStyle.SOLID);
  }
  sheet.setColumnWidth(1, 140); sheet.setColumnWidth(2, 180); sheet.setColumnWidth(3, 80);
  sheet.setColumnWidths(4, W - 3, 26);
  if (sheet.getMaxColumns() >= 36) sheet.setColumnWidth(36, 50);
}

// 元請別請求集計の「フィルタ用」シート（フラット構造）
// - マージなし・空白行なし → AutoFilter が全範囲で機能する
// - 列構成: 月(テキスト) / 会社名 / 現場名 / 名前(or 合計) / 1日〜31日 / 合計（計36列）
// - 月をまたいで会社名・現場名でフィルタを掛けて集計できる
// - 案件サブ合計行も同居（名前列=「合計」）
function generateBillingFilterSheet_(ss, records) {
  let sheet = ss.getSheetByName(BILLING_FILTER_SHEET);
  if (sheet) { sheet.clear(); sheet.clearFormats(); }
  else sheet = ss.insertSheet(BILLING_FILTER_SHEET);
  // 既存フィルタを除去（再生成時の衝突防止）
  try { const f = sheet.getFilter(); if (f) f.remove(); } catch (e) {}
  ensureColumns_(sheet, 36);

  // 休み/予定は除外。倉庫は genba='', loc='倉庫作業' に正規化してフィルタ用シートに含める
  // （通常の請求集計シート generateBillingSummary_ は変更なし — 倉庫除外のまま）
  const workRecords = records
    .filter(r => r.yakin !== '休み' && r.yakin !== '予定')
    .map(r => r.yakin === '倉庫'
      ? Object.assign({}, r, { genba: '', loc: '倉庫作業' })
      : r);
  const months = [...new Set(workRecords.map(r => r.month).filter(Boolean))].sort().reverse();
  const genbas = [...new Set(workRecords.map(r => r.genba).filter(Boolean))].sort();
  // 倉庫作業が存在する場合は空 genba ブロックを末尾に追加（倉庫は元請名が空）
  if (workRecords.some(r => r.loc === '倉庫作業' && !r.genba)) genbas.push('');

  // ヘッダー行
  const header = ['月', '会社名', '現場名', '名前'];
  for (let d = 1; d <= 31; d++) header.push(d + '日');
  header.push('合計');

  const rows = [header];
  // 合計行のインデックス（後で背景色つけるため）
  const totalRowIndices = [];

  months.forEach(month => {
    const parts = month.split('-');
    const year = Number(parts[0]);
    const mon = Number(parts[1]);
    const monthLabel = year + '年' + mon + '月';
    const daysInMonth = new Date(year, mon, 0).getDate();
    const mr = workRecords.filter(r => r.month === month);

    // 1日に複数現場行ったら 1/N で按分（昼/夜勤は別カウント）
    const sitesByPDN = {};
    mr.forEach(r => {
      const dn = r.yakin === '夜勤' ? 'N' : 'D';
      const k = r.name + '|' + r.date + '|' + dn;
      if (!sitesByPDN[k]) sitesByPDN[k] = new Set();
      sitesByPDN[k].add(r.genba + '|||' + (r.loc || '（現場名なし）'));
    });

    genbas.forEach(genba => {
      const gr = mr.filter(r => r.genba === genba);
      if (gr.length === 0) return;
      const locs = [...new Set(gr.map(r => r.loc || '（現場名なし）'))].sort();
      locs.forEach(loc => {
        const lr = gr.filter(r => (r.loc || '（現場名なし）') === loc);
        const namesInLoc = [...new Set(lr.map(r => r.name))].sort();
        const activeNames = namesInLoc.filter(name => calcEffective_(lr, name).kosu > 0);
        if (activeNames.length === 0) return;

        // 各人の行
        activeNames.forEach(name => {
          const row = [monthLabel, genba, loc, name];
          let rowTotal = 0;
          for (let d = 1; d <= 31; d++) {
            if (d > daysInMonth) { row.push(''); continue; }
            const dateStr = year + '-' + String(mon).padStart(2,'0') + '-' + String(d).padStart(2,'0');
            const dayRecs = lr.filter(r => r.name === name && r.date === dateStr);
            const hasDay = dayRecs.some(r => r.yakin !== '夜勤');
            const hasNight = dayRecs.some(r => r.yakin === '夜勤');
            let kosu = 0;
            if (hasDay) {
              const sCnt = (sitesByPDN[name + '|' + dateStr + '|D'] || new Set()).size || 1;
              kosu += 1 / sCnt;
            }
            if (hasNight) {
              const sCnt = (sitesByPDN[name + '|' + dateStr + '|N'] || new Set()).size || 1;
              kosu += 1 / sCnt;
            }
            row.push(kosu);
            rowTotal += kosu;
          }
          row.push(rowTotal);
          rows.push(row);
        });

        // 案件の合計行
        const totalRow = [monthLabel, genba, loc, '合計'];
        let grandTotal = 0;
        for (let d = 1; d <= 31; d++) {
          if (d > daysInMonth) { totalRow.push(''); continue; }
          const dateStr = year + '-' + String(mon).padStart(2,'0') + '-' + String(d).padStart(2,'0');
          let daySum = 0;
          activeNames.forEach(name => {
            const dayRecs = lr.filter(r => r.name === name && r.date === dateStr);
            const hasDay = dayRecs.some(r => r.yakin !== '夜勤');
            const hasNight = dayRecs.some(r => r.yakin === '夜勤');
            if (hasDay) {
              const sCnt = (sitesByPDN[name + '|' + dateStr + '|D'] || new Set()).size || 1;
              daySum += 1 / sCnt;
            }
            if (hasNight) {
              const sCnt = (sitesByPDN[name + '|' + dateStr + '|N'] || new Set()).size || 1;
              daySum += 1 / sCnt;
            }
          });
          totalRow.push(daySum);
          grandTotal += daySum;
        }
        totalRow.push(grandTotal);
        rows.push(totalRow);
        totalRowIndices.push(rows.length); // 1-based 合計行の最終 index
      });
    });
  });

  if (rows.length === 1) {
    // データなし: ヘッダーだけ書いて終了
    sheet.getRange(1, 1, 1, 36).setValues(rows)
      .setFontWeight('bold').setBackground('#E8F4FD').setHorizontalAlignment('center');
    return;
  }

  // 一括書き込み
  sheet.getRange(1, 1, rows.length, 36).setValues(rows);

  // ヘッダー＋合計行の背景・太字は全行グリッドで一括適用（以前は合計行ごとに2通信だった）
  {
    const n = rows.length;
    const bgsF = [], fwsF = [];
    for (let i = 0; i < n; i++) { bgsF.push(new Array(36).fill('#FFFFFF')); fwsF.push(new Array(36).fill('normal')); }
    for (let c = 0; c < 36; c++) { bgsF[0][c] = '#E8F4FD'; fwsF[0][c] = 'bold'; }
    totalRowIndices.forEach(r => { for (let c = 0; c < 36; c++) { bgsF[r - 1][c] = '#FFF9C4'; fwsF[r - 1][c] = 'bold'; } });
    sheet.getRange(1, 1, n, 36).setBackgrounds(bgsF).setFontWeights(fwsF);
  }
  sheet.getRange(1, 1, 1, 36).setHorizontalAlignment('center');

  // 月列をテキスト書式（Excel での日付シリアル化防止）
  sheet.getRange(2, 1, rows.length - 1, 1).setNumberFormat('@');

  // 日付列・合計列の数値書式（0は非表示）
  sheet.getRange(2, 5, rows.length - 1, 32).setNumberFormat('0.0;-0.0;').setHorizontalAlignment('center');

  // 列幅
  sheet.setColumnWidth(1, 80);    // 月
  sheet.setColumnWidth(2, 150);   // 会社名
  sheet.setColumnWidth(3, 220);   // 現場名
  sheet.setColumnWidth(4, 70);    // 名前
  sheet.setColumnWidths(5, 31, 32);  // 1〜31日
  sheet.setColumnWidth(36, 60);   // 合計

  // フリーズ（1行＋4列）
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(4);

  // AutoFilter（全範囲）
  try {
    sheet.getRange(1, 1, rows.length, 36).createFilter();
  } catch (e) { /* createFilter は既存フィルタがあれば失敗するが、上で除去済み */ }

  // 罫線
  sheet.getRange(1, 1, rows.length, 36)
    .setBorder(true, true, true, true, true, true, '#DDDDDD', SpreadsheetApp.BorderStyle.SOLID);
}

function generateDivisionAllocation_(ss, records) {
  let sheet = ss.getSheetByName(ALLOCATION_SHEET);
  if (sheet) { sheet.clear(); sheet.clearFormats(); } else { sheet = ss.insertSheet(ALLOCATION_SHEET); }

  const memberSheet = getOrCreateMemberSheet_(ss);
  const memberData = memberSheet.getDataRange().getValues();
  const memberDivision = {};
  const memberRate = {};
  for (let i = 1; i < memberData.length; i++) {
    const name = String(memberData[i][0] || '').trim();
    const div = String(memberData[i][2] || '').trim();
    const rate = Number(memberData[i][3] || 0);
    if (!name) continue;
    // 同名の重複行がある場合、非空の事業部を優先（空欄で上書きされないようにする）
    if (memberDivision[name] === undefined || (div && !memberDivision[name])) memberDivision[name] = div;
    if (rate) memberRate[name] = rate;
  }

  const jobSiteSheet = getOrCreateJobSiteSheet_(ss);
  const jobSiteData = jobSiteSheet.getDataRange().getValues();
  const siteJobNo = {};
  const siteInfo = {};
  const siteRevenue = {};
  const siteDivision = {};
  for (let i = 1; i < jobSiteData.length; i++) {
    const genba = String(jobSiteData[i][0] || '').trim();
    const loc = String(jobSiteData[i][1] || '').trim();
    const jobNo = String(jobSiteData[i][2] || '').trim();
    const divV = String(jobSiteData[i][3] || '').trim();
    const revenue = Number(jobSiteData[i][6] || 0);
    if (genba && jobNo) {
      siteJobNo[genba + '|||' + loc] = jobNo;
      siteInfo[jobNo] = { genba, loc };
      siteRevenue[jobNo] = revenue;
      siteDivision[jobNo] = divV;
    }
  }

  // 倉庫は工番なしのため事業部按分の対象外。旧データで工番マスタにヒットしてしまうケースを防ぐため明示的に除外
  const grRecords = records.filter(r => r.company === GROWISE && r.yakin !== '休み' && r.yakin !== '予定' && r.yakin !== '倉庫');

  // (氏名, 日付, 昼夜区分) → 行った jobNo のSet。1日に複数現場行ったら 1/N で按分する
  const jobsByPDN = {};
  grRecords.forEach(r => {
    const jobNo = siteJobNo[r.genba + '|||' + r.loc];
    if (!jobNo) return;
    const nf = r.yakin === '夜勤' ? 'N' : 'D';
    const k = r.name + '|' + r.date + '|' + nf;
    if (!jobsByPDN[k]) jobsByPDN[k] = new Set();
    jobsByPDN[k].add(jobNo);
  });

  const byKey = {};
  grRecords.forEach(r => {
    const jobNo = siteJobNo[r.genba + '|||' + r.loc];
    if (!jobNo) return;
    const nightFlag = r.yakin === '夜勤' ? 'N' : 'D';
    const pdnKey = r.name + '|' + r.date + '|' + nightFlag;
    const jobCount = (jobsByPDN[pdnKey] || new Set()).size || 1;
    const sharedKosu = 1 / jobCount; // 1人工を行った現場数で按分
    const key = jobNo + '|' + r.name + '|' + r.date + '|' + nightFlag;
    // 同じ key（同 jobNo+昼夜）で複数レコードある場合は1度だけ計上（重複登録対策）
    if (!byKey[key]) {
      byKey[key] = { jobNo, name: r.name, date: r.date, month: r.month, kosu: sharedKosu };
    }
  });

  const kosuTotalByJob = {};
  const kosuMonthlyByJob = {};
  const costTotalByJob = {};
  const costMonthlyByJob = {};
  const allDivs = new Set();

  Object.values(byKey).forEach(v => {
    let div = memberDivision[v.name];
    if (!div) div = siteDivision[v.jobNo] || '';
    if (!div) div = '不明';
    const rate = memberRate[v.name] || 0;
    const cost = v.kosu * rate;
    allDivs.add(div);
    if (!kosuTotalByJob[v.jobNo]) kosuTotalByJob[v.jobNo] = {};
    kosuTotalByJob[v.jobNo][div] = (kosuTotalByJob[v.jobNo][div] || 0) + v.kosu;
    if (!costTotalByJob[v.jobNo]) costTotalByJob[v.jobNo] = {};
    costTotalByJob[v.jobNo][div] = (costTotalByJob[v.jobNo][div] || 0) + cost;
    if (!kosuMonthlyByJob[v.month]) { kosuMonthlyByJob[v.month] = {}; costMonthlyByJob[v.month] = {}; }
    if (!kosuMonthlyByJob[v.month][v.jobNo]) { kosuMonthlyByJob[v.month][v.jobNo] = {}; costMonthlyByJob[v.month][v.jobNo] = {}; }
    kosuMonthlyByJob[v.month][v.jobNo][div] = (kosuMonthlyByJob[v.month][v.jobNo][div] || 0) + v.kosu;
    costMonthlyByJob[v.month][v.jobNo][div] = (costMonthlyByJob[v.month][v.jobNo][div] || 0) + cost;
  });

  const DIVS_ORDER = ['ICT', 'INF', 'MSC', 'GRB'];
  const divs = DIVS_ORDER.filter(d => allDivs.has(d));
  [...allDivs].sort().forEach(d => { if (!divs.includes(d)) divs.push(d); });
  if (divs.length === 0) divs.push('ICT');

  // 列構成: 工番 | 元請名 | 現場名 | 売上 | [div人工] | 合計人工 | [div%] | [div原価] | 合計原価 | 粗利 | 粗利率
  const numCols = 4 + divs.length + 1 + divs.length + divs.length + 1 + 2;
  const blank = () => Array(numCols).fill('');
  const rows = [];
  const formats = [];

  // 按分%の計算: 工番事業部に50%固定 + 残り50%を稼働した事業部(工番事業部含む)で人工比按分
  // 工番事業部が稼働ゼロ→工番事業部100% / 工番事業部不明→従来通り100%稼働按分
  function calcAllocPercent_(kosuBreakdown, jobNoDiv) {
    const totalKosu = divs.reduce((s, d) => s + (kosuBreakdown[d] || 0), 0);
    const result = {};
    divs.forEach(d => result[d] = 0);
    const hasJobNoDiv = jobNoDiv && divs.includes(jobNoDiv);
    if (hasJobNoDiv) {
      result[jobNoDiv] = 50;
      if (totalKosu > 0) {
        divs.forEach(d => { result[d] += 50 * (kosuBreakdown[d] || 0) / totalKosu; });
      } else {
        result[jobNoDiv] += 50; // 稼働ゼロ→残り50%も工番事業部に
      }
    } else if (totalKosu > 0) {
      divs.forEach(d => { result[d] = 100 * (kosuBreakdown[d] || 0) / totalKosu; });
    }
    return result;
  }

  function buildHeader() {
    const h = ['工番', '元請名', '現場名', '売上'];
    divs.forEach(d => h.push(d + '人工'));
    h.push('合計人工');
    divs.forEach(d => h.push(d + '%'));
    divs.forEach(d => h.push(d + '原価'));
    h.push('合計原価');
    h.push('粗利');
    h.push('粗利率');
    return h;
  }
  function buildRow(jobNo, kosuBreakdown, costBreakdown, revenue, showRevenue) {
    const info = siteInfo[jobNo] || { genba: '', loc: '' };
    const jobNoDiv = siteDivision[jobNo] || '';
    const totalKosu = divs.reduce((s, d) => s + (kosuBreakdown[d] || 0), 0);
    const totalCost = divs.reduce((s, d) => s + (costBreakdown[d] || 0), 0);
    const alloc = calcAllocPercent_(kosuBreakdown, jobNoDiv);
    const row = [jobNo, info.genba, info.loc, showRevenue ? (revenue || 0) : ''];
    divs.forEach(d => row.push(kosuBreakdown[d] || 0));
    row.push(totalKosu);
    divs.forEach(d => row.push(Math.round((alloc[d] || 0) * 10) / 10 + '%'));
    divs.forEach(d => row.push(Math.round(costBreakdown[d] || 0)));
    row.push(Math.round(totalCost));
    if (showRevenue) {
      const profit = (revenue || 0) - totalCost;
      row.push(Math.round(profit));
      row.push(revenue > 0 ? Math.round(profit / revenue * 1000) / 10 + '%' : '');
    } else {
      row.push(''); row.push('');
    }
    return row;
  }

  const titleRow = blank();
  titleRow[0] = '事業部別按分';
  titleRow[numCols - 1] = '更新日時: ' + new Date().toLocaleString('ja-JP');
  rows.push(titleRow);
  formats.push({ row: rows.length, type: 'title' });
  rows.push(blank());

  // 全期間累計
  const totalSectionRow = blank(); totalSectionRow[0] = '▶ 全期間累計（売上・粗利を計上）';
  rows.push(totalSectionRow);
  formats.push({ row: rows.length, type: 'section_total' });
  rows.push(buildHeader());
  formats.push({ row: rows.length, type: 'header' });
  const totalJobs = Object.keys(kosuTotalByJob).sort();
  let gKosu = {}, gCost = {}, gRev = 0;
  totalJobs.forEach(jobNo => {
    const rev = siteRevenue[jobNo] || 0;
    rows.push(buildRow(jobNo, kosuTotalByJob[jobNo], costTotalByJob[jobNo] || {}, rev, true));
    gRev += rev;
    divs.forEach(d => {
      gKosu[d] = (gKosu[d] || 0) + (kosuTotalByJob[jobNo][d] || 0);
      gCost[d] = (gCost[d] || 0) + ((costTotalByJob[jobNo] || {})[d] || 0);
    });
  });
  if (totalJobs.length > 0) {
    const totalKosu = divs.reduce((s, d) => s + (gKosu[d] || 0), 0);
    const totalCost = divs.reduce((s, d) => s + (gCost[d] || 0), 0);
    const profit = gRev - totalCost;
    const row = ['合計', '', '', gRev];
    divs.forEach(d => row.push(gKosu[d] || 0));
    row.push(totalKosu);
    divs.forEach(d => row.push(totalKosu > 0 ? Math.round((gKosu[d] || 0) / totalKosu * 1000) / 10 + '%' : '0%'));
    divs.forEach(d => row.push(Math.round(gCost[d] || 0)));
    row.push(Math.round(totalCost));
    row.push(Math.round(profit));
    row.push(gRev > 0 ? Math.round(profit / gRev * 1000) / 10 + '%' : '');
    rows.push(row);
    formats.push({ row: rows.length, type: 'total' });
  }
  rows.push(blank());

  // 月別
  const months = Object.keys(kosuMonthlyByJob).sort().reverse();
  months.forEach(month => {
    const parts = month.split('-');
    const label = parts[0] + '年' + Number(parts[1]) + '月';
    const sec = blank(); sec[0] = '▶ ' + label + '（月別人工・原価。売上は全期間のみ）';
    rows.push(sec);
    formats.push({ row: rows.length, type: 'section_month' });
    rows.push(buildHeader());
    formats.push({ row: rows.length, type: 'header' });
    const jobs = Object.keys(kosuMonthlyByJob[month]).sort();
    let mKosu = {}, mCost = {};
    jobs.forEach(jobNo => {
      rows.push(buildRow(jobNo, kosuMonthlyByJob[month][jobNo], costMonthlyByJob[month][jobNo] || {}, 0, false));
      divs.forEach(d => {
        mKosu[d] = (mKosu[d] || 0) + (kosuMonthlyByJob[month][jobNo][d] || 0);
        mCost[d] = (mCost[d] || 0) + ((costMonthlyByJob[month][jobNo] || {})[d] || 0);
      });
    });
    if (jobs.length > 0) {
      const totalKosu = divs.reduce((s, d) => s + (mKosu[d] || 0), 0);
      const totalCost = divs.reduce((s, d) => s + (mCost[d] || 0), 0);
      const row = ['合計', '', '', ''];
      divs.forEach(d => row.push(mKosu[d] || 0));
      row.push(totalKosu);
      divs.forEach(d => row.push(totalKosu > 0 ? Math.round((mKosu[d] || 0) / totalKosu * 1000) / 10 + '%' : '0%'));
      divs.forEach(d => row.push(Math.round(mCost[d] || 0)));
      row.push(Math.round(totalCost));
      row.push(''); row.push('');
      rows.push(row);
      formats.push({ row: rows.length, type: 'total' });
    }
    rows.push(blank());
  });

  if (rows.length > 0) {
    ensureColumns_(sheet, numCols);
    sheet.getRange(1, 1, rows.length, numCols).setValues(rows);
    // 書式は行×列グリッドで一括適用（以前は1行ずつ setBackground/setFontWeight を呼んでいた）
    {
      const n = rows.length;
      const bgsA = [], fwsA = [], hasA = [];
      for (let i = 0; i < n; i++) { bgsA.push(new Array(numCols).fill('#FFFFFF')); fwsA.push(new Array(numCols).fill('normal')); hasA.push(new Array(numCols).fill('left')); }
      const fontSizes = [];
      formats.forEach(f => {
        const ri = f.row - 1;
        if (f.type === 'title') { fwsA[ri][0] = 'bold'; fontSizes.push({ row: f.row, size: 14 }); }
        else if (f.type === 'section_total') { for (let c = 0; c < numCols; c++) bgsA[ri][c] = '#E8F5E9'; fwsA[ri][0] = 'bold'; fontSizes.push({ row: f.row, size: 12 }); }
        else if (f.type === 'section_month') { for (let c = 0; c < numCols; c++) bgsA[ri][c] = '#E3F2FD'; fwsA[ri][0] = 'bold'; fontSizes.push({ row: f.row, size: 12 }); }
        else if (f.type === 'header') { for (let c = 0; c < numCols; c++) { bgsA[ri][c] = '#F5F5F5'; fwsA[ri][c] = 'bold'; hasA[ri][c] = 'center'; } }
        else if (f.type === 'total') { for (let c = 0; c < numCols; c++) { bgsA[ri][c] = '#FFF9C4'; fwsA[ri][c] = 'bold'; } }
      });
      sheet.getRange(1, 1, n, numCols).setBackgrounds(bgsA).setFontWeights(fwsA).setHorizontalAlignments(hasA);
      fontSizes.forEach(fs => sheet.getRange(fs.row, 1).setFontSize(fs.size));
    }
    // 金額列に通貨書式
    const dataStartRow = 3;
    const dataEndRow = rows.length;
    if (dataEndRow >= dataStartRow) {
      const numRows = dataEndRow - dataStartRow + 1;
      // 売上
      sheet.getRange(dataStartRow, 4, numRows, 1).setNumberFormat('¥#,##0');
      // 人工列 (div人工 + 合計人工) を強制的に通常数値書式に（残存¥や%書式を排除）
      sheet.getRange(dataStartRow, 5, numRows, divs.length + 1).setNumberFormat('0.##');
      // % 列 (文字列 "48%" 等で格納)。書式を一般にして文字列をそのまま表示
      sheet.getRange(dataStartRow, 5 + divs.length + 1, numRows, divs.length).setNumberFormat('@');
      // 原価列 (div原価 + 合計原価)
      const costStart = 4 + divs.length + 1 + divs.length + 1;
      sheet.getRange(dataStartRow, costStart, numRows, divs.length + 1).setNumberFormat('¥#,##0');
      // 粗利
      sheet.getRange(dataStartRow, costStart + divs.length + 1, numRows, 1).setNumberFormat('¥#,##0');
      // 粗利率 (文字列)
      sheet.getRange(dataStartRow, costStart + divs.length + 2, numRows, 1).setNumberFormat('@');
    }
  }

  sheet.setColumnWidth(1, 110);
  sheet.setColumnWidth(2, 140);
  sheet.setColumnWidth(3, 160);
  sheet.setColumnWidth(4, 110);
  sheet.setColumnWidths(5, divs.length, 70);
  sheet.setColumnWidth(5 + divs.length, 80);
  sheet.setColumnWidths(6 + divs.length, divs.length, 60);
  sheet.setColumnWidths(6 + divs.length * 2, divs.length, 90);
  sheet.setColumnWidth(6 + divs.length * 3, 100);
  sheet.setColumnWidth(7 + divs.length * 3, 100);
  sheet.setColumnWidth(8 + divs.length * 3, 70);
}

function dailySummary() { generateSummary_(); }

// 元請名を「from」から「to」に統合（日報・アーカイブ・現場マスタ・元請マスタを全部書き換え）
function mergeGenba_(ss, fromName, toName) {
  const result = { nippoUpdated: 0, archiveUpdated: 0, jobsiteUpdated: 0, masterAction: 'none' };
  // 日報データ / アーカイブ
  [SHEET_NAME, ARCHIVE_SHEET].forEach((name, idx) => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return;
    const headers = data[0];
    const gCol = headers.indexOf('元請名');
    if (gCol < 0) return;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][gCol] || '').trim() === fromName) {
        sheet.getRange(i + 1, gCol + 1).setValue(toName);
        if (idx === 0) result.nippoUpdated++; else result.archiveUpdated++;
      }
    }
  });
  // 現場マスタ
  const jobSite = ss.getSheetByName(JOBSITE_SHEET);
  if (jobSite) {
    const data = jobSite.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').trim() === fromName) {
        jobSite.getRange(i + 1, 1).setValue(toName);
        result.jobsiteUpdated++;
      }
    }
  }
  // 元請マスタ
  const genbaSheet = ss.getSheetByName(GENBA_MASTER_SHEET);
  if (genbaSheet) {
    const data = genbaSheet.getDataRange().getValues();
    let fromRow = -1;
    let toExists = false;
    for (let i = 1; i < data.length; i++) {
      const n = String(data[i][0] || '').trim();
      if (n === fromName && fromRow < 0) fromRow = i;
      if (n === toName) toExists = true;
    }
    if (fromRow >= 0) {
      if (toExists) {
        genbaSheet.deleteRow(fromRow + 1);
        result.masterAction = 'deleted_duplicate';
      } else {
        genbaSheet.getRange(fromRow + 1, 1).setValue(toName);
        result.masterAction = 'renamed';
      }
    } else {
      result.masterAction = 'from_not_found';
    }
  }
  return result;
}

// 同じ元請内で「現場名」を fromLoc から toLoc に統合する
// 日報・アーカイブの現場名を書き換え、現場マスタも整理（toLoc 行があれば from を削除、なければ from を改名）
// 統合先の現場マスタに工番がある場合は日報側の工番もそちらに統一する
function mergeLoc_(ss, genba, fromLoc, toLoc) {
  const result = { nippoUpdated: 0, archiveUpdated: 0, masterAction: 'none', toJobNo: '' };

  // 1. 現場マスタを先に調べる（統合先の工番取得 + from/to 行の位置確認）
  const jobSite = ss.getSheetByName(JOBSITE_SHEET);
  let toJobNo = '';
  let fromRowIdx = -1;
  let toRowIdx = -1;
  if (jobSite) {
    const jData = jobSite.getDataRange().getValues();
    for (let i = 1; i < jData.length; i++) {
      const g = String(jData[i][0] || '').trim();
      const l = String(jData[i][1] || '').trim();
      if (g !== genba) continue;
      if (l === toLoc) { toRowIdx = i; toJobNo = String(jData[i][2] || ''); }
      if (l === fromLoc) { fromRowIdx = i; }
    }
  }
  result.toJobNo = toJobNo;

  // 2. 日報データ・アーカイブを書き換え（現場名 + 任意で工番）
  [SHEET_NAME, ARCHIVE_SHEET].forEach((name, idx) => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return;
    const headers = data[0];
    const gCol = headers.indexOf('元請名');
    const lCol = headers.indexOf('現場名');
    const jCol = headers.indexOf('工番');
    if (gCol < 0 || lCol < 0) return;
    let count = 0;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][gCol] || '').trim() !== genba) continue;
      if (String(data[i][lCol] || '').trim() !== fromLoc) continue;
      sheet.getRange(i + 1, lCol + 1).setValue(toLoc);
      if (toJobNo && jCol >= 0) sheet.getRange(i + 1, jCol + 1).setValue(toJobNo);
      count++;
    }
    if (idx === 0) result.nippoUpdated = count; else result.archiveUpdated = count;
  });

  // 3. 現場マスタを整理
  if (jobSite) {
    if (fromRowIdx > 0 && toRowIdx > 0) {
      // 両方ある → from 行を削除（売上などは to 側を残す）
      jobSite.deleteRow(fromRowIdx + 1);
      result.masterAction = 'deleted_duplicate';
    } else if (fromRowIdx > 0) {
      // from のみある → 現場名を toLoc に改名
      jobSite.getRange(fromRowIdx + 1, 2).setValue(toLoc);
      result.masterAction = 'renamed';
    } else if (toRowIdx > 0) {
      // to のみある → 何もしない（日報側だけ書き換えた）
      result.masterAction = 'to_only';
    } else {
      result.masterAction = 'none';
    }
  }
  return result;
}

// 工番を持つべきでないレコードの工番・事業部をクリア:
// - 休み/倉庫/予定 モードのレコード
// - 作業区分が「現場作業」以外のレコード（材料引取・現調・カギ借用・撤去品返却・着打ち・その他）
// （旧仕様時代のデータ清掃用 / これ以降は新規発行時に正しく空のまま）
function cleanupOrphanJobNos_(ss) {
  let cleaned = 0;
  [SHEET_NAME, ARCHIVE_SHEET].forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return;
    const headers = data[0];
    const yCol = headers.indexOf('夜勤');
    const dCol = headers.indexOf('事業部');
    const jCol = headers.indexOf('工番');
    const wtCol = headers.indexOf('作業区分');
    if (yCol < 0 || jCol < 0) return;
    for (let i = 1; i < data.length; i++) {
      const yakin = String(data[i][yCol] || '').trim();
      const jobNo = String(data[i][jCol] || '').trim();
      const div = dCol >= 0 ? String(data[i][dCol] || '').trim() : '';
      const wt = wtCol >= 0 ? String(data[i][wtCol] || '').trim() : '';
      const isMode = (yakin === '休み' || yakin === '倉庫' || yakin === '予定');
      const isNonGenba = (wt && wt !== '現場作業');
      if ((isMode || isNonGenba) && (jobNo || div)) {
        if (jobNo) sheet.getRange(i + 1, jCol + 1).setValue('');
        if (div && dCol >= 0) sheet.getRange(i + 1, dCol + 1).setValue('');
        cleaned++;
      }
    }
  });
  return cleaned;
}

// ========== 読み(フリガナ)バックフィル ==========
// スクリプトエディタから手動実行用。
// 元請マスタ/現場マスタの既存行で「読み」が空欄の項目に対し、Groqで読みを生成して書き込む。
// 実行前に「スクリプトプロパティ」に GROQ_API_KEY を設定してください。
function backfillAllYomi() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Groq通信を含む長時間処理なので、日報用ScriptLockを占有しない。
  const lock = LockService.getUserLock();
  if (!lock.tryLock(60000)) { Logger.log('ロック取得失敗'); return; }
  try {
    const key = PropertiesService.getScriptProperties().getProperty('GROQ_API_KEY');
    if (!key) { Logger.log('GROQ_API_KEY が未設定です。スクリプトプロパティに登録してください。'); return; }

    // 元請マスタ: A=元請名, B=会社, C=読み
    const gSheet = getOrCreateGenbaSheet_(ss);
    const gResult = _backfillYomiInSheet_(gSheet, 0, 2, '元請マスタ');

    // 現場マスタ: A=元請名, B=現場名, C=工番, ..., H=読み
    const jSheet = getOrCreateJobSiteSheet_(ss);
    const jResult = _backfillYomiInSheet_(jSheet, 1, 7, '現場マスタ');

    const msg = `完了 | 元請: ${gResult.filled}/${gResult.target}件 | 現場: ${jResult.filled}/${jResult.target}件`;
    Logger.log(msg);
    try { logOperation_(ss, 'backfill_yomi', 'マスタ一括', msg, 'system'); } catch (e) {}
  } finally {
    lock.releaseLock();
  }
}

// 指定シートの textColIdx(0ベース) 列の値を読んで、yomiColIdx 列が空なら読みを埋める
function _backfillYomiInSheet_(sheet, textColIdx, yomiColIdx, label) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) { Logger.log('[' + label + '] データなし'); return { target: 0, filled: 0 }; }

  // 要生成リストを作る
  const todo = [];
  for (let i = 1; i < data.length; i++) {
    const text = String(data[i][textColIdx] || '').trim();
    const currentYomi = String(data[i][yomiColIdx] || '').trim();
    if (!text) continue;
    if (currentYomi) continue;                // 既に入っている分はスキップ(手動入力を優先)
    if (!needsYomi_(text)) continue;          // 漢字を含まないものはスキップ
    todo.push({ row: i + 1, text: text });
  }
  Logger.log('[' + label + '] 要生成: ' + todo.length + '件');
  if (!todo.length) return { target: 0, filled: 0 };

  // 30件ずつ Groq にバッチ問合せ
  const BATCH = 30;
  let filled = 0;
  for (let i = 0; i < todo.length; i += BATCH) {
    const chunk = todo.slice(i, i + BATCH);
    const texts = chunk.map(function(c){ return c.text; });
    const readings = fetchYomiFromGroq_(texts);
    for (let k = 0; k < chunk.length; k++) {
      const y = String((readings[k] || '')).trim();
      if (y) {
        sheet.getRange(chunk[k].row, yomiColIdx + 1).setValue(y);
        filled++;
      }
    }
    Utilities.sleep(500);   // API負荷分散
  }
  Logger.log('[' + label + '] 書込: ' + filled + '件');
  return { target: todo.length, filled: filled };
}

// ========== 工番バックフィル（既存の工番未設定データに一括付与） ==========
function backfillJobNos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lock = getDailyDataLock_();
  if (!lock.tryLock(60000)) { Logger.log('ロック取得失敗'); return; }
  try {
    const main = backfillJobNosForSheet_(ss, SHEET_NAME);
    const archive = ss.getSheetByName(ARCHIVE_SHEET)
      ? backfillJobNosForSheet_(ss, ARCHIVE_SHEET)
      : null;
    const msg = '日報データ: 付与=' + main.assigned + ', 現場なしスキップ=' + main.skippedNoSite + ', 事業部不明スキップ=' + main.skippedNoDivision
      + (archive ? ' / アーカイブ: 付与=' + archive.assigned + ', 現場なし=' + archive.skippedNoSite + ', 事業部不明=' + archive.skippedNoDivision : '');
    Logger.log(msg);
    SpreadsheetApp.getActiveSpreadsheet().toast(msg, '工番バックフィル完了', 10);
    return { main, archive };
  } finally {
    lock.releaseLock();
  }
}

function backfillJobNosForSheet_(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { assigned: 0, skippedNoSite: 0, skippedNoDivision: 0 };
  ensureHeaders_(sheet);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { assigned: 0, skippedNoSite: 0, skippedNoDivision: 0 };

  const headers = data[0];
  const col = (n) => headers.indexOf(n);
  const gCol = col('元請名'), lCol = col('現場名'), rCol = col('役割'), nCol = col('氏名');
  const cCol = col('会社'), yCol = col('夜勤'), dCol = col('事業部'), jCol = col('工番');
  const wtCol = col('作業区分');

  // 代表者マップ: (元請, 現場) → 最初に出現した代表者名（現場作業のみ対象）
  const leaderByKey = {};
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][cCol] || '').trim() !== GROWISE) continue;
    if (String(data[i][rCol] || '').trim() !== '代表') continue;
    const yakin = String(data[i][yCol] || '').trim();
    if (yakin === '休み' || yakin === '予定' || yakin === '倉庫') continue;
    const wt = wtCol >= 0 ? String(data[i][wtCol] || '').trim() : '';
    if (wt && wt !== '現場作業') continue;
    const key = String(data[i][gCol] || '').trim() + '|||' + String(data[i][lCol] || '').trim();
    if (!leaderByKey[key]) leaderByKey[key] = String(data[i][nCol] || '').trim();
  }

  const jobNoCache = {};
  const divisionCache = {};
  let assigned = 0, skippedNoSite = 0, skippedNoDivision = 0;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][cCol] || '').trim() !== GROWISE) continue;
    const yakin = String(data[i][yCol] || '').trim();
    if (yakin === '休み' || yakin === '予定' || yakin === '倉庫') continue;
    const wt = wtCol >= 0 ? String(data[i][wtCol] || '').trim() : '';
    if (wt && wt !== '現場作業') continue;
    if (String(data[i][jCol] || '').trim()) continue; // 既に工番あり

    const genba = String(data[i][gCol] || '').trim();
    const loc = String(data[i][lCol] || '').trim();
    if (!genba) { skippedNoSite++; continue; }

    const key = genba + '|||' + loc;

    if (!jobNoCache[key]) {
      // まず現場マスタにあるか確認
      const existing = findExistingJobNo_(ss, genba, loc);
      if (existing && existing.jobNo) {
        jobNoCache[key] = existing.jobNo;
        divisionCache[key] = existing.division;
      } else {
        // 事業部を決定: 行の事業部列 > 代表者の職人マスタ
        let division = String(data[i][dCol] || '').trim();
        if (!division) {
          const leaderName = leaderByKey[key];
          if (leaderName) division = getMemberDivision_(ss, leaderName);
        }
        if (!division) { skippedNoDivision++; continue; }
        jobNoCache[key] = getOrGenerateJobNo_(ss, genba, loc, division);
        divisionCache[key] = division;
      }
    }

    sheet.getRange(i + 1, jCol + 1).setValue(jobNoCache[key]);
    if (divisionCache[key]) sheet.getRange(i + 1, dCol + 1).setValue(divisionCache[key]);
    assigned++;
  }

  logOperation_(ss, 'backfill_jobnos', sheetName, '付与=' + assigned + ' / 現場なし=' + skippedNoSite + ' / 事業部不明=' + skippedNoDivision, 'system');
  return { assigned, skippedNoSite, skippedNoDivision };
}

function findExistingJobNo_(ss, genba, loc) {
  const jobSiteSheet = getOrCreateJobSiteSheet_(ss);
  const data = jobSiteSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === genba && String(data[i][1]).trim() === loc) {
      return { jobNo: String(data[i][2] || ''), division: String(data[i][3] || '') };
    }
  }
  return null;
}

function archiveOldData_(ss, months) {
  months = months || 3;
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return 0;
  ensureHeaders_(sheet);
  const cutoff = new Date();
  cutoff.setMonth(cutoff.getMonth() - months);
  const tz = Session.getScriptTimeZone();
  let archiveSheet = ss.getSheetByName(ARCHIVE_SHEET);
  if (!archiveSheet) { archiveSheet = ss.insertSheet(ARCHIVE_SHEET); archiveSheet.appendRow(HEADERS); }
  const data = sheet.getDataRange().getValues();
  const rowsToArchive = [];
  const rowNumsToDelete = [];
  for (let i = data.length - 1; i >= 1; i--) {
    const dateVal = data[i][1];
    let rowDate = (dateVal instanceof Date) ? dateVal : new Date(String(dateVal));
    if (!isNaN(rowDate) && rowDate < cutoff) {
      const formatted = data[i].map((v, j) => {
        if (v instanceof Date) {
          if (j === 1) return fmtDate_(v, tz);
          if (j === 6 || j === 7) return fmtTime_(v, tz);
          return fmtDateTime_(v, tz);
        }
        return v;
      });
      rowsToArchive.push(formatted);
      rowNumsToDelete.push(i + 1);
    }
  }
  if (!rowsToArchive.length) return 0;

  // 2026-08-21 高速化: appendRow / deleteRow の1行ずつ繰り返しは、数百行に
  // なると6分の実行上限に届いて途中で止まる（＝いつまでも減らない）。
  // 追記は setValues 1回、削除は連続ブロックの deleteRows にまとめる。
  const width = HEADERS.length;
  const toWrite = rowsToArchive.reverse().map(row => {
    const out = row.slice(0, width);
    while (out.length < width) out.push('');
    return out;
  });
  if (archiveSheet.getMaxColumns() < width) {
    archiveSheet.insertColumnsAfter(archiveSheet.getMaxColumns(), width - archiveSheet.getMaxColumns());
  }
  const startRow = archiveSheet.getLastRow() + 1;
  const needRows = startRow + toWrite.length - 1;
  if (archiveSheet.getMaxRows() < needRows) {
    archiveSheet.insertRowsAfter(archiveSheet.getMaxRows(), needRows - archiveSheet.getMaxRows());
  }
  archiveSheet.getRange(startRow, 1, toWrite.length, width).setValues(toWrite);
  SpreadsheetApp.flush();

  // 削除は必ず「下から」。連番はまとめて1回の deleteRows にする。
  rowNumsToDelete.sort((a, b) => b - a);
  let blockEnd = rowNumsToDelete[0];
  let blockStart = blockEnd;
  for (let k = 1; k < rowNumsToDelete.length; k++) {
    const n = rowNumsToDelete[k];
    if (n === blockStart - 1) { blockStart = n; continue; }
    sheet.deleteRows(blockStart, blockEnd - blockStart + 1);
    blockEnd = n;
    blockStart = n;
  }
  sheet.deleteRows(blockStart, blockEnd - blockStart + 1);
  return toWrite.length;
}

function autoArchive() {
  const lock = getDailyDataLock_();
  if (!lock.tryLock(30000)) return;
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const archived = archiveOldData_(ss, 3);
    logOperation_(ss, 'auto_archive', '3ヶ月以前', '件数=' + archived, 'system');
  } finally {
    lock.releaseLock();
  }
}

function setupDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'dailySummary') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('dailySummary').timeBased().everyDays(1).atHour(6).create();
}

function setupArchiveTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'autoArchive') ScriptApp.deleteTrigger(t);
  });
  // 2026-08-21 変更: 月1回(1日3時)だと、その1回が失敗・未設置だと丸1ヶ月
  // 気づけず日報が溜まり続ける。毎週日曜3時にして取りこぼしを回収する。
  ScriptApp.newTrigger('autoArchive').timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(3).create();
}

// 現在どのトリガーが仕掛かっているかを確認する（実行ログに出す）。
// アーカイブが動いていない疑いが出たら、まずこれを実行する。
function checkTriggers() {
  const list = ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction() + ' / ' + t.getEventType());
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nippo = ss.getSheetByName(SHEET_NAME);
  const arc = ss.getSheetByName(ARCHIVE_SHEET);
  const msg = [
    '仕掛かっているトリガー: ' + (list.length ? list.join(' , ') : '（ゼロ件＝自動処理が一切動いていません）'),
    '日報データ 行数: ' + (nippo ? nippo.getLastRow() - 1 : '(シート無し)'),
    'アーカイブ 行数: ' + (arc ? Math.max(arc.getLastRow() - 1, 0) : '(シート無し)')
  ].join(String.fromCharCode(10));
  Logger.log(msg);
  return msg;
}

function setupAllTriggers() {
  setupDailyTrigger();
  setupArchiveTrigger();
}

function ok(data) {
  return ContentService.createTextOutput(JSON.stringify({status:'ok', ...data})).setMimeType(ContentService.MimeType.JSON);
}
function error(msg) {
  return ContentService.createTextOutput(JSON.stringify({status:'error', message: msg})).setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// 拠点の一括埋め（2026-08-26・利用者承認済み）
//   すでに入っている過去の予定には拠点が無い。会社から一括で埋める
//   （GRミツマ→関東支店／他→本社）。これまでの運用実態と一致する。
//
//   ★オーナーアカウントのApps Scriptエディタから backfillKyoten() を実行する。
//   ★何度実行しても安全: 空欄の行だけを埋める（すでに入っている値は触らない）。
//   ★1行ずつ書かない。setValues で列ごと一括で書く
//     （アーカイブ処理で「1行ずつだと6分の実行上限に達して落ちる」実績があるため）。
// ============================================================
function backfillKyoten() {
  // ★Codexレビュー[P2]#11: 列全体を読んで書き戻すため、実行中に誰かが予定を
  //   登録・編集すると、その変更を巻き戻す恐れがある。予定の書き込みと同じロックを取る。
  const dataLock = getDailyDataLock_();
  if (!dataLock.tryLock(30000)) {
    return '他の処理が予定を更新中のため中止しました。数十秒おいて実行し直してください。';
  }
  try {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const result = [];
  [SHEET_NAME, ARCHIVE_SHEET].forEach(function (name) {
    const sheet = ss.getSheetByName(name);
    if (!sheet) { result.push(name + ': シートなし'); return; }
    ensureHeaders_(sheet);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) { result.push(name + ': データなし'); return; }

    const kyotenCol = HEADERS.indexOf('拠点') + 1;
    const companyCol = HEADERS.indexOf('会社') + 1;

    const n = lastRow - 1;
    const kyotenVals  = sheet.getRange(2, kyotenCol,  n, 1).getValues();
    const companyVals = sheet.getRange(2, companyCol, n, 1).getValues();

    let filled = 0, kept = 0, skipped = 0;
    for (let i = 0; i < n; i++) {
      const cur = String(kyotenVals[i][0] || '').trim();
      if (KYOTEN_VALUES.indexOf(cur) >= 0) { kept++; continue; }   // 既に入っている＝触らない
      // ★拠点の軸を持たない会社（和信カインド・ラーテル・GRHD）は空欄のまま
      if (!hasKyotenAxis_(companyVals[i][0])) { skipped++; continue; }
      // ★Codexレビュー[P1]#8: 利用者が承認したのは「会社から一括で埋める」。
      //   現場マスタを優先すると、getJobsiteKyotenMap_ が現場名だけをキーにしている
      //   ため同名現場の衝突で誤った拠点を大量に確定させ、しかも再実行では
      //   「既に入っている」扱いになって自動修復できない。ここでは会社だけを使う。
      //   （新規登録では現場マスタ優先のままでよい＝1件ずつ目に見える形で入るため）
      kyotenVals[i][0] = defaultKyotenForCompany_(companyVals[i][0]);
      filled++;
    }
    // ★列ごと1回のsetValuesで書き戻す
    sheet.getRange(2, kyotenCol, n, 1).setValues(kyotenVals);
    result.push(name + ': ' + n + '行中 ' + filled + '行を埋めた（既に入っていた ' + kept
      + '行はそのまま／拠点の軸を持たない会社 ' + skipped + '行は空欄のまま）');
  });
  const msg = result.join(' / ');
  Logger.log(msg);
  return msg;
  } finally {
    dataLock.releaseLock();
  }
}
