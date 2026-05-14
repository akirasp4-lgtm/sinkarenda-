const SHEET_NAME = '日報データ';
const ARCHIVE_SHEET = 'アーカイブ';
const MEMBER_SHEET = '職人マスタ';
const GENBA_MASTER_SHEET = '元請マスタ';
const JOBSITE_SHEET = '現場マスタ';
const SUMMARY_COMPANY = '会社別集計';
const SUMMARY_MONTH = '月別集計';
const KAKUNIN_SHEET = '月別確認表';
const BILLING_SHEET = '元請別請求集計';
const ALLOCATION_SHEET = '事業部別按分';
const OPLOG_SHEET = '操作ログ';
const HEADERS = ['登録日時','作業日','元請名','現場名','氏名','役割','出勤','退勤','人工','メモ','夜勤','会社','ID','更新者','色','事業部','工番','作業区分','車両'];
const GROWISE = 'グローライズ';

// ==============================================================
// 社長専用カレンダー（極秘）
// シート名は意図的に内部呼称のみ。PIN認証でのみアクセス可能。
// ==============================================================
const PRES_SHEET = '社長予定';
const PRES_HEADERS = ['登録日時','タイトル','開始日','開始時刻','終了日','終了時刻','場所','メモ','カテゴリ','色','ID','更新者'];
const PRES_PIN = '1203';

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

function doPost(e) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return error('現在他の人が更新中です。数秒待ってから再度お試しください。');
  }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    const body = JSON.parse(e.postData.contents);
    const action = body.action || 'add';
    const updatedBy = String(body.updatedBy || '');

    ensureHeaders_(sheet);
    const idCol = getIdCol_();

    if (action === 'add') {
      const jobNoCache = {};
      let leaderDivision = null;
      const leaderRow = body.rows.find(r => r.role === '代表');
      const leaderName = leaderRow ? leaderRow.name : '';
      body.rows.forEach(row => {
        let division = '';
        let jobNo = '';
        // 工番発行は「グローライズ × 倉庫/休み/予定 のいずれでもない」場合のみ（夜勤は通常の工番発行対象）
        if (row.company === GROWISE && !row.souko && !row.yotei && !row.yasumi && row.workType === '現場作業') {
          const explicitDiv = String(row.jobNoDivision || '').trim();
          if (explicitDiv) {
            division = explicitDiv;
          } else {
            if (leaderDivision === null) {
              leaderDivision = getMemberDivision_(ss, leaderName);
            }
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
        sheet.appendRow([
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
          row.vehicle || ''
        ]);
      });
      logOperation_(ss, 'add', body.rows[0] && body.rows[0].genba + '/' + (body.rows[0].loc || ''), '行数=' + body.rows.length, updatedBy);
      return ok({count: body.rows.length});
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
      const ids = body.ids || [];
      if (ids.length > 0) {
        const data = sheet.getDataRange().getValues();
        const rowsToDelete = [];
        for (let i = data.length - 1; i >= 1; i--) {
          const rowId = String(data[i][idCol] || '').trim();
          if (rowId && ids.includes(rowId)) rowsToDelete.push(i + 1);
        }
        rowsToDelete.forEach(rowNum => sheet.deleteRow(rowNum));
      }
      const jobNoCache = {};
      let leaderDivision = null;
      const leaderRow = body.rows.find(r => r.role === '代表');
      const leaderName = leaderRow ? leaderRow.name : '';
      body.rows.forEach(row => {
        let division = '';
        let jobNo = '';
        // 工番発行は「グローライズ × 倉庫/休み/予定 のいずれでもない」場合のみ（夜勤は通常の工番発行対象）
        if (row.company === GROWISE && !row.souko && !row.yotei && !row.yasumi && row.workType === '現場作業') {
          const explicitDiv = String(row.jobNoDivision || '').trim();
          if (explicitDiv) {
            division = explicitDiv;
          } else {
            if (leaderDivision === null) {
              leaderDivision = getMemberDivision_(ss, leaderName);
            }
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
        sheet.appendRow([
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
          row.vehicle || ''
        ]);
      });
      logOperation_(ss, 'update', body.rows[0] && body.rows[0].genba + '/' + (body.rows[0].loc || ''), '行数=' + body.rows.length + ', 旧ID=' + (body.ids || []).length, updatedBy);
      return ok({updated: body.rows.length});
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

    if (action === 'summarize') {
      generateSummary_();
      return ok({message: '集計を更新しました'});
    }

    if (action === 'get_sheet') {
      const sheetName = body.sheet || '';
      const allowed = [SHEET_NAME, ARCHIVE_SHEET, MEMBER_SHEET, GENBA_MASTER_SHEET, JOBSITE_SHEET, SUMMARY_COMPANY, SUMMARY_MONTH, KAKUNIN_SHEET, BILLING_SHEET, ALLOCATION_SHEET, OPLOG_SHEET];
      if (!allowed.includes(sheetName)) return error('無効なシート名です');
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
          const body_ = data.slice(1).filter(row => {
            const v = row[dateColIdx];
            const d = v instanceof Date
              ? Utilities.formatDate(v, tz, 'yyyy-MM-dd')
              : String(v || '').slice(0, 10);
            if (dateFrom && d < dateFrom) return false;
            if (dateTo && d > dateTo) return false;
            return true;
          });
          filtered = head.concat(body_);
        }
      }
      const formatted = filtered.map(row => row.map(v => {
        if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd HH:mm:ss');
        return v;
      }));
      return ok({sheetName, data: formatted});
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
    // 社長専用カレンダー（極秘）
    // すべてのアクションで PIN チェックを行う。
    // ============================================================
    if (action === 'pres_list' || action === 'pres_add' || action === 'pres_update' || action === 'pres_delete') {
      if (String(body.pin || '') !== PRES_PIN) {
        return error('認証に失敗しました');
      }
      const presSheet = getOrCreatePresSheet_(ss);
      const tz = Session.getScriptTimeZone();

      if (action === 'pres_list') {
        const data = presSheet.getDataRange().getValues();
        let rows = [];
        if (data.length > 1) {
          const headers = data[0];
          rows = data.slice(1).map(r => {
            const obj = {};
            headers.forEach((h, j) => {
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
        return ok({rows});
      }

      if (action === 'pres_add') {
        const ev = body.event || {};
        const id = 'P' + new Date().getTime() + '_' + Math.floor(Math.random() * 10000);
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
        for (let i = 1; i < data.length; i++) {
          if (String(data[i][idCol]) === id) {
            presSheet.getRange(i + 1, 1, 1, PRES_HEADERS.length).setValues([[
              data[i][0] instanceof Date ? data[i][0] : new Date(),
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
            ]]);
            return ok({updated: id});
          }
        }
        return error('対象が見つかりませんでした');
      }

      if (action === 'pres_delete') {
        const id = String(body.id || '');
        if (!id) return error('IDが指定されていません');
        const data = presSheet.getDataRange().getValues();
        const idCol = PRES_HEADERS.indexOf('ID');
        for (let i = data.length - 1; i >= 1; i--) {
          if (String(data[i][idCol]) === id) {
            presSheet.deleteRow(i + 1);
            return ok({deleted: id});
          }
        }
        return error('対象が見つかりませんでした');
      }
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

    return ok({});
  } catch(err) {
    return error(err.toString());
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    const tz = Session.getScriptTimeZone();
    ensureHeaders_(sheet);
    const data = sheet.getDataRange().getValues();
    let rows = [];
    if (data.length > 1) {
      const headers = data[0];
      rows = data.slice(1).map(row => {
        const obj = {};
        headers.forEach((h, j) => {
          const v = row[j];
          if (h === '作業日') obj[h] = (v instanceof Date) ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : String(v || '');
          else if (h === '出勤' || h === '退勤') obj[h] = (v instanceof Date) ? Utilities.formatDate(v, tz, 'HH:mm') : String(v || '');
          else obj[h] = (v === undefined || v === null) ? '' : v;
        });
        return obj;
      });
    }
    const memberSheet = getOrCreateMemberSheet_(ss);
    const mData = memberSheet.getDataRange().getValues();
    const members = mData.length > 1 ? mData.slice(1).map(r => ({
      name: String(r[0]||''),
      company: String(r[1]||''),
      division: String(r[2]||''),
      rate: Number(r[3]||0)
    })) : [];

    const genbaSheet = getOrCreateGenbaSheet_(ss);
    const gData = genbaSheet.getDataRange().getValues();
    const genbaMaster = gData.length > 1 ? gData.slice(1).map(r => ({name: String(r[0]||''), company: String(r[1]||'')})).filter(g => g.name) : [];

    return ok({rows, members, genbaMaster});
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
    sheet.appendRow(['元請名', '現場名', '工番', '事業部', '年度', '連番', '売上', '読み']);
  } else {
    ensureColumns_(sheet, 8);
    const headers = sheet.getRange(1, 1, 1, 8).getValues()[0];
    if (String(headers[6] || '').trim() !== '売上') sheet.getRange(1, 7).setValue('売上');
    if (String(headers[7] || '').trim() !== '読み') sheet.getRange(1, 8).setValue('読み');
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

  const mainRecords = sheetToRecords(ss.getSheetByName(SHEET_NAME));
  const archiveRecords = sheetToRecords(ss.getSheetByName(ARCHIVE_SHEET));

  generateCompanySummary_(ss, mainRecords);
  generateMonthSummary_(ss, mainRecords);
  generateBillingSummary_(ss, mainRecords);

  const allRecords = [...mainRecords, ...archiveRecords];
  generateKakuninTable_(ss, allRecords);
  generateDivisionAllocation_(ss, allRecords);

  // 現場マスタの孤立行を掃除（日報データに1件も存在しない＆売上未入力のもの）
  try {
    const orphanCount = cleanupOrphanSites_(ss);
    if (orphanCount > 0) logOperation_(ss, 'cleanup_orphan_sites', '現場マスタ', '削除=' + orphanCount, 'auto');
  } catch (e) {}
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
    formats.forEach(f => {
      const range = sheet.getRange(f.row, 1, 1, 6);
      if (f.type === 'title') sheet.getRange(f.row, 1).setFontSize(14).setFontWeight('bold');
      else if (f.type === 'company') { range.setBackground('#E8F5E9'); sheet.getRange(f.row, 1).setFontSize(12).setFontWeight('bold'); }
      else if (f.type === 'header') range.setFontWeight('bold').setBackground('#F5F5F5');
      else if (f.type === 'total') range.setFontWeight('bold').setBackground('#FFF9C4');
    });
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
    formats.forEach(f => {
      const range = sheet.getRange(f.row, 1, 1, 6);
      if (f.type === 'title') sheet.getRange(f.row, 1).setFontSize(14).setFontWeight('bold');
      else if (f.type === 'month') { range.setBackground('#E3F2FD'); sheet.getRange(f.row, 1).setFontSize(12).setFontWeight('bold'); }
      else if (f.type === 'header') range.setFontWeight('bold').setBackground('#F5F5F5');
      else if (f.type === 'total') range.setFontWeight('bold').setBackground('#FFF9C4');
    });
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
  for (let c = 2; c <= 32; c++) sheet.setColumnWidth(c, 28);
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
    sheet.getRange(1, 1, outputData.length, maxCols).setValues(outputData);

    formatRules.forEach(rule => {
      const r = rule.row + 1;
      if (rule.type === 'title') {
        sheet.getRange(r, 1, 1, rule.cols).merge().setHorizontalAlignment('center').setFontSize(13).setFontWeight('bold').setBackground('#F9E400');
      } else if (rule.type === 'header') {
        const range = sheet.getRange(r, 1, 1, rule.cols);
        range.setFontWeight('bold').setBackground('#CCCCCC').setHorizontalAlignment('center');
        // 日付の数字(1〜31)が日付シリアルとして解釈されないよう、整数書式を明示
        sheet.getRange(r, 2, 1, rule.daysInMonth).setNumberFormat('0');
        for (let d = 1; d <= rule.daysInMonth; d++) {
          const dow = new Date(rule.year, rule.month, d).getDay();
          const cell = sheet.getRange(r, d + 1);
          if (dow === 0) cell.setFontColor('#CC0000');
          else if (dow === 6) cell.setFontColor('#0000CC');
        }
      } else if (rule.type === 'empty_data') {
        sheet.getRange(r, 1).setFontColor('#999999');
      } else if (rule.type === 'data') {
        const bg = rule.index % 2 === 0 ? '#FFFFFF' : '#F0FFF0';
        sheet.getRange(r, 1, 1, rule.cols).setBackground(bg);
        sheet.getRange(r, 1).setFontWeight('bold');
        sheet.getRange(r, 2, 1, rule.cols - 1).setNumberFormat('0.##');
        for (let d = 1; d <= rule.daysInMonth; d++) {
          const dow = new Date(rule.year, rule.month, d).getDay();
          const cell = sheet.getRange(r, d + 1);
          cell.setHorizontalAlignment('center');
          const val = outputData[rule.row][d];
          if (val === 0) cell.setFontColor('#CCCCCC');
          if (dow === 0) cell.setBackground('#FFE6E6');
          else if (dow === 6) cell.setBackground('#E6E6FF');
        }
        sheet.getRange(r, rule.cols).setFontWeight('bold').setHorizontalAlignment('center');
      } else if (rule.type === 'total') {
        sheet.getRange(r, 1, 1, rule.cols).setFontWeight('bold').setBackground('#FFF9C4');
        sheet.getRange(r, 2, 1, rule.cols - 1).setNumberFormat('0.##');
        for (let d = 1; d <= rule.daysInMonth; d++) sheet.getRange(r, d + 1).setHorizontalAlignment('center');
        sheet.getRange(r, rule.cols).setHorizontalAlignment('center');
        const startRow = r - rule.namesLength - 1;
        sheet.getRange(startRow, 1, rule.namesLength + 2, rule.cols).setBorder(true, true, true, true, true, true);
      }
    });
  }
}

function generateBillingSummary_(ss, records) {
  let sheet = ss.getSheetByName(BILLING_SHEET);
  if (sheet) { sheet.clear(); sheet.clearFormats(); } else { sheet = ss.insertSheet(BILLING_SHEET); }
  ensureColumns_(sheet, 35);
  // 倉庫は元請に請求しない作業のため除外（旧データで元請名が入っているものも対象外にする）
  const workRecords = records.filter(r => r.yakin !== '休み' && r.yakin !== '予定' && r.yakin !== '倉庫');
  const months = [...new Set(workRecords.map(r => r.month).filter(Boolean))].sort().reverse();
  const genbas = [...new Set(workRecords.map(r => r.genba).filter(Boolean))].sort();
  const DOW = ['日','月','火','水','木','金','土'];
  let currentRow = 1;
  months.forEach(month => {
    const parts = month.split('-');
    const year = Number(parts[0]);
    const mon = Number(parts[1]);
    const monthLabel = year + '年' + mon + '月';
    const daysInMonth = new Date(year, mon, 0).getDate();
    const totalCols = 3 + daysInMonth + 1;
    ensureColumns_(sheet, totalCols);
    const mr = workRecords.filter(r => r.month === month);
    // (氏名, 日付, 昼夜区分) → 行った現場のSet。1日に複数現場行ったら 1/N で按分する
    const sitesByPDN = {};
    mr.forEach(r => {
      const dn = r.yakin === '夜勤' ? 'N' : 'D';
      const k = r.name + '|' + r.date + '|' + dn;
      if (!sitesByPDN[k]) sitesByPDN[k] = new Set();
      sitesByPDN[k].add(r.genba + '|||' + (r.loc || '（現場名なし）'));
    });
    sheet.getRange(currentRow, 1, 1, totalCols).merge().setValue('▶ ' + monthLabel).setBackground('#1D9E75').setFontColor('#FFFFFF').setFontSize(12).setFontWeight('bold');
    currentRow++;
    const headerRow = ['会社名', '現場名', '名前'];
    for (let d = 1; d <= daysInMonth; d++) { const dow = new Date(year, mon - 1, d).getDay(); headerRow.push(d + ' ' + DOW[dow]); }
    headerRow.push('合計');
    sheet.getRange(currentRow, 1, 1, headerRow.length).setValues([headerRow]).setFontWeight('bold').setBackground('#CCCCCC').setHorizontalAlignment('center').setWrap(true);
    for (let d = 1; d <= daysInMonth; d++) { const dow = new Date(year, mon - 1, d).getDay(); const cell = sheet.getRange(currentRow, 3 + d); if (dow === 0) cell.setFontColor('#CC0000'); else if (dow === 6) cell.setFontColor('#0000CC'); }
    currentRow++;
    genbas.forEach(genba => {
      const gr = mr.filter(r => r.genba === genba);
      if (gr.length === 0) return;
      const locs = [...new Set(gr.map(r => r.loc || '（現場名なし）'))].sort();
      locs.forEach(loc => {
        const lr = gr.filter(r => (r.loc || '（現場名なし）') === loc);
        const namesInLoc = [...new Set(lr.map(r => r.name))].sort();
        const activeNames = namesInLoc.filter(name => calcEffective_(lr, name).kosu > 0);
        if (activeNames.length === 0) return;
        const blockStartRow = currentRow;
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
          sheet.getRange(currentRow, 1, 1, row.length).setValues([row]);
          const bg = ni % 2 === 0 ? '#FFFFFF' : '#F0FFF0';
          sheet.getRange(currentRow, 1, 1, totalCols).setBackground(bg);
          sheet.getRange(currentRow, 1).setFontWeight('bold');
          sheet.getRange(currentRow, 4, 1, row.length - 3).setNumberFormat('0.##');
          for (let d = 1; d <= daysInMonth; d++) { const dow = new Date(year, mon - 1, d).getDay(); const cell = sheet.getRange(currentRow, 3 + d); cell.setHorizontalAlignment('center'); const val = row[3 + d - 1]; if (val === 0) cell.setFontColor('#CCCCCC'); if (dow === 0) cell.setBackground('#FFE6E6'); else if (dow === 6) cell.setBackground('#E6E6FF'); }
          sheet.getRange(currentRow, 3 + daysInMonth + 1).setFontWeight('bold').setHorizontalAlignment('center');
          currentRow++;
        });
        const totalRow = ['', '', '合計'];
        let grandTotal = 0;
        for (let d = 1; d <= daysInMonth; d++) { const dateStr = year + '-' + String(mon).padStart(2,'0') + '-' + String(d).padStart(2,'0'); let daySum = 0; activeNames.forEach(name => { const dayRecs = lr.filter(r => r.name === name && r.date === dateStr); const hasDay = dayRecs.some(r => r.yakin !== '夜勤'); const hasNight = dayRecs.some(r => r.yakin === '夜勤'); if (hasDay) { const sCnt = (sitesByPDN[name + '|' + dateStr + '|D'] || new Set()).size || 1; daySum += 1 / sCnt; } if (hasNight) { const sCnt = (sitesByPDN[name + '|' + dateStr + '|N'] || new Set()).size || 1; daySum += 1 / sCnt; } }); totalRow.push(daySum > 0 ? daySum : 0); grandTotal += daySum; }
        totalRow.push(grandTotal);
        sheet.getRange(currentRow, 1, 1, totalRow.length).setValues([totalRow]).setFontWeight('bold').setBackground('#FFF9C4');
        sheet.getRange(currentRow, 4, 1, totalRow.length - 3).setNumberFormat('0.##');
        for (let d = 1; d <= daysInMonth; d++) sheet.getRange(currentRow, 3 + d).setHorizontalAlignment('center');
        sheet.getRange(currentRow, 3 + daysInMonth + 1).setHorizontalAlignment('center');
        if (activeNames.length > 1) { sheet.getRange(blockStartRow, 1, activeNames.length, 1).merge(); sheet.getRange(blockStartRow, 2, activeNames.length, 1).merge(); }
        sheet.getRange(blockStartRow, 1).setFontWeight('bold').setVerticalAlignment('middle');
        sheet.getRange(blockStartRow, 2).setFontWeight('bold').setVerticalAlignment('middle');
        currentRow++;
        currentRow++;
      });
    });
    currentRow++;
  });
  sheet.setColumnWidth(1, 140); sheet.setColumnWidth(2, 180); sheet.setColumnWidth(3, 80);
  const maxCols = sheet.getMaxColumns();
  for (let c = 4; c <= Math.min(maxCols, 35); c++) sheet.setColumnWidth(c, 26);
  if (maxCols >= 36) sheet.setColumnWidth(Math.min(maxCols, 36), 50);
  if (currentRow > 1) { const borderCols = Math.min(maxCols, 35); sheet.getRange(1, 1, currentRow - 1, borderCols).setBorder(true, true, true, true, true, true, '#DDDDDD', SpreadsheetApp.BorderStyle.SOLID); }
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
    formats.forEach(f => {
      const range = sheet.getRange(f.row, 1, 1, numCols);
      if (f.type === 'title') sheet.getRange(f.row, 1).setFontSize(14).setFontWeight('bold');
      else if (f.type === 'section_total') { range.setBackground('#E8F5E9'); sheet.getRange(f.row, 1).setFontSize(12).setFontWeight('bold'); }
      else if (f.type === 'section_month') { range.setBackground('#E3F2FD'); sheet.getRange(f.row, 1).setFontSize(12).setFontWeight('bold'); }
      else if (f.type === 'header') range.setFontWeight('bold').setBackground('#F5F5F5').setHorizontalAlignment('center');
      else if (f.type === 'total') range.setFontWeight('bold').setBackground('#FFF9C4');
    });
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
  for (let c = 5; c < 5 + divs.length; c++) sheet.setColumnWidth(c, 70);
  sheet.setColumnWidth(5 + divs.length, 80);
  for (let c = 6 + divs.length; c < 6 + divs.length * 2; c++) sheet.setColumnWidth(c, 60);
  for (let c = 6 + divs.length * 2; c < 6 + divs.length * 3; c++) sheet.setColumnWidth(c, 90);
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
  const lock = LockService.getScriptLock();
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
  const lock = LockService.getScriptLock();
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
          if (j === 1) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
          if (j === 6 || j === 7) return Utilities.formatDate(v, tz, 'HH:mm');
          return Utilities.formatDate(v, tz, 'yyyy-MM-dd HH:mm:ss');
        }
        return v;
      });
      rowsToArchive.push(formatted);
      rowNumsToDelete.push(i + 1);
    }
  }
  rowsToArchive.reverse().forEach(row => archiveSheet.appendRow(row));
  rowNumsToDelete.forEach(rowNum => sheet.deleteRow(rowNum));
  return rowsToArchive.length;
}

function autoArchive() {
  const lock = LockService.getScriptLock();
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
  ScriptApp.newTrigger('autoArchive').timeBased().onMonthDay(1).atHour(3).create();
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
