// 夜勤区分（勤務区分・夜勤手当・夜勤請求）— 2026-09-03
//
// ★依頼書「予定表カレンダー改修.txt」:
//   「Excel出力上で夜勤かどうかを判定できない。そのため夜勤に入った作業員への
//     夜勤手当、元請への夜勤応援分の請求を機械的に拾えない」
//
// ★このテストが守るもの（壊れたら赤くなる）:
//   1. 勤務区分は『夜勤』列から導出する。空欄＝日勤（＝改修前の全データ）
//   2. 夜勤手当・夜勤請求は空欄＝自動（夜勤ならYes）。'対象外' で個別に外せる
//   3. 依頼書の完了条件「既存の人工合計と、日勤＋夜勤の合計が一致する」
//      → 月別確認表の合計 − 夜勤確認表の夜勤人工合計 ＝ 日勤の人工
//   4. 元請別請求集計_フィルタ用 で日勤と夜勤が別行になり、合計行は改修前と同じ数字
//   5. 保存時に画面が値を送ってこなくても壊れない（空欄＝自動へ戻る）
import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import vm from 'node:vm';

const here = dirname(fileURLToPath(import.meta.url));
const CODE = readFileSync(join(here, '..', '..', 'gas.js'), 'utf8');

const EXPORT = `
;globalThis.__gas = {
  workClass_, isWorkClass_, yakinFlag_, yakinTeateOn_, yakinSeikyuOn_,
  yakinNeedsCheck_, yakinCheckNote_, normalizeYakinFlag_, dailyKosuBuckets_,
  generateKakuninTable_, generateNightKakuninTable_,
  generateWorkerDetailSheet_, generateBillingFilterSheet_,
  HEADERS, DETAIL_SHEET, NIGHT_KAKUNIN_SHEET, KAKUNIN_SHEET, BILLING_FILTER_SHEET
};
`;

// ---- 最小限の偽スプレッドシート ----------------------------------------
// setValues で書かれた 2D 配列だけ覚える。書式（色・幅・罫線）は呼ばれても捨てる。
function makeSheet(name) {
  const sheet = {
    name,
    written: null,            // 最後に setValues された表
    getMaxColumns: () => 200,
    insertColumnsAfter() {},
    clear() { sheet.written = null; },
    clearFormats() {},
    getFilter: () => null,
    setColumnWidth() {}, setColumnWidths() {},
    setFrozenRows() {}, setFrozenColumns() {},
    getRange(row, col, numRows, numCols) {
      const range = {
        setValues(vals) {
          // 左上が (1,1) の書き込みだけを「本文」として覚える
          if (row === 1 && col === 1) sheet.written = vals.map(r => r.slice());
          return range;
        },
        setValue() { return range; },
        setBackground: () => range, setBackgrounds: () => range,
        setFontWeight: () => range, setFontWeights: () => range,
        setFontColor: () => range, setFontColors: () => range,
        setFontSize: () => range, setFontSizes: () => range,
        setHorizontalAlignment: () => range, setHorizontalAlignments: () => range,
        setVerticalAlignment: () => range, setVerticalAlignments: () => range,
        setNumberFormat: () => range, setNumberFormats: () => range,
        setWrap: () => range, setWraps: () => range,
        setBorder: () => range, merge: () => range,
        setFormula: () => range, setFormulas: () => range,
        setFontStyle: () => range, setFontLine: () => range,
        createFilter: () => range
      };
      return range;
    }
  };
  return sheet;
}
function makeSS() {
  const sheets = {};
  return {
    sheets,
    getSheetByName: (n) => sheets[n] || null,
    insertSheet: (n) => (sheets[n] = makeSheet(n)),
    deleteSheet() {}
  };
}

function load() {
  const sandbox = vm.createContext({
    SpreadsheetApp: { BorderStyle: { SOLID: 'SOLID' }, flush() {} },
    Session: { getScriptTimeZone: () => 'Asia/Tokyo' },
    LockService: {}, Utilities: {}, ContentService: {}, UrlFetchApp: {},
    PropertiesService: {}, Logger: { log() {} }, console
  });
  vm.runInContext(CODE + EXPORT, sandbox, { filename: 'gas.js' });
  return sandbox.__gas;
}
const g = load();

// 今月の日付を作る（月別確認表は「来月〜2ヶ月前」の窓しか出さないため）
const NOW = new Date();
const Y = NOW.getFullYear();
const M = NOW.getMonth() + 1;
const MM = String(M).padStart(2, '0');
const day = (d) => `${Y}-${MM}-${String(d).padStart(2, '0')}`;

function rec(over) {
  return Object.assign({
    date: day(1), month: `${Y}-${MM}`, name: '田中', kosu: 1,
    company: 'グローライズ', genba: 'みなと建設', loc: '名古屋A現場',
    kyoten: '本社', yakin: '', start: '', end: '',
    memo: '', id: 'id1', teate: '', seikyu: ''
  }, over || {});
}

// =========================================================
describe('勤務区分は『夜勤』列から導出する（新しい列を作らない）', () => {
  it('空欄は日勤。改修前の全データがこれ', () => {
    expect(g.workClass_('')).toBe('日勤');
    expect(g.workClass_(null)).toBe('日勤');
    expect(g.workClass_(undefined)).toBe('日勤');
  });
  it('夜勤・休み・予定・倉庫はそのまま', () => {
    expect(g.workClass_('夜勤')).toBe('夜勤');
    expect(g.workClass_('休み')).toBe('休み');
    expect(g.workClass_('予定')).toBe('予定');
    expect(g.workClass_('倉庫')).toBe('倉庫');
  });
  it('前後の空白は無視する', () => {
    expect(g.workClass_(' 夜勤 ')).toBe('夜勤');
  });
  it('休み・予定は実働ではない。倉庫は実働', () => {
    expect(g.isWorkClass_('日勤')).toBe(true);
    expect(g.isWorkClass_('夜勤')).toBe(true);
    expect(g.isWorkClass_('倉庫')).toBe(true);
    expect(g.isWorkClass_('休み')).toBe(false);
    expect(g.isWorkClass_('予定')).toBe(false);
  });
});

describe('夜勤手当・夜勤請求は空欄＝自動、例外だけ手で外す', () => {
  it('空欄なら夜勤はYes・日勤はNo（45人ぶん毎日入力させないための既定）', () => {
    expect(g.yakinTeateOn_(rec({ yakin: '夜勤' }))).toBe(true);
    expect(g.yakinTeateOn_(rec({ yakin: '' }))).toBe(false);
    expect(g.yakinSeikyuOn_(rec({ yakin: '夜勤' }))).toBe(true);
    expect(g.yakinSeikyuOn_(rec({ yakin: '' }))).toBe(false);
  });
  it('夜勤でも「対象外」と書けば個別に外せる（依頼書2: 作業員ごとの上書き）', () => {
    expect(g.yakinTeateOn_(rec({ yakin: '夜勤', teate: '対象外' }))).toBe(false);
    expect(g.yakinSeikyuOn_(rec({ yakin: '夜勤', seikyu: '対象外' }))).toBe(false);
  });
  it('手当と請求は別々に判断できる（依頼書3）', () => {
    const r = rec({ yakin: '夜勤', teate: '対象', seikyu: '対象外' });
    expect(g.yakinTeateOn_(r)).toBe(true);
    expect(g.yakinSeikyuOn_(r)).toBe(false);
  });
  it('日勤でも「対象」と書けば手当を出せる', () => {
    expect(g.yakinTeateOn_(rec({ yakin: '', teate: '対象' }))).toBe(true);
  });
});

describe('保存の値の正規化（画面から何が来ても壊さない）', () => {
  it('対象・対象外・空 の3つだけを通す', () => {
    expect(g.normalizeYakinFlag_('対象')).toBe('対象');
    expect(g.normalizeYakinFlag_('対象外')).toBe('対象外');
    expect(g.normalizeYakinFlag_('')).toBe('');
  });
  it('画面が値を送ってこなくても落ちない（空＝自動へ戻る）', () => {
    expect(g.normalizeYakinFlag_(undefined)).toBe('');
    expect(g.normalizeYakinFlag_(null)).toBe('');
    expect(g.normalizeYakinFlag_('へんな値')).toBe('');
  });
  it('true/false や ○× も受け取れる', () => {
    expect(g.normalizeYakinFlag_(true)).toBe('対象');
    expect(g.normalizeYakinFlag_(false)).toBe('対象外');
    expect(g.normalizeYakinFlag_('○')).toBe('対象');
    expect(g.normalizeYakinFlag_('×')).toBe('対象外');
  });
});

describe('確認漏れの検知', () => {
  it('夜勤なのに出勤・退勤が空なら要確認', () => {
    expect(g.yakinCheckNote_(rec({ yakin: '夜勤' }))).toBe('要確認：時刻なし');
    expect(g.yakinCheckNote_(rec({ yakin: '夜勤', start: '22:00' }))).toBe('要確認：時刻なし');
  });
  it('日をまたぐ夜勤は正常（22:00〜05:00）', () => {
    expect(g.yakinCheckNote_(rec({ yakin: '夜勤', start: '22:00', end: '05:00' }))).toBe('');
  });
  it('日勤は時刻が空でも要確認にしない（夜勤手当の話ではないため）', () => {
    expect(g.yakinCheckNote_(rec({ yakin: '' }))).toBe('');
  });

  // ★2026-09-05 検品③で本番に見つかった入力ミス。利用者確認「時間の間違いやね」。
  //   9/17・18の夜勤が 08:00〜17:00 で登録されていた。時刻が空でないので拾えていなかった。
  it('★夜勤なのに丸ごと昼の時間帯なら要確認（08:00〜17:00 は入力ミス）', () => {
    expect(g.yakinCheckNote_(rec({ yakin: '夜勤', start: '08:00', end: '17:00' }))).toBe('要確認：昼の時刻');
    expect(g.yakinCheckNote_(rec({ yakin: '夜勤', start: '09:00', end: '18:00' }))).toBe('要確認：昼の時刻');
  });
  it('夕方以降に出勤する夜勤は正常', () => {
    expect(g.yakinCheckNote_(rec({ yakin: '夜勤', start: '18:00', end: '23:00' }))).toBe('');
    expect(g.yakinCheckNote_(rec({ yakin: '夜勤', start: '20:00', end: '23:59' }))).toBe('');
  });
  it('朝までに退勤する夜勤は正常', () => {
    expect(g.yakinCheckNote_(rec({ yakin: '夜勤', start: '00:00', end: '09:00' }))).toBe('');
  });
  it('読めない時刻は「時刻なし」扱い（安全側）', () => {
    expect(g.yakinCheckNote_(rec({ yakin: '夜勤', start: 'あさ', end: 'よる' }))).toBe('要確認：時刻なし');
    expect(g.yakinCheckNote_(rec({ yakin: '夜勤', start: '25:00', end: '05:00' }))).toBe('要確認：時刻なし');
  });
  it('日勤は昼の時刻でも要確認にしない', () => {
    expect(g.yakinCheckNote_(rec({ yakin: '', start: '08:00', end: '17:00' }))).toBe('');
  });
});

describe('人工の昼夜バケット（月別確認表と夜勤確認表の共通の土台）', () => {
  it('同じ日に昼と夜勤があれば 1.0 + 1.0 = 2.0', () => {
    const rs = [rec({ yakin: '' }), rec({ yakin: '夜勤' })];
    expect(g.dailyKosuBuckets_(rs, '田中', day(1))).toEqual({ day: 1, night: 1 });
  });
  it('同じバケットの重複入力は最大値を採る（二重に数えない）', () => {
    const rs = [rec({ kosu: 1 }), rec({ kosu: 0.5, loc: '別現場' })];
    expect(g.dailyKosuBuckets_(rs, '田中', day(1))).toEqual({ day: 1, night: 0 });
  });
  it('休み・予定は実働に数えない', () => {
    const rs = [rec({ yakin: '休み', kosu: 0 }), rec({ yakin: '予定', kosu: 1 })];
    expect(g.dailyKosuBuckets_(rs, '田中', day(1))).toEqual({ day: 0, night: 0 });
  });
  it('倉庫は昼のバケットに入る', () => {
    const rs = [rec({ yakin: '倉庫' })];
    expect(g.dailyKosuBuckets_(rs, '田中', day(1))).toEqual({ day: 1, night: 0 });
  });
});

// =========================================================
// 依頼書の完了条件をそのまま試す
// =========================================================
describe('★完了条件: 既存の人工合計と、日勤＋夜勤の合計が一致する', () => {
  // 田中: 1日 昼のみ / 2日 夜勤のみ / 3日 昼+夜勤 / 4日 休み
  // 佐藤: 1日 夜勤のみ
  const records = [
    rec({ name: '田中', date: day(1), yakin: '' }),
    rec({ name: '田中', date: day(2), yakin: '夜勤' }),
    rec({ name: '田中', date: day(3), yakin: '' }),
    rec({ name: '田中', date: day(3), yakin: '夜勤' }),
    rec({ name: '田中', date: day(4), yakin: '休み', kosu: 0 }),
    rec({ name: '佐藤', date: day(1), yakin: '夜勤' })
  ];

  function totalsOf(sheetRows, labelCol, totalCol, label) {
    // 「合計」行を除いた本文行の合計列を足す
    let sum = 0;
    sheetRows.forEach(r => {
      if (r[labelCol] === label) return;
      const v = r[totalCol];
      if (typeof v === 'number') sum += v;
    });
    return sum;
  }

  it('月別確認表の合計 − 夜勤確認表の夜勤人工合計 ＝ 日勤の人工', () => {
    const ss = makeSS();
    g.generateKakuninTable_(ss, records);
    g.generateNightKakuninTable_(ss, records);

    // --- 月別確認表: 名前行の合計列（daysInMonth+1 列目）を足す ---
    const kak = ss.sheets[g.KAKUNIN_SHEET].written;
    const daysInMonth = new Date(Y, M, 0).getDate();
    let kakTotal = 0;
    kak.forEach(row => {
      const nameCell = row[0];
      if (!nameCell || nameCell === '合計' || nameCell === '名前 ▼') return;
      if (String(nameCell).includes('年')) return;         // 月タイトル行
      if (nameCell === '（データなし）') return;
      const v = row[daysInMonth + 1];
      if (typeof v === 'number') kakTotal += v;
    });

    // --- 夜勤確認表: 名前行の「夜勤人工合計」列（36列目 = index 35） ---
    const night = ss.sheets[g.NIGHT_KAKUNIN_SHEET].written;
    let nightTotal = 0;
    night.slice(1).forEach(row => {
      if (row[1] === '合計') return;
      const v = row[35];
      if (typeof v === 'number') nightTotal += v;
    });

    // --- 期待値: 昼 田中3日ぶん? → 田中 1日+3日 = 2.0 昼 / 夜勤 田中2日+3日 + 佐藤1日 = 3.0
    expect(kakTotal).toBe(5);      // 昼2.0 + 夜3.0
    expect(nightTotal).toBe(3);
    expect(kakTotal - nightTotal).toBe(2);   // ＝日勤の人工
  });

  it('夜勤確認表には夜勤日数が出る（完了条件3: 作業員ごとの夜勤日数・夜勤人工）', () => {
    const ss = makeSS();
    g.generateNightKakuninTable_(ss, records);
    const rows = ss.sheets[g.NIGHT_KAKUNIN_SHEET].written;
    const header = rows[0];
    expect(header[0]).toBe('月');
    expect(header[1]).toBe('名前');
    expect(header[34]).toBe('夜勤日数');
    expect(header[35]).toBe('夜勤人工合計');

    const tanaka = rows.find(r => r[1] === '田中');
    expect(tanaka[34]).toBe(2);   // 2日・3日の2日
    expect(tanaka[35]).toBe(2);
    const sato = rows.find(r => r[1] === '佐藤');
    expect(sato[34]).toBe(1);
    expect(sato[35]).toBe(1);
  });

  it('夜勤が1件も無ければ夜勤確認表に人の行は出ない', () => {
    const ss = makeSS();
    g.generateNightKakuninTable_(ss, [rec({ yakin: '' })]);
    const rows = ss.sheets[g.NIGHT_KAKUNIN_SHEET].written;
    expect(rows.length).toBe(1);   // ヘッダーだけ
  });
});

describe('作業者日別明細シート（依頼書4の必須出力）', () => {
  const records = [
    rec({ name: '田中', date: day(2), yakin: '夜勤', start: '22:00', end: '05:00', memo: '応援' }),
    rec({ name: '佐藤', date: day(2), yakin: '夜勤' }),                 // 時刻なし → 要確認
    rec({ name: '鈴木', date: day(1), yakin: '', teate: '対象' }),      // 日勤だが手当あり
    rec({ name: '高橋', date: day(1), yakin: '休み', kosu: 0 })         // 明細には出さない
  ];

  it('依頼された11項目がその順番で並ぶ', () => {
    const ss = makeSS();
    g.generateWorkerDetailSheet_(ss, records);
    const rows = ss.sheets[g.DETAIL_SHEET].written;
    expect(rows[0].slice(0, 11)).toEqual([
      '日付', '支店', '元請', '現場', '作業員', '人工数', '勤務区分',
      '夜勤手当対象', '夜勤請求対象', '予定ID', 'メモ'
    ]);
  });

  it('勤務区分・手当・請求が行ごとに出る', () => {
    const ss = makeSS();
    g.generateWorkerDetailSheet_(ss, records);
    const rows = ss.sheets[g.DETAIL_SHEET].written;
    const tanaka = rows.find(r => r[4] === '田中');
    expect(tanaka[6]).toBe('夜勤');
    expect(tanaka[7]).toBe('○');    // 夜勤手当対象
    expect(tanaka[8]).toBe('○');    // 夜勤請求対象
    expect(tanaka[14]).toBe('');    // 時刻が入っているので要確認ではない

    const suzuki = rows.find(r => r[4] === '鈴木');
    expect(suzuki[6]).toBe('日勤');
    expect(suzuki[7]).toBe('○');    // 日勤だが「対象」と明示したので手当あり
    expect(suzuki[8]).toBe('');     // 請求は自動＝日勤なので対象外
  });

  it('夜勤なのに時刻が無い行は要確認になる（依頼書5）', () => {
    const ss = makeSS();
    g.generateWorkerDetailSheet_(ss, records);
    const rows = ss.sheets[g.DETAIL_SHEET].written;
    const sato = rows.find(r => r[4] === '佐藤');
    expect(sato[14]).toBe('要確認：時刻なし');
  });

  // ★2026-09-05 検品③の指摘P3: 1行目の注意書きを全15列に結合していたため、
  //   画面で見るとき「日付」列だけで横幅を使い切り、他の列が画面外へ出ていた。
  it('★1行目はヘッダー（長い注意書きの結合行を置かない）', () => {
    const ss = makeSS();
    g.generateWorkerDetailSheet_(ss, records);
    const rows = ss.sheets[g.DETAIL_SHEET].written;
    expect(rows[0][0]).toBe('日付');
    expect(String(rows[0][0]).length).toBeLessThan(20);
  });

  it('休み・予定は明細に出さない', () => {
    const ss = makeSS();
    g.generateWorkerDetailSheet_(ss, records);
    const rows = ss.sheets[g.DETAIL_SHEET].written;
    expect(rows.find(r => r[4] === '高橋')).toBeUndefined();
  });
});

describe('元請別請求集計_フィルタ用: 日勤と夜勤を別行にする（依頼書4）', () => {
  // 同一月・同一元請・同一現場・同一作業員で、日勤と夜勤の両方がある
  const records = [
    rec({ name: '田中', date: day(1), yakin: '' }),
    rec({ name: '田中', date: day(2), yakin: '夜勤' })
  ];

  it('勤務区分・夜勤請求の列が増えている', () => {
    const ss = makeSS();
    g.generateBillingFilterSheet_(ss, records);
    const rows = ss.sheets[g.BILLING_FILTER_SHEET].written;
    expect(rows[0].slice(0, 6)).toEqual(['月', '会社名', '現場名', '名前', '勤務区分', '夜勤請求']);
    expect(rows[0][37]).toBe('合計');
  });

  it('同じ人・同じ現場でも日勤と夜勤が別行に分かれる', () => {
    const ss = makeSS();
    g.generateBillingFilterSheet_(ss, records);
    const rows = ss.sheets[g.BILLING_FILTER_SHEET].written;
    const mine = rows.filter(r => r[3] === '田中');
    expect(mine.length).toBe(2);
    expect(mine.map(r => r[4]).sort()).toEqual(['夜勤', '日勤']);
    const nightRow = mine.find(r => r[4] === '夜勤');
    expect(nightRow[5]).toBe('○');       // 夜勤請求対象
    expect(nightRow[37]).toBe(1);
    const dayRow = mine.find(r => r[4] === '日勤');
    expect(dayRow[5]).toBe('');          // 日勤は請求対象外
    expect(dayRow[37]).toBe(1);
  });

  it('合計行は日勤＋夜勤で、改修前と同じ数字になる（既存の請求運用を壊さない）', () => {
    const ss = makeSS();
    g.generateBillingFilterSheet_(ss, records);
    const rows = ss.sheets[g.BILLING_FILTER_SHEET].written;
    const total = rows.find(r => r[3] === '合計');
    expect(total[37]).toBe(2);          // 日勤1 + 夜勤1
    expect(total[4]).toBe('');          // 合計行に区分は付けない（フィルタで拾わない）
  });

  it('夜勤請求を対象外にすると請求列が空になる', () => {
    const ss = makeSS();
    g.generateBillingFilterSheet_(ss, [rec({ date: day(2), yakin: '夜勤', seikyu: '対象外' })]);
    const rows = ss.sheets[g.BILLING_FILTER_SHEET].written;
    const nightRow = rows.find(r => r[4] === '夜勤');
    expect(nightRow[5]).toBe('');
    expect(nightRow[37]).toBe(1);       // 人工は消えない（請求判断と人工は別）
  });

  it('1日に2現場へ行けば夜勤も 0.5 ずつに按分される', () => {
    const ss = makeSS();
    g.generateBillingFilterSheet_(ss, [
      rec({ date: day(2), yakin: '夜勤', loc: 'A現場' }),
      rec({ date: day(2), yakin: '夜勤', loc: 'B現場' })
    ]);
    const rows = ss.sheets[g.BILLING_FILTER_SHEET].written;
    const nightRows = rows.filter(r => r[4] === '夜勤');
    expect(nightRows.length).toBe(2);
    nightRows.forEach(r => expect(r[37]).toBe(0.5));
  });
});

describe('日報データの列', () => {
  it('夜勤手当・夜勤請求が末尾に足されている（既存列の位置は動かさない）', () => {
    expect(g.HEADERS[g.HEADERS.length - 2]).toBe('夜勤手当');
    expect(g.HEADERS[g.HEADERS.length - 1]).toBe('夜勤請求');
    expect(g.HEADERS.indexOf('夜勤')).toBe(10);      // 既存の位置が動いていない
    expect(g.HEADERS.indexOf('部隊')).toBe(20);
  });
});
