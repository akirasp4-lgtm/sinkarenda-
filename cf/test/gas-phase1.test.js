// gas.js（Google Apps Script用）の「純粋な関数」だけを取り出して試験する。
//
// ★2026-08-27 実測で判明した注意点（この方式でないと動かない）:
//   vm に読み込んでも `const HEADERS = ...` は**コンテキストの属性にならない**
//   （const/let は字句束縛でグローバルオブジェクトに載らない。var と function だけが載る）。
//   そのため gas.js の末尾に「同じ字句スコープのまま外へ出す」1行を足してから実行する。
//   ここに列挙し忘れた名前はテストから見えないので、関数を足したら必ずここにも足すこと。
import { describe, it, expect, beforeAll } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import vm from 'node:vm';

const here = dirname(fileURLToPath(import.meta.url));
const GAS_PATH = join(here, '..', '..', 'gas.js');

const EXPORT_SNIPPET = `
;globalThis.__gas = {
  HEADERS, BUTAI_VALUES,
  normalizeButai_, resolveButai_, normalizeMemberActive_, BUTAI_LEADERS,
  SITE_STATUSES, SITE_STATUS_DONE, normalizeSiteStatus_, isSiteStatusDone_, isCompletedCell_,
  HISTORY_SHEET, HISTORY_HEADERS, HISTORY_MAX_ROWS,
  diffDailyRows_, rowSummary_, rowFullJson_, sortHistoryRows_, fmtDateTime_,
  // ★vm の外で作った Date は中の Date と別物になり instanceof が効かない。
  //   本番（GAS）は同一環境なので起きないが、テストでは中で作る必要がある。
  makeDate_: function (y, m, d, h, mi) { return new Date(y, m, d, h || 0, mi || 0); },
  KNOWN_COMPANIES, fixMojibakeCompany_, mergeMemberRows_
};`;

let ctx;   // sandbox.__gas
beforeAll(() => {
  const code = readFileSync(GAS_PATH, 'utf8');
  // Apps Script のグローバルを最低限だけ用意する（純粋関数の試験が目的）
  const sandbox = vm.createContext({
    SpreadsheetApp: { getActiveSpreadsheet: () => null, flush() {} },
    Session: { getScriptTimeZone: () => 'Asia/Tokyo' },
    LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock() {} }) },
    Utilities: {
      // ★fmtDate_ / fmtTime_ が使う。gas.js は起動時に tzFastOk_() でこれを試し、
      //   期待どおりの文字列が返れば以降は素のDateメソッドで組み立てる（速い経路）。
      //   ここを空にしておくと日付・時刻の変換が例外になり、履歴の突き合わせが壊れる。
      formatDate: (d, tz, fmt) => {
        const p = (n) => String(n).padStart(2, '0');
        return String(fmt)
          .replace('yyyy', d.getFullYear())
          .replace('MM', p(d.getMonth() + 1))
          .replace('dd', p(d.getDate()))
          .replace('HH', p(d.getHours()))
          .replace('mm', p(d.getMinutes()))
          .replace('ss', p(d.getSeconds()));
      }
    },
    ContentService: {}, PropertiesService: {},
    UrlFetchApp: {}, Logger: { log() {} }, console
  });
  vm.runInContext(code + EXPORT_SNIPPET, sandbox, { filename: 'gas.js' });
  ctx = sandbox.__gas;
});

describe('HEADERS', () => {
  it('21列で、21列目が部隊', () => {
    expect(ctx.HEADERS.length).toBe(21);
    expect(ctx.HEADERS[20]).toBe('部隊');
  });

  it('先頭19列は1つも動いていない', () => {
    expect(ctx.HEADERS.slice(0, 19)).toEqual([
      '登録日時', '作業日', '元請名', '現場名', '氏名', '役割', '出勤', '退勤',
      '人工', 'メモ', '夜勤', '会社', 'ID', '更新者', '色', '事業部', '工番', '作業区分', '車両'
    ]);
  });

  it('20列目は拠点のまま', () => {
    expect(ctx.HEADERS[19]).toBe('拠点');
  });
});

describe('normalizeButai_', () => {
  it('1〜4部隊はそのまま通す', () => {
    ['第一部隊', '第二部隊', '第三部隊', '第四部隊', '第五部隊', '第六部隊'].forEach(v =>
      expect(ctx.normalizeButai_(v)).toBe(v));
  });

  it('前後の空白を落とす', () => {
    expect(ctx.normalizeButai_('  第二部隊 ')).toBe('第二部隊');
  });

  it('知らない値は空にする', () => {
    // ★2026-08-27: 旧表記（1部隊〜4部隊）は組織図2026に無いので通さない。
    //   まだ誰にも部隊が入っていない段階で第一〜第六へ切り替えたため、
    //   旧表記のデータは1件も存在しない。
    ['第七部隊', '1部隊', '2部隊', '3部隊', '4部隊', '部隊', 'A班', '1', 1, null, undefined, ''].forEach(v =>
      expect(ctx.normalizeButai_(v)).toBe(''));
  });

  it('部隊長は6人（第六部隊は奥田・利用者確認済み）', () => {
    expect(ctx.BUTAI_LEADERS).toEqual({
      '第一部隊': '中島', '第二部隊': '前﨑', '第三部隊': '東',
      '第四部隊': '鈴木', '第五部隊': '高田', '第六部隊': '奥田'
    });
  });

  it('★組織図の「前崎」ではなく職人マスタの表記「前﨑」を使う（字が違う）', () => {
    expect(ctx.BUTAI_LEADERS['第二部隊']).toBe('前﨑');
    expect(ctx.BUTAI_LEADERS['第二部隊']).not.toBe('前崎');
  });

  it('部隊の値は第一〜第六の6つだけ（組織図2026・方針書Ver1.0）', () => {
    expect(ctx.BUTAI_VALUES).toEqual(['第一部隊','第二部隊','第三部隊','第四部隊','第五部隊','第六部隊']);
  });
});

describe('resolveButai_', () => {
  it('画面が値を送ってきたらそれを使う', () => {
    expect(ctx.resolveButai_({ butai: '第三部隊' }, '第一部隊')).toBe('第三部隊');
  });

  it('★画面が「空欄」を送ってきたら空欄のまま（既定値で上書きしない）', () => {
    // 事務所・休みなど「部隊に属さない」を明示できるようにするため。
    // 拠点で起きたバグ（手で消した値が既定値に戻る）を繰り返さない。
    expect(ctx.resolveButai_({ butai: '' }, '第一部隊')).toBe('');
  });

  it('画面が項目そのものを送ってこなければ職人マスタの既定部隊を使う', () => {
    expect(ctx.resolveButai_({}, '第一部隊')).toBe('第一部隊');
  });

  it('既定部隊も無ければ空', () => {
    expect(ctx.resolveButai_({}, '')).toBe('');
    expect(ctx.resolveButai_({}, undefined)).toBe('');
  });

  it('既定部隊が壊れた値でも空にする', () => {
    expect(ctx.resolveButai_({}, '第九部隊')).toBe('');
  });

  it('画面が送ってきた値が壊れていれば空（既定値へは戻さない）', () => {
    expect(ctx.resolveButai_({ butai: 'A班' }, '第一部隊')).toBe('');
  });

  it('rowがnull/undefinedでも落ちない', () => {
    expect(ctx.resolveButai_(null, '第二部隊')).toBe('第二部隊');
    expect(ctx.resolveButai_(undefined, '')).toBe('');
  });
});

describe('職人の有効/無効', () => {
  it('×だけが無効。それ以外は全部有効', () => {
    ['×', 'x', 'X', '✕'].forEach(v =>
      expect(ctx.normalizeMemberActive_(v)).toBe(false));
    ['○', 'o', '', '　', undefined, null, true].forEach(v =>
      expect(ctx.normalizeMemberActive_(v)).toBe(true));
  });

  it('★空欄は有効（既存71件を巻き込まないための既定）', () => {
    expect(ctx.normalizeMemberActive_('')).toBe(true);
  });
});

describe('案件ステータス（8段階）', () => {
  it('8つちょうど、順番も依頼どおり', () => {
    expect(ctx.SITE_STATUSES).toEqual([
      '見積中', '受注', '準備中', '施工中', '残工事', '完工', '延期', '中止'
    ]);
  });

  it('完了扱いは完工と中止だけ', () => {
    expect(ctx.SITE_STATUS_DONE).toEqual(['完工', '中止']);
  });

  it('保存済みの正しい値はそのまま返す', () => {
    ctx.SITE_STATUSES.forEach(s =>
      expect(ctx.normalizeSiteStatus_(s, false)).toBe(s));
  });

  it('★未設定（空欄）は 完了 列から導く＝既存184件を書き換えずに移行できる', () => {
    expect(ctx.normalizeSiteStatus_('', true)).toBe('完工');
    expect(ctx.normalizeSiteStatus_('', false)).toBe('施工中');
    expect(ctx.normalizeSiteStatus_(undefined, true)).toBe('完工');
  });

  it('★実際に保存されている値は ✓ である（TRUEではない）', () => {
    // gas.js の doGet は String(r[8]||'').trim() !== '' で completed を出しており、
    // 「完了にする」ボタンは '✓' を書く。ここを TRUE だけで見ると
    // 完工済みの7件が全部「施工中」に化ける。
    expect(ctx.normalizeSiteStatus_('', '✓')).toBe('完工');
    expect(ctx.normalizeSiteStatus_('', 'TRUE')).toBe('完工');
    expect(ctx.normalizeSiteStatus_('', '完了')).toBe('完工');
    expect(ctx.normalizeSiteStatus_('', '1')).toBe('完工');
  });

  it('空欄・false・FALSE は完了ではない', () => {
    expect(ctx.normalizeSiteStatus_('', '')).toBe('施工中');
    expect(ctx.normalizeSiteStatus_('', '  ')).toBe('施工中');
    expect(ctx.normalizeSiteStatus_('', false)).toBe('施工中');
    expect(ctx.normalizeSiteStatus_('', 'FALSE')).toBe('施工中');
    expect(ctx.normalizeSiteStatus_('', null)).toBe('施工中');
  });

  it('既存のcompleted判定（空でなければ完了）と完全に一致する', () => {
    // doGet: completed: String(r[8] || '').trim() !== ''
    ['✓', 'TRUE', '完了', '1', 'x', 'あ'].forEach(v => {
      const legacy = String(v || '').trim() !== '';
      expect(ctx.isCompletedCell_(v)).toBe(legacy);
    });
  });

  it('知らない値は 完了 列から導き直す（勝手な値を通さない）', () => {
    expect(ctx.normalizeSiteStatus_('進行中', false)).toBe('施工中');
    expect(ctx.normalizeSiteStatus_('やめた', true)).toBe('完工');
  });

  it('前後の空白を落とす', () => {
    expect(ctx.normalizeSiteStatus_(' 残工事 ', false)).toBe('残工事');
  });

  it('isSiteStatusDone_ が 完了 列に書く値を決める', () => {
    expect(ctx.isSiteStatusDone_('完工')).toBe(true);
    expect(ctx.isSiteStatusDone_('中止')).toBe(true);
    ['見積中', '受注', '準備中', '施工中', '残工事', '延期'].forEach(s =>
      expect(ctx.isSiteStatusDone_(s)).toBe(false));
  });
});

describe('変更履歴 diffDailyRows_', () => {
  const H = () => ctx.HEADERS;
  const mkRow = (over) => {
    const base = {
      '登録日時': '2026/08/27 10:00', '作業日': '2026-08-28', '元請名': 'きんでん西',
      '現場名': 'A現場', '氏名': '元', '役割': '代表', '出勤': '08:00', '退勤': '17:00',
      '人工': 1, 'メモ': '', '夜勤': '', '会社': 'グローライズ', 'ID': 'X1',
      '更新者': '向', '色': '', '事業部': 'INF', '工番': 'INF-26-001',
      '作業区分': '現場作業', '車両': '', '拠点': '本社', '部隊': '第一部隊'
    };
    Object.assign(base, over || {});
    return H().map(h => base[h]);
  };

  it('変わった項目だけを返す', () => {
    const d = ctx.diffDailyRows_(H(), [mkRow({ ID: 'X1' })], [mkRow({ ID: 'X2', '現場名': 'B現場' })]);
    const fields = d.map(x => x.field);
    expect(fields).toContain('現場名');
    expect(fields).not.toContain('元請名');
  });

  it('変更前と変更後の両方が残る（元の予定が確認できる）', () => {
    const d = ctx.diffDailyRows_(H(), [mkRow({ '人工': 1 })], [mkRow({ ID: 'X2', '人工': 0.5 })]);
    const k = d.find(x => x.field === '人工');
    expect(String(k.before)).toBe('1');
    expect(String(k.after)).toBe('0.5');
  });

  it('★旧IDと新IDが繋がる（編集するとIDが変わるため）', () => {
    const d = ctx.diffDailyRows_(H(), [mkRow({ ID: 'OLD' })], [mkRow({ ID: 'NEW', 'メモ': 'あ' })]);
    expect(d[0].oldId).toBe('OLD');
    expect(d[0].newId).toBe('NEW');
  });

  it('登録日時は毎回変わるので履歴に出さない', () => {
    const oldR = mkRow({ '登録日時': '2026/08/27 10:00' });
    const newR = mkRow({ ID: 'X2', '登録日時': '2026/08/27 11:00' });
    expect(ctx.diffDailyRows_(H(), [oldR], [newR]).map(x => x.field)).not.toContain('登録日時');
  });

  it('IDそのものは項目としては出さない（旧ID/新IDの欄で見えるため）', () => {
    const d = ctx.diffDailyRows_(H(), [mkRow({ ID: 'A' })], [mkRow({ ID: 'B', 'メモ': 'x' })]);
    expect(d.map(x => x.field)).not.toContain('ID');
  });

  it('★業務項目に差が無くてもIDが変われば連鎖だけは残す', () => {
    const d = ctx.diffDailyRows_(H(), [mkRow({ ID: 'A' })], [mkRow({ ID: 'B' })]);
    expect(d.length).toBe(1);
    expect(d[0].field).toBe('(ID引継ぎ)');
    expect(d[0].oldId).toBe('A');
    expect(d[0].newId).toBe('B');
  });

  it('本当に何も変わっていなければ（IDも同じなら）空を返す', () => {
    expect(ctx.diffDailyRows_(H(), [mkRow()], [mkRow()])).toEqual([]);
  });

  it('人が増えた（追加された）行は 追加 として出る', () => {
    const d = ctx.diffDailyRows_(H(), [mkRow({ '氏名': '元' })],
      [mkRow({ ID: 'X2', '氏名': '元' }), mkRow({ ID: 'X3', '氏名': '中島' })]);
    const add = d.find(x => x.field === '(追加)');
    expect(add).toBeTruthy();
    expect(add.after).toContain('中島');
  });

  it('人が減った（外された）行は 削除 として出る', () => {
    const d = ctx.diffDailyRows_(H(),
      [mkRow({ '氏名': '元' }), mkRow({ ID: 'X9', '氏名': '中島' })],
      [mkRow({ ID: 'X2', '氏名': '元' })]);
    const del = d.find(x => x.field === '(削除)');
    expect(del).toBeTruthy();
    expect(del.before).toContain('中島');
  });

  it('部隊の変更も拾う', () => {
    const d = ctx.diffDailyRows_(H(), [mkRow({ '部隊': '第一部隊' })],
      [mkRow({ ID: 'X2', '部隊': '第三部隊' })]);
    const k = d.find(x => x.field === '部隊');
    expect(k.before).toBe('第一部隊');
    expect(k.after).toBe('第三部隊');
  });

  it('日付や時刻の型が違っても文字列として比べる（誤検知しない）', () => {
    const d = ctx.diffDailyRows_(H(), [mkRow({ '人工': 1, ID: 'A' })], [mkRow({ '人工': '1', ID: 'A' })]);
    expect(d).toEqual([]);
  });

  it('★同じ人が同じ日に2件ある場合も、2件目の変更を取りこぼさない', () => {
    // 本番に250件ある形（現場＋事務所 など）。連番を鍵に混ぜていないと握りつぶされる。
    const oldRows = [
      mkRow({ ID: 'A1', '現場名': 'A現場', '作業区分': '現場作業' }),
      mkRow({ ID: 'A2', '現場名': '事務所', '作業区分': '事務所' })
    ];
    const newRows = [
      mkRow({ ID: 'B1', '現場名': 'A現場', '作業区分': '現場作業' }),
      mkRow({ ID: 'B2', '現場名': '事務所', '作業区分': '事務所', '人工': 0.5 })
    ];
    const d = ctx.diffDailyRows_(H(), oldRows, newRows);
    const k = d.find(x => x.field === '人工');
    expect(k).toBeTruthy();
    expect(k.oldId).toBe('A2');
    expect(k.newId).toBe('B2');
    expect(d.find(x => x.field === '(削除)')).toBeUndefined();
  });

  it('★同じ人が同じ日に2件→1件に減ったら、減った1件だけを削除として記録する', () => {
    const oldRows = [mkRow({ ID: 'A1', '現場名': 'A現場' }), mkRow({ ID: 'A2', '現場名': '事務所' })];
    const newRows = [mkRow({ ID: 'B1', '現場名': 'A現場' })];
    const d = ctx.diffDailyRows_(H(), oldRows, newRows);
    const del = d.filter(x => x.field === '(削除)');
    expect(del.length).toBe(1);
    expect(del[0].oldId).toBe('A2');
  });
});

describe('変更履歴シートの形', () => {
  it('8列で、依頼どおりの並び', () => {
    expect(ctx.HISTORY_HEADERS).toEqual(
      ['日時', '操作', '旧ID', '新ID', '項目', '変更前', '変更後', '実行者']);
  });

  it('シート名は 変更履歴', () => {
    expect(ctx.HISTORY_SHEET).toBe('変更履歴');
  });

  it('★削除時は21列すべてをJSONで残す（要約では復元できない）', () => {
    const headers = ctx.HEADERS;
    const arr = headers.map((h, i) => 'v' + i);
    const o = JSON.parse(ctx.rowFullJson_(headers, arr));
    expect(Object.keys(o).length).toBe(21);
    expect(o['部隊']).toBe('v20');
    expect(o['人工']).toBe('v8');
  });

  it('rowFullJson_ は null/undefined を空文字にする', () => {
    const o = JSON.parse(ctx.rowFullJson_(['a', 'b'], [null, undefined]));
    expect(o).toEqual({ a: '', b: '' });
  });
});

describe('履歴の取り出し', () => {
  it('新しい順に並べ替える', () => {
    const rows = [
      ['2026/08/25 10:00', 'update', 'A', 'B', 'メモ', '', 'あ', '向'],
      ['2026/08/27 09:00', 'update', 'C', 'D', 'メモ', '', 'い', '元'],
      ['2026/08/26 12:00', 'add', '', 'E', '(新規)', '', 'う', '中島']
    ];
    const out = ctx.sortHistoryRows_(rows);
    expect(out[0][0]).toBe('2026/08/27 09:00');
    expect(out[2][0]).toBe('2026/08/25 10:00');
  });

  it('件数の上限で切る', () => {
    const rows = Array.from({ length: 700 }, (_, i) =>
      ['2026/08/' + String((i % 28) + 1).padStart(2, '0') + ' 10:00', 'update', '', '', 'メモ', '', 'x', '向']);
    expect(ctx.sortHistoryRows_(rows, 500).length).toBe(500);
  });

  it('空でも落ちない', () => {
    expect(ctx.sortHistoryRows_([])).toEqual([]);
    expect(ctx.sortHistoryRows_(null)).toEqual([]);
  });
});

describe('データ掃除', () => {
  const MOJI = '�';   // 文字化けを表す記号

  it('文字化けした会社名を直す', () => {
    expect(ctx.fixMojibakeCompany_('グロ' + MOJI + 'ライズ')).toBe('グローライズ');
    expect(ctx.fixMojibakeCompany_('グロ?ライズ')).toBe('グローライズ');
  });

  it('正しい会社名はそのまま返す', () => {
    ['グローライズ', '和信カインド', 'GRミツマ', 'GRHD', 'ラーテル'].forEach(c =>
      expect(ctx.fixMojibakeCompany_(c)).toBe(c));
  });

  it('関係ない文字列は触らない', () => {
    expect(ctx.fixMojibakeCompany_('よその会社')).toBe('よその会社');
    expect(ctx.fixMojibakeCompany_('')).toBe('');
  });

  it('★手がかりが半分も残っていなければ触らない（推測しない）', () => {
    // 全部化けている＝1文字も手がかりが無い。長さが偶然一致するだけで
    // 会社を決めてはいけない（'?????' は GRミツマ と同じ5文字）。
    const v = MOJI.repeat(5);
    expect(ctx.fixMojibakeCompany_(v)).toBe(v);
    const v2 = 'G' + MOJI.repeat(4);
    expect(ctx.fixMojibakeCompany_(v2)).toBe(v2);
  });

  it('手がかりが半分以上残っていれば直す', () => {
    // 本番で実際に起きた形: 6文字中1文字だけ化けている
    expect(ctx.fixMojibakeCompany_('グロ' + MOJI + 'ライズ')).toBe('グローライズ');
    expect(ctx.fixMojibakeCompany_('和信カイン' + MOJI)).toBe('和信カインド');
  });

  it('★重複行は非空の値を寄せて統合する（先勝ちで捨てない）', () => {
    const r = ctx.mergeMemberRows_([
      ['元', 'グローライズ', '', 0, '', ''],
      ['元', 'グローライズ', 'INF', 25000, '第二部隊', '']
    ]);
    expect(r.conflicts.length).toBe(0);
    expect(r.merged.length).toBe(1);
    expect(r.merged[0][2]).toBe('INF');
    expect(r.merged[0][3]).toBe(25000);      // ★単価を失わない
    expect(r.merged[0][4]).toBe('第二部隊');
  });

  it('★値が食い違ったら統合せず conflicts に出す（勝手に決めない）', () => {
    const r = ctx.mergeMemberRows_([
      ['元', 'グローライズ', 'INF', 25000, '', ''],
      ['元', 'グローライズ', 'ICT', 30000, '', '']
    ]);
    expect(r.conflicts.length).toBeGreaterThan(0);
    expect(r.conflicts[0].name).toBe('元');
  });

  it('★単価の食い違いは必ず conflicts に出す（給料の元数字を機械が選ばない）', () => {
    const r = ctx.mergeMemberRows_([
      ['元', 'グローライズ', 'INF', 25000, '', ''],
      ['元', 'グローライズ', 'INF', 30000, '', '']
    ]);
    expect(r.conflicts.map(c => c.field)).toContain('単価');
  });

  it('会社が違えば別人として扱う（同姓の別会社を潰さない）', () => {
    const r = ctx.mergeMemberRows_([
      ['元', 'グローライズ', 'INF', 25000, '', ''],
      ['元', '和信カインド', '', 0, '', '']
    ]);
    expect(r.merged.length).toBe(2);
    expect(r.conflicts.length).toBe(0);
  });

  it('氏名が空の行は捨てる', () => {
    const r = ctx.mergeMemberRows_([['', 'グローライズ', '', 0, '', ''], ['元', 'グローライズ', '', 0, '', '']]);
    expect(r.merged.length).toBe(1);
  });

  it('元の並び順を保つ', () => {
    const r = ctx.mergeMemberRows_([
      ['中島', 'グローライズ', '', 0, '', ''],
      ['元', 'グローライズ', '', 0, '', ''],
      ['中島', 'グローライズ', 'INF', 0, '', '']
    ]);
    expect(r.merged.map(x => x[0])).toEqual(['中島', '元']);
  });

  it('★「人でない枠」を機械で判定する関数は存在しない（推測しない設計）', () => {
    // 予定が0件の14人には 川端・井上・作本・児玉・杉本仁（兄）・いくや など
    // 実在の職人が多数含まれる。予定が無いことと人でないことは別の話。
    expect(ctx.looksLikeNonPerson_).toBeUndefined();
  });
});

describe('変更履歴 突き合わせ（Codexレビュー[P2]#4#5の追試）', () => {
  const H = () => ctx.HEADERS;
  const mk = (over) => {
    const base = {
      '登録日時': '2026/08/27 10:00', '作業日': '2026-08-28', '元請名': 'きんでん西',
      '現場名': 'A現場', '氏名': '元', '役割': '代表', '出勤': '08:00', '退勤': '17:00',
      '人工': 1, 'メモ': '', '夜勤': '', '会社': 'グローライズ', 'ID': 'X1',
      '更新者': '向', '色': '', '事業部': 'INF', '工番': 'INF-26-001',
      '作業区分': '現場作業', '車両': '', '拠点': '本社', '部隊': '第一部隊'
    };
    Object.assign(base, over || {});
    return H().map(h => base[h]);
  };

  it('★Codexの例: 旧[A現場, 事務所] → 新[事務所] は「A現場を削除」と記録する', () => {
    // 素朴な出現順の鍵だと「A現場→事務所へ変更」＋「事務所を削除」と誤記録された
    const oldRows = [
      mk({ ID: 'A1', '現場名': 'A現場', '作業区分': '現場作業' }),
      mk({ ID: 'A2', '現場名': '事務所', '作業区分': '事務所' })
    ];
    const newRows = [mk({ ID: 'B2', '現場名': '事務所', '作業区分': '事務所' })];
    const d = ctx.diffDailyRows_(H(), oldRows, newRows);
    const del = d.filter(x => x.field === '(削除)');
    expect(del.length).toBe(1);
    expect(del[0].oldId).toBe('A1');            // ★消えたのは A現場
    expect(d.find(x => x.field === '現場名')).toBeUndefined();
  });

  it('★先頭に人が増えても、既存の行が「変更」に化けない', () => {
    const oldRows = [mk({ ID: 'A1', '氏名': '元' })];
    const newRows = [mk({ ID: 'B0', '氏名': '中島' }), mk({ ID: 'B1', '氏名': '元' })];
    const d = ctx.diffDailyRows_(H(), oldRows, newRows);
    const add = d.filter(x => x.field === '(追加)');
    expect(add.length).toBe(1);
    expect(add[0].newId).toBe('B0');
    expect(d.filter(x => x.field === '(削除)').length).toBe(0);
  });

  it('並び順が入れ替わっただけなら何も記録しない（ID引継ぎだけ）', () => {
    const a = mk({ ID: 'A1', '氏名': '元' });
    const b = mk({ ID: 'A2', '氏名': '中島' });
    const a2 = mk({ ID: 'B1', '氏名': '元' });
    const b2 = mk({ ID: 'B2', '氏名': '中島' });
    const d = ctx.diffDailyRows_(H(), [a, b], [b2, a2]);
    expect(d.filter(x => x.field === '(削除)').length).toBe(0);
    expect(d.filter(x => x.field === '(追加)').length).toBe(0);
  });

  it('現場を変えた編集は「現場名 A現場 → B現場」と読める形で残る', () => {
    const d = ctx.diffDailyRows_(H(), [mk({ ID: 'A1', '現場名': 'A現場' })],
      [mk({ ID: 'B1', '現場名': 'B現場' })]);
    const k = d.find(x => x.field === '現場名');
    expect(k).toBeTruthy();
    expect(k.before).toBe('A現場');
    expect(k.after).toBe('B現場');
  });

  it('★シートのDate値と画面の文字列を同じものとして扱う（履歴が全滅しない）', () => {
    // 旧行はシートから読むので Date、新行は画面由来なので文字列になる
    const oldRow = mk({ ID: 'A1' });
    oldRow[H().indexOf('作業日')] = ctx.makeDate_(2026, 7, 28);      // 2026-08-28
    oldRow[H().indexOf('出勤')] = ctx.makeDate_(1899, 11, 30, 8, 0);
    oldRow[H().indexOf('退勤')] = ctx.makeDate_(1899, 11, 30, 17, 0);
    const newRow = mk({ ID: 'B1' });                             // 文字列のまま
    const d = ctx.diffDailyRows_(H(), [oldRow], [newRow]);
    // 中身は同じなので「変わった項目」は出ず、ID引継ぎだけが残るはず
    expect(d.filter(x => x.field === '作業日').length).toBe(0);
    expect(d.filter(x => x.field === '出勤').length).toBe(0);
    expect(d.filter(x => x.field === '(削除)').length).toBe(0);
    expect(d.filter(x => x.field === '(追加)').length).toBe(0);
  });

  it('人工の 1 と "1" と 1.0 を同じ値として扱う', () => {
    const o = mk({ ID: 'A1', '人工': 1 });
    const n = mk({ ID: 'A1', '人工': '1.0' });
    expect(ctx.diffDailyRows_(H(), [o], [n]).filter(x => x.field === '人工').length).toBe(0);
  });

  it('人工が本当に変わったときは記録する', () => {
    const o = mk({ ID: 'A1', '人工': 1 });
    const n = mk({ ID: 'B1', '人工': 0.5 });
    const k = ctx.diffDailyRows_(H(), [o], [n]).find(x => x.field === '人工');
    expect(k.before).toBe('1');
    expect(k.after).toBe('0.5');
  });
});

describe('変更履歴の読みやすさ（2026-08-27 実機テストで見つけた欠陥）', () => {
  const H = () => ctx.HEADERS;

  it('★削除の記録の日付が「2028-12-31」の形で残る（生のDate文字列にしない）', () => {
    // 実機で "Sun Dec 31 2028 00:00:00 GMT+0900 (日本標準時)" になっていた。
    // 人が読めないし、ここから予定を作り直すのも難しい。
    const arr = H().map(h => {
      if (h === '作業日') return ctx.makeDate_(2028, 11, 31);
      if (h === '出勤') return ctx.makeDate_(1899, 11, 30, 8, 0);
      if (h === '退勤') return ctx.makeDate_(1899, 11, 30, 17, 0);
      if (h === '登録日時') return ctx.makeDate_(2026, 7, 27, 18, 45);
      if (h === '氏名') return '元';
      if (h === '部隊') return '第三部隊';
      return '';
    });
    const o = JSON.parse(ctx.rowFullJson_(H(), arr));
    expect(o['作業日']).toBe('2028-12-31');
    expect(o['出勤']).toBe('08:00');
    expect(o['退勤']).toBe('17:00');
    expect(o['登録日時']).not.toContain('GMT');
    expect(o['氏名']).toBe('元');
    expect(o['部隊']).toBe('第三部隊');
    expect(Object.keys(o).length).toBe(21);
  });

  it('文字列で入っている日付はそのまま通す（画面由来の行）', () => {
    const arr = H().map(h => (h === '作業日' ? '2028-12-31' : h === '出勤' ? '08:00' : ''));
    const o = JSON.parse(ctx.rowFullJson_(H(), arr));
    expect(o['作業日']).toBe('2028-12-31');
    expect(o['出勤']).toBe('08:00');
  });

  it('★追加の記録（要約）の日付も読める形にする', () => {
    const arr = H().map(h => {
      if (h === '作業日') return ctx.makeDate_(2028, 11, 31);
      if (h === '氏名') return '元';
      if (h === '元請名') return 'きんでん西';
      if (h === '現場名') return 'A現場';
      if (h === '作業区分') return '現場作業';
      return '';
    });
    const sum = ctx.rowSummary_(H(), arr);
    expect(sum).toContain('2028-12-31');
    expect(sum).not.toContain('GMT');
    expect(sum).toContain('元');
  });
});
