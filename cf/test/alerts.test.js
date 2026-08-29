// 毎朝のアラート（依頼文の要件9）。2026-08-29。
//
// ★一番大事なのは2つ:
//   1. 画面の重複判定と**1文字も違わない結果**になること（下の parity の describe）
//   2. **問題が無い日は送らない**こと。毎日必ず届く通知は読まれなくなる
import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import vm from 'node:vm';
import {
  buildAlerts, formatAlertsText, hasProblem, toRecords, findConflicts,
  activeRoster, addDays, qualStatus, rosterKey, usualHeadcount
} from '../src/alerts.js';

const here = dirname(fileURLToPath(import.meta.url));
const read = (f) => readFileSync(join(here, '..', '..', f), 'utf8');

const H = ['登録日時', '作業日', '元請名', '現場名', '氏名', '役割', '出勤', '退勤', '人工', 'メモ',
  '夜勤', '会社', 'ID', '更新者', '色', '事業部', '工番', '作業区分', '車両', '拠点', '部隊'];
const row = (o = {}) => H.map(h => {
  const m = {
    作業日: '2026-08-31', 元請名: 'きんでん東', 現場名: 'A現場', 氏名: 'A', 役割: '代表',
    人工: 1, 会社: 'グローライズ', ID: 'id-' + Math.abs(H.length), 作業区分: '現場作業', 拠点: '本社'
  };
  const v = Object.prototype.hasOwnProperty.call(o, h) ? o[h] : m[h];
  return v === undefined ? '' : v;
});
const payload = (rows, extra = {}) => ({
  headers: H, rows,
  members: [{ name: 'A', company: 'グローライズ', active: true },
            { name: 'B', company: 'グローライズ', active: true },
            { name: 'C', company: 'グローライズ', active: true }],
  genbaMaster: [], jobsites: [], qualifications: [], ...extra
});
const D = '2026-08-31';
const opt = { date: D, today: '2026-08-30', company: '全社' };

describe('★画面の重複判定と結果が完全に一致する（食い違うと通知と画面で件数が違う）', () => {
  // 画面側（index.html の PHASE2 ブロック）を vm で動かし、同じデータを両方に通す
  function screenFindConflicts(nippos) {
    const src = read('index.html');
    const B = '// ===== PHASE2-CONFLICT-RULE:BEGIN =====', E = '// ===== PHASE2-CONFLICT-RULE:END =====';
    const code = src.slice(src.indexOf(B) + B.length, src.indexOf(E));
    const sandbox = vm.createContext({ console, Map, Set, String, Object, Array });
    sandbox.globalThis = sandbox;
    vm.runInContext(code + ';globalThis.__f=findConflicts;', sandbox, { filename: 'index.html' });
    return sandbox.__f(nippos, {});
  }

  const cases = [
    ['同じ人が同じ日に別の現場2つ', [
      row({ 氏名: 'A', 元請名: 'きんでん東', 現場名: 'X' }),
      row({ 氏名: 'A', 元請名: 'きんでん東', 現場名: 'Y' })]],
    ['同じ現場なら重複ではない', [
      row({ 氏名: 'A', 現場名: 'X' }), row({ 氏名: 'A', 現場名: 'X' })]],
    ['昼と夜勤は別枠', [
      row({ 氏名: 'A', 現場名: 'X' }), row({ 氏名: 'A', 現場名: 'Y', 夜勤: '夜勤' })]],
    ['📌予定・休みは対象外', [
      row({ 氏名: 'A', 現場名: 'X' }), row({ 氏名: 'A', 現場名: 'Y', 夜勤: '予定' }),
      row({ 氏名: 'A', 現場名: 'Z', 夜勤: '休み' })]],
    ['現場系でない作業区分は対象外', [
      row({ 氏名: 'A', 現場名: 'X' }), row({ 氏名: 'A', 現場名: 'Y', 作業区分: '事務所' })]],
    ['★グローライズとGRミツマは1つの名簿として見る', [
      row({ 氏名: '高田（関東）', 会社: 'グローライズ', 現場名: 'X' }),
      row({ 氏名: '高田（関東）', 会社: 'GRミツマ', 現場名: 'Y' })]],
    ['会社が違えば別人として扱う', [
      row({ 氏名: '奥田', 会社: 'グローライズ', 現場名: 'X' }),
      row({ 氏名: '奥田', 会社: 'GRHD', 現場名: 'Y' })]],
    ['置局・着打ち・撤去品返却も現場系', [
      row({ 氏名: 'A', 現場名: 'X', 作業区分: '置局' }),
      row({ 氏名: 'A', 現場名: 'Y', 作業区分: '着打ち' })]]
  ];

  cases.forEach(([title, rows]) => {
    it(title, () => {
      const recs = toRecords(payload(rows));
      const mine = findConflicts(recs, {});
      const screen = screenFindConflicts(recs);
      expect(JSON.stringify(mine), title + ' で画面と食い違っている').toBe(JSON.stringify(screen));
    });
  });
});

describe('問題が無ければ送らない', () => {
  it('★何も無い日は空文字（＝送信しない）', () => {
    const a = buildAlerts(payload([row({ 氏名: 'A' })]), opt);
    expect(hasProblem(a)).toBe(false);
    expect(formatAlertsText(a)).toBe('');
  });
  it('重複があれば送る', () => {
    const a = buildAlerts(payload([
      row({ 氏名: 'A', 現場名: 'X' }), row({ 氏名: 'A', 現場名: 'Y' })]), opt);
    expect(hasProblem(a)).toBe(true);
    expect(formatAlertsText(a)).toContain('予定が重なっています');
  });
  it('★「翌日の現場」「空き人員」だけでは送らない（問題ではなくお知らせ）', () => {
    const a = buildAlerts(payload([row({ 氏名: 'A' })]), opt);
    expect(a.siteCount).toBe(1);
    expect(a.freeCount).toBe(2);
    expect(hasProblem(a)).toBe(false);
  });
});

describe('責任者がいない現場（依頼の「人員不足」の代わり）', () => {
  it('代表が1人もいない現場を出す', () => {
    const a = buildAlerts(payload([
      row({ 氏名: 'A', 役割: '同行', 現場名: 'X' }),
      row({ 氏名: 'B', 役割: '同行', 現場名: 'X' })]), opt);
    expect(a.noLead).toHaveLength(1);
    expect(a.noLead[0].loc).toBe('X');
    expect(formatAlertsText(a)).toContain('責任者がいない現場');
  });
  it('代表が1人でもいれば出さない', () => {
    const a = buildAlerts(payload([
      row({ 氏名: 'A', 役割: '代表', 現場名: 'X' }),
      row({ 氏名: 'B', 役割: '同行', 現場名: 'X' })]), opt);
    expect(a.noLead).toHaveLength(0);
  });
  it('事務所・倉庫作業・移動は現場として見ない', () => {
    const a = buildAlerts(payload([
      row({ 氏名: 'A', 役割: '同行', 現場名: '本社', 作業区分: '事務所' })]), opt);
    expect(a.noLead).toHaveLength(0);
    expect(a.siteCount).toBe(0);
  });
});

describe('資格（依頼の「資格不足」の代わり）', () => {
  const q = (name, qual, expires) => ({ name, company: 'グローライズ', qual, kind: '技能講習', expires });
  it('★まもなく切れる物だけ出す。切れている物は出さない', () => {
    // ★切れている物を毎朝出すと、同じ警告が未来永劫出続けて誰も読まなくなる。
    //   利用者判断（2026-08-29）「その期限切れの資格は一旦ほっといていい」
    const a = buildAlerts(payload([row({ 氏名: 'A' })], {
      qualifications: [q('A', '切れてる', '2024-05-31'), q('A', 'もうすぐ', '2026-09-13'),
        q('A', '読めない', '?'), q('A', '期限なし', ''), q('A', 'まだ先', '2030-01-01')]
    }), opt);   // today=2026-08-30 から 2026-09-13 は14日前＝節目
    expect(a.quals.map(x => x.qual)).toEqual(['もうすぐ']);
  });
  it('★[P2] 節目の日だけ出す（60日間ずっと毎朝出さない）', () => {
    // ★Codexレビュー[P2]（2026-08-29）: 「60日以内」だけだと、その人が出る日は
    //   最大60日ぶん毎朝同じ警告が出る。節目（60/30/14/7/3/1/0日前）に絞る。
    const at = (days) => {
      const d = new Date(Date.parse('2026-08-30T00:00:00Z') + days * 86400000)
        .toISOString().slice(0, 10);
      return buildAlerts(payload([row({ 氏名: 'A' })], {
        qualifications: [q('A', '玉掛け', d)] }), opt).quals.length;
    };
    [60, 30, 14, 7, 3, 1, 0].forEach(d => expect(at(d), d + '日前は出すべき').toBe(1));
    [59, 45, 29, 20, 13, 8, 5, 2].forEach(d => expect(at(d), d + '日前は出さない').toBe(0));
  });
  it('★同じ人の同じ資格が2行あっても1回だけ', () => {
    const a = buildAlerts(payload([row({ 氏名: 'A' })], {
      qualifications: [q('A', '玉掛け', '2026-09-13'), q('A', '玉掛け', '2026-09-13')]
    }), opt);
    expect(a.quals).toHaveLength(1);
  });
  it('その日に出ない人の資格は出さない', () => {
    const a = buildAlerts(payload([row({ 氏名: 'A' })], {
      qualifications: [q('B', '玉掛け', '2026-09-10')]
    }), opt);
    expect(a.quals).toHaveLength(0);
  });
});

describe('拠点をまたぐ移動（依頼の「移動時間」の代わり）', () => {
  it('★[P2] 見るのは「その日→翌日」だけ（同じ移動を2朝続けて出さない）', () => {
    // ★Codexレビュー[P2]（2026-08-29）: 前後どちらも見ると、同じ移動が
    //   「明日の分」と「今日の分」で2回通知される。
    //   毎朝 date=明日 で動くので、この向きだけで全部の移動が1回ずつ出る。
    const a = buildAlerts(payload([
      row({ 氏名: 'A', 作業日: D, 拠点: '本社' }),
      row({ 氏名: 'A', 作業日: '2026-09-01', 拠点: '関東支店' })]), opt);
    expect(a.moves).toHaveLength(1);
    expect(a.moves[0]).toMatchObject({ name: 'A', fromKyoten: '本社', toKyoten: '関東支店' });
    expect(formatAlertsText(a)).toContain('拠点をまたぐ移動');
  });
  it('★前の日から来る分は出さない（前日の朝にもう知らせてある）', () => {
    const a = buildAlerts(payload([
      row({ 氏名: 'A', 作業日: '2026-08-30', 拠点: '関東支店' }),
      row({ 氏名: 'A', 作業日: D, 拠点: '本社' })]), opt);
    expect(a.moves).toHaveLength(0);
  });
  it('★「両方」は移動として数えない', () => {
    const a = buildAlerts(payload([
      row({ 氏名: 'A', 作業日: D, 拠点: '両方' }),
      row({ 氏名: 'A', 作業日: '2026-09-01', 拠点: '本社' })]), opt);
    expect(a.moves).toHaveLength(0);
  });
  it('同じ拠点なら出さない', () => {
    const a = buildAlerts(payload([
      row({ 氏名: 'A', 作業日: D, 拠点: '本社' }),
      row({ 氏名: 'A', 作業日: '2026-09-01', 拠点: '本社' })]), opt);
    expect(a.moves).toHaveLength(0);
  });
  it('★2日以上あいていれば出さない（移動する時間があるため）', () => {
    const a = buildAlerts(payload([
      row({ 氏名: 'A', 作業日: D, 拠点: '本社' }),
      row({ 氏名: 'A', 作業日: '2026-09-02', 拠点: '関東支店' })]), opt);
    expect(a.moves).toHaveLength(0);
  });
  it('名簿に載っていない人は出さない', () => {
    const a = buildAlerts(payload([
      row({ 氏名: '知らない人', 作業日: D, 拠点: '本社' }),
      row({ 氏名: '知らない人', 作業日: '2026-09-01', 拠点: '関東支店' })]), opt);
    expect(a.moves).toHaveLength(0);
  });
});

describe('延期・中止なのに人が入っている', () => {
  it('現場マスタが延期なのに予定があれば出す', () => {
    const a = buildAlerts(payload([row({ 氏名: 'A', 現場名: 'X' })], {
      jobsites: [{ genba: 'きんでん東', loc: 'X', status: '延期' }]
    }), opt);
    expect(a.stoppedWithPeople).toHaveLength(1);
    expect(formatAlertsText(a)).toContain('延期・中止なのに人が入っています');
  });
  it('施工中なら出さない', () => {
    const a = buildAlerts(payload([row({ 氏名: 'A', 現場名: 'X' })], {
      jobsites: [{ genba: 'きんでん東', loc: 'X', status: '施工中' }]
    }), opt);
    expect(a.stoppedWithPeople).toHaveLength(0);
  });
});

describe('空き人員・現場の数', () => {
  it('休みの人も倉庫の人も空きに数えない', () => {
    const a = buildAlerts(payload([
      row({ 氏名: 'A' }), row({ 氏名: 'B', 夜勤: '休み' }), row({ 氏名: 'C', 夜勤: '倉庫' })]), opt);
    expect(a.freeCount).toBe(0);
    expect(a.workingCount).toBe(2);   // A と 倉庫のC
  });
  it('📌予定の行は「出ている」と数えない', () => {
    const a = buildAlerts(payload([row({ 氏名: 'A', 夜勤: '予定' })]), opt);
    expect(a.workingCount).toBe(0);
    expect(a.freeCount).toBe(3);
  });
});

describe('会社での絞り込み', () => {
  it('★グローライズを指定するとGRミツマも入る（1つの名簿）', () => {
    expect(rosterKey('GRミツマ')).toBe(rosterKey('グローライズ'));
  });
  it('他社の重複は混ざらない', () => {
    const rows = [row({ 氏名: '元', 会社: '和信カインド', 現場名: 'X' }),
                  row({ 氏名: '元', 会社: '和信カインド', 現場名: 'Y' })];
    const glo = buildAlerts(payload(rows), { ...opt, company: 'グローライズ' });
    const all = buildAlerts(payload(rows), { ...opt, company: '全社' });
    expect(glo.conflicts).toHaveLength(0);
    expect(all.conflicts).toHaveLength(1);
  });
});

describe('こまごました物', () => {
  it('日付の足し算', () => {
    expect(addDays('2026-08-31', 1)).toBe('2026-09-01');
    expect(addDays('2026-09-01', -1)).toBe('2026-08-31');
    expect(addDays('へんな文字', 1)).toBe('');
  });
  it('資格の期限の判定は画面と同じ考え方', () => {
    expect(qualStatus('', '2026-08-30')).toBe('none');
    expect(qualStatus('2026-08-29', '2026-08-30')).toBe('expired');
    expect(qualStatus('2026-08-30', '2026-08-30')).toBe('soon');
    expect(qualStatus('2026-02-31', '2026-08-30')).toBe('unknown');
  });
  it('無効にした人は名簿に入らない', () => {
    expect(activeRoster([{ name: 'A', company: 'グローライズ', active: false },
      { name: 'B', company: 'グローライズ', active: true }], '全社')).toEqual(['B']);
  });
  it('空の応答でも落ちない', () => {
    const a = buildAlerts({ headers: H, rows: [], members: [] }, opt);
    expect(hasProblem(a)).toBe(false);
    expect(formatAlertsText(a)).toBe('');
  });
});


describe('★Codexレビュー[P1]の再発防止（2026-08-29）', () => {
  it('★倉庫の人を「現場」「責任者なし」に数えない', () => {
    // Codexが実際に動かして発見: 「倉庫・同行」1行で
    // 現場1件・責任者なし1件 の誤通知が出ていた
    const a = buildAlerts(payload([
      row({ 氏名: 'A', 役割: '同行', 現場名: '倉庫', 夜勤: '倉庫' })]), opt);
    expect(a.siteCount, '倉庫を現場に数えている').toBe(0);
    expect(a.noLead, '倉庫で責任者なしを出している').toHaveLength(0);
    expect(hasProblem(a)).toBe(false);
  });

  it('★受注が決まっていないのに人が入っている案件を出す（依頼の「未確定案件」）', () => {
    const a = buildAlerts(payload([row({ 氏名: 'A', 現場名: 'X' })], {
      jobsites: [{ genba: 'きんでん東', loc: 'X', status: '見積中' }]
    }), opt);
    expect(a.unconfirmedWithPeople).toHaveLength(1);
    expect(hasProblem(a)).toBe(true);
    expect(formatAlertsText(a)).toContain('受注が決まっていないのに人が入っています');
  });

  it('★見積中でも人が入っていなければ問題にしない（毎朝同じ数字を出さない）', () => {
    const a = buildAlerts(payload([row({ 氏名: 'A', 現場名: 'X' })], {
      jobsites: [{ genba: 'きんでん東', loc: 'ほかの現場', status: '見積中' }]
    }), opt);
    expect(a.unconfirmedWithPeople).toHaveLength(0);
    expect(hasProblem(a)).toBe(false);
    expect(a.unconfirmed).toHaveLength(1);   // 総数はまとめ行に出る
  });

  it('★延期・中止の総数もまとめ行に出す（依頼の「延期案件」）', () => {
    const a = buildAlerts(payload([
      row({ 氏名: 'A', 現場名: 'X' }), row({ 氏名: 'A', 現場名: 'Y' })], {
      jobsites: [{ genba: 'きんでん東', loc: 'Z', status: '延期' }]
    }), opt);
    expect(a.stoppedAll).toBe(1);
    expect(formatAlertsText(a)).toContain('延期・中止の案件 1件');
  });
});

describe('★画面の parseRows と同じ読み方をしている（2026-08-29 Codexレビュー[P2]）', () => {
  it('★「夜勤」列は trim せず完全一致で見る', () => {
    // ★Codexが発見: Workerだけ trim していたため「休み␣」が
    //   Workerでは休み・画面では通常勤務になり、件数が食い違った。
    const recs = toRecords(payload([
      row({ 氏名: 'A', 夜勤: '休み' }), row({ 氏名: 'B', 夜勤: '休み ' }),
      row({ 氏名: 'C', 夜勤: '夜勤' }), row({ 氏名: 'D', 夜勤: '予定' }),
      row({ 氏名: 'E', 夜勤: '倉庫' }), row({ 氏名: 'F', 夜勤: '' })]));
    const by = {};
    recs.forEach(r => { by[r.name] = r; });
    expect(by['A'].yasumi).toBe(true);
    expect(by['B'].yasumi, '空白付きを休みにしてしまっている（画面は通常勤務）').toBe(false);
    expect(by['C'].yakin).toBe(true);
    expect(by['D'].yotei).toBe(true);
    expect(by['E'].souko).toBe(true);
    expect(by['F'].yakin || by['F'].yotei || by['F'].yasumi || by['F'].souko).toBe(false);
  });

  it('★画面(index.html)の判定式と同じ書き方であること', () => {
    // 画面: yakin:String(r['夜勤']||'')==='夜勤' … trimしていない
    const src = read('index.html');
    expect(src).toContain("yasumi:String(r['夜勤']||'')==='休み'");
    const mine = readFileSync(join(here, '..', 'src', 'alerts.js'), 'utf8');
    expect(mine, 'Worker側が trim している').not.toContain("get(r, '夜勤') || '').trim()");
  });
});


// ============================================================
// 人員不足 = 「いつもより人が少ない現場」（2026-08-29）
//
// ★依頼書の「人員不足」。現場マスタに「必要人数」の欄が無いので、
//   その現場の過去の実績から「いつも何人か」を出して比べる。入力ゼロで効く。
// ★一番怖いのは **鳴りすぎ**（毎朝出ると読まれなくなる）と
//   **実績の浅い現場で誤報**（1日しか実績が無い現場に「いつも」は無い）。
//   その2つを重点的に見張る。
// ============================================================

// n日ぶん、同じ現場に people 人ずつ入れる
function history(genba, loc, days) {
  const out = [];
  days.forEach(([ymd, n]) => {
    for (let k = 0; k < n; k++) {
      out.push(row({ 作業日: ymd, 元請名: genba, 現場名: loc, 氏名: '人' + k,
                     役割: k === 0 ? '代表' : '' }));
    }
  });
  return out;
}

describe('人員不足 — いつもより人が少ない現場', () => {
  const past = [['2026-08-01', 4], ['2026-08-02', 4], ['2026-08-03', 4],
                ['2026-08-04', 4], ['2026-08-05', 4]];

  it('いつも4人の現場がその日2人なら知らせる', () => {
    const rows = history('きんでん西', 'SB心斎橋', past.concat([[D, 2]]));
    const a = buildAlerts(payload(rows), opt);
    expect(a.shortStaff).toHaveLength(1);
    expect(a.shortStaff[0]).toMatchObject({ genba: 'きんでん西', loc: 'SB心斎橋', usual: 4, count: 2 });
    expect(hasProblem(a)).toBe(true);
    expect(formatAlertsText(a)).toContain('いつもより人が少ない現場');
    expect(formatAlertsText(a)).toContain('いつも4人 → 2人');
  });

  it('★いつもどおりの人数なら鳴らない（これが鳴ると毎朝出る）', () => {
    const rows = history('きんでん西', 'SB心斎橋', past.concat([[D, 4]]));
    const a = buildAlerts(payload(rows), opt);
    expect(a.shortStaff).toEqual([]);
    expect(hasProblem(a)).toBe(false);
    expect(formatAlertsText(a)).toBe('');
  });

  it('1人少ないだけでは鳴らない（4人→3人。日常の増減で毎朝出てしまう）', () => {
    const rows = history('きんでん西', 'SB心斎橋', past.concat([[D, 3]]));
    expect(buildAlerts(payload(rows), opt).shortStaff).toEqual([]);
  });

  it('★大きい現場で2人減っても鳴らない（10人→8人。割合の判定が効いているか）', () => {
    // ★わざと壊して確認（2026-08-29）: 割合の判定を消してもテストが緑のままだった。
    //   人数差の判定だけが効いていて、割合の判定を誰も見張っていなかった。
    //   10人の現場が8人になるのは日常。ここが鳴ると大きい現場で毎朝出る。
    const big = [['2026-08-01', 10], ['2026-08-02', 10], ['2026-08-03', 10],
                 ['2026-08-04', 10], ['2026-08-05', 10]];
    const rows = history('きんでん西', '大現場', big.concat([[D, 8]]));
    expect(buildAlerts(payload(rows), opt).shortStaff).toEqual([]);
  });

  it('大きい現場でも半分以下なら鳴る（10人→5人）', () => {
    const big = [['2026-08-01', 10], ['2026-08-02', 10], ['2026-08-03', 10],
                 ['2026-08-04', 10], ['2026-08-05', 10]];
    const rows = history('きんでん西', '大現場', big.concat([[D, 5]]));
    const a = buildAlerts(payload(rows), opt);
    expect(a.shortStaff).toHaveLength(1);
    expect(a.shortStaff[0]).toMatchObject({ usual: 10, count: 5 });
  });

  it('★実績が4日しかない現場は判定しない（「いつも」が決められない）', () => {
    const rows = history('きんでん西', '新規現場',
      [['2026-08-01', 4], ['2026-08-02', 4], ['2026-08-03', 4], ['2026-08-04', 4], [D, 1]]);
    expect(buildAlerts(payload(rows), opt).shortStaff).toEqual([]);
  });

  it('★その日が初日の現場は絶対に鳴らない（1日で終わる現場が大多数）', () => {
    const rows = history('きんでん西', '単発現場', [[D, 1]]);
    expect(buildAlerts(payload(rows), opt).shortStaff).toEqual([]);
  });

  it('中央値を使う＝1日だけ大人数を入れた日に引っ張られない', () => {
    // 実績 3,3,3,3,20 → 平均7.2だが中央値3。その日2人でも「いつも3人」で判定
    const rows = history('きんでん西', 'B現場',
      [['2026-08-01', 3], ['2026-08-02', 3], ['2026-08-03', 3], ['2026-08-04', 3],
       ['2026-08-05', 20], [D, 2]]);
    const a = buildAlerts(payload(rows), opt);
    expect(a.shortStaff).toEqual([]);   // 中央値3 × 0.6 = 1.8 → 2人は下回らない
  });

  it('★休み・予定の行は人数に数えない', () => {
    const rows = history('きんでん西', 'SB心斎橋', past).concat([
      row({ 作業日: D, 元請名: 'きんでん西', 現場名: 'SB心斎橋', 氏名: '人0', 役割: '代表' }),
      row({ 作業日: D, 元請名: 'きんでん西', 現場名: 'SB心斎橋', 氏名: '人1', 夜勤: '休み' }),
      row({ 作業日: D, 元請名: 'きんでん西', 現場名: 'SB心斎橋', 氏名: '人2', 夜勤: '予定' })]);
    const a = buildAlerts(payload(rows), opt);
    expect(a.shortStaff).toHaveLength(1);
    expect(a.shortStaff[0].count).toBe(1);
  });

  it('★倉庫・事務所は現場として数えない（過去の集計でも当日でも同じ条件）', () => {
    const rows = history('グローライズ自社', '事務所', past.concat([[D, 1]]))
      .map(r => { const c = r.slice(); c[H.indexOf('作業区分')] = '事務所'; return c; });
    expect(buildAlerts(payload(rows), opt).shortStaff).toEqual([]);
  });

  it('同じ人が同じ日に2行あっても1人と数える', () => {
    const rows = history('きんでん西', 'SB心斎橋', past).concat([
      row({ 作業日: D, 元請名: 'きんでん西', 現場名: 'SB心斎橋', 氏名: '人0', 役割: '代表' }),
      row({ 作業日: D, 元請名: 'きんでん西', 現場名: 'SB心斎橋', 氏名: '人0' })]);
    const a = buildAlerts(payload(rows), opt);
    expect(a.shortStaff[0].count).toBe(1);
  });

  it('usualHeadcount は現場ごとに中央値と日数を返す', () => {
    const recs = toRecords(payload(history('元A', '現場A', past)));
    const u = usualHeadcount(recs);
    const k = Object.keys(u)[0];
    expect(u[k]).toEqual({ usual: 4, days: 5 });
  });
});


// ============================================================
// 人員不足 — Codexレビューが「見張れていない」と指摘した穴（2026-08-29）
// 指摘された9項目をここで固定する。
// ============================================================

// 会社・夜勤・作業区分を変えられる履歴ヘルパー
function hist(genba, loc, days, o = {}) {
  const out = [];
  days.forEach(([ymd, n]) => {
    for (let k = 0; k < n; k++) {
      out.push(row(Object.assign({
        作業日: ymd, 元請名: genba, 現場名: loc,
        氏名: (o.氏名接頭 || '人') + k, 役割: k === 0 ? '代表' : ''
      }, o.row || {})));
    }
  });
  return out;
}
const P5 = [['2026-08-01', 4], ['2026-08-02', 4], ['2026-08-03', 4],
            ['2026-08-04', 4], ['2026-08-05', 4]];

describe('人員不足 — レビュー指摘の穴を塞ぐ', () => {

  it('★判定日より後の予定を「いつも」に入れない（毎朝は翌日を見るので未来が必ず存在する）', () => {
    const rows = hist('きんでん西', '新規現場',
      [['2026-08-01', 4], ['2026-08-02', 4], ['2026-08-03', 4], ['2026-08-04', 4],
       [D, 1], ['2026-09-05', 4]]);
    expect(buildAlerts(payload(rows), opt).shortStaff).toEqual([]);
  });

  it('★未来に大人数の予定があっても「いつも」を押し上げない', () => {
    const rows = hist('きんでん西', 'C現場', P5.concat([[D, 3], ['2026-09-10', 20]]));
    expect(buildAlerts(payload(rows), opt).shortStaff).toEqual([]);
  });

  it('★全社で見るとき、別会社の同じ元請・現場名を合算しない', () => {
    const rows = hist('きんでん西', '同名現場', P5.map(([y]) => [y, 10]),
                      { row: { 会社: '和信カインド' }, 氏名接頭: '和' })
      .concat(hist('きんでん西', '同名現場',
        [['2026-08-01', 2], ['2026-08-02', 2], ['2026-08-03', 2],
         ['2026-08-04', 2], ['2026-08-05', 2], [D, 2]], { 氏名接頭: 'グ' }));
    const a = buildAlerts(payload(rows), { ...opt, company: '全社' });
    expect(a.shortStaff).toEqual([]);
  });

  it('★グローライズとGRミツマは1つの現場として合算する（統合済みのため）', () => {
    const rows = hist('きんでん東', '関東現場',
      [['2026-08-01', 2], ['2026-08-02', 2], ['2026-08-03', 2], ['2026-08-04', 2]],
      { row: { 会社: 'GRミツマ' }, 氏名接頭: 'ミ' })
      .concat(hist('きんでん東', '関東現場', [['2026-08-05', 2]], { 氏名接頭: 'グ' }));
    const recs = toRecords(payload(rows)).filter(r => r.date < D);
    const u = usualHeadcount(recs);
    expect(Object.keys(u)).toHaveLength(1);
    expect(u[Object.keys(u)[0]]).toEqual({ usual: 2, days: 5 });
  });

  it('★昼勤と夜勤を別々に数える（重複判定が昔から別枠なのに合算すると誤報）', () => {
    const rows = hist('きんでん西', '夜あり現場', P5)
      .concat(hist('きんでん西', '夜あり現場', P5, { row: { 夜勤: '夜勤' }, 氏名接頭: '夜' }))
      .concat(hist('きんでん西', '夜あり現場', [[D, 4]]));
    expect(buildAlerts(payload(rows), opt).shortStaff).toEqual([]);
  });

  it('夜勤側だけ人が減っていれば夜勤として知らせる', () => {
    const rows = hist('きんでん西', '夜あり現場', P5)
      .concat(hist('きんでん西', '夜あり現場', P5, { row: { 夜勤: '夜勤' }, 氏名接頭: '夜' }))
      .concat(hist('きんでん西', '夜あり現場', [[D, 4]]))
      .concat(hist('きんでん西', '夜あり現場', [[D, 1]], { row: { 夜勤: '夜勤' }, 氏名接頭: '夜' }));
    const a = buildAlerts(payload(rows), opt);
    expect(a.shortStaff).toHaveLength(1);
    expect(a.shortStaff[0]).toMatchObject({ yakin: true, usual: 4, count: 1 });
    expect(formatAlertsText(a)).toContain('（夜勤）');
  });

  it('★過去側でも 予定・休み・倉庫 を数えない', () => {
    const rows = hist('きんでん西', 'D現場', P5)
      .concat(hist('きんでん西', 'D現場', P5, { row: { 夜勤: '予定' }, 氏名接頭: 'よ' }))
      .concat(hist('きんでん西', 'D現場', P5, { row: { 夜勤: '休み' }, 氏名接頭: 'や' }))
      .concat(hist('きんでん西', 'D現場', P5, { row: { 夜勤: '倉庫' }, 氏名接頭: 'そ' }));
    const recs = toRecords(payload(rows)).filter(r => r.date < D);
    const u = usualHeadcount(recs);
    expect(u[Object.keys(u)[0]]).toEqual({ usual: 4, days: 5 });
  });

  it('★作業区分が「休み」でモード列が空の行も数えない（旧データにある）', () => {
    // ★わざと壊して確認したら、前のテストは「休み」を数えても数えなくても
    //   結果が同じで、この除外を全く見張れていなかった（2026-08-29）。
    //   出るのは1人だけ・残り3人は作業区分「休み」。
    //   休みを数えてしまうと4人＝平常に見えて、鳴るべき日が鳴らなくなる。
    const rows = hist('きんでん西', 'E現場', P5)
      .concat(hist('きんでん西', 'E現場', [[D, 1]]))
      .concat(hist('きんでん西', 'E現場', [[D, 3]],
                   { row: { 作業区分: '休み' }, 氏名接頭: 'や' }));
    const a = buildAlerts(payload(rows), opt);
    expect(a.shortStaff).toHaveLength(1);
    expect(a.shortStaff[0]).toMatchObject({ usual: 4, count: 1 });
  });

  it('作業区分が「休み」の人しか居ない日は、現場そのものが立たない', () => {
    const rows = hist('きんでん西', 'E2現場', P5)
      .concat(hist('きんでん西', 'E2現場', [[D, 4]], { row: { 作業区分: '休み' } }));
    expect(buildAlerts(payload(rows), opt).shortStaff).toEqual([]);
  });

  it('★現場名が「事務所」の行は人員不足の対象外（実データで464行が作業区分その他）', () => {
    const rows = hist('グローライズ自社', '事務所',
      P5.map(([y]) => [y, 9]).concat([[D, 3]]), { row: { 作業区分: 'その他' } });
    expect(buildAlerts(payload(rows), opt).shortStaff).toEqual([]);
  });

  it('「倉庫材料準備」のような本物の作業は外さない（部分一致で消さない）', () => {
    const rows = hist('グローライズ自社', '倉庫材料準備',
      P5.concat([[D, 1]]), { row: { 作業区分: 'その他' } });
    expect(buildAlerts(payload(rows), opt).shortStaff).toHaveLength(1);
  });

  it('★ちょうど60%は鳴らない（境界）', () => {
    const five = [['2026-08-01', 5], ['2026-08-02', 5], ['2026-08-03', 5],
                  ['2026-08-04', 5], ['2026-08-05', 5]];
    expect(buildAlerts(payload(hist('きんでん西', 'F現場', five.concat([[D, 3]]))), opt)
      .shortStaff).toEqual([]);
    expect(buildAlerts(payload(hist('きんでん西', 'F現場', five.concat([[D, 2]]))), opt)
      .shortStaff).toHaveLength(1);
  });

  it('実績が偶数日のときの中央値（真ん中2つの平均）', () => {
    const six = [['2026-08-01', 2], ['2026-08-02', 2], ['2026-08-03', 4],
                 ['2026-08-04', 4], ['2026-08-05', 6], ['2026-08-06', 6]];
    const recs = toRecords(payload(hist('きんでん西', 'G現場', six))).filter(r => r.date < D);
    const u = usualHeadcount(recs);
    expect(u[Object.keys(u)[0]]).toEqual({ usual: 4, days: 6 });
  });

  it('★11件以上でも見出しの件数と本文が食い違わない（ほかN件を出す）', () => {
    let rows = [];
    for (let i = 0; i < 12; i++) {
      rows = rows.concat(hist('きんでん西', '現場' + i, P5.concat([[D, 1]]),
                              { 氏名接頭: 'p' + i + '_' }));
    }
    const a = buildAlerts(payload(rows), opt);
    expect(a.shortStaff).toHaveLength(12);
    const t = formatAlertsText(a);
    expect(t).toContain('いつもより人が少ない現場 12件');
    expect(t).toContain('・ほか 2件');
  });

  it('足りない人数が多い順に並ぶ（10件で切るとき何が落ちるかを決める）', () => {
    const rows = hist('きんでん西', '軽い現場', P5.concat([[D, 2]]))
      .concat(hist('きんでん西', '重い現場',
        [['2026-08-01', 8], ['2026-08-02', 8], ['2026-08-03', 8],
         ['2026-08-04', 8], ['2026-08-05', 8], [D, 1]], { 氏名接頭: 'h' }));
    const a = buildAlerts(payload(rows), opt);
    expect(a.shortStaff.map(s => s.loc)).toEqual(['重い現場', '軽い現場']);
  });
});
