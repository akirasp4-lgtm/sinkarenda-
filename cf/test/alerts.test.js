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
  activeRoster, addDays, qualStatus, rosterKey
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
      qualifications: [q('A', '切れてる', '2024-05-31'), q('A', 'もうすぐ', '2026-09-10'),
        q('A', '読めない', '?'), q('A', '期限なし', ''), q('A', 'まだ先', '2030-01-01')]
    }), opt);
    expect(a.quals.map(x => x.qual)).toEqual(['もうすぐ']);
  });
  it('★同じ人の同じ資格が2行あっても1回だけ', () => {
    const a = buildAlerts(payload([row({ 氏名: 'A' })], {
      qualifications: [q('A', '玉掛け', '2026-09-10'), q('A', '玉掛け', '2026-09-10')]
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
  it('前の日と拠点が違えば出す', () => {
    const a = buildAlerts(payload([
      row({ 氏名: 'A', 作業日: '2026-08-30', 拠点: '関東支店' }),
      row({ 氏名: 'A', 作業日: D, 拠点: '本社' })]), opt);
    expect(a.moves).toHaveLength(1);
    expect(a.moves[0]).toMatchObject({ name: 'A', fromKyoten: '関東支店', toKyoten: '本社' });
    expect(formatAlertsText(a)).toContain('拠点をまたぐ移動');
  });
  it('翌日と拠点が違っても出す', () => {
    const a = buildAlerts(payload([
      row({ 氏名: 'A', 作業日: D, 拠点: '本社' }),
      row({ 氏名: 'A', 作業日: '2026-09-01', 拠点: '関東支店' })]), opt);
    expect(a.moves).toHaveLength(1);
  });
  it('同じ拠点なら出さない', () => {
    const a = buildAlerts(payload([
      row({ 氏名: 'A', 作業日: '2026-08-30', 拠点: '本社' }),
      row({ 氏名: 'A', 作業日: D, 拠点: '本社' })]), opt);
    expect(a.moves).toHaveLength(0);
  });
  it('★2日以上あいていれば出さない（移動する時間があるため）', () => {
    const a = buildAlerts(payload([
      row({ 氏名: 'A', 作業日: '2026-08-28', 拠点: '関東支店' }),
      row({ 氏名: 'A', 作業日: D, 拠点: '本社' })]), opt);
    expect(a.moves).toHaveLength(0);
  });
  it('名簿に載っていない人は出さない', () => {
    const a = buildAlerts(payload([
      row({ 氏名: '知らない人', 作業日: '2026-08-30', 拠点: '関東支店' }),
      row({ 氏名: '知らない人', 作業日: D, 拠点: '本社' })]), opt);
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
