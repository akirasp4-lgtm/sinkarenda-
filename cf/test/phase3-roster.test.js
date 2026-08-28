// 同姓同名の人を、どの会社の人として扱うか（2026-08-28）。
//
// ★Codexレビューが実際に動かして見つけた欠陥:
//   activeRosterMembers() が「職人マスタで先に見つかった1件」で会社を決めていた。
//   **奥田さんはグローライズとGRHDの両方に、川端さんはグローライズとラーテルに実在する。**
//   並び順しだいで、グローライズの画面なのに GRHD の奥田さんとして資格を引き、
//   持っているはずの資格が出なくなる。
//
// ★phase3-qual-select.test.js は activeRosterMembers を偽物に置き換えているので、
//   この欠陥を検出できない。ここでは**本物を動かす**。
import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import vm from 'node:vm';

const here = dirname(fileURLToPath(import.meta.url));
const read = (f) => readFileSync(join(here, '..', '..', f), 'utf8');

function pick(src, name) {
  const start = src.indexOf('function ' + name + '(');
  if (start < 0) return null;
  let depth = 0;
  for (let i = src.indexOf('{', start); i < src.length; i++) {
    if (src[i] === '{') depth++;
    else if (src[i] === '}') { depth--; if (depth === 0) return src.slice(start, i + 1); }
  }
  return null;
}

// 本物の activeRosterMembel を、必要最小限の周りだけ用意して動かす
function run(file, { members, company, names }) {
  const src = read(file);
  const fn = pick(src, 'activeRosterMembers');
  if (!fn) throw new Error(file + ' に activeRosterMembers が無い');
  const sandbox = vm.createContext({
    console,
    allMembers: members,
    currentCompany: company,
    // 本物と同じ定義（index.html / admin.html の hasKyotenAxis）
    hasKyotenAxis: (c) => ['グローライズ', 'GRミツマ'].indexOf(String(c || '').trim()) >= 0,
    getActiveShokunin: () => names
  });
  sandbox.globalThis = sandbox;
  vm.runInContext(fn, sandbox, { filename: file });
  return sandbox.activeRosterMembers();
}

const OKUDA_GRHD_FIRST = [
  { name: '奥田', company: 'GRHD' },
  { name: '奥田', company: 'グローライズ' }
];
const OKUDA_GLO_FIRST = [
  { name: '奥田', company: 'グローライズ' },
  { name: '奥田', company: 'GRHD' }
];

describe.each(['index.html', 'admin.html'])('同姓同名の会社の決め方（%s）', (file) => {
  it('★グローライズの画面では、並び順がどちらでもグローライズの奥田さんになる', () => {
    expect(run(file, { members: OKUDA_GRHD_FIRST, company: 'グローライズ', names: ['奥田'] }))
      .toEqual([{ name: '奥田', company: 'グローライズ' }]);
    expect(run(file, { members: OKUDA_GLO_FIRST, company: 'グローライズ', names: ['奥田'] }))
      .toEqual([{ name: '奥田', company: 'グローライズ' }]);
  });

  it('★GRHDの画面では、並び順がどちらでもGRHDの奥田さんになる', () => {
    expect(run(file, { members: OKUDA_GRHD_FIRST, company: 'GRHD', names: ['奥田'] }))
      .toEqual([{ name: '奥田', company: 'GRHD' }]);
    expect(run(file, { members: OKUDA_GLO_FIRST, company: 'GRHD', names: ['奥田'] }))
      .toEqual([{ name: '奥田', company: 'GRHD' }]);
  });

  it('★川端さん（グローライズとラーテル）も同じ', () => {
    const ms = [{ name: '川端（達）', company: 'ラーテル' }, { name: '川端（達）', company: 'グローライズ' }];
    expect(run(file, { members: ms, company: 'グローライズ', names: ['川端（達）'] }))
      .toEqual([{ name: '川端（達）', company: 'グローライズ' }]);
    expect(run(file, { members: ms, company: 'ラーテル', names: ['川端（達）'] }))
      .toEqual([{ name: '川端（達）', company: 'ラーテル' }]);
  });

  it('★グローライズの画面ではGRミツマ所属も「合う」とみなす（1つの名簿）', () => {
    const ms = [{ name: '江頭', company: 'GRHD' }, { name: '江頭', company: 'GRミツマ' }];
    expect(run(file, { members: ms, company: 'グローライズ', names: ['江頭'] }))
      .toEqual([{ name: '江頭', company: 'GRミツマ' }]);
  });

  it('その会社に居ない人でも、名簿に載っていれば何かの会社を返す（空にしない）', () => {
    const ms = [{ name: '奥田', company: 'GRHD' }];
    const out = run(file, { members: ms, company: 'グローライズ', names: ['奥田'] });
    expect(out[0].name).toBe('奥田');
    expect(out[0].company).toBe('GRHD');
  });

  it('全社の画面では並び順どおり（アプリ全体が氏名で動いているのと同じ扱い）', () => {
    expect(run(file, { members: OKUDA_GRHD_FIRST, company: '全社', names: ['奥田'] }))
      .toEqual([{ name: '奥田', company: 'GRHD' }]);
  });

  it('職人マスタに居ない名前でも落ちない', () => {
    expect(run(file, { members: [], company: 'グローライズ', names: ['知らない人'] }))
      .toEqual([{ name: '知らない人', company: '' }]);
  });

  it('★出す人の集合は getActiveShokunin() と完全に同じ（増やさない・減らさない）', () => {
    const ms = [{ name: 'A', company: 'グローライズ' }, { name: 'B', company: 'グローライズ' }];
    const out = run(file, { members: ms, company: 'グローライズ', names: ['A'] });
    expect(out.map(m => m.name)).toEqual(['A']);
  });
});
