# フェーズ2（見つける）実装計画

> **作業する人へ:** 1タスクずつ、テストを先に書いて、赤を見てから直すこと。
> 各タスクの終わりで `cd cf && npx vitest run` が**全部緑**になってからコミットする。

**目的:** 「誰がいつ空いているか」「同じ人が2つの現場に入っていないか」「条件で予定を探せるか」を
画面で分かるようにする。

**方針:** **画面だけで完結させる。gas.js・Cloudflare Worker・スプレッドシートは1文字も触らない。**
判定に必要なデータ（作業区分・夜勤列・元請・現場名・氏名・日付）は既に21列に揃っていて、
`allNippos` に全部載っている。サーバを触らなければ、失敗しても画面を戻すだけで済む。

**仕様の出どころ:** `docs/superpowers/specs/2026-08-27-zensha-jinin-haichi-design.md`
の §1.1（重複の判定ルール）と §6（フェーズ2の範囲）。

---

## 全体の制約（全タスク共通）

1. **`class="tab"` を新しく足さない。**
   `switchTab()` は `document.querySelectorAll('.tab').forEach((el,i)=>...)` と
   `tabs=['list','avail','genba','vehicle','jimu']` を**添字で対応**させている（index.html:1957-1958）。
   `.tab` の要素が1つ増えるだけで、下部ナビの選択表示が全部ずれる。
   トグルは既存の `.btn-secondary` + `.active`（`.gm-mode-bar` と同じ形）を使う。

2. **重複と空き人員の母集団は「拠点の絞り込みを無視する」。**
   本社ビューで見ていても、その人が関東の現場と重なっていれば重複。
   ただし**会社（法人）は尊重する** — 和信カインドの「元」とグローライズの人は別人。
   → `filteredNippos()` ではなく、新しい `companyNippos()` を使う。

3. **「詳しく探す」の結果に一括編集・一括削除を出さない。**
   絞り込んだ結果は複数の現場・複数の日にまたがる。そこで一括削除を押せると、
   関係のない予定まで消える。**読むだけ。**

4. **保存の入口（`rows.push`）は1文字も変えない。** フェーズ2は読み取りだけ。

---

## ファイル構成

| ファイル | 何をする |
|---|---|
| `index.html` | 判定ルール（純関数ブロック）・保存時の警告の差し替え・重複バナー・詳しく探す・空き人員 |
| `admin.html` | 判定ルール（同じブロックをそのまま）・保存時の警告の差し替え |
| `cf/test/phase2-conflict.test.js` | 新規。判定ルールを**実際に動かす**テスト（vm） |
| `cf/test/ui-wiring.test.js` | 追記。配線もれの見張り |

**判定ルールは純関数ブロックとして切り出し、目印コメントで囲む。**
テストはその区間だけを取り出して vm で動かす（`gas-phase1.test.js` と同じやり方）。

```
// ===== PHASE2-CONFLICT-RULE:BEGIN =====
（DOMを一切触らない関数だけ）
// ===== PHASE2-CONFLICT-RULE:END =====
```

**index.html と admin.html でこのブロックは1文字も違わないこと**（テストで固定する）。

---

## Task 1: 重複の判定ルールを純関数にする

**Files:**
- Modify: `index.html`（`parseRows` の手前あたり、DOMに触らない場所）
- Modify: `admin.html`（同じブロックを同じ内容で）
- Test: `cf/test/phase2-conflict.test.js`（新規）

**Interfaces（後のタスクが使う名前）:**
- `GENBA_WORKTYPES` : `['現場作業','置局','着打ち','撤去品返却']`
- `isGenbaWork(n)` → boolean
- `conflictBucket(n)` → `'night'` | `'day'`
- `jobKey(n)` → `元請名 + '' + 現場名`
- `findConflicts(nippos, opts)` → `[{date, name, company, jobs:[{genba,loc,ids,butai}]}]`
  - `opts.from` を渡すと `date >= from` の分だけ返す（省略時は全部）
  - 返りは `date` 昇順 → `name` 昇順

- [ ] **Step 1: 失敗するテストを書く**

`cf/test/phase2-conflict.test.js`:

```js
// 人員の重複判定を「実際に動かして」確かめる。
//
// ★なぜvmで動かすか: 画面のコードを正規表現で見張るだけだと
//   「書いてあるが動かない」を通してしまう。判定は数字が出る場所なので実際に動かす。
import { describe, it, expect, beforeAll } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import vm from 'node:vm';

const here = dirname(fileURLToPath(import.meta.url));
const read = (f) => readFileSync(join(here, '..', '..', f), 'utf8');

const BEGIN = '// ===== PHASE2-CONFLICT-RULE:BEGIN =====';
const END = '// ===== PHASE2-CONFLICT-RULE:END =====';

function extract(file) {
  const src = read(file);
  const i = src.indexOf(BEGIN), j = src.indexOf(END);
  if (i < 0 || j < 0) throw new Error(file + ' に判定ルールのブロックが無い');
  return src.slice(i + BEGIN.length, j);
}

// const は vm のコンテキストのプロパティにならないので、同じ塊の末尾で外へ出す
const EXPORT = `
;globalThis.__p2 = { GENBA_WORKTYPES, isGenbaWork, conflictBucket, jobKey, findConflicts };
`;

let P;
beforeAll(() => {
  const sandbox = vm.createContext({ globalThis: {}, console });
  sandbox.globalThis = sandbox;
  vm.runInContext(extract('index.html') + EXPORT, sandbox, { filename: 'index.html' });
  P = sandbox.__p2;
});

const row = (o) => Object.assign({
  date: '2026-09-01', name: '中島', company: 'グローライズ',
  genba: 'きんでん西', loc: 'A現場', workType: '現場作業',
  yakin: false, yasumi: false, yotei: false, souko: false, isGhost: false, id: 'x'
}, o);

describe('作業区分の判定', () => {
  it('現場系の4つだけを現場作業とみなす', () => {
    expect(P.GENBA_WORKTYPES).toEqual(['現場作業', '置局', '着打ち', '撤去品返却']);
    ['現場作業', '置局', '着打ち', '撤去品返却'].forEach(w =>
      expect(P.isGenbaWork(row({ workType: w })), w).toBe(true));
    ['現調', '事務所', '移動', 'カギ借用', '材料引取・検品', '倉庫作業', '休み', 'その他', '前乗り', ''].forEach(w =>
      expect(P.isGenbaWork(row({ workType: w })), w).toBe(false));
  });

  it('前後の空白を落として判定する', () => {
    expect(P.isGenbaWork(row({ workType: ' 現場作業 ' }))).toBe(true);
  });
});

describe('重複の判定（設計書§1.1のルール）', () => {
  it('★同じ人・同じ日・別の現場（現場系）は重複', () => {
    const c = P.findConflicts([
      row({ loc: 'A現場' }), row({ loc: 'B現場' })
    ]);
    expect(c.length).toBe(1);
    expect(c[0].name).toBe('中島');
    expect(c[0].date).toBe('2026-09-01');
    expect(c[0].jobs.length).toBe(2);
  });

  it('同じ現場が2行（責任者と班員）は重複ではない', () => {
    expect(P.findConflicts([row({ name: '中島' }), row({ name: '中島' })]).length).toBe(0);
  });

  it('元請だけ違えば別の現場として数える', () => {
    expect(P.findConflicts([
      row({ genba: 'きんでん西', loc: 'A現場' }),
      row({ genba: 'ナンジョウ', loc: 'A現場' })
    ]).length).toBe(1);
  });

  it('★現場作業＋事務所 は重複ではない（同じ日に両立する）', () => {
    expect(P.findConflicts([
      row({ loc: 'A現場', workType: '現場作業' }),
      row({ loc: '本社', workType: '事務所' })
    ]).length).toBe(0);
  });

  it('★昼と夜勤は別枠。重ならない', () => {
    expect(P.findConflicts([
      row({ loc: 'A現場', yakin: false }),
      row({ loc: 'B現場', yakin: true })
    ]).length).toBe(0);
  });

  it('夜勤どうしが別現場なら重複', () => {
    expect(P.findConflicts([
      row({ loc: 'A現場', yakin: true }),
      row({ loc: 'B現場', yakin: true })
    ]).length).toBe(1);
  });

  it('★「予定」「休み」の行は数えない', () => {
    expect(P.findConflicts([row({ loc: 'A現場' }), row({ loc: 'B現場', yotei: true })]).length).toBe(0);
    expect(P.findConflicts([row({ loc: 'A現場' }), row({ loc: 'B現場', yasumi: true })]).length).toBe(0);
  });

  it('ゴースト行（夜勤の翌日ぶん）は数えない', () => {
    expect(P.findConflicts([row({ loc: 'A現場' }), row({ loc: 'B現場', isGhost: true })]).length).toBe(0);
  });

  it('日が違えば重複ではない', () => {
    expect(P.findConflicts([
      row({ date: '2026-09-01', loc: 'A現場' }),
      row({ date: '2026-09-02', loc: 'B現場' })
    ]).length).toBe(0);
  });

  it('人が違えば重複ではない', () => {
    expect(P.findConflicts([row({ name: '中島' }), row({ name: '東', loc: 'B現場' })]).length).toBe(0);
  });

  it('★会社が違う同姓同名は別人として扱う', () => {
    expect(P.findConflicts([
      row({ name: '元', company: '和信カインド', loc: 'A現場' }),
      row({ name: '元', company: 'ラーテル', loc: 'B現場' })
    ]).length).toBe(0);
  });

  it('3つ重なったら1件にまとめて jobs が3つ', () => {
    const c = P.findConflicts([row({ loc: 'A' }), row({ loc: 'B' }), row({ loc: 'C' })]);
    expect(c.length).toBe(1);
    expect(c[0].jobs.length).toBe(3);
  });

  it('opts.from より前の日は返さない', () => {
    const rows = [
      row({ date: '2026-06-29', loc: 'A' }), row({ date: '2026-06-29', loc: 'B' }),
      row({ date: '2026-09-01', loc: 'A' }), row({ date: '2026-09-01', loc: 'B' })
    ];
    expect(P.findConflicts(rows).length).toBe(2);
    expect(P.findConflicts(rows, { from: '2026-08-27' }).length).toBe(1);
    expect(P.findConflicts(rows, { from: '2026-08-27' })[0].date).toBe('2026-09-01');
  });

  it('日付順→氏名順に並ぶ', () => {
    const c = P.findConflicts([
      row({ date: '2026-09-02', name: '東', loc: 'A' }), row({ date: '2026-09-02', name: '東', loc: 'B' }),
      row({ date: '2026-09-01', name: '中島', loc: 'A' }), row({ date: '2026-09-01', name: '中島', loc: 'B' }),
      row({ date: '2026-09-01', name: '鈴木', loc: 'A' }), row({ date: '2026-09-01', name: '鈴木', loc: 'B' })
    ]);
    expect(c.map(x => x.date + '/' + x.name))
      .toEqual(['2026-09-01/中島', '2026-09-01/鈴木', '2026-09-02/東']);
  });

  it('空の配列でも落ちない', () => {
    expect(P.findConflicts([])).toEqual([]);
    expect(P.findConflicts(null)).toEqual([]);
  });
});

describe('画面2つで判定ルールが1文字も違わないこと', () => {
  it('index.html と admin.html のブロックが同一', () => {
    expect(extract('admin.html')).toBe(extract('index.html'));
  });
});
```

- [ ] **Step 2: 赤を見る**

```bash
cd cf && npx vitest run test/phase2-conflict.test.js
```

期待: `index.html に判定ルールのブロックが無い` で落ちる。

- [ ] **Step 3: index.html に純関数ブロックを足す**

`parseRows` の定義の直前に置く（DOMに触らない場所であればどこでもよい）。

```js
// ===== PHASE2-CONFLICT-RULE:BEGIN =====
// 人員の重複判定（設計書 §1.1・2026-08-27 実データで確定したルール）
//
//   同じ人・同じ日に、作業区分が現場系で、元請＋現場名が異なる行が2件以上ある。
//   昼と夜勤は別枠。夜勤列が「予定」「休み」の行は対象外。
//
// ★このルールに至った経緯（実データ2,664行で数えた）:
//     同じ人・同じ日に2件以上 → 250件（現場＋安全会議など、正常が大半）
//     出退勤の時間が重なる     → 217件（08:00-17:00の定型なのでほぼ全部）
//     人工の合計が1.0超        → 189件（人工は給料の数字。全行1で時間の意味がない）
//     同じ日に別々の現場系      →  47件（うち今日以降12件）★これだけが本物
//   ゆるいルールにすると「毎日200件の警告」になり、誰も読まなくなる。
//
// ★このブロックは admin.html にも同じ物が入っている。片方だけ直さないこと
//   （cf/test/phase2-conflict.test.js が1文字でも違えば落とす）。
const GENBA_WORKTYPES = ['現場作業', '置局', '着打ち', '撤去品返却'];

function isGenbaWork(n) {
  return GENBA_WORKTYPES.indexOf(String((n && n.workType) || '').trim()) >= 0;
}

// 昼と夜勤は同じ日でも別枠。夜勤明けに昼の現場へ行くのは普通にある
function conflictBucket(n) { return (n && n.yakin) ? 'night' : 'day'; }

// 「どの現場か」の同一性。元請が違えば別の現場として数える
function jobKey(n) {
  return String((n && n.genba) || '').trim() + '' + String((n && n.loc) || '').trim();
}

// 重複だけを拾う。数えない行はここで落とす
function countsForConflict(n) {
  if (!n || n.isGhost) return false;      // 夜勤の翌日ぶんの影は実体ではない
  if (n.yotei || n.yasumi) return false;  // 「予定」「休み」は現場に立っていない
  if (!isGenbaWork(n)) return false;      // 事務所・現調・移動などは現場と両立する
  if (!n.date || !n.name) return false;
  return true;
}

function findConflicts(nippos, opts) {
  const from = (opts && opts.from) || '';
  const map = new Map();
  (nippos || []).forEach(function (n) {
    if (!countsForConflict(n)) return;
    if (from && String(n.date) < from) return;
    // 会社を混ぜない。和信カインドの「元」とラーテルの「元」は別人
    const key = [n.date, String(n.company || '').trim(), String(n.name).trim(), conflictBucket(n)].join('');
    if (!map.has(key)) map.set(key, new Map());
    const jobs = map.get(key);
    const jk = jobKey(n);
    if (!jobs.has(jk)) jobs.set(jk, { genba: String(n.genba || ''), loc: String(n.loc || ''), butai: String(n.butai || ''), ids: [] });
    if (n.id) jobs.get(jk).ids.push(n.id);
  });
  const out = [];
  map.forEach(function (jobs, key) {
    if (jobs.size < 2) return;
    const p = key.split('');
    out.push({ date: p[0], company: p[1], name: p[2], bucket: p[3], jobs: Array.from(jobs.values()) });
  });
  out.sort(function (a, b) {
    return a.date < b.date ? -1 : a.date > b.date ? 1
      : a.name < b.name ? -1 : a.name > b.name ? 1 : 0;
  });
  return out;
}
// ===== PHASE2-CONFLICT-RULE:END =====
```

- [ ] **Step 4: admin.html にも同じブロックを入れる**（コピーではなく同一内容）

- [ ] **Step 5: 緑にする**

```bash
cd cf && npx vitest run
```

期待: 430 + 新規 が全部 PASS。

- [ ] **Step 6: コミット**

```bash
git add index.html admin.html cf/test/phase2-conflict.test.js && git commit -m "feat(画面): 人員の重複判定を1つの純関数にする（設計書§1.1のルール）"
```

---

## Task 2: 保存時の警告を新ルールへ差し替える

**Files:**
- Modify: `index.html:2433-2446`（新規登録）/ `index.html:3720-3736`（編集）
- Modify: `admin.html`（同じ2箇所）
- Test: `cf/test/phase2-conflict.test.js` に追記 + `cf/test/ui-wiring.test.js` に追記

**今の何が問題か:** 「同じ日に何か1件でもあれば警告」。実データだと250件が該当し、
その大半（事務所・安全会議・現調との併記）は正常。**毎回出る警告は読まれなくなる。**

**Interfaces:**
- `conflictsIfAdded(existing, candidates)` → `findConflicts(existing.concat(candidates))` のうち
  **candidates が原因で新しく生まれたものだけ** を返す

- [ ] **Step 1: 失敗するテストを書く**（`phase2-conflict.test.js` に追記）

```js
describe('保存しようとしている予定が重複を生むか', () => {
  it('★事務所の予定がある日に現場を入れても警告しない（今までは警告していた）', () => {
    const existing = [row({ loc: '本社', workType: '事務所' })];
    const cand = [row({ loc: 'A現場', id: '' })];
    expect(P.conflictsIfAdded(existing, cand).length).toBe(0);
  });

  it('★別の現場が既にある日に現場を入れたら警告する', () => {
    const existing = [row({ loc: 'A現場' })];
    const cand = [row({ loc: 'B現場', id: '' })];
    const c = P.conflictsIfAdded(existing, cand);
    expect(c.length).toBe(1);
    expect(c[0].name).toBe('中島');
  });

  it('同じ現場に班員として足すだけなら警告しない', () => {
    expect(P.conflictsIfAdded([row({ loc: 'A現場' })], [row({ loc: 'A現場', id: '' })]).length).toBe(0);
  });

  it('★元から重なっていた分は「今回のせい」ではないので出さない', () => {
    const existing = [row({ loc: 'A現場' }), row({ loc: 'B現場' })];
    const cand = [row({ date: '2026-09-05', loc: 'Z現場', id: '' })];
    expect(P.conflictsIfAdded(existing, cand).length).toBe(0);
  });

  it('候補が空なら何も出ない', () => {
    expect(P.conflictsIfAdded([row({ loc: 'A現場' })], []).length).toBe(0);
  });
});
```

- [ ] **Step 2: 赤を見る** → `P.conflictsIfAdded is not a function`

- [ ] **Step 3: `conflictsIfAdded` をブロックに足す**（index.html / admin.html 両方。END の直前）

```js
// 「今から保存しようとしている予定」が新しく重複を生むかだけを見る。
// 元から重なっていた分まで出すと、毎回同じ警告が出て読まれなくなる。
function conflictsIfAdded(existing, candidates) {
  if (!candidates || !candidates.length) return [];
  const sig = function (c) { return c.date + '' + c.company + '' + c.name + '' + c.bucket; };
  const before = new Set(findConflicts(existing).map(sig));
  return findConflicts((existing || []).concat(candidates)).filter(function (c) { return !before.has(sig(c)); });
}
```

- [ ] **Step 4: 新規登録の警告を差し替える**（index.html:2433-2446）

```js
  // ★2026-08-27 フェーズ2: 判定を設計書§1.1のルールに差し替えた。
  //   今までは「同じ日に何か1件でもあれば警告」で、実データ250件が該当していた。
  //   その大半は「現場＋安全会議」「現場＋事務所」で正常。毎回出る警告は読まれない。
  const candidates = [];
  dates.forEach(date => members.forEach(m => candidates.push({
    date, name: m.name, company: currentCompany, genba, loc: location,
    workType, yakin, yasumi, yotei, souko, isGhost: false, id: ''
  })));
  const conflicts = conflictsIfAdded(companyNippos(), candidates);
  if (conflicts.length > 0) {
    const lines = conflicts.map(c =>
      `・${formatDate(c.date)}　${c.name}\n　　${c.jobs.map(j => j.genba + '／' + j.loc).join('\n　　')}`
    ).join('\n');
    const proceed = confirm(`同じ日に別の現場に入っています。\n\n${lines}\n\nそれでも登録しますか？`);
    if (!proceed) return;
  }
```

**注意:** `genba` `location` `workType` `yakin` `yasumi` `yotei` `souko` は、
その関数の中で既に組み立てられている変数名に合わせること（違う名前なら合わせる）。
**`rows.push` には一切触らない。**

- [ ] **Step 5: 編集の警告を差し替える**（index.html:3720-3736）

`excludeIds`（編集中のグループ自身）を母集団から除く点だけが違う:

```js
    const base = companyNippos().filter(n => !excludeIds.has(n.id));
    const conflicts = conflictsIfAdded(base, candidates);
```

- [ ] **Step 6: admin.html の同じ2箇所も差し替える**

- [ ] **Step 7: 見張りを `ui-wiring.test.js` に追記**

```js
describe('保存時の重複警告が新ルールを使っていること（2026-08-27 フェーズ2）', () => {
  ['index.html', 'admin.html'].forEach(f => {
    it(f + ': 古い「同じ日に1件でもあれば警告」が残っていない', () => {
      const src = read(f);
      expect(src).not.toContain('既に予定が入っています');
      expect(src).not.toContain('既に他の予定が入っています');
      expect(src).toContain('conflictsIfAdded(');
    });
    it(f + ': 判定ブロックが入っている', () => {
      expect(read(f)).toContain('// ===== PHASE2-CONFLICT-RULE:BEGIN =====');
    });
  });
});
```

- [ ] **Step 8: 緑にしてコミット**

```bash
cd cf && npx vitest run
```

```bash
git add index.html admin.html cf/test && git commit -m "fix(画面): 保存時の重複警告を実態に合うルールへ差し替える（250件の誤警告を12件に）"
```

---

## Task 3: 重複の一覧をカレンダー画面に出す

**Files:**
- Modify: `index.html`（`screen-list` の見出し下・`renderList()` から呼ぶ）
- Test: `cf/test/ui-wiring.test.js` に追記

**Interfaces:**
- `companyNippos()` → 会社だけで絞った配列（**拠点は無視する**）
- `renderConflictBanner()` → バナーの描画。0件なら要素ごと隠す
- `openConflictList()` → 一覧のモーダル

- [ ] **Step 1: `companyNippos()` を足す**（`filteredNippos()` の隣・index.html:1938付近）

```js
// 重複チェックと空き人員の母集団。
// ★拠点（本社/関東支店）の絞り込みは通さない。
//   本社ビューで見ていても、その人が関東の現場と重なっていれば重複だから。
//   会社（法人）は尊重する。和信カインドの「元」とラーテルの「元」は別人。
function companyNippos() {
  if (currentCompany === '全社') return allNippos;
  if (hasKyotenAxis(currentCompany)) return allNippos.filter(n => hasKyotenAxis(n.company));
  return allNippos.filter(n => n.company === currentCompany);
}
```

- [ ] **Step 2: バナーのHTMLを足す**（`screen-list` の `.page-title-row` の直後）

`.tab` を使わないこと。既存の `.notice-bar` の並びに置く。

```html
<div class="conflict-bar" id="conflict-bar" style="display:none" onclick="openConflictList()"></div>
```

CSS（`.notice-bar` の隣に）:

```css
.conflict-bar{display:none;background:#FCEBEB;color:#A32D2D;border:1px solid #E8384F;border-radius:10px;padding:10px 14px;margin-bottom:12px;font-size:14px;font-weight:600;cursor:pointer;min-height:44px;box-sizing:border-box}
```

- [ ] **Step 3: 描画関数を足して `renderList()` の先頭から呼ぶ**

```js
// 今日以降だけを出す。過ぎた日の重複は直しようがないので出しても意味がない
// （実データでは全47件のうち今日以降が12件）。0件のときはバナーごと消す。
function renderConflictBanner() {
  const el = document.getElementById('conflict-bar');
  if (!el) return;
  const list = findConflicts(companyNippos(), { from: todayStr() });
  if (!list.length) { el.style.display = 'none'; return; }
  el.style.display = 'block';
  el.textContent = `⚠️ 同じ日に別の現場に入っている人が ${list.length} 件あります（タップで一覧）`;
}
```

`todayStr()` が既にあるか確認し、無ければ既存の日付整形関数を使う。

- [ ] **Step 4: 一覧のモーダル**（既存の `.modal-bg` / `.modal` を使う）

1件ずつ「日付・氏名・重なっている現場」を出す。**編集ボタンは出さない**（読むだけ）。

- [ ] **Step 5: 見張りを追記**

```js
describe('重複バナー（2026-08-27 フェーズ2）', () => {
  const src = read('index.html');
  it('★拠点で絞らない母集団を使っている', () => {
    expect(src).toContain('function companyNippos(');
    expect(src).toMatch(/findConflicts\(companyNippos\(\)/);
  });
  it('★今日以降だけを出す', () => {
    expect(src).toMatch(/findConflicts\(companyNippos\(\),\s*\{\s*from:/);
  });
  it('0件のときはバナーを消す', () => {
    expect(src).toMatch(/if\s*\(!list\.length\)\s*\{\s*el\.style\.display\s*=\s*'none'/);
  });
  it('★新しいUIに class="tab" を使っていない（下部ナビの添字がずれる）', () => {
    expect((src.match(/class="tab"/g) || []).length).toBe(4);   // 下部ナビの4個だけ
  });
});
```

※ `class="tab"` の実測値は着手時に数え直して基準値にすること。

- [ ] **Step 6: 緑にしてコミット**

---

## Task 4: 「詳しく探す」（要件2の絞り込み）

**Files:**
- Modify: `index.html`（`screen-genba` に4つ目のモードを足す）
- Test: `cf/test/ui-wiring.test.js` に追記

既に `gmMode` が `'genba' | 'person' | 'pin'` の3つある（index.html:3775）。
**4つ目 `'search'` を足す。** モードのボタンは `.gm-mode-bar` の `.btn-secondary`（`.active` で緑）。

絞り込む軸（設計書§6のとおり）:
社員名 / 部隊 / 拠点 / 元請 / 現場名（部分一致）/ 作業区分 / 日付（から・まで）

- [ ] **Step 1: 見張りテストを先に書く**

```js
describe('詳しく探す（2026-08-27 フェーズ2）', () => {
  const src = read('index.html');
  it('4つ目のモードがある', () => {
    expect(src).toContain("setGmMode('search')");
    expect(src).toContain('gm-filter-search');
  });
  it('★7つの軸がそろっている', () => {
    ['gm-sc-name','gm-sc-butai','gm-sc-kyoten','gm-sc-genba','gm-sc-loc','gm-sc-worktype','gm-sc-from','gm-sc-to']
      .forEach(id => expect(src, id).toContain('id="' + id + '"'));
  });
  it('★結果に一括編集・一括削除を出さない', () => {
    const m = src.match(/function renderSearchResults\(\)[\s\S]*?\n\}/);
    expect(m).toBeTruthy();
    expect(m[0]).not.toContain('gm-delete-bar');
    expect(m[0]).not.toContain('gm-edit-btn');
    expect(m[0]).not.toContain('toggleGmSelect');
  });
});
```

- [ ] **Step 2: 赤を見る → 実装 → 緑 → コミット**

実装の要点:
- `searchNippos()` は `filteredNippos()` を起点にする（探す画面では拠点の絞り込みを**効かせる**。
  重複・空き人員とは目的が違う）
- 部隊は `n.butai`、拠点は `matchKyoten(n, 値)` を再利用
- 現場名だけ部分一致、他は完全一致。空欄の軸は無視
- 結果は `groupNippos()` して日付順。件数を上に出す
- **1件も選ばせない・押させない。読むだけ**

---

## Task 5: 空き人員の名前リスト（要件4）と、延期・中止で空いた人（要件8後半）

**Files:**
- Modify: `index.html`（`screen-avail`）
- Test: `cf/test/ui-wiring.test.js` に追記

今ある物: 週の色分け表（`renderAvailWeek`）と、スマホ用の日ごとリスト（`renderAvailList`）。
**足りない物: 「この日の空きは誰か」を1日ぶんはっきり出すこと**（PCの表だと色しか分からない）。

- [ ] **Step 1: 日付を選ぶ欄と結果の枠を足す**

```html
<input type="date" id="avail-day" onchange="renderAvailDay()">
<div id="avail-day-result"></div>
```

- [ ] **Step 2: `renderAvailDay()`**

```js
// その日に予定が1件も無い人＝空き。
// ★母集団は companyNippos()（拠点で絞らない）。本社ビューで見ていても、
//   関東の現場に入っている人は「空き」ではない。
// ★名簿は getActiveShokunin()（職人マスタの 有効=×  を外した一覧）。
function renderAvailDay() {
  const d = document.getElementById('avail-day').value;
  const box = document.getElementById('avail-day-result');
  if (!d) { box.innerHTML = ''; return; }
  const cn = companyNippos().filter(n => !n.isGhost && n.date === d);
  const busy = new Set(cn.map(n => n.name));
  const free = getActiveShokunin().filter(name => !busy.has(name));
  ...
}
```

休み・倉庫は「空き」ではないので、既存の `renderAvailList` と同じく別枠で出す。

- [ ] **Step 3: 延期・中止で空いた人**

`allJobsites` の `status` が `延期` `中止` の現場に入っている人を、
「予定は入っているが現場は止まっている＝動かせる候補」として別枠で出す。

```js
// 要件8の後半。延期・中止にした現場の人は、予定表の上では埋まって見えるが実際は空く。
function releasedByStatus(dateFrom) { ... }
```

`allJobsites` のキー名（`genba` / `loc` / `status`）は着手時に実物を確認すること。

- [ ] **Step 4: 見張りテスト → 緑 → コミット**

---

## 最後に必ずやること

- [ ] `cd cf && npx vitest run` が全部緑
- [ ] GitHub へ push（画面だけなので **GASのデプロイは不要**）
- [ ] **Chromeで実機確認**（利用者指示「人と同じようにつかうんだよ」）
  - カレンダー: バナーが出る／0件のとき消える／タップで一覧
  - 保存: 事務所の日に現場を入れて**警告が出ないこと**、別現場で**出ること**
  - 詳しく探す: 7つの軸が効く／削除ボタンが無い
  - 空き確認: 日付を選ぶと名前が出る
  - スマホ幅（375px）で崩れない／コンソールエラー0
- [ ] `引き継ぎ.md` に §3.11 として記録
