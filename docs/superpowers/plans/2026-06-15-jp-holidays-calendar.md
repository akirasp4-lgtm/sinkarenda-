# 祝日カレンダー表示 実装プラン

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 3画面（index/president/admin）のカレンダーに日本の祝日を「赤い日付＋祝日名」で表示する。

**Architecture:** 各HTMLに自己完結の祝日ブロック（内蔵フォールバック＋holidays-jp自動取得＋localStorageキャッシュ）を埋め込み、`renderCalendar` に「祝日なら赤クラス＋名前」を足す。GAS・スプレッドシートは一切触らない。デプロイはHTMLのpushのみ。

**Tech Stack:** バニラJS（HTML埋め込み）、fetch、localStorage、CSS。回帰テストは Node（`tools/holidays/test_holidays.mjs`）＋ブラウザ目視（Claude Preview）。

**設計書:** `docs/superpowers/specs/2026-06-15-jp-holidays-calendar-design.md`

---

## 前提・共通部品

### 埋め込む祝日ブロック（3ファイル共通・一字一句同じものを各 `renderCalendar` の直前に挿入）

```javascript
// ===== 祝日表示（自動取得＋端末記憶＋内蔵フォールバック）=====
const JPH_FALLBACK = {
  "2025-01-01":"元日","2025-01-13":"成人の日","2025-02-11":"建国記念の日","2025-02-23":"天皇誕生日","2025-02-24":"休日","2025-03-20":"春分の日","2025-04-29":"昭和の日","2025-05-03":"憲法記念日","2025-05-04":"みどりの日","2025-05-05":"こどもの日","2025-05-06":"休日","2025-07-21":"海の日","2025-08-11":"山の日","2025-09-15":"敬老の日","2025-09-23":"秋分の日","2025-10-13":"スポーツの日","2025-11-03":"文化の日","2025-11-23":"勤労感謝の日","2025-11-24":"休日",
  "2026-01-01":"元日","2026-01-12":"成人の日","2026-02-11":"建国記念の日","2026-02-23":"天皇誕生日","2026-03-20":"春分の日","2026-04-29":"昭和の日","2026-05-03":"憲法記念日","2026-05-04":"みどりの日","2026-05-05":"こどもの日","2026-05-06":"休日","2026-07-20":"海の日","2026-08-11":"山の日","2026-09-21":"敬老の日","2026-09-22":"国民の休日","2026-09-23":"秋分の日","2026-10-12":"スポーツの日","2026-11-03":"文化の日","2026-11-23":"勤労感謝の日",
  "2027-01-01":"元日","2027-01-11":"成人の日","2027-02-11":"建国記念の日","2027-02-23":"天皇誕生日","2027-03-21":"春分の日","2027-03-22":"休日","2027-04-29":"昭和の日","2027-05-03":"憲法記念日","2027-05-04":"みどりの日","2027-05-05":"こどもの日","2027-07-19":"海の日","2027-08-11":"山の日","2027-09-20":"敬老の日","2027-09-23":"秋分の日","2027-10-11":"スポーツの日","2027-11-03":"文化の日","2027-11-23":"勤労感謝の日"
};
let JPH = Object.assign({}, JPH_FALLBACK);
const JPH_CACHE_KEY = 'jph_cache_v1';
const JPH_URL = 'https://holidays-jp.github.io/api/v1/date.json';
function jphApplyCache(){
  try{
    const raw = localStorage.getItem(JPH_CACHE_KEY);
    if(!raw) return null;
    const o = JSON.parse(raw);
    if(o && o.data){ JPH = Object.assign({}, JPH_FALLBACK, o.data); return o; }
  }catch(e){}
  return null;
}
function jphFresh(o){
  if(!o || !o.fetchedAt) return false;
  return (Date.now() - o.fetchedAt) < 30*24*60*60*1000; // 30日以内なら再取得不要
}
function loadHolidays(){
  const cached = jphApplyCache();                 // 手元キャッシュを反映（無ければフォールバックのまま）
  if(jphFresh(cached)) return;                     // 新しければ取得しない
  fetch(JPH_URL).then(r=>r.ok?r.json():Promise.reject()).then(data=>{
    if(data && typeof data==='object'){
      JPH = Object.assign({}, JPH_FALLBACK, data);
      try{ localStorage.setItem(JPH_CACHE_KEY, JSON.stringify({fetchedAt:Date.now(), data})); }catch(e){}
      try{ if(typeof renderCalendar==='function') renderCalendar(); }catch(e){}
    }
  }).catch(()=>{ /* 失敗時はキャッシュ/フォールバックのまま */ });
}
loadHolidays();
```

**ポイント**: `JPH` は「日付文字列→祝日名」。`renderCalendar` 側は `JPH[dateStr]` を見るだけ。
初期描画はフォールバック/キャッシュで即出る → 取得完了で `renderCalendar()` を呼び直して最新化。

### CSS（index.html / admin.html 共通）
`.cal-cell.sat .cal-day-num` 行の直後に追加：
```css
.cal-cell.holiday .cal-day-num{color:#E8384F}
.cal-holiday-name{font-size:9px;color:#E8384F;text-align:center;line-height:1.1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
```

### CSS（president.html）
`.day-num.sat{color:#36c}` 行の直後に追加：
```css
.day-num.holiday{color:#d33}
.hol-name{font-size:8px;color:#d33;line-height:1.05;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:100%}
```

---

### Task 1: 回帰テストを書く（先に失敗させる）

**Files:**
- Create: `tools/holidays/test_holidays.mjs`

- [ ] **Step 1: テストを作成**

```javascript
// 祝日表示の回帰テスト: 3 HTML に祝日ブロック/CSS/描画分岐が入っているか検査
import fs from 'node:fs';
const root = new URL('../../', import.meta.url);
let fail = 0;
const ok = (name, cond) => { if(!cond){ console.error('FAIL:', name); fail++; } else console.log('ok :', name); };
const read = f => fs.readFileSync(new URL(f, root), 'utf8');

for(const f of ['index.html','admin.html','president.html']){
  const s = read(f);
  ok(f+': JPH_FALLBACK 定義', /JPH_FALLBACK\s*=/.test(s));
  ok(f+': 元日 2026-01-01', /"2026-01-01"\s*:\s*"元日"/.test(s));
  ok(f+': 海の日 2026-07-20', /"2026-07-20"\s*:\s*"海の日"/.test(s));
  ok(f+': holidays-jp 取得', /holidays-jp\.github\.io/.test(s));
  ok(f+': loadHolidays 定義', /function loadHolidays/.test(s));
  ok(f+': loadHolidays 起動', /\bloadHolidays\(\)/.test(s));
}
for(const f of ['index.html','admin.html']){
  const s = read(f);
  ok(f+": holiday クラス付与", /\+=' holiday'/.test(s));
  ok(f+': cal-holiday-name CSS', /\.cal-holiday-name\{/.test(s));
  ok(f+': 祝日番号 赤 CSS', /\.cal-cell\.holiday \.cal-day-num\{color:#E8384F\}/.test(s));
  ok(f+': 描画で祝日名挿入', /cal-holiday-name">\$\{esc\(holiday\)\}/.test(s));
}
{
  const s = read('president.html');
  ok('president: hol-name CSS', /\.hol-name\{/.test(s));
  ok('president: day-num.holiday CSS', /\.day-num\.holiday\{color:#d33\}/.test(s));
  ok('president: holiday クラス付与', /hol \? ' holiday' : ''/.test(s));
  ok('president: 祝日名 要素追加', /hn\.className = 'hol-name'/.test(s));
}
if(fail){ console.error('\n'+fail+' checks failed'); process.exit(1); }
console.log('\nALL PASS');
```

- [ ] **Step 2: テストを走らせて失敗を確認**

Run: `node "G:/マイドライブ/Claude/予定管理アプリ作成/tools/holidays/test_holidays.mjs"`
Expected: FAIL（多数の `FAIL:` 行＋ `checks failed`。まだ何も実装していないため）

- [ ] **Step 3: コミット**

```bash
cd "G:/マイドライブ/Claude/予定管理アプリ作成"
git add tools/holidays/test_holidays.mjs 2>/dev/null && git commit -m "test: 祝日表示の回帰テストを追加" || echo "git未管理ならスキップ"
```

---

### Task 2: index.html（職人用）に祝日を実装

**Files:**
- Modify: `index.html`（CSS: line 170 直後 / 祝日ブロック: `function renderCalendar(){`（line 2099）の直前 / 描画: line 2117-2151）

- [ ] **Step 1: CSS を追加**

`index.html` の `.cal-cell.sat .cal-day-num{color:#3B82F6}`（line 170）の直後に挿入：
```css
.cal-cell.holiday .cal-day-num{color:#E8384F}
.cal-holiday-name{font-size:9px;color:#E8384F;text-align:center;line-height:1.1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
```

- [ ] **Step 2: 祝日ブロックを挿入**

`index.html` の `function renderCalendar(){`（line 2099）の直前の行に、本プラン冒頭「埋め込む祝日ブロック」を**そのまま**貼る。

- [ ] **Step 3: 描画に祝日分岐を追加**

`index.html` line 2120-2123 を：
```javascript
    let cls='cal-cell';
    if(dow===0)cls+=' sun';
    if(dow===6)cls+=' sat';
    if(isToday)cls+=' today';
```
→ 次に置換（`holiday` 行を追加）：
```javascript
    let cls='cal-cell';
    if(dow===0)cls+=' sun';
    if(dow===6)cls+=' sat';
    const holiday=JPH[dateStr];
    if(holiday)cls+=' holiday';
    if(isToday)cls+=' today';
```

そして line 2148-2151 を：
```javascript
    html+=`<div class="${cls}" onclick="openCalDay('${dateStr}')">
      <div class="cal-day-num">${d}</div>
      ${eventsHtml}
    </div>`;
```
→ 次に置換（祝日名を day-num の直後に挿入）：
```javascript
    html+=`<div class="${cls}" onclick="openCalDay('${dateStr}')">
      <div class="cal-day-num">${d}</div>
      ${holiday?`<div class="cal-holiday-name">${esc(holiday)}</div>`:''}${eventsHtml}
    </div>`;
```

- [ ] **Step 4: テストの index 部分が通ることを確認**

Run: `node "G:/マイドライブ/Claude/予定管理アプリ作成/tools/holidays/test_holidays.mjs"`
Expected: `index.html:` から始まる行がすべて `ok :`（admin/president はまだ FAIL のままで可）

- [ ] **Step 5: コミット**

```bash
cd "G:/マイドライブ/Claude/予定管理アプリ作成"
git add index.html && git commit -m "feat: 職人用カレンダーに祝日表示（赤＋祝日名）" || echo "git未管理ならスキップ"
```

---

### Task 3: admin.html（管理者用）に祝日を実装

**Files:**
- Modify: `admin.html`（CSS: line 169 直後 / 祝日ブロック: `function renderCalendar(){`（line 2190）の直前 / 描画: line 2204, 2225）

- [ ] **Step 1: CSS を追加**

`admin.html` の `.cal-cell.sat .cal-day-num{color:#3B82F6}`（line 169）の直後に挿入：
```css
.cal-cell.holiday .cal-day-num{color:#E8384F}
.cal-holiday-name{font-size:9px;color:#E8384F;text-align:center;line-height:1.1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
```

- [ ] **Step 2: 祝日ブロックを挿入**

`admin.html` の `function renderCalendar(){`（line 2190）の直前の行に、本プラン冒頭「埋め込む祝日ブロック」を**そのまま**貼る。

- [ ] **Step 3: 描画に祝日分岐を追加**

`admin.html` line 2204：
```javascript
    let cls='cal-cell';if(dow===0)cls+=' sun';if(dow===6)cls+=' sat';if(isToday)cls+=' today';
```
→ 置換：
```javascript
    const holiday=JPH[dateStr];
    let cls='cal-cell';if(dow===0)cls+=' sun';if(dow===6)cls+=' sat';if(holiday)cls+=' holiday';if(isToday)cls+=' today';
```

`admin.html` line 2225：
```javascript
    html+=`<div class="${cls}" onclick="openCalDay('${dateStr}')"><div class="cal-day-num">${d}</div>${eventsHtml}</div>`;
```
→ 置換：
```javascript
    html+=`<div class="${cls}" onclick="openCalDay('${dateStr}')"><div class="cal-day-num">${d}</div>${holiday?`<div class="cal-holiday-name">${esc(holiday)}</div>`:''}${eventsHtml}</div>`;
```

- [ ] **Step 4: テストの admin 部分が通ることを確認**

Run: `node "G:/マイドライブ/Claude/予定管理アプリ作成/tools/holidays/test_holidays.mjs"`
Expected: `index.html:` と `admin.html:` の行がすべて `ok :`（president はまだ FAIL で可）

- [ ] **Step 5: コミット**

```bash
cd "G:/マイドライブ/Claude/予定管理アプリ作成"
git add admin.html && git commit -m "feat: 管理者用カレンダーに祝日表示（赤＋祝日名）" || echo "git未管理ならスキップ"
```

---

### Task 4: president.html（社長用）に祝日を実装

**Files:**
- Modify: `president.html`（CSS: line 61 直後 / 祝日ブロック: `function renderCalendar(){`（line 489）の直前 / 描画: line 525-528）

- [ ] **Step 1: CSS を追加**

`president.html` の `.day-num.sat{color:#36c}`（line 61）の直後に挿入：
```css
.day-num.holiday{color:#d33}
.hol-name{font-size:8px;color:#d33;line-height:1.05;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:100%}
```

- [ ] **Step 2: 祝日ブロックを挿入**

`president.html` の `function renderCalendar(){`（line 489）の直前の行に、本プラン冒頭「埋め込む祝日ブロック」を**そのまま**貼る。

- [ ] **Step 3: 描画に祝日分岐を追加**

`president.html` line 525-528：
```javascript
    const numEl = document.createElement('div');
    numEl.className = 'day-num' + (c.wd === 0 ? ' sun' : c.wd === 6 ? ' sat' : '');
    numEl.textContent = c.d;
    cell.appendChild(numEl);
```
→ 置換（祝日クラス＋祝日名要素を追加）：
```javascript
    const numEl = document.createElement('div');
    const hol = JPH[c.dateStr];
    numEl.className = 'day-num' + (c.wd === 0 ? ' sun' : c.wd === 6 ? ' sat' : '') + (hol ? ' holiday' : '');
    numEl.textContent = c.d;
    cell.appendChild(numEl);
    if(hol){
      const hn = document.createElement('div');
      hn.className = 'hol-name';
      hn.textContent = hol;
      cell.appendChild(hn);
    }
```

- [ ] **Step 4: テスト全体が通ることを確認**

Run: `node "G:/マイドライブ/Claude/予定管理アプリ作成/tools/holidays/test_holidays.mjs"`
Expected: 末尾に `ALL PASS`（全行 `ok :`）

- [ ] **Step 5: コミット**

```bash
cd "G:/マイドライブ/Claude/予定管理アプリ作成"
git add president.html && git commit -m "feat: 社長用カレンダーに祝日表示（赤＋祝日名）" || echo "git未管理ならスキップ"
```

---

### Task 5: ブラウザで目視確認（Claude Preview）

このプロジェクトに JS のユニットテスト土台は無い（引き継ぎ.md 記載のとおり）。
最終確認は実画面の目視で行う。

- [ ] **Step 1: プレビュー起動**

`preview_start` で本フォルダを配信し、`index.html` を開く（静的配信が難しければ、利用者にローカルで `index.html` をダブルクリックで開いてもらい確認を依頼）。

- [ ] **Step 2: 既知の祝日を確認**

カレンダーを **2026年7月** に送り、`20日（海の日）` のマスで **日付が赤＋「海の日」表示** を確認。
`preview_console_logs` で **JSエラーが無い** ことを確認。`preview_screenshot` を取得。

- [ ] **Step 3: 既存挙動が無事か確認**

同じ画面で、既存の **予定チップ／土日色／今日マーク** が従来どおり出ていることを確認（祝日が予定表示を壊していない）。

- [ ] **Step 4: 社長画面の小型セルを確認**

`president.html` を開き、祝日マスで赤番号＋極小の祝日名が**はみ出さず**収まることを確認、`preview_screenshot` を取得。

- [ ] **Step 5: （任意）オフライン耐性メモ**

`localStorage` を空にし通信を切った状態でも当年の祝日（フォールバック）が出ることを確認（環境的に難しければスキップ可。フォールバック自体は Task1 のテストで担保済み）。

---

### Task 6: 引き継ぎ更新とデプロイ

**Files:**
- Modify: `引き継ぎ.md`（最上部に当セッションの記録を追記）

- [ ] **Step 1: 引き継ぎ.md を更新**

最上部に「2026-06-15 祝日表示を3画面に追加（自動取得＋フォールバック・GAS不変・HTMLのみ）」の節を追記。
来年以降の保守メモ（「holidays-jp が自動更新するが、念のためフォールバックは数年ごとに延長可」）を残す。

- [ ] **Step 2: 変更を push（GitHub Pages 自動反映）**

```bash
cd "G:/マイドライブ/Claude/予定管理アプリ作成"
git add 引き継ぎ.md && git commit -m "docs: 祝日表示の実装を引き継ぎに記録" || echo "git未管理ならスキップ"
git push origin main || echo "リポジトリ未設定ならローカル(別PC)でpush"
```

- [ ] **Step 3: 本番URLで最終確認**

GitHub Pages 反映後（1-2分）、`index.html?c=グローライズ` を **Ctrl+Shift+R** で開き、2026年7月20日が祝日表示になることを確認。
**GAS の再デプロイは不要**（バックエンド未変更）。

---

## Self-Review メモ

- **Spec 網羅**: 見せ方（赤＋名前）=Task2-4 / 3画面=Task2,3,4 / 自動取得＋キャッシュ＋フォールバック=共通ブロック / GAS不変=Task6 Step3。すべて対応済み。
- **Placeholder**: なし（全コード実体を記載）。
- **整合**: `JPH` / `loadHolidays` / `cal-holiday-name`(index,admin) / `hol-name`(president) / 色 `#E8384F`(index,admin) `#d33`(president) を全タスクで一貫使用。
- **既知の注意**: このフォルダはローカルで git 未管理の場合あり（環境依存）。その時は commit/push をスキップし、別PC（リポジトリ設定済み）で push する運用。
