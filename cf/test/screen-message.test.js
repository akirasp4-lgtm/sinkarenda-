// 画面内メッセージ（showMessage）の見張り。
//
// ★なぜ必要か（2026-09-05）:
//   ブラウザ標準の警告ダイアログ（alert）は画面全体を止めてしまう。
//   ・動作確認をブラウザ操作でやると、そこで固まって文面が読めない。
//   ・事務がOKを押すと文面が消え、エラーの中身を後から確認できない。
//   （実例: 「集計エラー：Failed to fetch」が出たが、履歴が残らず口頭で聞くしかなかった）
//   そこで画面の一番上に出す帯（showMessage）へ全部置き換えた。
//   置き換えは機械的にやったので、1か所でも戻ると同じ困りごとが再発する。ここで見張る。
import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import vm from 'node:vm';

const here = dirname(fileURLToPath(import.meta.url));
const read = (f) => readFileSync(join(here, '..', '..', f), 'utf8');
const FILES = ['index.html', 'admin.html'];

// showMessage を定義している <script> ブロックだけを取り出す
const extract = (src) => {
  const clean = src.replace(/<!--[\s\S]*?-->/g, '');
  const blocks = [...clean.matchAll(/<script(?![^>]*\ssrc=)[^>]*>([\s\S]*?)<\/script>/g)].map(m => m[1]);
  const hit = blocks.filter(b => b.indexOf('function showMessage(') >= 0);
  expect(hit.length, 'showMessage を持つ <script> は1つだけのはず').toBe(1);
  return hit[0];
};

// 画面のかわりになる最小の DOM。本物の showMessage をそのまま動かして中身を見る
const makeDoc = () => {
  const el = (tag) => ({
    tagName: tag, id: '', className: '', textContent: '', type: '',
    onclick: null, attrs: {}, children: [], parentNode: null,
    get firstChild() { return this.children[0] || null; },
    setAttribute(k, v) { this.attrs[k] = String(v); },
    getAttribute(k) { return k in this.attrs ? this.attrs[k] : null; },
    appendChild(c) { c.parentNode = this; this.children.push(c); return c; },
    removeChild(c) {
      const i = this.children.indexOf(c);
      if (i >= 0) { this.children.splice(i, 1); c.parentNode = null; }
      return c;
    },
  });
  const body = el('body');
  const find = (n, id) => {
    if (n.id === id) return n;
    for (const c of n.children) { const r = find(c, id); if (r) return r; }
    return null;
  };
  return {
    body,
    documentElement: body,
    createElement: el,
    getElementById(id) { return find(body, id); },
  };
};

const load = (file) => {
  const timers = [];
  const doc = makeDoc();
  const box = doc.createElement('div');
  box.id = 'app-msg';                       // 本物のHTMLと同じく置き場所は先にある
  doc.body.appendChild(box);
  const sandbox = {
    document: doc,
    console: { error() {} },
    setTimeout: (fn) => { timers.push(fn); return timers.length; },
  };
  vm.createContext(sandbox);
  vm.runInContext(extract(read(file)), sandbox, { filename: file });
  return { sandbox, doc, box, timers };
};

// 帯に出ている文面と色を読みやすい形にする
const shown = (box) => box.children.map(c => ({
  type: c.getAttribute('data-type'),
  text: c.children[1].textContent,
}));

describe.each(FILES)('%s: 警告ダイアログが残っていないこと', (file) => {
  it('alert( の呼び出しが1つも無い', () => {
    expect(read(file)).not.toContain('alert(');
  });

  it('置き場所の div と3色ぶんのCSSがある', () => {
    const src = read(file);
    expect(src).toContain('<div id="app-msg"');
    expect(src).toContain('.app-msg-success{');
    expect(src).toContain('.app-msg-warn{');
    expect(src).toContain('.app-msg-error{');
  });
});

it('index.html と admin.html の showMessage が完全に同じ', () => {
  expect(extract(read('admin.html'))).toBe(extract(read('index.html')));
});

describe.each(FILES)('%s: 文面から色を決める', (file) => {
  const { sandbox } = load(file);
  const guess = sandbox.guessMessageType;

  // 実際に画面で出る文面をそのまま使う
  it.each([
    ['エラー：Failed to fetch', 'error'],
    ['集計エラー：Failed to fetch', 'error'],
    ['削除エラー：通信できませんでした', 'error'],
    ['方式の保存に失敗しました（通信エラー）', 'error'],
    ['✓ 登録しました', 'success'],
    ['集計シートを更新しました！', 'success'],
    ['12件をアーカイブしました', 'success'],
    ['更新完了\n旧: A\n新: B', 'success'],
    ['名前を選択してください', 'warn'],
    ['この予定はまだ送信中です。送信が終わってから編集してください', 'warn'],
    ['開始日が終了日より後です', 'warn'],
    ['削除対象が見つかりませんでした。', 'warn'],
    ['最新データを読み込めていないため実行できません。「更新」を押してから、もう一度お試しください。', 'warn'],
  ])('「%s」→ %s', (text, expected) => {
    expect(guess(text)).toBe(expected);
  });
});

describe.each(FILES)('%s: 帯の出し方', (file) => {
  it('文面は要約せず原文のまま出す', () => {
    const { sandbox, box } = load(file);
    const long = '集計エラー：Failed to fetch\n（回線が切れているか、サーバーが落ちています）';
    sandbox.showMessage(long);
    expect(shown(box)).toEqual([{ type: 'error', text: long }]);
  });

  it('色を指定して呼べる', () => {
    const { sandbox, box } = load(file);
    // 文面から決めると warn になる言い回しでも、指定した色が優先される
    sandbox.showMessage('success', '名前を選択してください');
    sandbox.showMessage('error', '倉庫作業を1つ以上選んでください');
    expect(shown(box).map(m => m.type)).toEqual(['success', 'error']);
  });

  it('★赤（エラー）は自動で消えない。閉じるボタンで消える', () => {
    const { sandbox, box, timers } = load(file);
    sandbox.showMessage('エラー：Failed to fetch');
    timers.forEach(fn => fn());        // 自動で消す仕掛けが動いても
    expect(box.children.length, '赤が消えてしまった').toBe(1);
    box.children[0].children[2].onclick();   // ✕ を押す
    expect(box.children.length).toBe(0);
  });

  it('★黄（警告）も自動で消えない', () => {
    const { sandbox, box, timers } = load(file);
    sandbox.showMessage('名前を選択してください');
    timers.forEach(fn => fn());
    expect(box.children.length).toBe(1);
  });

  it('緑（成功）だけ時間がたつと消える', () => {
    const { sandbox, box, timers } = load(file);
    sandbox.showMessage('✓ 登録しました');
    expect(box.children.length).toBe(1);
    timers.forEach(fn => fn());
    expect(box.children.length).toBe(0);
  });

  it('同じ文面を連打しても1件しか出ない', () => {
    const { sandbox, box } = load(file);
    for (let i = 0; i < 5; i++) sandbox.showMessage('エラー：Failed to fetch');
    expect(box.children.length).toBe(1);
  });

  it('違う文面が続いても5件までしか溜めない', () => {
    const { sandbox, box } = load(file);
    for (let i = 0; i < 8; i++) sandbox.showMessage('エラー：' + i);
    expect(box.children.length).toBe(5);
    expect(shown(box)[0].text, '古いものから消えるはず').toBe('エラー：3');
  });

  it('1件も無いときは帯そのものを隠す', () => {
    const { sandbox, box } = load(file);
    expect(box.className).toBe('');
    sandbox.showMessage('名前を選択してください');
    expect(box.className).toBe('show');
    sandbox.clearMessages();
    expect(box.className).toBe('');
  });

  it('文面はHTMLとして解釈しない（textContent に入れる）', () => {
    const { sandbox, box } = load(file);
    sandbox.showMessage('エラー：<img src=x onerror="alert(1)">');
    expect(box.children[0].children[1].textContent).toBe('エラー：<img src=x onerror="alert(1)">');
  });
});
