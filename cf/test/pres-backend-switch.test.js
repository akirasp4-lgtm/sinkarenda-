// president.html の「読み取り先の切り替え」の検証。
//
// ★これまで社長用の検証台は _local/verify/ にあり .gitignore 対象＝1台のPCにしか
//   存在しなかった（引き継ぎ§3.7の宿題）。ここでは president.html の <script> を
//   Node の vm にそのまま読み込んで vitest から検証する。実装とテストが同じコードを
//   見るので、テストだけが実装と乖離する事故が起きない（sync-guard.js と同じ方針）。
import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import vm from 'node:vm';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..', '..');
const html = readFileSync(join(ROOT, 'president.html'), 'utf8');
const pageJs = [...html.matchAll(/<script>([\s\S]*?)<\/script>/g)].map(m => m[1]).join('\n');
const guardJs = readFileSync(join(ROOT, 'sync-guard.js'), 'utf8');
const queueJs = readFileSync(join(ROOT, 'send-queue.js'), 'utf8');
const REAL_IDS = new Set([...html.matchAll(/id="([^"]+)"/g)].map(m => m[1]));


// ★呼ぶたびに1ms進む時計。Date.now() がたまたま同じ値を返すことで隠れる不具合
//   （beginAttempt と mark が同一ミリ秒なら偶然うまくいく等）を確実に暴くために使う。
function makeTickingDate() {
  let t = Date.now();
  const D = function (...a) { return a.length ? new Date(...a) : new Date(t); };
  D.now = () => ++t;
  D.parse = Date.parse;
  D.UTC = Date.UTC;
  D.prototype = Date.prototype;
  return D;
}

function mkEl(id) {
  const e = { id, _cls: new Set(), value: '', textContent: '', innerHTML: '', dataset: {},
              style: { cssText: '' }, children: [], offsetHeight: 44 };
  e.classList = { add: (...c) => c.forEach(x => e._cls.add(x)), remove: (...c) => c.forEach(x => e._cls.delete(x)),
                  toggle: (c, o) => { o ? e._cls.add(c) : e._cls.delete(c); }, contains: c => e._cls.has(c) };
  e.addEventListener = () => {}; e.appendChild = c => { e.children.push(c); return c; };
  e.insertBefore = c => { e.children.unshift(c); return c; };
  e.querySelectorAll = () => []; e.querySelector = () => null;
  e.focus = () => {}; e.remove = () => {}; e.setAttribute = () => {};
  return e;
}
function mkStorage() {
  const m = new Map();
  return { getItem: k => (m.has(String(k)) ? m.get(String(k)) : null),
           setItem: (k, v) => { m.set(String(k), String(v)); },
           removeItem: k => { m.delete(String(k)); }, clear: () => m.clear(),
           key: i => (Array.from(m.keys())[i] ?? null), get length() { return m.size; } };
}

/**
 * @param {object} o
 *   backendJson: backend.json の中身。null なら取得失敗（＝「不明」）を再現
 *   d1: /api/president の応答を返す関数。例外を投げれば通信失敗を再現
 *   gas: GASの応答を返す関数
 */
function makeApp(o = {}) {
  const els = new Map();
  const getEl = id => { if (els.has(id)) return els.get(id);
                        if (!REAL_IDS.has(id)) return null;
                        const e = mkEl(id); els.set(id, e); return e; };
  const document = { getElementById: getEl, createElement: () => mkEl('created'),
                     querySelectorAll: () => [], querySelector: () => null,
                     addEventListener: () => {}, body: mkEl('body') };
  document.body.appendChild = c => { document.body.children.push(c); if (c.id) els.set(c.id, c); return c; };

  const app = { hits: [] };
  const respond = body => Promise.resolve({
    ok: true, status: 200,
    clone: () => ({ json: () => Promise.resolve(body) }),
    json: () => Promise.resolve(body)
  });

  const sandbox = {
    document, console, localStorage: mkStorage(), sessionStorage: mkStorage(),
    location: { search: '' }, navigator: { userAgent: 'node' },
    setTimeout, clearTimeout, setInterval, clearInterval,
    URLSearchParams, AbortController, AbortSignal, Date: (o.tickingClock ? makeTickingDate() : Date), Math, JSON, Map, Set,
    Array, Object, String, Number, Promise, Error, TextEncoder,
    crypto: { randomUUID: () => 'x'.repeat(32) },
    confirm: () => true, prompt: () => null, alert: () => {},
    fetch: (url, opts) => {
      const u = String(url);
      if (u.includes('backend.json')) {
        app.hits.push('backend.json');
        if (o.backendJson === null || o.backendJson === undefined) {
          return Promise.resolve({ ok: false, json: () => Promise.reject(new Error('取得失敗')) });
        }
        return respond(o.backendJson);
      }
      const body = opts && opts.body ? JSON.parse(opts.body) : {};
      if (u.includes('/api/president')) {
        app.hits.push('d1:' + u);
        app.lastD1Body = body;
        return (o.d1 || (() => { throw new Error('D1未設定'); }))(body, respond);
      }
      if (u.includes('/api/pres-sync')) {
        app.hits.push('pres-sync');
        return (o.presSync || (() => respond({ status: 'ok', rows: 1, skipped: false })))(body, respond);
      }
      if (u.includes('script.google.com')) {
        app.hits.push('gas:' + (body.action || '?'));
        app.lastGasBody = body;
        return (o.gas || (() => respond({ status: 'ok', rows: [] })))(body, respond);
      }
      // ★それ以外（祝日API等）。president.html は起動時に外部の祝日一覧も取りに行くため、
      //   これをGAS呼び出しとして数えると「GASを何回叩いたか」の検証が狂う。
      app.hits.push('other:' + u);
      return respond([]);
    }
  };
  sandbox.addEventListener = () => {}; sandbox.removeEventListener = () => {};
  sandbox.window = sandbox; sandbox.globalThis = sandbox;
  vm.createContext(sandbox);
  const BRIDGE = `\n;globalThis.__T={get state(){return state},set state(v){state=v},
    get presLoadOk(){return presLoadOk},
    get tracker(){return presPreferGasTracker}, get PIN(){return PIN},
    get presQueue(){return presQueue}};`;
  vm.runInContext(guardJs + '\n' + queueJs + '\n' + pageJs + BRIDGE, sandbox, { filename: 'president' });
  app.s = sandbox; app.T = sandbox.__T;
  // ★描画に必要な初期値。入れ忘れると renderCalendar() が例外を投げ、loadEvents の
  //   catch へ落ちて presLoadOk が false のままになる（＝「取れているのに保存できない」
  //   という、製品ではなく模擬側の偽陽性になる）。旧検証台も同じ初期化をしている。
  app.T.state.viewYM = { y: 2026, m: 7 };
  app.T.state.editing = null;
  return app;
}

const D1_CFG = { backend: 'd1', workerUrl: 'https://w.test' };

describe('presFetchList（読み取り先の切り替え）', () => {
  it('backend.json が d1 なら Cloudflare から読む（GASは呼ばない）', async () => {
    const app = makeApp({
      backendJson: D1_CFG,
      d1: (b, respond) => respond({ status: 'ok', rows: [{ ID: 'P1', 'タイトル': 'A' }] })
    });
    const { data, err } = await app.s.presFetchList(false);
    expect(err).toBe(null);
    expect(data.rows).toHaveLength(1);
    expect(app.hits.some(h => h.startsWith('d1:'))).toBe(true);
    expect(app.hits.some(h => h.startsWith('gas:'))).toBe(false);
  });

  it('backend.json が gas ならGASだけを読む（Cloudflareを呼ばない）', async () => {
    const app = makeApp({
      backendJson: { backend: 'gas' },
      gas: (b, respond) => respond({ status: 'ok', rows: [] })
    });
    await app.s.presFetchList(false);
    expect(app.hits.some(h => h.startsWith('d1:'))).toBe(false);
    expect(app.hits.some(h => h.startsWith('gas:'))).toBe(true);
  });

  it('★backend.json が取れない（不明）ときはGASを読む＝安全側', async () => {
    const app = makeApp({ backendJson: null, gas: (b, respond) => respond({ status: 'ok', rows: [] }) });
    await app.s.presFetchList(false);
    expect(app.hits.some(h => h.startsWith('d1:'))).toBe(false);
    expect(app.hits.some(h => h.startsWith('gas:'))).toBe(true);
  });

  it('★Cloudflareがエラーを返したら黙ってGASへ落ちる', async () => {
    const app = makeApp({
      backendJson: D1_CFG,
      d1: (b, respond) => respond({ status: 'error', message: 'まだ取り込みが行われていません' }),
      gas: (b, respond) => respond({ status: 'ok', rows: [{ ID: 'P9', 'タイトル': 'GASの中身' }] })
    });
    const { data, err } = await app.s.presFetchList(false);
    expect(err).toBe(null);
    expect(data.rows[0]['タイトル']).toBe('GASの中身');
    expect(app.hits.some(h => h.startsWith('gas:'))).toBe(true);
  });

  it('★Cloudflareが通信ごと落ちてもGASへ落ちる', async () => {
    const app = makeApp({
      backendJson: D1_CFG,
      d1: () => Promise.reject(new Error('繋がらない')),
      gas: (b, respond) => respond({ status: 'ok', rows: [] })
    });
    const { data, err } = await app.s.presFetchList(false);
    expect(err).toBe(null);
    expect(data.status).toBe('ok');
  });

  it('★GASも2回とも失敗したら err を返す（画面は前回の内容を残す側へ進む）', async () => {
    const app = makeApp({
      backendJson: D1_CFG,
      d1: () => Promise.reject(new Error('D1だめ')),
      gas: () => Promise.reject(new Error('GASもだめ'))
    });
    const { data, err } = await app.s.presFetchList(false);
    expect(data).toBe(null);
    expect(err).toBeTruthy();
    expect(app.hits.filter(h => h.startsWith('gas:'))).toHaveLength(2);   // GASは2回試す
  });

  it('forceGas を指定すればCloudflareを候補にしない', async () => {
    const app = makeApp({
      backendJson: D1_CFG,
      d1: (b, respond) => respond({ status: 'ok', rows: [] }),
      gas: (b, respond) => respond({ status: 'ok', rows: [] })
    });
    await app.s.presFetchList(true);
    expect(app.hits.some(h => h.startsWith('d1:'))).toBe(false);
  });

  it('★PINは本文で送る。URLには載せない（履歴・アクセスログに残さない）', async () => {
    const app = makeApp({
      backendJson: D1_CFG,
      d1: (b, respond) => respond({ status: 'ok', rows: [] })
    });
    await app.s.presFetchList(false);
    expect(app.lastD1Body.pin).toBe(app.T.PIN);
    expect(app.hits.find(h => h.startsWith('d1:'))).not.toContain(app.T.PIN);
  });

  it('応答の rows が配列でなければ受け入れずGASへ落ちる', async () => {
    const app = makeApp({
      backendJson: D1_CFG,
      d1: (b, respond) => respond({ status: 'ok', rows: { おかしい: true } }),
      gas: (b, respond) => respond({ status: 'ok', rows: [] })
    });
    const { data } = await app.s.presFetchList(false);
    expect(Array.isArray(data.rows)).toBe(true);
    expect(app.hits.some(h => h.startsWith('gas:'))).toBe(true);
  });
});

describe('書き込み直後（refreshInBackground）', () => {
  const settle = () => new Promise(r => setTimeout(r, 30));

  it('★同期が確実に成功していないときはGASを読む（書き込み前の古いD1を読まない）', async () => {
    const app = makeApp({
      backendJson: D1_CFG,
      // 同期が「進行中でスキップ」＝確実成功ではない
      presSync: (b, respond) => respond({ status: 'ok', rows: 0, skipped: true }),
      d1: (b, respond) => respond({ status: 'ok', rows: [] }),
      gas: (b, respond) => respond({ status: 'ok', rows: [] })
    });
    app.s.refreshInBackground();
    await settle();
    expect(app.hits).toContain('pres-sync');
    expect(app.hits.some(h => h.startsWith('gas:'))).toBe(true);
    expect(app.hits.some(h => h.startsWith('d1:'))).toBe(false);
  });

  it('★同期に失敗したときもGASを読む', async () => {
    const app = makeApp({
      backendJson: D1_CFG,
      presSync: () => Promise.reject(new Error('同期できない')),
      d1: (b, respond) => respond({ status: 'ok', rows: [] }),
      gas: (b, respond) => respond({ status: 'ok', rows: [] })
    });
    app.s.refreshInBackground();
    await settle();
    expect(app.hits.some(h => h.startsWith('d1:'))).toBe(false);
    expect(app.hits.some(h => h.startsWith('gas:'))).toBe(true);
  });

  it('同期が確実に成功したときはCloudflareを読んでよい', async () => {
    const app = makeApp({
      backendJson: D1_CFG,
      presSync: (b, respond) => respond({ status: 'ok', rows: 3, skipped: false }),
      d1: (b, respond) => respond({ status: 'ok', rows: [] }),
      gas: (b, respond) => respond({ status: 'ok', rows: [] })
    });
    app.s.refreshInBackground();
    await settle();
    expect(app.hits).toContain('pres-sync');
    expect(app.hits.some(h => h.startsWith('d1:'))).toBe(true);
  });

  it('★一度GAS優先になったら、次の読み取りも確認が取れるまでGASのまま', async () => {
    const app = makeApp({
      backendJson: D1_CFG,
      presSync: (b, respond) => respond({ status: 'ok', rows: 0, skipped: true }),  // 確実成功ではない
      d1: (b, respond) => respond({ status: 'ok', rows: [] }),
      gas: (b, respond) => respond({ status: 'ok', rows: [] })
    });
    app.s.refreshInBackground();
    await settle();
    app.hits.length = 0;
    // forceGas を渡さない普通の読み取りでも、trackerがブロックしているのでGASを読む
    await app.s.presFetchList(false);
    expect(app.hits.some(h => h.startsWith('d1:'))).toBe(false);
    expect(app.hits.some(h => h.startsWith('gas:'))).toBe(true);
  });
});

// ★③（通信が全部だめでも画面を空にしない）の検証をここへ恒久化する。
//   これまで _local/verify/verify_pres.js にしか無く .gitignore 対象だった。
//   さらに、その検証台は「fetch 1回につき用意した応答を1つ消費する」作りのため、
//   backend.json の取得が1つ食う今回の変更で誤検知を出した（製品側の欠陥ではない）。
//   URLで振り分ける下の模擬なら、その取り違えが起きない。
describe('通信が全部だめなとき（画面を空にしない）', () => {
  const settle = () => new Promise(r => setTimeout(r, 30));

  it('★D1もGASも全滅しても、前回の予定を消さず・保存も禁止のまま・赤い帯を出す', async () => {
    const app = makeApp({
      backendJson: D1_CFG,
      d1: () => Promise.reject(new Error('HTTP 404')),
      gas: () => Promise.reject(new Error('HTTP 404'))
    });
    app.T.state.events = [{ id: 'P1', title: '前回の予定', startDate: '2026-08-20', color: '#36c' }];
    await app.s.loadEvents();
    await settle();

    expect(app.T.state.events).toHaveLength(1);          // 消えていない
    expect(app.T.state.events[0].title).toBe('前回の予定');
    expect(app.T.presLoadOk).toBe(false);                // 保存は禁止のまま＝二重登録を防ぐ
    const badge = app.s.document.body.children.find(c => c.id === 'pres-stale-badge');
    expect(badge).toBeTruthy();                          // 赤い帯が出る
    expect(badge.style.background).toBe('#E8384F');
  });

  it('★1回目が失敗しても2回目で取れれば、利用者には何も見せずに復帰する', async () => {
    let n = 0;
    const app = makeApp({
      backendJson: null,                                  // D1は使わない（GASのみ）
      gas: (b, respond) => {
        n++;
        if (n === 1) return Promise.reject(new Error('HTTP 404'));
        return respond({ status: 'ok', rows: [{ ID: 'P9', 'タイトル': '本物の予定', '開始日': '2026-08-26' }] });
      }
    });
    await app.s.loadEvents();
    await settle();
    expect(app.T.state.events).toHaveLength(1);
    expect(app.T.state.events[0].title).toBe('本物の予定');
    expect(app.T.presLoadOk).toBe(true);                 // 復帰したので保存できる
  });
});

describe('社長用のtrackerが社員用と混ざらないこと', () => {
  it('保存キーが社員用と別（同じ端末で職人の書き込みに引きずられない）', () => {
    const app = makeApp({ backendJson: D1_CFG });
    app.T.tracker.mark();
    const keys = [];
    for (let i = 0; i < app.s.localStorage.length; i++) keys.push(app.s.localStorage.key(i));
    expect(keys).toContain('pres-prefer-gas-v1');
  });
});

// ★2026-08-26 Codexレビュー[P1]#1 で見つかった欠陥の再発防止。
describe('書き込みがD1へ確実に伝わること（Codexレビュー[P1]#1）', () => {
  const settle = () => new Promise(r => setTimeout(r, 40));

  it('★同期の結果を待つ前に、その場でGAS優先を立てる（待っている間に別タブ・再読込が古いD1を読まない）', async () => {
    let releaseSync;
    const app = makeApp({
      backendJson: D1_CFG,
      // 同期は「まだ返ってこない」状態にする＝書き込み直後の数秒間を再現
      presSync: () => new Promise(r => { releaseSync = r; }),
      d1: (b, respond) => respond({ status: 'ok', rows: [] }),
      gas: (b, respond) => respond({ status: 'ok', rows: [] })
    });
    app.s.refreshInBackground();
    await settle();                       // 同期はまだ返っていない
    app.hits.length = 0;
    await app.s.presFetchList(false);     // この隙に別タブが読みに来た想定
    expect(app.hits.some(h => h.startsWith('d1:'))).toBe(false);   // 古いD1を読まない
    expect(app.hits.some(h => h.startsWith('gas:'))).toBe(true);
    if (releaseSync) releaseSync({ ok:true, status:200, clone:()=>({json:()=>Promise.resolve({})}), json:()=>Promise.resolve({status:'ok',rows:1,skipped:false}) });
  });

  it('★再送で「もう入っていた」と分かった時もD1へ取り込ませる（画面から予定が消えない）', async () => {
    const app = makeApp({
      backendJson: D1_CFG,
      gas: (b, respond) => {
        if (b.action === 'pres_list') return respond({ status:'ok', rows:[{ ID:'PDUP1', 'タイトル':'先に届いていた予定', '開始日':'2026-08-27' }] });
        return respond({ status: 'ok' });
      },
      presSync: (b, respond) => respond({ status:'ok', rows:1, skipped:false }),
      d1: (b, respond) => respond({ status: 'ok', rows: [] })
    });
    // 「一度送信に失敗した未送信」をキューに積む
    const q = app.T.presQueue;
    q.enqueue({ id:'PDUP1', rows:[{ id:'PDUP1', title:'先に届いていた予定', startDate:'2026-08-27' }] }, Date.now());
    const item = q.list()[0];
    // ★1度送信を試みて失敗させる＝次回は wasRetry になり「もう入っていないか」の確認が走る
    const firstClaim = q.beginSend(item.id, Date.now());
    q.markFailed(item.id, firstClaim.token, 'HTTP 404', Date.now() - 60000);
    expect(q.nextDue(Date.now())).toBeTruthy();   // 再送の順番が回ってくる状態

    app.hits.length = 0;
    await app.s.presDrainQueue();
    await settle();
    // 着地確認(pres_list)のあと、D1への取り込みを必ず呼ぶ
    expect(app.hits).toContain('pres-sync');
  });
});

// ★2026-08-26 Codex再レビューで、上の修正そのものに見つかった欠陥の再発防止。
describe('GAS優先の立て方・外し方（Codex再レビュー[P1]）', () => {
  const settle = () => new Promise(r => setTimeout(r, 40));

  it('★同期が確実に成功したら必ず解除される（解除できないと社長用が15分間ずっと遅いまま）', async () => {
    const app = makeApp({
      tickingClock: true,     // beginAttempt と mark が別ミリ秒になる＝実機で普通に起きる状況
      backendJson: D1_CFG,
      presSync: (b, respond) => respond({ status:'ok', rows:2, skipped:false }),   // 確実成功
      d1: (b, respond) => respond({ status:'ok', rows: [] }),
      gas: (b, respond) => respond({ status:'ok', rows: [] })
    });
    app.s.refreshInBackground();
    await settle();
    expect(app.T.tracker.status()).toBe('trust');   // 解除されている
    app.hits.length = 0;
    await app.s.presFetchList(false);
    expect(app.hits.some(h => h.startsWith('d1:'))).toBe(true);   // 次からD1を読める＝速い
  });

  // ★このテストの期待値は Codex 再レビュー2回目で反転させた。
  //   当初は「backendがgasなら無用なブロックを残さない(trust)」を正しいとしたが、
  //   clear すると「D1がまだこの書き込みを取り込んでいない」証拠を捨てることになり、
  //   あとで d1 へ切り替えた直後に古いD1を即信用してしまう。gas運用中は読み先が
  //   そもそもGASなので、ブロックが残っていても実害はゼロ。安全側を採る。
  it('★backendがgasのときはブロックを解除しない（D1未取り込みの証拠を捨てない）', async () => {
    const app = makeApp({
      tickingClock: true,
      backendJson: { backend: 'gas' },
      gas: (b, respond) => respond({ status:'ok', rows: [] })
    });
    app.s.refreshInBackground();
    await settle();
    expect(app.T.tracker.status()).not.toBe('trust');
    // ただし読み先はGASなので、この状態でも普通に読める（実害ゼロの確認）
    app.hits.length = 0;
    const { data } = await app.s.presFetchList(false);
    expect(data.status).toBe('ok');
    expect(app.hits.some(h => h.startsWith('d1:'))).toBe(false);
  });

  it('★並行する2つの保存で、古い方の失敗が新しい方の記録を巻き戻さない', async () => {
    // A(先に開始)は同期失敗、B(後に開始)は成功。Bの成功でブロックが解けてはいけない。
    const app = makeApp({ tickingClock: true, backendJson: D1_CFG });
    const tr = app.T.tracker;
    tr.mark();                      // A: 先に立てる
    const a = tr.beginAttempt();
    tr.mark();                      // B: あとから立てる（前へ進む）
    const b = tr.beginAttempt();
    tr.mark();                      // A の失敗（今の時刻で前へ進む）
    expect(tr.clear(b)).toBe(false);        // Bの成功では解除できない
    expect(tr.status()).not.toBe('trust');
    expect(tr.clear(a)).toBe(false);        // Aの成功でも当然解除できない
  });

  it('★未送信を消す前にGAS優先を立てる（消してから立てるまでの隙にタブが閉じても安全側）', async () => {
    const app = makeApp({
      backendJson: D1_CFG,
      gas: (b, respond) => {
        if (b.action === 'pres_list') return respond({ status:'ok', rows:[{ ID:'PORD1', 'タイトル':'先に届いていた', '開始日':'2026-08-27' }] });
        return respond({ status:'ok' });
      },
      presSync: (b, respond) => respond({ status:'ok', rows:1, skipped:false }),
      d1: (b, respond) => respond({ status:'ok', rows: [] })
    });
    const q = app.T.presQueue;
    q.enqueue({ id:'PORD1', rows:[{ id:'PORD1', title:'先に届いていた', startDate:'2026-08-27' }] }, Date.now());
    const it0 = q.list()[0];
    const c0 = q.beginSend(it0.id, Date.now());
    q.markFailed(it0.id, c0.token, 'HTTP 404', Date.now() - 60000);

    // markSent が呼ばれた瞬間の tracker の状態を記録する
    let statusAtMarkSent = null;
    const origMarkSent = q.markSent.bind(q);
    q.markSent = (...args) => { statusAtMarkSent = app.T.tracker.status(); return origMarkSent(...args); };

    await app.s.presDrainQueue();
    await settle();
    expect(statusAtMarkSent).not.toBe(null);
    expect(statusAtMarkSent).not.toBe('trust');   // 消す時点で既にGAS優先が立っている
  });
});
