// GASの応答とWorkerの応答を突き合わせる関門。
// 完全一致しない限り移行を進めてはいけない。
//
// ★2026-08-24 改訂: 以前は「ID＋作業日＋氏名」をキーにしたMapで対応づけていたが、
//   実データには氏名が空の行（車検期限リマインダー）も、同一内容の重複行（職人マスタ）も
//   実在することが判明した。キー方式ではそれらを取りこぼす／誤って対にするため、
//   **並び順のまま1行ずつ突き合わせる方式**に変えた。D1側も並び順を保持している。
//
// 使い方: node cf/test/compare.mjs <GAS_URL> <WORKER_URL>
const [gasUrl, workerUrl, companyArg] = process.argv.slice(2);
if (!gasUrl || !workerUrl) {
  console.error('使い方: node cf/test/compare.mjs <GAS_URL> <WORKER_URL> [会社名]');
  console.error('  会社名を省略すると全社。実際のアプリは会社を選んだ状態で叩くため、');
  console.error('  切り替え前には主要な会社名でも1回ずつ実行すること（レビュー指摘）。');
  process.exit(2);
}
const company = companyArg || '';
console.log('=== 対象: ' + (company || '全社') + ' ===');

const g = await (await fetch(gasUrl + '?compact=1&company=' + encodeURIComponent(company) + '&t=' + Date.now())).json();
const w = await (await fetch(workerUrl + '/api/schedule?company=' + encodeURIComponent(company))).json();

if (g.status !== 'ok') { console.error('GAS側がエラー:', g.message); process.exit(1); }
if (w.status !== 'ok') { console.error('Worker側がエラー:', w.message); process.exit(1); }

let ok = true;
const ng = (msg) => { ok = false; console.log('  NG:', msg); };

// ── ヘッダ ──
console.log('■ ヘッダ');
if (JSON.stringify(g.headers) !== JSON.stringify(w.headers)) {
  ng(`19列の並びが違う\n    GAS: ${JSON.stringify(g.headers)}\n    CF : ${JSON.stringify(w.headers)}`);
} else console.log('  OK: 19列が順番まで一致');

// ── 日報の行（順序どおり1行ずつ）──
console.log('■ 日報データ');
console.log(`  GAS ${g.rows.length}行 / CF ${w.rows.length}行`);
if (g.rows.length !== w.rows.length) ng(`行数が違う（差 ${g.rows.length - w.rows.length}）`);

const n = Math.min(g.rows.length, w.rows.length);
let rowDiff = 0;
for (let i = 0; i < n; i++) {
  const a = g.headers.map((_, j) => String(g.rows[i][j] ?? ''));
  const b = w.headers.map((_, j) => String(w.rows[i][j] ?? ''));
  if (JSON.stringify(a) === JSON.stringify(b)) continue;
  rowDiff++;
  if (rowDiff <= 5) {
    const cols = g.headers.map((h, j) => (a[j] !== b[j] ? `${h}: GAS「${a[j]}」/ CF「${b[j]}」` : null)).filter(Boolean);
    console.log(`  NG: ${i + 1}行目 → ${cols.join(' , ')}`);
  }
}
if (rowDiff) { ok = false; console.log(`  NG: 中身が違う行 ${rowDiff}件`); }
else if (g.rows.length === w.rows.length) console.log('  OK: 全行が順序も中身も一致');

// GAS側にしか無い行を分かりやすく出す（行数が違うときの手掛かり）
if (g.rows.length > w.rows.length) {
  const iId = g.headers.indexOf('ID'), iD = g.headers.indexOf('作業日'), iN = g.headers.indexOf('氏名');
  console.log('  CF側に無い可能性のある行（先頭5件）:');
  for (const r of g.rows.slice(w.rows.length, w.rows.length + 5)) {
    console.log(`    ID=${r[iId]} 作業日=${r[iD]} 氏名=${r[iN] || '(空)'}`);
  }
}

// ── マスタ3種（順序どおり1件ずつ）──
const cmpList = (label, ga, wa, keys) => {
  console.log(`■ ${label}`);
  console.log(`  GAS ${ga.length}件 / CF ${wa.length}件`);
  if (ga.length !== wa.length) ng(`件数が違う（差 ${ga.length - wa.length}）`);
  const m = Math.min(ga.length, wa.length);
  let d = 0;
  for (let i = 0; i < m; i++) {
    for (const k of keys) {
      if (String(ga[i][k] ?? '') !== String(wa[i][k] ?? '')) {
        d++;
        if (d <= 3) console.log(`  NG: ${i + 1}件目 ${k}: GAS「${ga[i][k]}」/ CF「${wa[i][k]}」`);
        break;
      }
    }
  }
  if (d) { ok = false; console.log(`  NG: 中身が違う ${d}件`); }
  else if (ga.length === wa.length) console.log('  OK: 全件が順序も中身も一致');
};

cmpList('職人マスタ', g.members, w.members, ['name', 'company', 'division']);
cmpList('元請マスタ', g.genbaMaster, w.genbaMaster, ['name', 'company']);
cmpList('現場マスタ', g.jobsites, w.jobsites, ['genba', 'loc', 'jobNo', 'completed', 'billingMethod']);

// ── 単価がCF側へ漏れていないこと（給料情報をD1に置かない方針）──
console.log('■ 単価(rate)の混入チェック');
if (w.members.every(m => !('rate' in m))) console.log('  OK: CF側に単価は含まれていない');
else ng('CF側の職人マスタに単価が混入している');

console.log('');
console.log(ok ? '=> 完全一致。移行を進めてよい' : '=> 不一致あり。移行を進めないこと');
process.exit(ok ? 0 : 1);
