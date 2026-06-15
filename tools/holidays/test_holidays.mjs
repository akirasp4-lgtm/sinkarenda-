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
