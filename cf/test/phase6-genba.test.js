// 元請プルダウンの検索と、表記ゆれ拾い（2026-08-31 利用者要望）
//
// 利用者の言葉:
//   「直接入力で入れた元請けを検索できるようにしてほしい」
//   「グローライズ自社って元請けがあるのにグローライズで入れられてて
//     検索も出来ないし経費精算アプリにもでてこないんだよ」
//   「お前に直せって言えばできるのはわかってるけどそれじゃ事務員が直せない」
//
// ★AIがデータを直すのではなく、事務員が画面から直せる形にした。
//   その判定部分（誰と誰が同じ会社か）をここで固定する。
import { describe, it, expect, beforeAll } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import vm from 'node:vm';

const here = dirname(fileURLToPath(import.meta.url));
const read = (f) => readFileSync(join(here, '..', '..', f), 'utf8');

const BEGIN = '// ===== PHASE6-GENBA-RULE:BEGIN =====';
const END = '// ===== PHASE6-GENBA-RULE:END =====';
function extract(file) {
  const src = read(file);
  const i = src.indexOf(BEGIN), j = src.indexOf(END);
  if (i < 0 || j < 0) throw new Error(file + ' に元請のルールブロックが無い');
  return src.slice(i + BEGIN.length, j);
}

const EXPORT = ';globalThis.__g6 = { genbaNorm, genbaSelectChoices, genbaCanon, '
  + 'genbaUsageList, genbaVariantGroups, genbaUnregistered };';

let G;
beforeAll(() => {
  const sandbox = vm.createContext({ console, String, Object, Array, Number, Boolean });
  sandbox.globalThis = sandbox;
  vm.runInContext(extract('index.html') + EXPORT, sandbox, { filename: 'index.html' });
  G = sandbox.__g6;
});

it('index.html と admin.html が完全に同じ', () => {
  expect(extract('admin.html')).toBe(extract('index.html'));
});

// 予定1行を作る近道
const n = (genba, o) => Object.assign({
  genba, date: '2026-08-01', company: 'グローライズ', isGhost: false
}, o || {});

describe('全角と半角を同じ物として扱う', () => {
  it('ＨＳＪ と HSJ は同じ芯になる', () => {
    expect(G.genbaNorm('ＨＳＪ')).toBe('hsj');
    expect(G.genbaNorm('HSJ')).toBe('hsj');
  });

  it('★ひらがなでも引ける（日本語入力の変換前で0件にしない）', () => {
    // 本番の画面で確かめたら「ぐろ」で0件だった。
    // 日本語入力は変換前がひらがななので、打っている最中ずっと
    // 「見つかりません」に見えてしまう。
    expect(G.genbaNorm('ぐろーらいず')).toBe(G.genbaNorm('グローライズ'));
    expect(G.genbaNorm('きんでん')).toBe(G.genbaNorm('キンデン'));
  });

  it('カタカナ以外は変えない（漢字・英字はそのまま）', () => {
    expect(G.genbaNorm('中村電設')).toBe('中村電設');
    expect(G.genbaNorm('JTE')).toBe('jte');
  });

  it('全角スペースも普通の空白として扱う', () => {
    expect(G.genbaNorm('公共工事　入札')).toBe('公共工事 入札');
  });

  it('前後の空白は落とす／空でも落ちない', () => {
    // ★ひらがなはカタカナに寄せてから比べる（上の「ひらがなでも引ける」参照）ので、
    //   ここでの期待値もカタカナになる。比較用の芯なので画面の表示は変わらない。
    expect(G.genbaNorm('  きんでん東  ')).toBe('キンデン東');
    expect(G.genbaNorm(null)).toBe('');
    expect(G.genbaNorm(undefined)).toBe('');
  });

  it('株式会社などの飾りを外すと同じ会社だと分かる', () => {
    expect(G.genbaCanon('株式会社エイシン')).toBe(G.genbaCanon('エイシン'));
    expect(G.genbaCanon('株式会社HSJ')).toBe(G.genbaCanon('ＨＳＪ'));
    expect(G.genbaCanon('ハイテックス（北九州）')).toBe('ハイテックス北九州');
  });
});

describe('プルダウンの中身（検索）', () => {
  const master = ['きんでん東', 'きんでん西', 'グローライズ自社'];
  const nippos = [
    n('株式会社HSJ', { date: '2026-09-04' }),
    n('ＨＳＪ', { date: '2026-08-26' }),
    n('ＨＳＪ', { date: '2026-08-20' }),
    n('きんでん東'),
    n('他社の元請', { company: 'GRミツマ' })
  ];

  it('★直接入力した元請もプルダウンに出る（これが今回の要望）', () => {
    const c = G.genbaSelectChoices(master, nippos, 'グローライズ', '', '');
    expect(c.used.map((u) => u.name)).toEqual(['株式会社HSJ', 'ＨＳＪ']);
  });

  it('マスタにある元請は「直接入力ぶん」に二重で出さない', () => {
    const c = G.genbaSelectChoices(master, nippos, 'グローライズ', '', '');
    expect(c.used.some((u) => u.name === 'きんでん東')).toBe(false);
    expect(c.reg).toEqual(['きんでん東', 'きんでん西', 'グローライズ自社'].sort());
  });

  it('直接入力ぶんは最近使った順に並ぶ', () => {
    const c = G.genbaSelectChoices(master, nippos, 'グローライズ', '', '');
    expect(c.used[0].name).toBe('株式会社HSJ');   // 2026-09-04
    expect(c.used[0].count).toBe(1);
    expect(c.used[1].count).toBe(2);              // ＨＳＪ は2件
  });

  it('★他の会社の元請は混ぜない', () => {
    const c = G.genbaSelectChoices(master, nippos, 'グローライズ', '', '');
    expect(c.used.some((u) => u.name === '他社の元請')).toBe(false);
  });

  it('全社なら会社で絞らない', () => {
    const c = G.genbaSelectChoices(master, nippos, '全社', '', '');
    expect(c.used.some((u) => u.name === '他社の元請')).toBe(true);
  });

  it('★半角で打っても全角の元請が見つかる', () => {
    const c = G.genbaSelectChoices(master, nippos, 'グローライズ', 'hsj', '');
    expect(c.used.map((u) => u.name)).toEqual(['株式会社HSJ', 'ＨＳＪ']);   // 最近使った順
    expect(c.hit).toBe(2);
    expect(c.total).toBe(5);   // マスタ3 + 直接入力2
  });

  it('絞り込みは登録済みの元請にも効く', () => {
    const c = G.genbaSelectChoices(master, nippos, 'グローライズ', 'きんでん', '');
    expect(c.reg).toEqual(['きんでん東', 'きんでん西']);
  });

  it('★今選んでいる元請は、絞り込みで外れても必ず残す（選択を勝手に捨てない）', () => {
    const c = G.genbaSelectChoices(master, nippos, 'グローライズ', 'きんでん', 'グローライズ自社');
    expect(c.reg).toContain('グローライズ自社');
  });

  it('予定がゼロでも落ちない', () => {
    const c = G.genbaSelectChoices(master, [], 'グローライズ', '', '');
    expect(c.used).toEqual([]);
    expect(c.total).toBe(3);
  });

  it('取り消し線の行（ゴースト）は数えない', () => {
    const c = G.genbaSelectChoices([], [n('幽霊', { isGhost: true })], '全社', '', '');
    expect(c.used).toEqual([]);
  });
});

describe('表記ゆれを拾う（事務員が直すための候補）', () => {
  it('★グローライズ と グローライズ自社 が同じ組になる（利用者が困っていた例）', () => {
    const list = G.genbaUsageList(['グローライズ自社'], [
      n('グローライズ自社'), n('グローライズ自社'), n('グローライズ')
    ], 'グローライズ');
    const g = G.genbaVariantGroups(list);
    expect(g.length).toBe(1);
    expect(g[0].map((x) => x.name).sort()).toEqual(['グローライズ', 'グローライズ自社']);
  });

  it('★残す候補（マスタ登録済み）が先頭に来る', () => {
    const list = G.genbaUsageList(['グローライズ自社'], [
      n('グローライズ'), n('グローライズ'), n('グローライズ'), n('グローライズ自社')
    ], 'グローライズ');
    const g = G.genbaVariantGroups(list);
    expect(g[0][0].name).toBe('グローライズ自社');   // 件数は少なくてもマスタ登録済みが先
    expect(g[0][0].inMaster).toBe(true);
  });

  it('★HSJ の表記ゆれ3つは1組、HSJ-KNSI は代表との2つ組で出す', () => {
    // ★Codexレビュー（2026-08-31）で作り直した。
    //   以前は「片方が丸ごと含まれる」だけで数珠つなぎにまとめていたので、
    //   「大阪電設」「大阪電設工業」「電設工業」まで1組になった（次のテスト）。
    //   芯が完全に同じ物だけを1組にし、含む関係は2つずつの組にする。
    const list = G.genbaUsageList([], [
      n('HSJ'), n('ＨＳＪ'), n('株式会社HSJ'), n('HSJ-KNSI')
    ], '全社');
    const g = G.genbaVariantGroups(list);
    const exact = g.find((x) => x.length === 3);
    expect(exact, '芯が同じ3つが1組になっていない').toBeTruthy();
    expect(exact.map((x) => x.name).sort()).toEqual(['HSJ', 'ＨＳＪ', '株式会社HSJ'].sort());
    // HSJ-KNSI も候補から落とさない（代表との2つ組で出す）
    const pair = g.find((x) => x.some((y) => y.name === 'HSJ-KNSI'));
    expect(pair, 'HSJ-KNSI が候補から落ちている').toBeTruthy();
    expect(pair.length).toBe(2);
  });

  it('★無関係な元請が橋渡しで数珠つなぎにならない（Codexが再現した壊れ方）', () => {
    // 「大阪電設」と「電設工業」は互いに似ていないのに、
    // 真ん中の「大阪電設工業」が橋渡しして1組になっていた。
    // 取り消せない統一の候補にこれを並べるのは危ない。
    const list = G.genbaUsageList([], [
      n('大阪電設'), n('大阪電設工業'), n('電設工業')
    ], '全社');
    const g = G.genbaVariantGroups(list);
    g.forEach((grp) => {
      const names = grp.map((x) => x.name);
      expect(names.includes('大阪電設') && names.includes('電設工業'),
        '似ていない2つが同じ組に入っている').toBe(false);
      expect(grp.length).toBe(2);
    });
  });

  it('★きんでん東 と きんでん西 は別物として扱う（違う元請を勝手に束ねない）', () => {
    const list = G.genbaUsageList(['きんでん東', 'きんでん西'], [n('きんでん東'), n('きんでん西')], '全社');
    expect(G.genbaVariantGroups(list)).toEqual([]);
  });

  it('★2文字以下の短い名前は「含む」だけで束ねない（誤爆よけ）', () => {
    const list = G.genbaUsageList([], [n('BF'), n('BFコーポレーション'), n('サンBF電設')], '全社');
    expect(G.genbaVariantGroups(list)).toEqual([]);
  });

  it('似ている物が無ければ空（1件だけの組は返さない）', () => {
    const list = G.genbaUsageList(['きんでん東'], [n('きんでん東')], '全社');
    expect(G.genbaVariantGroups(list)).toEqual([]);
  });

  it('影響の大きい組が上に来る', () => {
    const list = G.genbaUsageList([], [
      n('あああ会社'), n('あああ会社支店'),
      n('いいい'), n('いいい'), n('いいい'), n('いいい商会')
    ], '全社');
    const g = G.genbaVariantGroups(list);
    expect(g[0].map((x) => x.name)).toContain('いいい');   // 4件 > 2件
  });

  it('件数と最終使用日を数える（どちらを残すか決める材料）', () => {
    const list = G.genbaUsageList(['きんでん東'], [
      n('きんでん東', { date: '2026-08-01' }),
      n('きんでん東', { date: '2026-09-15' })
    ], 'グローライズ');
    const e = list.find((x) => x.name === 'きんでん東');
    expect(e.count).toBe(2);
    expect(e.last).toBe('2026-09-15');
    expect(e.inMaster).toBe(true);
  });

  it('マスタにあるだけで一度も使われていない元請は0件として出る', () => {
    const list = G.genbaUsageList(['使ってない元請'], [], 'グローライズ');
    expect(list[0]).toMatchObject({ name: '使ってない元請', count: 0, inMaster: true });
  });
});

describe('マスタに載っていない元請の一覧', () => {
  it('★直接入力しっぱなしの元請だけを出す', () => {
    const list = G.genbaUsageList(['きんでん東'], [n('きんでん東'), n('株式会社鈴開興産')], 'グローライズ');
    expect(G.genbaUnregistered(list).map((x) => x.name)).toEqual(['株式会社鈴開興産']);
  });

  it('一度も使われていないマスタ行は出さない（登録を勧める意味がない）', () => {
    const list = G.genbaUsageList(['使ってない元請'], [], 'グローライズ');
    expect(G.genbaUnregistered(list)).toEqual([]);
  });

  it('件数の多い順に並ぶ', () => {
    const list = G.genbaUsageList([], [n('A社'), n('B社'), n('B社')], '全社');
    expect(G.genbaUnregistered(list).map((x) => x.name)).toEqual(['B社', 'A社']);
  });
});

describe('画面の配線', () => {
  ['index.html', 'admin.html'].forEach((f) => {
    const src = read(f);

    it(f + ': 入力・編集の両方に元請の絞り込み欄がある', () => {
      expect(src).toContain('id="s-genba-search"');
      expect(src).toContain('id="e-genba-search"');
      expect(src).toContain("onGenbaSearch('s')");
      expect(src).toContain("onGenbaSearch('e')");
    });

    it(f + ': ★名前をHTMLへ直接埋め込まず、番号で引く（引用符事故よけ）', () => {
      expect(src).toContain('data-gfixmerge=');
      expect(src).toContain('data-gfixadd=');
      expect(src).not.toContain('onclick="gfixMerge(\'');
      expect(src).not.toContain('onclick="gfixAdd(\'');
    });

    it(f + ': 統一する前に必ず確認を出す（取り消せない操作）', () => {
      const i = src.indexOf('async function gfixMerge');
      const body = src.slice(i, src.indexOf('\n}', i));
      expect(body).toContain('confirm(');
      expect(body).toContain('取り消せません');
    });

    it(f + ': ★通信中の二重押しを止める（同じマージが2回走らない）', () => {
      // マージは取り消せない。回線が遅いときに事務員がもう一度押すのは自然な動作なので、
      // 押せてしまうと同じ書き換えが2回走る。フラグで必ず止める。
      expect(src).toContain('let gfixBusy = false;');
      ['async function gfixMerge', 'async function gfixAdd'].forEach((fn) => {
        const i = src.indexOf(fn);
        const body = src.slice(i, src.indexOf('\n}', i));
        expect(body, fn + ' に二重押し止めが無い').toContain('if(gfixBusy)');
        expect(body, fn + ' がフラグを立てていない').toContain('gfixBusy=true;');
        expect(body, fn + ' がフラグを戻していない').toContain('finally{gfixBusy=false;');
      });
    });

    it(f + ': ★会社をまたいで書き換えない（Codexレビュー P1）', () => {
      // 実データ: 「不動産」はGRHDに21件・グローライズに1件、
      //           「オリエンス」は和信カインドに114件・グローライズに2件。
      //   グローライズの画面で「1件」と確認して実行したのに、他社の行まで
      //   書き換わっていた。「和信カインド・ラーテル・GRHDは触らない」に反する。
      const i = src.indexOf('async function gfixMerge');
      const body = src.slice(i, src.indexOf('\n}', i));
      // 会社を限定する新しい入口だけを呼ぶ（古い merge_genba は呼ばない）
      expect(body).toContain("action:'merge_genba_company'");
      expect(body).not.toContain("action:'merge_genba'");
      expect(body).toContain('company:targetCompany');
      // 会社が決まっていない「全社」では実行させない
      expect(body).toContain("targetCompany==='全社'");
    });

    it(f + ': ★押した時点の会社を握って離さない（通信中の会社切替で別会社を触らない）', () => {
      ['async function gfixMerge', 'async function gfixAdd'].forEach((fn) => {
        const i = src.indexOf(fn);
        const body = src.slice(i, src.indexOf('\n}', i));
        expect(body, fn + ' が会社を控えていない').toContain('const targetCompany=currentCompany;');
        // await より後ろで currentCompany を読み直していないこと
        const after = body.slice(body.indexOf('await fetch'));
        expect(after, fn + ' が通信後に currentCompany を読み直している')
          .not.toContain('currentCompany');
      });
    });

    it(f + ': ★古い画面（キャッシュ表示中）からは実行させない', () => {
      // 通信できずキャッシュを出しているだけのとき、他の人が既に直した後かもしれない。
      // 逆向きに統一して、直した内容を巻き戻す事故が起きる。
      ['async function gfixMerge', 'async function gfixAdd'].forEach((fn) => {
        const i = src.indexOf(fn);
        const body = src.slice(i, src.indexOf('\n}', i));
        expect(body, fn + ' が dataLoadOk を見ていない').toContain('if(!dataLoadOk)');
      });
    });

    it(f + ': ★ボタンの向きが読み違えられないようにする（取り消せない操作）', () => {
      // 以前は「→ これに統一」と書いてあったが、実際に消えるのは押した行の方。
      // 非エンジニアが逆に読むと、残したい名前を消してしまう。
      expect(src).toContain('この名前をやめる');
      expect(src).not.toContain('これに統一</button>');
      // 「残す名前」に選ばれている行は押せなくする
      expect(src).toContain('function gfixMarkKeep(gi){');
      expect(src).toContain('b.disabled=isKeep;');
    });

    it(f + ': ★統一先が未登録なら、その場で元請マスタへ登録できる', () => {
      // GAS の mergeGenba_ は「消す方」がマスタに載っているときだけ改名する。
      // HSJ / ＨＳＪ / 株式会社HSJ のようにどれも未登録だと、統一しても
      // 未登録のままでプルダウンに出てこない＝困りごとが解決しない。
      const i = src.indexOf('async function gfixMerge');
      const body = src.slice(i, src.indexOf('\n}', i));
      expect(body).toContain('!keep.inMaster');
      expect(body).toContain("action:'add_genba'");
    });

    it(f + ': ★属性値には escAttr を使う（esc は引用符を逃がさない）', () => {
      const i = src.indexOf('function populateGenbaSelect');
      const body = src.slice(i, src.indexOf('\n}', i));
      // option の value に生の esc() を使っていないこと
      expect(body).not.toMatch(/value="\$\{esc\(/);
      expect(body).toMatch(/value="\$\{escAttr\(/);
    });

    it(f + ': 事務タブに表記ゆれカードが置いてある', () => {
      expect(src).toContain('id="gfix-body"');
      expect(src).toContain('renderGenbaFix()');
    });
  });
});
