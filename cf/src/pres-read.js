// 社長予定の読み取り。GASの pres_list と "同じ形" を返す。
// 画面側（president.html）の分岐を増やさないため、キー名も形も1つも変えない。
//   pres_list の応答: { status:'ok', rows:[ {登録日時,タイトル,開始日,...}, ... ] }
//
// ★社員用（cf/src/read.js）との違い
//   - 参照するテーブルは pres_snapshot / pres_sync_log のみ。
//     社員用の snapshot / sync_log には絶対に触れない（cf/schema-president.sql の
//     コメント参照。混ぜると社員用の鮮度判定とレート制限が壊れる）。
//   - 会社での絞り込みが無い（社長予定は会社をまたがない単一のシート）。
//     保存したものをそのまま返すだけ。

// 直近の同期成功からこの時間より古ければ「もう正常データとは言えない」とみなして
// エラーを返す。社員用（read.js）と同じ15分＝Cron3回分の猶予。
// 一時的な取得失敗が1〜2回続く程度では無用にフォールバックさせない。
export const PRES_FRESHNESS_THRESHOLD_MS = 15 * 60 * 1000;

// ★変更なしスキップ（pres-sync.js）も ok=1 で記録する。「変更が無いことを確認できた」
// のも成功だからで、これが無いと予定を触らない日が続くだけで鮮度ガードが誤発火する。
async function getLastSuccessAt(env) {
  const res = await env.DB.prepare(
    'SELECT at FROM pres_sync_log WHERE ok = 1 ORDER BY at DESC LIMIT 1'
  ).all();
  const row = (res.results && res.results[0]) || null;
  return row ? row.at : null;
}

/**
 * 社長予定を返す。失敗しても例外は投げず {status:'error', message} を返す契約。
 * 画面側は status!=='ok' を見て自動的にGASへ切り替わるため、利用者にはエラーが
 * 見えない（遅くなるだけで済む）。
 */
export async function readPresident(env) {
  const res = await env.DB.prepare('SELECT payload FROM pres_snapshot WHERE id = 1').all();
  const row = (res.results && res.results[0]) || null;
  if (!row) {
    // まだ一度も取り込みが成功していない。空のD1を「予定ゼロ件」として返すと
    // 社長のカレンダーが空に見えてしまうため、必ずエラーで返す。
    return { status: 'error', message: 'まだ取り込みが行われていません' };
  }

  // 鮮度ガード。snapshotが「存在するだけ」で正常返却すると、同期が何日失敗し続けても
  // 最後に成功した古い内容を無条件に正常として返し続けてしまう（社員用で実測再現された欠陥）。
  const lastSuccessAt = await getLastSuccessAt(env);
  if (!lastSuccessAt) {
    return { status: 'error', message: '同期の成功記録がありません' };
  }
  const lastSuccessMs = Date.parse(lastSuccessAt);
  if (!Number.isFinite(lastSuccessMs) || Date.now() - lastSuccessMs > PRES_FRESHNESS_THRESHOLD_MS) {
    return {
      status: 'error',
      message: `同期が長時間成功していません（最終成功: ${lastSuccessAt}）。最新のデータではない可能性があるため取得を中止しました`
    };
  }

  let rows;
  try {
    rows = JSON.parse(row.payload);
  } catch (e) {
    // 理論上は起こらない（書き込み前にJSON.stringifyしたものしか入らない）が、
    // 万一壊れていてもクラッシュさせず、GASへフォールバックさせる。
    return { status: 'error', message: '保存済みデータの形式が壊れています' };
  }
  if (!Array.isArray(rows)) {
    return { status: 'error', message: '保存済みデータの形式が想定と違います' };
  }

  return { status: 'ok', rows };
}
