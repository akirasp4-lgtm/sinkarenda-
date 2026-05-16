# 予定管理アプリ

建設現場向けの日報管理・粗利管理 Web アプリ。

- **フロントエンド**: 静的 HTML + JS（GitHub Pages 配信）
- **バックエンド**: Google Apps Script（GAS）
- **データ保管**: Google スプレッドシート（バインドスクリプト）

---

## 📁 リポジトリ構成

```
sinkarenda-/
├── README.md              ← このファイル
├── 引き継ぎ.md            ← プロジェクト全体の引継書（最重要）
├── index.html             一般ユーザー用カレンダー
├── admin.html             管理者用画面（PIN: 8800）
├── president.html         社長専用カレンダー（PIN: 1203）
├── links.html             配布URL一覧
├── manifest.json          PWA設定
├── icon.png               アプリアイコン
├── gas.js                 GAS バックエンド（手動で Apps Script に貼る）
├── docs/
│   └── 使い方ガイド.md
└── ラインボットからの依頼.md / ラインボットへの返信.md
                           ← 車両予約 GAS 連携の経緯
```

---

## 🚀 別 PC でセットアップする手順

### 1. リポジトリを clone

```bash
git clone https://github.com/akirasp4-lgtm/sinkarenda-.git
cd sinkarenda-
```

### 2. Git ユーザー設定（このリポジトリ内だけ）

```bash
git config user.name "akirasp4-lgtm"
git config user.email "akirasp4@gmail.com"
```

### 3. アクセス権限の確認

別 PC でも以下が必要：

| 何を | どこで |
|---|---|
| GitHub の push 権限 | `akirasp4-lgtm` アカウントでログイン（または PAT） |
| スプレッドシート編集権限 | 所有者 Google アカウントでアクセス |
| GAS 編集権限 | 同上（バインドスクリプト） |

**Groq API キー** は GAS のスクリプトプロパティ `GROQ_API_KEY` にすでに登録済みなので、開発側に持つ必要なし。

### 4. 編集 → push → デプロイ

#### HTML の変更（GitHub Pages）

```bash
# 編集
vi index.html  # またはお好みのエディタ

# commit & push
git add index.html
git commit -m "..."
git push origin main
```

→ GitHub Pages の自動デプロイで 1-2 分後に本番反映。

#### GAS の変更

1. `gas.js` を編集 & commit & push（GitHub 上のコードは記録目的）
2. **別途、手動で Apps Script エディタに貼り直す**:
   - https://docs.google.com/spreadsheets/d/【スプシID】/edit を開く
   - 「拡張機能」→「Apps Script」
   - エディタ内で `Ctrl + A` → `Ctrl + V`（gas.js の中身を貼る）
   - `Ctrl + S` で保存
   - 右上「デプロイ」→「デプロイを管理」→ 鉛筆 → バージョン「**新しいバージョン**」→ 「デプロイ」
3. WebApp URL は変わらない（同じデプロイの新バージョン扱い）

---

## 📋 配布 URL

| 対象 | URL |
|---|---|
| グローライズ | `https://akirasp4-lgtm.github.io/sinkarenda-/index.html?c=グローライズ` |
| 和信カインド | `https://akirasp4-lgtm.github.io/sinkarenda-/index.html?c=和信カインド` |
| ラーテル | `https://akirasp4-lgtm.github.io/sinkarenda-/index.html?c=ラーテル` |
| GRHD | `https://akirasp4-lgtm.github.io/sinkarenda-/index.html?c=GRHD` |
| GRミツマ | `https://akirasp4-lgtm.github.io/sinkarenda-/index.html?c=GRミツマ` |
| 管理画面 | `https://akirasp4-lgtm.github.io/sinkarenda-/admin.html`（PIN: 8800） |
| 社長画面 | `https://akirasp4-lgtm.github.io/sinkarenda-/president.html`（PIN: 1203） |
| リンク集 | `https://akirasp4-lgtm.github.io/sinkarenda-/links.html` |

---

## 📚 詳しい情報

- **引き継ぎ.md** — 全体仕様・直近の変更履歴・既知の課題
- **docs/使い方ガイド.md** — 一般ユーザー向けマニュアル

---

## ⚠️ 注意

- **このリポジトリは Public** です。本物の機密情報（API キーなど）は git に含めないこと
- 確認用の PIN（8800 / 1203）は HTML 内にハードコードされており、すでに事実上公開状態。**機密性が必要な場合は別途認証の仕組みを足す**こと
- スプレッドシートと GAS は Google アカウント認可で保護されている。これが事実上のセキュリティ境界
