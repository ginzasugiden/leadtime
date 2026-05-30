# リードタイム修正侍 — Claude Code 開発ガイド

楽天店舗運営者向けの**複数店舗対応Webツール**。楽天RMS上の商品リードタイム(LT)設定を管理する。
バックエンド = Google Apps Script(GAS)、フロント = GitHub Pages。
旧来は自社のみのスプレッドシート作業だったが、各店舗が個別アカウントでログインして使うWeb UIへ移行した。

## 現在のステータス：稼働中
ログイン → 商品検索 → LT取得（現在値をドロップダウンにプリセット） → LT変更 → RMS反映 まで一通り動作確認済み。
このガイドは「これから作る指示書」ではなく、**既存の動く実装のリファレンス**として扱うこと。

---

## 構成
```
X:\git\leadtime\
├── gas\   ← clasp管理。本番ロジックは doPost.js に集約
└── web\   ← GitHub Pages UI（index.html 1ファイル）
```
- GitHub Pages: `https://ginzasugiden.github.io/leadtime/web/index.html`
- GAS scriptId: `1Zjv7sNl5rKuPdP8aRrS4yWnmOQvuJ4CCdDkWEL_3LWSNXeM2K80p9nbQ`
- 現行 Web App URL（要確認）: `https://script.google.com/macros/s/AKfycbxsWq3cKOCzHvGpEmvB-PM0MQzR0opgiJEy9vnUt4hfOP7UvdaCqTm8iViB3f4QzNSusg/exec`

---

## アーキテクチャ
```
[web/index.html]  --JSONP-->  [gas/doPost.js]  --ESA auth-->  [楽天RMS API]
ログイン画面      id+pw         シートで認証→token
操作画面          token+action  licenseKey/serviceSecretを内部取得→楽天API呼び出し→JSON返却
```
- **licenseKey/serviceSecret は GAS内部のみで使用。フロントには絶対に返さない。**
- **CORS対策：GAS呼び出しは fetch ではなく JSONP を使う。**

---

## ファイルマップ（gas/）
- **`doPost.js`** ← 本番。以下をすべて内包する。
  - 認証: `safeBase64Decode_` / `getUserFromSheet_` / `createSession_` / `validateSession_` / `deleteSession_` / `buildEsaAuthHeader_`
  - エントリ: `doGet` / `doPost` / `handleAction_`
  - 機能: `getLeadTimeListJson_` / `fetchInventoryLT_` / `searchItemsJson_` / `searchItemsWithLTJson_` / `updateLeadTimeJson_`
  - レスポンス: `createJsonResponse_`
- 旧シート版の名残ファイル: `商品情報を取得.js` / `商品検索.js` / `商品名検索.js` / `LT一覧取得.js` / `LT更新.js` / `出荷LT取得.js` / `定期実行.js` / `■在庫あり納期管理更新.js` ほか
  → **現行Webフローでは未使用**。新規実装時に参照・流用しないこと。整理候補。

## action 一覧（handleAction_）
- 認証不要: `login` / `logout`
- 認証必要（token必須）: `getLeadTimeList` / `searchItems` / `searchItemsWithLT` / `updateLeadTime`

---

## 認証
- **`auth.js` は存在しない。認証はすべて `doPost.js` に直接実装**（過去に `gas-auth` ライブラリ依存があったが廃止済み・redundant）。
  - 理由: `CacheService` が GAS のライブラリ境界をまたいで正しく動かないため。
- ユーザー/APIキーシート: `1iYeV2SbOVoRH8Qjm2d1w5tWmhlE_zcc-yO1tDSLN7Rk` の `api_key` タブ
  - 列: A=id / B=CHATGPT_API_KEY / C=licenseKey / D=serviceSecret / E=download / F=pw / G=sid / H=sname / I=email / J=flag(0=有効) / K=expiry
  - **B/C/D/F は `BASE64:` プレフィックス付きで格納**。`safeBase64Decode_` でプレフィックスを剥がしてからデコードする。デコードは認証/セッション層で**1回だけ**行い、下流関数には平文を渡す。
  - pw は BASE64 デコードして照合。`flag=0` かつ `expiry` が未来のユーザーのみ有効。
- 認証ヘッダー: `buildEsaAuthHeader_(session)` に一本化。形式は `ESA Base64(serviceSecret:licenseKey)`。
- セッション: `CacheService`。フロントのトークンは `sessionStorage`。401/認証エラー時はログイン画面へ戻す。

---

## 楽天RMS API
- **ItemAPI 2.0**: `/es/2.0/items/search`（検索） / `/es/2.0/items/inventory-related-settings/`
  - 検索は **`manageNumber` パラメータ**を使う（`title` ではない）。レスポンスは `results[].item` をパース。
  - 注意: `manageNumber` は完全一致〜前方一致で、中間一致（contains）はできない。中間一致が必要ならページング取得して GAS 側で `String.includes()` フィルタする。
- **Shop API**: `/es/1.0/shop/operationLeadTime` / `/es/1.0/shop/delvdateMaster`
- **Inventories API（ES 2.1）**: `/es/2.1/inventories/manage-numbers/{manageNumber}/variants/{variantId}`
  → **LTの取得・更新はこのAPIを使う**。`normalDeliveryTimeId` を `operationLeadTimeId` 体系で返すため、LTマスターと直接マッチする（旧シート版と同じ考え方）。
- LT更新（PUT）は全体上書きなので、**更新前に現在設定をGETし、変更フィールドだけ差し替えてからPUT**する（在庫切れ時納期などを消さないため）。

### 【最重要】ID体系の落とし穴
- `inventory-related-settings.get` が返す `normalDeliveryDateId` は **delvdateNumber 体系**（1, 5, 27 のような小さい整数）。
- `operationLeadTime` が使う `operationLeadTimeId` は**大きい整数**（12392, 12920 等）。
- **両者は非互換**。LTデータの正は **Inventories API（operationLeadTimeId 体系）**。混同しないこと。

---

## レート制限（QPS）
- 楽天API呼び出しに sleep を挟む。現状は **1100ms + 429時リトライ**（過去の 1500ms から変更済み）。

---

## 【最重要】デプロイ手順 — 絶対厳守
- **`clasp deploy` は絶対に使わない**。デプロイ種別が Web App → Library に変わってしまうため。
- 反映は **`clasp push --force` のみ**。
- そのうえで **GASエディタで手動バージョン更新**: デプロイ → デプロイを管理 → 編集(鉛筆) → 新バージョン → デプロイ。
- `web/index.html` の `GAS_URL` 更新も手動。
- `appsscript.json` の webapp は `executeAs: USER_DEPLOYING`, `access: ANYONE_ANONYMOUS`。

---

## 開発フロー（Claude Code）
- 一時的な作業指示は `.md` でディレクトリに置き、**使用後は削除**。恒久的に残すのは `CLAUDE.md` のみ。
- 変更前に `git status` を確認し、意味のある単位でコミットする。
- デバッグ時は各関数に `Logger.log` を入れて GAS 実行ログで追う。
- 大きめの実装は subagent に分割してよい（`.claude/agents/` 参照）。

## やらないこと（禁止事項）
- `clasp deploy`
- `gas-auth` ライブラリの再導入
- licenseKey / serviceSecret をフロントへ返す
- 旧シート版ファイルを現行Webフローに組み込む
- `manageNumber` で中間一致を期待する実装
