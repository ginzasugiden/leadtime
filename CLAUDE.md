# リードタイム修正侍 - 開発指示

## プロジェクト構成
```
X:\git\leadtime\
├── gas\         ← 既存GASファイルあり + 新規ファイル追加
└── web\         ← GitHub Pages UI
```

---

## アーキテクチャ

```
[web/index.html]          [gas/doPost.js]              [楽天API]
ログイン画面     →  id+pw送信  →  スプレッドシートで認証
                ←  sessionToken  ←
操作画面        →  token+action  →  licenseKey/serviceSecretを内部取得
                ←  結果JSON      ←  楽天APIを呼び出し → 結果を返す
```

**重要: licenseKey/serviceSecretはGAS内部のみで使用。フロントに返さない。**

---

## ユーザー管理スプレッドシート

- スプレッドシートID: `1iYeV2SbOVoRH8Qjm2d1w5tWmhlE_zcc-yO1tDSLN7Rk`
- シート名: `api_key`
- 列構成:
  - A列: id（ログインID）
  - B列: CHATGPT_API_KEY
  - C列: licenseKey（楽天API）
  - D列: serviceSecret（楽天API）
  - E列: download
  - F列: pw（パスワード ※BASE64エンコード済み）
  - G列: sid（楽天店舗ID）
  - H列: sname（店舗名）
  - I列: email
  - J列: flag（0=有効）
  - K列: expiry（有効期限）

---

## タスク1: gas/auth.js を新規作成

### セッション管理
- `createSession(userId)`: CacheServiceでセッショントークンを生成・保存（有効期限2時間）
- `validateSession(token)`: トークンからuserIdを取得、licenseKey/serviceSecretを返す
- `getUserFromSheet(userId, password)`: スプレッドシートでid+pwを照合

### 認証ロジック
- pwはBASE64デコードして比較
- flag=0のユーザーのみ有効
- expiryが現在日時より未来のユーザーのみ有効

---

## タスク2: gas/doPost.js を新規作成（認証付き）

### エントリポイント
doGet / doPost を実装。全レスポンスにCORSヘッダーを付ける。

### action一覧

**認証不要:**
- `login`: id+pwを受け取り → getUserFromSheet → createSession → tokenとsname返す

**認証必要（tokenが必須）:**
- `getLeadTimeList`: getRakutenLeadTime()の結果をJSONで返す
- `searchItems`: searchRakutenItems()の結果をJSONで返す
- `updateLeadTime`: updateInventoryAndLeadTime()を実行

### 認証フロー
```javascript
// 認証必要なactionの処理例
const token = params.token;
const creds = validateSession(token);
if (!creds) return errorResponse('認証エラー');
// credsにlicenseKeyとserviceSecretが入っている
// 既存関数に渡して実行
```

---

## タスク3: web/index.html を新規作成

### 画面構成
1. **ログイン画面**: id・パスワード入力 → loginアクション → token保存（sessionStorage）
2. **メイン画面**（ログイン後）:
   - ヘッダー: 「リードタイム修正侍」+ 店舗名表示 + ログアウトボタン
   - タブ: [商品検索] [LT一括変更] [スケジュール]
3. **商品検索タブ**: キーワード入力 → searchItems → テーブル表示
4. **LT一括変更タブ**: LT一覧セレクトボックス → 対象商品チェック → 一括更新
5. **スケジュールタブ**: 日付指定での自動変更設定

### セキュリティ
- tokenはsessionStorageに保存（タブを閉じると消える）
- 全APIリクエストにtokenを含める
- 401/認証エラー時はログイン画面に戻す

### 定数（先頭に定義）
```javascript
const GAS_URL = 'YOUR_GAS_DEPLOY_URL_HERE';
```

### デザイン
- 和モダン: 深緑(#2d5a27)・金(#c9a84c)・白(#fafaf5)基調
- フォント: Noto Sans JP
- レスポンシブ対応

---

## 注意事項
- APIキーはGASのPropertiesServiceで管理（既存）
- ユーザー認証はスプレッドシートで管理
- licenseKey/serviceSecretは絶対にフロントに返さない
- 楽天APIのQPS制限: sleep(1500)が必要
- GAS_URLはデプロイ後に差し替える
