# GAS Auth Setup

このサイトでは、個人アカウント情報をブラウザではなく既存の Google Apps Script 側に保存する前提に切り替えています。

## 追加するもの

既存の GAS プロジェクトに [gas-auth-extension.gs](/Users/ootamatadashi/Desktop/就活ナビHP2/gas-auth-extension.gs) を追加してください。

この GAS は、すでに以下の処理を受けている前提です。

- `action: 'read'`
- `action: 'generate'`
- `action: 'save'`

## doPost への組み込み

既存の `doPost(e)` で、JSON を1回だけパースしたあとに認証処理を差し込めば動きます。

```javascript
function doPost(e) {
  var payload = JSON.parse(e.postData.contents || '{}');

  var authResponse = handleAuthAction_(payload);
  if (authResponse) return authResponse;

  // 既存の read / generate / save ロジック...
}
```

すでに `payload` を作っている場合は、その変数を再利用して `handleAuthAction_(payload)` の行だけ追加すれば大丈夫です。

## 自動で作られるシート

認証拡張を入れると、必要に応じて次のシートを自動作成します。

- `auth_users`
- `auth_sessions`
- `auth_liked`

## 必要なら設定する Script Property

Apps Script が対象スプレッドシートに紐づいていない場合は、Script Properties に次を設定してください。

- `AUTH_SPREADSHEET_ID`
- `AUTH_LINE_QR_FOLDER_ID`
- `SITE_BASE_URL`

`AUTH_SPREADSHEET_ID` を未設定にすると `SpreadsheetApp.getActiveSpreadsheet()` を使います。  
`AUTH_LINE_QR_FOLDER_ID` を未設定にすると、LINE QR 画像は実行アカウントのマイドライブ直下に保存されます。
`SITE_BASE_URL` には本番サイトのベース URL を入れてください。例: `https://example.com`

`SITE_BASE_URL` は次の用途で必須です。

- パスワードリセットメール内のリンク生成

## フロントから呼ばれる新しい action

- `authRegister`
- `authLogin`
- `authLogout`
- `authUpdateProfile`
- `authChangePassword`
- `authSetLikedCompanies`

## 保存される内容

- アカウント情報は `auth_users`
- セッショントークンは `auth_sessions`
- 企業マッチの「いいね」は `auth_liked`
- `auth_users` には `志望業界 / 第1〜第3志望企業 / LINE名 / LINE QR の Drive URL` も保存されます

ブラウザ側にも表示高速化用の一時キャッシュは持ちますが、正本データは GAS / スプレッドシート側です。

## 運営側で見る場所

- 登録情報は `auth_users` シートで確認できます
- LINE QR は `auth_users.lineQrDriveUrl` に Drive のURLが入るので、そこから確認できます
- 運営で見やすくするなら `AUTH_LINE_QR_FOLDER_ID` に専用フォルダを指定しておくのがおすすめです

## 反映手順

1. GAS に [gas-auth-extension.gs](/Users/ootamatadashi/Desktop/就活ナビHP2/gas-auth-extension.gs) を追加する
2. `doPost(e)` に `handleAuthAction_(payload)` を組み込む
3. Script Properties に `SITE_BASE_URL` を設定する
4. Web アプリとして再デプロイする
5. デプロイ URL が変わった場合だけ、以下の `GAS_PROXY_URL` を更新する

- [match.html](/Users/ootamatadashi/Desktop/就活ナビHP2/match.html)
- [members.html](/Users/ootamatadashi/Desktop/就活ナビHP2/members.html)
- [orientation.html](/Users/ootamatadashi/Desktop/就活ナビHP2/orientation.html)
- [auth.js](/Users/ootamatadashi/Desktop/就活ナビHP2/auth.js)

## 補足

このリポジトリ側のフロント実装はサーバー保存前提に切り替えてありますが、実際に登録・ログイン・いいね保存をサーバー側へ流すには、上の GAS 更新と再デプロイが必要です。
