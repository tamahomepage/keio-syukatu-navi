# 選考情報スプレッドシート GAS登録手順

## 追加するスプレッドシートID

```
1YRfephgVy8moJLN76MFaIRYUy1xfi4pm7E5JZLvsg34
```

---

## STEP 1：スプレッドシートの列を整える

Googleフォームの回答スプレッドシートに「**selection**」という名前のシートを作成し、
以下の順番で列を並べてください（A列から順番に）。

| 列 | 列名（1行目に入力） | 内容の例 |
|----|------------------|---------|
| A  | 企業名           | 三菱商事（サマーインターン） |
| B  | 締切日           | 2025/06/10 |
| C  | 種別             | インターン または 本選考 |
| D  | リンク           | https://... （選考URLがあれば） |
| E  | 業界             | 商社（コンサル/商社/金融/メーカー/広告・出版・マスコミ/サービス・インフラ/ソフトウェア/ベンチャー/官公庁/その他） |
| F  | 特集             | 大手 または ベンチャー（どちらでもなければ空欄） |

※ Googleフォームの回答が自動でこのシートに入る設定か、
　 または手動でこの形式にコピーして使ってください。

---

## STEP 2：GASスクリプトに追記する

既存の GAS プロジェクトを開き、`doPost(e)` を処理している関数を探して
以下の `SHEET_KEYS` マップに `selection` の行を追加してください。

### 追加するコード

既存スクリプト内に以下のような `sheetKey` → `スプレッドシートID` のマップがあれば、
`selection` の行を追加します：

```javascript
var SHEET_KEYS = {
  // 既存のキー...
  'selection': {
    spreadsheetId: '1YRfephgVy8moJLN76MFaIRYUy1xfi4pm7E5JZLvsg34',
    sheetName: 'selection'  // ← STEP1で作成したシート名
  }
};
```

### `action: 'read'` の処理部分（参考）

GASの `doPost` 内で `action === 'read'` を処理している箇所に
以下のような読み込みロジックが必要です：

```javascript
if (payload.action === 'read') {
  var key = payload.sheetKey;
  var config = SHEET_KEYS[key];
  if (!config) {
    return jsonResponse_({ status: 'error', message: 'sheetKey not found: ' + key });
  }

  // event / faq / shinkan / orientation 以外のシートはログイン済みセッションを必須にする
  if (key !== 'event' && key !== 'faq' && key !== 'shinkan' && key !== 'orientation') {
    getActiveSessionOrThrow_(payload.sessionToken);
  }

  var ss = SpreadsheetApp.openById(config.spreadsheetId);
  var sheet = ss.getSheetByName(config.sheetName);
  if (!sheet) {
    return jsonResponse_({ status: 'error', message: 'Sheet not found: ' + config.sheetName });
  }

  var values = sheet.getDataRange().getValues();
  if (values.length <= 1) {
    return jsonResponse_({ status: 'ok', rows: [] });
  }

  var rows = values.slice(1); // 1行目はヘッダーなのでスキップ
  return jsonResponse_({ status: 'ok', rows: rows });
}
```

---

## STEP 3：スクリプトプロパティに追加（必要な場合）

GASがスプレッドシートを別ファイルとして参照する場合、
アクセス権限のエラーが出ることがあります。

Apps Script → プロジェクトの設定 → スクリプトプロパティ に以下を追加：

```
SELECTION_SPREADSHEET_ID = 1YRfephgVy8moJLN76MFaIRYUy1xfi4pm7E5JZLvsg34
```

コード側で使う場合：
```javascript
var spreadsheetId = PropertiesService.getScriptProperties()
  .getProperty('SELECTION_SPREADSHEET_ID');
var ss = SpreadsheetApp.openById(spreadsheetId);
```

---

## STEP 4：再デプロイ

GASに変更を加えた後は必ず再デプロイが必要です。

1. GAS エディタ右上の「デプロイ」→「デプロイを管理」
2. 既存のデプロイを選択 →「編集」→「バージョン：新しいバージョン」
3. 「デプロイ」をクリック

※ デプロイURLは変わらないため、HTML側の `GAS_PROXY_URL` の更新は不要です。

---

## 確認方法

デプロイ後、ログイン済みの状態で members.html を開き、
エラーが出ていなければ接続成功です。

選考情報セクションにサンプルデータではなく実際のデータが表示されれば完了です。

---

## 列の対応まとめ（再掲）

members.html が期待するスプレッドシートの列順：

```
A列: 企業名（選考の種類も含めて記載）例：三菱商事（サマーインターン）
B列: 締切日　例：2025/06/10
C列: 種別　  例：インターン / 本選考
D列: リンク　例：https://...（なければ空欄）
E列: 業界　  例：商社
F列: 特集　  例：大手 / ベンチャー（なければ空欄）
```
