# ebay-gas-automation プロジェクト

## プロジェクト概要
eBay出品自動化のためのGoogle Apps Script（GAS）と
Pythonスクリプトを管理するリポジトリ。

## 技術スタック
- Google Apps Script（GAS）: スプレッドシート操作・eBay API連携
- Python: eBay APIからデータ一括取得
- eBay Metadata API: カテゴリ・コンディション情報取得
- eBay Trading API: 商品出品（AddFixedPriceItem）

## ファイル構成
- gas/: Google Apps Scriptファイル（.gs）
- python/: Pythonスクリプト
- docs/: ドキュメント・仕様書

## コーディングルール
- GASはV8ランタイムで動作するためconst/let/アロー関数などES6+の記法を使用してよい
- 関数名はキャメルケース（例: getConditionsByGroupId）
- コメントは日本語でOK
- エラーハンドリングは必ず入れる
- API認証情報はコードに直接書かない（Secretsを使う）

### 戻り値・エラーハンドリングの原則（CRITICAL）

**全関数の戻り値は必ず `success` 判定を含め、無条件で `success: true` を返さない。**

- オブジェクトを返す関数は `{ success: boolean, ... }` 形式を使う
- 呼び出し側は必ず `if (!result || !result.success)` を確認してから成功処理に進む
- 失敗時は `Logger.log('❌ [関数名] 失敗: ' + errorMsg)` を必ず記録する
- `result.success` チェックなしに `result.itemId` などのフィールドにアクセスしない

```javascript
// NG: 失敗を無視
const result = doSomething();
Logger.log('成功: ' + result.itemId); // result が失敗でも実行される

// OK: 必ず失敗チェックを先に行う
const result = doSomething();
if (!result || !result.success) {
  Logger.log('❌ doSomething 失敗: ' + (result && result.message));
  return { success: false, message: result && result.message };
}
Logger.log('成功: ' + result.itemId);
```

### 失敗パスの実装チェックリスト

Claude Codeへの実装指示に「`result.success` が false の場合のパスも必ず実装し、Logger.log で記録すること」を毎回含める。

## 出品系変更時のテストチェックリスト

出品ツール（listing）に変更を加えた場合、以下3ケースを必ず動作確認する：

1. **正常系**: 全フィールド入力済みで出品 → 成功メッセージ・Item IDが表示される
2. **Brand未入力**: Brand列を空にして出品 → eBay APIがFailureを返し、失敗メッセージが表示される（成功メッセージが出ないこと）
3. **画像なし**: 画像列を全て空にして出品 → EPSアップロードなしで進み、結果（成功 or 失敗）が正しく表示される

## eBay API仕様
- Trading API: XML/SOAP形式
- Metadata API: REST形式
- 認証: OAuth 2.0（Access Token・Refresh Token）
- Access Tokenの有効期限: 2時間（自動更新済み）

## スプレッドシートのシート名
- category_master: カテゴリIDマスター（group_id列あり）
- condition_master: コンディションIDマスター（16行）
- condition_group_map: グループ別conditions_json（26行）
- category_condition_map: カテゴリ×コンディションマッピング

## 開発フロー
- 1機能1PR を原則とする
- develop で実装 → 動作確認 → PR作成 → main マージ（本番反映）
- 複数機能をまとめてdevelopに積み上げない
- PR前に必ず動作確認を行う

## やってはいけないこと
- .envファイルをコミットしない
- APIキー・トークンをコードに直接書かない
- GASでvarを使わない（V8対応のconst/letを使う）
- category_masterを手動で削除しない

## デプロイ方法
- gas/listing/standalone/ のファイルを変更してmainにマージ
  → listing スタンドアロンGAS（本番）に自動でclasp pushされる
- gas/listing/container/ のファイルを変更してmainにマージ
  → listing コンテナバインドGAS（本番）に自動でclasp pushされる
- gas/research/container/ のファイルを変更してmainにマージ
  → research コンテナバインドGAS（本番）に自動でclasp pushされる
- gas/ebay-db/container/ のファイルを変更してdevelopにマージ
  → ebay-db コンテナバインドGAS（開発）に自動でclasp pushされる
- developブランチへのマージ時は各GASの開発用に自動デプロイされる

## 仕様書・ドキュメントの場所
- 出品ツール仕様書: docs/listing/
- リサーチツール仕様書: docs/research/
- **eBay DBシステム設計定義書: ebay-db/docs/ebay_db_design_report.md**
  - ebay-db配下の実装を進める際は必ずこのファイルを参照すること
  - カテゴリマスター・コンディションマスターのスキーマ定義、GASの関数仕様、GitHub Actionsワークフロー、データフローが記載されている

## @claudeへの指示の出し方
- IssueまたはPRのコメントで @claude と書いて指示する
- 例: @claude getConditionsByGroupId関数を実装してください
- 日本語で指示してOK
