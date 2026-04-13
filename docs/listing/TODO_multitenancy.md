# マルチテナント対応 TODO

## 設計方針（確定）
- マスター台帳SS（1つ）のIDのみスクリプトプロパティに保持
- クライアント一覧はマスター台帳SSのシートで管理
- 各クライアントの設定（出品DB等）は各自のツール設定シートから取得

## 実装待ちタスク

### 1. Config.gs: CLIENTSをシート管理に移行
- getEnabledClients() をスクリプトプロパティではなくマスター台帳SSから読む形に変更
- getClientInfo() も同様に変更
- スクリプトプロパティに CLIENTS_MASTER_SS_ID のみ残す

### 2. ClientManager.gs: processAllClients() の拡張
- 現状はポリシー取得のみ対応
- 出品処理のマルチテナント対応を追加
- readClientPolicies() の未実装部分を完成させる

### 3. マスター台帳SSの作成
- シート構成: クライアントID | SpreadsheetID | 有効/無効
- クライアント追加はシートに1行追加するだけで対応可能

## 現状
- 単一テナント運用（DEFAULT_SPREADSHEET_ID）は正常動作確認済み
- CLIENTS はスクリプトプロパティにJSON形式で手動管理中
