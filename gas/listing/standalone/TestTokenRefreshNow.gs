/**
 * 特定のスプレッドシートIDでトークン更新を実行
 * スプレッドシート1: 149FTxOzVtA2FOVKi39swynl8Rfj8meHN5PkE4OrF3wqiw5X0M05GOWjT
 *
 * この関数を実行すると、指定したスプレッドシートのトークンをチェックして
 * 必要に応じて自動更新します。
 */
function testTokenRefreshNow() {
  const spreadsheetId = '149FTxOzVtA2FOVKi39swynl8Rfj8meHN5PkE4OrF3wqiw5X0M05GOWjT';

  try {
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('🧪 トークン更新テスト');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('');
    Logger.log('Spreadsheet ID: ' + spreadsheetId);
    Logger.log('');

    // Step 1: 現在のトークン状態を確認
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('Step 1: トークン状態確認');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('');

    const tokenStatus = debugTokenStatus(spreadsheetId);

    Logger.log('');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('Step 2: トークンチェック・更新実行');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('');

    // Step 2: トークンチェック・更新
    const status = checkAndRefreshToken(spreadsheetId);

    Logger.log('');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('結果');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('');

    if (status.refreshed) {
      Logger.log('✅ トークン自動更新完了');
      Logger.log('');
      Logger.log('新しいトークンがツール設定シートに保存されました。');
      Logger.log('次の有効期限: 約2時間後');
    } else if (status.valid) {
      Logger.log('✅ トークンは有効です（更新不要）');
      Logger.log('');
      Logger.log('現在のトークンをそのまま使用できます。');
    } else {
      Logger.log('❌ トークン更新失敗');
      Logger.log('');
      Logger.log('エラー: ' + (status.error || '不明なエラー'));
      Logger.log('');
      Logger.log('対処方法:');
      Logger.log('1. Refresh Tokenが有効か確認');
      Logger.log('2. 必要に応じて再認証（OAuth認証URL生成 → OAuth認証実行）');
    }

    Logger.log('');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('Step 3: Refresh Token期限チェック');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('');

    // Step 3: Refresh Token期限チェック
    checkRefreshTokenExpiry(spreadsheetId);

    Logger.log('');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('🎉 テスト完了');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');

    return status;

  } catch (error) {
    Logger.log('');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('❌ テストエラー');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('');
    Logger.log('エラー詳細: ' + error.toString());
    Logger.log('');
    Logger.log('スタックトレース:');
    if (error.stack) {
      Logger.log(error.stack);
    }

    throw error;
  }
}

/**
 * 特定のスプレッドシートIDでトークン更新を実行
 * スプレッドシート2: 1GslBHPDPfkQE7XLEpvOgkgCCj30udEFsVLI0adxO_TJRd2GBCYzx3142
 *
 * この関数を実行すると、指定したスプレッドシートのトークンをチェックして
 * 必要に応じて自動更新します。
 */
function testTokenRefreshNow2() {
  const spreadsheetId = '1GslBHPDPfkQE7XLEpvOgkgCCj30udEFsVLI0adxO_TJRd2GBCYzx3142';

  try {
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('🧪 トークン更新テスト（スプレッドシート2）');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('');
    Logger.log('Spreadsheet ID: ' + spreadsheetId);
    Logger.log('');

    // Step 1: 現在のトークン状態を確認
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('Step 1: トークン状態確認');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('');

    const tokenStatus = debugTokenStatus(spreadsheetId);

    Logger.log('');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('Step 2: トークンチェック・更新実行');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('');

    // Step 2: トークンチェック・更新
    const status = checkAndRefreshToken(spreadsheetId);

    Logger.log('');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('結果');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('');

    if (status.refreshed) {
      Logger.log('✅ トークン自動更新完了');
      Logger.log('');
      Logger.log('新しいトークンがツール設定シートに保存されました。');
      Logger.log('次の有効期限: 約2時間後');
    } else if (status.valid) {
      Logger.log('✅ トークンは有効です（更新不要）');
      Logger.log('');
      Logger.log('現在のトークンをそのまま使用できます。');
    } else {
      Logger.log('❌ トークン更新失敗');
      Logger.log('');
      Logger.log('エラー: ' + (status.error || '不明なエラー'));
      Logger.log('');
      Logger.log('対処方法:');
      Logger.log('1. Refresh Tokenが有効か確認');
      Logger.log('2. 必要に応じて再認証（OAuth認証URL生成 → OAuth認証実行）');
    }

    Logger.log('');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('Step 3: Refresh Token期限チェック');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('');

    // Step 3: Refresh Token期限チェック
    checkRefreshTokenExpiry(spreadsheetId);

    Logger.log('');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('🎉 テスト完了');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');

    return status;

  } catch (error) {
    Logger.log('');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('❌ テストエラー');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('');
    Logger.log('エラー詳細: ' + error.toString());
    Logger.log('');
    Logger.log('スタックトレース:');
    if (error.stack) {
      Logger.log(error.stack);
    }

    throw error;
  }
}
