/**
 * OAuth設定テスト・実行用スクリプト
 *
 * Apps Scriptエディタから直接実行してください
 */

/**
 * 【テスト1】設定シートの内容確認
 *
 * RuName、App ID、Cert IDなどが正しく設定されているか確認
 *
 * @param {string} spreadsheetId スプレッドシートID（省略時はデフォルト）
 */
function testCheckOAuthSettings(spreadsheetId) {
  try {
    // バインドスクリプトから呼び出される場合はspreadsheetIdが渡される
    if (!spreadsheetId) {
      spreadsheetId = '1gGoJSu-ckMllYWuFCoERGVIPBDGvpVVRHDStx58MEgQ';
    }
    CURRENT_SPREADSHEET_ID = spreadsheetId;

    Logger.log('=== OAuth設定確認 ===');
    Logger.log('');

    const config = getEbayConfig();

    Logger.log('App ID: ' + (config.appId ? '✅ 設定済み (' + config.appId.substring(0, 20) + '...)' : '❌ 未設定'));
    Logger.log('Cert ID: ' + (config.certId ? '✅ 設定済み (' + config.certId.substring(0, 20) + '...)' : '❌ 未設定'));
    Logger.log('Dev ID: ' + (config.devId ? '✅ 設定済み (' + config.devId.substring(0, 20) + '...)' : '❌ 未設定'));
    Logger.log('RuName: ' + (config.ruName ? '✅ 設定済み (' + config.ruName + ')' : '❌ 未設定'));
    Logger.log('User Token: ' + (config.userToken ? '✅ 設定済み (' + config.userToken.substring(0, 20) + '...)' : '⚠️ 未設定（OAuth認証後に自動入力）'));
    Logger.log('Refresh Token: ' + (config.refreshToken ? '✅ 設定済み (' + config.refreshToken.substring(0, 20) + '...)' : '⚠️ 未設定（OAuth認証後に自動入力）'));
    Logger.log('Token Expiry: ' + (config.tokenExpiry ? '✅ 設定済み (' + config.tokenExpiry + ')' : '⚠️ 未設定（OAuth認証後に自動入力）'));
    Logger.log('');

    // 必須項目チェック
    const requiredFields = [
      { name: 'App ID', value: config.appId },
      { name: 'Cert ID', value: config.certId },
      { name: 'Dev ID', value: config.devId },
      { name: 'RuName', value: config.ruName }
    ];

    const missingFields = requiredFields.filter(field => !field.value);

    if (missingFields.length > 0) {
      Logger.log('❌ 以下の必須項目が未設定です:');
      missingFields.forEach(field => Logger.log('   - ' + field.name));
      Logger.log('');
      Logger.log('ツール設定シートで以下を確認してください:');
      Logger.log('1. A列に項目名が入力されている');
      Logger.log('2. B列に対応する値が入力されている');
      return { success: false, missingFields: missingFields };
    }

    Logger.log('✅ 必須項目はすべて設定されています');
    Logger.log('');
    Logger.log('次のステップ: testGenerateAuthUrl() を実行してOAuth認証URLを生成してください');

    return { success: true, config: config };

  } catch (error) {
    Logger.log('❌ エラー: ' + error.toString());
    Logger.log('');
    Logger.log('エラーの原因:');
    Logger.log('1. スプレッドシートIDが正しいか確認');
    Logger.log('2. "ツール設定"シートが存在するか確認');
    Logger.log('3. A列に項目名、B列に値が入力されているか確認');
    return { success: false, error: error.toString() };
  } finally {
    CURRENT_SPREADSHEET_ID = null;
  }
}

/**
 * 【テスト2】OAuth認証URL生成
 *
 * このURLをブラウザで開いてeBayにサインインします
 *
 * @param {string} spreadsheetId スプレッドシートID（省略時はデフォルト）
 */
function testGenerateAuthUrl(spreadsheetId) {
  try {
    if (!spreadsheetId) {
      spreadsheetId = '1gGoJSu-ckMllYWuFCoERGVIPBDGvpVVRHDStx58MEgQ';
    }
    CURRENT_SPREADSHEET_ID = spreadsheetId;

    Logger.log('=== OAuth認証URL生成 ===');
    Logger.log('');

    const authUrl = generateAuthUrl();

    Logger.log('');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('📋 次のステップ:');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('');
    Logger.log('1. 上記のURLをコピー');
    Logger.log('2. ブラウザの新しいタブで開く');
    Logger.log('3. eBayにサインイン');
    Logger.log('4. アプリケーションへのアクセスを許可');
    Logger.log('5. リダイレクト後のURLから "code=" の値をコピー');
    Logger.log('6. testExchangeTokens("コピーしたコード") を実行');
    Logger.log('');
    Logger.log('⚠️ Authorization Codeは5分で期限切れになります');
    Logger.log('');

    return { success: true, authUrl: authUrl };

  } catch (error) {
    Logger.log('❌ エラー: ' + error.toString());
    return { success: false, error: error.toString() };
  } finally {
    CURRENT_SPREADSHEET_ID = null;
  }
}

/**
 * 【テスト3】Authorization CodeをAccess Token + Refresh Tokenに交換
 *
 * @param {string} spreadsheetId スプレッドシートID
 * @param {string} authorizationCode testGenerateAuthUrl()で取得したコード
 */
function testExchangeTokens(spreadsheetId, authorizationCode) {
  try {
    // バインドスクリプトから呼び出される場合、第1引数がspreadsheetId
    // スタンドアロンから呼び出される場合、第1引数がauthorizationCode
    if (authorizationCode === undefined) {
      // スタンドアロンから呼び出し: testExchangeTokens("code")
      authorizationCode = spreadsheetId;
      spreadsheetId = '1gGoJSu-ckMllYWuFCoERGVIPBDGvpVVRHDStx58MEgQ';
    }
    CURRENT_SPREADSHEET_ID = spreadsheetId;

    if (!authorizationCode) {
      Logger.log('❌ エラー: Authorization Codeが指定されていません');
      Logger.log('');
      Logger.log('使い方:');
      Logger.log('testExchangeTokens("v%5E1.1%...your-code...")');
      return { success: false, error: 'Authorization Code not provided' };
    }

    Logger.log('=== トークン取得開始 ===');
    Logger.log('');

    const result = exchangeCodeForTokens(authorizationCode);

    Logger.log('');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('✅ OAuth認証完了！');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('');
    Logger.log('ツール設定シートに以下が保存されました:');
    Logger.log('• User Token (Access Token)');
    Logger.log('• Refresh Token');
    Logger.log('• Token Expiry');
    Logger.log('');
    Logger.log('次のステップ: testAutoRefresh() でトークン自動更新をテストしてください');
    Logger.log('');

    return { success: true, result: result };

  } catch (error) {
    Logger.log('❌ エラー: ' + error.toString());
    Logger.log('');
    Logger.log('よくあるエラー:');
    Logger.log('1. Authorization Codeの期限切れ（5分）');
    Logger.log('   → testGenerateAuthUrl() から再実行');
    Logger.log('2. RuNameの不一致');
    Logger.log('   → eBay Developer Portalで設定したRuNameと一致しているか確認');
    Logger.log('3. App ID / Cert IDの誤り');
    Logger.log('   → ツール設定シートの値を確認');
    return { success: false, error: error.toString() };
  } finally {
    CURRENT_SPREADSHEET_ID = null;
  }
}

/**
 * 【テスト4】トークン自動更新のテスト
 *
 * @param {string} spreadsheetId スプレッドシートID（省略時はデフォルト）
 */
function testAutoRefresh(spreadsheetId) {
  try {
    if (!spreadsheetId) {
      spreadsheetId = '1gGoJSu-ckMllYWuFCoERGVIPBDGvpVVRHDStx58MEgQ';
    }
    CURRENT_SPREADSHEET_ID = spreadsheetId;

    Logger.log('=== トークン自動更新テスト ===');
    Logger.log('');

    // トークン有効期限チェック
    const isExpired = isTokenExpired();

    if (isExpired) {
      Logger.log('⚠️ トークンが期限切れです。自動更新を実行します...');
      Logger.log('');

      const result = refreshEbayAccessToken(spreadsheetId);

      if (result.success) {
        Logger.log('✅ トークン自動更新成功！');
        Logger.log('新しい有効期限: ' + result.expiryDate.toLocaleString('ja-JP'));
      } else {
        Logger.log('❌ トークン更新失敗: ' + result.error);
      }
    } else {
      Logger.log('✅ トークンは有効です');
      Logger.log('');
      Logger.log('強制的に更新をテストする場合:');
      Logger.log('1. ツール設定シートの"Token Expiry"を過去の日時に変更');
      Logger.log('2. 再度このテストを実行');
    }

    Logger.log('');
    Logger.log('次のステップ: testGetPolicies() でポリシー取得をテストしてください');

    return { success: true };

  } catch (error) {
    Logger.log('❌ エラー: ' + error.toString());
    return { success: false, error: error.toString() };
  } finally {
    CURRENT_SPREADSHEET_ID = null;
  }
}

/**
 * 【テスト5】ポリシー取得テスト（トークン自動更新統合確認）
 *
 * @param {string} spreadsheetId スプレッドシートID（省略時はデフォルト）
 */
function testGetPolicies(spreadsheetId) {
  try {
    if (!spreadsheetId) {
      spreadsheetId = '1gGoJSu-ckMllYWuFCoERGVIPBDGvpVVRHDStx58MEgQ';
    }

    Logger.log('=== ポリシー取得テスト ===');
    Logger.log('');
    Logger.log('トークン自動更新機能が正しく統合されているか確認します');
    Logger.log('');

    const result = menuGetPolicies(spreadsheetId);

    if (result.success) {
      Logger.log('');
      Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
      Logger.log('✅ 全テスト完了！');
      Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
      Logger.log('');
      Logger.log('OAuth自動更新機能が正常に動作しています:');
      Logger.log('• トークン有効期限チェック: OK');
      Logger.log('• 自動更新: OK');
      Logger.log('• eBay API呼び出し: OK');
      Logger.log('');
      Logger.log('今後、トークンは自動的に更新されます（2時間ごと）');
      Logger.log('Refresh Tokenの有効期限: 18ヶ月');
      Logger.log('');
    } else {
      Logger.log('❌ ポリシー取得失敗: ' + result.error);
    }

    return result;

  } catch (error) {
    Logger.log('❌ エラー: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * 【まとめ】全テストを順番に実行するガイド
 *
 * @param {string} spreadsheetId スプレッドシートID（省略可）
 */
function showTestGuide(spreadsheetId) {
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('📋 OAuth設定・テスト実行ガイド');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('');
  Logger.log('【事前準備】ツール設定シートに以下を設定してください:');
  Logger.log('');
  Logger.log('| 項目名 (A列)   | 値 (B列)                    |');
  Logger.log('|----------------|----------------------------|');
  Logger.log('| App ID         | eBay Developer Portalから取得 |');
  Logger.log('| Cert ID        | eBay Developer Portalから取得 |');
  Logger.log('| Dev ID         | eBay Developer Portalから取得 |');
  Logger.log('| RuName         | eBay Developer Portalで生成   |');
  Logger.log('| User Token     | （空欄でOK・自動入力）        |');
  Logger.log('| Refresh Token  | （空欄でOK・自動入力）        |');
  Logger.log('| Token Expiry   | （空欄でOK・自動入力）        |');
  Logger.log('');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('【テスト実行手順】');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('');
  Logger.log('1️⃣ testCheckOAuthSettings()');
  Logger.log('   → 設定値の確認');
  Logger.log('');
  Logger.log('2️⃣ testGenerateAuthUrl()');
  Logger.log('   → OAuth認証URL生成');
  Logger.log('   → URLをブラウザで開いてAuthorization Code取得');
  Logger.log('');
  Logger.log('3️⃣ testExchangeTokens("取得したコード")');
  Logger.log('   → Access Token + Refresh Token取得');
  Logger.log('');
  Logger.log('4️⃣ testAutoRefresh()');
  Logger.log('   → トークン自動更新テスト');
  Logger.log('');
  Logger.log('5️⃣ testGetPolicies()');
  Logger.log('   → ポリシー取得テスト（統合確認）');
  Logger.log('');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('');
  Logger.log('準備ができたら 1️⃣ から順番に実行してください');
  Logger.log('');
}
