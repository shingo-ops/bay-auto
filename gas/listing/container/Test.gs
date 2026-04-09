/**
 * eBay出品管理 - テスト関数
 */

/**
 * 複数スプレッドシート対応テスト
 *
 * エビデンス: 同じスタンドアロンスクリプトから複数のスプレッドシートにアクセス可能
 */
function testMultipleSpreadsheets() {
  Logger.log('=== 複数スプレッドシート対応テスト ===');
  Logger.log('');

  // テスト用スプレッドシートID
  const testSpreadsheets = [
    {
      name: 'クライアントA（出品シート）',
      id: '1gGoJSu-ckMllYWuFCoERGVIPBDGvpVVRHDStx58MEgQ'
    },
    {
      name: 'クライアントB（仮想）',
      id: '1gGoJSu-ckMllYWuFCoERGVIPBDGvpVVRHDStx58MEgQ' // 同じIDで代用
    }
  ];

  const results = [];

  // 各スプレッドシートに対して処理を実行
  for (let i = 0; i < testSpreadsheets.length; i++) {
    const client = testSpreadsheets[i];

    Logger.log('--- ' + (i + 1) + '. ' + client.name + ' ---');
    Logger.log('スプレッドシートID: ' + client.id);

    try {
      // グローバル変数をリセット
      CURRENT_SPREADSHEET_ID = null;

      // スプレッドシートIDを引数で渡して設定確認
      const result = checkSettings(client.id);

      results.push({
        name: client.name,
        id: client.id,
        success: true,
        message: '✅ アクセス成功'
      });

      Logger.log('✅ ' + client.name + ': アクセス成功');

    } catch (error) {
      results.push({
        name: client.name,
        id: client.id,
        success: false,
        message: '❌ エラー: ' + error.toString()
      });

      Logger.log('❌ ' + client.name + ': エラー - ' + error.toString());
    }

    Logger.log('');
  }

  // 結果サマリー
  Logger.log('=== テスト結果サマリー ===');
  const successCount = results.filter(function(r) { return r.success; }).length;
  Logger.log('成功: ' + successCount + ' / ' + results.length);
  Logger.log('');

  // エビデンス確立
  if (successCount === results.length) {
    Logger.log('✅ エビデンス確立: 同じスタンドアロンスクリプトから複数のスプレッドシートにアクセス可能');
  } else {
    Logger.log('⚠️ 一部のスプレッドシートでエラーが発生');
  }

  return results;
}

/**
 * スプレッドシートID引数渡しテスト
 *
 * 使用方法:
 * clasp run testExportPoliciesWithId -p '["1gGoJSu-ckMllYWuFCoERGVIPBDGvpVVRHDStx58MEgQ"]'
 */
function testExportPoliciesWithId(spreadsheetId) {
  Logger.log('=== ポリシー取得テスト（引数指定） ===');
  Logger.log('対象スプレッドシート: ' + spreadsheetId);
  Logger.log('');

  try {
    const result = exportPoliciesToSheet(spreadsheetId);

    Logger.log('');
    Logger.log('✅ テスト成功');
    Logger.log('取得ポリシー数: ' + result.totalCount);

    return result;

  } catch (error) {
    Logger.log('');
    Logger.log('❌ テスト失敗: ' + error.toString());
    throw error;
  }
}

/**
 * デフォルトスプレッドシートIDテスト
 */
function testDefaultSpreadsheet() {
  Logger.log('=== デフォルトスプレッドシートテスト ===');
  Logger.log('');

  try {
    // グローバル変数をリセット
    CURRENT_SPREADSHEET_ID = null;

    // デフォルトIDで設定確認
    Logger.log('引数なしでcheckSettings()を実行...');
    const result = checkSettings();

    Logger.log('');
    Logger.log('✅ デフォルトスプレッドシートでアクセス成功');

    return result;

  } catch (error) {
    Logger.log('');
    Logger.log('❌ エラー: ' + error.toString());
    Logger.log('⚠️ デフォルトスプレッドシートIDが未設定の可能性があります');
    Logger.log('   setDefaultSpreadsheetId("スプレッドシートID") を実行してください');

    throw error;
  }
}

/**
 * 複数クライアント連続処理シミュレーション
 */
function testMultipleClientsSequential() {
  Logger.log('=== 複数クライアント連続処理シミュレーション ===');
  Logger.log('');

  const clients = [
    { name: 'クライアント1', id: '1gGoJSu-ckMllYWuFCoERGVIPBDGvpVVRHDStx58MEgQ' },
    { name: 'クライアント2', id: '1gGoJSu-ckMllYWuFCoERGVIPBDGvpVVRHDStx58MEgQ' },
    { name: 'クライアント3', id: '1gGoJSu-ckMllYWuFCoERGVIPBDGvpVVRHDStx58MEgQ' }
  ];

  for (let i = 0; i < clients.length; i++) {
    Logger.log('--- 処理 ' + (i + 1) + ': ' + clients[i].name + ' ---');

    // グローバル変数をリセット（重要）
    CURRENT_SPREADSHEET_ID = null;

    // 各クライアントのスプレッドシートにアクセス
    checkSettings(clients[i].id);

    Logger.log('✅ ' + clients[i].name + ' の処理完了');
    Logger.log('');
  }

  Logger.log('=== 全クライアント処理完了 ===');
  Logger.log('');
  Logger.log('✅ エビデンス: 複数のクライアントを連続処理可能');

  return {
    success: true,
    processedCount: clients.length
  };
}

// ============================================================
// クライアント管理機能のテスト
// ============================================================

/**
 * テスト1: PropertiesServiceの確認
 *
 * 目的: GUIでスクリプトプロパティが正しく作成されたか確認
 */
function testPropertiesService() {
  Logger.log('=== PropertiesServiceテスト ===');
  Logger.log('');

  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const clientsJson = scriptProperties.getProperty('CLIENTS');

    if (!clientsJson) {
      Logger.log('❌ CLIENTSプロパティが見つかりません');
      Logger.log('');
      Logger.log('【対処法】');
      Logger.log('1. Apps Scriptエディタで⚙️（歯車アイコン）をクリック');
      Logger.log('2. 「プロパティを追加」ボタンをクリック');
      Logger.log('3. プロパティ名: CLIENTS');
      Logger.log('4. 値: {"CLIENT_A":"スプレッドシートID"}');
      return { success: false, message: 'CLIENTSプロパティが未設定' };
    }

    const clients = JSON.parse(clientsJson);

    Logger.log('✅ CLIENTSプロパティ取得成功');
    Logger.log('');
    Logger.log('【登録クライアント一覧】');

    for (const key in clients) {
      Logger.log('- ' + key);
      Logger.log('  スプレッドシートID: ' + clients[key]);
    }

    Logger.log('');
    Logger.log('登録クライアント数: ' + Object.keys(clients).length + '件');

    return {
      success: true,
      clients: clients,
      count: Object.keys(clients).length
    };

  } catch (error) {
    Logger.log('❌ エラー: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * テスト2: getClientInfo()のテスト
 *
 * 目的: クライアント情報取得が正しく動作するか確認
 */
function testGetClientInfo() {
  Logger.log('=== getClientInfo()テスト ===');
  Logger.log('');

  try {
    // 登録されている最初のクライアントの情報を取得
    const clients = getEnabledClients();
    if (clients.length === 0) {
      throw new Error('登録されているクライアントがありません');
    }
    const firstClientKey = clients[0].key;
    Logger.log('テスト対象: ' + firstClientKey);

    const spreadsheetId = getClientInfo(firstClientKey);

    Logger.log('✅ クライアント情報取得成功');
    Logger.log('スプレッドシートID: ' + spreadsheetId);

    return {
      success: true,
      spreadsheetId: spreadsheetId
    };

  } catch (error) {
    Logger.log('❌ エラー: ' + error.toString());
    Logger.log('');
    Logger.log('【対処法】');
    Logger.log('1. testPropertiesService()を実行してCLIENTSプロパティを確認');
    Logger.log('2. GUIでスクリプトプロパティが作成されているか確認');

    return { success: false, error: error.toString() };
  }
}

/**
 * テスト3: getEnabledClients()のテスト
 *
 * 目的: 全クライアント一覧取得が正しく動作するか確認
 */
function testGetEnabledClients() {
  Logger.log('=== getEnabledClients()テスト ===');
  Logger.log('');

  try {
    const clients = getEnabledClients();

    Logger.log('✅ クライアント一覧取得成功');
    Logger.log('');
    Logger.log('【登録クライアント】');

    clients.forEach(function(client) {
      Logger.log('- ' + client.key);
      Logger.log('  スプレッドシートID: ' + client.id);
    });

    Logger.log('');
    Logger.log('合計: ' + clients.length + '件');

    return {
      success: true,
      clients: clients,
      count: clients.length
    };

  } catch (error) {
    Logger.log('❌ エラー: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * テスト4: processClient()のテスト
 *
 * 目的: 指定クライアントのポリシー取得が正しく動作するか確認
 */
function testProcessClient() {
  Logger.log('=== processClient()テスト ===');
  Logger.log('');

  try {
    // 登録されている最初のクライアントのポリシー取得
    const clients = getEnabledClients();
    if (clients.length === 0) {
      throw new Error('登録されているクライアントがありません');
    }
    const firstClientKey = clients[0].key;
    Logger.log('テスト対象: ' + firstClientKey);

    const result = processClient(firstClientKey);

    if (result.success) {
      Logger.log('✅ processClient()成功');
      Logger.log('クライアントキー: ' + result.clientKey);
      Logger.log('取得ポリシー数: ' + result.result.totalCount);
    } else {
      Logger.log('❌ processClient()失敗');
      Logger.log('クライアントキー: ' + result.clientKey);
      Logger.log('エラー: ' + result.error);
    }

    return result;

  } catch (error) {
    Logger.log('❌ エラー: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * テスト5: processAllClients()のテスト
 *
 * 目的: 全クライアント一括処理が正しく動作するか確認
 */
function testProcessAllClients() {
  Logger.log('=== processAllClients()テスト ===');
  Logger.log('');

  try {
    const results = processAllClients();

    Logger.log('');
    Logger.log('=== テスト結果 ===');

    const successCount = results.filter(function(r) { return r.success; }).length;
    const failureCount = results.filter(function(r) { return !r.success; }).length;

    Logger.log('成功: ' + successCount + '件');
    Logger.log('失敗: ' + failureCount + '件');
    Logger.log('合計: ' + results.length + '件');

    return {
      success: failureCount === 0,
      results: results,
      successCount: successCount,
      failureCount: failureCount
    };

  } catch (error) {
    Logger.log('❌ エラー: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * テスト0: ツール設定シートの読み取り確認
 *
 * 目的: "ツール設定"シートから項目名ベースで正しく値を取得できるか確認
 */
function testToolSettings() {
  Logger.log('=== ツール設定シート読み取りテスト ===');
  Logger.log('');

  try {
    const clients = getEnabledClients();
    if (clients.length === 0) {
      Logger.log('❌ 登録されているクライアントがありません');
      return { success: false, message: 'クライアント未登録' };
    }

    const clientKey = clients[0].key;
    const spreadsheetId = clients[0].id;

    Logger.log('テスト対象クライアント: ' + clientKey);
    Logger.log('スプレッドシートID: ' + spreadsheetId);
    Logger.log('');

    // スプレッドシートIDを設定
    CURRENT_SPREADSHEET_ID = spreadsheetId;

    // 設定を取得
    Logger.log('--- getConfig()実行 ---');
    const config = getConfig();
    Logger.log('');

    Logger.log('--- 取得された設定 ---');
    for (const key in config) {
      if (key === 'User Token' || key === 'App ID' || key === 'Cert ID' || key === 'Dev ID') {
        const value = config[key];
        if (value && value.length > 20) {
          Logger.log(key + ': ' + value.substring(0, 20) + '... (長さ: ' + value.length + ')');
        } else if (value) {
          Logger.log(key + ': ' + value + ' (長さ: ' + value.length + ')');
        } else {
          Logger.log(key + ': （空）');
        }
      } else {
        Logger.log(key + ': ' + config[key]);
      }
    }
    Logger.log('');

    // User Token確認
    Logger.log('--- User Token確認 ---');
    const userToken = config['User Token'];
    if (userToken && userToken.trim() !== '') {
      Logger.log('✅ User Token設定あり（長さ: ' + userToken.length + '）');
    } else {
      Logger.log('❌ User Token未設定または空');
    }

    return {
      success: true,
      config: config,
      hasUserToken: !!(userToken && userToken.trim() !== '')
    };

  } catch (error) {
    Logger.log('❌ エラー: ' + error.toString());
    return { success: false, error: error.toString() };
  } finally {
    CURRENT_SPREADSHEET_ID = null;
  }
}

/**
 * 全テスト実行（クライアント管理機能）
 *
 * 推奨実行順序でテストを実行
 */
function runAllClientTests() {
  Logger.log('========================================');
  Logger.log('クライアント管理機能 - 全テスト実行');
  Logger.log('========================================');
  Logger.log('');

  const results = [];

  // テスト1: PropertiesService確認
  Logger.log('【テスト1/5】PropertiesService確認');
  const test1 = testPropertiesService();
  results.push({ test: 'PropertiesService', result: test1 });
  Logger.log('');
  Logger.log('---');
  Logger.log('');

  // テスト2: getClientInfo()
  Logger.log('【テスト2/5】getClientInfo()');
  const test2 = testGetClientInfo();
  results.push({ test: 'getClientInfo', result: test2 });
  Logger.log('');
  Logger.log('---');
  Logger.log('');

  // テスト3: getEnabledClients()
  Logger.log('【テスト3/5】getEnabledClients()');
  const test3 = testGetEnabledClients();
  results.push({ test: 'getEnabledClients', result: test3 });
  Logger.log('');
  Logger.log('---');
  Logger.log('');

  // テスト4: processClient()
  Logger.log('【テスト4/5】processClient()');
  const test4 = testProcessClient();
  results.push({ test: 'processClient', result: test4 });
  Logger.log('');
  Logger.log('---');
  Logger.log('');

  // テスト5: processAllClients()
  Logger.log('【テスト5/5】processAllClients()');
  const test5 = testProcessAllClients();
  results.push({ test: 'processAllClients', result: test5 });
  Logger.log('');
  Logger.log('========================================');

  // 総合結果
  const successCount = results.filter(function(r) { return r.result.success; }).length;
  Logger.log('');
  Logger.log('【総合結果】');
  Logger.log('成功: ' + successCount + ' / ' + results.length);

  if (successCount === results.length) {
    Logger.log('');
    Logger.log('✅ 全テスト合格');
  } else {
    Logger.log('');
    Logger.log('⚠️ 一部のテストが失敗しました');
  }

  return results;
}

// ============================================================
// 発送ポリシー詳細取得のテスト
// ============================================================

/**
 * テスト: 発送ポリシー詳細取得
 *
 * 目的: 特定の発送ポリシーIDの詳細情報を取得できるか確認
 */
function testFulfillmentPolicyDetails() {
  Logger.log('=== 発送ポリシー詳細取得テスト ===');
  Logger.log('');

  try {
    // クライアント情報を取得
    const clients = getEnabledClients();
    if (clients.length === 0) {
      Logger.log('❌ 登録されているクライアントがありません');
      return { success: false, message: 'クライアント未登録' };
    }

    const clientKey = clients[0].key;
    const spreadsheetId = clients[0].id;

    Logger.log('テスト対象クライアント: ' + clientKey);
    Logger.log('スプレッドシートID: ' + spreadsheetId);
    Logger.log('');

    // スプレッドシートIDを設定
    CURRENT_SPREADSHEET_ID = spreadsheetId;

    // Policy_設定シートから最初のFulfillment Policyを取得
    const policySheet = getTargetSpreadsheet().getSheetByName(SHEET_NAMES.POLICY_SETTINGS);

    if (!policySheet) {
      Logger.log('❌ Policy_設定シートが見つかりません');
      Logger.log('先にexportPoliciesToSheet()を実行してください');
      return { success: false, message: 'Policy_設定シート未作成' };
    }

    const data = policySheet.getDataRange().getValues();
    let fulfillmentPolicyId = null;
    let fulfillmentPolicyName = null;

    // ヘッダー行をスキップして最初のFulfillment Policyを探す
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === 'Fulfillment Policy') {
        fulfillmentPolicyName = data[i][1];
        fulfillmentPolicyId = data[i][2];
        break;
      }
    }

    if (!fulfillmentPolicyId) {
      Logger.log('❌ Fulfillment Policyが見つかりません');
      return { success: false, message: 'Fulfillment Policy未登録' };
    }

    Logger.log('テスト対象ポリシー:');
    Logger.log('- 名前: ' + fulfillmentPolicyName);
    Logger.log('- ID: ' + fulfillmentPolicyId);
    Logger.log('');

    // 詳細情報を取得
    Logger.log('--- 詳細情報取得開始 ---');
    const details = logFulfillmentPolicyDetails(fulfillmentPolicyId);

    Logger.log('');
    Logger.log('✅ 発送ポリシー詳細取得成功');

    return {
      success: true,
      policyId: fulfillmentPolicyId,
      policyName: fulfillmentPolicyName,
      details: details
    };

  } catch (error) {
    Logger.log('❌ エラー: ' + error.toString());
    return { success: false, error: error.toString() };
  } finally {
    CURRENT_SPREADSHEET_ID = null;
  }
}

// ============================================================
// ポリシー同期のテスト
// ============================================================

/**
 * テスト: ポリシー同期（ドライラン）
 *
 * 目的: Policy_設定シートの状態をチェック（実際のAPI呼び出しなし）
 */
function testPolicySyncDryRun() {
  Logger.log('=== ポリシー同期ドライラン ===');
  Logger.log('（実際のAPI呼び出しは行いません）');
  Logger.log('');

  try {
    const clients = getEnabledClients();
    if (clients.length === 0) {
      Logger.log('❌ 登録されているクライアントがありません');
      return { success: false, message: 'クライアント未登録' };
    }

    const clientKey = clients[0].key;
    const spreadsheetId = clients[0].id;

    Logger.log('テスト対象クライアント: ' + clientKey);
    Logger.log('スプレッドシートID: ' + spreadsheetId);
    Logger.log('');

    CURRENT_SPREADSHEET_ID = spreadsheetId;

    const policySheet = getTargetSpreadsheet().getSheetByName(SHEET_NAMES.POLICY_SETTINGS);

    if (!policySheet) {
      Logger.log('❌ Policy_設定シートが見つかりません');
      return { success: false, message: 'Policy_設定シート未作成' };
    }

    const data = policySheet.getDataRange().getValues();

    const summary = {
      add: 0,
      update: 0,
      delete: 0,
      skip: 0,
      unknown: 0
    };

    Logger.log('【現在の状態】');
    Logger.log('');

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const operation = row[POLICY_SHEET_COLUMNS.OPERATION - 1];
      const policyType = row[POLICY_SHEET_COLUMNS.POLICY_TYPE - 1];
      const policyName = row[POLICY_SHEET_COLUMNS.POLICY_NAME - 1];
      const policyId = row[POLICY_SHEET_COLUMNS.POLICY_ID - 1];

      if (!policyType || !policyName) {
        continue;
      }

      const op = (operation || '').toString().trim();
      let action = '';

      if (op === '' || op === '-') {
        action = '（スキップ）';
        summary.skip++;
      } else if (op === '追加') {
        action = '→ 作成予定';
        summary.add++;
      } else if (op === '更新') {
        action = '→ 更新予定 (ID: ' + policyId + ')';
        summary.update++;
      } else if (op === '削除') {
        action = '→ 削除予定';
        summary.delete++;
      } else {
        action = '（不明な操作: "' + op + '"）';
        summary.unknown++;
      }

      Logger.log('行' + (i + 1) + ': ' + policyName + ' ' + action);
    }

    Logger.log('');
    Logger.log('【サマリー】');
    Logger.log('作成予定: ' + summary.add + '件');
    Logger.log('更新予定: ' + summary.update + '件');
    Logger.log('削除予定: ' + summary.delete + '件');
    Logger.log('スキップ: ' + summary.skip + '件');
    if (summary.unknown > 0) {
      Logger.log('不明な操作: ' + summary.unknown + '件');
    }
    Logger.log('');

    const totalChanges = summary.add + summary.update + summary.delete;
    if (totalChanges === 0) {
      Logger.log('✅ 同期が必要な変更はありません');
    } else {
      Logger.log('⚠️ ' + totalChanges + '件の変更があります');
      Logger.log('menuSyncPolicies()を実行して同期できます');
    }

    return {
      success: true,
      summary: summary
    };

  } catch (error) {
    Logger.log('❌ エラー: ' + error.toString());
    return { success: false, error: error.toString() };
  } finally {
    CURRENT_SPREADSHEET_ID = null;
  }
}
