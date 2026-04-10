/**
 * eBay OAuth 完全自動更新システム
 *
 * 誰もシートを触らなくても18ヶ月動き続ける設計
 * - 1時間ごとの自動チェック（時間トリガー）
 * - Refresh Token期限の監視（30日前・7日前にSlack通知）
 * - LockServiceで同時実行を防止
 */

/**
 * 定期実行トリガー関数（1時間ごとに実行）
 *
 * この関数を時間トリガーに設定することで、
 * 誰もスプレッドシートを開かなくても自動的にトークンが更新されます
 */
function scheduledTokenRefresh() {
  const lock = LockService.getScriptLock();

  try {
    // 同時実行の競合を防ぐ（複数トリガーが同時に走っても安全）
    if (!lock.tryLock(10000)) {
      Logger.log('⚠️ 他の処理が実行中のためスキップしました');
      return;
    }

    Logger.log('=== 定期トークンチェック開始 ===');

    // 全クライアントのスプレッドシートIDを取得
    const clientIds = getAllClientSpreadsheetIds();

    if (clientIds.length === 0) {
      Logger.log('⚠️ 登録されているクライアントがありません');
      return;
    }

    Logger.log('対象クライアント数: ' + clientIds.length + '件');

    // 各クライアントのトークンをチェック・更新
    clientIds.forEach(function(spreadsheetId) {
      try {
        Logger.log('');
        Logger.log('--- クライアント: ' + spreadsheetId + ' ---');

        // トークン状態をチェック
        const status = checkAndRefreshToken(spreadsheetId);

        if (status.refreshed) {
          Logger.log('✅ トークン自動更新完了');
        } else if (status.valid) {
          Logger.log('✅ トークン有効（更新不要）');
        }

        // Refresh Token期限の警告チェック
        checkRefreshTokenExpiry(spreadsheetId);

      } catch (error) {
        Logger.log('❌ エラー: ' + error.toString());
      }
    });

    Logger.log('');
    Logger.log('=== 定期トークンチェック完了 ===');

  } catch (error) {
    Logger.log('❌ 定期チェックエラー: ' + error.toString());
  } finally {
    lock.releaseLock();
  }
}

/**
 * トークンをチェックして必要に応じて更新
 *
 * @param {string} spreadsheetId スプレッドシートID
 * @returns {Object} { valid: boolean, refreshed: boolean, error?: string }
 */
function checkAndRefreshToken(spreadsheetId) {
  try {
    // トークン期限をチェック
    const isExpired = isTokenExpired(spreadsheetId);

    if (!isExpired) {
      return { valid: true, refreshed: false };
    }

    // 期限切れ → 自動更新
    Logger.log('トークンが期限切れ → 自動更新を実行');
    const result = refreshEbayAccessToken(spreadsheetId);

    if (result.success) {
      return { valid: true, refreshed: true };
    } else {
      return { valid: false, refreshed: false, error: result.error };
    }

  } catch (error) {
    return { valid: false, refreshed: false, error: error.toString() };
  }
}

/**
 * Refresh Token の期限をチェックしてSlack通知
 *
 * @param {string} spreadsheetId スプレッドシートID
 */
function checkRefreshTokenExpiry(spreadsheetId) {
  try {
    if (spreadsheetId) {
      CURRENT_SPREADSHEET_ID = spreadsheetId;
    }

    const config = getConfig();
    const refreshTokenExpiryStr = config['Refresh Token Expiry'];

    if (!refreshTokenExpiryStr) {
      Logger.log('⚠️ Refresh Token Expiryが設定されていません');
      return;
    }

    const expiryDate = new Date(refreshTokenExpiryStr);

    if (isNaN(expiryDate.getTime())) {
      Logger.log('⚠️ Refresh Token Expiryのパースに失敗: ' + refreshTokenExpiryStr);
      return;
    }

    const now = new Date();
    const remainingMs = expiryDate.getTime() - now.getTime();
    const remainingDays = Math.floor(remainingMs / (1000 * 60 * 60 * 24));

    Logger.log('Refresh Token残り日数: ' + remainingDays + '日');

    // 期限切れ
    if (remainingDays < 0) {
      const message = '🚨 *eBay Refresh Token期限切れ*\n\n' +
                      'スプレッドシート: ' + spreadsheetId + '\n' +
                      '期限: ' + expiryDate.toLocaleString('ja-JP') + '\n\n' +
                      '対処: 手動で再認証が必要です。\n' +
                      '「初期設定」→「🔗 OAuth認証URL生成」から再認証してください。';

      sendSlackNotification(message);
      Logger.log('❌ Refresh Token期限切れ（Slack通知送信）');
      return;
    }

    // 7日前警告
    if (remainingDays <= 7) {
      const message = '🔴 *eBay Refresh Token残り' + remainingDays + '日*\n\n' +
                      'スプレッドシート: ' + spreadsheetId + '\n' +
                      '期限: ' + expiryDate.toLocaleString('ja-JP') + '\n\n' +
                      '⚠️ 今すぐ再認証してください！\n' +
                      '放置すると自動化が停止します。\n\n' +
                      '「初期設定」→「🔗 OAuth認証URL生成」から再認証してください。';

      sendSlackNotification(message);
      Logger.log('🔴 Refresh Token残り7日以内（Slack通知送信）');
      return;
    }

    // 30日前警告
    if (remainingDays <= 30) {
      const message = '⚠️ *eBay Refresh Token残り' + remainingDays + '日*\n\n' +
                      'スプレッドシート: ' + spreadsheetId + '\n' +
                      '期限: ' + expiryDate.toLocaleString('ja-JP') + '\n\n' +
                      '30日以内に再認証することを推奨します。\n\n' +
                      '「初期設定」→「🔗 OAuth認証URL生成」から再認証してください。';

      sendSlackNotification(message);
      Logger.log('⚠️ Refresh Token残り30日以内（Slack通知送信）');
      return;
    }

  } catch (error) {
    Logger.log('❌ Refresh Token期限チェックエラー: ' + error.toString());
  } finally {
    CURRENT_SPREADSHEET_ID = null;
  }
}

/**
 * Slack通知を送信
 *
 * @param {string} message 通知メッセージ
 */
function sendSlackNotification(message) {
  try {
    const webhookUrl = getSlackWebhookUrl();

    if (!webhookUrl) {
      Logger.log('⚠️ Slack Webhook URLが設定されていません（通知スキップ）');
      return;
    }

    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ text: message }),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(webhookUrl, options);
    const statusCode = response.getResponseCode();

    if (statusCode === 200) {
      Logger.log('✅ Slack通知送信成功');
    } else {
      Logger.log('❌ Slack通知送信失敗: ' + statusCode);
    }

  } catch (error) {
    Logger.log('❌ Slack通知エラー: ' + error.toString());
  }
}

/**
 * Slack Webhook URLを取得
 *
 * @returns {string} Webhook URL（未設定の場合は空文字列）
 */
function getSlackWebhookUrl() {
  // スクリプトプロパティから取得
  const props = PropertiesService.getScriptProperties();
  const url = props.getProperty('SLACK_WEBHOOK_URL');

  if (url) {
    return url;
  }

  // 環境変数から取得（将来の拡張用）
  return '';
}

/**
 * 時間トリガーを設定（1時間ごとに実行）
 *
 * この関数を1回だけ実行することで、以降は自動的に1時間ごとに
 * scheduledTokenRefresh()が実行されます
 */
function setupTimeTrigger() {
  try {
    // 既存のトリガーを削除（重複防止）
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(function(trigger) {
      if (trigger.getHandlerFunction() === 'scheduledTokenRefresh') {
        ScriptApp.deleteTrigger(trigger);
        Logger.log('既存のトリガーを削除しました');
      }
    });

    // 1時間ごとに実行するトリガーを作成
    ScriptApp.newTrigger('scheduledTokenRefresh')
      .timeBased()
      .everyHours(1)
      .create();

    Logger.log('✅ 時間トリガー設定完了');
    Logger.log('');
    Logger.log('設定内容:');
    Logger.log('- 関数: scheduledTokenRefresh');
    Logger.log('- 実行間隔: 1時間ごと');
    Logger.log('');
    Logger.log('これで誰もシートを触らなくても自動的にトークンが更新されます。');
    Logger.log('Refresh Tokenの有効期限（約18ヶ月）まで完全自動で動作します。');

  } catch (error) {
    Logger.log('❌ トリガー設定エラー: ' + error.toString());
    throw error;
  }
}

/**
 * 時間トリガーを削除
 *
 * 自動更新を停止したい場合に実行
 */
function removeTimeTrigger() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    let count = 0;

    triggers.forEach(function(trigger) {
      if (trigger.getHandlerFunction() === 'scheduledTokenRefresh') {
        ScriptApp.deleteTrigger(trigger);
        count++;
      }
    });

    if (count > 0) {
      Logger.log('✅ 時間トリガーを削除しました（' + count + '件）');
    } else {
      Logger.log('⚠️ 削除対象のトリガーが見つかりませんでした');
    }

  } catch (error) {
    Logger.log('❌ トリガー削除エラー: ' + error.toString());
    throw error;
  }
}

/**
 * Slack Webhook URLを設定
 *
 * @param {string} webhookUrl Slack Incoming Webhook URL
 */
function setSlackWebhookUrl(webhookUrl) {
  try {
    const props = PropertiesService.getScriptProperties();
    props.setProperty('SLACK_WEBHOOK_URL', webhookUrl);

    Logger.log('✅ Slack Webhook URL設定完了');
    Logger.log('');
    Logger.log('テスト通知を送信します...');

    // テスト通知
    sendSlackNotification(
      '✅ *eBay OAuth自動更新システム*\n\n' +
      'Slack通知が正常に設定されました。\n\n' +
      '今後、以下のタイミングで通知が送信されます:\n' +
      '- Refresh Token残り30日\n' +
      '- Refresh Token残り7日\n' +
      '- Refresh Token期限切れ\n' +
      '- トークン自動更新失敗'
    );

  } catch (error) {
    Logger.log('❌ Slack Webhook URL設定エラー: ' + error.toString());
    throw error;
  }
}

/**
 * 現在のトリガー状態を確認
 */
function checkTriggerStatus() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    const scheduledTriggers = triggers.filter(function(trigger) {
      return trigger.getHandlerFunction() === 'scheduledTokenRefresh';
    });

    Logger.log('=== トリガー状態確認 ===');
    Logger.log('');

    if (scheduledTriggers.length === 0) {
      Logger.log('❌ 時間トリガーが設定されていません');
      Logger.log('');
      Logger.log('setupTimeTrigger() を実行してトリガーを設定してください。');
      return false;
    }

    Logger.log('✅ 時間トリガーが設定されています');
    Logger.log('');
    Logger.log('トリガー数: ' + scheduledTriggers.length + '件');

    scheduledTriggers.forEach(function(trigger, index) {
      Logger.log('');
      Logger.log('トリガー ' + (index + 1) + ':');
      Logger.log('- 関数: ' + trigger.getHandlerFunction());
      Logger.log('- 種類: ' + trigger.getEventType());

      // 時間ベーストリガーの詳細
      const triggerSource = trigger.getTriggerSource();
      if (triggerSource === ScriptApp.TriggerSource.CLOCK) {
        Logger.log('- 実行間隔: 1時間ごと');
      }
    });

    Logger.log('');
    Logger.log('======================');

    return true;

  } catch (error) {
    Logger.log('❌ トリガー状態確認エラー: ' + error.toString());
    return false;
  }
}

/**
 * 【手動テスト】トークン自動更新を1回だけ実行
 *
 * この関数を実行すると、scheduledTokenRefresh()と同じ処理が1回だけ実行されます。
 * トリガーを待たずに、今すぐ自動更新をテストできます。
 *
 * 使い方:
 * 1. Apps Scriptエディタで「実行する関数」から testTokenRefreshOnce を選択
 * 2. 「実行」ボタンをクリック
 * 3. ログを確認
 */
function testTokenRefreshOnce() {
  try {
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('🧪 手動テスト: トークン自動更新');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('');

    // 全クライアントのスプレッドシートIDを取得
    const clientIds = getAllClientSpreadsheetIds();

    if (clientIds.length === 0) {
      Logger.log('❌ 登録されているクライアントがありません');
      Logger.log('');
      Logger.log('対処方法:');
      Logger.log('1. ClientManager.gsでクライアントを登録');
      Logger.log('2. または、testTokenRefreshForClient("spreadsheetId") を実行');
      return;
    }

    Logger.log('テスト対象クライアント数: ' + clientIds.length + '件');
    Logger.log('');

    // 各クライアントのトークンをチェック・更新
    clientIds.forEach(function(spreadsheetId, index) {
      Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
      Logger.log('クライアント ' + (index + 1) + '/' + clientIds.length);
      Logger.log('Spreadsheet ID: ' + spreadsheetId);
      Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
      Logger.log('');

      try {
        // トークン状態を確認
        const tokenStatus = debugTokenStatus(spreadsheetId);
        Logger.log('');

        // トークンチェック・更新
        const status = checkAndRefreshToken(spreadsheetId);

        if (status.refreshed) {
          Logger.log('✅ トークン自動更新完了');
        } else if (status.valid) {
          Logger.log('✅ トークン有効（更新不要）');
        } else {
          Logger.log('❌ トークン更新失敗: ' + (status.error || '不明なエラー'));
        }

        Logger.log('');

        // Refresh Token期限チェック
        checkRefreshTokenExpiry(spreadsheetId);

      } catch (error) {
        Logger.log('❌ エラー: ' + error.toString());
      }

      Logger.log('');
    });

    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('✅ テスト完了');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');

  } catch (error) {
    Logger.log('❌ テストエラー: ' + error.toString());
  }
}

/**
 * 【手動テスト】特定のクライアントのトークン自動更新を1回だけ実行
 *
 * @param {string} spreadsheetId テストするスプレッドシートID
 */
function testTokenRefreshForClient(spreadsheetId) {
  try {
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('🧪 手動テスト: トークン自動更新（単一クライアント）');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('');
    Logger.log('Spreadsheet ID: ' + spreadsheetId);
    Logger.log('');

    // トークン状態を確認
    Logger.log('--- トークン状態確認 ---');
    const tokenStatus = debugTokenStatus(spreadsheetId);
    Logger.log('');

    // トークンチェック・更新
    Logger.log('--- トークンチェック・更新 ---');
    const status = checkAndRefreshToken(spreadsheetId);

    if (status.refreshed) {
      Logger.log('✅ トークン自動更新完了');
    } else if (status.valid) {
      Logger.log('✅ トークン有効（更新不要）');
    } else {
      Logger.log('❌ トークン更新失敗: ' + (status.error || '不明なエラー'));
    }

    Logger.log('');

    // Refresh Token期限チェック
    Logger.log('--- Refresh Token期限チェック ---');
    checkRefreshTokenExpiry(spreadsheetId);

    Logger.log('');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('✅ テスト完了');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');

  } catch (error) {
    Logger.log('❌ テストエラー: ' + error.toString());
  }
}
