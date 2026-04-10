/**
 * eBay出品管理 - クライアント側バインドスクリプト
 *
 * このスクリプトは各クライアントのスプレッドシートに設置します
 *
 * セットアップ手順:
 * 1. スプレッドシートで「拡張機能」→「Apps Script」を開く
 * 2. このコードを貼り付け
 * 3. 「ライブラリ」→「ライブラリを追加」
 * 4. スクリプトID: ツール設定シートのライブラリスクリプトID（B18）を使用
 * 5. 識別子: EbayLib
 * 6. バージョン: Head（開発中）または最新バージョン（本番）
 * 7. 保存
 *
 * 図形ボタンに割り当てる関数:
 * - authorizeScript: 権限承認 + handleEdit トリガー登録（初回セットアップ時に実行）
 * - setupEbayManager: eBay API セットアップ
 * - menuGetPolicies: ポリシー取得
 * - menuSyncPolicies: ポリシー更新
 */

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 編集トリガー
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * 【インストール済みトリガー】handleEdit
 *
 * 出品シートの編集時に自動呼び出しされる。
 * authorizeScript() を実行するとトリガーが登録される。
 *
 * ・カテゴリID 列が変更された場合
 *     → 確認ダイアログを表示し、YES なら EbayLib.applyCategoryChange() を実行
 * ・それ以外の列
 *     → EbayLib.processOnEdit() に委譲（タイトル文字数更新など）
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e
 */
function handleEdit(e) {
  try {
    if (!e || !e.range) return;

    const range     = e.range;
    const sheet     = range.getSheet();
    const sheetName = sheet.getName();
    const row       = range.getRow();
    const col       = range.getColumn();

    // 出品シート以外は無視
    if (sheetName !== '出品') return;

    // ヘッダー行（3行目まで）は無視
    if (row <= 3) return;

    const spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();

    // カテゴリID 列の変更かどうかを確認
    const categoryIdCol = EbayLib.getCategoryIdColumnNumber(spreadsheetId);

    if (categoryIdCol && col === categoryIdCol) {
      _handleCategoryIdChange(e, spreadsheetId, sheetName, row);
    } else {
      // 他の列は既存処理（タイトル文字数更新など）に委譲
      EbayLib.processOnEdit(e, spreadsheetId);
    }

  } catch (error) {
    Logger.log('handleEdit エラー: ' + error.toString());
    try {
      SpreadsheetApp.getUi().alert(
        'エラー',
        '❌ 編集処理中にエラーが発生しました:\n' + error.toString(),
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } catch (uiError) {
      // UI 表示自体が失敗した場合は握り潰す
    }
  }
}

/**
 * カテゴリID 列変更時の処理（handleEdit から呼び出す）
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e
 * @param {string} spreadsheetId
 * @param {string} sheetName
 * @param {number} row
 */
function _handleCategoryIdChange(e, spreadsheetId, sheetName, row) {
  const newCategoryId = String(e.value     !== undefined ? e.value     : '').trim();
  const oldCategoryId = String(e.oldValue  !== undefined ? e.oldValue  : '').trim();

  // 値が変わっていなければスキップ
  if (newCategoryId === oldCategoryId) return;

  const oldCategoryName = EbayLib.getCategoryNameById(spreadsheetId, oldCategoryId)
    || (oldCategoryId ? oldCategoryId : '（なし）');
  const newCategoryName = EbayLib.getCategoryNameById(spreadsheetId, newCategoryId)
    || (newCategoryId ? newCategoryId : '（不明）');

  const ui = SpreadsheetApp.getUi();

  // 4-1. 確認ポップアップ
  const response = ui.alert(
    'カテゴリ変更確認',
    '現在: ' + oldCategoryId + ' (' + oldCategoryName + ')\n' +
    '変更先: ' + newCategoryId + ' (' + newCategoryName + ')\n\n' +
    'カテゴリを変更するとスペック情報が初期化されます。\n変更しますか？',
    ui.ButtonSet.YES_NO
  );

  if (response === ui.Button.YES) {
    // 4-2. YES: カテゴリ変更処理を実行
    const result = EbayLib.applyCategoryChange(spreadsheetId, sheetName, row, newCategoryId);

    // コンディション値が新カテゴリと非互換の場合は通知
    if (result && result.conditionIncompatible) {
      ui.alert(
        'コンディション再選択',
        '選択されていた「' + result.oldConditionValue + '」は新しいカテゴリでは使用できません。\n' +
        'プルダウンから再度選択してください。',
        ui.ButtonSet.OK
      );
    }

  } else {
    // 4-3. NO: category_id 列を元の値に戻す
    EbayLib.revertCategoryId(spreadsheetId, sheetName, row, oldCategoryId);
  }
}

/**
 * 【権限承認 + handleEdit トリガー登録】
 *
 * 初回セットアップ時に実行する。図形ボタンに割り当て可。
 *
 * 実行内容:
 *   1. スプレッドシート / ドライブ / 外部URL の権限を一括承認
 *   2. handleEdit トリガーを登録（既存トリガーがある場合はスキップ）
 */
function authorizeScript() {
  const ui = SpreadsheetApp.getUi();
  try {
    // ── 1. 権限承認 ────────────────────────────────────────────────
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    sheet.getRange('A1').getValue();                                        // スプレッドシート権限
    const folders = DriveApp.getFolders();                                  // ドライブ権限
    if (folders.hasNext()) { folders.next(); }
    UrlFetchApp.fetch('https://www.google.com', { muteHttpExceptions: true }); // 外部URL権限

    // ── 2. handleEdit トリガー登録（重複チェック付き）──────────────
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const existing = ScriptApp.getUserTriggers(ss).filter(function(t) {
      return t.getHandlerFunction() === 'handleEdit';
    });

    let triggerRegistered;
    if (existing.length > 0) {
      Logger.log('handleEdit トリガーは既に登録済みのためスキップ（' + existing.length + '件）');
      triggerRegistered = false;
    } else {
      ScriptApp.newTrigger('handleEdit')
        .forSpreadsheet(ss)
        .onEdit()
        .create();
      Logger.log('✅ handleEdit トリガーを新規登録しました');
      triggerRegistered = true;
    }

    ui.alert(
      '権限承認完了',
      '✅ すべての権限が正常に承認されました。\n\n' +
      (triggerRegistered
        ? '✅ handleEdit トリガーを新規登録しました。'
        : 'ℹ️ handleEdit トリガーはすでに登録済みのためスキップしました。'),
      ui.ButtonSet.OK
    );
    Logger.log('✅ authorizeScript 完了');

  } catch (error) {
    ui.alert('エラー', '❌ 権限承認中にエラーが発生しました:\n' + error.toString(), ui.ButtonSet.OK);
    Logger.log('❌ authorizeScript エラー: ' + error.toString());
  }
}


/**
 * 【ポリシー取得ボタン】
 *
 * 図形ボタンから呼び出す関数
 * スクリプト割り当て: menuGetPolicies
 */
function menuGetPolicies() {
  try {
    const spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
    const ui = SpreadsheetApp.getUi();

    // 確認ダイアログ
    const response = ui.alert(
      'ポリシー取得',
      'eBayからポリシーを取得してシートを更新します。\n既存のデータは上書きされます（操作列とプルダウンは保持）。\n\n実行しますか？',
      ui.ButtonSet.OK_CANCEL
    );

    if (response !== ui.Button.OK) {
      Logger.log('キャンセルされました');
      return;
    }

    // ライブラリ経由でポリシー取得実行
    const result = EbayLib.exportPoliciesToSheet(spreadsheetId);

    // 完了メッセージ
    ui.alert(
      '取得完了',
      '✅ ポリシーを取得しました\n\n' +
      '- Fulfillment Policy: ' + result.fulfillmentCount + '件\n' +
      '- Return Policy: ' + result.returnCount + '件\n' +
      '- Payment Policy: ' + result.paymentCount + '件\n' +
      '合計: ' + result.totalCount + '件',
      ui.ButtonSet.OK
    );

  } catch (error) {
    const ui = SpreadsheetApp.getUi();
    ui.alert('エラー', '❌ ' + error.toString(), ui.ButtonSet.OK);
    Logger.log('❌ ポリシー取得エラー: ' + error.toString());
  }
}

/**
 * 【ポリシー更新ボタン】
 *
 * 図形ボタンから呼び出す関数
 * スクリプト割り当て: menuSyncPolicies
 */
function menuSyncPolicies() {
  try {
    const spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
    const ui = SpreadsheetApp.getUi();

    // 確認ダイアログ
    const response = ui.alert(
      'ポリシー更新',
      'シートの変更をeBayに反映します。\n\n' +
      '- 操作列が「追加」→ 新規作成\n' +
      '- 操作列が「更新」→ 更新\n' +
      '- 操作列が「削除」→ 削除\n' +
      '- 操作列が「-」または空欄 → スキップ\n\n' +
      '実行しますか？',
      ui.ButtonSet.OK_CANCEL
    );

    if (response !== ui.Button.OK) {
      Logger.log('キャンセルされました');
      return;
    }

    // ライブラリ経由でポリシー同期実行
    const result = EbayLib.syncPoliciesToEbay(spreadsheetId);

    // 完了メッセージ
    let message = '✅ 同期が完了しました\n\n';
    message += '作成: ' + result.created.length + '件\n';
    message += '更新: ' + result.updated.length + '件\n';
    message += '削除: ' + result.deleted.length + '件\n';
    message += 'スキップ: ' + result.skipped + '件\n';

    if (result.errors.length > 0) {
      message += '\n⚠️ エラー: ' + result.errors.length + '件\n';
      message += '詳細はログを確認してください';
    }

    ui.alert('同期完了', message, ui.ButtonSet.OK);

  } catch (error) {
    const ui = SpreadsheetApp.getUi();
    ui.alert('エラー', '❌ ' + error.toString(), ui.ButtonSet.OK);
    Logger.log('❌ ポリシー同期エラー: ' + error.toString());
  }
}

/**
 * 【初回セットアップ】
 * 図形ボタンに割り当てる関数
 *
 * 前提条件: EbayLib ライブラリが追加済みであること
 */
function setupEbayManager() {
  try {
    const spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
    const ui = SpreadsheetApp.getUi();

    Logger.log('=== eBay出品管理 - 初回セットアップ開始 ===');
    Logger.log('スプレッドシートID: ' + spreadsheetId);

    // ライブラリ経由でセットアップ実行
    const result = EbayLib.setupEbayManager(spreadsheetId);

    // 結果をUIに表示
    if (result.success) {
      ui.alert(
        'セットアップ完了',
        result.message,
        ui.ButtonSet.OK
      );
      Logger.log('✅ セットアップ完了');
    } else {
      ui.alert(
        'エラー',
        result.message,
        ui.ButtonSet.OK
      );
      Logger.log('❌ セットアップ失敗: ' + result.error);
    }

    Logger.log('=== 初回セットアップ処理完了 ===');
    return result;

  } catch (error) {
    const ui = SpreadsheetApp.getUi();
    ui.alert('エラー', '❌ ' + error.toString(), ui.ButtonSet.OK);
    Logger.log('❌ セットアップエラー: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * ツール設定シートからライブラリ情報を取得
 *
 * ライブラリが未追加の場合に使用
 *
 * @returns {Object} { scriptId: string, identifier: string }
 */
function getLibraryInfoFromSheet() {
  try {
    const settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ツール設定');

    if (!settingsSheet) {
      // フォールバック
      return {
        scriptId: '13B_QVLCmt-KuxsyytDsS-2Ca6S_PLyNb-ZlEVbpg0T5-vEvM3otTLn1Y',
        identifier: 'EbayLib'
      };
    }

    const data = settingsSheet.getDataRange().getValues();

    for (let i = 0; i < data.length; i++) {
      const key = data[i][0];
      const value = data[i][1];

      if (key === 'ライブラリスクリプトID') {
        return {
          scriptId: value,
          identifier: 'EbayLib'
        };
      }
    }

    // 見つからない場合はフォールバック
    return {
      scriptId: '13B_QVLCmt-KuxsyytDsS-2Ca6S_PLyNb-ZlEVbpg0T5-vEvM3otTLn1Y',
      identifier: 'EbayLib'
    };

  } catch (e) {
    Logger.log('⚠️ ライブラリ情報取得エラー: ' + e.toString());
    return {
      scriptId: '13B_QVLCmt-KuxsyytDsS-2Ca6S_PLyNb-ZlEVbpg0T5-vEvM3otTLn1Y',
      identifier: 'EbayLib'
    };
  }
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// OAuth設定テスト関数（Apps Scriptエディタから実行）
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * 【テスト0】テストガイド表示
 */
function showOAuthTestGuide() {
  const spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  EbayLib.showTestGuide();
}

/**
 * 【テスト1】OAuth設定確認
 */
function testCheckOAuthSettings() {
  const spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  return EbayLib.testCheckOAuthSettings(spreadsheetId);
}

/**
 * 【テスト2】OAuth認証URL生成
 */
function testGenerateAuthUrl() {
  const spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  return EbayLib.testGenerateAuthUrl(spreadsheetId);
}

/**
 * 【テスト3】トークン取得
 *
 * 使い方: testExchangeTokens("ここにコピーしたAuthorization Codeを貼り付け")
 */
function testExchangeTokens(authorizationCode) {
  const spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  return EbayLib.testExchangeTokens(spreadsheetId, authorizationCode);
}

/**
 * 【テスト4】トークン自動更新テスト
 */
function testAutoRefresh() {
  const spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  return EbayLib.testAutoRefresh(spreadsheetId);
}

/**
 * 【テスト5】ポリシー取得（統合テスト）
 */
function testGetPolicies() {
  const spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  return EbayLib.testGetPolicies(spreadsheetId);
}
