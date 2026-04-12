/**
 * eBay出品管理 - バインドスクリプト
 *
 * このスクリプトをバインドスクリプトに追加して使用します
 *
 * 図形ボタンに割り当てる関数:
 * - menuSetupEbayManager: 初回セットアップ実行
 * - menuGetPolicies: ポリシー取得
 * - menuSyncPolicies: ポリシー更新
 *
 * 前提条件: EbayLib ライブラリが追加済みであること
 */

/**
 * 【初回セットアップ】
 * 図形ボタンに割り当てる関数
 *
 * 前提条件: EbayLib ライブラリが追加済みであること
 */
function menuSetupEbayManager() {
  try {
    const spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
    const ui = SpreadsheetApp.getUi();

    // ライブラリ経由でセットアップ実行
    const result = EbayLib.setupEbayManager(spreadsheetId);

    // 結果をUIに表示
    if (result.success) {
      ui.alert('セットアップ完了', result.message, ui.ButtonSet.OK);
    } else {
      ui.alert('エラー', result.message, ui.ButtonSet.OK);
    }

  } catch (error) {
    SpreadsheetApp.getUi().alert('エラー', '❌ ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
    Logger.log('❌ セットアップエラー: ' + error.toString());
  }
}

/**
 * 【ポリシー取得】
 * 図形ボタンに割り当てる関数
 *
 * 前提条件: EbayLib ライブラリが追加済みであること
 */
function menuGetPolicies() {
  try {
    const spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
    const ui = SpreadsheetApp.getUi();

    const response = ui.alert(
      'ポリシー取得',
      'eBayからポリシーを取得してシートを更新します。\n既存のデータは上書きされます（操作列とプルダウンは保持）。\n\n実行しますか？',
      ui.ButtonSet.OK_CANCEL
    );

    if (response !== ui.Button.OK) {
      return;
    }

    const result = EbayLib.menuGetPolicies(spreadsheetId);

    if (result.success) {
      ui.alert('取得完了', result.message, ui.ButtonSet.OK);
    } else {
      ui.alert('エラー', result.message, ui.ButtonSet.OK);
    }

  } catch (error) {
    SpreadsheetApp.getUi().alert('エラー', '❌ ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 【ポリシー更新】
 * 図形ボタンに割り当てる関数
 *
 * 前提条件: EbayLib ライブラリが追加済みであること
 */
function menuSyncPolicies() {
  try {
    const spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
    const ui = SpreadsheetApp.getUi();

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
      return;
    }

    const result = EbayLib.menuSyncPolicies(spreadsheetId);

    if (result.success) {
      ui.alert('同期完了', result.message, ui.ButtonSet.OK);
    } else {
      ui.alert('エラー', result.message, ui.ButtonSet.OK);
    }

  } catch (error) {
    SpreadsheetApp.getUi().alert('エラー', '❌ ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}
