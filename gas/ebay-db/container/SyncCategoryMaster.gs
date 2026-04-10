/**
 * SyncCategoryMaster.gs
 * このブック（category_master原本）から参照用ブックへカテゴリマスタを転記
 *
 * ソース: このスクリプトがバインドされているスプレッドシート（category_master原本）
 * ターゲット: スクリプトプロパティ SERVICE_BOOK_ID で指定された参照用ブック
 */

const SYNC_SHEET_NAMES = [
  'category_master_EBAY_US',
  'category_master_EBAY_GB',
  'category_master_EBAY_DE',
  'category_master_EBAY_AU',
  'category_master_EBAY_JP',
  'condition_ja_map'
];

/**
 * category_master原本から参照用ブックへ全シートを転記
 */
function syncCategoryMasterSheets() {
  const targetId = PropertiesService.getScriptProperties().getProperty('SERVICE_BOOK_ID');
  if (!targetId) {
    throw new Error('スクリプトプロパティ SERVICE_BOOK_ID が未設定です');
  }

  const sourceSs = SpreadsheetApp.getActiveSpreadsheet();
  const targetSs = SpreadsheetApp.openById(targetId);

  const results = [];

  SYNC_SHEET_NAMES.forEach(function(sheetName) {
    const sourceSheet = sourceSs.getSheetByName(sheetName);

    if (!sourceSheet) {
      Logger.log('スキップ: ' + sheetName + ' がソースブックに存在しません');
      results.push(sheetName + ': ソースなし（スキップ）');
      return;
    }

    const lastRow = sourceSheet.getLastRow();
    const lastCol = sourceSheet.getLastColumn();

    if (lastRow === 0) {
      Logger.log('スキップ: ' + sheetName + ' にデータがありません');
      results.push(sheetName + ': データなし（スキップ）');
      return;
    }

    const data = sourceSheet.getRange(1, 1, lastRow, lastCol).getValues();

    let targetSheet = targetSs.getSheetByName(sheetName);
    if (!targetSheet) {
      targetSheet = targetSs.insertSheet(sheetName);
    } else {
      targetSheet.clearContents();
    }

    targetSheet.getRange(1, 1, data.length, data[0].length).setValues(data);

    const rowCount = lastRow - 1;
    Logger.log('転記完了: ' + sheetName + ' (' + rowCount + '件)');
    results.push(sheetName + ': ' + rowCount + '件');
  });

  Logger.log('=== 転記完了 ===\n' + results.join('\n'));
}
