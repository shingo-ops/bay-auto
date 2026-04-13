/**
 * シート双方向同期機能（スタンドアロン）
 * container/Code.gs の ⚙️メニューから呼び出す
 */

// 同期対象シート名リスト（出品シートは絶対に含めない）
const SYNC_TARGET_SHEETS = [
  'Vero/禁止ワード',
  '状態_テンプレ',
  'Description_テンプレ',
  '担当者管理',
  'ポリシー管理',
  'ツール設定',
  'HARU_CSV',
  'セルスタ_CSV',
  'プルダウン管理'
];

// 絶対に触ってはいけないシート
const PROTECTED_SHEETS = ['出品'];

/**
 * シート同期実行
 * @param {string} spreadsheetId 出品スプレッドシートID
 * @param {string} sheetName 同期するシート名
 * @param {string} direction 'ss_to_db' または 'db_to_ss'
 * @returns {{ success: boolean, message: string }}
 */
function syncSheet(spreadsheetId, sheetName, direction) {
  try {
    Logger.log('=== シート同期開始: ' + sheetName + ' / 方向: ' + direction + ' ===');

    // 保護シートチェック
    if (PROTECTED_SHEETS.indexOf(sheetName) !== -1) {
      return { success: false, message: '「' + sheetName + '」は同期対象外のシートです。' };
    }

    // 対象シートチェック
    if (SYNC_TARGET_SHEETS.indexOf(sheetName) === -1) {
      return { success: false, message: '「' + sheetName + '」は同期対象のシートではありません。' };
    }

    if (spreadsheetId) CURRENT_SPREADSHEET_ID = spreadsheetId;

    // 出品DBのIDを取得
    const config = getEbayConfig();
    const outputDbId = config.outputDbSpreadsheetId;
    if (!outputDbId) {
      return { success: false, message: '出品DBが設定されていません。ツール設定を確認してください。' };
    }

    // 出品SS・出品DBを開く
    const sourceSS = direction === 'ss_to_db'
      ? getTargetSpreadsheet(spreadsheetId)
      : SpreadsheetApp.openById(outputDbId);
    const destSS   = direction === 'ss_to_db'
      ? SpreadsheetApp.openById(outputDbId)
      : getTargetSpreadsheet(spreadsheetId);

    const sourceLabel = direction === 'ss_to_db' ? '出品スプレッドシート' : '出品DB';
    const destLabel   = direction === 'ss_to_db' ? '出品DB' : '出品スプレッドシート';

    // コピー元シートを取得
    const sourceSheet = sourceSS.getSheetByName(sheetName);
    if (!sourceSheet) {
      return { success: false, message: sourceLabel + 'に「' + sheetName + '」シートが見つかりません。' };
    }

    // コピー先に既存シートがあれば削除（copyTo後にリネームするため先に除去）
    const existingDestSheet = destSS.getSheetByName(sheetName);
    if (existingDestSheet) {
      // シートが1枚しかない場合は削除できないためダミーシートを挿入
      if (destSS.getSheets().length === 1) {
        destSS.insertSheet('__temp__');
      }
      destSS.deleteSheet(existingDestSheet);
      Logger.log('既存の「' + sheetName + '」シートを削除しました');
    }

    // copyTo() で完全コピー（列幅・行高さ・書式・数式・プルダウン・結合すべて引き継ぎ）
    const copiedSheet = sourceSheet.copyTo(destSS);
    copiedSheet.setName(sheetName);
    Logger.log('✅ copyTo() 完了: 「' + copiedSheet.getName() + '」');

    // ダミーシートが残っていれば削除
    const tempSheet = destSS.getSheetByName('__temp__');
    if (tempSheet) destSS.deleteSheet(tempSheet);

    const lastRow = sourceSheet.getLastRow();
    const lastCol = sourceSheet.getLastColumn();

    return {
      success: true,
      message: '「' + sheetName + '」を ' + sourceLabel + ' → ' + destLabel + ' に同期しました。\n' +
               '列幅・行高さ・書式・数式をすべて引き継ぎました。\n（' + lastRow + '行 × ' + lastCol + '列）'
    };

  } catch (e) {
    Logger.log('❌ シート同期エラー: ' + e.toString());
    return { success: false, message: 'シート同期エラー: ' + e.toString() };
  } finally {
    CURRENT_SPREADSHEET_ID = null;
  }
}

/**
 * 同期対象シートリストを返す（container側のHTML生成用）
 */
function getSyncTargetSheets() {
  return SYNC_TARGET_SHEETS;
}
