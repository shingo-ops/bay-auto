/**
 * "状態_テンプレ"シート構造確認スクリプト
 *
 * clasp runで実行してヘッダー構造を確認
 */

/**
 * "状態_テンプレ"シートのヘッダー構造を確認
 *
 * @param {string} spreadsheetId スプレッドシートID（省略時はデフォルト）
 */
function debugConditionTemplateHeaders(spreadsheetId) {
  try {
    if (spreadsheetId) {
      CURRENT_SPREADSHEET_ID = spreadsheetId;
    }

    const ss = getTargetSpreadsheet(spreadsheetId);
    const templateSheet = ss.getSheetByName('状態_テンプレ');

    if (!templateSheet) {
      Logger.log('❌ "状態_テンプレ"シートが見つかりません');
      return {
        success: false,
        error: '"状態_テンプレ"シートが存在しません'
      };
    }

    Logger.log('=== "状態_テンプレ"シート構造 ===');
    Logger.log('');

    // シート情報
    const lastRow = templateSheet.getLastRow();
    const lastCol = templateSheet.getLastColumn();

    Logger.log('最終行: ' + lastRow);
    Logger.log('最終列: ' + lastCol);
    Logger.log('');

    // ヘッダー行を確認（1行目、2行目、3行目の可能性）
    for (let headerRow = 1; headerRow <= 3; headerRow++) {
      Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
      Logger.log(headerRow + '行目:');
      Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');

      const headers = templateSheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];

      for (let i = 0; i < headers.length; i++) {
        const colLetter = String.fromCharCode(65 + i); // A=65
        const headerName = headers[i];

        if (headerName) {
          Logger.log(colLetter + '列（' + (i + 1) + '列目）: "' + headerName + '"');
        } else {
          Logger.log(colLetter + '列（' + (i + 1) + '列目）: （空）');
        }
      }
      Logger.log('');
    }

    // データ行のサンプルを確認（4-8行目）
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('データサンプル（4-8行目）:');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('');

    if (lastRow >= 4) {
      const sampleRows = Math.min(8, lastRow);
      const sampleData = templateSheet.getRange(4, 1, sampleRows - 3, lastCol).getValues();

      for (let i = 0; i < sampleData.length; i++) {
        const rowNum = i + 4;
        Logger.log('--- ' + rowNum + '行目 ---');

        for (let j = 0; j < sampleData[i].length; j++) {
          const colLetter = String.fromCharCode(65 + j);
          const cellValue = sampleData[i][j];

          if (cellValue) {
            // 長い文字列は省略
            const displayValue = String(cellValue).length > 50
              ? String(cellValue).substring(0, 50) + '...'
              : String(cellValue);

            Logger.log(colLetter + '列: ' + displayValue);
          }
        }
        Logger.log('');
      }
    } else {
      Logger.log('⚠️ データ行が存在しません（ヘッダーのみ）');
    }

    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('✅ デバッグ完了');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');

    return {
      success: true,
      lastRow: lastRow,
      lastCol: lastCol
    };

  } catch (error) {
    Logger.log('❌ デバッグエラー: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  } finally {
    CURRENT_SPREADSHEET_ID = null;
  }
}

/**
 * "出品"シートのヘッダーも確認（状態テンプレ、状態説明列を探す）
 *
 * @param {string} spreadsheetId スプレッドシートID（省略時はデフォルト）
 */
function debugListingSheetHeaders(spreadsheetId) {
  try {
    if (spreadsheetId) {
      CURRENT_SPREADSHEET_ID = spreadsheetId;
    }

    const ss = getTargetSpreadsheet(spreadsheetId);
    const listingSheet = ss.getSheetByName('出品');

    if (!listingSheet) {
      Logger.log('❌ "出品"シートが見つかりません');
      return {
        success: false,
        error: '"出品"シートが存在しません'
      };
    }

    Logger.log('=== "出品"シート構造 ===');
    Logger.log('');

    // ヘッダー行（3行目）
    const headerRow = 3;
    const lastCol = listingSheet.getLastColumn();
    const headers = listingSheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];

    Logger.log('ヘッダー行: ' + headerRow + '行目');
    Logger.log('最終列: ' + lastCol);
    Logger.log('');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('全ヘッダー:');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');

    const targetHeaders = ['状態テンプレ', '状態説明', 'タイトル', '文字数'];

    for (let i = 0; i < headers.length; i++) {
      const colLetter = String.fromCharCode(65 + (i % 26)); // 簡易的な列名
      const headerName = headers[i];

      if (headerName) {
        // 対象ヘッダーの場合は目立たせる
        if (targetHeaders.indexOf(headerName) !== -1) {
          Logger.log('★ 列' + (i + 1) + ': "' + headerName + '"');
        } else {
          Logger.log('  列' + (i + 1) + ': "' + headerName + '"');
        }
      }
    }

    Logger.log('');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('対象列の確認:');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');

    for (let j = 0; j < targetHeaders.length; j++) {
      const targetHeader = targetHeaders[j];
      let found = false;

      for (let i = 0; i < headers.length; i++) {
        if (headers[i] === targetHeader) {
          Logger.log('✅ "' + targetHeader + '": 列' + (i + 1));
          found = true;
          break;
        }
      }

      if (!found) {
        Logger.log('❌ "' + targetHeader + '": 見つかりません');
      }
    }

    Logger.log('');
    Logger.log('✅ デバッグ完了');

    return {
      success: true,
      lastCol: lastCol
    };

  } catch (error) {
    Logger.log('❌ デバッグエラー: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  } finally {
    CURRENT_SPREADSHEET_ID = null;
  }
}
