/**
 * ヘッダー構造確認ツール（clasp run対応）
 *
 * 使い方:
 * clasp run getListingHeaders -p '[{"spreadsheetId":"YOUR_SPREADSHEET_ID"}]'
 */

/**
 * "出品"シートの3行目ヘッダーを取得
 *
 * @param {string} spreadsheetId スプレッドシートID
 * @returns {Object} { success: boolean, headers: Array, mapping: Object }
 */
function getListingHeaders(spreadsheetId) {
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

    const headerRow = 1;
    const lastCol = listingSheet.getLastColumn();
    const headers = listingSheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];

    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('📋 "出品"シート - 1行目ヘッダー構造');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('');
    Logger.log('スプレッドシートID: ' + spreadsheetId);
    Logger.log('最終列: ' + lastCol);
    Logger.log('');

    const mapping = {};
    const headerList = [];

    for (let i = 0; i < headers.length; i++) {
      const headerName = headers[i];
      const colNum = i + 1;

      if (headerName) {
        mapping[headerName] = colNum;
        headerList.push({
          column: colNum,
          name: headerName
        });

        Logger.log('列' + colNum + ': "' + headerName + '"');
      } else {
        Logger.log('列' + colNum + ': （空）');
      }
    }

    Logger.log('');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('🔍 重要な列の確認');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('');

    const importantHeaders = [
      'SKU',
      'タイトル',
      '文字数',
      '状態テンプレ',
      '状態説明',
      'Description',
      'カテゴリID',
      '個数',
      '売値($)',
      'Shipping Policy',
      'Return Policy',
      'Payment Policy'
    ];

    for (let i = 0; i < importantHeaders.length; i++) {
      const targetHeader = importantHeaders[i];
      const colNum = mapping[targetHeader];

      if (colNum) {
        Logger.log('✅ "' + targetHeader + '": 列' + colNum);
      } else {
        Logger.log('❌ "' + targetHeader + '": 見つかりません');
      }
    }

    Logger.log('');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('✅ ヘッダー取得完了');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');

    return {
      success: true,
      headers: headerList,
      mapping: mapping,
      lastColumn: lastCol
    };

  } catch (error) {
    Logger.log('❌ エラー: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  } finally {
    CURRENT_SPREADSHEET_ID = null;
  }
}

/**
 * "状態_テンプレ"シートのヘッダーを取得
 *
 * @param {string} spreadsheetId スプレッドシートID
 * @param {number} headerRow ヘッダー行番号（省略時は1, 2, 3行目を全て確認）
 * @returns {Object} { success: boolean, headers: Array, mapping: Object }
 */
function getConditionTemplateHeaders(spreadsheetId, headerRow) {
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

    const lastCol = templateSheet.getLastColumn();
    const lastRow = templateSheet.getLastRow();

    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('📋 "状態_テンプレ"シート - ヘッダー構造');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('');
    Logger.log('スプレッドシートID: ' + spreadsheetId);
    Logger.log('最終行: ' + lastRow);
    Logger.log('最終列: ' + lastCol);
    Logger.log('');

    // ヘッダー行が指定されている場合はその行のみ
    const rowsToCheck = headerRow ? [headerRow] : [1, 2, 3];

    let bestMapping = null;
    let bestHeaderRow = null;

    for (let i = 0; i < rowsToCheck.length; i++) {
      const row = rowsToCheck[i];

      Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
      Logger.log(row + '行目:');
      Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');

      const headers = templateSheet.getRange(row, 1, 1, lastCol).getValues()[0];
      const mapping = {};
      const headerList = [];

      for (let j = 0; j < headers.length; j++) {
        const headerName = headers[j];
        const colNum = j + 1;

        if (headerName) {
          mapping[headerName] = colNum;
          headerList.push({
            column: colNum,
            name: headerName
          });

          Logger.log('列' + colNum + ': "' + headerName + '"');
        } else {
          Logger.log('列' + colNum + ': （空）');
        }
      }

      Logger.log('');

      // "コンディション"と"テンプレート(英語)"の両方があればベストヘッダー
      if (mapping['コンディション'] && mapping['テンプレート(英語)']) {
        bestMapping = mapping;
        bestHeaderRow = row;
        Logger.log('✅ このヘッダー行が最適です（コンディション列とテンプレート(英語)列を含む）');
        Logger.log('');
      }
    }

    if (bestHeaderRow) {
      Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
      Logger.log('🔍 推奨ヘッダー行: ' + bestHeaderRow + '行目');
      Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
      Logger.log('');
      Logger.log('✅ "コンディション": 列' + bestMapping['コンディション']);
      Logger.log('✅ "テンプレート(英語)": 列' + bestMapping['テンプレート(英語)']);
      Logger.log('');

      // データサンプルを表示
      Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
      Logger.log('📝 データサンプル（先頭5件）');
      Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
      Logger.log('');

      const dataStartRow = bestHeaderRow + 1;
      const sampleCount = Math.min(5, lastRow - bestHeaderRow);

      if (sampleCount > 0) {
        const sampleData = templateSheet.getRange(dataStartRow, 1, sampleCount, lastCol).getValues();

        for (let i = 0; i < sampleData.length; i++) {
          const rowNum = dataStartRow + i;
          const conditionValue = sampleData[i][bestMapping['コンディション'] - 1];
          const templateValue = sampleData[i][bestMapping['テンプレート(英語)'] - 1];

          Logger.log('【' + rowNum + '行目】');
          Logger.log('コンディション: ' + conditionValue);

          if (templateValue) {
            const preview = String(templateValue).length > 60
              ? String(templateValue).substring(0, 60) + '...'
              : String(templateValue);
            Logger.log('テンプレート: ' + preview);
          } else {
            Logger.log('テンプレート: （空）');
          }
          Logger.log('');
        }
      }
    } else {
      Logger.log('⚠️ "コンディション"列または"テンプレート(英語)"列が見つかりません');
    }

    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('✅ ヘッダー取得完了');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');

    return {
      success: true,
      bestHeaderRow: bestHeaderRow,
      mapping: bestMapping,
      lastColumn: lastCol,
      lastRow: lastRow
    };

  } catch (error) {
    Logger.log('❌ エラー: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  } finally {
    CURRENT_SPREADSHEET_ID = null;
  }
}

/**
 * 両方のシートのヘッダーをまとめて確認
 *
 * @param {string} spreadsheetId スプレッドシートID
 * @returns {Object} 結果
 */
function getAllHeaders(spreadsheetId) {
  Logger.log('');
  Logger.log('╔═══════════════════════════════════════════╗');
  Logger.log('║   ヘッダー構造 - 一括確認                ║');
  Logger.log('╚═══════════════════════════════════════════╝');
  Logger.log('');

  const listingResult = getListingHeaders(spreadsheetId);

  Logger.log('');
  Logger.log('');

  const templateResult = getConditionTemplateHeaders(spreadsheetId);

  Logger.log('');
  Logger.log('');
  Logger.log('╔═══════════════════════════════════════════╗');
  Logger.log('║   確認完了                                ║');
  Logger.log('╚═══════════════════════════════════════════╝');

  return {
    listing: listingResult,
    template: templateResult
  };
}
