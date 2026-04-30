/**
 * ConditionDropdown.gs
 *
 * eBay Condition のドロップダウン管理機能
 * category_master.conditions_json で対象 condition_id リストを取得し、
 * condition_ja_map.ja_display を日本語表示名としてD列プルダウンに表示します。
 * ユーザー選択後は condition_id（数値）をE列に自動入力し、eBay API 送信に利用します。
 *
 * シート列構成（定数で変更可）:
 *   C列（CATEGORY_COLUMN）     : カテゴリID
 *   D列（CONDITION_COLUMN）    : Condition 表示名（ja_display）
 *   E列（CONDITION_ID_COLUMN） : condition_id（eBay API 送信用、自動入力）
 */

const CATEGORY_COLUMN     = 3; // C列: カテゴリID
const CONDITION_COLUMN    = 4; // D列: Condition 表示名（ja_display）
const CONDITION_ID_COLUMN = 5; // E列: condition_id（eBay API 送信用）

/**
 * カテゴリIDに対応する condition_id リストを category_master から取得
 *
 * @param {Spreadsheet} ss - スプレッドシートオブジェクト
 * @param {string} categoryId - eBay カテゴリID
 * @returns {Array<string>} condition_id の文字列配列（例: ["1000", "3000"]）
 */
function getConditionIdsByCategoryId(ss, categoryId) {
  if (!ss || !categoryId) {
    Logger.log('getConditionIdsByCategoryId: 無効な引数');
    return [];
  }

  try {
    const sheet = ss.getSheetByName('category_master');
    if (!sheet) {
      Logger.log('category_master シートが見つかりません');
      return [];
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const categoryIdIndex    = headers.indexOf('category_id');
    const conditionsJsonIndex = headers.indexOf('conditions_json');

    if (categoryIdIndex === -1 || conditionsJsonIndex === -1) {
      Logger.log('必要な列が見つかりません: category_id または conditions_json');
      return [];
    }

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][categoryIdIndex]) === String(categoryId)) {
        const conditionsJson = data[i][conditionsJsonIndex];
        if (!conditionsJson) return [];
        try {
          const parsed = JSON.parse(conditionsJson);
          // [{id: "1000", ...}] 形式と [1000, 3000] 形式の両方に対応
          return parsed.map(item =>
            (typeof item === 'object' && item !== null) ? String(item.id) : String(item)
          );
        } catch (e) {
          Logger.log('conditions_json パースエラー: ' + e.toString());
          return [];
        }
      }
    }

    Logger.log('カテゴリIDが見つかりません: ' + categoryId);
    return [];
  } catch (error) {
    Logger.log('getConditionIdsByCategoryId エラー: ' + error.toString());
    return [];
  }
}

/**
 * condition_ja_map シートを参照し、condition_id に対応する ja_display リストを構築
 * 順序は conditionIds の並び順を維持します
 *
 * @param {Spreadsheet} ss - スプレッドシートオブジェクト
 * @param {Array<string>} conditionIds - 対象の condition_id リスト
 * @returns {Array<{id: string, jaDisplay: string}>}
 */
function buildConditionDisplayList(ss, conditionIds) {
  const sheet = ss.getSheetByName('condition_ja_map');
  if (!sheet) {
    Logger.log('condition_ja_map シートが見つかりません。condition_id をそのまま表示します');
    return conditionIds.map(id => ({ id: id, jaDisplay: id }));
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIdx        = headers.indexOf('condition_id');
  const jaDisplayIdx = headers.indexOf('ja_display');

  if (idIdx === -1 || jaDisplayIdx === -1) {
    Logger.log('必要な列が見つかりません: condition_id または ja_display');
    return conditionIds.map(id => ({ id: id, jaDisplay: id }));
  }

  // condition_id -> ja_display の逆引きマップを構築
  const jaMap = {};
  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][idIdx]);
    if (id) jaMap[id] = data[i][jaDisplayIdx] || id;
  }

  return conditionIds.map(id => ({
    id: id,
    jaDisplay: jaMap[id] || id // 未登録は id をそのまま表示
  }));
}

/**
 * ja_display から Condition 情報を逆引き（eBay API 送信時に呼び出す）
 *
 * condition_ja_map シートのヘッダー:
 *   condition_id   : eBay の数値 ID（例: 3000）
 *   condition_name : eBay API に送る英語名（例: "Used"）
 *   condition_enum : eBay API に送る列挙値（例: "USED_EXCELLENT"）
 *   ja_display     : プルダウン表示用の日本語名（例: "やや傷や汚れあり"）
 *
 * @param {string} jaDisplay - D列で選択された日本語表示名
 * @returns {{condition_id: number, condition_name: string, condition_enum: string}|null}
 */
function getConditionByJaDisplay(jaDisplay) {
  if (!jaDisplay) return null;

  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('condition_ja_map');
    if (!sheet) {
      Logger.log('condition_ja_map シートが見つかりません');
      return null;
    }

    const data    = sheet.getDataRange().getValues();
    const headers = data[0];
    const idIdx        = headers.indexOf('condition_id');
    const nameIdx      = headers.indexOf('condition_name');
    const enumIdx      = headers.indexOf('condition_enum');
    const jaDisplayIdx = headers.indexOf('ja_display');

    if (idIdx === -1 || jaDisplayIdx === -1) {
      Logger.log('必要な列が見つかりません: condition_id または ja_display');
      return null;
    }

    for (let i = 1; i < data.length; i++) {
      if (data[i][jaDisplayIdx] === jaDisplay) {
        return {
          condition_id:   Number(data[i][idIdx]),
          condition_name: nameIdx !== -1 ? String(data[i][nameIdx]) : '',
          condition_enum: enumIdx !== -1 ? String(data[i][enumIdx]) : ''
        };
      }
    }

    Logger.log('ja_display に対応するエントリが見つかりません: ' + jaDisplay);
    return null;
  } catch (error) {
    Logger.log('getConditionByJaDisplay エラー: ' + error.toString());
    return null;
  }
}

/**
 * C列（カテゴリID）変更時の処理
 * D列のドロップダウン（ja_display）を更新し、D/E列をクリアします
 *
 * @param {Spreadsheet} ss
 * @param {Sheet} sheet
 * @param {number} row
 * @param {*} categoryId
 */
function handleCategoryChange(ss, sheet, row, categoryId) {
  const conditionCell   = sheet.getRange(row, CONDITION_COLUMN);
  const conditionIdCell = sheet.getRange(row, CONDITION_ID_COLUMN);

  if (!categoryId) {
    conditionCell.clearDataValidations();
    conditionCell.clearContent();
    conditionIdCell.clearContent();
    return;
  }

  const conditionIds = getConditionIdsByCategoryId(ss, String(categoryId));
  if (conditionIds.length === 0) {
    Logger.log('カテゴリID ' + categoryId + ' に対応する Condition が見つかりません');
    return;
  }

  const displayList  = buildConditionDisplayList(ss, conditionIds);
  const displayNames = displayList.map(c => c.jaDisplay);

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(displayNames, true)
    .setAllowInvalid(false)
    .build();
  conditionCell.setDataValidation(rule);

  // 既存値が新リストにない場合は D/E 列をクリア
  const currentValue = conditionCell.getValue();
  if (currentValue && !displayNames.includes(currentValue)) {
    conditionCell.clearContent();
    conditionIdCell.clearContent();
  }

  Logger.log('カテゴリID ' + categoryId + ' の Condition ドロップダウンを更新しました（' + displayNames.length + '件）');
}

/**
 * D列（Condition 表示名）変更時の処理
 * 選択された ja_display から condition_id を解決して E列に数値で書き込みます
 *
 * @param {Sheet} sheet
 * @param {number} row
 * @param {string} jaDisplay - 選択された表示名
 */
function handleConditionChange(sheet, row, jaDisplay) {
  const conditionIdCell = sheet.getRange(row, CONDITION_ID_COLUMN);

  if (!jaDisplay) {
    conditionIdCell.clearContent();
    return;
  }

  const cond = getConditionByJaDisplay(jaDisplay);
  if (cond) {
    conditionIdCell.setValue(cond.condition_id);
    Logger.log('condition_id を設定しました: ' + cond.condition_id + ' (' + cond.condition_enum + ')');
  } else {
    conditionIdCell.clearContent();
    Logger.log('ja_display に対応する condition_id が見つかりません: ' + jaDisplay);
  }
}

/**
 * テスト用: カテゴリIDから Condition 表示リストを確認
 */
function testGetConditionsByCategoryId() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const conditionIds = getConditionIdsByCategoryId(ss, '261581');
  Logger.log('カテゴリID 261581 の condition_ids: ' + JSON.stringify(conditionIds));
  const displayList = buildConditionDisplayList(ss, conditionIds);
  Logger.log('表示リスト: ' + JSON.stringify(displayList));
}

/**
 * テスト用: ja_display から Condition 情報を逆引き確認
 */
function testGetConditionByJaDisplay() {
  const cond = getConditionByJaDisplay('やや傷や汚れあり');
  Logger.log('逆引き結果: ' + JSON.stringify(cond));
  // 期待値例: { condition_id: 3000, condition_name: "Used", condition_enum: "USED_GOOD" }
}

/**
 * 状態テンプレ列変更時：テンプレートシートから状態説明テキストを取得してセット
 * @param {string} spreadsheetId
 * @param {string} sheetName
 * @param {number} row
 * @param {string} selectedCondition
 */
function handleConditionTemplateChange(spreadsheetId, sheetName, row, selectedCondition) {
  try {
    Logger.log('状態テンプレ変更: ' + selectedCondition);

    const ss = getTargetSpreadsheet(spreadsheetId);
    const templateSheet = ss.getSheetByName('状態_テンプレ');
    if (!templateSheet) {
      Logger.log('⚠️ 状態_テンプレシートが見つかりません');
      return;
    }

    const lastCol = templateSheet.getLastColumn();
    const lastRow = templateSheet.getLastRow();
    if (lastRow < 2) return;

    const headers = templateSheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const conditionIdx = headers.findIndex(function(h) {
      return String(h || '').trim() === 'コンディション';
    });
    const templateIdx = headers.findIndex(function(h) {
      return String(h || '').trim() === 'テンプレート(英語)';
    });

    if (conditionIdx === -1 || templateIdx === -1) {
      Logger.log('⚠️ 状態_テンプレシートに「コンディション」または「テンプレート(英語)」列が見つかりません');
      return;
    }

    const data = templateSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    let templateText = '';
    for (let i = 0; i < data.length; i++) {
      const condition = String(data[i][conditionIdx] || '').trim();
      if (condition === selectedCondition) {
        templateText = String(data[i][templateIdx] || '').replace(/<[^>]*>/g, '').trim();
        break;
      }
    }

    if (!templateText) {
      Logger.log('⚠️ 「' + selectedCondition + '」のテンプレートが見つかりません');
      return;
    }

    const headerMapping = buildListingHeaderMapping(spreadsheetId, sheetName);
    const conditionDescCol = headerMapping['状態説明'];
    if (!conditionDescCol) {
      Logger.log('⚠️ 「状態説明」列が見つかりません');
      return;
    }

    const sheet = ss.getSheetByName(sheetName);
    sheet.getRange(row, conditionDescCol).setValue(templateText);
    Logger.log('✅ 状態説明を更新: ' + templateText.substring(0, 50) + '...');

  } catch(e) {
    Logger.log('handleConditionTemplateChange エラー: ' + e.toString());
  }
}

/**
 * 発送業者列変更時：プルダウン管理シートから発送方法リストを取得して発送方法列にプルダウン展開
 * @param {string} spreadsheetId
 * @param {string} sheetName
 * @param {number} row
 * @param {string} selectedShipper
 */
function handleShipperChange(spreadsheetId, sheetName, row, selectedShipper) {
  try {
    Logger.log('発送業者変更: ' + selectedShipper);

    const ss = getTargetSpreadsheet(spreadsheetId);
    const pulldownSheet = ss.getSheetByName('プルダウン管理');
    if (!pulldownSheet) {
      Logger.log('⚠️ プルダウン管理シートが見つかりません');
      return;
    }

    const lastCol = pulldownSheet.getLastColumn();
    const lastRow = pulldownSheet.getLastRow();
    if (lastRow < 2) return;

    const headers = pulldownSheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const shipperColIdx = headers.findIndex(function(h) {
      return String(h || '').trim() === selectedShipper;
    });

    if (shipperColIdx === -1) {
      Logger.log('⚠️ プルダウン管理シートに「' + selectedShipper + '」列が見つかりません');
      return;
    }

    const methodValues = pulldownSheet.getRange(2, shipperColIdx + 1, lastRow - 1, 1).getValues();
    const methodList = methodValues
      .map(function(r) { return String(r[0] || '').trim(); })
      .filter(function(v) { return v !== ''; });

    if (methodList.length === 0) {
      Logger.log('⚠️ 「' + selectedShipper + '」の発送方法リストが空です');
      return;
    }

    Logger.log('発送方法リスト: ' + methodList.join(', '));

    const headerMapping = buildListingHeaderMapping(spreadsheetId, sheetName);
    const shippingMethodCol = headerMapping['発送方法'];
    if (!shippingMethodCol) {
      Logger.log('⚠️ 「発送方法」列が見つかりません');
      return;
    }

    const sheet = ss.getSheetByName(sheetName);
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(methodList, true)
      .setAllowInvalid(false)
      .build();

    const methodCell = sheet.getRange(row, shippingMethodCol);
    methodCell.setDataValidation(rule);

    const currentMethod = String(methodCell.getValue() || '').trim();
    if (currentMethod && methodList.indexOf(currentMethod) === -1) {
      methodCell.clearContent();
      Logger.log('発送方法をクリア（前の値が新リストにない）: ' + currentMethod);
    }

    if (!currentMethod) {
      methodCell.setValue(methodList[0]);
    }

    Logger.log('✅ 発送方法プルダウンを更新: ' + methodList.length + '件');

  } catch(e) {
    Logger.log('handleShipperChange エラー: ' + e.toString());
  }
}
