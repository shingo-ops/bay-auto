/**
 * ConditionDropdown.gs
 *
 * リサーチシートの「状態」セル（E8）に、カテゴリIDに対応する
 * eBayコンディションのプルダウンを生成します。
 *
 * データソース: カテゴリマスタスプレッドシート（ツール設定の「カテゴリマスタ」）
 *   - category_master_EBAY_US シート: condition_group 列でグループを取得
 *   - condition_ja_map シート: ja_map_json でグループ別の日本語表示名を取得
 *
 * condition_ja_map スキーマ（1グループ1行）:
 *   condition_group    : グループラベル（A/B/C...）
 *   condition_ids_json : [1000, 3000] 等
 *   ja_map_json        : {"1000":"新品、未使用","3000":"やや傷や汚れあり"} 等
 *   category_count     : 該当カテゴリ数
 *   example_categories : 代表カテゴリ名3つ
 *   last_synced
 *
 * トリガー経路:
 *   1. G8（カテゴリID）を手動編集 → handleEdit → setConditionDropdown
 *   2. B8（Item URL）入力 → fetchCategoryFromUrl がG8を自動セット → setConditionDropdown
 */

/**
 * カテゴリマスタスプレッドシートを開く
 * ツール設定の「カテゴリマスタ」に設定されたIDを使用
 *
 * @returns {Spreadsheet|null}
 */
function openCategoryMasterSs() {
  const config = getEbayConfig();
  const spreadsheetId = config.categoryMasterSpreadsheetId;

  if (!spreadsheetId) {
    Logger.log('⚠️ カテゴリマスタのスプレッドシートIDが設定されていません（ツール設定の「カテゴリマスタ」を確認）');
    return null;
  }

  try {
    return SpreadsheetApp.openById(spreadsheetId);
  } catch (e) {
    Logger.log('❌ カテゴリマスタスプレッドシートを開けません: ' + e.toString());
    return null;
  }
}

/**
 * カテゴリIDに対応する condition_group を category_master_EBAY_US から取得
 *
 * @param {Spreadsheet} categoryMasterSs
 * @param {string} categoryId
 * @returns {string|null} グループラベル（例: "A"）または null
 */
function getConditionGroupForCategory(categoryMasterSs, categoryId) {
  if (!categoryMasterSs || !categoryId) return null;

  try {
    const sheet = categoryMasterSs.getSheetByName(SHEET_NAMES.CATEGORY_MASTER);
    if (!sheet) {
      Logger.log('⚠️ ' + SHEET_NAMES.CATEGORY_MASTER + ' シートが見つかりません');
      return null;
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return null;

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const catIdIdx = headers.indexOf('category_id');
    const groupIdx = headers.indexOf('condition_group');

    if (catIdIdx === -1) {
      Logger.log('⚠️ category_id 列が見つかりません');
      return null;
    }
    if (groupIdx === -1) {
      Logger.log('⚠️ condition_group 列が見つかりません。category_master を最新版に更新してください');
      return null;
    }

    const data = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][catIdIdx]) === String(categoryId)) {
        return data[i][groupIdx] || null;
      }
    }

    Logger.log('カテゴリID ' + categoryId + ' が ' + SHEET_NAMES.CATEGORY_MASTER + ' に見つかりません');
    return null;

  } catch (error) {
    Logger.log('❌ getConditionGroupForCategory エラー: ' + error.toString());
    return null;
  }
}

/**
 * condition_group に対応する ja_map_json をパースして返す
 *
 * @param {Spreadsheet} categoryMasterSs
 * @param {string} group グループラベル（例: "A"）
 * @returns {Object|null} {conditionId: jaDisplay} 形式のオブジェクト
 */
function getJaMapForGroup(categoryMasterSs, group) {
  if (!categoryMasterSs || !group) return null;

  try {
    const sheet = categoryMasterSs.getSheetByName(SHEET_NAMES.CONDITION_JA_MAP);
    if (!sheet) {
      Logger.log('⚠️ ' + SHEET_NAMES.CONDITION_JA_MAP + ' シートが見つかりません');
      return null;
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const groupIdx  = headers.indexOf('condition_group');
    const jaMapIdx  = headers.indexOf('ja_map_json');

    if (groupIdx === -1 || jaMapIdx === -1) {
      Logger.log('⚠️ condition_group または ja_map_json 列が見つかりません。condition_ja_map を最新版に更新してください');
      return null;
    }

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][groupIdx]) === String(group)) {
        const jaMapJson = data[i][jaMapIdx];
        if (!jaMapJson) return null;
        try {
          return JSON.parse(jaMapJson);
        } catch (e) {
          Logger.log('❌ ja_map_json パースエラー（グループ' + group + '）: ' + e.toString());
          return null;
        }
      }
    }

    Logger.log('グループ ' + group + ' が ' + SHEET_NAMES.CONDITION_JA_MAP + ' に見つかりません');
    return null;

  } catch (error) {
    Logger.log('❌ getJaMapForGroup エラー: ' + error.toString());
    return null;
  }
}

/**
 * リサーチシートのE8（状態）セルに、カテゴリIDに対応する状態プルダウンを設定
 *
 * フロー:
 *   1. category_master から condition_group を取得
 *   2. condition_ja_map で ja_map_json をJSON.parse
 *   3. values（ja_display）をプルダウン選択肢に設定
 *
 * @param {string} categoryId カテゴリID（G8の値）
 * @param {Sheet} sheet リサーチシート
 */
function setConditionDropdown(categoryId, sheet) {
  const conditionCell = sheet.getRange(
    RESEARCH_ITEM_LIST.DATA_ROW,
    RESEARCH_ITEM_LIST.COLUMNS.CONDITION.col
  );

  // カテゴリIDが空 → プルダウンをクリア
  if (!categoryId || String(categoryId).trim() === '') {
    conditionCell.clearDataValidations();
    Logger.log('カテゴリIDが空のため状態プルダウンをクリアしました');
    return;
  }

  const categoryMasterSs = openCategoryMasterSs();
  if (!categoryMasterSs) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'カテゴリマスタが未設定のため状態プルダウンを生成できません。\nツール設定の「カテゴリマスタ」を確認してください。',
      '⚠️ 状態プルダウン',
      8
    );
    return;
  }

  // 1. condition_group を取得
  const group = getConditionGroupForCategory(categoryMasterSs, String(categoryId));
  if (!group) {
    conditionCell.clearDataValidations();
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'カテゴリID ' + categoryId + ' の状態グループが見つかりません',
      '⚠️ 状態プルダウン',
      5
    );
    return;
  }

  // 2. ja_map_json を取得・パース
  const jaMap = getJaMapForGroup(categoryMasterSs, group);
  if (!jaMap || Object.keys(jaMap).length === 0) {
    conditionCell.clearDataValidations();
    Logger.log('グループ ' + group + ' の ja_map_json が空です');
    return;
  }

  // 3. プルダウン選択肢 = ja_map_json の値（ja_display）
  const displayOptions = Object.values(jaMap).filter(function(v) { return v && v.trim() !== ''; });

  if (displayOptions.length === 0) {
    conditionCell.clearDataValidations();
    return;
  }

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(displayOptions, true)
    .setAllowInvalid(false)
    .build();
  conditionCell.setDataValidation(rule);

  // 既存値がリストにない場合はクリア
  const currentValue = conditionCell.getValue();
  if (currentValue && displayOptions.indexOf(String(currentValue)) === -1) {
    conditionCell.clearContent();
    Logger.log('既存の状態値がリストにないためクリアしました: ' + currentValue);
  }

  Logger.log('✅ 状態プルダウン設定: カテゴリID=' + categoryId + ' → グループ' + group + ' / ' + displayOptions.length + '件');
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '状態プルダウンを設定しました（グループ' + group + ' / ' + displayOptions.length + '件）',
    '✅ 状態',
    2
  );
}

/**
 * カテゴリIDに対応するコンディション DataValidation ルールを構築して返す
 * リサーチシート・出品シート共通で使用
 *
 * @param {string} categoryId
 * @returns {DataValidation|null} ルール（カテゴリ未登録・マスタ未設定時は null）
 */
function buildConditionValidationRule(categoryId) {
  if (!categoryId || String(categoryId).trim() === '') return null;

  const categoryMasterSs = openCategoryMasterSs();
  if (!categoryMasterSs) return null;

  const group = getConditionGroupForCategory(categoryMasterSs, String(categoryId));
  if (!group) {
    Logger.log('[buildConditionValidationRule] グループ未登録: ' + categoryId);
    return null;
  }

  const jaMap = getJaMapForGroup(categoryMasterSs, group);
  if (!jaMap || Object.keys(jaMap).length === 0) return null;

  const displayOptions = Object.values(jaMap).filter(function(v) { return v && v.trim() !== ''; });
  if (displayOptions.length === 0) return null;

  Logger.log('[buildConditionValidationRule] カテゴリ=' + categoryId + ' グループ=' + group + ' 選択肢=' + displayOptions.length + '件');
  return SpreadsheetApp.newDataValidation()
    .requireValueInList(displayOptions, true)
    .setAllowInvalid(false)
    .build();
}

/**
 * ja_display から condition_id を逆引き（出品データ転記時に呼び出す）
 *
 * カテゴリIDからcondition_groupを特定し、
 * ja_map_json（{conditionId: jaDisplay}）を逆引きしてcondition_idを返す。
 *
 * @param {string} categoryId カテゴリID（G8の値）
 * @param {string} jaDisplay  選択された日本語表示名（E8の値）
 * @returns {{condition_id: number}|null}
 */
function getConditionIdByJaDisplay(categoryId, jaDisplay) {
  if (!categoryId || !jaDisplay) return null;

  const categoryMasterSs = openCategoryMasterSs();
  if (!categoryMasterSs) return null;

  const group = getConditionGroupForCategory(categoryMasterSs, String(categoryId));
  if (!group) return null;

  const jaMap = getJaMapForGroup(categoryMasterSs, group);
  if (!jaMap) return null;

  // 逆引き: jaDisplay に一致する conditionId を探す
  const keys = Object.keys(jaMap);
  for (let i = 0; i < keys.length; i++) {
    if (jaMap[keys[i]] === jaDisplay) {
      const condId = keys[i];
      return {
        condition_id: condId.match(/^\d+$/) ? parseInt(condId, 10) : condId
      };
    }
  }

  Logger.log('ja_display に対応する condition_id が見つかりません: ' + jaDisplay);
  return null;
}

/**
 * condition_ja_map の全グループの ja_map_json を確認用に返す（clasp run 用）
 */
function inspectConditionJaMap() {
  const categoryMasterSs = openCategoryMasterSs();
  if (!categoryMasterSs) return 'カテゴリマスタが開けません';

  const sheet = categoryMasterSs.getSheetByName(SHEET_NAMES.CONDITION_JA_MAP);
  if (!sheet) return SHEET_NAMES.CONDITION_JA_MAP + ' シートが見つかりません';

  const data    = sheet.getDataRange().getValues();
  const headers = data[0];
  const groupIdx  = headers.indexOf('condition_group');
  const jaMapIdx  = headers.indexOf('ja_map_json');

  const result = [];
  for (let i = 1; i < data.length; i++) {
    const group = String(data[i][groupIdx] || '');
    if (!group) continue;
    let jaMap = {};
    try { jaMap = JSON.parse(String(data[i][jaMapIdx] || '{}')); } catch (e) {}
    result.push('=== グループ ' + group + ' ===');
    Object.keys(jaMap).forEach(function(id) { result.push('  ' + id + ': ' + jaMap[id]); });
  }
  return result.join('\n');
}

/**
 * 指定グループの condition_id "3000" を「中古品」に更新（F・H 用）
 */
function updateCondition3000GroupsFH() {
  return _updateCondition3000ForGroups(['F', 'H']);
}

/**
 * メルカリ表現に合わせて全グループの ja_map_json を一括更新
 *
 * 変更内容:
 *   E  : 2750 ほぼ新品          → 目立った傷や汚れなし
 *   G  : 2750 ほぼ新品          → 未使用に近い
 *        4000 やや傷や汚れあり  → 目立った傷や汚れなし
 *        5000 傷や汚れあり      → やや傷や汚れあり
 *   K  : 3000 目立った傷や汚れなし（使用感あり） → やや傷や汚れあり
 *   M  : 3000 目立った傷や汚れなし（使用感あり） → やや傷や汚れあり
 *   Q  : 3010 傷や汚れあり      → やや傷や汚れあり
 *   R  : 3010 傷や汚れあり      → やや傷や汚れあり
 *   W  : 2750 ほぼ新品          → 目立った傷や汚れなし
 */
function updateAllGroupsMercariAligned() {
  const PATCH = {
    'E': { '2750': '目立った傷や汚れなし' },
    'G': { '2750': '未使用に近い', '4000': '目立った傷や汚れなし', '5000': 'やや傷や汚れあり' },
    'K': { '3000': 'やや傷や汚れあり' },
    'M': { '3000': 'やや傷や汚れあり' },
    'Q': { '3010': 'やや傷や汚れあり' },
    'R': { '3010': 'やや傷や汚れあり' },
    'W': { '2750': '目立った傷や汚れなし' }
  };

  const categoryMasterSs = openCategoryMasterSs();
  if (!categoryMasterSs) return 'カテゴリマスタが開けません';

  const sheet = categoryMasterSs.getSheetByName(SHEET_NAMES.CONDITION_JA_MAP);
  if (!sheet) return SHEET_NAMES.CONDITION_JA_MAP + ' シートが見つかりません';

  const data     = sheet.getDataRange().getValues();
  const headers  = data[0];
  const groupIdx = headers.indexOf('condition_group');
  const jaMapIdx = headers.indexOf('ja_map_json');

  if (groupIdx === -1 || jaMapIdx === -1) return 'condition_group または ja_map_json 列が見つかりません';

  const log = [];

  for (let i = 1; i < data.length; i++) {
    const group = String(data[i][groupIdx]);
    if (!PATCH[group]) continue;

    let jaMap;
    try { jaMap = JSON.parse(String(data[i][jaMapIdx] || '{}')); } catch (e) { continue; }

    const patch   = PATCH[group];
    let changed   = false;

    Object.keys(patch).forEach(function(id) {
      if (jaMap.hasOwnProperty(id)) {
        const oldVal = jaMap[id];
        jaMap[id]    = patch[id];
        log.push('グループ ' + group + ' [' + id + ']: "' + oldVal + '" → "' + patch[id] + '"');
        changed = true;
      }
    });

    if (changed) {
      sheet.getRange(i + 1, jaMapIdx + 1).setValue(JSON.stringify(jaMap));
    }
  }

  const summary = log.length > 0
    ? '更新完了（' + log.length + '件）:\n' + log.join('\n')
    : '更新対象なし';
  Logger.log('[updateAllGroupsMercariAligned] ' + summary);
  return summary;
}

/**
 * グループG に属するカテゴリ一覧を返す（調査用）
 */
function inspectGroupGCategories() {
  const categoryMasterSs = openCategoryMasterSs();
  if (!categoryMasterSs) return 'カテゴリマスタが開けません';

  const sheet = categoryMasterSs.getSheetByName(SHEET_NAMES.CATEGORY_MASTER);
  if (!sheet) return SHEET_NAMES.CATEGORY_MASTER + ' シートが見つかりません';

  const data    = sheet.getDataRange().getValues();
  const headers = data[0];
  const catIdIdx   = headers.indexOf('category_id');
  const catNameIdx = headers.indexOf('category_name');
  const groupIdx   = headers.indexOf('condition_group');

  if (catIdIdx === -1 || groupIdx === -1) return 'category_id または condition_group 列が見つかりません';

  const result = ['=== グループ G のカテゴリ一覧 ==='];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][groupIdx]) !== 'G') continue;
    const id   = String(data[i][catIdIdx]   || '');
    const name = catNameIdx >= 0 ? String(data[i][catNameIdx] || '') : '';
    result.push(id + (name ? ': ' + name : ''));
  }
  if (result.length === 1) result.push('（該当カテゴリなし）');
  return result.join('\n');
}

/**
 * 共通: 指定グループの 3000 を「中古品」に更新
 */
function _updateCondition3000ForGroups(targetGroups) {
  const categoryMasterSs = openCategoryMasterSs();
  if (!categoryMasterSs) return 'カテゴリマスタが開けません';

  const sheet = categoryMasterSs.getSheetByName(SHEET_NAMES.CONDITION_JA_MAP);
  if (!sheet) return SHEET_NAMES.CONDITION_JA_MAP + ' シートが見つかりません';

  const data     = sheet.getDataRange().getValues();
  const headers  = data[0];
  const groupIdx = headers.indexOf('condition_group');
  const jaMapIdx = headers.indexOf('ja_map_json');

  if (groupIdx === -1 || jaMapIdx === -1) return 'condition_group または ja_map_json 列が見つかりません';

  const log = [];
  for (let i = 1; i < data.length; i++) {
    const group = String(data[i][groupIdx]);
    if (targetGroups.indexOf(group) === -1) continue;

    let jaMap;
    try { jaMap = JSON.parse(String(data[i][jaMapIdx] || '{}')); } catch (e) { continue; }

    if (jaMap.hasOwnProperty('3000')) {
      const oldValue = jaMap['3000'];
      jaMap['3000'] = '中古品';
      sheet.getRange(i + 1, jaMapIdx + 1).setValue(JSON.stringify(jaMap));
      log.push('グループ ' + group + ': "' + oldValue + '" → "中古品"');
    }
  }

  const summary = log.length > 0
    ? '更新完了（' + log.length + '件）: ' + log.join(' / ')
    : '更新対象なし';
  Logger.log('[updateCondition3000] ' + summary);
  return summary;
}

/**
 * condition_ja_map のグループ A〜D における condition_id "3000" の表示を「中古品」に一括更新
 */
function updateCondition3000ToChukuhin() {
  return _updateCondition3000ForGroups(['A', 'B', 'C', 'D']);
}
