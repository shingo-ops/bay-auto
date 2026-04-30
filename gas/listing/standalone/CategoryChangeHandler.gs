/**
 * CategoryChangeHandler.gs
 *
 * 出品シートで category_id 列が変更された時のデータ更新処理
 *
 * 呼び出し元: listing/container/Code.gs の handleEdit
 *
 * 依存シート:
 *   - category_master_EBAY_US（カテゴリマスタ外部ブック）
 *   - condition_ja_map（カテゴリマスタ外部ブック）
 *   - _cache（現スプレッドシート内、存在する場合のみ使用）
 */

// スペック優先度別の文字色
const SPEC_COLOR_REQUIRED    = '#CC0000'; // 必須: 赤
const SPEC_COLOR_RECOMMENDED = '#1155CC'; // 推奨: 青
const SPEC_COLOR_OPTIONAL    = '#666666'; // 任意: グレー

// 出品シートのヘッダー行
const LISTING_HEADER_ROW = 1;

// スペック列の最大組数
const MAX_SPEC_PAIRS = 30;

/**
 * 出品シートの「カテゴリID」列番号を返す
 *
 * @param {string} spreadsheetId
 * @returns {number|null} 列番号（1-based）
 */
function getCategoryIdColumnNumber(spreadsheetId) {
  try {
    if (spreadsheetId) CURRENT_SPREADSHEET_ID = spreadsheetId;
    const mapping = buildHeaderMapping();
    return mapping['カテゴリID'] || null;
  } catch (e) {
    Logger.log('getCategoryIdColumnNumber エラー: ' + e.toString());
    return null;
  } finally {
    CURRENT_SPREADSHEET_ID = null;
  }
}

/**
 * カテゴリIDからカテゴリ名を取得（category_master_EBAY_US を参照）
 *
 * @param {string} spreadsheetId 出品スプレッドシートID
 * @param {string} categoryId
 * @returns {string} カテゴリ名（見つからない場合は空文字）
 */
function getCategoryNameById(spreadsheetId, categoryId) {
  if (!categoryId || String(categoryId).trim() === '') return '';
  try {
    const row = _getCategoryMasterRow(spreadsheetId, categoryId);
    return row ? row.categoryName : '';
  } catch (e) {
    Logger.log('getCategoryNameById エラー: ' + e.toString());
    return '';
  }
}

/**
 * カテゴリ変更を適用する（YES が選択された場合に呼び出す）
 *
 * 実行内容:
 *   1. category_name 列を新カテゴリ名に更新
 *   2. スペック列（項目名・内容 × MAX_SPEC_PAIRS）をクリア
 *   3. 新カテゴリの required / recommended / optional specs でスペック列を再生成
 *      （項目名に文字色、内容にプルダウン）
 *   4. _cache の item_specs_json に一致する値を自動入力
 *   5. condition_group から状態プルダウンを再生成
 *      ※ 既存値が新リストに存在しない場合: クリアして conditionIncompatible=true を返す
 *
 * @param {string} spreadsheetId
 * @param {string} sheetName  シート名（通常 '出品'）
 * @param {number} row        対象データ行（1-based）
 * @param {string} newCategoryId
 * @returns {{ conditionIncompatible: boolean, oldConditionValue: string }}
 */
function applyCategoryChange(spreadsheetId, sheetName, row, newCategoryId) {
  try {
    if (spreadsheetId) CURRENT_SPREADSHEET_ID = spreadsheetId;

    const ss    = getTargetSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error(sheetName + ' シートが見つかりません');

    // ヘッダーマッピングを構築
    const headerMap = _buildListingHeaderMap(sheet);

    // カテゴリマスタからデータ取得
    const catData = _getCategoryMasterRow(spreadsheetId, newCategoryId);

    // ── 1. category_name 列を更新 ───────────────────────────
    const categoryNameCol = headerMap['カテゴリ'];
    if (categoryNameCol) {
      sheet.getRange(row, categoryNameCol).setValue(catData ? catData.categoryName : '');
    }

    // ── 2. スペック列をクリア ────────────────────────────────
    for (let n = 1; n <= MAX_SPEC_PAIRS; n++) {
      const nameCol  = headerMap['項目名（' + n + '）'];
      const valueCol = headerMap['内容（' + n + '）'];
      if (nameCol) {
        sheet.getRange(row, nameCol)
          .clearContent()
          .clearDataValidations()
          .setFontColor(null);
      }
      if (valueCol) {
        sheet.getRange(row, valueCol)
          .clearContent()
          .clearDataValidations();
      }
    }

    // カテゴリマスタが見つからない場合はスペック再生成をスキップ
    if (!catData) {
      Logger.log('⚠️ カテゴリID ' + newCategoryId + ' がカテゴリマスタに見つかりません。スペック再生成をスキップ。');
      return { conditionIncompatible: false, oldConditionValue: '' };
    }

    // ── 3. スペック列を再生成 ────────────────────────────────
    // required / recommended / optional の順に最大 MAX_SPEC_PAIRS 組
    const specEntries = _buildSpecEntries(catData);
    const aspectValues = catData.aspectValues || {};

    // ── 4. _cache から自動入力データを取得 ──────────────────
    const cachedSpecs = _readCachedSpecs(ss, sheet, row, headerMap);

    const limit = Math.min(specEntries.length, MAX_SPEC_PAIRS);
    for (let i = 0; i < MAX_SPEC_PAIRS; i++) {
      const nameCol  = headerMap['項目名（' + (i + 1) + '）'];
      const valueCol = headerMap['内容（' + (i + 1) + '）'];
      if (i >= limit) break;

      const entry = specEntries[i];

      // 項目名セルに名前と文字色を設定
      if (nameCol) {
        sheet.getRange(row, nameCol)
          .setValue(entry.name)
          .setFontColor(entry.color);
      }

      // 内容セルにプルダウンと自動入力
      if (valueCol) {
        const allowed = aspectValues[entry.name];
        if (Array.isArray(allowed) && allowed.length > 0) {
          const rule = SpreadsheetApp.newDataValidation()
            .requireValueInList(allowed.map(String), true)
            .setAllowInvalid(true)
            .build();
          sheet.getRange(row, valueCol).setDataValidation(rule);
        }
        // _cache に一致する値があれば自動入力
        if (cachedSpecs[entry.name] !== undefined && cachedSpecs[entry.name] !== '') {
          sheet.getRange(row, valueCol).setValue(cachedSpecs[entry.name]);
        }
      }
    }

    // ── 5. 状態プルダウンを再生成 ─────────────────────────────
    const conditionResult = _rebuildConditionDropdown(
      spreadsheetId, sheet, row, headerMap, catData.conditionGroup
    );

    Logger.log('✅ カテゴリ変更適用完了: row=' + row + ' categoryId=' + newCategoryId);
    return conditionResult;

  } catch (e) {
    Logger.log('applyCategoryChange エラー: ' + e.toString());
    throw e;
  } finally {
    CURRENT_SPREADSHEET_ID = null;
  }
}

/**
 * category_id 列を元の値に戻す（NO が選択された場合に呼び出す）
 *
 * @param {string} spreadsheetId
 * @param {string} sheetName
 * @param {number} row
 * @param {string} oldCategoryId
 */
function revertCategoryId(spreadsheetId, sheetName, row, oldCategoryId) {
  try {
    if (spreadsheetId) CURRENT_SPREADSHEET_ID = spreadsheetId;

    const ss    = getTargetSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error(sheetName + ' シートが見つかりません');

    const headerMap = _buildListingHeaderMap(sheet);
    const categoryIdCol = headerMap['カテゴリID'];
    if (categoryIdCol) {
      sheet.getRange(row, categoryIdCol).setValue(oldCategoryId);
      Logger.log('カテゴリID を元に戻しました: row=' + row + ' value=' + oldCategoryId);
    }

  } catch (e) {
    Logger.log('revertCategoryId エラー: ' + e.toString());
    throw e;
  } finally {
    CURRENT_SPREADSHEET_ID = null;
  }
}

// ─────────────────────────────────────────────────────────────
// プライベートヘルパー
// ─────────────────────────────────────────────────────────────

/**
 * 出品シートのヘッダーマッピングを構築
 * （buildHeaderMapping() は Config.gs で定義されているが CURRENT_SPREADSHEET_ID に依存するため
 *  直接 sheet オブジェクトを受け取るローカル版を用意）
 *
 * @param {Sheet} sheet
 * @returns {Object} {ヘッダー名: 列番号(1-based)}
 */
function _buildListingHeaderMap(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return {};
  const headers = sheet.getRange(LISTING_HEADER_ROW, 1, 1, lastCol).getValues()[0];
  const map = {};
  for (let i = 0; i < headers.length; i++) {
    if (headers[i]) map[String(headers[i])] = i + 1;
  }
  return map;
}

/**
 * category_master_EBAY_US から 1 カテゴリ分のデータを取得
 *
 * @param {string} spreadsheetId 出品スプレッドシートID（CURRENT_SPREADSHEET_ID 設定済み前提）
 * @param {string} categoryId
 * @returns {Object|null}
 */
function _getCategoryMasterRow(spreadsheetId, categoryId) {
  if (!categoryId || String(categoryId).trim() === '') return null;
  try {
    // CURRENT_SPREADSHEET_ID はすでに設定済みとして getEbayConfig() を呼ぶ
    const savedId = CURRENT_SPREADSHEET_ID;
    if (spreadsheetId) CURRENT_SPREADSHEET_ID = spreadsheetId;
    const config = getEbayConfig();
    CURRENT_SPREADSHEET_ID = savedId;

    const masterSpreadsheetId = config.categoryMasterSpreadsheetId;
    if (!masterSpreadsheetId) {
      Logger.log('⚠️ カテゴリマスタのspreadsheetIDが未設定（ツール設定 > カテゴリマスタ）');
      return null;
    }

    const masterSs = SpreadsheetApp.openById(masterSpreadsheetId);
    const sheet    = masterSs.getSheetByName('category_master_EBAY_US');
    if (!sheet) {
      Logger.log('⚠️ category_master_EBAY_US シートが見つかりません');
      return null;
    }

    const data    = sheet.getDataRange().getValues();
    const headers = data[0];

    const idx = {
      catId:   headers.indexOf('category_id'),
      catName: headers.indexOf('category_name'),
      req:     headers.indexOf('required_specs_json'),
      rec:     headers.indexOf('recommended_specs_json'),
      opt:     headers.indexOf('optional_specs_json'),
      aspVal:  headers.indexOf('aspect_values_json'),
      group:   headers.indexOf('condition_group')
    };

    if (idx.catId === -1) {
      Logger.log('⚠️ category_id 列が見つかりません');
      return null;
    }

    const parseArr = function(val) {
      if (!val) return [];
      try { return JSON.parse(val); } catch (e) { return []; }
    };
    const parseObj = function(val) {
      if (!val) return {};
      try { return JSON.parse(val); } catch (e) { return {}; }
    };

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idx.catId]) === String(categoryId)) {
        return {
          categoryId:        String(data[i][idx.catId]),
          categoryName:      idx.catName !== -1 ? String(data[i][idx.catName] || '') : '',
          requiredSpecs:     idx.req  !== -1 ? parseArr(data[i][idx.req])  : [],
          recommendedSpecs:  idx.rec  !== -1 ? parseArr(data[i][idx.rec])  : [],
          optionalSpecs:     idx.opt  !== -1 ? parseArr(data[i][idx.opt])  : [],
          aspectValues:      idx.aspVal !== -1 ? parseObj(data[i][idx.aspVal]) : {},
          conditionGroup:    idx.group !== -1 ? String(data[i][idx.group] || '') : ''
        };
      }
    }

    Logger.log('カテゴリID ' + categoryId + ' が category_master_EBAY_US に見つかりません');
    return null;

  } catch (e) {
    Logger.log('_getCategoryMasterRow エラー: ' + e.toString());
    return null;
  }
}

/**
 * required / recommended / optional スペックを優先度順に結合して返す
 *
 * @param {Object} catData _getCategoryMasterRow() の戻り値
 * @returns {Array<{name: string, color: string}>}
 */
function _buildSpecEntries(catData) {
  const entries = [];
  const pushSpecs = function(specs, color) {
    (specs || []).forEach(function(spec) {
      const name = typeof spec === 'object' && spec !== null
        ? (spec.name || spec.localizedAspectName || JSON.stringify(spec))
        : String(spec);
      if (name) entries.push({ name: name, color: color });
    });
  };
  pushSpecs(catData.requiredSpecs,    SPEC_COLOR_REQUIRED);
  pushSpecs(catData.recommendedSpecs, SPEC_COLOR_RECOMMENDED);
  pushSpecs(catData.optionalSpecs,    SPEC_COLOR_OPTIONAL);
  return entries;
}

/**
 * _cache シートから対象行の item_url で item_specs_json を取得
 *
 * ItemURL → スペックURL の順で試みる。
 * _cache シートが存在しない・ヒットしない場合は空オブジェクトを返す。
 *
 * @param {Spreadsheet} ss
 * @param {Sheet} sheet
 * @param {number} row
 * @param {Object} headerMap
 * @returns {Object} {specName: specValue, ...}
 */
function _readCachedSpecs(ss, sheet, row, headerMap) {
  try {
    const cacheSheet = ss.getSheetByName('_cache');
    if (!cacheSheet) {
      Logger.log('_readCachedSpecs: _cache シートが見つかりません');
      return {};
    }

    const lastRow = cacheSheet.getLastRow();
    if (lastRow < 2) {
      Logger.log('_readCachedSpecs: _cache にデータがありません');
      return {};
    }

    const cacheHeaders = cacheSheet.getRange(1, 1, 1, cacheSheet.getLastColumn()).getValues()[0];
    const urlIdx   = cacheHeaders.indexOf('item_url');
    const specsIdx = cacheHeaders.indexOf('item_specs_json');
    Logger.log('_readCachedSpecs: _cache headers=' + JSON.stringify(cacheHeaders));
    if (urlIdx === -1 || specsIdx === -1) {
      Logger.log('_readCachedSpecs: item_url または item_specs_json 列が見つかりません');
      return {};
    }

    const cacheData = cacheSheet.getRange(2, 1, lastRow - 1, cacheHeaders.length).getValues();

    // 出品シートのヘッダー一覧をログ出力（列名確認用）
    Logger.log('_readCachedSpecs: headerMap keys=' + JSON.stringify(Object.keys(headerMap)));

    // 対象行の URL を取得（ItemURL → スペックURL の順）
    const urlCandidates = ['ItemURL', 'スペックURL'].map(function(h) {
      const col = headerMap[h];
      const val = col ? String(sheet.getRange(row, col).getValue() || '') : '';
      Logger.log('_readCachedSpecs: ' + h + ' col=' + col + ' val=' + val);
      return val;
    }).filter(function(u) { return u !== ''; });

    Logger.log('_readCachedSpecs: urlCandidates=' + JSON.stringify(urlCandidates));

    const normalize = function(url) {
      return url.trim().split('?')[0].split('#')[0].replace(/\/$/, '').toLowerCase();
    };

    for (let ci = 0; ci < urlCandidates.length; ci++) {
      const target = normalize(urlCandidates[ci]);
      for (let ri = 0; ri < cacheData.length; ri++) {
        const cacheUrl = normalize(String(cacheData[ri][urlIdx] || ''));
        if (cacheUrl === target) {
          const specsJson = String(cacheData[ri][specsIdx] || '{}');
          Logger.log('_readCachedSpecs: キャッシュヒット row=' + (ri + 2));
          try {
            return JSON.parse(specsJson);
          } catch (e) {
            Logger.log('_readCachedSpecs: item_specs_json パースエラー: ' + e.toString());
            return {};
          }
        }
      }
    }

    Logger.log('_readCachedSpecs: キャッシュにURLが見つかりませんでした');
    return {};
  } catch (e) {
    Logger.log('_readCachedSpecs エラー: ' + e.toString());
    return {};
  }
}

/**
 * 状態（コンディション）プルダウンを再生成
 *
 * condition_group → ja_map_json の values をプルダウン選択肢に設定。
 * 既存値が新リストに存在しない場合: 値をクリアして conditionIncompatible=true を返す。
 *
 * @param {string} spreadsheetId
 * @param {Sheet} sheet
 * @param {number} row
 * @param {Object} headerMap
 * @param {string} conditionGroup
 * @returns {{ conditionIncompatible: boolean, oldConditionValue: string }}
 */
function _rebuildConditionDropdown(spreadsheetId, sheet, row, headerMap, conditionGroup) {
  const result = { conditionIncompatible: false, oldConditionValue: '' };

  const conditionCol = headerMap['状態'];
  if (!conditionCol || !conditionGroup) return result;

  try {
    const savedId = CURRENT_SPREADSHEET_ID;
    if (spreadsheetId) CURRENT_SPREADSHEET_ID = spreadsheetId;
    const config = getEbayConfig();
    CURRENT_SPREADSHEET_ID = savedId;

    const masterSpreadsheetId = config.categoryMasterSpreadsheetId;
    if (!masterSpreadsheetId) return result;

    const masterSs = SpreadsheetApp.openById(masterSpreadsheetId);
    const jaMapSheet = masterSs.getSheetByName('condition_ja_map');
    if (!jaMapSheet) return result;

    const data    = jaMapSheet.getDataRange().getValues();
    const headers = data[0];
    const groupIdx  = headers.indexOf('condition_group');
    const jaMapIdx  = headers.indexOf('ja_map_json');
    if (groupIdx === -1 || jaMapIdx === -1) return result;

    let jaMap = null;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][groupIdx]) === String(conditionGroup)) {
        try {
          jaMap = JSON.parse(data[i][jaMapIdx] || '{}');
        } catch (e) {
          Logger.log('❌ ja_map_json パースエラー: ' + e.toString());
          jaMap = null;
        }
        break;
      }
    }

    if (!jaMap) {
      Logger.log('condition_group ' + conditionGroup + ' の ja_map_json が見つかりません');
      return result;
    }

    const displayOptions = Object.values(jaMap).filter(function(v) {
      return v && String(v).trim() !== '';
    });
    if (displayOptions.length === 0) return result;

    // 現在の状態値を取得しておく
    result.oldConditionValue = String(sheet.getRange(row, conditionCol).getValue() || '');

    // プルダウンを設定
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(displayOptions, true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(row, conditionCol).setDataValidation(rule);

    // 既存値が新リストにない場合はクリアしてフラグを立てる
    if (result.oldConditionValue && displayOptions.indexOf(result.oldConditionValue) === -1) {
      sheet.getRange(row, conditionCol).clearContent();
      result.conditionIncompatible = true;
      Logger.log('⚠️ 既存コンディション値が新カテゴリに存在しないためクリア: ' + result.oldConditionValue);
    }

    Logger.log('✅ 状態プルダウン再生成: group=' + conditionGroup + ' / ' + displayOptions.length + '件');
    return result;

  } catch (e) {
    Logger.log('_rebuildConditionDropdown エラー: ' + e.toString());
    return result;
  }
}
