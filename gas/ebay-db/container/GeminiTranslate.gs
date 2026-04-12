/**
 * GeminiTranslate.gs - Gemini 2.5 Flash-Lite による日本語コンディション自動補完
 * ebay-db 原本ブック専用
 *
 * 新スキーマ対応 (condition_ja_map 1グループ1行):
 *   - condition_group  : グループラベル
 *   - ja_map_json      : {"1000":"新品、未使用",...}
 *
 * fillMissingJaDisplay() は ja_map_json に空値が含まれるグループを検出し、
 * Gemini API で日本語表示名を生成して補完します。
 */

var GEMINI_API_ENDPOINT = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash-lite:generateContent';

// condition_id → 英語名のフォールバックマップ（APIから取得できない場合用）
var CONDITION_NAME_FALLBACK = {
  '1000': 'New',
  '1500': 'New other (see details)',
  '1750': 'New with defects',
  '1900': 'Unused',
  '2000': 'Certified refurbished',
  '2010': 'Excellent - Refurbished',
  '2020': 'Very Good - Refurbished',
  '2030': 'Good - Refurbished',
  '2500': 'Seller refurbished',
  '2750': 'Like New',
  '2990': 'Pre-owned - Excellent',
  '3000': 'Used',
  '3010': 'Pre-owned - Fair',
  '4000': 'Very Good',
  '5000': 'Good',
  '6000': 'Acceptable',
  '7000': 'For parts or not working'
};

/**
 * condition_ja_map シートの ja_map_json に空値があるグループを Gemini で自動補完
 *
 * 新スキーマ: 1グループ1行。ja_map_json の各 condition_id の値が空の場合に補完。
 * 標準 condition_id は generate_csv.py の JA_DISPLAY_DEFAULT で事前に埋まるため、
 * このステップは未知の condition_id（将来の新規追加等）向けのフォールバック。
 *
 * @returns {Object} { filled: number, failed: number }
 */
function fillMissingJaDisplay() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('condition_ja_map');
  if (!sheet) {
    throw new Error('condition_ja_map シートが見つかりません');
  }

  var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    throw new Error('GEMINI_API_KEY がスクリプトプロパティに設定されていません');
  }

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var colIdx = {};
  headers.forEach(function(h, i) { colIdx[h] = i; });

  var groupIdx  = colIdx['condition_group'];
  var jaMapIdx  = colIdx['ja_map_json'];

  if (groupIdx === undefined || jaMapIdx === undefined) {
    throw new Error('condition_ja_map のスキーマが不正です。condition_group または ja_map_json 列が見つかりません');
  }

  var filled = 0;
  var failed = 0;

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var group = String(row[groupIdx]);
    var jaMapJson = row[jaMapIdx];

    // ja_map_json をパース
    var jaMap = {};
    try {
      if (jaMapJson && jaMapJson !== '{}') {
        jaMap = JSON.parse(jaMapJson);
      }
    } catch (e) {
      Logger.log('ja_map_json パースエラー（グループ' + group + '）: ' + e.toString());
    }

    // 空値の condition_id を抽出
    var emptyIds = Object.keys(jaMap).filter(function(cid) { return !jaMap[cid]; });

    if (emptyIds.length === 0) continue; // 全て埋まっている

    Logger.log('グループ' + group + ': ' + emptyIds.length + '件の空 ja_display を補完します');
    var changed = false;

    for (var k = 0; k < emptyIds.length; k++) {
      var cid = emptyIds[k];
      var conditionName = CONDITION_NAME_FALLBACK[cid] || ('Condition ' + cid);

      try {
        var result = callGeminiForJaDisplay(apiKey, conditionName, '', '');
        if (result.ja_display) {
          jaMap[cid] = result.ja_display;
          changed = true;
          filled++;
          Logger.log('  condition_id=' + cid + ' → ' + result.ja_display);
        } else {
          failed++;
          Logger.log('  condition_id=' + cid + ': Gemini 応答なし');
        }
      } catch (e) {
        failed++;
        Logger.log('  Gemini API エラー (condition_id=' + cid + '): ' + e.toString());
      }

      Utilities.sleep(1000);
    }

    if (changed) {
      sheet.getRange(i + 1, jaMapIdx + 1).setValue(JSON.stringify(jaMap));
      Logger.log('グループ' + group + ': ja_map_json 更新完了');
    }
  }

  Logger.log('ja_map_json 補完完了: filled=' + filled + ', failed=' + failed);
  return { filled: filled, failed: failed };
}

/**
 * Gemini API を呼び出して ja_display と ja_description を生成
 *
 * @param {string} apiKey
 * @param {string} conditionName - 英語デフォルト名
 * @param {string} conditionEnum - enum値（空可）
 * @param {string} categoryDisplay - カテゴリ固有表示名（空可）
 * @returns {Object} { ja_display: string, ja_description: string }
 */
function callGeminiForJaDisplay(apiKey, conditionName, conditionEnum, categoryDisplay) {
  var displayTarget = categoryDisplay || conditionName;
  var prompt = 'あなたはeBayの日本語出品サポートAIです。\n\n'
    + 'eBayコンディション「' + displayTarget + '」'
    + (conditionEnum ? '（enum: ' + conditionEnum + '）' : '') + 'の\n'
    + '日本語表記を以下のJSON形式で返してください。\n'
    + 'メルカリ・ヤフオクの出品者が直感的に選択できる表現にしてください。\n\n'
    + '出力形式（JSONのみ）:\n'
    + '{"ja_display": "新品・未使用", "ja_description": "未開封・未使用の新品状態"}';

  var requestBody = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      responseMimeType: 'application/json',
      responseSchema: {
        type: 'OBJECT',
        properties: {
          ja_display:     { type: 'STRING' },
          ja_description: { type: 'STRING' }
        },
        required: ['ja_display', 'ja_description']
      }
    }
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(requestBody),
    muteHttpExceptions: true
  };

  var url = GEMINI_API_ENDPOINT + '?key=' + apiKey;
  var response = UrlFetchApp.fetch(url, options);

  if (response.getResponseCode() !== 200) {
    throw new Error('Gemini API エラー: ' + response.getResponseCode() + ' ' + response.getContentText());
  }

  var json = JSON.parse(response.getContentText());
  var text = json.candidates[0].content.parts[0].text;
  return JSON.parse(text);
}
