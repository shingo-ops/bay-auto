/**
 * AnalyticsTest.gs
 *
 * eBay Analytics API (sell/analytics/v1) のテスト関数
 * GASエディタから直接実行して動作を検証する
 */

/**
 * Traffic Report API テスト（昨日〜今日のインプレッション・ビュー数を取得）
 *
 * 実行方法: GASエディタで testTrafficReport を選択して実行
 * 確認ポイント:
 *   - HTTP 200 が返るか
 *   - dimensionMetrics / metricData の構造
 *   - 各listing のデータが取得できるか
 */
function testTrafficReport() {
  Logger.log('=== Analytics API testTrafficReport 開始 ===');

  // トークンを確認・必要なら自動更新
  const tokenStatus = checkAndRefreshToken(null);
  Logger.log('トークン状態: ' + JSON.stringify(tokenStatus));

  const token = getAccessToken();
  if (!token) {
    Logger.log('❌ アクセストークンが取得できません。再認証してください。');
    return;
  }
  Logger.log('✅ トークン取得OK');

  // 日付範囲: 過去30日〜昨日
  const today     = new Date();
  const dateTo    = new Date(today.getTime() - 1 * 24 * 60 * 60 * 1000);
  const dateFrom  = new Date(today.getTime() - 30 * 24 * 60 * 60 * 1000);
  const dateFromStr = Utilities.formatDate(dateFrom, 'GMT', 'yyyyMMdd');
  const dateToStr   = Utilities.formatDate(dateTo,   'GMT', 'yyyyMMdd');
  Logger.log('期間: ' + dateFromStr + ' 〜 ' + dateToStr);

  const filter = 'marketplace_ids:{EBAY_US},date_range:[' + dateFromStr + '..' + dateToStr + ']';
  const metrics = [
    'LISTING_IMPRESSION_TOTAL',
    'LISTING_VIEWS_TOTAL',
    'CLICK_THROUGH_RATE',
    'SALES_CONVERSION_RATE'
  ].join(',');

  const url = 'https://api.ebay.com/sell/analytics/v1/traffic_report'
    + '?dimension=LISTING'
    + '&filter=' + encodeURIComponent(filter)
    + '&metric=' + metrics;

  Logger.log('URL: ' + url);

  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: {
        'Authorization': 'Bearer ' + token,
        'Content-Type':  'application/json'
      },
      muteHttpExceptions: true
    });

    const statusCode   = response.getResponseCode();
    const responseText = response.getContentText();

    Logger.log('HTTP ステータス: ' + statusCode);

    if (statusCode === 200) {
      const json = JSON.parse(responseText);

      // ヘッダー情報（メトリクス定義）
      const headers = json.header ? json.header.dimensionMetrics || [] : [];
      Logger.log('=== ヘッダー（メトリクス定義） ===');
      headers.forEach(function(h) {
        Logger.log('  dimension: ' + JSON.stringify(h.dimensionKeys));
        (h.metrics || []).forEach(function(m) {
          Logger.log('  metric: ' + m.metricKey + ' / type: ' + m.dataType);
        });
      });

      // データ行（最初の5件のみ）
      const records = json.records || [];
      Logger.log('=== レコード数: ' + records.length + '件（最初の5件表示） ===');
      records.slice(0, 5).forEach(function(record, i) {
        Logger.log('--- レコード ' + (i + 1) + ' ---');
        (record.dimensionMetrics || []).forEach(function(dm) {
          Logger.log('  dimensionValues: ' + JSON.stringify(dm.dimensionValues));
          (dm.metricValues || []).forEach(function(mv) {
            Logger.log('  ' + mv.metricKey + ': ' + JSON.stringify(mv.value));
          });
        });
      });

      Logger.log('✅ testTrafficReport 完了');
    } else {
      Logger.log('❌ エラー応答: ' + responseText.substring(0, 1000));
    }

  } catch (e) {
    Logger.log('❌ 例外: ' + e.toString());
  }
}

/**
 * Traffic Report API テスト（単一Listing IDの詳細取得）
 *
 * 対象Listing IDを指定して詳細メトリクスを確認する
 * LISTING_ID に出品シートの Item ID をセットして実行
 */
function testTrafficReportForListing() {
  const LISTING_ID = ''; // ← テスト対象の Item ID をここに入力

  if (!LISTING_ID) {
    Logger.log('❌ LISTING_ID を設定してください');
    return;
  }

  Logger.log('=== Analytics API testTrafficReportForListing 開始: ' + LISTING_ID + ' ===');

  const tokenStatus = checkAndRefreshToken(null);
  const token = getAccessToken();
  if (!token) {
    Logger.log('❌ アクセストークンが取得できません');
    return;
  }

  const today     = new Date();
  const dateTo    = new Date(today.getTime() - 1 * 24 * 60 * 60 * 1000);
  const dateFrom  = new Date(today.getTime() - 30 * 24 * 60 * 60 * 1000);
  const dateFromStr = Utilities.formatDate(dateFrom, 'GMT', 'yyyyMMdd');
  const dateToStr   = Utilities.formatDate(dateTo,   'GMT', 'yyyyMMdd');

  const filter = 'marketplace_ids:{EBAY_US}'
    + ',date_range:[' + dateFromStr + '..' + dateToStr + ']'
    + ',listing_ids:{' + LISTING_ID + '}';

  const url = 'https://api.ebay.com/sell/analytics/v1/traffic_report'
    + '?dimension=LISTING'
    + '&filter=' + encodeURIComponent(filter)
    + '&metric=LISTING_IMPRESSION_TOTAL,LISTING_VIEWS_TOTAL,CLICK_THROUGH_RATE,SALES_CONVERSION_RATE';

  Logger.log('URL: ' + url);

  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: {
        'Authorization': 'Bearer ' + token,
        'Content-Type':  'application/json'
      },
      muteHttpExceptions: true
    });

    Logger.log('HTTP ステータス: ' + response.getResponseCode());
    Logger.log('レスポンス: ' + response.getContentText().substring(0, 3000));

  } catch (e) {
    Logger.log('❌ 例外: ' + e.toString());
  }
}
