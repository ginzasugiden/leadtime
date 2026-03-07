function fetchRakutenItemInfo() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("商品情報");
  const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("log") || SpreadsheetApp.getActiveSpreadsheet().insertSheet("log");
  logSheet.clear();  // ログをクリア
  logSheet.appendRow(["管理番号", "APIレスポンス"]);  // ヘッダー追加

  const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1);
  const manageNumbers = dataRange.getValues();
  const serviceSecret = PropertiesService.getScriptProperties().getProperty('serviceSecret');
  const licenseKey = PropertiesService.getScriptProperties().getProperty('licenseKey');

  manageNumbers.forEach((row, index) => {
  const manageNumber = row[0];
  if (manageNumber) {
    const url = `https://api.rms.rakuten.co.jp/es/2.0/items/manage-numbers/${manageNumber}`;
    const options = {
      method: 'get',
      headers: {
        'Authorization': 'ESA ' + Utilities.base64Encode(serviceSecret + ':' + licenseKey),
        'Content-Type': 'application/json'
      }
    };
    try {
      const response = UrlFetchApp.fetch(url, options);
      const json = JSON.parse(response.getContentText());
      Logger.log(json);

      // ログシートにAPIレスポンスを記録
      logSheet.appendRow([manageNumber, JSON.stringify(json, null, 2), new Date()]);

      let itemNumber = json.itemNumber || '該当なし';
      let skuNumber = '該当なし';
      let systemSkuNumber = '該当なし';

      // SKU情報を取得 (variantsのキーを直接取得)
      if (json.variants) {
        const variantKeys = Object.keys(json.variants);
        if (variantKeys.length > 0) {
          skuNumber = variantKeys[0];  // 最初のキーを取得
          const item = json.variants[skuNumber];
          systemSkuNumber = item.merchantDefinedSkuId || '該当なし';
        }
      }

      // HTMLタグを除去する関数
      function stripHtmlTags(str) {
        return str.replace(/<[^>]*>/g, '');
      }

      const title = stripHtmlTags(json.title || 'タイトルなし');
      const tagline = stripHtmlTags(json.tagline || 'キャッチコピーなし');

      // シートに記入
      sheet.getRange(index + 2, 2).setValue(itemNumber);       // 商品番号
      sheet.getRange(index + 2, 3).setValue(skuNumber);        // SKU管理番号 (variantsのキー)
      sheet.getRange(index + 2, 4).setValue(systemSkuNumber);  // システム連携用SKU番号
      sheet.getRange(index + 2, 5).setValue(title);            // 商品名
      sheet.getRange(index + 2, 7).setValue(tagline);          // キャッチコピー

      // 次のリクエストまで1秒待機
      Utilities.sleep(1000);

    } catch (e) {
      Logger.log(`Error fetching data for ${manageNumber}: ${e}`);
      logSheet.appendRow([manageNumber, `Error: ${e.message}`, new Date()]);
      sheet.getRange(index + 2, 5).setValue('APIエラー');
      
      // エラーが発生しても次のリクエストまで1秒待機
      Utilities.sleep(300);
    }
  }
});

}
