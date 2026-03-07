function updateRakutenItems() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();  // 2行目からデータ取得 (9列)
  
  const serviceSecret = PropertiesService.getScriptProperties().getProperty('serviceSecret');
  const licenseKey = PropertiesService.getScriptProperties().getProperty('licenseKey');

  const headers = {
    'Authorization': 'ESA ' + Utilities.base64Encode(serviceSecret + ':' + licenseKey),
    'Content-Type': 'application/json'
  };

  data.forEach(row => {
    const manageNumber = row[0];  // 商品管理番号 (A列)
    const skuNumber = row[2];     // SKU管理番号 (C列)
    const title = row[4];         // 商品名 (E列)
    const tagline = row[6];       // キャッチコピー (G列)
    const normalDeliveryDateId = row[8];  // 在庫あり時納期管理番号 (I列)

    const payload = {
      title: title,
      tagline: tagline,
      variants: [
        {
          variantId: skuNumber,
          normalDeliveryDateId: normalDeliveryDateId  // SKUごとに在庫納期を設定
        }
      ]
    };

    const options = {
      method: 'patch',
      headers: headers,
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const url = 'https://api.rms.rakuten.co.jp/es/2.0/items/manage-numbers/' + manageNumber;
    
    Logger.log('リクエストURL: ' + url);
    Logger.log('リクエストボディ: ' + JSON.stringify(payload, null, 2));

    const response = UrlFetchApp.fetch(url, options);

    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    Logger.log('商品管理番号: ' + manageNumber);
    Logger.log('SKU管理番号: ' + skuNumber);
    Logger.log('レスポンスコード: ' + responseCode);
    Logger.log('レスポンス: ' + responseText);
  });
}
