function updateDeliveryLeadTimeFromSheet() {
  const serviceSecret = "SP240364_MUraZF2oFfd6wq19";
  const licenseKey = "SL240364_Jj6TnbFE5FbEsSQ6";
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  
  const headers = {
    'Authorization': 'ESA ' + Utilities.base64Encode(serviceSecret + ':' + licenseKey),
    'Content-Type': 'application/json'
  };

  for (let i = 1; i < data.length; i++) {
    const manageNumber = data[i][0];  // 商品管理番号（1列目）
    const deliveryLeadTime = data[i][8];  // 在庫あり時納期管理番号（9列目）

    if (manageNumber && deliveryLeadTime) {
      const url = `https://api.rms.rakuten.co.jp/es/2.0/items/manage-numbers/${manageNumber}`;
      
      const payload = JSON.stringify({
        "deliveryLeadTime": deliveryLeadTime
      });

      const options = {
        'method': 'PATCH',
        'headers': headers,
        'payload': payload,
        'muteHttpExceptions': true
      };

      try {
        const response = UrlFetchApp.fetch(url, options);
        const responseCode = response.getResponseCode();
        if (responseCode === 200) {
          Logger.log(`商品管理番号 ${manageNumber} 更新成功`);
          sheet.getRange(i + 1, 10).setValue('更新成功');  // 10列目に結果を記録
        } else {
          Logger.log(`商品管理番号 ${manageNumber} 更新失敗: ${response.getContentText()}`);
          sheet.getRange(i + 1, 10).setValue('更新失敗');
        }
      } catch (e) {
        Logger.log(`商品管理番号 ${manageNumber} エラー: ${e}`);
        sheet.getRange(i + 1, 10).setValue('エラー');
      }
    } else {
      Logger.log(`行 ${i + 1}: データ不足`);
      sheet.getRange(i + 1, 10).setValue('データ不足');
    }
  }
}
