function fetchLeadTime() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('商品情報');
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  
  const serviceSecret = PropertiesService.getScriptProperties().getProperty('serviceSecret');
  const licenseKey = PropertiesService.getScriptProperties().getProperty('licenseKey');
  const authorization = Utilities.base64Encode(serviceSecret + ':' + licenseKey);

  const batchSize = 2;  // 1回のリクエスト数を制限
  const stockResults = [];
  const leadTimeResults = [];

  for (let i = 0; i < data.length; i += batchSize) {
    const batch = data.slice(i, i + batchSize);
    const urls = [];
    const manageNumbers = [];
    const variantIds = [];
    
    batch.forEach((row) => {
      const manageNumber = row[0];
      const variantId = row[2];
      
      if (manageNumber && variantId) {
        const url = `https://api.rms.rakuten.co.jp/es/2.1/inventories/manage-numbers/${manageNumber}/variants/${variantId}`;
        urls.push(url);
        manageNumbers.push(manageNumber);
        variantIds.push(variantId);
      }
    });

    const options = {
      'method': 'get',
      'headers': {
        'Authorization': `ESA ${authorization}`
      },
      'muteHttpExceptions': true
    };

    const responses = UrlFetchApp.fetchAll(urls.map(url => ({url, ...options})));

    responses.forEach((response, index) => {
      let stockCount = 'N/A';
      let leadTime = 'N/A';
      
      try {
        const responseCode = response.getResponseCode();
        
        if (responseCode === 200) {
          const json = JSON.parse(response.getContentText());
          stockCount = json.quantity !== undefined ? json.quantity : 'N/A';
          leadTime = json.operationLeadTime?.normalDeliveryTimeId ?? 'N/A';
        } else {
          Logger.log(`Error fetching data for ${manageNumbers[index]}: ${response.getContentText()}`);
          stockCount = 'Error';
          leadTime = 'Error';
        }
      } catch (e) {
        Logger.log(`Unexpected error for ${manageNumbers[index]} - ${variantIds[index]}: ${e.message}`);
        stockCount = 'Error';
        leadTime = 'Error';
      }
      
      stockResults.push([stockCount]);
      leadTimeResults.push([leadTime]);
    });

    // QPS制限を避けるために5秒待機
    Utilities.sleep(800);
  }


  // 書き込み前に必ずチェックを入れる
  if (stockResults.length > 0) {
    sheet.getRange(2, 9, stockResults.length, 1)
         .setValues(stockResults);
  }
  if (leadTimeResults.length > 0) {
    sheet.getRange(2, 10, leadTimeResults.length, 1)
         .setValues(leadTimeResults);
  }

}
