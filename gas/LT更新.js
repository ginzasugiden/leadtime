function updateInventoryAndLeadTime() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const productSheet = ss.getSheetByName('商品情報');
  const leadTimeSheet = ss.getSheetByName('リードタイム');
  
  const apiUrl = 'https://api.rms.rakuten.co.jp/es/2.1/inventories/manage-numbers/';
  const serviceSecret = PropertiesService.getScriptProperties().getProperty('serviceSecret');
  const licenseKey = PropertiesService.getScriptProperties().getProperty('licenseKey');
  const authorization = 'ESA ' + Utilities.base64Encode(serviceSecret + ':' + licenseKey);
  
  // リードタイムIDマッピングの作成
  const leadTimeData = leadTimeSheet.getDataRange().getValues();
  const nameToIdMap = {};
  
  for (let i = 1; i < leadTimeData.length; i++) {
    const id = leadTimeData[i][0];
    const name = leadTimeData[i][1];
    nameToIdMap[name] = id;
  }
  
  // 商品情報シートのデータ取得
  const data = productSheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    const manageNumber = data[i][0];  // A列 - 商品管理番号
    const variantId = data[i][2];     // C列 - SKU管理番号
    const quantity = data[i][8];      // I列 - 在庫数
    let leadTime = data[i][9];        // J列 - 出荷リードタイム (名称 or ID)
    
    if (!manageNumber || !variantId || !quantity || !leadTime) {
      Logger.log(`行 ${i + 1}: データ不足 - スキップ`);
      continue;
    }
    
    // リードタイムの変換 (名称 → ID)
    if (isNaN(leadTime)) {  // 数値でなければ名称と判断
      if (nameToIdMap[leadTime]) {
        leadTime = nameToIdMap[leadTime];
      } else {
        Logger.log(`行 ${i + 1}: リードタイム名称「${leadTime}」が見つかりません - スキップ`);
        continue;
      }
    }
    
    const endpoint = `${apiUrl}${manageNumber}/variants/${variantId}`;
    
    const payload = {
      mode: 'ABSOLUTE',
      quantity: quantity,
      operationLeadTime: {
        normalDeliveryTimeId: leadTime
      }
    };
    
    const options = {
      method: 'put',
      headers: {
        'Authorization': authorization,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    try {
      const response = UrlFetchApp.fetch(endpoint, options);
      const responseCode = response.getResponseCode();
      
      if (responseCode === 204) {
        Logger.log(`行 ${i + 1}: 更新成功 - 在庫数: ${quantity}, リードタイム: ${leadTime}`);
      } else {
        Logger.log(`行 ${i + 1}: 更新失敗 - ${response.getContentText()}`);
      }
      
      // QPS制限対策
      Utilities.sleep(1500);
      
    } catch (error) {
      Logger.log(`行 ${i + 1}: エラー発生 - ${error.message}`);
    }
  }
}
