function getRakutenLeadTime() {
  const endpoint = "https://api.rms.rakuten.co.jp/es/1.0/shop/operationLeadTime";
  const serviceSecret = PropertiesService.getScriptProperties().getProperty('serviceSecret');
  const licenseKey = PropertiesService.getScriptProperties().getProperty('licenseKey');
  const sheetName = "リードタイム";
  
  // 認証情報の作成
  const authHeader = Utilities.base64Encode(`${serviceSecret}:${licenseKey}`);
  
  // APIリクエストのオプション設定
  const options = {
    method: "get",
    headers: {
      "Authorization": `ESA ${authHeader}`,
      "Content-Type": "application/xml; charset=UTF-8"
    },
    muteHttpExceptions: true
  };
  
  // APIリクエストを送信
  const response = UrlFetchApp.fetch(endpoint, options);
  const responseCode = response.getResponseCode();
  
  if (responseCode === 200) {
    const xmlData = response.getContentText();
    const parsedData = XmlService.parse(xmlData);
    const root = parsedData.getRootElement();
    const leadTimeList = root.getChild('result')
                              .getChild('operationLeadTimeList')
                              .getChildren('operationLeadTime');

    // スプレッドシートへの書き込み
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    sheet.clear();  // 既存データをクリア
    sheet.appendRow(["ID", "名称", "日数", "在庫ありフラグ", "在庫切れフラグ"]);

    leadTimeList.forEach(item => {
      const id = item.getChildText("operationLeadTimeId");
      const name = item.getChildText("name");
      const days = item.getChildText("numberOfDays");
      const inStockFlag = item.getChildText("inStockDefaultFlag");
      const outOfStockFlag = item.getChildText("outOfStockDefaultFlag");
      
      sheet.appendRow([id, name, days, inStockFlag, outOfStockFlag]);
    });

    Logger.log("データをスプレッドシートに書き出しました。");
  } else {
    Logger.log(`APIリクエスト失敗: ${responseCode}`);
    Logger.log(response.getContentText());
  }
}
