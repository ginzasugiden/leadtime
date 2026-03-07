function searchItemsFromA1() {
  // 対象シート「商品検索」を取得
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("商品検索");
  if (!sheet) {
    Logger.log("シート『商品検索』が見つかりません。");
    return;
  }
  
  // セル A1 の値（検索文字列）を取得
  var searchTerm = sheet.getRange("A1").getValue();
  if (!searchTerm) {
    Logger.log("セルA1に検索文字列がありません。");
    return;
  }
  
  // APIのエンドポイントと認証情報の設定
  var endpoint = "https://api.rms.rakuten.co.jp/es/2.0/items/search";
  var serviceSecret = PropertiesService.getScriptProperties().getProperty('serviceSecret');
  var licenseKey = PropertiesService.getScriptProperties().getProperty('licenseKey');
  var authHeader = "ESA " + Utilities.base64Encode(serviceSecret + ":" + licenseKey);
  
  var options = {
    "method": "get",
    "headers": {
      "Authorization": authHeader,
      "Content-Type": "application/json"
    },
    "muteHttpExceptions": true
  };
  
  // 検索対象フィールド（部分一致検索）
  var searchFields = ["title", "tagline", "manageNumber", "itemNumber", "articleNumber", "variantId"];
  var results = [];
  
  // 各検索項目ごとにAPI呼び出しを実施
  searchFields.forEach(function(field) {
    // クエリパラメータ：各フィールドに対して検索文字列を設定し、最大100件取得
    var queryStr = encodeURIComponent(field) + "=" + encodeURIComponent(searchTerm) + "&hits=100";
    var url = endpoint + "?" + queryStr;
    
    try {
      var response = UrlFetchApp.fetch(url, options);
      var json = JSON.parse(response.getContentText());
      
      if (json.results) {
        json.results.forEach(function(itemObj) {
          var product = itemObj.item;
          var manageNumber = product.manageNumber || "N/A";
          var itemNumber = product.itemNumber || "N/A";
          var title = product.title || "N/A";
          var tagline = product.tagline || "N/A";
          var articleNumber = product.articleNumber || "N/A";
          var variants = product.variants || {};
          
          // 複数のSKU（variantId）がある場合はそれぞれを出力
          var variantKeys = Object.keys(variants);
          if (variantKeys.length > 0) {
            variantKeys.forEach(function(vId) {
              results.push([manageNumber, itemNumber, title, tagline, articleNumber, vId]);
            });
          } else {
            // SKUがない場合は "N/A" として出力
            results.push([manageNumber, itemNumber, title, tagline, articleNumber, "N/A"]);
          }
        });
      }
    } catch (e) {
      Logger.log("フィールド " + field + " でエラー: " + e);
    }
  });
  
  // 商品管理番号の重複を除去（最初の出現のみ採用）
  var uniqueResults = [];
  var seenManageNumbers = {};
  results.forEach(function(row) {
    var manageNumber = row[0]; // 0列目が商品管理番号
    if (!seenManageNumbers[manageNumber]) {
      uniqueResults.push(row);
      seenManageNumbers[manageNumber] = true;
    }
  });
  
  // シートの5行目以降に結果を書き出す
  // まず、既存の5行目以降をクリア
  var lastRow = sheet.getLastRow();
  if (lastRow >= 5) {
    sheet.getRange(5, 1, lastRow - 4, sheet.getLastColumn()).clearContent();
  }
  
  // ヘッダーの書き込み（例：各列の項目名）
  sheet.getRange(4, 1, 1, 6).setValues([["商品管理番号", "商品番号", "商品名", "キャッチコピー", "カタログID", "SKU管理番号"]]);
  
  if (uniqueResults.length > 0) {
    sheet.getRange(5, 1, uniqueResults.length, uniqueResults[0].length).setValues(uniqueResults);
  } else {
    sheet.getRange(5, 1).setValue("検索結果なし");
  }
  
  Logger.log("商品検索完了（重複除外）。");
}
