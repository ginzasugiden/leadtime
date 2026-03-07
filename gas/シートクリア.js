function clearSheetExceptHeader() {
  // 対象のスプレッドシートとシートを取得（シート名は適宜変更してください）
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("商品情報");
  
  // シートの最終行を取得
  var lastRow = sheet.getLastRow();
  
  // もしデータが2行目以降に存在すれば、その範囲のセルの値をクリアする
  if (lastRow > 1) {
    // clearContent() はセルの値だけを消すため、
    // 書式や図形（ボタン）はそのまま維持されます
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }
}
