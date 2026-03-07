/* 商品名検索に統合
async function searchItemsfetchLeadTime() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const keyword = sheet.getRange("A2").getValue().toString().trim();
  if (!keyword) {
    SpreadsheetApp.getUi().alert("A2 に検索キーワードを入力してください。");
    return;
  }

  // A2 だけを keyword にして検索 → シートに結果を書き出す
  searchRakutenItems(keyword);

  // 結果があれば在庫・リードタイム更新
  if (sheet.getLastRow() > 1) {
    fetchLeadTime();
    autoReplaceLeadTime();
  }
}
*/