function onEdit(e) {
  var sheet = e.source.getSheetByName('商品情報');  // 商品情報シートを指定
  var range = e.range;
  var column = range.getColumn();
  var row = range.getRow();
  
  // 編集がE列またはG列で、2行目以降の場合に処理
  if (sheet && row > 1 && (column == 5 || column == 7)) {
    var value = range.getValue();
    var charCount = Math.floor(countRakutenChars(value));  // 文字数を切り捨て
    
    if (column == 5) {
      var maxChars = 127;
      var result = charCount + "文字 / 全角" + maxChars + "文字";
      sheet.getRange(row, 6).setValue(result);
    }
    
    if (column == 7) {
      var maxChars = 87;
      var result = charCount + "文字 / 全角" + maxChars + "文字";
      sheet.getRange(row, 8).setValue(result);
    }
  }
}

// 楽天RMSの文字数計算方式に基づいて全角・半角を区別してカウント
function countRakutenChars(str) {
  var fullWidth = str.match(/[^\x01-\x7E\xA1-\xDF]/g) || [];  // 全角文字
  var halfWidth = str.match(/[\x01-\x7E\xA1-\xDF]/g) || [];   // 半角文字
  return fullWidth.length + (halfWidth.length * 0.5);  // 全角は1文字、半角は0.5文字
}
