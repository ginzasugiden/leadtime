/**
 * 高速化版：商品検索とリードタイム変換
 */
async function searchItemsfetchLeadTime() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const sheet   = ss.getSheetByName('商品情報');
  const keyword = sheet.getRange('A2').getValue().toString().trim();

  if (!keyword) {
    SpreadsheetApp.getUi().alert('A2 に検索キーワード（商品管理番号など）を入力してください。');
    return;
  }

  // 商品情報を楽天から検索してシートに出力
  searchRakutenItems(keyword, sheet);

  // 結果が3行目以降に存在する場合に実行
  if (sheet.getLastRow() > 2) {
    fetchLeadTime();
    
    // 高速化版のリードタイム変換を実行
    try {
      if (typeof RakutenAPI !== 'undefined' && typeof RakutenAPI.autoReplaceLeadTimeFast === 'function') {
        RakutenAPI.autoReplaceLeadTimeFast();  // 外部から呼び出し
      } else if (typeof autoReplaceLeadTimeFast === 'function') {
        autoReplaceLeadTimeFast();             // ライブラリ内から呼び出し
      } else {
        Logger.log('autoReplaceLeadTimeFast 関数が見つかりません');
      }
    } catch (e) {
      Logger.log('autoReplaceLeadTimeFast の呼び出し時にエラー: ' + e.message);
    }
  }
}

/**
 * 高速化版：リードタイム変換（一括処理）
 */
function autoReplaceLeadTimeFast() {
  const startTime = new Date();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const productSheet = ss.getSheetByName('商品情報');
  const leadTimeSheet = ss.getSheetByName('リードタイム');
  
  if (!leadTimeSheet) {
    Logger.log('リードタイムシートが見つかりません');
    return;
  }
  
  // リードタイムマッピングの作成
  const leadTimeData = leadTimeSheet.getDataRange().getValues();
  const idToNameMap = {};
  
  for (let i = 1; i < leadTimeData.length; i++) {
    const id = leadTimeData[i][0];
    const name = leadTimeData[i][1];
    if (id && name) {
      idToNameMap[id] = name;
    }
  }
  
  Logger.log('ID→名称マッピング件数:', Object.keys(idToNameMap).length);
  
  // 商品情報シートのヘッダー確認
  const headerRow = productSheet.getRange(1, 1, 1, productSheet.getLastColumn()).getValues()[0];
  const leadTimeIdColIndex = headerRow.indexOf('リードタイムID');
  const leadTimeNameColIndex = headerRow.indexOf('リードタイム名称');
  
  if (leadTimeIdColIndex === -1) {
    Logger.log('リードタイムID列が見つかりません');
    return;
  }
  
  // データ範囲の取得（一括）
  const lastRow = productSheet.getLastRow();
  if (lastRow <= 2) {
    Logger.log('変換対象のデータがありません');
    return;
  }
  
  // 全データを一括取得
  const dataRange = productSheet.getRange(3, 1, lastRow - 2, productSheet.getLastColumn());
  const allData = dataRange.getValues();
  
  // 変換処理（メモリ上で実行）
  let convertedCount = 0;
  const updatedData = allData.map((row, index) => {
    const leadTimeIdValue = row[leadTimeIdColIndex];
    
    if (leadTimeIdValue && leadTimeIdValue !== "N/A" && idToNameMap[leadTimeIdValue]) {
      // リードタイム名称列が存在する場合は更新
      if (leadTimeNameColIndex !== -1) {
        row[leadTimeNameColIndex] = idToNameMap[leadTimeIdValue];
        convertedCount++;
      }
    }
    
    return row;
  });
  
  // 変更されたデータを一括書き込み
  if (convertedCount > 0) {
    dataRange.setValues(updatedData);
  }
  
  const endTime = new Date();
  const executionTime = (endTime - startTime) / 1000;
  
  Logger.log(`リードタイム変換完了。変換件数: ${convertedCount}, 実行時間: ${executionTime}秒`);
}

/**
 * 高速化版：商品検索結果出力
 */
function searchRakutenItems(keyword, sheet) {
  const endpoint      = "https://api.rms.rakuten.co.jp/es/2.0/items/search";
  const props         = PropertiesService.getScriptProperties();
  const serviceSecret = props.getProperty('serviceSecret');
  const licenseKey    = props.getProperty('licenseKey');
  const authHeader    = "ESA " + Utilities.base64Encode(serviceSecret + ":" + licenseKey);

  // ── 1) 既存の検索結果（3行目以降）だけをクリア ──
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow > 2) {
    sheet.getRange(3, 1, lastRow - 2, lastCol).clearContent();
  }

  // ── 2) ヘッダー行の確認と必要な列の追加 ──
  const headerRow = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 10)).getValues()[0];
  
  // 必要な列が存在しない場合は追加
  const requiredColumns = ['管理番号', '商品番号', 'バリエーションID', 'SKU ID', '商品名', '商品名文字数', 'キャッチコピー', 'キャッチコピー文字数', 'リードタイムID', 'リードタイム名称'];
  let needsHeaderUpdate = false;
  
  for (let i = 0; i < requiredColumns.length; i++) {
    if (headerRow[i] !== requiredColumns[i]) {
      headerRow[i] = requiredColumns[i];
      needsHeaderUpdate = true;
    }
  }
  
  // ヘッダーを一括更新
  if (needsHeaderUpdate) {
    sheet.getRange(1, 1, 1, requiredColumns.length).setValues([headerRow.slice(0, requiredColumns.length)]);
  }

  // ── 3) API コールして results 配列を作成 ──
  const results = [];
  let cursorMark = "";

  do {
    // 管理番号で完全一致検索
    let url = `${endpoint}?manageNumber=${encodeURIComponent(keyword)}&hits=100`;
    if (cursorMark) {
      url += `&cursorMark=${encodeURIComponent(cursorMark)}`;
    }

    const resp = UrlFetchApp.fetch(url, {
      method:  "get",
      headers: {
        "Authorization": authHeader,
        "Content-Type":  "application/json"
      },
      muteHttpExceptions: true
    });
    const json = JSON.parse(resp.getContentText());

    if (Array.isArray(json.results)) {
      json.results.forEach(entry => {
        const p       = entry.item;
        const mn      = p.manageNumber   || "";
        const inum    = p.itemNumber     || "";
        const title   = p.title          || "";
        const tagline = p.tagline        || "";
        const tlen    = countCustomLength(title);
        const tcLen   = countCustomLength(tagline);

        Object.keys(p.variants || {}).forEach(variantId => {
          const skuId = p.variants[variantId]?.merchantDefinedSkuId || "";
          
          // 結果配列に追加
          results.push([
            mn, inum, variantId, skuId,
            title, tlen, tagline, tcLen,
            "N/A", "N/A"
          ]);
        });
      });
    }

    // ページング判定
    if (!json.nextCursorMark || json.nextCursorMark === cursorMark) {
      break;
    }
    cursorMark = json.nextCursorMark;

  } while (true);

  // ── 4) 結果を一括書き出し ──
  if (results.length > 0) {
    sheet.getRange(3, 1, results.length, requiredColumns.length).setValues(results);
    Logger.log(`${results.length}件の商品情報を出力しました`);
  } else {
    sheet.getRange(3, 1, 1, 1).setValue("検索結果なし");
  }
}

/**
 * 既存のリードタイム取得機能（fetchLeadTime）
 * ※この関数は別途定義されている前提
 */

/**
 * 文字数カウント関数（変更なし）
 */
function countCustomLength(text) {
  const fw = (text.match(/[^\x01-\x7E]/g) || []).length;
  const sp = (text.match(/ /g)        || []).length;
  return fw + Math.floor(sp / 2);
}