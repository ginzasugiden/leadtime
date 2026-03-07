/** ### 商品名検索.gs ###
 * 【A】 メニュー登録用トリガー関数
 * （現在のスプレッドシート上で「拡張機能 → カスタムメニュー」を出すため）
 */

/**
 * onOpen() トリガー：メニューを追加します
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('商品情報')
    .addItem('商品情報検索・A2セルに入力', 'searchItemsfetchLeadTime')
    .addSeparator()
    .addItem('リードタイム更新','updateInventoryAndLeadTime')
    .addSeparator()
    .addItem('シートクリア', 'clearSheetExceptHeader')
    .addToUi();
}

// —— ライブラリ側（商品情報API プロジェクト） ——
/**
 * 実際にメニューを組み立てる関数
 * ※onOpenではなく、名前を変えてエクスポートします
 */
function createMenu() {
  SpreadsheetApp.getUi()
    .createMenu('商品情報')
    .addItem('商品情報検索: A2セルに入力', 'searchItemsfetchLeadTime')
    .addSeparator()
    .addItem('リードタイム更新', 'updateInventoryAndLeadTime')
    .addSeparator()
    .addItem('リードタイム取得','getRakutenLeadTime')
    .addItem('シートクリア', 'clearSheetExceptHeader')
    .addToUi();
}


/**
 * シート上のボタンに割り当てるエントリポイント
 * ライブラリ内（内部）・ライブラリ外（外部）の両方に対応
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

    // —— 内外判定で LT書き換え実行 ——
    try {
      if (typeof RakutenAPI !== 'undefined' && typeof RakutenAPI.autoReplaceLeadTime === 'function') {
        RakutenAPI.autoReplaceLeadTime();  // 外部から呼び出し
      } else if (typeof autoReplaceLeadTime === 'function') {
        autoReplaceLeadTime();             // ライブラリ内（自身）から呼び出し
      } else {
        Logger.log('autoReplaceLeadTime 関数が見つかりません');
      }
    } catch (e) {
      Logger.log('autoReplaceLeadTime の呼び出し時にエラー: ' + e.message);
    }
  }
}



/**
 * A2 の文字列を管理番号完全一致検索（manageNumber）に
 * → ３行目以降にヘッダ行をそのまま残しつつ結果をセット
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

  // ── 2) API コールして results 配列を作成 ──
  const results    = [];
  let cursorMark   = "";

  do {
    // 管理番号で完全一致検索
    let url = `${endpoint}?manageNumber=${encodeURIComponent(keyword)}&hits=100`;
    if (cursorMark) {
      url += `&cursorMark=${encodeURIComponent(cursorMark)}`;
    }
    Logger.log("→ Request URL: " + url);

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
          results.push([
            mn, inum, variantId, skuId,
            title, tlen,
            tagline, tcLen,
            "N/A","N/A"
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

  // ── 3) 結果を 3行目以降に書き出し ──
  if (results.length > 0) {
    sheet.getRange(3, 1, results.length, results[0].length)
         .setValues(results);
  } else {
    sheet.getRange(3, 1, 1, 1).setValue("検索結果なし");
  }
}


/**
 * （そのまま使ってOK）全角1／半角スペース2つで1文字カウント
 */
function countCustomLength(text) {
  const fw = (text.match(/[^\x01-\x7E]/g) || []).length;
  const sp = (text.match(/ /g)        || []).length;
  return fw + Math.floor(sp / 2);
}
