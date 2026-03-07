/**
 * Web API エントリポイント（GasAuth ライブラリで認証）
 * GitHub Pages から doGet (JSONP風) で呼び出す
 * ※ GAS Web App は POST で CORS ヘッダーを返せないため GET に統一
 */

/**
 * JSONレスポンスを返すヘルパー
 */
function createJsonResponse_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * GETリクエスト（全アクションをURLパラメータで受け取る）
 * items等の複雑なデータは JSON文字列 として渡す
 */
function doGet(e) {
  var params = (e && e.parameter) || {};
  var action = params.action || '';

  // items パラメータがJSON文字列の場合はパースする
  if (params.items) {
    try { params.items = JSON.parse(params.items); } catch (_) { params.items = []; }
  }

  return handleAction_(action, params);
}

/**
 * POSTリクエスト（フォールバック用）
 */
function doPost(e) {
  var params = JSON.parse(e.postData.contents || '{}');
  var action = params.action || '';
  return handleAction_(action, params);
}

/**
 * アクション振り分け
 */
function handleAction_(action, params) {
  try {
    // ── 認証不要: ログイン ──
    if (action === 'login') {
      var userId   = (params && params.userId) || '';
      var password = (params && params.password) || '';
      if (!userId || !password) {
        return createJsonResponse_({ error: 'userId と password は必須です' });
      }
      var result = GasAuth.getUserFromSheet(userId, password);
      if (!result.success) {
        return createJsonResponse_({ error: result.message });
      }
      var token = GasAuth.createSession(result.userId, result.licenseKey, result.serviceSecret);
      return createJsonResponse_({ token: token, sname: result.sname });
    }

    // ── 認証不要: ログアウト ──
    if (action === 'logout') {
      var logoutToken = (params && params.token) || '';
      GasAuth.deleteSession(logoutToken);
      return createJsonResponse_({ message: 'ログアウトしました' });
    }

    // ── 認証必要なアクション ──
    var token = (params && params.token) || '';
    var creds = GasAuth.validateSession(token);
    if (!creds) {
      return createJsonResponse_({ error: 'セッションが無効です。再ログインしてください。', status: 401 });
    }

    var authHeader = 'ESA ' + Utilities.base64Encode(creds.serviceSecret + ':' + creds.licenseKey);

    switch (action) {

      case 'getLeadTimeList':
        return createJsonResponse_(getLeadTimeListJson_(authHeader));

      case 'searchItems':
        var keyword = (params && params.keyword) || '';
        if (!keyword) {
          return createJsonResponse_({ error: 'keyword は必須です' });
        }
        return createJsonResponse_(searchItemsJson_(keyword, authHeader));

      case 'updateLeadTime':
        var items = (params && params.items) || [];
        if (!items.length) {
          return createJsonResponse_({ error: 'items 配列は必須です' });
        }
        return createJsonResponse_(updateLeadTimeJson_(items, authHeader));

      default:
        return createJsonResponse_({ error: '不明な action: ' + action });
    }
  } catch (err) {
    return createJsonResponse_({ error: err.message });
  }
}

// ────────────────────────────────────────
// Wrapper 関数（セッションの認証情報を使用）
// ────────────────────────────────────────

/**
 * リードタイム一覧を配列で返す
 */
function getLeadTimeListJson_(authHeader) {
  var endpoint = 'https://api.rms.rakuten.co.jp/es/1.0/shop/operationLeadTime';

  var response = UrlFetchApp.fetch(endpoint, {
    method: 'get',
    headers: {
      'Authorization': authHeader,
      'Content-Type': 'application/xml; charset=UTF-8'
    },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    return { error: 'API失敗: ' + response.getResponseCode() };
  }

  var root = XmlService.parse(response.getContentText()).getRootElement();
  var list = root.getChild('result')
                 .getChild('operationLeadTimeList')
                 .getChildren('operationLeadTime');

  var results = [];
  for (var i = 0; i < list.length; i++) {
    var item = list[i];
    results.push({
      id: item.getChildText('operationLeadTimeId'),
      name: item.getChildText('name'),
      days: item.getChildText('numberOfDays'),
      inStockFlag: item.getChildText('inStockDefaultFlag'),
      outOfStockFlag: item.getChildText('outOfStockDefaultFlag')
    });
  }
  return { leadTimeList: results };
}

/**
 * 管理番号で検索し、結果を配列で返す
 */
function searchItemsJson_(keyword, authHeader) {
  var endpoint = 'https://api.rms.rakuten.co.jp/es/2.0/items/search';
  var results = [];
  var cursorMark = '';

  do {
    var url = endpoint + '?manageNumber=' + encodeURIComponent(keyword) + '&hits=100';
    if (cursorMark) {
      url += '&cursorMark=' + encodeURIComponent(cursorMark);
    }

    var resp = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: {
        'Authorization': authHeader,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    });
    var json = JSON.parse(resp.getContentText());

    if (Array.isArray(json.results)) {
      for (var i = 0; i < json.results.length; i++) {
        var p = json.results[i].item;
        var variantKeys = Object.keys(p.variants || {});
        for (var j = 0; j < variantKeys.length; j++) {
          var vid = variantKeys[j];
          var skuId = (p.variants[vid] && p.variants[vid].merchantDefinedSkuId) || '';
          results.push({
            manageNumber: p.manageNumber || '',
            itemNumber: p.itemNumber || '',
            variantId: vid,
            skuId: skuId,
            title: p.title || '',
            tagline: p.tagline || ''
          });
        }
      }
    }

    if (!json.nextCursorMark || json.nextCursorMark === cursorMark) break;
    cursorMark = json.nextCursorMark;
  } while (true);

  return { items: results };
}

/**
 * items配列の各要素 { manageNumber, variantId, quantity, leadTimeId } を更新
 */
function updateLeadTimeJson_(items, authHeader) {
  var results = [];

  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    var endpoint = 'https://api.rms.rakuten.co.jp/es/2.1/inventories/manage-numbers/'
      + item.manageNumber + '/variants/' + item.variantId;

    var payload = {
      mode: 'ABSOLUTE',
      quantity: item.quantity,
      operationLeadTime: {
        normalDeliveryTimeId: item.leadTimeId
      }
    };

    try {
      var response = UrlFetchApp.fetch(endpoint, {
        method: 'put',
        headers: {
          'Authorization': authHeader,
          'Content-Type': 'application/json'
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });

      var code = response.getResponseCode();
      results.push({
        manageNumber: item.manageNumber,
        variantId: item.variantId,
        success: code === 204,
        status: code
      });
    } catch (err) {
      results.push({
        manageNumber: item.manageNumber,
        variantId: item.variantId,
        success: false,
        error: err.message
      });
    }

    Utilities.sleep(1500);
  }

  return { results: results };
}
