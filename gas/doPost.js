/**
 * Web API エントリポイント（認証機能内蔵）
 * GitHub Pages から doGet (JSONP) で呼び出す
 */

// ── 認証設定 ──
var AUTH_SHEET_ID = '1iYeV2SbOVoRH8Qjm2d1w5tWmhlE_zcc-yO1tDSLN7Rk';
var AUTH_SHEET_NAME = 'api_key';
var SESSION_TTL = 7200; // 2時間（秒）

// ────────────────────────────────────────
// 認証・セッション管理
// ────────────────────────────────────────

/**
 * スプレッドシートで id + pw を照合
 * @returns {{ success:true, userId, sname, sid, email, licenseKey, serviceSecret } | { success:false, message }}
 */
function getUserFromSheet_(userId, password) {
  var ss = SpreadsheetApp.openById(AUTH_SHEET_ID);
  var sheet = ss.getSheetByName(AUTH_SHEET_NAME);
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (String(row[0]) !== userId) continue;

    // flag=0 のみ有効
    if (String(row[9]) !== '0') {
      return { success: false, message: 'このアカウントは無効です' };
    }

    // 有効期限チェック
    var expiry = row[10];
    if (expiry) {
      if (new Date(expiry) <= new Date()) {
        return { success: false, message: 'アカウントの有効期限が切れています' };
      }
    }

    // パスワード照合（「BASE64:」プレフィックスを除去してデコード）
    var pwRaw = String(row[5]);
    if (pwRaw.indexOf('BASE64:') === 0) pwRaw = pwRaw.substring(7);
    var decoded = Utilities.newBlob(Utilities.base64Decode(pwRaw)).getDataAsString();
    if (decoded !== password) {
      return { success: false, message: 'ログインIDまたはパスワードが違います' };
    }

    return {
      success: true,
      userId: String(row[0]),
      licenseKey: String(row[2]),
      serviceSecret: String(row[3]),
      sid: String(row[6]),
      sname: String(row[7]),
      email: String(row[8])
    };
  }

  return { success: false, message: 'ログインIDまたはパスワードが違います' };
}

/**
 * セッショントークンを生成・CacheServiceに保存
 * @returns {string} token
 */
function createSession_(userId, licenseKey, serviceSecret) {
  var token = Utilities.getUuid();
  var cache = CacheService.getScriptCache();
  cache.put('session_' + token, JSON.stringify({
    userId: userId,
    licenseKey: licenseKey,
    serviceSecret: serviceSecret
  }), SESSION_TTL);
  return token;
}

/**
 * トークンからセッション情報を取得
 * @returns {{ userId, licenseKey, serviceSecret } | null}
 */
function validateSession_(token) {
  if (!token) return null;
  var cache = CacheService.getScriptCache();
  var data = cache.get('session_' + token);
  if (!data) return null;
  return JSON.parse(data);
}

/**
 * セッション削除
 */
function deleteSession_(token) {
  if (!token) return;
  var cache = CacheService.getScriptCache();
  cache.remove('session_' + token);
}

// ────────────────────────────────────────
// エントリポイント
// ────────────────────────────────────────

/**
 * JSONP / JSON レスポンスを返すヘルパー
 */
function createJsonResponse_(data, callback) {
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + JSON.stringify(data) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * GETリクエスト（JSONP対応）
 */
function doGet(e) {
  var params = (e && e.parameter) || {};
  var action = params.action || '';
  var callback = params.callback || '';

  if (params.items) {
    try { params.items = JSON.parse(params.items); } catch (_) { params.items = []; }
  }

  return handleAction_(action, params, callback);
}

/**
 * POSTリクエスト（フォールバック用）
 */
function doPost(e) {
  var params = JSON.parse(e.postData.contents || '{}');
  var action = params.action || '';
  return handleAction_(action, params, '');
}

/**
 * アクション振り分け
 */
function handleAction_(action, params, callback) {
  function resp_(data) { return createJsonResponse_(data, callback); }

  try {
    // ── 認証不要: ログイン ──
    if (action === 'login') {
      var userId   = (params && params.userId) || '';
      var password = (params && params.password) || '';
      if (!userId || !password) {
        return resp_({ error: 'userId と password は必須です' });
      }
      var result = getUserFromSheet_(userId, password);
      if (!result.success) {
        return resp_({ error: result.message });
      }
      var token = createSession_(result.userId, result.licenseKey, result.serviceSecret);
      return resp_({ token: token, sname: result.sname });
    }

    // ── 認証不要: ログアウト ──
    if (action === 'logout') {
      var logoutToken = (params && params.token) || '';
      deleteSession_(logoutToken);
      return resp_({ message: 'ログアウトしました' });
    }

    // ── 認証必要なアクション ──
    var token = (params && params.token) || '';
    var creds = validateSession_(token);
    if (!creds) {
      return resp_({ error: 'セッションが無効です。再ログインしてください。', status: 401 });
    }

    var authHeader = 'ESA ' + Utilities.base64Encode(creds.serviceSecret + ':' + creds.licenseKey);

    switch (action) {

      case 'getLeadTimeList':
        return resp_(getLeadTimeListJson_(authHeader));

      case 'searchItems':
        var keyword = (params && params.keyword) || '';
        if (!keyword) {
          return resp_({ error: 'keyword は必須です' });
        }
        return resp_(searchItemsJson_(keyword, authHeader));

      case 'updateLeadTime':
        var items = (params && params.items) || [];
        if (!items.length) {
          return resp_({ error: 'items 配列は必須です' });
        }
        return resp_(updateLeadTimeJson_(items, authHeader));

      default:
        return resp_({ error: '不明な action: ' + action });
    }
  } catch (err) {
    return resp_({ error: err.message });
  }
}

// ────────────────────────────────────────
// 楽天API Wrapper
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
