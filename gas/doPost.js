/**
 * Web API エントリポイント（認証機能内蔵・GasAuthライブラリ不要）
 * GitHub Pages から doGet (JSONP) で呼び出す
 */

// ────────────────────────────────────────
// 認証・セッション管理
// ────────────────────────────────────────

/**
 * BASE64値を安全にデコードする
 * "BASE64:" プレフィックスがあれば除去してからデコード
 * デコード失敗時は元の値をそのまま返す
 */
function safeBase64Decode_(value) {
  if (!value) return '';
  var raw = String(value);
  if (raw.indexOf('BASE64:') === 0) {
    raw = raw.substring(7);
  }
  try {
    return Utilities.newBlob(Utilities.base64Decode(raw)).getDataAsString();
  } catch (e) {
    Logger.log('[safeBase64Decode_] failed: ' + e.message + ' raw prefix=' + raw.substring(0, 10));
    return raw;
  }
}

/**
 * スプレッドシートの api_key タブからユーザー情報を取得
 * @param {string} userId - A列のid
 * @param {string} password - 平文パスワード
 * @returns {object|null} ユーザー情報 or null
 */
function getUserFromSheet_(userId, password) {
  var SHEET_ID = '1iYeV2SbOVoRH8Qjm2d1w5tWmhlE_zcc-yO1tDSLN7Rk';
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('api_key');
  var data = sheet.getDataRange().getValues();

  Logger.log('[getUserFromSheet_] rows=' + data.length);

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var rowId = String(row[0]).trim(); // A列: id

    if (rowId !== userId) continue;

    Logger.log('[getUserFromSheet_] found userId at row ' + (i + 1));

    // flag チェック（J列 = index 9）
    var flag = row[9];
    Logger.log('[getUserFromSheet_] flag=' + flag);
    if (flag == 1) return null;

    // expiry チェック（K列 = index 10）
    var expiry = row[10];
    if (expiry) {
      var expiryDate = new Date(expiry);
      if (expiryDate < new Date()) {
        Logger.log('[getUserFromSheet_] expired: ' + expiryDate);
        return null;
      }
    }

    // パスワード照合（F列 = index 5）
    var storedPw = safeBase64Decode_(row[5]);
    Logger.log('[getUserFromSheet_] pw match=' + (storedPw === password));

    if (storedPw !== password) return null;

    // ユーザー情報を返却（licenseKey / serviceSecret は RAW のまま返す）
    var result = {
      id: rowId,
      licenseKey: String(row[2]),      // C列: BASE64のまま
      serviceSecret: String(row[3]),   // D列: BASE64のまま
      sid: String(row[6]),             // G列: 店舗ID
      sname: String(row[7]),           // H列: 店舗名
      email: String(row[8]),           // I列
      role: String(row[11])            // L列
    };

    Logger.log('[getUserFromSheet_] success: sname=' + result.sname);
    return result;
  }

  Logger.log('[getUserFromSheet_] user not found: ' + userId);
  return null;
}

/**
 * セッション作成
 * @returns {string} token
 */
function createSession_(userId, userData) {
  var token = Utilities.getUuid();
  var cache = CacheService.getScriptCache();
  var sessionData = JSON.stringify({
    userId: userId,
    licenseKey: userData.licenseKey,       // BASE64のまま保存
    serviceSecret: userData.serviceSecret, // BASE64のまま保存
    sid: userData.sid,
    sname: userData.sname,
    email: userData.email,
    role: userData.role,
    created: new Date().toISOString()
  });
  cache.put('session_' + token, sessionData, 3600);
  Logger.log('[createSession_] userId=' + userId + ' token=' + token);
  return token;
}

/**
 * セッション検証
 * @returns {object|null} セッションデータ or null
 */
function validateSession_(token) {
  if (!token) return null;
  Logger.log('[validateSession_] token=' + token.substring(0, 8) + '...');
  var cache = CacheService.getScriptCache();
  var data = cache.get('session_' + token);
  if (!data) {
    Logger.log('[validateSession_] not found');
    return null;
  }
  var session = JSON.parse(data);
  Logger.log('[validateSession_] userId=' + session.userId);
  return session;
}

/**
 * セッション削除
 */
function deleteSession_(token) {
  if (!token) return;
  var cache = CacheService.getScriptCache();
  cache.remove('session_' + token);
  Logger.log('[deleteSession_] removed');
}

/**
 * 楽天RMS API用の ESA 認証ヘッダーを生成
 * セッションから licenseKey / serviceSecret を取得し、デコードしてヘッダーを作る
 * @param {object} session - validateSession_ の戻り値
 * @returns {string} "ESA xxxx" 形式のヘッダー値
 */
function buildEsaAuthHeader_(session) {
  var lk = safeBase64Decode_(session.licenseKey);
  var ss = safeBase64Decode_(session.serviceSecret);
  Logger.log('[buildEsaAuthHeader_] lk length=' + lk.length + ' ss length=' + ss.length);
  Logger.log('[buildEsaAuthHeader_] lk prefix=' + lk.substring(0, 4) + '...');
  Logger.log('[buildEsaAuthHeader_] ss prefix=' + ss.substring(0, 4) + '...');
  var authHeader = 'ESA ' + Utilities.base64Encode(ss + ':' + lk);
  Logger.log('[buildEsaAuthHeader_] header prefix=' + authHeader.substring(0, 15) + '...');
  return authHeader;
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
  Logger.log('[doGet] action=' + action + ' callback=' + (callback ? 'yes' : 'no'));

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
  Logger.log('[doPost] action=' + action);
  return handleAction_(action, params, '');
}

/**
 * アクション振り分け
 */
function handleAction_(action, params, callback) {
  function resp_(data) {
    Logger.log('[handleAction_] response: ' + JSON.stringify(data).substring(0, 200));
    return createJsonResponse_(data, callback);
  }

  try {
    Logger.log('[handleAction_] action=' + action);

    // ── 認証不要: ログイン ──
    if (action === 'login') {
      var userId   = (params && params.userId) || '';
      var password = (params && params.password) || '';
      Logger.log('[handleAction_] login: userId=' + userId);
      if (!userId || !password) {
        return resp_({ error: 'userId と password は必須です' });
      }
      var user = getUserFromSheet_(userId, password);
      if (!user) {
        Logger.log('[handleAction_] login failed');
        return resp_({ error: 'ログインIDまたはパスワードが違います' });
      }
      var token = createSession_(userId, user);
      return resp_({ token: token, sname: user.sname });
    }

    // ── 認証不要: ログアウト ──
    if (action === 'logout') {
      var logoutToken = (params && params.token) || '';
      Logger.log('[handleAction_] logout');
      deleteSession_(logoutToken);
      return resp_({ message: 'ログアウトしました' });
    }

    // ── 認証必要なアクション ──
    var token = (params && params.token) || '';
    var session = validateSession_(token);
    if (!session) {
      Logger.log('[handleAction_] auth failed: token=' + (token ? token.substring(0, 8) + '...' : 'empty'));
      return resp_({ error: 'セッションが無効です。再ログインしてください。', status: 401 });
    }
    Logger.log('[handleAction_] authenticated: userId=' + session.userId);

    Logger.log('[handleAction_] calling buildEsaAuthHeader_...');
    var authHeader = buildEsaAuthHeader_(session);
    Logger.log('[handleAction_] authHeader generated, length=' + authHeader.length);

    switch (action) {

      case 'getLeadTimeList':
        Logger.log('[handleAction_] getLeadTimeList');
        return resp_(getLeadTimeListJson_(authHeader));

      case 'searchItems':
        var keyword = (params && params.keyword) || '';
        Logger.log('[handleAction_] searchItems: keyword=' + keyword);
        if (!keyword) {
          return resp_({ error: 'keyword は必須です' });
        }
        return resp_(searchItemsJson_(keyword, authHeader));

      case 'searchItemsWithLT':
        var kwLT = (params && params.keyword) || '';
        Logger.log('[handleAction_] searchItemsWithLT: keyword=' + kwLT);
        if (!kwLT) {
          return resp_({ error: 'keyword は必須です' });
        }
        return resp_(searchItemsWithLTJson_(kwLT, authHeader));

      case 'updateLeadTime':
        var items = (params && params.items) || [];
        Logger.log('[handleAction_] updateLeadTime: items=' + items.length);
        if (!items.length) {
          return resp_({ error: 'items 配列は必須です' });
        }
        return resp_(updateLeadTimeJson_(items, authHeader));

      default:
        return resp_({ error: '不明な action: ' + action });
    }
  } catch (err) {
    Logger.log('[handleAction_] ERROR: ' + err.message + '\n' + err.stack);
    return resp_({ error: err.message, stack: err.stack });
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

  Logger.log('[getLeadTimeListJson_] endpoint=' + endpoint);
  var response = UrlFetchApp.fetch(endpoint, {
    method: 'get',
    headers: {
      'Authorization': authHeader,
      'Content-Type': 'application/xml; charset=UTF-8'
    },
    muteHttpExceptions: true
  });

  var respCode = response.getResponseCode();
  Logger.log('[getLeadTimeListJson_] status=' + respCode);
  if (respCode !== 200) {
    Logger.log('[getLeadTimeListJson_] body=' + response.getContentText());
    return { error: 'API失敗: ' + respCode, detail: response.getContentText() };
  }

  var rawText = response.getContentText();
  Logger.log('[getLeadTimeListJson_] raw response=' + rawText.substring(0, 500));

  var root = XmlService.parse(rawText).getRootElement();
  var list = root.getChild('result')
                 .getChild('operationLeadTimeList')
                 .getChildren('operationLeadTime');

  // 最初の要素の全子要素名をログ出力（フィールド特定用）
  if (list.length > 0) {
    var children = list[0].getChildren();
    var childNames = [];
    for (var c = 0; c < children.length; c++) {
      childNames.push(children[c].getName() + '=' + children[c].getText());
    }
    Logger.log('[getLeadTimeListJson_] first item children: ' + childNames.join(', '));
  }

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

  // operationLeadTimeId 昇順ソートし、1始まりの連番(delvdateNumber)を振る
  results.sort(function(a, b) { return Number(a.id) - Number(b.id); });
  for (var i = 0; i < results.length; i++) {
    results[i].delvdateNumber = i + 1;
  }
  Logger.log('[getLeadTimeListJson_] results: ' + JSON.stringify(results));
  return { leadTimeList: results };
}

/**
 * 商品検索 + 各商品の現在のLT設定を取得して返す
 */
function searchItemsWithLTJson_(keyword, authHeader) {
  var searchResult = searchItemsJson_(keyword, authHeader);
  var items = searchResult.items || [];

  // manageNumber の重複を排除（同一商品の複数variant対応）
  var seen = {};
  var uniqueManageNumbers = [];
  for (var i = 0; i < items.length; i++) {
    if (!seen[items[i].manageNumber]) {
      seen[items[i].manageNumber] = true;
      uniqueManageNumbers.push(items[i].manageNumber);
    }
  }

  Logger.log('[searchItemsWithLTJson_] unique manageNumbers=' + uniqueManageNumbers.length);

  // 各商品のLT設定を取得
  var settingsMap = {};
  for (var i = 0; i < uniqueManageNumbers.length; i++) {
    try {
      var settings = getInventorySettings_(uniqueManageNumbers[i], authHeader);
      // 429 エラーの場合はリトライ
      if (settings === 429) {
        Logger.log('[searchItemsWithLTJson_] 429 rate limit, retrying after 2000ms...');
        Utilities.sleep(2000);
        settings = getInventorySettings_(uniqueManageNumbers[i], authHeader);
      }
      if (settings && settings !== 429) {
        settingsMap[uniqueManageNumbers[i]] = settings;
      }
      if (i < uniqueManageNumbers.length - 1) {
        Utilities.sleep(1100);
      }
    } catch (e) {
      Logger.log('[searchItemsWithLTJson_] error for ' + uniqueManageNumbers[i] + ': ' + e.message);
    }
  }

  // 商品情報 + LT設定をマージ
  var itemsWithLT = [];
  for (var i = 0; i < uniqueManageNumbers.length; i++) {
    var mn = uniqueManageNumbers[i];
    var settings = settingsMap[mn];
    // items から title を取得
    var title = '';
    for (var k = 0; k < items.length; k++) {
      if (items[k].manageNumber === mn) { title = items[k].title; break; }
    }

    if (settings && settings.variants) {
      var variantKeys = Object.keys(settings.variants);
      for (var j = 0; j < variantKeys.length; j++) {
        var vId = variantKeys[j];
        var variant = settings.variants[vId];
        itemsWithLT.push({
          manageNumber: mn,
          title: title,
          variantId: vId,
          normalDeliveryDateId: variant.normalDeliveryDateId || null,
          backOrderDeliveryDateId: variant.backOrderDeliveryDateId || null,
          backOrderFlag: variant.backOrderFlag || false
        });
      }
    } else {
      itemsWithLT.push({
        manageNumber: mn,
        title: title,
        variantId: '',
        normalDeliveryDateId: null,
        backOrderDeliveryDateId: null,
        backOrderFlag: false
      });
    }
  }

  Logger.log('[searchItemsWithLTJson_] total items with LT=' + itemsWithLT.length);
  return { items: itemsWithLT };
}

/**
 * 管理番号で検索し、結果を配列で返す
 */
function searchItemsJson_(keyword, authHeader) {
  Logger.log('[searchItemsJson_] authHeader length=' + (authHeader ? authHeader.length : 'null'));
  var endpoint = 'https://api.rms.rakuten.co.jp/es/2.0/items/search';
  var results = [];
  var cursorMark = '';

  do {
    var url = endpoint + '?manageNumber=' + encodeURIComponent(keyword) + '&hits=100';
    if (cursorMark) {
      url += '&cursorMark=' + encodeURIComponent(cursorMark);
    }
    Logger.log('[searchItemsJson_] url=' + url);

    var response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: {
        'Authorization': authHeader
      },
      muteHttpExceptions: true
    });

    var status = response.getResponseCode();
    Logger.log('[searchItemsJson_] status=' + status);

    if (status !== 200) {
      Logger.log('[searchItemsJson_] error body=' + response.getContentText().substring(0, 500));
      break;
    }

    var responseText = response.getContentText();
    Logger.log('[searchItemsJson_] raw response (first 500)=' + responseText.substring(0, 500));

    var json = JSON.parse(responseText);
    var numFound = json.numFound || 0;
    Logger.log('[searchItemsJson_] numFound=' + numFound);

    // ★ パース: results[].item からデータを取得
    Logger.log('[searchItemsJson_] json.results exists=' + !!json.results);
    Logger.log('[searchItemsJson_] json.results.length=' + (json.results ? json.results.length : 0));
    if (json.results && json.results.length > 0) {
      Logger.log('[searchItemsJson_] first result keys=' + Object.keys(json.results[0]).join(','));
      Logger.log('[searchItemsJson_] first result.item exists=' + !!json.results[0].item);
      if (json.results[0].item) {
        Logger.log('[searchItemsJson_] first item keys=' + Object.keys(json.results[0].item).join(','));
        Logger.log('[searchItemsJson_] first item.manageNumber=' + json.results[0].item.manageNumber);
        Logger.log('[searchItemsJson_] first item.variants type=' + typeof json.results[0].item.variants);
        Logger.log('[searchItemsJson_] first item.variants=' + JSON.stringify(json.results[0].item.variants).substring(0, 200));
      }
      for (var i = 0; i < json.results.length; i++) {
        var item = json.results[i].item;
        if (!item) { Logger.log('[searchItemsJson_] result[' + i + '].item is null/undefined'); continue; }
        var variants = item.variants || {};
        var variantKeys = Object.keys(variants);
        Logger.log('[searchItemsJson_] item[' + i + '] manageNumber=' + item.manageNumber + ' variantKeys=' + variantKeys.length);
        if (variantKeys.length > 0) {
          for (var j = 0; j < variantKeys.length; j++) {
            var vid = variantKeys[j];
            var vdata = variants[vid] || {};
            results.push({
              manageNumber: item.manageNumber || '',
              itemNumber: item.itemNumber || '',
              variantId: vid,
              skuId: vdata.merchantDefinedSkuId || '',
              title: item.title || '',
              tagline: item.tagline || ''
            });
          }
        } else {
          // variants がない商品もそのまま追加
          results.push({
            manageNumber: item.manageNumber || '',
            itemNumber: item.itemNumber || '',
            variantId: '',
            skuId: '',
            title: item.title || '',
            tagline: item.tagline || ''
          });
        }
      }
      Logger.log('[searchItemsJson_] parsed ' + json.results.length + ' items, total=' + results.length);
    }

    // ページネーション: cursorMark
    var nextCursorMark = json.nextCursorMark || '';
    if (nextCursorMark && nextCursorMark !== cursorMark) {
      cursorMark = nextCursorMark;
    } else {
      cursorMark = '';
    }

  } while (cursorMark && results.length < 1000);

  Logger.log('[searchItemsJson_] total results=' + results.length);
  return { items: results, numFound: results.length };
}

/**
 * 商品の在庫関連設定（納期設定含む）を取得
 * @param {string} manageNumber - 商品管理番号
 * @param {string} authHeader - ESA認証ヘッダー
 * @returns {object|null} 設定データ
 */
function getInventorySettings_(manageNumber, authHeader) {
  var url = 'https://api.rms.rakuten.co.jp/es/2.0/items/inventory-related-settings/manage-numbers/' +
    encodeURIComponent(manageNumber);

  Logger.log('[getInventorySettings_] url=' + url);

  try {
    var response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { 'Authorization': authHeader },
      muteHttpExceptions: true
    });
    var status = response.getResponseCode();
    Logger.log('[getInventorySettings_] status=' + status);

    if (status === 200) {
      var data = JSON.parse(response.getContentText());
      Logger.log('[getInventorySettings_] variants keys=' + Object.keys(data.variants || {}).join(','));
      return data;
    } else if (status === 429) {
      Logger.log('[getInventorySettings_] 429 rate limited');
      return 429;
    } else {
      Logger.log('[getInventorySettings_] error: ' + response.getContentText().substring(0, 300));
      return null;
    }
  } catch (e) {
    Logger.log('[getInventorySettings_] exception: ' + e.message);
    return null;
  }
}

/**
 * items配列の各要素を更新
 * { manageNumber, variantId, normalDeliveryDateId, backOrderDeliveryDateId }
 * または旧形式 { manageNumber, variantId, leadTimeId }
 * GET で現在の設定を取得し、指定フィールドだけ差し替えて PUT する
 */
function updateLeadTimeJson_(items, authHeader) {
  var results = [];

  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    var manageNumber = item.manageNumber;
    var variantId = item.variantId;

    Logger.log('[updateLeadTimeJson_] processing ' + manageNumber + ' variant=' + variantId);

    // Step 1: 現在の設定を取得
    var currentSettings = getInventorySettings_(manageNumber, authHeader);

    if (!currentSettings) {
      Logger.log('[updateLeadTimeJson_] failed to get current settings for ' + manageNumber);
      results.push({
        manageNumber: manageNumber,
        variantId: variantId,
        success: false,
        status: 'GET失敗'
      });
      Utilities.sleep(1500);
      continue;
    }

    // Step 2: 指定フィールドを変更
    var variants = currentSettings.variants || {};
    // 新形式: normalDeliveryDateId / backOrderDeliveryDateId
    // 旧形式互換: leadTimeId → normalDeliveryDateId
    var newNormal = item.normalDeliveryDateId != null ? item.normalDeliveryDateId : (item.leadTimeId != null ? item.leadTimeId : null);
    var newBackOrder = item.backOrderDeliveryDateId != null ? item.backOrderDeliveryDateId : null;

    if (variantId && variants[variantId]) {
      if (newNormal != null) variants[variantId].normalDeliveryDateId = Number(newNormal);
      if (newBackOrder != null) variants[variantId].backOrderDeliveryDateId = Number(newBackOrder);
      Logger.log('[updateLeadTimeJson_] updated variant ' + variantId +
        ' normal=' + variants[variantId].normalDeliveryDateId +
        ' backOrder=' + (variants[variantId].backOrderDeliveryDateId || 'none'));
    } else {
      var variantKeys = Object.keys(variants);
      for (var j = 0; j < variantKeys.length; j++) {
        if (newNormal != null) variants[variantKeys[j]].normalDeliveryDateId = Number(newNormal);
        if (newBackOrder != null) variants[variantKeys[j]].backOrderDeliveryDateId = Number(newBackOrder);
      }
    }

    // Step 3: 全データを PUT で送信
    var payload = {
      unlimitedInventoryFlag: currentSettings.unlimitedInventoryFlag,
      features: currentSettings.features,
      variants: variants
    };

    var url = 'https://api.rms.rakuten.co.jp/es/2.0/items/inventory-related-settings/manage-numbers/' +
      encodeURIComponent(manageNumber);

    Logger.log('[updateLeadTimeJson_] PUT url=' + url);
    Logger.log('[updateLeadTimeJson_] payload variants count=' + Object.keys(variants).length);

    try {
      var response = UrlFetchApp.fetch(url, {
        method: 'put',
        headers: {
          'Authorization': authHeader,
          'Content-Type': 'application/json'
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });

      var code = response.getResponseCode();
      var body = response.getContentText();
      Logger.log('[updateLeadTimeJson_] status=' + code + ' body=' + body.substring(0, 200));

      results.push({
        manageNumber: manageNumber,
        variantId: variantId,
        success: (code === 200 || code === 204),
        status: code
      });
    } catch (e) {
      Logger.log('[updateLeadTimeJson_] exception: ' + e.message);
      results.push({
        manageNumber: manageNumber,
        variantId: variantId,
        success: false,
        status: e.message
      });
    }

    Utilities.sleep(1500);
  }

  return { results: results };
}
