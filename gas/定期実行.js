/**
 * 定期実行: スケジュール設定シートに基づき全店舗のLTを自動更新
 * トリガー: 1時間ごと（時間ベースのタイマー）
 * マルチアカウント対応: api_key シートの全有効ユーザーをループ
 */
function scheduledLeadTimeUpdater() {
  const now = new Date();
  const jst = new Date(now.getTime() + 9 * 60 * 60 * 1000);
  const day   = jst.getUTCDay();    // 0=日 〜 6=土
  const hour  = jst.getUTCHours();
  const month = jst.getUTCMonth() + 1;
  const date  = jst.getUTCDate();

  Logger.log('[scheduler] 実行時刻(JST): ' + Utilities.formatDate(jst, 'UTC', 'yyyy-MM-dd HH:mm') +
    ' day=' + day + ' hour=' + hour);

  // ── api_key シートから有効ユーザーを取得 ──
  const apiKeySS = SpreadsheetApp.openById('1iYeV2SbOVoRH8Qjm2d1w5tWmhlE_zcc-yO1tDSLN7Rk');
  const apiKeySheet = apiKeySS.getSheetByName('api_key');
  const apiKeyData = apiKeySheet.getDataRange().getValues();

  // ── スケジュール設定シートを取得 ──
  const schedSS = SpreadsheetApp.openById('1JICZkk2GzcGIt3VWAZHzI1wMKQMXhWLw5qQNVIgamAQ');
  const schedSheet = schedSS.getSheetByName('スケジュール設定');
  const logSheet = schedSS.getSheetByName('LT実行ログ') || schedSS.insertSheet('LT実行ログ');

  if (!schedSheet) {
    Logger.log('[scheduler] スケジュール設定シートが見つかりません');
    return;
  }

  const schedData = schedSheet.getDataRange().getValues();
  // ヘッダー行スキップ: A=sid B=管理番号 C=タイミング種別 D=曜日 E=月 F=日 G=時刻 H=LT名称

  const timestamp = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
  const logs = [];

  // ── ユーザーループ ──
  for (let u = 1; u < apiKeyData.length; u++) {
    const row = apiKeyData[u];
    const sid   = String(row[6]).trim();   // G列: 店舗ID
    const flag  = row[9];                  // J列: flag（1=無効）
    const expiry = row[10];                // K列: expiry

    if (flag == 1) continue;
    if (expiry && new Date(expiry) < now) continue;
    if (!sid) continue;

    // 認証ヘッダー生成
    const licenseKey    = safeBase64Decode_(String(row[2]));  // C列
    const serviceSecret = safeBase64Decode_(String(row[3]));  // D列
    if (!licenseKey || !serviceSecret) continue;
    const authHeader = 'ESA ' + Utilities.base64Encode(serviceSecret + ':' + licenseKey);

    // LT名称→IDマップ（Shop API から取得）
    const ltMap = getLeadTimeMap_(authHeader);

    // このsidに該当するスケジュールを抽出
    for (let s = 1; s < schedData.length; s++) {
      const sc = schedData[s];
      const scSid     = String(sc[0]).trim();  // A: sid
      const manageNum = String(sc[1]).trim();  // B: 管理番号
      const timing    = String(sc[2]).trim();  // C: weekly / yearly
      const scDay     = Number(sc[3]);         // D: 曜日(0-6)
      const scMonth   = Number(sc[4]);         // E: 月
      const scDate    = Number(sc[5]);         // F: 日
      const scHour    = Number(sc[6]);         // G: 時刻(0-23)
      const ltName    = String(sc[7]).trim();  // H: LT名称

      if (scSid !== sid) continue;
      if (!manageNum || !ltName) continue;

      // タイミング判定
      let matched = false;
      if (timing === 'weekly' && day === scDay && hour === scHour) {
        matched = true;
      } else if (timing === 'yearly' && month === scMonth && date === scDate && hour === scHour) {
        matched = true;
      }
      if (!matched) continue;

      const ltId = ltMap[ltName];
      if (!ltId) {
        logs.push([timestamp, sid, manageNum, ltName, 'LT名称が見つかりません']);
        continue;
      }

      // 商品のvariantId取得
      try {
        const skuUrl = 'https://api.rms.rakuten.co.jp/es/2.0/items/manage-numbers/' + manageNum;
        const skuResp = UrlFetchApp.fetch(skuUrl, {
          method: 'get',
          headers: { 'Authorization': authHeader },
          muteHttpExceptions: true
        });
        const skuJson = JSON.parse(skuResp.getContentText());
        const variantKeys = Object.keys(skuJson.variants || {});
        if (variantKeys.length === 0) {
          logs.push([timestamp, sid, manageNum, ltName, 'SKU取得失敗']);
          continue;
        }

        // 全バリエーションに適用
        for (let v = 0; v < variantKeys.length; v++) {
          const variantId = variantKeys[v];
          const invUrl = 'https://api.rms.rakuten.co.jp/es/2.1/inventories/manage-numbers/' +
            manageNum + '/variants/' + variantId;
          const invResp = UrlFetchApp.fetch(invUrl, {
            method: 'get',
            headers: { 'Authorization': authHeader },
            muteHttpExceptions: true
          });
          const invJson = JSON.parse(invResp.getContentText());
          const quantity = ('quantity' in invJson) ? invJson.quantity : 0;

          const putResp = UrlFetchApp.fetch(invUrl, {
            method: 'put',
            headers: { 'Authorization': authHeader, 'Content-Type': 'application/json' },
            payload: JSON.stringify({
              mode: 'ABSOLUTE',
              quantity: quantity,
              operationLeadTime: { normalDeliveryTimeId: Number(ltId) }
            }),
            muteHttpExceptions: true
          });
          const code = putResp.getResponseCode();
          logs.push([timestamp, sid, manageNum + '/' + variantId, ltName,
            (code === 200 || code === 204) ? '成功(在庫:' + quantity + ')' : 'API失敗(' + code + ')']);
          Utilities.sleep(800);
        }
      } catch(e) {
        logs.push([timestamp, sid, manageNum, ltName, '例外: ' + e.message]);
      }
    }
  }

  if (logs.length === 0) {
    Logger.log('[scheduler] 該当スケジュールなし');
    return;
  }

  // ログ書き込み
  if (logSheet.getLastRow() === 0) {
    logSheet.appendRow(['日時', '店舗ID', '管理番号/バリアント', 'LT名称', '結果']);
  }
  logSheet.getRange(logSheet.getLastRow() + 1, 1, logs.length, 5).setValues(logs);

  Logger.log('[scheduler] 完了: ' + logs.length + '件処理');
}

/**
 * Shop APIからLT名称→IDのマップを取得
 */
function getLeadTimeMap_(authHeader) {
  const map = {};
  try {
    const resp = UrlFetchApp.fetch('https://api.rms.rakuten.co.jp/es/1.0/shop/operationLeadTime', {
      method: 'get',
      headers: { 'Authorization': authHeader, 'Content-Type': 'application/xml; charset=UTF-8' },
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) return map;
    const root = XmlService.parse(resp.getContentText()).getRootElement();
    const list = root.getChild('result').getChild('operationLeadTimeList').getChildren('operationLeadTime');
    for (let i = 0; i < list.length; i++) {
      map[list[i].getChildText('name')] = list[i].getChildText('operationLeadTimeId');
    }
  } catch(e) {
    Logger.log('[getLeadTimeMap_] ' + e.message);
  }
  return map;
}
