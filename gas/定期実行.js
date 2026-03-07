function scheduledLeadTimeUpdater_v4() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const execSheet = ss.getSheetByName('定期実行');
  const leadTimeSheet = ss.getSheetByName('リードタイム');
  const logSheet = ss.getSheetByName('LT実行ログ') || ss.insertSheet('LT実行ログ');

  const serviceSecret = PropertiesService.getScriptProperties().getProperty('serviceSecret');
  const licenseKey = PropertiesService.getScriptProperties().getProperty('licenseKey');
  const authHeader = 'ESA ' + Utilities.base64Encode(serviceSecret + ':' + licenseKey);

  // リードタイム名称 → ID マップ
  const leadTimeData = leadTimeSheet.getDataRange().getValues();
  const nameToIdMap = {};
  for (let i = 1; i < leadTimeData.length; i++) {
    nameToIdMap[leadTimeData[i][1]] = leadTimeData[i][0];
  }

  // 実行タイミングでリードタイムを自動決定
  const now = new Date();
  const day = now.getDay();
  const hour = now.getHours();
  const month = now.getMonth() + 1;  // 1-12
  const date = now.getDate();

  let targetLeadTimeName = null;

  // ===== 年末年始の特別設定（優先判定） =====
  if (month === 12 && date === 28 && hour === 15) {
    targetLeadTimeName = '出荷リードタイム10日';
  } else if (month === 1 && date === 2 && hour === 13) {
    targetLeadTimeName = '出荷リードタイム5日';
  } else if (month === 1 && date === 4 && hour === 13) {
    targetLeadTimeName = '出荷リードタイム3日';
  }
  // ===== 通常の週次設定 =====
  else if (day === 4 && hour === 13) {
    targetLeadTimeName = '出荷リードタイム5日';
  } else if (day === 0 && hour === 13) {
    targetLeadTimeName = '出荷リードタイム3日';
  } else {
    Logger.log('曜日・時間外のため終了');
    return;
  }

  const targetLeadTimeId = nameToIdMap[targetLeadTimeName];
  if (!targetLeadTimeId) {
    Logger.log(`リードタイム「${targetLeadTimeName}」のIDが見つかりません`);
    return;
  }

  const execData = execSheet.getRange(2, 1, execSheet.getLastRow() - 1, 2).getValues();
  const timestamp = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
  const logs = [];

  for (let i = 0; i < execData.length; i++) {
    const manageNumber = execData[i][0];
    const flag = execData[i][1];

    if (flag !== 1 && flag !== '1') continue;

    try {
      const skuUrl = `https://api.rms.rakuten.co.jp/es/2.0/items/manage-numbers/${manageNumber}`;
      const skuOptions = {
        method: 'get',
        headers: {
          'Authorization': authHeader,
          'Content-Type': 'application/json'
        },
        muteHttpExceptions: true
      };
      const skuResp = UrlFetchApp.fetch(skuUrl, skuOptions);
      const skuJson = JSON.parse(skuResp.getContentText());

      const variantKeys = Object.keys(skuJson.variants || {});
      if (variantKeys.length === 0) {
        logs.push([timestamp, manageNumber, targetLeadTimeName, 'SKUが取得できません']);
        continue;
      }

      const variantId = variantKeys[0];

      const inventoryUrl = `https://api.rms.rakuten.co.jp/es/2.1/inventories/manage-numbers/${manageNumber}/variants/${variantId}`;
      const invResp = UrlFetchApp.fetch(inventoryUrl, skuOptions);
      const invJson = JSON.parse(invResp.getContentText());

      if (!('quantity' in invJson)) {
        logs.push([timestamp, manageNumber, targetLeadTimeName, '在庫数が取得できません']);
        continue;
      }

      const quantity = invJson.quantity;

      const putUrl = `https://api.rms.rakuten.co.jp/es/2.1/inventories/manage-numbers/${manageNumber}/variants/${variantId}`;
      const payload = {
        mode: 'ABSOLUTE',
        quantity: quantity,
        operationLeadTime: {
          normalDeliveryTimeId: targetLeadTimeId
        }
      };

      const putOptions = {
        method: 'put',
        headers: {
          'Authorization': authHeader,
          'Content-Type': 'application/json'
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };

      const putResp = UrlFetchApp.fetch(putUrl, putOptions);
      const code = putResp.getResponseCode();
      if (code === 204) {
        logs.push([timestamp, manageNumber, targetLeadTimeName, `成功（在庫数: ${quantity}）`]);
      } else {
        logs.push([timestamp, manageNumber, targetLeadTimeName, `API失敗(${code}): ${putResp.getContentText()}`]);
      }

      Utilities.sleep(1500);

    } catch (e) {
      logs.push([timestamp, manageNumber, targetLeadTimeName, `例外: ${e.message}`]);
      Utilities.sleep(1500);
    }
  }

  if (logSheet.getLastRow() === 0) {
    logSheet.appendRow(['日時', '商品管理番号', '設定リードタイム', '結果']);
  }
  logSheet.getRange(logSheet.getLastRow() + 1, 1, logs.length, logs[0].length).setValues(logs);

  MailApp.sendEmail({
    to: 'tokyoflowercoltd@gmail.com',
    subject: `リードタイム更新結果（${timestamp}）`,
    body: logs.map(row => row.join(' | ')).join('\n')
  });
}