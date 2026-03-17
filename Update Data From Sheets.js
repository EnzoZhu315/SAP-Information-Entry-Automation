function syncSmartsheetToGoogleSheet() {
  // --- 1. 获取配置信息 (从脚本属性中读取) ---
  const scriptProperties = PropertiesService.getScriptProperties();
  
  const SMARTSHEET_TOKEN = scriptProperties.getProperty('SMARTSHEET_TOKEN');
  const SMARTSHEET_ID = scriptProperties.getProperty('SMARTSHEET_ID');
  const TARGET_GS_ID = scriptProperties.getProperty('TARGET_GS_ID');
  const TARGET_SHEET_NAME = scriptProperties.getProperty('TARGET_SHEET_NAME') || 'Sheet1';

  // 安全检查：确保必要配置已存在
  if (!SMARTSHEET_TOKEN || !SMARTSHEET_ID) {
    throw new Error('未发现必要的配置信息，请在“脚本属性”中设置。');
  }

  // --- 2. 调用 Smartsheet API 获取数据 ---
  const url = `https://api.smartsheet.com/2.0/sheets/${SMARTSHEET_ID}`;
  const options = {
    'method': 'get',
    'headers': {
      'Authorization': 'Bearer ' + SMARTSHEET_TOKEN
    },
    'muteHttpExceptions': true
  };

  const response = UrlFetchApp.fetch(url, options);
  if (response.getResponseCode() !== 200) {
    Logger.log('API 请求失败: ' + response.getContentText());
    return;
  }
  
  const data = JSON.parse(response.getContentText());
  
  // 建立列名到列 ID 的映射
  const columnMap = {};
  data.columns.forEach(col => {
    columnMap[col.title] = col.id;
  });

  // --- 3. 筛选逻辑 ---
  const filteredRows = data.rows.filter(row => {
    const cells = row.cells;
    
    // 获取对应列的值
    const procType = getCellValue(cells, columnMap['Procurement Type']);
    const plant = getCellValue(cells, columnMap['Plant/Department']); 
    const done4 = getCellValue(cells, columnMap['Done-4']);

    // 逻辑判定
    return procType === 'Outsource' && 
           String(plant).includes('CN12') && 
           (done4 === false || done4 === null || done4 === "" || done4 === 0);
  });

  // --- 4. 写入 Google Sheet ---
  const ss = SpreadsheetApp.openById(TARGET_GS_ID);
  const sheet = ss.getSheetByName(TARGET_SHEET_NAME);
  
  sheet.clear(); 

  // 写入表头
  const headers = data.columns.map(col => col.title);
  sheet.appendRow(headers);

  // 批量写入筛选后的数据
  if (filteredRows.length > 0) {
    const outputData = filteredRows.map(row => {
      return data.columns.map(col => getCellValue(row.cells, col.id));
    });
    
    sheet.getRange(2, 1, outputData.length, outputData[0].length).setValues(outputData);
    Logger.log('同步成功！共导入 ' + filteredRows.length + ' 条数据。');
  } else {
    Logger.log('没有发现符合条件的数据。');
  }
}

/**
 * 辅助函数：根据列 ID 获取单元格内容
 */
function getCellValue(cells, columnId) {
  const cell = cells.find(c => c.columnId === columnId);
  if (!cell) return "";
  return cell.value !== undefined ? cell.value : "";
}
