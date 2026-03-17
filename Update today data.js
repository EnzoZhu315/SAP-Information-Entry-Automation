function syncDailyDataToSubSheet() {
  // --- 1. 配置信息 (根据实际需求修改) ---
  const CONFIG = {
    SOURCE_SHEET_INDEX: 0,              // 源工作表索引（0 为第一个）
    TARGET_SHEET_NAME: "Daily_Update",   // 目标工作表名称
    DATE_COLUMN_NAME: "Date",           // 用于筛选的日期列表头名称
    TIMEZONE: "GMT+8",                  // 时区
    DATE_FORMAT: "yyyy-MM-dd"           // 日期比对格式
  };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheets()[CONFIG.SOURCE_SHEET_INDEX];
  
  // 确保目标子表存在，不存在则创建
  let targetSheet = ss.getSheetByName(CONFIG.TARGET_SHEET_NAME);
  if (!targetSheet) {
    targetSheet = ss.insertSheet(CONFIG.TARGET_SHEET_NAME);
  }

  // --- 2. 获取源数据 ---
  const fullData = sourceSheet.getDataRange().getValues();
  if (fullData.length <= 1) {
    SpreadsheetApp.getUi().alert("源表中没有数据（仅有表头或为空）。");
    return;
  }

  const headers = fullData[0];
  const dateIdx = headers.indexOf(CONFIG.DATE_COLUMN_NAME);
  
  if (dateIdx === -1) {
    SpreadsheetApp.getUi().alert(`未找到列名: "${CONFIG.DATE_COLUMN_NAME}"，请检查配置。`);
    return;
  }

  // 获取今天的日期字符串
  const todayStr = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, CONFIG.DATE_FORMAT);

  // --- 3. 筛选逻辑 ---
  const filteredData = fullData.slice(1).filter(row => {
    let cellValue = row[dateIdx];
    if (!cellValue) return false;

    let datePart = "";
    if (cellValue instanceof Date) {
      datePart = Utilities.formatDate(cellValue, CONFIG.TIMEZONE, CONFIG.DATE_FORMAT);
    } else {
      // 尝试处理字符串格式 (例如 2026-03-17T08:00:00)
      datePart = String(cellValue).split(' ')[0].split('T')[0]; 
    }
    return datePart === todayStr;
  });

  // --- 4. 执行更新 ---
  targetSheet.clear(); // 清空旧数据实现覆盖
  targetSheet.appendRow(headers); // 重新写入表头

  if (filteredData.length > 0) {
    targetSheet.getRange(2, 1, filteredData.length, headers.length).setValues(filteredData);
    Logger.log(`同步成功：已更新 ${filteredData.length} 条记录。`);
    SpreadsheetApp.getUi().alert(`✅ 同步完成！已更新 ${filteredData.length} 条 [${todayStr}] 的数据。`);
  } else {
    Logger.log("未发现今日数据。");
    SpreadsheetApp.getUi().alert(`ℹ️ 未发现日期为 ${todayStr} 的数据。`);
  }
}
