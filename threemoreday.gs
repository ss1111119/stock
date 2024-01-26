function fetchDataForDateRange(startDate, endDate) {
  // 獲取名為 "三大法人每日" 的工作表
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("三大法人每日");
  if (!sheet) {
    Logger.log("找不到名為 '三大法人每日' 的工作表");
    return;
  }

  // 設置起始和結束日期
  var start = new Date(startDate);
  var end = new Date(endDate);

  // 遍歷起始和結束日期之間的每一天
  for (var d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) {
    var dateString = Utilities.formatDate(d, "GMT+8", "yyyyMMdd");
    Logger.log("正在取得 " + dateString + " 的資料");
    // 檢查該日期的資料是否已經存在於工作表中
    if (!isDataExists(sheet, dateString)) {
      fetchDataAndWriteToSheet(dateString, sheet);
    }
    Utilities.sleep(10000); // 每次請求之間暫停 10 秒以避免超過 API 限制
  }
}

function fetchDataAndWriteToSheet(dateString, sheet) {
  // 設定請求的 URL 和日期
  var url = 'https://www.twse.com.tw/rwd/zh/fund/BFI82U?response=json&dayDate=' + dateString;

  try {
    // 發送請求並獲取回應
    var response = UrlFetchApp.fetch(url);
    var content = response.getContentText();
    var jsonData = JSON.parse(content);

    // 檢查是否成功獲取到數據
    if (jsonData['stat'] === 'OK') {
      var data = [
        jsonData['date'],
        jsonData['data'][0][1], jsonData['data'][0][2],
        jsonData['data'][1][1], jsonData['data'][1][2],
        jsonData['data'][2][1], jsonData['data'][2][2],
        jsonData['data'][3][1], jsonData['data'][3][2],
        jsonData['data'][4][1], jsonData['data'][4][2]
      ];

      // 將數據寫入工作表
      sheet.appendRow(data);
      Logger.log("日期 " + dateString + " 的數據已成功寫入");
    } else {
      Logger.log("未能成功獲取日期 " + dateString + " 的數據");
    }
  } catch (error) {
    Logger.log("請求出現錯誤: " + error.toString());
  }
}

function isDataExists(sheet, dateString) {
  // 獲取工作表中的日期列
  var dateColumn = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat();
  // 檢查日期是否存在於日期列中
  return dateColumn.some(function(date) {
    return Utilities.formatDate(new Date(date), "GMT+8", "yyyyMMdd") === dateString;
  });
}

function testFetchDataForDateRange() {
  // 設定想要獲取資料的日期範圍
  var startDate = "2023-01-01"; // 起始日期
  var endDate = "2023-01-31"; // 結束日期
  // 呼叫函數並傳入起始和結束日期
  fetchDataForDateRange(startDate, endDate);
}
