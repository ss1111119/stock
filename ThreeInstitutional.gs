function fetchDataAndWriteToSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("三大法人統計");

  if (!sheet) {
    console.error("找不到名稱為 '三大法人統計' 的工作表");
    return;
  }

  var existingDates = []; // 定义 existingDates 变量

  var currentDate = new Date();
  var formattedStartDate = Utilities.formatDate(currentDate, "GMT+8", "yyyyMMdd");
  Logger.log("正在取得日期 " + formattedStartDate + " 的數據");

  // 清空 existingDates 数组
  existingDates.length = 0;

  // 重新获取 existingDates 数据
  existingDates = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues().flat();

  Logger.log("existingDates: " + existingDates);
  Logger.log("formattedStartDate: " + formattedStartDate);

  if (existingDates.indexOf(formattedStartDate) !== -1) {
    Logger.log("日期 " + formattedStartDate + " 的數據已存在，不再寫入。");
    return; // 立即停止执行
  }


  Logger.log("继续执行其他操作，因为日期 " + formattedStartDate + " 的數據不存在。");


  var url = 'https://www.twse.com.tw/rwd/zh/fund/BFI82U';
  var payload = {
    'response': 'json',
    'dayDate': formattedStartDate
  };

  var response = UrlFetchApp.fetch(url, { 'method': 'get', 'muteHttpExceptions': true, 'payload': payload });
  var content = response.getContentText();

  var jsonData = JSON.parse(content);

  if (jsonData['stat'] === 'OK') {
    var data = [
      jsonData['date'],
      jsonData['data'][0][1], jsonData['data'][0][2],
      jsonData['data'][1][1], jsonData['data'][1][2],
      jsonData['data'][2][1], jsonData['data'][2][2],
      jsonData['data'][3][1], jsonData['data'][3][2],
      jsonData['data'][4][1], jsonData['data'][4][2]
    ];

    sheet.appendRow(data);
    Logger.log("日期 " + formattedStartDate + " 的數據已成功寫入。");
  } else {
    Logger.log("無法取得日期 " + formattedStartDate + " 的數據。");
  }
}

function runDataFetchLoop() {
  fetchDataAndWriteToSheet();
}
