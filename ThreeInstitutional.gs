function fetchDataAndWriteToSheet(testDate) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("三大法人統計");
  if (!sheet) {
    Logger.log("找不到名稱為 '三大法人統計' 的工作表");
    return;
  }

  var currentDate = testDate ? new Date(testDate) : new Date();
  var formattedStartDate = Utilities.formatDate(currentDate, "GMT+8", "yyyyMMdd");
  Logger.log("正在取得日期 " + formattedStartDate + " 的資料");

  var existingDates = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues().flat();
  Logger.log("existingDates (before check): " + existingDates.join(", "));

  var dateExists = false;
  for (var i = 0; i < existingDates.length; i++) {
    Logger.log("比對 " + existingDates[i] + " 與 " + formattedStartDate);
    if (existingDates[i].toString() === formattedStartDate) {
      dateExists = true;
      break;
    }
  }

  if (dateExists) {
    Logger.log("日期 " + formattedStartDate + " 的資料已存在，不再寫入。");
    return;
  }

  Logger.log("繼續執行其他操作，因為日期 " + formattedStartDate + " 的資料不存在。");
  // 準備請求並獲取資料
  var url = 'https://www.twse.com.tw/rwd/zh/fund/BFI82U';
  var payload = {
    'response': 'json',
    'dayDate': formattedStartDate
  };
  Logger.log("請求URL: " + url + "，參數: " + JSON.stringify(payload));

  var response = UrlFetchApp.fetch(url, { 'method': 'get', 'muteHttpExceptions': true, 'payload': payload });
  var content = response.getContentText();
  Logger.log("響應內容: " + content);

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
    Logger.log("日期 " + formattedStartDate + " 的資料已成功寫入。");
  } else {
    Logger.log("無法獲取日期 " + formattedStartDate + " 的資料。");
  }
}


function testFetchDataForSpecificDate() {
  Logger.log("開始測試特定日期20240124的資料抓取");
  fetchDataAndWriteToSheet('2024-01-24');
}
