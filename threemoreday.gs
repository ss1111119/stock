function getTWSEDataForOneYear() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var columns = ['日期', '自營商自行買賣_買進金額', '自營商自行買賣_賣出金額', '自營商避險_買進金額', '自營商避險_賣出金額', '投信_買進金額', '投信_賣出金額', '外資及陸資不含外資自營商_買進金額', '外資及陸資不含外資自營商_賣出金額', '外資自營商_買進金額', '外資自營商_賣出金額'];
  sheet.getRange(1, 1, 1, columns.length).setValues([columns]);

  var startDate = new Date();
  startDate.setDate(startDate.getDate() - 365); // 一年前的日期

  var currentDate = new Date();
  var data = [];

  while (currentDate >= startDate) {
    var formattedDate = Utilities.formatDate(currentDate, "GMT+8", "yyyyMMdd");
    Logger.log(formattedDate);
    var rowData = getTWSEData(formattedDate);
    if (rowData.length > 0) {
      data.push(rowData);
      Utilities.sleep(10000); // 每次請求之間休息10秒
    }
    currentDate.setDate(currentDate.getDate() - 1);
  }

  // 將抓取的資料寫入工作表
  var lastRow = sheet.getLastRow();
  if (data.length > 0) {
    sheet.getRange(lastRow + 1, 1, data.length, data[0].length).setValues(data);
  }
}

function getTWSEData(dayDate) {
  var url = 'https://www.twse.com.tw/rwd/zh/fund/BFI82U';
  var payload = {
    'response': 'json',
    'dayDate': dayDate
  };
  var response = UrlFetchApp.fetch(url, { 'method': 'get', 'muteHttpExceptions': true, 'payload': payload });
  var content = response.getContentText();
  var jsonData = JSON.parse(content);

  if (jsonData['stat'] === 'OK') {
    var data = [
      dayDate,
      jsonData['data'][0][1], jsonData['data'][0][2],
      jsonData['data'][1][1], jsonData['data'][1][2],
      jsonData['data'][2][1], jsonData['data'][2][2],
      jsonData['data'][3][1], jsonData['data'][3][2],
      jsonData['data'][4][1], jsonData['data'][4][2]
    ];
    return data;
  } else {
    return [];
  }
}
