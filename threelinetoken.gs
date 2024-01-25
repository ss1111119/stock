var token1 = "YOUR LINE Notify1"; //因為有兩個群組，所以我設定兩個權杖通知
var token2 = "YOUR LINE Notify2";

function calculateDifferencesForSpecificDateAndNotify() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("三大法人統計");
  var range = sheet.getDataRange();
  var values = range.getValues();

  // 自動獲取當天日期並格式化為yyyyMMdd格式
  var currentDate = new Date();
  var year = currentDate.getFullYear();
  var month = (currentDate.getMonth() + 1).toString().padStart(2, '0');
  var day = currentDate.getDate().toString().padStart(2, '0');
  var testDate = year + month + day;

  Logger.log('當天日期：' + testDate); // 記錄當天日期

  var found = false; // 標記是否找到匹配的日期

  // 遍歷數據
  values.forEach(function(row, index) {
    // 跳過標題行
    if (index === 0) return;

    // 檢查日期是否匹配
    var rowDate = row[0].toString();
    if (rowDate === testDate) {
      found = true; // 找到匹配的日期
      // 計算差額
      var diffs = [
        row[1] - row[2], // B - C
        row[3] - row[4], // D - E
        row[5] - row[6], // F - G
        row[7] - row[8], // H - I
        row[9] - row[10] // J - K
      ];

      // 構建消息
      var message = '\n' + testDate + '\n' +
              '自營商自行買賣: ' + formatNumber(diffs[0]) + '\n' +
              '自營商避險: ' + formatNumber(diffs[1]) + '\n' +
              '投信: ' + formatNumber(diffs[2]) + '\n' +
              '外資及陸資不含外資自營商: ' + formatNumber(diffs[3]) + '\n' +
              '外資自營商: ' + formatNumber(diffs[4]);

      function formatNumber(number) {
        var isNegative = false;
        if (number < 0) {
          isNegative = true;
          number = Math.abs(number);
        }
        
        if (number >= 100000000) {
          number = (number / 100000000).toFixed(2) + '億';
        } else if (number >= 10000) {
          number = (number / 10000).toFixed(2) + '萬';
        } else if (number >= 1000) {
          number = (number / 1000).toFixed(2) + '千';
        } else {
          number = number.toString();
        }
        
        if (isNegative) {
          number = '-' + number;
        }
        
        return number;
      }

      // 發送 LINE 通知給兩個不同的 token
      sendLine(message, token1);
      sendLine(message, token2);
    }
  });

  // 如果沒有找到匹配的日期，記錄在日誌中並發送通知
  if (!found) {
    var message = '沒有找到匹配的日期：' + testDate;
    Logger.log(message);
    // 發送 LINE 通知給兩個不同的 token
    sendLine(message, token1);
    sendLine(message, token2);
  }
}

function sendLine(message, token) {
  UrlFetchApp.fetch('https://notify-api.line.me/api/notify', {
    'headers': {
      'Authorization': 'Bearer ' + token,
    },
    'method': 'post',
    'payload': {
      'message': message,
    }
  });
}
