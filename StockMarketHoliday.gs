var HOLIDAYS = [
  '2024-01-01', // 格式YYYY-MM-DD 113年休市日期
  '2024-02-03',
  '2024-02-04',
  '2024-02-05',
  '2024-02-07',
  '2024-02-08',
  '2024-02-09',
  '2024-02-10',
  '2024-02-11',
  '2024-02-12',
  '2024-02-13',
  '2024-02-14',
  '2024-02-17',
  '2024-02-18',
  '2024-02-24',
  '2024-02-25',
  '2024-02-28',
  '2024-03-02',
  '2024-03-03',
  '2024-03-09',
  '2024-03-10',
  '2024-03-16',
  '2024-03-17',
  '2024-03-23',
  '2024-03-24',
  '2024-03-30',
  '2024-03-31',
  '2024-04-04',
  '2024-04-05',
  '2024-04-06',
  '2024-04-07',
  '2024-04-13',
  '2024-04-14',
  '2024-04-20',
  '2024-04-21',
  '2024-04-27',
  '2024-04-28',
  '2024-05-01',
  '2024-05-04',
  '2024-05-05',
  '2024-05-11',
  '2024-05-12',
  '2024-05-18',
  '2024-05-19',
  '2024-05-25',
  '2024-05-26',
  '2024-06-01',
  '2024-06-02',
  '2024-06-08',
  '2024-06-09',
  '2024-06-10',
  '2024-06-15',
  '2024-06-16',
  '2024-06-22',
  '2024-06-23',
  '2024-06-29',
  '2024-06-30',
  '2024-07-06',
  '2024-07-07',
  '2024-07-13',
  '2024-07-14',
  '2024-07-20',
  '2024-07-21',
  '2024-07-27',
  '2024-07-28',
  '2024-08-03',
  '2024-08-04',
  '2024-08-10',
  '2024-08-11',
  '2024-08-17',
  '2024-08-18',
  '2024-08-24',
  '2024-08-25',
  '2024-08-31',
  '2024-09-01',
  '2024-09-07',
  '2024-09-08',
  '2024-09-14',
  '2024-09-15',
  '2024-09-17',
  '2024-09-21',
  '2024-09-22',
  '2024-09-28',
  '2024-09-29',
  '2024-10-05',
  '2024-10-06',
  '2024-10-10',
  '2024-10-12',
  '2024-10-13',
  '2024-10-19',
  '2024-10-20',
  '2024-10-26',
  '2024-10-27',
  '2024-11-02',
  '2024-11-03',
  '2024-11-09',
  '2024-11-10',
  '2024-11-16',
  '2024-11-17',
  '2024-11-23',
  '2024-11-24',
  '2024-11-30',
  '2024-12-01',
  '2024-12-07',
  '2024-12-08',
  '2024-12-14',
  '2024-12-15',
  '2024-12-21',
  '2024-12-22',
  '2024-12-28',
  '2024-12-29',
  // ...添加其他假日...
];

function isHoliday(date) {
  // 格式化日期為YYYY-MM-DD
  console.log(isHoliday);
  var dateString = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  // 檢查今天是否是假日
  return HOLIDAYS.indexOf(dateString) > -1;
}

function myFunction() {
  var today = new Date();
  var testDate = Utilities.formatDate(today, "GMT+8", "yyyy-MM-dd");
  if (!isHoliday(today)) {
    fetchDataAndWriteToSheet(testDate);  // 傳遞有效的日期字符串
    console.log('執行了任務，因為今天不是假日。');
  } else {
    console.log('今天是休市日，跳過任務。');
  }
}

function createDailyTrigger() {
  // 創建一個每天下午3點30分左右運行的觸發器
  ScriptApp.newTrigger('myFunction')
      .timeBased()
      .everyDays(1)
      .atHour(15) // 設定為下午3點
      .nearMinute(30) // 接近30分
      .create();
}