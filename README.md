# 三大法人交易日統計(Apps Script)
1.爬取台灣證交所三大法人買賣金額統計。

2.將資料存在google sheet。

3.為避免占用空間，所以紀錄計算，改使用程式即時運算。

4.目前為當日抓取。

5.利用google apps script來執行。

6.可以配合使用google apps script定時觸發裝置，來每日執行。

7.增加判斷113年股票休市日期，如果休市就不抓取。

# Line 通知
1.將每日爬取台灣證交所三大法人買賣金額統計資料後，透過Line notify通知。

# 三大法人交易日統計歷史抓取
1.可以自動抓取三大法人每日交易量。

2.可以設定抓取區間(建議請設定一個月)，apps scrip只能執行6分鐘。

3.避免伺服器阻擋，有每10秒才抓取一次延遲。

## 已知BUG
1.無法判斷資料是否已存在sheet中。

2.調教chatgpt已修正無法判斷資料是否存在sheet中。


