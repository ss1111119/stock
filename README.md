# 三大法人交易日統計
1.爬取台灣證交所三大法人買賣金額統計。

2.將資料存在google sheet。

3.為避免占用空間，所以紀錄計算，改使用程式即時運算。

4.目前為當日抓取。

5.利用google apps script來執行。

6.可以配合使用google apps script定時觸發裝置，來每日執行。

# Line 通知
1.將每日爬取台灣證交所三大法人買賣金額統計資料後，透過Line notify通知。

## 已知BUG
1.無法判斷資料是否已存在sheet中。

2.調教chatgpt已修正無法判斷資料是否存在sheet中。


