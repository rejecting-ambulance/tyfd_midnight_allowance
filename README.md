# TYFD_深夜危勞性勤務津貼處理器
  
> The restaurant for TYFD firefighters.  
> Let's burn the midnight oil.

## 使用方式
1. 點選 midnightoil.exe，輸入帳戶、查詢年月。
1. 請開始吃著火鍋唱著歌，
2. 程式結束，輸出Excel
  * 深夜食堂 - 修ㄟ味噌湯：列出待確認項(修正)
  * 深夜世堂 - 千層明太子：簽名確認表(簽名)


## 出納清冊(大隊承辦)
* 檔案內有 冰箱、廚房、吧檯 3 個分頁。
* 把下載的內容放到冰箱。
* 不用管廚房在做什麼
* 到吧檯吃飯。  

(是不是跟把大象放進冰箱一樣簡單？)


## 稽核錯誤項目 
* 出勤案件編號相同    (系統紅底)
* 結束時間與下一筆開始時間重疊
* 當筆過長或過短，要確認有無實際出勤(工作紀錄)
  
#### 你喜歡紅蘿蔔嗎？
* If all correct, the carrot will appear.
* 如果全對，你會發現胡蘿蔔喔。


---
# 改動紀錄
### 114/7/7  
1. 圖形化介面(GUI)取代json讀取輸入資料
2. 修正查詢常崩潰的錯誤StaleElementReferenceException，捕捉並自動重試。
3. 還有一些美化，GUI用起來很直觀，每個功能都是血汗啊。

## About me
* AI Assistance : ChatGPT, Claude
* Contact : zhandezhonghenry@gmail.com
  
  
  