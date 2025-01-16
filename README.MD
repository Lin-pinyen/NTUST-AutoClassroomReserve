# 欲使用classroom_crawler來進行場地預約有兩個步驟要執行

1. 打開classroom_crawler資料夾中的setting.txt，可能會看到如以下畫面

{
    "username": "Your email",
    "password": "Your password",
    "target_room_url": "https://yo-1.ct.ntust.edu.tw/rms/loans/roomReserve/77",
    "reason": "營建工程整合設計實務",
    "weekday": 1,
    "desired_months": ["February","March", "April", "May", "June"],
    "start_time_hour": 15,
    "start_time_minute": 30,
    "end_time_hour": 17,
    "end_time_minute": 0,
    "start_date": [2025, 2, 17],
    "max_days_to_borrow": 16
}

username: 您的帳號

password: 您的密碼

target_room_url: 這邊預設為T206未來教室，可以自行更換網址

reason: 借用理由

weekday: 借用的日期是禮拜幾(0: 星期日，1:星期一,...6: 星期六)

desired_months: 這邊要輸入會用到的月份名稱，如我們要從2月借到6月，就要向上方範例的寫法

desired_months: ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]

start_time_hour: 教室借用起始的小時

start_time_minute: 教室借用的起始分鐘

end_time_hour: 教室借用結束的小時

end_time_minute: 教室借用結束分鐘

start_date : 要借的第一天是幾年幾月幾日

max_days_to_borrow: 要連續借幾週

2. 在./classroom_crawler 目錄下輸入 python main.py


## 補充:

- 記得要去看egde driver的版本是否與你的電腦版本相同

- 這次最難實作的功能為set_time_slider，選完時間後要記得模擬滑鼠操作才能達到選擇時間的效果
