📊 Field Sales Visit Report Project
本專案旨在自動化處理並產出海外事業處業務報表，整合來自 API 與內部系統（如飛騰）的多項資料來源，透過 Python 程式處理拜訪紀錄、客戶分類、專案時數與請假資訊，並以彙整報表形式輸出，協助管理者掌握團隊外勤狀況與資源分配。

Field_Sales_Visit_Report_Project/
│
├── tw_token.py                 # 存取 API Token，讀取環境變數 (.env)
├── README.md                   # 專案說明文件
├── TW_trackingrecord.py        # 處理台灣業務的客戶追蹤紀錄
├── TW_exhibition.py            # 展廳接待資料整理
├── Customer_Category.py        # 客戶類型的標籤分類與清洗
├── Sales_Visit_Pivot_Table.py  # 外勤拜訪紀錄的樞紐分析（按地區/客戶類別統計）
├── Customer_Relationship.py    # 客戶關係人（業務負責人）清單處理
├── Project_Hours.py            # 專案時數申請與統計
├── Hrs_day_off.py              # 請假記錄處理（整合飛騰請假資料）
├── Company_Task.py             # 公司內部交辦事項整合
└── Multiple_Visit_List.py      # 處理多位人員拜訪同一客戶之清單

🚀 功能特色
📌 自動化 API 存取與資料合併處理

📅 彙整每日外勤活動、請假與展廳接待記錄

🧩 客戶標籤分類、客戶負責人對應

📊 匯出可視化報表與地區別統計分析

✅ 支援資料稽核與多人訪客清單處理