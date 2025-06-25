# 📅 Daily Task 自動化任務說明

本專案旨在實現每日 CRM 數據自動化處理任務，透過 Python 腳本串接 CRM API 並進行資料清洗、轉換與自動提交流程。整體流程涵蓋從權限驗證、資料取得、ETL 處理到審批提交，全程自動執行，大幅減少人工作業時間與錯誤率。

## 📌 核心功能

1. **獲取 CRM Token（OAuth 認證）**
   - 通過 RESTful API 登入系統，動態取得 Access Token。
   - 自動刷新機制確保長時間運行不中斷。

2. **串接各項 CRM 表單 API**
   - 包含：客戶基本資料、聯絡人、拜訪紀錄、業務審批流程等。
   - 利用 `requests` 模組進行多表同步，支援分頁與異常重試機制。

3. **ETL 處理（Extract → Transform → Load）**
   - Extract：擷取多來源 JSON 資料。
   - Transform：清洗欄位、格式統一、時間戳轉換、邏輯分類（如 K大客戶標註）。
   - Load：儲存為本地 CSV/Excel，或同步更新至資料庫。

4. **自動化批次審批提交**
   - 批次送出每日待處理紀錄至 CRM 系統，觸發審批流程。
   - 實作大量寫入與例外處理邏輯，保證資料一致性與穩定性。

5. **定時排程（Scheduler）**
   - 結合 Windows 工作排程 / crontab / `schedule` 套件每日自動執行。
   - 可搭配郵件發送結果摘要與錯誤日誌。

## 🚀 專案效益

- 節省每日約 **2~3 小時** 重複性人工作業
- 提高數據處理正確性與一致性
- 實現「無人介入」的每日 CRM 數據流通與審批
- 強化團隊決策依據的即時性與透明度

## 🔧 使用技術

- Python 3.x
- `requests`, `pandas`, `schedule`, `json`, `openpyxl`
- CRM API（OAuth 2.0 認證）
- Excel / CSV 匯出
- Windows Task Scheduler or Linux crontab

## 📂 專案結構

```bash
daily_task/
├── auth.py           # Token 授權處理
├── fetch_data.py     # 各表單 API 抓取模組
├── transform.py      # 資料轉換與清洗
├── submit.py         # 批次審批任務提交
├── schedule_task.py  # 定時執行主程式
├── config.json       # API 金鑰與參數設定
├── README.md         # 專案說明文件
