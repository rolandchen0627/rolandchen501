# Chiayi_Apply_Sanitized.py
import pandas as pd
import json
import requests
import time
from concurrent.futures import ThreadPoolExecutor
import numpy as np
from datetime import datetime, timedelta
import os
from workalendar.asia import Taiwan
import sys
from dotenv import load_dotenv

# 載入 .env 環境變數
load_dotenv()

# 初始化行事曆
cal = Taiwan()
current_year = datetime.now().year
holidays_year = cal.holidays(current_year)
holiday_dates = [d for d, _ in holidays_year]

today = datetime.now().date()
if today in holiday_dates:
    print(f"今天是假日: {today}")
    sys.exit(0)

# 取得 access token
url_token = os.getenv("API_TOKEN_URL")
payload = {
    "grant_type": "password",
    "client_id": os.getenv("CLIENT_ID"),
    "client_secret": os.getenv("CLIENT_SECRET"),
    "redirect_uri": os.getenv("REDIRECT_URI"),
    "username": os.getenv("USERNAME"),
    "password": os.getenv("PASSWORD")
}
response = requests.post(url_token, data=payload)
content = response.json()
ac_token = content["access_token"]

# 查詢 userId 對應 token
headers = {"Authorization": f"Bearer {ac_token}", "Content-Type":"application/x-www-form-urlencoded"}
data = {"xoql": "select id, name as DisplayName from user", "batchCount": 2000}
response = requests.post(os.getenv("API_QUERY_URL"), headers=headers, data=data)
user_df = pd.json_normalize(response.json()["data"]["records"])
user_df = user_df[user_df['DisplayName'] == os.getenv("TARGET_USER")]

# 取得委任 Token
base_url = os.getenv("API_DELEGATE_TOKEN_URL")

def fetch_data(assigneeId):
    url = f"{base_url}{assigneeId}"
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        result = response.json().get("result")
        if result:
            return json.dumps({"assigneeId": assigneeId, "access_token": result["access_token"]})
    return None

with ThreadPoolExecutor(2) as executor:
    results = list(executor.map(fetch_data, user_df['id']))

valid_results = [json.loads(r) for r in results if r]
Tasks_df3 = pd.DataFrame(valid_results)
ac_token = Tasks_df3.iloc[0]['access_token']

# 此處開始處理撿貨申請的邏輯流程

# 判斷下一個工作日

def get_next_workday(start_date):
    while cal.is_holiday(start_date) or start_date.weekday() >= 5:
        start_date += timedelta(days=1)
    return start_date

next_date = datetime.now() + timedelta(days=1)
next_workday = get_next_workday(next_date)
next_str = next_workday.strftime('%y%m%d')
today_str = datetime.now().strftime('%y%m%d')

# 以下略過實際查詢資料，改為結構示意：
# 假設我們取得了 TW 與 OV 當日與次日的最大值
max_tw_num = 123
max_tw_next = 45
max_ov_num = 67
max_ov_next = 89

# 示意：分配流水碼格式
# TWyyyyMMdd-xxx 或 OSyyyyMMdd-xxx

def assign_sequence_code(prefix, date_str, start_seq, length):
    return [f"{prefix}{date_str}-{i:03}" for i in range(start_seq + 1, start_seq + 1 + length)]

# 示意資料框（請依實際應用邏輯產出）
tw_sm = pd.DataFrame({'id': ['a1', 'a2']})
tw_big = pd.DataFrame({'id': ['a3']})
ov_sm = pd.DataFrame({'id': ['b1']})
ov_big = pd.DataFrame({'id': ['b2', 'b3']})

# 分配代碼欄位
sequence_tw_sm = assign_sequence_code('TW', today_str, max_tw_num, len(tw_sm))
tw_sm['customItem156__c'] = sequence_tw_sm

sequence_tw_big = assign_sequence_code('TW', next_str, max_tw_next, len(tw_big))
tw_big['customItem156__c'] = sequence_tw_big

sequence_ov_sm = assign_sequence_code('OS', today_str, max_ov_num, len(ov_sm))
ov_sm['customItem156__c'] = sequence_ov_sm

sequence_ov_big = assign_sequence_code('OS', next_str, max_ov_next, len(ov_big))
ov_big['customItem156__c'] = sequence_ov_big

# 合併資料後上傳
final_df = pd.concat([tw_sm, tw_big, ov_sm, ov_big])
final_df['customItem157__c'] = 'mock_time_value'  # 示例欄位

# 建立批次任務（示意）
bulk_job_url = os.getenv("API_BULK_JOB_URL")
bulk_batch_url = os.getenv("API_BULK_BATCH_URL")

headers = {"Authorization": f"Bearer {ac_token}", "Content-Type":"application/json"}
job_data = {
    "data": {
        "operation": "update",
        "object": "customEntity25__c"
    }
}
response = requests.post(bulk_job_url, headers=headers, json=job_data)
bulk_id = response.json()["result"]["id"]

# 批次分批上傳
batch_size = 5000
for i in range(0, len(final_df), batch_size):
    batch_df = final_df.iloc[i:i+batch_size]
    batch_data = {
        "data": {
            "jobId": bulk_id,
            "datas": batch_df.to_dict(orient='records')
        }
    }
    requests.post(bulk_batch_url, headers=headers, json=batch_data)

print("✅ 撿貨申請資料已成功處理並送出")
