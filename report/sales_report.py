import pandas as pd
import numpy as np
import requests
from datetime import datetime ,timedelta
from dateutil.relativedelta import relativedelta
import time
import calendar
import json
import os

# # 📁 Field_Sales_Visit_Report_Project/
# │
# ├── tw_token.py                 # 存取 API Token，讀取環境變數 (.env)
# ├── README.md                   # 專案說明文件
# ├── TW_trackingrecord.py        # 台灣追蹤紀錄
# ├── TW_exhibition.py            # 展廳接待
# ├── Customer_Category.py        # 客戶類型分類與標籤
# ├── Sales_Visit_Pivot_Table.py  # 外勤拜訪記錄的樞紐分析（依地區/類別彙整）
# ├── Customer_Relationship.py    # 客戶關係人
# ├── Project_Hours.py            # 專案時數申請
# ├── Hrs_day_off.py              # 請假記錄（飛騰系統）
# ├── Company_Task.py             # 公司交辦
# └── Multiple_Visit_List.py      # 多位拜訪名單

start_time_計時 = time.time()
today = datetime.now()

# 判斷是否為本月第一天
if today.day == 1:
    # 計算上個月的年份與月份
    last_month = today.month - 1 if today.month > 1 else 12
    last_year = today.year if today.month > 1 else   today.year - 1
    
    # 計算上個月第一天與最後一天
    start_date_of_month = datetime(last_year, last_month, 1, 0, 0, 0)
    last_day = calendar.monthrange(last_year, last_month)[1]  # 上個月的天數
    end_date_of_month = datetime(last_year, last_month, last_day, 23, 59, 59)
else:
    # 計算當月第一天與最後一天
    start_date_of_month = datetime(today.year, today.month, 1, 0, 0, 0)
    last_day = calendar.monthrange(today.year, today.month)[1]  # 當月天數
    end_date_of_month = datetime(today.year, today.month, last_day, 23, 59, 59)

# 轉換為毫秒級時間戳
start_timestamp_of_month = int(start_date_of_month.timestamp() * 1000)
end_timestamp_of_month = int(end_date_of_month.timestamp() * 1000)

print("開始日期:", start_date_of_month.strftime("%Y-%m-%d"))
print("結束日期:", end_date_of_month.strftime("%Y-%m-%d"))


time.sleep(5)

'''
get token TW&CN
'''
from tw_token import get_access_token
ac_token = get_access_token()

'''
行事曆
'''
from get_calendarid import get_calendarid 
print("=" * 50)
print(" " * 15 + "📅 Loading.. 集團行事曆 📅")
print("=" * 50)
time.sleep(5)
final_calendarid = get_calendarid()

'''
台灣追蹤記錄
註解：台灣追蹤記錄中的建檔日期近兩個月的資料
'''
from TW_trackingrecord import TW_trackingrecord
print(" " * 15 + "📅 Loading.. 台灣追蹤記錄 📅")
print("=" * 50)
time.sleep(5)
final_tracking = TW_trackingrecord()

'''
台灣展廳接待
預約參訪日當月
'''
from TW_exhibition import TW_exhibition
print(" " * 15 + "📅 Loading.. 台灣展廳接待 📅")
print("=" * 50)
time.sleep(5)
twdata_reception = TW_exhibition()

'''
客戶類別
'''
from Customer_Category import Customer_Category
print('-'*20+'客戶類別'+'-'*20)
final_Customer_Category = Customer_Category()
Excel_客戶類別_全 = final_Customer_Category.copy()
time.sleep(2)

'''
集團外勤業務人員
'''
from Sales_Visit_Pivot_Table import Sales_Visit_Pivot_Table
print(" " * 15 + "📅 Loading.. 外勤業務人員 📅")
print("=" * 50)
time.sleep(5)
final_sale = Sales_Visit_Pivot_Table()
Excel_業務樞紐 = final_sale.copy()

twdata_reception_list = final_sale.loc[final_sale['區域'] == '專案', '人員姓名(繁中)'].unique().tolist()
twdata_reception = twdata_reception[twdata_reception['接待人員'].isin(twdata_reception_list)]
Excel_展廳接待 = twdata_reception.copy()

sales_list = final_sale['用戶'].unique().tolist()
sales_list_tw = final_sale.loc[
    (final_sale['國家'] == 'TW') & (~final_sale['區域'].str.contains('專案', na=False)),
    '用戶'
].unique().tolist()

sales_list_pro = final_sale.loc[
    (final_sale['國家'] == 'TW') & (final_sale['區域'].str.contains('專案', na=False)),
    '用戶'
].unique().tolist()

sales_list_os = final_sale.loc[
    (final_sale['國家'] != 'TW'),
    '用戶'
].unique().tolist()

sales_list_os_chinses = final_sale.loc[
    (final_sale['國家'] != 'TW'),
    '人員姓名(繁中)'
].unique().tolist()
sales_list_os = sales_list_os + sales_list_os_chinses
sales_list = sales_list + sales_list_os_chinses
final_tracking = final_tracking[final_tracking['業務人員姓名'].isin(sales_list)]
Excel_台灣追蹤記錄 = final_tracking.copy()
'''
TW客戶關係聯絡人
'''
from Customer_Relationship import Customer_Relationship
print(" " * 15 + "📅 Loading.. 客戶關係聯絡人 📅")
print("=" * 50)
time.sleep(5)
twdata_customer = Customer_Relationship()
Excel_客戶關係聯絡人 = twdata_customer.copy()
'''
TW專案時數
執行日期當月第一天-當月最後一天
'''
from Project_Hours import Project_Hours
print(" " * 15 + "📅 Loading.. 專案時數 📅")
print("=" * 50)
time.sleep(5)
final_proHours = Project_Hours()
final_proHours = final_proHours[final_proHours['申請人'].isin(sales_list)]

'''
請假明細
開始日期前三個月~當月最後一天
'''
from Hrs_day_off import Day_Off
print(" " * 15 + "📅 Loading.. 請假明細 📅")
print("=" * 50)
time.sleep(5)
final_dayoff = Day_Off()
final_dayoff = final_dayoff[final_dayoff['員工姓名'].isin(sales_list)]
Excel_請假明細 =final_dayoff.copy()
final_proHours['休假追帳不計算'] = ""
# 計算符合 COUNTIFS 條件的次數
counts = []
for _, row in final_proHours.iterrows():
    count = final_dayoff.loc[
        (final_dayoff['員工姓名'] == row['申請人']) &
        (final_dayoff['開始日期'] == row['執行日期']) &
        (final_dayoff['實際請假時數'] == 8)
    ].shape[0]
    counts.append(count)

# 將計算的次數保存回 `final_proHours`
final_proHours['COUNTIFS結果'] = counts
# FIND 部分：判斷 '專案內容說明' 中是否包含 "追帳"
find_flags = []
for _, row in final_proHours.iterrows():
    find_flag = 1 if "追帳" in row['專案內容說明'] else 0
    find_flags.append(find_flag)

# 保存 FIND 的結果
final_proHours['FIND結果'] = find_flags
# 判斷空值：'專案內容說明' 是否為空
empty_flags = []
for _, row in final_proHours.iterrows():
    empty_flag = 1 if row['專案內容說明'] == "" else 0
    empty_flags.append(empty_flag)

# 保存空值判斷結果
final_proHours['空值判斷'] = empty_flags
final_proHours['主管是否同意申請'] = final_proHours['主管是否同意申請'].str.get(0)
final_proHours['主管是否同意申請'] = final_proHours['主管是否同意申請'].fillna("")
# 判斷 '主管是否同意申請' 是否不等於 "是"
not_equal_flags = []
for _, row in final_proHours.iterrows():
    not_equal_flag = 1 if row['主管是否同意申請'] != "是" else 0
    not_equal_flags.append(not_equal_flag)
# 保存不等於 "是" 的判斷結果
final_proHours['主管同意結果'] = not_equal_flags

# 條件加總
final_flags = []
for _, row in final_proHours.iterrows():
    total = (
        row['COUNTIFS結果'] +
        row['FIND結果'] +
        row['空值判斷'] +
        row['主管同意結果']
    )
    final_flag = 1 if total > 0 else 0
    final_flags.append(final_flag)
# 保存最終結果到 '休假追帳不計算'
final_proHours['休假追帳不計算'] = final_flags
# 刪除中間過程的輔助欄位
final_proHours.drop(['COUNTIFS結果', 'FIND結果', '空值判斷', '主管同意結果'], axis=1, inplace=True)
Excel_專案時數 = final_proHours.copy()

time.sleep(5)

'''
公司交辦
'''
from Company_Task import Company_Task
print(" " * 15 + "📅 Loading.. 公司交辦 📅")
print("=" * 50)
time.sleep(5)
final_task = Company_Task()
Excel_台海交辦 = final_task.copy()
'''
多位拜訪明細
'''
from Multiple_Visit_List import Multiple_Visit_List
print(" " * 15 + "📅 Loading.. 多位拜訪明細 📅")
print("=" * 50)
time.sleep(5)
Multiple_Visit_List = Multiple_Visit_List()
Excel_多位拜訪 = Multiple_Visit_List.copy()
time.sleep(2)

# -------------------------------- 追蹤記錄 - 明細 -------------------------------------
print(" " * 15 + "追蹤記錄-明細 ")
print("=" * 50)

final_tracking['拜訪開始時間'] = pd.to_datetime(final_tracking['拜訪開始時間'], errors='coerce')

# 篩選條件：排除包含特定字串的公司名稱
wordlist = "反映與建議|明日行程"
final_tracking = final_tracking[~final_tracking['公司名稱'].astype(str).str.contains(wordlist, case=False, na=False)]

final_tracking = final_tracking[
    ((final_tracking['拜訪開始時間'] >= start_date_of_month) & 
     (final_tracking['拜訪開始時間'] <= end_date_of_month))
]

# 新增欄位清單
new_columns = [
    "客戶類別", "計算判斷", "同公司最多2", "同公司同聯絡人最多1", "同天同公司最多1",
    "是否特例(公司)", "是否有效聯絡人", "是否公司交辦", "是否新公司", "是否支援", "是否客訴", "拜訪時長是否>=15", "F類", "K類", "SE類", 
    "是否偏遠地區" ,"是否必拜訪","廣度觸及人數","其他客戶","同天同拜訪時間",
    "是否客訴","海外用-開發經營","檢視重複","檢視未計算",
     "人+客戶+公司", "人+公司", "日期"
]
for column in new_columns:
    final_tracking[column] = ""
time.sleep(2)
# -------------------------------- 追蹤記錄 - 日期 -------------------------------------
target_date = start_date_of_month
final_tracking['拜訪開始時間'] = pd.to_datetime(final_tracking['拜訪開始時間'], format='%Y/%m/%d')

# 自定義函數
def find_matching_date(row, target_date):
    target_month = target_date.month
    
    if row['拜訪開始時間'].month == target_month:
        return row['拜訪開始時間']
    elif row['建檔日期(實際執行日期)'].month == target_month:
        return row['建檔日期(實際執行日期)']
try:
    final_tracking['日期'] = final_tracking.apply(find_matching_date, target_date=target_date, axis=1)
except Exception:
    pass
time.sleep(2)
# -------------------------------- 追蹤記錄 - 每日最早到訪時間 -------------------------------------
final_tracking['拜訪開始時間'] = pd.to_datetime(final_tracking['拜訪開始時間'])
final_tracking['拜訪結束時間'] = pd.to_datetime(final_tracking['拜訪結束時間'])

final_tracking_time = final_tracking.copy()
final_tracking = final_tracking[final_tracking['觸客類型']!="['A2 無效拜訪']"]
final_tracking_time = final_tracking_time[final_tracking_time['業務工作記錄代號'].str.contains('GTR')]

# 計算每個人每一天的最早拜訪時間
final_tracking_time['首間客戶到訪時間'] = (
    final_tracking_time.groupby(['業務人員姓名', '日期'])['拜訪開始時間']
    .transform('min')  
)

final_tracking_time['末間客戶離訪時間'] = (
    final_tracking_time.groupby(['業務人員姓名', '日期'])['拜訪結束時間']
    .transform('max')  
)

final_tracking_time['日期'] = final_tracking_time['首間客戶到訪時間'].dt.date  # 提取日期部分

final_tracking_first_time = final_tracking_time.sort_values(by=['日期', '首間客戶到訪時間'])
final_tracking_first_time = final_tracking_first_time.groupby(['日期', '業務人員姓名']).first().reset_index()
final_tracking_first_time['首間客戶到訪時間'] = pd.to_datetime(final_tracking_first_time['首間客戶到訪時間'], errors='coerce')
final_tracking_first_time['首間客戶到訪時間'] = final_tracking_first_time['首間客戶到訪時間'].dt.strftime('%H:%M:%S')
final_tracking_first_time = final_tracking_first_time[['日期','業務工作記錄代號',	'業務人員姓名','首間客戶到訪時間','交辦管理','公司代號','公司名稱','拜訪開始時間','觸客類型']]
final_tracking_first_time.rename(columns={'業務工作記錄代號':'首間追蹤記錄'}, inplace=True)

final_tracking_last_time = final_tracking_time.sort_values(by=['日期', '末間客戶離訪時間'])
final_tracking_last_time = final_tracking_last_time.groupby(['日期', '業務人員姓名']).last().reset_index()
final_tracking_last_time['末間客戶離訪時間'] = pd.to_datetime(final_tracking_last_time['末間客戶離訪時間'], errors='coerce')
final_tracking_last_time['末間客戶離訪時間'] = final_tracking_last_time['末間客戶離訪時間'].dt.strftime('%H:%M:%S')
final_tracking_last_time = final_tracking_last_time[['日期','業務工作記錄代號',	'業務人員姓名','末間客戶離訪時間','交辦管理','公司代號','公司名稱','拜訪結束時間','觸客類型']]
final_tracking_last_time.rename(columns={'業務工作記錄代號':'離訪追蹤記錄'}, inplace=True)

final_tracking_time = pd.merge(
    final_tracking_first_time,
    final_tracking_last_time[['離訪追蹤記錄','業務人員姓名', '日期', '末間客戶離訪時間','拜訪結束時間']],
    on=['業務人員姓名', '日期'],
    how='outer'  
)

final_tracking_time = final_tracking_time[['日期','業務人員姓名','首間追蹤記錄','首間客戶到訪時間','離訪追蹤記錄','末間客戶離訪時間','交辦管理','公司代號','公司名稱','拜訪開始時間','拜訪結束時間','觸客類型']]

Excel_到訪時間 = final_tracking_time.copy()
time.sleep(2)

# -------------------------------- 拜訪統計 - 業務樞紐 -------------------------------------

print(" " * 15 + "💡 計算..拜訪統計 - 業務樞紐 💡")
unique_values = final_sale['區域'].unique()
unique_values_list = unique_values.tolist()

data = {'大區': unique_values_list}
final_Report = pd.DataFrame(data)


##外勤業務拜訪報表##
expanded_rows = []  

for _, row in final_Report.iterrows():
    target_region = row['大區']
    matches = final_sale[final_sale['區域'] == target_region]
    
    if not matches.empty:
        for _, match_row in matches.iterrows():
            # 複製原始行數據，添加對應人員姓名（繁中與英文）
            new_row = row.copy()
            new_row['國家'] = match_row['國家']
            new_row['對應人員姓名(繁中)'] = match_row['人員姓名(繁中)']
            new_row['對應人員姓名(英文)'] = match_row['人員姓名(英文)']

            expanded_rows.append(new_row)
    else:
        # 如果沒有匹配，添加空值
        new_row = row.copy()
        new_row['國家'] = ""
        new_row['對應人員姓名(繁中)'] = ""
        new_row['對應人員姓名(英文)'] = ""
        expanded_rows.append(new_row)

expanded_final_Report = pd.DataFrame(expanded_rows)
expanded_final_Report.reset_index(drop=True, inplace=True)

# 遍歷 expanded_final_Report 的每一行
for index, row in expanded_final_Report.iterrows():
    # 取得當前行的對應人員姓名
    person_name = row['對應人員姓名(繁中)']
    
    # 檢查 final_sale 中是否有匹配的 '人員姓名(繁中)'
    match = final_sale[final_sale['人員姓名(繁中)'] == person_name]
    
    if not match.empty:
        # 如果匹配，將 '負責區域' 填入 expanded_final_Report 的 '區域'
        expanded_final_Report.at[index, '區域'] = match['負責區域'].values[0]
    else:
        # 如果沒有匹配，保留原值或設定為空
        expanded_final_Report.at[index, '區域'] = None
time.sleep(2)        
 
expanded_final_Report.rename(columns={'對應人員姓名(繁中)': '外勤業務姓名', '對應人員姓名(英文)': '外勤業務英文名' }, inplace=True)
final_Report = expanded_final_Report[['國家','大區','區域','外勤業務姓名','外勤業務英文名']]


print(" " * 15 + "💡 計算..到職日 💡")
time.sleep(2)
final_Report['到職日/接區日']=""
final_Report['到職日/接區日'] = pd.to_datetime(final_Report['到職日/接區日'])
# 遍歷 expanded_final_Report 的每一行
for index, row in final_Report.iterrows():
    # 取得當前行的對應人員姓名
    person_name = row['外勤業務姓名']
    
    # 檢查 final_sale 中是否有匹配的 '人員姓名(繁中)'
    match = final_sale[final_sale['人員姓名(繁中)'] == person_name]
    
    if not match.empty:
        # 如果匹配，將 '負責區域' 填入 expanded_final_Report 的 '區域'
        final_Report.at[index, '到職日/接區日'] = match['接區日期'].values[0]
    else:
        # 如果沒有匹配，保留原值或設定為空
        final_Report.at[index, '到職日/接區日'] = None
final_Report['到職日/接區日'] = pd.to_datetime(final_Report['到職日/接區日'], errors='coerce')
final_Report['到職日/接區日'] = final_Report['到職日/接區日'].dt.strftime('%Y/%m/%d')


weekday_today = datetime.now().weekday()

# 如果今天是星期一 (0)，則使用 days=3，否則使用 days=1
days_to_subtract = 3 if weekday_today == 0 else 1

# 計算前幾天的 0:00:00 和 23:59:59 的時間戳
date_start = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0) - timedelta(days=days_to_subtract)
end_date_yesterday = date_start + timedelta(hours=0, minutes=0)

print(" " * 15 + "💡 計算..拜訪統計 - 本月休假時數 💡")
time.sleep(2)
final_Report['本月休假時數'] = 0
final_Report['起日期']=start_date_of_month
final_Report['迄日期']=end_date_yesterday
final_Report['起日期']=final_Report['起日期'].dt.strftime('%Y/%m/%d')
final_Report['迄日期']=final_Report['迄日期'].dt.strftime('%Y/%m/%d')
final_dayoff['開始日期'] = pd.to_datetime(final_dayoff['開始日期'])
final_dayoff['結束日期'] = pd.to_datetime(final_dayoff['結束日期'])
final_Report['到職日/接區日'] = pd.to_datetime(final_Report['到職日/接區日'])
for i, row in final_Report.iterrows():
    if row['到職日/接區日'] > start_date_of_month:
        result = final_dayoff.loc[
            ( (final_dayoff['員工姓名'] == row['外勤業務姓名']) )  &
            (final_dayoff['開始日期'] >= row['到職日/接區日']) &
            (final_dayoff['結束日期'] <= end_date_yesterday),
            '實際請假時數'
        ].sum()
    else:
        result = final_dayoff.loc[
            ( (final_dayoff['員工姓名'] == row['外勤業務姓名']) )  &
            (final_dayoff['開始日期'] >= start_date_of_month) &
            (final_dayoff['結束日期'] <= end_date_yesterday),
            '實際請假時數'
        ].sum()
    
    # 將結果填入 final_Report 的對應欄位
    final_Report.at[i, '本月休假時數'] = result

print(" " * 15 + "💡 計算..拜訪統計 - 專案時數 💡")
time.sleep(2)
final_Report['專案時數'] = ""
final_proHours['專案時長'] = pd.to_numeric(final_proHours['專案時長'], errors='coerce')
final_proHours['執行日期'] = pd.to_datetime(final_proHours['執行日期'])
# 計算專案時數
for i, row in final_Report.iterrows():
    if row['到職日/接區日'] > start_date_of_month:
        # 從到職日計算專案時數
        result = (
            final_proHours.loc[
                ((final_proHours['申請人'] == row['外勤業務姓名']) | 
                 (final_proHours['申請人'] == row['外勤業務英文名'])) &
                (final_proHours['執行日期'] >= row['到職日/接區日']) &
                (final_proHours['休假追帳不計算'] == 0) &
                (final_proHours['執行日期'] <= end_date_yesterday),
                '專案時長'
            ].sum() / 60
        ).round(1)
    else:
        # 從當月第一天計算專案時數
        result = (
            final_proHours.loc[
                ((final_proHours['申請人'] == row['外勤業務姓名']) | 
                 (final_proHours['申請人'] == row['外勤業務英文名'])) &
                (final_proHours['執行日期'] >= start_date_of_month) &
                (final_proHours['休假追帳不計算'] == 0) &
                (final_proHours['執行日期'] <= end_date_yesterday),
                '專案時長'
            ].sum() / 60
        ).round(1)
    final_Report.at[i, '專案時數'] = result

print(" " * 15 + "💡 計算..拜訪統計 - 工作天 💡")
time.sleep(2)
final_Report['工作天']=""
final_calendarid['日期'] = pd.to_datetime(final_calendarid['日期'])
final_Report['到職日/接區日'] = pd.to_datetime(final_Report['到職日/接區日'])

# 過濾出工作日
workdays = final_calendarid.loc[final_calendarid['特殊假'].isna(), ['國家', '日期']]
today = datetime.today()
previous_workday = today - pd.tseries.offsets.BDay(1)
previous_workday = previous_workday.strftime('%Y-%m-%d')
previous_workday = pd.to_datetime(previous_workday)
print(previous_workday)
def calculate_daily_workdays(row):
    start_date = max(row['到職日/接區日'], start_date_of_month)
    country_workdays = workdays.loc[workdays['國家'] == row['國家']]
    filtered_workdays = country_workdays.loc[
        (country_workdays['日期'] >= start_date) & (country_workdays['日期'] <= previous_workday)
    ]

    days_count = len(filtered_workdays)
    
    return days_count
final_Report['工作天'] = final_Report.apply(calculate_daily_workdays, axis=1)

# 先計算每人「當月有效工作日」數量（已完成邏輯）
def calculate_daily_workdays(row):
    start_date = max(row['到職日/接區日'], start_date_of_month)
    country_workdays = workdays[workdays['國家'] == row['國家']]
    filtered_workdays = country_workdays[
        (country_workdays['日期'] >= start_date) & (country_workdays['日期'] <= previous_workday)
    ]
    return filtered_workdays['日期'].tolist()  # 回傳實際日期清單

final_Report['工作日清單'] = final_Report.apply(calculate_daily_workdays, axis=1)

# 初始化請假 & 時數欄位
final_Report['本月休假時數'] = final_Report['本月休假時數'].fillna(0).astype(float)
final_Report['專案時數'] = final_Report['專案時數'].fillna(0).astype(float)

# 計算請假影響後的有效出勤日（依 8 小時換算）
final_Report['實際工作天數'] = final_Report['工作日清單'].apply(len)
final_Report['請假折抵天數'] = ((final_Report['本月休假時數'] + final_Report['專案時數']) / 8).round(1)
final_Report['實際工作日'] = final_Report['實際工作天數'] - final_Report['請假折抵天數']

time.sleep(2)
# -------------------------------- 追蹤記錄 - 同公司同聯絡人最多1 -------------------------------------
final_tracking['人+公司'] = final_tracking['業務人員姓名'] + final_tracking['公司代號']
final_tracking['人+客戶+公司'] = final_tracking['業務人員姓名'] + final_tracking['客戶關係連絡人']+final_tracking['公司代號']

final_tracking['同公司最多2'] = final_tracking.groupby('人+公司').cumcount() + 1
final_tracking.loc[final_tracking['公司名稱'].str.contains('廣度經營') , '同公司最多2'] = 1

final_tracking['同公司同聯絡人最多1'] = final_tracking.groupby('人+客戶+公司').cumcount() + 1
final_tracking.loc[final_tracking['公司名稱'].str.contains('廣度經營') , '同公司同聯絡人最多1'] = 1

final_tracking['同天同公司最多1'] = final_tracking.groupby(['人+公司', '日期']).cumcount() + 1
final_tracking.loc[final_tracking['公司名稱'].str.contains('廣度經營') , '同天同公司最多1'] = 1
# -------------------------------- 追蹤記錄 - 同天同拜訪時間 -------------------------------------
final_tracking['同天同拜訪時間'] = 0
final_tracking['同天同拜訪時間'] = (
    final_tracking.groupby(['人+公司', '拜訪開始時間'])
    .cumcount() + 1
)

# -------------------------------- 追蹤記錄 - 是否特例(公司) -------------------------------------

final_tracking['是否特例(公司)'] = 0
# 條件邏輯
condition_1 = (
    (final_tracking['業務類型'] == '案例-追蹤紀錄') & 
    (final_tracking['業務人員姓名'].isin(
        final_Report.loc[final_Report['大區'] == '專案', '外勤業務姓名']
    ))
)

# 設定結果
final_tracking['是否特例(公司)'] = np.where(condition_1, 1, 0)

# -------------------------------- 追蹤記錄 - 是否公司交辦 -------------------------------------

final_tracking['是否公司交辦'] = 0

# 條件 1：台海交辦的匹配
condition_1 = (
    final_tracking['公司型態'].astype(str).str.contains('Z', na=False) & 
    final_tracking['公司名稱'].astype(str).str.contains('廣度經營', na=False)
)


# 條件 2：台海交辦的 CRM 案例匹配
condition_2 = final_tracking['人+公司'].isin(final_task['人+公司'])

# 條件 4：交辦管理的 CRM 案例匹配
condition_3 = (
    final_tracking['交辦管理'].notna() & 
    final_tracking['交辦管理'].isin(
        final_task.loc[
            final_task['工作主旨'].str.contains('CRM案例', na=False), '交辦管理編號'
        ]
    )
)

# 組合所有條件
final_tracking['是否公司交辦'] = np.where(
    condition_1 | condition_2 | condition_3 ,
    1,
    0
)

time.sleep(2)
# -------------------------------- 追蹤記錄 - 客戶類別 -------------------------------------
  
final_tracking['客戶類別'] = ""

# 条件1：拜访时长 > 10
final_tracking['拜訪時長'] = pd.to_numeric(final_tracking['拜訪時長'])
condition1 = final_tracking['拜訪時長'] >= 10

# 条件2：公司型态为 C、D、Z 且公司名称包含关键字
condition2 = (
    final_tracking['公司型態'].isin(['C', 'D', 'Z']) &
    final_tracking['公司名稱'].astype(str).str.contains('廣度經營', na=False) &
    final_tracking['業務類型'].fillna('').apply(
        lambda x: any(keyword in x for keyword in ['一般記錄', '支援記錄'])
    )
)

# 条件3：台交办或大陆交办的任务符合条件
task_condition_tw = final_task['工作主旨'].str.contains('拜訪廣度', na=False)

condition3 = (
    final_tracking['人+公司'].isin(final_task.loc[task_condition_tw, '人+公司']))

# 应用条件
final_tracking['客戶類別'] = '指標'
final_tracking.loc[condition1 & (condition2 | condition3  ), '客戶類別'] = '廣度'
    
    
# =IF(COUNTIFS(台海交辦!$B:$B, AA2, 台海交辦!$I:$I, "*主管指定*") > 0, "指標", A2)

def check_condition(row):
    # 检查 final_task 中是否有任何符合条件的行
    count = len(final_task[(final_task['人+公司'] == row['人+公司']) & 
                            (final_task['工作主旨'].str.contains("(F類)主管指定拜訪"))])
    
    if count > 0:
        return '指標'
    else:
        return row['客戶類別']

final_tracking['客戶類別'] = final_tracking.apply(check_condition, axis=1)

# -------------------------------- 追蹤記錄 - 是否客訴 -------------------------------------
final_tracking['是否客訴']=""
final_tracking['是否客訴'] = np.where(
    final_tracking['工作類別名稱'].str.contains('客訴|客诉', regex=True),
    1,
    0
)
time.sleep(2)
# -------------------------------- 追蹤記錄 - 是否有效聯絡人 -------------------------------------
final_tracking['是否有效聯絡人'] = 0  # 預設為 0

# 先篩選 twdata_customer 內 '是否計算' 為 1 的客戶關係連絡人編號
valid_contacts = twdata_customer.loc[twdata_customer['是否計算'] == 1, '客戶關係連絡人編號']
# 檢查 final_tracking['客戶關係連絡人'] 是否在 valid_contacts 內
final_tracking.loc[final_tracking['客戶關係連絡人'].isin(valid_contacts), '是否有效聯絡人'] = 1
time.sleep(2)
# -------------------------------- 追蹤記錄 - 是否新公司 -------------------------------------
final_tracking['是否新公司'] = 0

final_tracking['CRM建檔日期']=pd.to_datetime(final_tracking['CRM建檔日期'])
final_tracking['公司建檔日期年月'] = final_tracking['CRM建檔日期'].dt.to_period('M')
final_tracking['拜訪開始時間年月'] = final_tracking['拜訪開始時間'].dt.to_period('M')

final_tracking['是否新公司'] = (
    (final_tracking['公司建檔日期年月'] == final_tracking['拜訪開始時間年月'])
).astype(int)  

# -------------------------------- 追蹤記錄 - 是否支援 -------------------------------------
final_tracking['是否支援'] = 0
final_tracking['是否支援'] = final_tracking['業務類型'].str.contains('支援', na=False).astype(int)



# -------------------------------- 追蹤記錄 - 海外用開發經營 -------------------------------------

final_tracking['海外用-開發經營'] = ""

# 先轉成 set 加快查找效率
task_lookup = dict(zip(final_task['人+公司'], final_task['開發經營']))

# 客戶類別先做成查找表，避免多次 filter
customer_category_lookup = final_Customer_Category.set_index('公司代號')['目標客戶類型'].to_dict()

def determine_category(row):
    # 條件 1：是否支援
    if row['是否支援'] == 1:
        return "開發"
    
    # 條件 2：台海交辦
    if row['人+公司'] in task_lookup:
        return task_lookup[row['人+公司']]
    
    # 條件 3/4：從公司代號判斷客戶類別
    company_id = row['公司代號']
    category_desc = customer_category_lookup.get(company_id, "")
    
    if "開發" in category_desc or "开发" in category_desc:
        return "開發"
    elif "經營" in category_desc or "经营" in category_desc:
        return "經營"
    
    # 預設
    return "開發"

final_tracking['海外用-開發經營'] = final_tracking.apply(determine_category, axis=1)


# -------------------------------- 追蹤記錄 - 拜訪時長是否>=15 -------------------------------------

conditions = [    (final_tracking['客戶類別'] == "廣度") & (final_tracking['拜訪時長'] >= 10),
    (final_tracking['拜訪時長'] >= 15)
]
choices = [1, 1]
final_tracking['拜訪時長是否>=15'] = np.select(conditions, choices, default=0)

# -------------------------------- 追蹤記錄 - 專案F類K類SE類 -------------------------------------

final_tracking['F類'] = 0

final_tracking['F類'] = final_tracking['公司型態'].isin(['FA', 'FB', 'FD', 'FZ']).astype(int)

matching_companies = final_task.loc[final_task['工作主旨'] == '(F類)主管指定拜訪', '人+公司']

final_tracking.loc[final_tracking['人+公司'].isin(matching_companies), 'F類'] = 1

final_tracking['K類'] = final_tracking['公司型態'].isin(['KZ']).astype(int)
final_tracking['SE類'] = final_tracking['公司型態'].isin(['SE']).astype(int)
final_tracking.loc[final_tracking['公司名稱'].str.contains('廣度經營'), 'SE類'] = 1

# -------------------------------- 追蹤記錄 - 是否偏遠地區 -------------------------------------
final_tracking['是否偏遠地區'] = 0
def calculate_and_accumulate(row):
    # 篩選出符合條件的記錄
    count = final_task[
        (row['人+公司']== final_task['人+公司']) &
        (final_task['是否為偏遠地區'] == '是')
    ].shape[0] 
    return row['是否偏遠地區'] + count
final_tracking['是否偏遠地區'] = final_tracking.apply(calculate_and_accumulate, axis=1) #J

# -------------------------------- 追蹤記錄 - 是否必拜訪 -------------------------------------

final_tracking['是否必拜訪'] = 0
def check_condition1(row):
    # 检查 final_task 中是否有任何符合条件的行
    count = len(final_task[(final_task['人+公司'] == row['人+公司']) & 
                            (final_task['工作主旨'].str.contains("必拜訪|必拜访"))])
    
    if count > 0:
        return 1
    else:
        return 0

# 使用 apply 函数来应用该逻辑
final_tracking['是否必拜訪'] = final_tracking.apply(check_condition1, axis=1)

# -------------------------------- 追蹤記錄 - 其他客戶 -------------------------------------
time.sleep(2)
final_tracking['其他客戶'] = 0
# 條件 1：拜訪時長 >= 15
condition1 = final_tracking['拜訪時長'] >= 15
# 條件 2：指定欄位值全部為 0
columns_to_check = [
    '是否特例(公司)', '是否有效聯絡人',
    '是否公司交辦', '是否新公司', 
    '是否支援','是否客訴'
]
condition2 = (final_tracking[columns_to_check] == 0).all(axis=1)
# 條件 3：必須同時滿足的多個條件
condition3 = (
    (final_tracking['同公司最多2'] <= 2) & 
    (final_tracking['同公司同聯絡人最多1'] <= 1) &
    (final_tracking['同天同公司最多1'] <= 1) &
    (final_tracking['同天同拜訪時間'] <= 1)
)
# 結合所有條件並更新 '其他客戶' 欄位
final_tracking['其他客戶'] = (condition1 & condition2 & condition3).astype(int)
final_tracking.loc[final_tracking['公司代號']=='GAC00067708' , '其他客戶'] = 1 
final_tracking.loc[final_tracking['業務工作記錄代號']=='GTR03828805' , '其他客戶'] = 1 

time.sleep(2)
# -------------------------------- 追蹤記錄 - 廣度觸及人數 -------------------------------------
time.sleep(2)
final_tracking['廣度觸及人數'] = ""

# 遍歷 final_tracking 中的每一行
for i, row in final_tracking.iterrows():
    ab_value = row['業務工作記錄代號']  # 代表 AB4
    ad_value = row['業務人員姓名']  # 代表 AD4

    # 檢查 AB4 是否在之前的 AB 列中出現過
    count_ab = (final_tracking['業務工作記錄代號'][:i+1] == ab_value).sum()

    if count_ab > 1:
        # 如果 AB4 已經出現過，則設置為 0
        final_tracking.at[i, '廣度觸及人數'] = 0
    else:
        # 如果 AB4 是第一次出現，使用 COUNTIFS 進行條件匹配
        condition = (
            (Multiple_Visit_List['業務工作記錄代號'] == ab_value) & 
            (Multiple_Visit_List['業務人員姓名'] == ad_value)
        )
        
        count_matches = condition.sum()
        
        # 設置 '是否公司交辦' 根據匹配數量
        final_tracking.at[i, '廣度觸及人數'] = count_matches
time.sleep(2)
print(" " * 15 + "💡 計算..追蹤記錄 - 計算判斷 💡")
# -------------------------------- 追蹤記錄 - 計算判斷 -------------------------------------
final_tracking['計算判斷'] = 0
final_tracking.loc[final_tracking['是否支援']==1 , '計算判斷'] = 1 
condition1 = final_tracking['拜訪時長是否>=15'] == 1
condition2 = final_tracking[['是否特例(公司)', '是否有效聯絡人', '是否公司交辦', '是否新公司', '是否支援']].any(axis=1)
condition3 = final_tracking['是否客訴'] == 1
condition4 = (final_tracking['同公司最多2'] <=2) & (final_tracking['同公司同聯絡人最多1']<=1) 
final_tracking.loc[condition1 &  condition2 & condition4  , '計算判斷'] = 1
final_tracking.loc[condition1 &  condition3   , '計算判斷'] = 1
print(final_tracking['計算判斷'].sum())

final_tracking = final_tracking[[
    '客戶類別', '計算判斷', '同公司最多2', '同公司同聯絡人最多1', '同天同公司最多1',
    "同天同拜訪時間","是否客訴",
    '是否特例(公司)', '是否有效聯絡人', '是否公司交辦', '是否新公司', '是否支援', '拜訪時長是否>=15',  
    'F類', 'K類', 'SE類',"是否偏遠地區","是否必拜訪",
    '廣度觸及人數',"其他客戶","海外用-開發經營","檢視重複","檢視未計算",
    '人+客戶+公司', '人+公司',"打卡記錄", 
    '日期', '業務工作記錄代號', '所屬部門', '業務人員姓名', '公司名稱', '創建日期(資料建立日期)', '公司型態', '最近工作內容_一', '案例代號', '案例名稱', 
    '客戶關係連絡人', '連絡人', '公司代號', '觸客類型', '拜訪開始時間', '拜訪結束時間', 
    '拜訪時長', '交辦管理', '工作類別名稱', '建檔日期(實際執行日期)', "是否邀約K大","講解分鐘數",'業務類型','CRM建檔日期','職務類別','特殊認列職務類別' 
]]

final_tracking['日期'] = final_tracking['日期'].dt.strftime('%Y/%m/%d')

# 確保日期格式一致
final_Report['迄日期'] = pd.to_datetime(final_Report['迄日期'], format='%Y/%m/%d')
final_dayoff['開始日期'] = pd.to_datetime(final_dayoff['開始日期'], format='%Y/%m/%d')


print(" " * 15 + "💡 計算..拜訪統計 - 當日是否休假.... 💡")
time.sleep(2)
final_Report['當日是否休假'] = ""

# 合併 DataFrame
merged_off = pd.merge(
    final_Report,
    final_dayoff,
    left_on=['外勤業務姓名', '迄日期'],  
    right_on=['員工姓名', '開始日期'],  
    how='left'
)

# 確保 `實際請假時數` 中的 NaN 值處理為 0，避免計算錯誤
merged_off['實際請假時數'] = merged_off['實際請假時數'].fillna(0)

# 判斷是否休假且實際請假時數需 >= 8
merged_off['當日是否休假'] = np.where(
    (merged_off['員工姓名'].notna()) & (merged_off['實際請假時數'] >= 8),
    '是',
    '否'
)

final_Report['當日是否休假'] = merged_off['當日是否休假']

final_Report['當日休假時數'] = 0

merged_time = final_Report.merge(
    final_dayoff,
    left_on=['迄日期', '外勤業務姓名'],
    right_on=['開始日期', '員工姓名'],
    how='left'
)

# 按照 "外勤業務姓名" 和 "開始日期" 分組並計算 "實際請假時數" 的總和
grouped_time = merged_time.groupby(
    ['外勤業務姓名', '開始日期'], as_index=False
).agg({'實際請假時數': 'sum'})

# 將 grouped_time 與 final_Report 合併
final_Report = final_Report.merge(
    grouped_time,
    left_on=['迄日期', '外勤業務姓名'],
    right_on=['開始日期', '外勤業務姓名'],
    how='left'
)

final_Report['實際請假時數'] = final_Report['實際請假時數'].fillna(0)
# 更新 "當日休假時數"
final_Report['當日休假時數'] = np.where(
    final_Report['當日是否休假'] == '否',
    final_Report['實際請假時數'],
    0
)
final_Report.drop(['開始日期' ,'實際請假時數'] ,axis=1 , inplace=True)


final_Report['當日專案時數'] = 0
final_Report = pd.merge(final_Report, final_sale[['人員姓名(繁中)','用戶']] , left_on='外勤業務姓名',right_on='人員姓名(繁中)' , how='left')
merged_protime = final_proHours[final_proHours['休假追帳不計算']==0]
# merged_protime = merged_protime[~merged_protime['專案內容說明'].isin(['5/2支援k大', '5/6支援k大', '5/8支援k大', '5/9支援k大'])]


merged_protime = final_Report.merge(
    merged_protime[['執行日期', '申請人', '專案時長']],  # 先篩選需要的欄位
    left_on=['迄日期', '用戶'],  
    right_on=['執行日期', '申請人'],  
    how='left'
)

final_Report['當日專案時數'] = merged_protime['專案時長'].fillna(0) / 60

final_Report.drop(['人員姓名(繁中)','用戶'],axis=1,inplace=True)
# 打印結果
print(final_Report)
print(" " * 15 + "💡 計算..拜訪統計 - 當日國家假日 💡")
time.sleep(2)
final_Report['當日國家假日'] = 0

merged_holiday = final_Report.merge(
    final_calendarid[['國家', '日期', '特殊假']],  
    left_on=['迄日期', '國家'], 
    right_on=['日期', '國家'],
    how='left'  
)

final_Report['當日國家假日'] = merged_holiday['特殊假']

final_Report.loc[final_Report['當日國家假日'].notna(), '當日是否休假'] = '是'
final_Report.loc[final_Report['當日是否休假'] =='是','當日休假時數'] =8


boss_list = ['北區主管','中區主管','南區主管','專案區主管']
final_Report['是否為主管'] = np.where(final_Report['大區'].isin(boss_list), '是', '否')


# 重複檢視
final_tracking['檢視重複'] =""
final_tracking.loc[
    (final_tracking['業務人員姓名'].isin(sales_list_tw)) &
    (final_tracking['同公司最多2'] >= 2) &
    (~final_tracking['公司名稱'].astype(str).str.startswith('廣度經營')),
    '檢視重複'
] = '重複'

# 專案重複檢視

final_tracking.loc[
    (final_tracking['業務人員姓名'].isin(sales_list_pro)) &
    (final_tracking['同公司最多2'] >=2 ) ,
    '檢視重複'
] = '重複'

# 海外重複檢視
final_tracking.loc[
    (final_tracking['業務人員姓名'].isin(sales_list_os)) &
    (final_tracking['同公司最多2'] >=2 ) &
    (final_tracking['同公司同聯絡人最多1'] >1 )    , 
    '檢視重複'
] = '重複'

# 未計算
final_tracking['檢視未計算'] =""
final_tracking.loc[
    (final_tracking['業務人員姓名'].isin(sales_list_tw)) &
    (final_tracking['計算判斷'] ==0) &
    (final_tracking['其他客戶'] == 0) &
    (final_tracking['同天同拜訪時間'] ==1)  ,
    '檢視未計算'
] = '未計算'


# 專案未計算
final_tracking.loc[
    (final_tracking['業務人員姓名'].isin(sales_list_pro)) &
    (final_tracking['拜訪時長是否>=15']==0) ,
    '檢視未計算'
] = '未計算'

final_tracking.loc[
    (final_tracking['業務人員姓名'].isin(sales_list_pro)) &
    (final_tracking['F類'] == 0) &
    (final_tracking['K類'] == 0) &
    (final_tracking['SE類'] == 0)   ,
    '檢視未計算'
] = '未計算'
t = final_tracking[final_tracking['業務人員姓名']=='蔡家維']
# 海外未計算

final_tracking.loc[
    (final_tracking['業務人員姓名'].isin(sales_list_os)) &
    (final_tracking['計算判斷'] ==0) &
    (final_tracking['其他客戶'] == 0) &
    (final_tracking['同天同拜訪時間'] ==1)  ,
    '檢視未計算'
] = '未計算'

'''
*特殊認列名單:
  一、新建 CRM 客戶，且上級客戶近兩個月內未被拜訪
  二、追蹤紀錄與交辦的公司相符，聯絡人不符，但拜訪的職務類別為老闆、 設計
  師、 設計總監/設計主管類、KEY MAN   vsale_list
  三、支援(需於「近工作內容_一」登打「支援記錄」的字樣) v
  四、客訴處理認列指標客戶 v

'''

  # 二、追蹤紀錄與交辦的公司相符，聯絡人不符，但拜訪的職務類別為老闆、 v
condition = (
    (final_tracking['是否公司交辦'] == 1) &
    (final_tracking['是否有效聯絡人'] == 0) &
    (final_tracking['特殊認列職務類別'] == 1)
)

final_tracking.loc[condition, '客戶類別'] = "指標"
final_tracking.loc[condition, '海外用-開發經營'] = "開發"

  # 三、支援認列指標客戶 v

condition = (
    (final_tracking['業務類型']=='支援記錄')
)

final_tracking.loc[condition, '客戶類別'] = "指標"
final_tracking.loc[condition, '海外用-開發經營'] = "開發"
  # 四、支援(需於「近工作內容_一」登打「支援記錄」的字樣) v

condition = (
    (final_tracking['是否客訴']==1)
)
final_tracking.loc[condition, '客戶類別'] = "指標"
final_tracking.loc[condition, '檢視重複'] = "" #若客訴拜訪,不卡重覆
final_tracking.loc[condition, '海外用-開發經營'] = "開發"
# 確保欄位為數值型
final_tracking['同公司最多2'] = pd.to_numeric(final_tracking['同公司最多2'], errors='coerce')
final_tracking['同公司同聯絡人最多1'] = pd.to_numeric(final_tracking['同公司同聯絡人最多1'], errors='coerce')

# 應用條件更新，減 1 並最低保留 1
final_tracking.loc[condition, '同公司最多2'] = (
    final_tracking.loc[condition, '同公司最多2'] - 1
).clip(lower=1)

final_tracking.loc[condition, '同公司同聯絡人最多1'] = (
    final_tracking.loc[condition, '同公司同聯絡人最多1'] - 1
).clip(lower=1)

condition = (
    (final_tracking['案例代號'].str.strip() != "") &
    (final_tracking['是否公司交辦'] == 1)
)
final_tracking.loc[condition, '客戶類別'] = "指標"
final_tracking.loc[condition, '海外用-開發經營'] = "開發"
time.sleep(1) 

Excel_追蹤記錄 = final_tracking.copy()
###########################分類各區統計表###########################
###########################分類各區統計表###########################
###########################分類各區統計表###########################

required_columns = [
    '國家', '大區', '區域', '外勤業務姓名', '實際工作日', '本月休假時數', '專案時數',
    '指標客戶', '本日目標(指標)', '本日指標(補數)', '廣度客戶', '本日目標(廣度)', '本日廣度(補數)',
    '客訴拜訪', '其他客戶', '拜訪總計', '本日總達成率', '本月累積指標目標', '本月指標目標(20%)',
    '本月指標客戶', '本月指標客戶(終)', '本月指標達成率', '本月累積廣度目標', '本月廣度目標(20%)',
    '本月廣度客戶', '本月廣度客戶(終)', '本月廣度達成率', '本月其他客戶', '本月必拜訪數', '本月必拜訪目標',
    '偏遠地區完成數', '本月偏遠地區目標', '重複拜訪', '未計算', '誤登打(同天同拜訪時間)',
    '本月拜訪總計', '首間客戶到訪時間','末間客戶離訪時間', '累計觸及人數', '累積達成率', '外勤業務英文名', '外勤業務簡中名',
    '工作天', '指標目標(值)', '廣度目標(值)', '到職日/接區日', '起日期', '迄日期', '當日是否休假',
    '當日休假時數', '當日國家假日', '是否為主管','當日專案時數','任職天數','工作日清單','請假折抵天數'
]
final_Report["任職天數"] = (end_date_yesterday - final_Report["到職日/接區日"]).dt.days
# 確保所有必要的欄位都存在
for col in required_columns:
    if col not in final_Report.columns:
        final_Report[col] = None  # 新增欄位，並填充為空值

# 確保欄位順序與 required_columns 一致
final_Report = final_Report[required_columns]

final_Report_TW = final_Report[final_Report['外勤業務姓名'].isin(sales_list_tw)]
final_Report_OS = final_Report[final_Report['外勤業務姓名'].isin(sales_list_os)]
final_Report_Pro = final_Report[final_Report['外勤業務姓名'].isin(sales_list_pro)]

###########################台灣統計判斷###########################
###########################台灣統計判斷###########################
final_Report_TW['指標目標(值)'] = 5 #輔助欄
final_Report_TW['指標目標(值)2'] = 5 #輔助欄
final_Report_TW['廣度目標(值)'] = 0 #輔助欄
final_Report_TW.loc[final_Report_TW['大區'].isin(['北區主管', '中區主管', '南區主管']),'指標目標(值)'] = 1.2 #輔助欄
final_Report_TW.loc[final_Report_TW['大區'].isin(['北區主管', '中區主管', '南區主管']),'指標目標(值)2'] = 1.2 #輔助欄
final_Report_TW.loc[final_Report_TW['大區'].isin(['北區主管', '中區主管', '南區主管']),'廣度目標(值)'] = 0 #輔助欄


# 再根據天數去套入指標目標(值)
final_Report_TW["指標目標(值)"] = np.where(
    final_Report_TW["任職天數"] <= 60,
    final_Report_TW["指標目標(值)"] * 0.8,
    np.where(
        (final_Report_TW["任職天數"] <= 120) & (final_Report_TW["任職天數"] > 60),
        final_Report_TW["指標目標(值)"] * 0.9,
        final_Report_TW["指標目標(值)"]
    )
)


'''
本日統計計算
'''

final_Report_TW['本日目標(指標)'] = final_Report_TW['指標目標(值)'] -  (((final_Report_TW['當日休假時數'] + final_Report_TW['當日專案時數']) / 8) * final_Report_TW['指標目標(值)'] )

final_Report_TW['指標客戶']=0 #G
final_tracking['日期'] = pd.to_datetime(final_tracking['日期'])

# 預先篩選符合條件的 final_tracking
filtered_tracking = final_tracking[
    (final_tracking['計算判斷'] == 1) &
    (final_tracking['客戶類別'] == '指標') &
    (final_tracking['同公司最多2'] <= 2) &
    (final_tracking['同公司同聯絡人最多1'] == 1) &
    (final_tracking['同天同拜訪時間'] == 1 )
]

def calculate_and_accumulate(row):
    relevant_rows = filtered_tracking[
        (filtered_tracking['業務人員姓名'] == row['外勤業務姓名']) &
        ( row['是否為主管'] =='否') &
        (filtered_tracking['日期'] == row['迄日期'])
    ]
    
    # 判斷是否為偏遠地區並計算權重
    count = relevant_rows.shape[0]
    remote_count = relevant_rows[relevant_rows['是否偏遠地區'] == 1].shape[0]
    # 偏遠地區的行數 * 1.2，加上其他行數
    weighted_count = count + remote_count * 0.2
    return row['指標客戶'] + weighted_count

final_Report_TW['指標客戶'] = final_Report_TW.apply(calculate_and_accumulate, axis=1)


filtered_tracking = final_tracking[
    (final_tracking['計算判斷'] == 1) &
    (final_tracking['同公司最多2'] <= 2) &
    (final_tracking['同公司同聯絡人最多1'] == 1)
]

def calculate_and_accumulate(row):
    # 篩選出符合條件的記錄
    relevant_rows = filtered_tracking[
        (filtered_tracking['業務人員姓名'] == row['外勤業務姓名']) &
        (filtered_tracking['日期'] == row['迄日期'])
    ]
    
    # 如果是主管，累加篩選結果行數
    if row['是否為主管'] == '是':
        return row['指標客戶'] + relevant_rows.shape[0]
    
    # 如果不是主管，保持原本的值
    return row['指標客戶']

# 計算並更新欄位
final_Report_TW['指標客戶'] = final_Report_TW.apply(calculate_and_accumulate, axis=1)



time.sleep(1)


final_Report_TW['本日目標(廣度)'] = final_Report_TW['廣度目標(值)'] -  (((final_Report_TW['當日休假時數'] + final_Report_TW['當日專案時數']) / 8) * final_Report_TW['廣度目標(值)'] )

final_Report_TW['廣度客戶']=0 #I
def calculate_and_accumulate(row):
    # 篩選出符合條件的記錄
    count = final_tracking[
        (final_tracking['業務人員姓名'] == row['外勤業務姓名']) &  # 條件1
        (final_tracking['計算判斷'] == 1) &                    # 條件2
        (final_tracking['同公司最多2'] <= 2) &                 # 條件3
        (final_tracking['同公司同聯絡人最多1'] == 1) &          # 條件4
        (final_tracking['日期'] == row['迄日期']) &            # 條件5
        (final_tracking['客戶類別'] == '廣度')                 # 條件6
    ].shape[0]  # 計算符合條件的記錄數量

    # 返回累加結果
    return row['廣度客戶'] + count

# 更新欄位值
final_Report_TW['廣度客戶'] = final_Report_TW.apply(calculate_and_accumulate, axis=1) #I

time.sleep(1)

final_Report_TW['其他客戶']=0 #J
def calculate_and_accumulate(row):
    # 篩選出符合條件的記錄
    count = final_tracking[
        (final_tracking['業務人員姓名'] == row['外勤業務姓名']) &  # 條件1
        (final_tracking['其他客戶'] == 1) &                    # 條件2
        (final_tracking['同公司最多2'] <= 2) &                 # 條件3
        (final_tracking['同公司同聯絡人最多1'] == 1) &          # 條件4
        (final_tracking['是否客訴'] == 0) &          # 條件4
        (final_tracking['日期'] == row['迄日期'])
    ].shape[0]  # 計算符合條件的記錄數量
    # 返回累加結果
    return row['其他客戶'] + count
# 更新欄位值
final_Report_TW['其他客戶'] = final_Report_TW.apply(calculate_and_accumulate, axis=1) #J

# F8 = final_Report_TW['指標客戶']
# G8 = final_Report_TW['本日目標(指標)']
# I8 = final_Report_TW['廣度客戶']
# J8 = final_Report_TW['本日目標(廣度)']

final_Report_TW['本日指標(補數)'] = 0
def calculate_result(row):
    row['指標客戶'], row['本日目標(指標)'], row['廣度客戶'], row['本日目標(廣度)'] = row['指標客戶'], row['本日目標(指標)'], row['廣度客戶'], row['本日目標(廣度)']
    
    if row['指標客戶'] <= row['本日目標(指標)'] or row['廣度客戶'] >= row['本日目標(廣度)']:
        return row['指標客戶']
    elif row['廣度客戶'] < row['本日目標(廣度)'] and row['指標客戶'] > row['本日目標(指標)'] and ((row['指標客戶'] - row['本日目標(指標)']) * 2) >= (row['本日目標(廣度)'] - row['廣度客戶']):
        return ((((row['指標客戶'] - row['本日目標(指標)']) * 2) - (row['本日目標(廣度)'] - row['廣度客戶'])) / 2) + row['本日目標(指標)']
    elif row['廣度客戶'] < row['本日目標(廣度)'] and row['指標客戶'] > row['本日目標(指標)'] and ((row['指標客戶'] - row['本日目標(指標)']) * 2) < (row['本日目標(廣度)'] - row['廣度客戶']):
        return row['本日目標(指標)']
    else:
        return row['本日目標(指標)'] - ((row['指標客戶'] - row['本日目標(指標)']) * 2)

final_Report_TW['本日指標(補數)'] = final_Report_TW.apply(calculate_result, axis=1)
final_Report_TW.loc[(final_Report_TW['是否為主管']=='是', '本日指標(補數)' )] = 0
final_Report_TW['本日廣度(補數)'] = 0
def calculate_result(row):
    row['指標客戶'], row['本日目標(指標)'], row['廣度客戶'], row['本日目標(廣度)'] = row['指標客戶'], row['本日目標(指標)'], row['廣度客戶'], row['本日目標(廣度)']
    
    if row['廣度客戶']>=row['本日目標(廣度)'] or row['指標客戶']<=row['本日目標(指標)'] :
        return row['廣度客戶']
    elif row['廣度客戶'] < row['本日目標(廣度)'] and row['指標客戶']>row['本日目標(指標)'] and ((row['指標客戶'] - row['本日目標(指標)']) * 2) > (row['本日目標(廣度)'] - row['廣度客戶']):
        return row['本日目標(廣度)']
    elif row['廣度客戶'] < row['本日目標(廣度)'] and row['指標客戶'] > row['本日目標(指標)'] and ((row['指標客戶'] - row['本日目標(指標)']) * 2) <= (row['本日目標(廣度)'] - row['廣度客戶']):
        return ((row['指標客戶']-row['本日目標(指標)'])*2)+row['廣度客戶']
    else:
        return ((row['指標客戶']-row['本日目標(指標)'])*2)+(row['本日目標(廣度)']-row['廣度客戶'])

final_Report_TW['本日廣度(補數)'] = final_Report_TW.apply(calculate_result, axis=1)

final_Report_TW['客訴拜訪'] = 0
def calculate_and_accumulate(row):
    # 篩選出符合條件的記錄
    count = final_tracking[
        (final_tracking['業務人員姓名'] == row['外勤業務姓名']) &  # 條件1
        (row['是否為主管'] == '否') &                 # 條件3
        (final_tracking['計算判斷'] == 1) &                    # 條件2
        (final_tracking['是否客訴'] == 1) &                 # 條件3
        (final_tracking['日期'] == row['迄日期'])
    ].shape[0]  
    # 返回累加結果
    return row['客訴拜訪'] + count
# 更新欄位值
final_Report_TW['客訴拜訪'] = final_Report_TW.apply(calculate_and_accumulate, axis=1) #L


final_Report_TW['拜訪總計'] = final_Report_TW['指標客戶'] + final_Report_TW['廣度客戶'] + final_Report_TW['其他客戶']
final_Report_TW['本日總達成率'] = 0 
# 判斷是否為主管，並進行不同計算
def calculate_total_achievement(row):
    # 計算分母
    denominator = row['本日目標(指標)'] + row['本日目標(廣度)']
    if denominator == 0:  # 分母為 0 時，直接回傳 0
        return 0
    
    if '主管' in row['大區']:
        # 主管計算邏輯
        return np.minimum(
            (row['指標客戶'] + row['其他客戶']) / denominator,
            1
        )
    else:
        # 非主管計算邏輯
        return np.minimum(
            (row['本日指標(補數)'] + row['本日廣度(補數)']) / denominator,
            1
        )

# 計算本日總達成率，並將其轉換為百分比格式
final_Report_TW['本日總達成率'] = final_Report_TW.apply(calculate_total_achievement, axis=1)

# 計算本日總達成率，並將其轉換為百分比格式
# final_Report_TW['本日總達成率'] = final_Report_TW.apply(calculate_total_achievement, axis=1) * 100

# final_Report_TW['本日總達成率'] = final_Report_TW['本日總達成率'].apply(lambda x: f"{x:.0f}%")


columns_to_update = ['指標客戶', '本日目標(指標)', '本日指標(補數)', '廣度客戶', '本日目標(廣度)', '本日廣度(補數)',
'客訴拜訪', '其他客戶', '拜訪總計', '本日總達成率']
# 建立條件
mask = final_Report_TW['當日是否休假'] == '是'
mask2 = final_Report_TW['當日專案時數'] >= 8
# 對多個欄位批量更新
final_Report_TW.loc[(mask | mask2), columns_to_update] = '休'


'''
本月統計
'''

final_Report_TW['本月指標客戶']=0 #G

# 預先篩選符合條件的 final_tracking
filtered_tracking = final_tracking[
    (final_tracking['計算判斷'] == 1) &
    (final_tracking['客戶類別'] == '指標') &
    (final_tracking['同公司最多2'] <= 2) &
    (final_tracking['同公司同聯絡人最多1'] == 1) &
    (final_tracking['同天同拜訪時間'] == 1 )
]

def calculate_and_accumulate(row):
    relevant_rows = filtered_tracking[
        (filtered_tracking['業務人員姓名'] == row['外勤業務姓名']) &
        ( row['是否為主管'] =='否') &
        (filtered_tracking['日期'] >= row['起日期'] ) &
        (filtered_tracking['日期'] <= row['迄日期'] )
    ]
    
    # 判斷是否為偏遠地區並計算權重
    count = relevant_rows.shape[0]
    remote_count = relevant_rows[relevant_rows['是否偏遠地區'] == 1].shape[0]
    # 偏遠地區的行數 * 1.2，加上其他行數
    weighted_count = count + remote_count * 0.2
    return row['本月指標客戶'] + weighted_count

final_Report_TW['本月指標客戶'] = final_Report_TW.apply(calculate_and_accumulate, axis=1)


# 過濾符合條件的 tracking 資料
filtered_tracking_boss = final_tracking[
    (final_tracking['計算判斷'] == 1) &
    (final_tracking['同公司最多2'] <= 2) &
    (final_tracking['同公司同聯絡人最多1'] == 1) 
]

# 計算本月指標客戶（主管邏輯）
def calculate_and_accumulate(row):
    if row['是否為主管'] == '是':  
        relevant_rows_boss = filtered_tracking_boss[
            (filtered_tracking_boss['業務人員姓名'] == row['外勤業務姓名']) &
            (filtered_tracking_boss['日期'] >= row['起日期']) &
            (filtered_tracking_boss['日期'] <= row['迄日期'])
        ]
        count = relevant_rows_boss.shape[0]
        return row['本月指標客戶'] + count 
    else:
        return row['本月指標客戶']  

final_Report_TW['本月指標客戶'] = final_Report_TW.apply(calculate_and_accumulate, axis=1)



final_Report_TW['本月廣度客戶']=0 #I

# # 預先篩選符合條件的 final_tracking
# filtered_tracking = final_tracking[
#     (final_tracking['計算判斷'] == 1) &
#     (final_tracking['客戶類別'] == '廣度') &
#     (final_tracking['同天同拜訪時間'] == 1 )
# ]

# final_Report_TW['本月廣度客戶'] = final_Report_TW.apply(lambda row: 
#     filtered_tracking[
#         (filtered_tracking['業務人員姓名'] == row['外勤業務姓名']) &
#         (row['是否為主管'] == '否') &
#         (filtered_tracking['日期'] >= row['起日期']) &
#         (filtered_tracking['日期'] <= row['迄日期'])
#     ].shape[0], axis=1
# )


final_Report_TW['重複拜訪'] = 0
def duplicates(row):
    # 篩選出符合條件的記錄
    count = final_tracking[
        (final_tracking['業務人員姓名'] == row['外勤業務姓名']) &  # 條件1
        (final_tracking['同公司最多2'] >= 2) &                    # 條件2
        (~final_tracking['公司名稱'].astype(str).str.startswith('廣度經營')) 
    ].shape[0]  
    return row['重複拜訪'] + count
# 更新欄位值
final_Report_TW['重複拜訪'] = final_Report_TW.apply(duplicates, axis=1) #L


final_Report_TW['未計算'] = 0

def duplicates(row):
    # 篩選出符合條件的記錄
    count = final_tracking[
        (final_tracking['業務人員姓名'] == row['外勤業務姓名']) &  # 條件1
        (final_tracking['計算判斷'] == 0) &                    # 條件2
        (final_tracking['其他客戶'] == 0) &  # 條件3
        (final_tracking['同天同拜訪時間'] == 1)
        
    ].shape[0]  
    return row['未計算'] + count
# 更新欄位值
final_Report_TW['未計算'] = final_Report_TW.apply(duplicates, axis=1) #L

# time.sleep(1)
final_Report_TW['本月其他客戶']=0 #Z
def calculate_and_accumulate(row):
    # 篩選出符合條件的記錄
    count = final_tracking[
        (row['是否為主管'] == '否') &
        (final_tracking['業務人員姓名'] == row['外勤業務姓名']) &  # 條件1
        (final_tracking['其他客戶'] == 1)  &
        (final_tracking['同公司最多2'] <= 2) &
        (final_tracking['同公司同聯絡人最多1'] == 1) &
        (final_tracking['是否客訴'] == 0)
    ].shape[0] 
    return row['本月其他客戶'] + count
final_Report_TW['本月其他客戶'] = final_Report_TW.apply(calculate_and_accumulate, axis=1) #J


final_Report_TW['本月必拜訪數']=0 #Z
# def calculate_and_accumulate(row):
#     # 篩選出符合條件的記錄
#     count = final_tracking[
#         (row['是否為主管']=='否') &
#         (final_tracking['業務人員姓名'] == row['外勤業務姓名']) &  # 條件1
#         (final_tracking['是否必拜訪'] == 1)                   # 條件2
#     ].shape[0] 
#     return row['本月必拜訪數'] + count
# final_Report_TW['本月必拜訪數'] = final_Report_TW.apply(calculate_and_accumulate, axis=1) #J
# time.sleep(1)


final_Report_TW['本月必拜訪目標'] = 0 
# counts = []
# for index, row in final_Report_TW.iterrows():
#     if row['是否為主管'] == '否':  # 非主管
#         count = (
#             final_task[
#                 (final_task['執行人'] == row['外勤業務姓名']) &
#                 (final_task['工作主旨'].str.contains("必拜訪"))
#             ].shape[0]
            
#         )
#         count = max(0, count)  # 如果 count 小於 0，則設定為 0
#     else:  
#         count = 0
#     counts.append(count)
# final_Report_TW['本月必拜訪目標'] = counts
# time.sleep(1)


final_Report_TW['偏遠地區完成數'] = 0 
def calculate_and_accumulate(row):
    # 篩選出符合條件的記錄
    count = final_tracking[
        (row['是否為主管']=='否') &
        (final_tracking['業務人員姓名'] == row['外勤業務姓名']) &  # 條件1
        (final_tracking['是否偏遠地區'] == 1)                   # 條件2
    ].shape[0] 
    return row['偏遠地區完成數'] + count
final_Report_TW['偏遠地區完成數'] = final_Report_TW.apply(calculate_and_accumulate, axis=1) #J
time.sleep(1)

final_Report_TW['本月偏遠地區目標'] = 0 
counts = []
for index, row in final_Report_TW.iterrows():
    if row['是否為主管'] == '否':  
        count = (
            final_task[
                (final_task['執行人'] == row['外勤業務姓名']) &
                (final_task['業務類型'].str.contains("電/拜訪交辦")) &
                (final_task['工作主旨'].str.contains("指標")) &
                (final_task['是否為偏遠地區'] == '是')
            ].shape[0]
        )
        count = max(0, count)  
    else:  
        count = 0
    counts.append(count)
final_Report_TW['本月偏遠地區目標'] = counts
# time.sleep(1)
#平均到訪時間
final_arrival_time = final_tracking_time.copy()
final_arrival_time['日期'] = pd.to_datetime(final_arrival_time['日期'])
final_arrival_time.sort_values(by='日期' , ascending=True , inplace =True)
#排序時間並篩選唯一
final_arrival_time.drop_duplicates(subset=['日期', '業務人員姓名'], keep='first', inplace=True)
# final_tracking_time['到訪時間'] = pd.to_datetime(final_tracking_time['到訪時間'], format='%H:%M:%S').dt.time

final_Report_TW['首間客戶到訪時間'] = 0
counts = [] 
final_arrival_time
for index, row in final_Report_TW.iterrows():
    filtered_data = final_arrival_time[
        (final_arrival_time['業務人員姓名'] == row['外勤業務姓名']) &
        (final_arrival_time['日期'] >= start_date_of_month) &
        (final_arrival_time['日期'] <= end_date_of_month)
    ]
    
    if not filtered_data.empty:
        filtered_data['首間客戶到訪時間'] = pd.to_datetime(filtered_data['首間客戶到訪時間'].astype(str), format='%H:%M:%S')
        avg_time = filtered_data['首間客戶到訪時間'].mean()
        avg_time_str = avg_time.strftime('%H:%M:%S')  
    else:
        avg_time_str = '00:00:00'  
    
    counts.append(avg_time_str)  
final_Report_TW['首間客戶到訪時間'] = counts

final_Report_TW['末間客戶離訪時間'] = 0
counts = [] 
final_arrival_time
for index, row in final_Report_TW.iterrows():
    filtered_data = final_arrival_time[
        (final_arrival_time['業務人員姓名'] == row['外勤業務姓名']) &
        (final_arrival_time['日期'] >= start_date_of_month) &
        (final_arrival_time['日期'] <= end_date_of_month)
    ]
    
    if not filtered_data.empty:
        filtered_data['末間客戶離訪時間'] = pd.to_datetime(filtered_data['末間客戶離訪時間'].astype(str), format='%H:%M:%S')
        avg_time = filtered_data['末間客戶離訪時間'].mean()
        avg_time_str = avg_time.strftime('%H:%M:%S')  
    else:
        avg_time_str = '00:00:00'  
    
    counts.append(avg_time_str)  
final_Report_TW['末間客戶離訪時間'] = counts

final_Report_TW['本月累積指標目標'] = 0
# final_Report_TW['本月累積指標目標'] = (
#     final_Report_TW['實際工作日'] * final_Report_TW['指標目標(值)']
# ).astype(float)

default_target = 5

def calculate_cumulative_target(row):
    # 若兩個欄位值相等 → 直接使用乘法
    if row['指標目標(值)'] == row['指標目標(值)2']:
        return round(row['實際工作日'] * row['指標目標(值)'], 2)

    # 否則跑逐日判斷邏輯
    onboard = row['到職日/接區日']
    work_dates = row['工作日清單']
    leave_days = int(row['請假折抵天數'])
    daily_target = row['指標目標(值)2']

    cumulative = 0
    skip = 0

    for day in sorted(work_dates):
        if skip < leave_days:
            skip += 1
            continue

        tenure = (day - onboard).days
        if tenure < 0:
            continue
        elif tenure <= 60:
            multiplier = 0.8
        elif 61 <= tenure <= 119:
            multiplier = 0.9
        else:
            multiplier = 1.0

        cumulative += daily_target * multiplier

    return round(cumulative, 2)

# 套用邏輯計算
final_Report_TW['本月累積指標目標'] = final_Report_TW.apply(calculate_cumulative_target, axis=1)


final_Report_TW2 = final_Report_TW[['外勤業務姓名','本月累積指標目標']]
final_Report_TW['本月累積廣度目標'] = 0

# final_Report_TW['本月累積廣度目標'] = (
#     final_Report_TW['實際工作日'] * final_Report_TW['廣度目標(值)']
# ).astype(float)
final_Report_TW['本月廣度目標(20%)'] = (final_Report_TW['本月累積廣度目標'] * 0.2).round(0) #U
final_Report_TW['本月指標目標(20%)'] = (final_Report_TW['本月累積指標目標'] * 0.2 + final_Report_TW['本月廣度目標(20%)'] ).round(0)
time.sleep(1) 


# '''
# 輔助欄
# '''
# import numpy as np

# def calculate_indicator_less_than_target(row):
#     if (np.round(row['本月指標客戶'], 0) < np.round(row['本月累積指標目標'], 0)) and (row['本月其他客戶'] < row['本月指標目標(20%)']):
#         return row['本月指標客戶'] + row['本月其他客戶']
#     elif (row['本月指標客戶'] < np.round(row['本月累積指標目標'], 0)) and (row['本月其他客戶'] > row['本月指標目標(20%)']):
#         return row['本月指標客戶'] + row['本月指標目標(20%)']
#     return 0

# def calculate_indicator_greater_than_target(row):
#     if row['本月指標客戶'] >= np.round(row['本月累積指標目標'], 0):
#         return row['本月指標客戶'] - np.round(row['本月累積指標目標'], 0)
#     return 0

# def calculate_other_customer_compensation(row):
#     if (row['本月指標客戶'] >= np.round(row['本月累積指標目標'], 0)) or (row['指標大於目標'] > 0):
#         return min(row['本月其他客戶'], row['本月指標目標(20%)'])
#     elif (row['本月指標客戶'] < np.round(row['本月累積指標目標'], 0)) and ((row['本月指標客戶'] + min(row['本月其他客戶'], row['本月指標目標(20%)'])) > np.round(row['本月累積指標目標'], 0)):
#         return (row['本月指標客戶'] + min(row['本月其他客戶'], row['本月指標目標(20%)'])) - np.round(row['本月累積指標目標'], 0)
#     return 0

# def calculate_remain_indicator_compensation(row):
#     total_compensation = row['指標客戶1補2'] + row['其他客戶1補1']
#     if row['指標客戶1補2'] - row['可補廣度'] <= 0 and total_compensation - row['可補廣度'] <= 0:
#         return 0
#     elif row['指標客戶1補2'] - row['可補廣度'] <= 0 and total_compensation - row['可補廣度'] > 0:
#         return total_compensation - row['可補廣度']
#     elif row['指標客戶1補2'] >= row['需補廣度']:
#         return ((row['指標客戶1補2'] - row['可補廣度']) / 2) + row['其他客戶1補1']
#     return 0

# def calculate_final_boss_target(row, final_tracking):
#     if row['是否為主管'] == '是':
#         relevant_rows_boss = final_tracking[(final_tracking['業務人員姓名'] == row['外勤業務姓名']) & (final_tracking['其他客戶'] == 1)]
#         return row['本月指標客戶(終)'] + relevant_rows_boss.shape[0]
#     return row['本月指標客戶(終)']

# # Main logic
# final_Report_TW['指標小於目標'] = final_Report_TW.apply(calculate_indicator_less_than_target, axis=1)
# final_Report_TW['指標大於目標'] = final_Report_TW.apply(calculate_indicator_greater_than_target, axis=1)
# final_Report_TW['指標客戶1補2'] = final_Report_TW['指標大於目標'] * 2
# final_Report_TW['其他客戶1補1'] = final_Report_TW.apply(calculate_other_customer_compensation, axis=1)

# final_Report_TW['需補廣度'] = np.where(
#     final_Report_TW['本月廣度客戶'] > np.round(final_Report_TW['本月累積廣度目標'], 0),
#     0,
#     np.where(
#         (final_Report_TW['指標客戶1補2'] > (np.round(final_Report_TW['本月累積廣度目標'], 0) - final_Report_TW['本月廣度客戶'])) &
#         ((np.round(final_Report_TW['本月累積廣度目標'], 0) - final_Report_TW['本月廣度客戶']) % 2 == 1),
#         (np.round(final_Report_TW['本月累積廣度目標'], 0) - final_Report_TW['本月廣度客戶'] + 1),
#         (np.round(final_Report_TW['本月累積廣度目標'], 0) - final_Report_TW['本月廣度客戶'])
#     )
# )

# final_Report_TW['可補廣度'] = np.where(
#     (final_Report_TW['指標客戶1補2'] == 0) & (final_Report_TW['其他客戶1補1'] == 0),
#     0,
#     np.where(
#         final_Report_TW['指標客戶1補2'] >= final_Report_TW['需補廣度'],
#         final_Report_TW['需補廣度'],
#         np.where(
#             (final_Report_TW['指標客戶1補2'] + final_Report_TW['其他客戶1補1']) >= final_Report_TW['需補廣度'],
#             final_Report_TW['需補廣度'],
#             final_Report_TW['指標客戶1補2'] + final_Report_TW['其他客戶1補1']
#         )
#     )
# )

# final_Report_TW['剩餘可補回指標'] = final_Report_TW.apply(calculate_remain_indicator_compensation, axis=1)
final_Report_TW['本月指標客戶(終)'] = 0

# 如果 本月指標客戶 < 累積指標目標 且 其他客戶 <= 指標目標(20%) → 加總兩者
# 如果 本月指標客戶 < 累積指標目標 且 其他客戶 > 指標目標(20%) → 客戶 + 20%

final_Report_TW['本月指標客戶(終)'] = np.where(
    final_Report_TW['是否為主管'] == '是',
    0,
    np.where(
        (final_Report_TW['本月指標客戶'] < final_Report_TW['本月累積指標目標']) & 
        (final_Report_TW['本月其他客戶'] <= final_Report_TW['本月指標目標(20%)']),
        final_Report_TW['本月指標客戶'] + final_Report_TW['本月其他客戶'],
        np.where(
            (final_Report_TW['本月指標客戶'] < final_Report_TW['本月累積指標目標']) & 
            (final_Report_TW['本月其他客戶'] > final_Report_TW['本月指標目標(20%)']),
            final_Report_TW['本月指標客戶'] + final_Report_TW['本月指標目標(20%)'],
            final_Report_TW['本月指標客戶']
        )
    )
)


def calculate_and_accumulate(row):
    if row['是否為主管'] == '是':
        count = final_tracking[
            (final_tracking['業務人員姓名'] == row['外勤業務姓名']) &
            (final_tracking['計算判斷'] == 0) &
            (final_tracking['其他客戶'] == 1)
        ].shape[0]
        return row['本月指標客戶'] + count
    else:
        return row['本月指標客戶(終)']  # 非主管用原來計算好的值（np.where 產出的）

final_Report_TW['本月指標客戶(終)'] = final_Report_TW.apply(calculate_and_accumulate, axis=1)

final_Report_TW['本月指標達成率'] = (final_Report_TW['本月指標客戶(終)'] / final_Report_TW['本月累積指標目標']).clip(upper=1).fillna(0)

# final_Report_TW['本月廣度客戶(終)'] = np.where(
#     (final_Report_TW['本月廣度客戶'] < final_Report_TW['本月累積廣度目標']) & (final_Report_TW['可補廣度'] > 0),
#     final_Report_TW['本月廣度客戶'] + final_Report_TW['可補廣度'],
#     final_Report_TW['本月廣度客戶']
# )

# final_Report_TW['本月廣度達成率'] = (final_Report_TW['本月廣度客戶(終)'] / final_Report_TW['本月累積廣度目標']).clip(upper=1).fillna(0)

# final_Report_TW['本月廣度達成率'] = np.where(
#     final_Report_TW['本月廣度客戶(終)'] / final_Report_TW['本月累積廣度目標'] > 1,
#     1,
#     final_Report_TW['本月廣度客戶(終)'] / final_Report_TW['本月累積廣度目標']
# )
# final_Report_TW['本月廣度達成率'] = final_Report_TW['本月廣度達成率'].fillna(0) 
# final_Report_TW['本月廣度達成率'] = final_Report_TW['本月廣度達成率'] * 100  
# final_Report_TW['本月廣度達成率'] = final_Report_TW['本月廣度達成率'].apply(lambda x: f"{x:.0f}%")# 轉換為百分比


print(final_Report_TW)
time.sleep(1)

final_Report_TW['誤登打(同天同拜訪時間)'] = 0

def calculate_and_accumulate(row):
    # 篩選出符合條件的記錄
    count = final_tracking[
        (final_tracking['業務人員姓名'] == row['外勤業務姓名']) &  # 配對業務人員姓名
        (final_tracking['同天同拜訪時間'] >= 2) &                  # 同天同拜訪時間 >= 2
        (final_tracking['日期'] >= row['起日期']) &               # 日期 >= 起日期
        (final_tracking['日期'] <= row['迄日期'])                 # 日期 <= 迄日期
    ].shape[0]  # 計算符合條件的行數
    return count  # 返回符合條件的行數

# 應用該函數到每一行並計算結果
final_Report_TW['誤登打(同天同拜訪時間)'] = final_Report_TW.apply(calculate_and_accumulate, axis=1)


final_Report_TW['本月拜訪總計'] = 0

def calculate_and_accumulate(row):
    # 篩選出符合條件的記錄
    count = final_tracking[
        (final_tracking['業務人員姓名'] == row['外勤業務姓名']) & 
        (final_tracking['日期'] >= row['起日期']) &           
        (final_tracking['日期'] <= row['迄日期'])                 
    ].shape[0]  
    return count  - row['誤登打(同天同拜訪時間)']

# 應用該函數到每一行並計算結果
final_Report_TW['本月拜訪總計'] = final_Report_TW.apply(calculate_and_accumulate, axis=1)


final_Report_TW['累計觸及人數'] = 0 
def calculate_and_accumulate(row):
    # 篩選出符合條件的記錄，並計算觸及人數
    filtered = final_tracking[
        (row['是否為主管'] == '否') &
        (final_tracking['業務人員姓名'] == row['外勤業務姓名'])
    ]
    total_count = filtered['廣度觸及人數'].sum()
    total_count = row['本月拜訪總計']  + total_count  # 累加各項數據
    return total_count

final_Report_TW['累計觸及人數'] = final_Report_TW.apply(calculate_and_accumulate, axis=1)


final_Report_TW['累積達成率'] = 0
final_Report_TW['累積達成率'] = np.where(
    ( final_Report_TW['本月指標客戶(終)']  )  / ( final_Report_TW['本月累積指標目標']  )  > 1,
    1,
    ( final_Report_TW['本月指標客戶(終)']  )  / ( final_Report_TW['本月累積指標目標']  )
)



final_Report_TW['累積達成率'] = final_Report_TW['累積達成率'].fillna(0) 
# final_Report_TW['累積達成率'] = final_Report_TW['累積達成率'] * 100  
# final_Report_TW['累積達成率'] = final_Report_TW['累積達成率'].apply(lambda x: f"{x:.0f}%")# 轉換為百分比


Excel_台灣報表  = final_Report_TW.copy()

Excel_海外報表  = final_Report_OS.copy()
Excel_專案報表  = final_Report_Pro.copy()
end_time_計時 = time.time()
today_計時 = datetime.today().strftime('%Y-%m-%d')
execution_time = end_time_計時 - start_time_計時
hours, rem = divmod(execution_time, 3600)
minutes, seconds = divmod(rem, 60)

execution_time_info = f'{today_計時}***程式執行時間:{int(hours)}小時{int(minutes)}分{seconds:.2f}秒\n'
file_name = '外勤報表執行時間記錄.txt'

