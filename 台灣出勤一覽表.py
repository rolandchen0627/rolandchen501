import pandas as pd  
import requests
import json
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import numpy as np
import glob
import os

today = datetime.today()

# 計算上個月26號
last_month_26 = (today.replace(day=1) - relativedelta(months=1)).replace(day=26)

# 計算本月25號
this_month_25 = (today.replace(day=1) - relativedelta(months=0)).replace(day=25)

# 轉換為 "YYYY/MM/DD" 格式
last_month_26_str = last_month_26.strftime("%Y/%m/%d")
this_month_25_str = this_month_25.strftime("%Y/%m/%d")

start_timestamp_of_month = int(last_month_26.timestamp() * 1000)
end_timestamp_of_month = int(this_month_25.timestamp() * 1000)


'''
login
'''
def fn_login(CompanyID, pwd):
    IP_adrs = ""
    url_login = f""
    headers = {"Content-Type": "application/json"}

    payload = {
        "Action": "Login",
        "Value": {
            "$type": "",
            "CompanyID": CompanyID,
            "UserID": "",
            "Password": pwd,
            "LanguageId": ""
        }
    }

    response = requests.post(url_login, headers = headers, json=payload)
    content = response.json()
    SessionGuid = content['SessionGuid']

    return SessionGuid

'''
all_staff_Resign
'''
def all_staff_Resign(SessionGuid, CompanyID):
    IP_adrs = ""
    url_obj = f""
    
    payload_obj = {
        "Action": "",
        "SessionGuid": SessionGuid,
        "ProgID": "",
        "Value": {
            "$type": "",
            "UIType": "",
            "ReportID": "",
            "ReportTailID": "",
            "FilterItems": [
              {
                    "$type": "",
                    "FieldName": "",
                    "FilterValue": CompanyID
                },
             {
                    "$type": "",
                    "FieldName": "",
                    "FilterValue":last_month_26_str ,
                    "ComparisonOperator": ""
                  },
                  {
                    "$type": "",
                    "FieldName": "",
                    "FilterValue": this_month_25_str,
                    "ComparisonOperator": ""
                  }
            ],
            "UserFilter": ""
        }
    }
    response = requests.post(url_obj, json=payload_obj)
    content_obj = response.json()
    all_staff_Resign  = pd.json_normalize(content_obj["DataSet"]["ReportBody"])
    return all_staff_Resign



pwd_TW = ""
CompanyID_TW = ""
SessionGuid_TW = fn_login(CompanyID_TW, pwd_TW)

### get all_staff_Resign
HRS_LEAVE_TW = all_staff_Resign(SessionGuid_TW, CompanyID_TW)

keyword = "人事資料表"
file_paths = glob.glob(r'C:\Users\人事資料表(共用)\*{}*.xlsx'.format(keyword))

if file_paths:
    file_path = file_paths[0]

    # 讀取 Excel（第一個符合的檔案）
    depart_df = pd.read_excel(file_path, sheet_name='人事資料表_全集團', index_col=False)
    depart_df = depart_df[
        (depart_df['在職狀態'] == '正式') &
        (depart_df['人事資料表']=='TW')
        ]
else:
    print("找不到符合的檔案")
    depart_df = pd.DataFrame()  

today = datetime.today().replace(hour=0, minute=0, second=0, microsecond=0)


def calc_experience(start_date):
    if pd.isnull(start_date):
        return None
    rd = relativedelta(today, start_date)
    exp_year = rd.years + rd.months / 12
    return round(exp_year, 1)

# 計算工作年資
depart_df['工作年資'] = depart_df['到職日期'].apply(calc_experience)
depart_df['員工工號'] = depart_df['員工工號'].astype(str)

rename_dict = {
    'EMPLOYEEID': '員工工號',
    'EMPLOYEENAME': '員工姓名',
    'SYS_DATE': '單據日期',
    'SYS_VIEWID':'單據編號',
    'SVACATIONNAME': '假別類別',
    'CUTDATE': '開始日期',
    'ENDDATE': '結束日期',
    'SSTARTDATE':'報到日',
    'TMP_WORKPLACENAME':'工作地點',
    'JOBCODENAME':'職稱',
    'GSPECAILYEARS':'工作年資'

}



# TW
HRS_LEAVE_TW = HRS_LEAVE_TW.rename(columns=rename_dict)
HRS_LEAVE_TW = HRS_LEAVE_TW[(HRS_LEAVE_TW['SYS_FLOWFORMSTATUS']== 1) | (HRS_LEAVE_TW['SYS_FLOWFORMSTATUS']== 2) ]

final_dayoff = HRS_LEAVE_TW.copy()
final_dayoff = final_dayoff[['員工工號','員工姓名','單據日期','假別類別','開始日期','結束日期']]
final_dayoff['開始日期'] = pd.to_datetime(final_dayoff['開始日期']).dt.strftime('%Y/%m/%d')
global_staff = pd.merge(
    depart_df,
    final_dayoff[['員工工號', '單據日期', '假別類別', '開始日期', '結束日期']],
    on='員工工號',
    how='left'
)
global_staff.loc[global_staff['部門'].isna(), '部門'] = global_staff['單位']

# 引用設定好的字典(假別類別、職務、部門)
from leave_dict import get_leave_category, leave_category_dict 
from leave_dict import get_job_type, jobtype_dict
from leave_dict import get_leave_code, leave_dict
from leave_dict import get_depart_code, depart_dict

# 存放變數
leave_category = leave_category_dict  
job_type_dict = jobtype_dict  
leave_code_dict = leave_dict  
depart_dict = depart_dict

global_staff['未提前申請之特休'] = ""

global_staff['未提前申請之特休'] = np.where(
    (global_staff['假別類別'] == '特別休假') &  
    (global_staff['單據日期'] >= global_staff['開始日期']),
    9,  
    ""
)
global_staff.loc[global_staff['未提前申請之特休']== '9' , '假別類別'] ='未提前申請之特休'

global_staff["對應值"] = global_staff["假別類別"].map(leave_dict).fillna("無")
global_staff["對應值"] = global_staff["對應值"].map(leave_category_dict)
global_staff = global_staff[global_staff['人事資料表']=='TW']

pivot_df = global_staff.pivot_table(
    index=["部門", "員工工號", "中文姓名"],
    columns="對應值",
    aggfunc="size",
    fill_value=0
)

pivot_df = pivot_df.rename(columns=leave_category_dict).reset_index()
pivot_df = global_staff[["單位","部門", "員工工號","中文姓名","在職狀態","課別",'職務','工作地點名稱','工作年資']].drop_duplicates().merge(pivot_df, on=["部門","員工工號", "中文姓名"], how="left")

columns_to_sum = list(leave_category_dict.values())
# Ensure the columns exist in the DataFrame
columns_to_sum = [col for col in columns_to_sum if col in pivot_df.columns]
# Calculate the sum across the selected columns and store it in '合計'
pivot_df['合計'] = pivot_df[columns_to_sum].sum(axis=1)

pivot_df['不含特休小計']=""
if '特休' in pivot_df.columns:
    pivot_df['不含特休小計'] = pivot_df['合計'] - pivot_df['特休']
else:
    print("資料框中沒有 '特休' 欄位，請確認欄位名稱是否正確。")

required_columns = [
    '單位', '部門', "課別", "員工工號", "中文姓名", "職務", "工作地點名稱", "工作年資",	
    "合計", "不含特休小計",
    "病假", "生理", "家庭", "事假", "特休", "補休帶薪假", "其他", "未提前申請之特休"
]

# 如果缺少欄位，就補空值欄位
for col in required_columns:
    if col not in pivot_df.columns:
        pivot_df[col] = ""

# 重新排列欄位順序
pivot_df = pivot_df[required_columns]

pivot_df.rename(columns={'職務': '職稱','中文姓名':'員工姓名', '工作地點名稱': '工作地點'}, inplace=True)

base_path = r'C:\Users\CALENDARID\2025_Global_calendar.xlsx'
HOLIDAY_Global = pd.read_excel(base_path)
HOLIDAY_Global.rename(columns={'WDATE':'holiday_date','VACATIONNAME':'day_type'}, inplace=True)

#修改國家 
HOLIDAY_Global = HOLIDAY_Global[HOLIDAY_Global['國家']=='台灣']

HOLIDAY_Global = HOLIDAY_Global[['英文','holiday_date' , 'day_type']]
HOLIDAY_Global.loc[~HOLIDAY_Global['day_type'].isna() ,'day_type'] = 1
HOLIDAY_Global.loc[HOLIDAY_Global['day_type'].isna() ,'day_type'] = 0
HOLIDAY_Global['holiday_date'] = pd.to_datetime(HOLIDAY_Global['holiday_date'].astype(str), format='%Y%m%d')
HOLIDAY_Global['holiday_date'] = HOLIDAY_Global['holiday_date'].dt.strftime('%Y/%m/%d')


date_range = pd.date_range(start=last_month_26_str, end=this_month_25_str, freq="D")

df = pd.DataFrame({
    "日期": date_range,
    "星期": date_range.strftime('%a')  
})

# 将日期格式化为月/日
df['日期'] = df['日期'].dt.strftime('%Y/%m/%d')
df['type'] = ""

# 合併 DataFrame
merger_df = pd.merge(df, HOLIDAY_Global, left_on='日期', right_on='holiday_date', how='left')
# 判斷 day_type = 0 的條件，並設定 type = '休'
merger_df['type'] = merger_df['day_type'].apply(lambda x: '休' if x == 1 else "")
# 選擇需要的欄位
    
merger_df = merger_df[merger_df['英文']=='TW']    
result_df = merger_df[['日期', '星期', 'type']]
result_df =result_df.set_index('日期').T

result = pd.concat( [global_staff ,result_df]  )
result.rename(columns={'職務': '職稱','中文姓名':'員工姓名', '工作地點名稱': '工作地點'}, inplace=True)
reverse_leave_category_dict = {v: k for k, v in leave_category_dict.items()}
# 確保 "對應值" 欄位存在
if "對應值" in result.columns:
    result['假別類別分類'] = result['對應值'].map(reverse_leave_category_dict)
else:
    result['假別類別分類'] = None  # 或者用其他預設值
result['假別類別分類'] = result['對應值'].map(reverse_leave_category_dict).astype('Int64')
result['假別類別分類'] = result['假別類別分類'].astype(str).replace("<NA>", "未知")


df_grouped = result.groupby("員工工號").agg({
    "單據日期": lambda x: ', '.join(x.astype(str)),  # 日期轉字串用逗號連接
    "假別類別分類": lambda x: list(x),  # 保持為列表
    "開始日期": lambda x: list(x)      # 保持為列表
}).reset_index()

# 建立所有可能的「開始日期」
result["開始日期"] = pd.to_datetime(result["開始日期"], errors="coerce")  # 將開始日期轉為日期格式，無效值會變為 NaT
result = result.dropna(subset=["開始日期"])  # 移除無效的開始日期
result["開始日期"] = result["開始日期"].dt.strftime("%Y/%m/%d")  # 將日期格式化為 YYYY/MM/DD
result = result.sort_values(by="單據日期")  # 根據單據日期排序

    # 建立完整欄位順序（用 date_range）
date_str_list = [d.strftime('%Y/%m/%d') for d in date_range]
for date in date_str_list:
    df_grouped[date] = ""

# 根據請假資料填入 leave_type
for index, row in df_grouped.iterrows():
    for date, leave_type in zip(row["開始日期"], row["假別類別分類"]):
        if date in df_grouped.columns:
            df_grouped.at[index, date] = leave_type

# 移除舊的「開始日期」和「假別類別」欄位
df_grouped.drop(columns=["單據日期",'假別類別分類'], inplace=True)
df_grouped = df_grouped.merge(depart_df[['員工工號']] , on="員工工號", how='left')


print(df_grouped)

final_df = pivot_df.merge(df_grouped, on="員工工號", how="left")
final_df['合計'] = pd.to_numeric(final_df['合計'], errors='coerce')
final_df = final_df.sort_values(by='合計', ascending=False)

final_df = pd.concat([final_df,result_df])
week_col = final_df.loc['星期'].to_frame().T
type_col = final_df.loc['type'].to_frame().T

# 刪除 '星期' 和 'type' 的原索引
final_df = final_df.drop(index=['星期', 'type'])

final_df = pd.concat([final_df.iloc[:0], week_col, type_col, final_df]).reset_index(drop=True)
final_df = final_df.fillna("")
# 找出 index=1 (第二行) 為 "休" 的欄位
columns_to_fill = final_df.loc[1] == "休"
columns_to_fill = columns_to_fill[columns_to_fill].index.tolist() 

# 只對這些欄位補 "休"
for col in columns_to_fill:
    final_df[col] = final_df[col].replace("", "休").fillna("休")
final_df.columns = pd.to_datetime(final_df.columns, errors='ignore')

final_df = final_df.drop(index=1)

# 按照 leave_category_dict 對應欄位名稱來填
index_to_modify = 0  

# 遍歷 leave_category_dict，根據字典的 key 在 final_df 中更新對應欄位的值
for key, leave in leave_category_dict.items():
    if leave in final_df.columns:
        # 在 final_df 中找到對應的欄位，並填入 key（數字）
        final_df.at[index_to_modify, leave] = key

# 顯示結果
print(final_df)

final_df["幹部基層"]=""
final_df["幹部基層"] = final_df["職稱"].map(jobtype_dict).fillna("")
final_df["部門人數"] = final_df.groupby("部門")["部門"].transform("count")
final_df.loc[final_df["部門人數"] < 20, "幹部基層"] = ""
# final_df.loc[(final_df["部門人數"] >= 20) & (final_df["幹部基層"]==""), "幹部基層"] = "基層"
empty_depart_df = final_df[ ( final_df['部門人數'] >=20 ) & (final_df['幹部基層']=="" ) ]
final_df.drop(columns=["部門人數"], inplace=True)
# 調整欄位順序，把「幹部基層」與「部門人數」移到最前面
cols = final_df.columns.tolist()
cols.remove("幹部基層")
new_order = ["幹部基層"] + cols
final_df = final_df[new_order]
final_df['轉出'] =""
final_df['轉出'] = final_df['部門'].map(depart_dict).fillna("")
cols = final_df.columns.tolist()
cols.remove("轉出")
new_order = ["轉出"] + cols
final_df = final_df[new_order]
cols = ['合計', '不含特休小計', '病假', '生理', '家庭', '事假', '特休', '補休帶薪假', '其他', '未提前申請之特休']
final_df[cols] = final_df[cols].replace("", 0)
final_df[cols] = final_df[cols].replace(0, "")
file_path = r'C:\Users\Desktop\Github' 
file_final = f"{file_path}\\TW_每月出勤表.xlsx"
final_df.to_excel(file_final, index=False)