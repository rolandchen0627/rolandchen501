import k_event_filters as kd
import pandas as pd
import json
import requests
import time
import threading
import pyodbc
from concurrent.futures import ThreadPoolExecutor
import ast
import numpy as np
from sqlalchemy import create_engine
from datetime import datetime, timedelta, date
import os

from tw_token import get_access_token

ac_token = get_access_token()

print(ac_token)


date_0 = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
date_0_scrm = int(date_0.timestamp() * 1000)

date_2 = datetime.now() - timedelta(days=2)
date_2_scrm = int(date_2.timestamp() * 1000)

date_3 = datetime.now() - timedelta(days=3)
date_3_scrm = int(date_3.timestamp() * 1000)

date_7 = datetime.now() - timedelta(days=7)
date_7_scrm = int(date_7.timestamp() * 1000)

date_10 = datetime.now() - timedelta(days=10)
date_10_scrm = int(date_10.timestamp() * 1000)

date_14 = datetime.now() - timedelta(days=14)
date_14_scrm = int(date_14.timestamp() * 1000)

date_30 = datetime.now() - timedelta(days=30)
date_30_scrm = int(date_30.timestamp() * 1000)

date_60 = datetime.now() - timedelta(days=60)
date_60_scrm = int(date_60.timestamp() * 1000)

date_90 = datetime.now() - timedelta(days=90)
date_90_scrm = int(date_90.timestamp() * 1000)

date_6m = datetime.now() - timedelta(days=6*30)
date_6m_scrm = int(date_6m.timestamp() * 1000)

'''
電訪人員
'''
url_2 = "https://api-p10.xiaoshouyi.com/rest/data/v2.0/query/xoqlScroll"
headers = {
    "Authorization": f"Bearer {ac_token}",
    "Content-Type": "application/x-www-form-urlencoded"
    # Replace with your actual access token
}
queryLocator = ''
call_center = pd.DataFrame()

while True:
    data = {
        "xoql": '''SELECT customItem1__c.name as name, customItem1__c as EmplID, customItem3__c as call_number, customItem5__c as region,
                          dimDepart.departName
                            from customEntity42__c
                   ''',
        "batchCount": 2000,
        "queryLocator": queryLocator
    }
    response = requests.post(url_2, headers=headers, data=data)
    crm = response.json()
    data = pd.DataFrame(crm["data"]["records"])
    call_center = pd.concat([call_center, data], ignore_index=True, sort=False)

    if not crm['queryLocator']:
        break
    queryLocator = crm['queryLocator']
call_center['region'] = call_center['region'].str.get(0)
call_center = call_center[call_center['region'] != '台灣']
call_center = call_center[~call_center['name'].str.contains('*', regex=False)]
# call_center = call_center[call_center['region']=='香港']
# call_center = call_center[(call_center['name'] == 'NGUYEN THI LE (阮雲麗)')]

region_to_abbr = {
    '印尼': 'ID',
    '美國': 'US',
    '香港': 'HK',
    '菲律賓': 'PH',
    '越南': 'VN',
    '新加坡': 'SG',
    '泰國': 'TH',
    '印度': 'IN',
    '韓國': 'KR',
    '日本': 'JP',
    '英國': 'GB'
}

# final_3.loc[final_3['DataRegion']=='印度' , 'DataRegion'] = 'IN'
# first_nopick.loc[first_nopick['DataRegion']=='印度' , 'DataRegion'] = 'IN'


'''
select CompanyContact data
'''
url_2 = ""
headers = {
    "Authorization": f"Bearer {ac_token}",
    "Content-Type": ""
    # Replace with your actual access token
}
queryLocator = ''
comp_cont = pd.DataFrame()

while True:
    data = {
        "xoql": '''
            SELECT id,
                   entityType,
            WHERE customItem37__c  NOT LIKE 'TW%' 
            AND customItem37__c  NOT LIKE 'GB%' 
            AND customItem37__c  NOT LIKE 'JP%' 
            AND customItem37__c NOT LIKE 'KDED%'
            AND (customItem5__c LIKE '%C%'
            OR customItem5__c LIKE '%D%')
            ''',
        "batchCount": 2000,
        "queryLocator": queryLocator
    }
    response = requests.post(url_2, headers=headers, data=data)
    crm = response.json()
    data = pd.DataFrame(crm["data"]["records"])
    comp_cont = pd.concat([comp_cont, data], ignore_index=True, sort=False)

    if not crm['queryLocator']:
        break
    queryLocator = crm['queryLocator']
comp_cont = comp_cont[~comp_cont['CoFullName'].str.contains('Dup>')]
print(comp_cont['DataRegion'].unique())
comp_cont.value_counts('entityType')
time.sleep(2)
# JobtypeID
comp_cont['JobTypeID'] = comp_cont['JobTypeID'].str.get(0).str[:3]
comp_cont['JobTypeID'] = comp_cont['JobTypeID'].fillna('')
comp_cont = comp_cont[comp_cont['JobTypeID'].isin(
    ["001", "002", "003", "004", "005", "006", "007", "010", "011", "015", ""])]
# Contact Mobile
comp_cont['MobilePhone'] = comp_cont['MobilePhone'].str.replace(
    r'\D', '', regex=True)
comp_cont.loc[comp_cont['MobilePhone'].str.contains(
    "000000"), 'MobilePhone'] = ""
comp_cont.loc[comp_cont['MobilePhone'].str.len() < 6, 'MobilePhone'] = ""

# Company Phone
comp_cont['Company_Phone'] = comp_cont['Company_Phone'].str.replace(
    r'\D', '', regex=True)
comp_cont.loc[comp_cont['Company_Phone'].str.len() < 7, 'Company_Phone'] = ""

# Co Name
wordlist = "Closed|搬遷|倒閉|歇業|停業|轉行|退休|過世|廢止|解散|燈箱|群組|支援|留守|教育訓練|無效拜訪|資料不全"
comp_cont = comp_cont[~comp_cont['CoFullName'].str.contains(wordlist)]
comp_cont = comp_cont[~comp_cont['CoShortName'].str.contains(wordlist)]

# employed
comp_cont.sort_values(by='id', ascending=False, inplace=True)
comp_cont.drop_duplicates(subset='contact_code', keep='first', inplace=True)
comp_cont['employed'] = comp_cont['employed'].str.get(0)
comp_cont.loc[comp_cont['employed'].str.contains(
    "離職|离职|Resign", regex=True, na=False), 'Company_Phone'] = ""

# Invalid
comp_cont['Invalid'] = comp_cont['Invalid'].str.get(0)
comp_cont['Inactive'] = comp_cont['Inactive'].str.get(0)
comp_cont['Invalid'] = comp_cont['Invalid'].fillna("否")
comp_cont['停機'] = comp_cont['停機'].fillna("0")
comp_cont['空號'] = comp_cont['空號'].fillna("0")
comp_cont['Inactive'] = comp_cont['Inactive'].fillna("否")
comp_cont.loc[(comp_cont['Invalid'] == "是") | (comp_cont['停機'] ==
                                               "1") | (comp_cont['空號'] == "1"), 'MobilePhone'] = ""
comp_cont.loc[comp_cont['Inactive'] == "是", 'Company_Phone'] = ""

# SAP credit
# comp_cont.loc[comp_cont['SAP_CompanyID'].isin(credit_data[credit_data['SAP信用管制'] == "Restrain"]['SAP公司代號']), 'Company_Phone'] = ""

# 勿擾
comp_cont['勿擾'] = comp_cont['勿擾'].fillna('')
comp_cont['勿擾'] = comp_cont['勿擾'].astype(str)
comp_cont = comp_cont[~comp_cont['勿擾'].str.contains("勿電訪|勿傳送簡訊")]

# Remove blank mobile/phone and remove duplicate
comp_cont = comp_cont[~((comp_cont['MobilePhone'] == "")
                        & (comp_cont['Company_Phone'] == ""))]
comp_cont['LastWorkDate'] = comp_cont['LastWorkDate'].replace('', np.nan)
comp_cont['LastWorkDate'] = comp_cont['LastWorkDate'].astype(float)
comp_cont['LastWorkDate'] = comp_cont['LastWorkDate'].apply(
    lambda x: pd.to_datetime(x / 1000.0, unit='s', utc=True))
comp_cont['LastWorkDate'] = comp_cont['LastWorkDate'].dt.tz_convert(
    'Asia/Taipei')
comp_cont['LastWorkDate'] = comp_cont['LastWorkDate'].dt.strftime(
    '%Y-%m-%d %H:%M:%S')
comp_cont_mobile = comp_cont[comp_cont['MobilePhone'] != ""]
comp_cont_no_mobile = comp_cont[comp_cont['MobilePhone'] == ""]

comp_cont_mobile = comp_cont_mobile.sort_values(
    by='LastWorkDate', ascending=False)
comp_cont_mobile = comp_cont_mobile.drop_duplicates(subset=['MobilePhone'])
comp_cont_no_mobile = comp_cont_no_mobile.sort_values(
    by='LastWorkDate', ascending=False)
comp_cont_no_mobile = comp_cont_no_mobile.drop_duplicates(
    subset=['displayname', 'Company_Phone'])

tab = comp_cont_no_mobile['Company_Phone'].value_counts().reset_index()
tab.columns = ['Company_Phone', 'Freq']
comp_cont_no_mobile_nodup = comp_cont_no_mobile[comp_cont_no_mobile['Company_Phone'].isin(
    tab[tab['Freq'] < 2]['Company_Phone'])]
comp_cont_no_mobile_dup = comp_cont_no_mobile[~comp_cont_no_mobile['Company_Phone'].isin(
    tab[tab['Freq'] < 2]['Company_Phone'])]


comp_cont_final = pd.concat([comp_cont_mobile, comp_cont_no_mobile_nodup])
print(comp_cont_final['DataRegion'].unique())


'''
special exclude
'''

comp_cont_final = comp_cont_final[~comp_cont_final['ContactID'].isin(
    ["KHLXRGX20231102114151", "KHLXRGX2023110298720", "KHLXRGX2023110228629", "KHLXRGX2023110221895", "KHLXRGX202310042112", "KHLXRGX2023110234140", "KHLXRGX2023110225511", "KHLXRGX2023110293730"])]
comp_cont_final = comp_cont_final[~comp_cont_final['CompanyID'].isin(
    ["GAC00102921", "GAC00096758", "GAC00075392", "GAC00071671", "GAC00074439", "GAC00096845", "GAC00067286", "GAC00096812"])]
comp_cont_final = comp_cont_final.sort_values(by='LastWorkDate')

comp_cont_final['Operator'] = pd.NA
comp_cont_final['EmplID'] = pd.NA


#  呼叫函式


# ----1.	排除三個月內曾拜訪過 K 大的關係聯絡人。---------##(K大參會人)(K大預約表)
comp_cont_final, OS_K_SCRM_in_two = kd.filter_reserve_mrk(comp_cont_final)
comp_cont_final, K_visit_df1_in_two = kd.filter_k_3M(comp_cont_final)

# ----2.	排除六個月內曾於新加坡拜訪過的客戶。---------##(追蹤紀錄)
comp_cont_final, TWOS_tracking = kd.sg_visits_6M(comp_cont_final)

# ----3.	排除不同邏輯---------##(海外交辦)
# o	排除 [電訪類型] 為「無效」
# o	排除建檔日 90 天內 [是否邀約K大] 為「已完成邀約」
# o	排除建檔日 180 天內 [電訪類型] 為「客戶要求勿聯繫」（半年一次聯繫）
# o	若為寄送型電訪，建檔日 2 週內非「未接」或「接通後掛電話」亦排除
# o	排除 14 天內【是否邀約K大】=”沒興趣、客戶拒絕邀約”
comp_cont_final, duplicate_records, duplicate_records_JP, calling_record_SCRM_pick, calling_record_10D = kd.filter_scrm_calling(
    comp_cont_final)


# ----4.  排除 7 天內出現下列狀況---------##(海外交辦)
# o	任務完成（不含已完成K大）
# o	接通後掛電話
comp_cont_final, exist_SCRM = kd.exclude_recent_contacts(comp_cont_final)


# ----5. 排除 2 天內未接---------##(海外交辦)
comp_cont_final, nopick_df = kd.Missed_2D(comp_cont_final)

# ----6. 排除 下一次聯絡日期---------##(海外交辦)
comp_cont_final, next_contact = kd.next_contact(comp_cont_final)

# ----7. 排除 追蹤紀錄---------##(追蹤紀錄)
# 觸客類型：B1電訪, 工作類別=K大邀約, 是否邀約K大 = 成功邀約K大
comp_cont_final, tr_df_in_two = kd.tr_df(comp_cont_final)


# 8.	曾做過 K 大者：
# o	主營城市，工作主旨為 Daily Invite Mr.K(M) – Ver. 2
# o	非主營城市，工作主旨為 Daily Invite Mr.K(N) – Ver. 2
# 9.	未做過 K 大者：
# o	主營城市，交辦主旨為 Daily Invite Mr.K(M)
# o	非主營城市，交辦主旨為 Daily Invite Mr.K(N)
# 10.	SG（新加坡）與 HK（香港）不區分主營與否，統一使用：
# o	Daily Invite Mr.K (M)、Daily Invite Mr.K (N)

comp_cont_final['customItem2__c'].replace("nan", np.nan, inplace=True)
comp_cont_final['customItem2__c'].fillna('Daily Invite Mr.K', inplace=True)

comp_cont_final = kd.update_invite_version(
    comp_cont_final, 'Daily Invite Mr.K- Ver. 2', city_pattern)
comp_cont_final = kd.update_invite_version(
    comp_cont_final, 'Daily Invite Mr.K', city_pattern)


'''
處理之前有過任務完成但沒有K大過
'''
oversea_df = pd.merge(comp_cont_final, calling_record_SCRM_pick,
                      left_on='ContactID', right_on='customItem14__c.name', how='left')
oversea_df = oversea_df.sort_values(by='LastWorkDate')
# oversea_df['DataRegion2'] = oversea_df['DataRegion'].str[:2]
# region_counts = oversea_df.groupby('DataRegion2').size().reset_index(name='count')

oversea_df['nopick'] = False
# Update nopick column based on condition
oversea_df.loc[oversea_df['customItem38__c'] == '未接', 'nopick'] = True
oversea_df['customItem4__c'] = np.where(oversea_df['nopick'], '2', '1')
final_1 = oversea_df[oversea_df['customItem14__c.name'].isna()]
final_2 = oversea_df[oversea_df['customItem38__c'] == '未接']
final_3 = oversea_df[~oversea_df['id'].isin(
    final_1['id']) & ~oversea_df['id'].isin(final_2['id'])]


def process_data_region(region):
    if region.startswith('KDED'):
        return region
    else:
        return region[:2]


final_1['DataRegion'] = final_1['DataRegion'].apply(process_data_region)
final_2['DataRegion'] = final_2['DataRegion'].apply(process_data_region)
final_3['DataRegion'] = final_3['DataRegion'].apply(process_data_region)

final_3 = final_3.sort_values(by='customItem7__c', ascending=True)
first_nopick = pd.concat([final_1, final_2], ignore_index=True)


final_3.loc[final_3['DataRegion'] == '香港', 'DataRegion'] = 'HK'
first_nopick.value_counts('DataRegion')
time.sleep(2)

'''
開始分配交辦, 分成兩個dataframe處理
'''
call_center['abbr'] = call_center['region'].apply(
    lambda x: region_to_abbr.get(x, ''))
call_center = call_center[~((call_center['call_number'] == '0') | (
    call_center['call_number'] == ''))]
call_center['call_number'] = call_center['call_number'].astype(int)
call_center1 = call_center.copy()
call_center2 = call_center.copy()
call_center1['call_number'] = call_center['call_number']*0.8
call_center2['call_number'] = call_center['call_number']*0.6
first_nopick = first_nopick.sort_values(
    by='LastWorkDate', ascending=True).reset_index(drop=True)
final_3 = final_3.sort_values(
    by='LastWorkDate', ascending=True).reset_index(drop=True)
'''
call_center1 for first_nopick dataframe
'''
total_cases = first_nopick.groupby(
    'DataRegion').size().reset_index(name='total_number')
total_call_numbers = call_center1.groupby(
    'abbr')['call_number'].sum().reset_index(name='total_call_numbers')
call_center1 = call_center1.merge(
    total_cases, left_on='abbr', right_on='DataRegion',  how='left')
call_center1 = call_center1.merge(total_call_numbers, on='abbr',  how='left')
condition = call_center1['total_number'] < call_center1['total_call_numbers']
# Calculate adjusted_number based on the condition
call_center1['adjusted_number'] = np.floor(
    call_center1['total_number'] * call_center1['call_number'] / call_center1['total_call_numbers'])
# Where total_number <= total_call_numbers, set adjusted_number equal to call_number
call_center1.loc[~condition, 'adjusted_number'] = call_center1['call_number']
'''
call_center2 for final_3 dataframe
'''
total_cases = final_3.groupby(
    'DataRegion').size().reset_index(name='total_number')
total_call_numbers = call_center2.groupby(
    'abbr')['call_number'].sum().reset_index(name='total_call_numbers')
call_center2 = call_center2.merge(
    total_cases, left_on='abbr', right_on='DataRegion',  how='left')
call_center2 = call_center2.merge(total_call_numbers, on='abbr',  how='left')
condition = call_center2['total_number'] < call_center2['total_call_numbers']
# Calculate adjusted_number based on the condition
call_center2['adjusted_number'] = np.floor(
    call_center2['total_number'] * call_center2['call_number'] / call_center2['total_call_numbers'])
# Where total_number <= total_call_numbers, set adjusted_number equal to call_number
call_center2.loc[~condition, 'adjusted_number'] = call_center2['call_number']

first_nopick = first_nopick.sample(frac=1).reset_index(drop=True)
'''
for the first call and the no_pick dataframe
'''
employee_calls = {}  # Initialize the dictionary

for index, row in call_center1.iterrows():
    # Initialize each employee's call count
    employee_calls[row['name']] = int(row['adjusted_number'])
    region = row['abbr']
    call_number = int(row['adjusted_number'])
    region_data = first_nopick[first_nopick['DataRegion'] == region].sort_values(
        by='LastWorkDate', ascending=True)
    remaining_calls = len(region_data)
    sum_number = sum(int(call_num)
                     for call_num in call_center1[call_center1['abbr'] == region]['adjusted_number'])
    same_region_people = call_center1[call_center1['abbr'] == region]['name']
    # Count of people in the same region, excluding the current person
    count = len(same_region_people)
    # Total number of people in the same region, including the current person
    total_people = count

    if sum_number < remaining_calls:
        cases_to_assign = min(call_number, len(region_data))
        if cases_to_assign > 0:
            employee_name = row['name']  # Get the employee name
            # Get indices of rows where Operator is not assigned
            employee_index = region_data[region_data['Operator'].isnull(
            )].index
            cases_for_employee = min(cases_to_assign, len(
                employee_index))  # Limit cases to available rows

            # Assign Operator and EmplID to the selected rows
            first_nopick.loc[employee_index[:cases_for_employee],
                             'Operator'] = employee_name
            first_nopick.loc[employee_index[:cases_for_employee],
                             'EmplID'] = row['EmplID']

            remaining_calls -= cases_for_employee
            employee_calls[employee_name] -= cases_for_employee
    else:
        # Distribute remaining calls evenly
        cases_to_assign = remaining_calls // total_people
        for employee_name in employee_calls:
            employee_index = region_data[region_data['Operator'].isnull(
            )].index
            cases_for_employee = min(cases_to_assign, len(employee_index))
            first_nopick.loc[employee_index[:cases_for_employee],
                             'Operator'] = employee_name
            first_nopick.loc[employee_index[:cases_for_employee],
                             'EmplID'] = row['EmplID']
            remaining_calls -= cases_for_employee
            employee_calls[employee_name] -= cases_for_employee
# id_df = first_nopick[first_nopick['DataRegion']=='ID']
'''
for the past is 
'''
employee_calls = {}  # Initialize the dictionary

for index, row in call_center2.iterrows():
    # Initialize each employee's call count
    employee_calls[row['name']] = int(row['adjusted_number'])
    region = row['abbr']
    call_number = int(row['adjusted_number'])
    region_data = final_3[final_3['DataRegion'] == region].sort_values(
        by='LastWorkDate', ascending=True)
    remaining_calls = len(region_data)
    sum_number = sum(int(call_num)
                     for call_num in call_center2[call_center2['abbr'] == region]['adjusted_number'])
    same_region_people = call_center2[call_center2['abbr'] == region]['name']
    # Count of people in the same region, excluding the current person
    count = len(same_region_people)
    # Total number of people in the same region, including the current person
    total_people = count

    if sum_number < remaining_calls:
        cases_to_assign = min(call_number, len(region_data))
        if cases_to_assign > 0:
            employee_name = row['name']  # Get the employee name
            # Get indices of rows where Operator is not assigned
            employee_index = region_data[region_data['Operator'].isnull(
            )].index
            cases_for_employee = min(cases_to_assign, len(
                employee_index))  # Limit cases to available rows

            # Assign Operator and EmplID to the selected rows
            final_3.loc[employee_index[:cases_for_employee],
                        'Operator'] = employee_name
            final_3.loc[employee_index[:cases_for_employee],
                        'EmplID'] = row['EmplID']

            remaining_calls -= cases_for_employee
            employee_calls[employee_name] -= cases_for_employee
    else:
        # Distribute remaining calls evenly
        cases_to_assign = remaining_calls // total_people
        for employee_name in employee_calls:
            employee_index = region_data[region_data['Operator'].isnull(
            )].index
            cases_for_employee = min(cases_to_assign, len(employee_index))
            final_3.loc[employee_index[:cases_for_employee],
                        'Operator'] = employee_name
            final_3.loc[employee_index[:cases_for_employee],
                        'EmplID'] = row['EmplID']
            remaining_calls -= cases_for_employee
            employee_calls[employee_name] -= cases_for_employee
final = pd.concat([first_nopick, final_3], ignore_index=True)
final.drop(['nopick', 'customItem14__c.name', 'customItem7__c',
           'customItem38__c'], axis=1, inplace=True)
final = final[final['Operator'].notnull()]
final.rename(columns={'id': 'customItem14__c',
             'EmplID': 'customItem1__c'}, inplace=True)
final = final[['customItem14__c', 'customItem8__c', 'dimDepart',
               'ownerId', 'customItem1__c', 'customItem4__c', 'customItem2__c']]
final['customItem7__c'] = date_0_scrm
index_length = len(final)
values = [['1', '3']] * index_length
final['customItem3__c'] = values
final['entityType'] = '3093637455518085'
final['customItem2__c'].replace("nan", np.nan, inplace=True)
