#clean environment
from IPython import get_ipython
get_ipython().magic('reset -f')
import pandas as pd
import json
import requests
import time
import threading
from concurrent.futures import ThreadPoolExecutor
import numpy as np
import pymysql
from sqlalchemy import create_engine, Table, MetaData, Column, String, inspect, Text

# 客戶分級標籤自動化工具

# ## 📌 專案簡介
# 針對集團內部客戶資料，根據交易記錄與聯繫情形，進行條件式分群與標籤分類，協助業務團隊辨識經營、開發與沉默客戶。

# ## 📂 使用技術
# - Python：資料清洗、條件判斷
# - pandas：資料處理
# - SQLAlchemy / PyMySQL：連接 MySQL 資料庫
# - Excel：最終報表輸出

# ## 🔍 分析流程簡述
# 1. **資料載入與清洗**：讀取 Excel 客戶明細與資料庫中公司主檔
# 2. **條件式分群邏輯**：
#    - 經營客戶：近三年有交易記錄或為主管指定名單
#    - 開發中：一年內曾聯繫或有送樣/拜訪紀錄
#    - 開發客戶：無近期交易紀錄
#    - 沉默客戶：開發客戶中近半年聯絡不上者
# 3. **上層關聯比對**：處理 accountCode__c 與 COMPANYID 上下層關係
# 4. **輸出 Excel**：將結果存為 `.xlsx` 報表供業務使用

# ## 📁 輸出結果
# 最終產出檔案：`0219經營開發客戶標籤.xlsx`，每家公司對應一個標籤欄位（經營/開發/沉默）


c_df = pd.read_excel('集團客戶型態統計數據_0219.xlsx', sheet_name = '(集團客戶型態統計表) 公司明細')
c_df = c_df[c_df['CompanyTypeID']!='DP']
c_df = c_df[c_df['accountCode__c'].str.contains('GAC')]
c_df = c_df[['accountCode__c', 'CompanyTypeID', '近三年交易_項目成交', '訂單重點人物', '業務主管指定名單', '服務費設計師', '近一年完成K大或拜訪', '近一年送樣', '項目_詢價', '近半年聯繫不上']]

condition_1 = (c_df['CompanyTypeID'].str.contains('C|D') &
              (c_df[['近三年交易_項目成交', '訂單重點人物', '業務主管指定名單', '服務費設計師']].any(axis=1))
)
# Condition 2
condition_2 = (c_df['CompanyTypeID'].str.contains('C|D') &
               (c_df[['近一年完成K大或拜訪', '近一年送樣', '項目_詢價']].any(axis=1)))

# Condition 3
condition_3 = (c_df['CompanyTypeID'].str.contains('C|D') &
               (~c_df[['近三年交易_項目成交', '訂單重點人物', '業務主管指定名單', '服務費設計師', '近一年完成K大或拜訪', '近一年送樣', '項目_詢價']].any(axis=1)))

# Assign values based on conditions
c_df.loc[condition_1, 'mark'] = '經營客戶'
c_df.loc[condition_2 & c_df['mark'].isna(), 'mark'] = '開發中'
c_df.loc[condition_3 & c_df['mark'].isna(), 'mark'] = '開發客戶'
#condition_silence = (c_df['mark'] == '開發客戶') & (c_df['近半年聯繫不上'] == True)
#c_df.loc[condition_silence, 'mark'] = '沉默客戶'

'''
mysql
'''
# 数据库连接信息
host = ''
port = 
user = ''
password = ''
database = ''

# 创建到 MySQL 的连接引擎


# 读取 related_company 表的数据
account_df = pd.read_sql_table('related_company', engine)
account_df = account_df[account_df['RELATED_FINAL'].str.contains('GAC')]
account_df = account_df[['COMPANYID', 'COMPANYTYPE', 'RELATED_FINAL', 'COMPANYTYPE_FINAL']]
account_df.rename(columns={'RELATED_FINAL': 'accountCode__c'}, inplace=True)
c_df = pd.merge(c_df, account_df, on='accountCode__c', how='left')
c_df['third_mark'] = c_df['mark']

no_up_down = c_df[c_df['accountCode__c']==c_df['COMPANYID']]
has_up_down = c_df[c_df['accountCode__c']!=c_df['COMPANYID']]
third_up_down = has_up_down[['COMPANYID', 'third_mark', '近半年聯繫不上']]
third_up_down.rename(columns={'COMPANYID': 'accountCode__c', 'third_mark':'mark'}, inplace=True)
no_up_down = no_up_down[~no_up_down['mark'].isna()]
no_up_down = no_up_down[['COMPANYID', 'mark', '近半年聯繫不上']]
no_up_down.rename(columns={'COMPANYID': 'accountCode__c'}, inplace=True)
final_1 = pd.concat([third_up_down, no_up_down], ignore_index=True)
final_1 = final_1.drop_duplicates(subset=['accountCode__c'], keep='first')
final_1 = final_1[~final_1['mark'].isna()]

'''
two_yes
'''
exist_account_df = pd.read_sql_table('related_company', engine)
exist_account_df = exist_account_df[exist_account_df['RELATED_FINAL'].str.contains('GAC')]
two_yes = exist_account_df[exist_account_df['COMPANYTYPE'].str.contains('C|D', na=False) & ~exist_account_df['COMPANYTYPE_FINAL'].str.contains('C|D', na=False)]
two_yes = pd.merge(two_yes, c_df, on = 'COMPANYID', how = 'left')
two_yes = two_yes[~two_yes['accountCode__c'].isna()]
two_yes = two_yes[['COMPANYID', 'COMPANYTYPE_x', 'RELATED_FINAL', 'COMPANYTYPE_FINAL_x', 'accountCode__c', 'CompanyTypeID', 
                   '近三年交易_項目成交', '訂單重點人物', '業務主管指定名單', '服務費設計師', '近一年完成K大或拜訪', '近一年送樣', '項目_詢價', '近半年聯繫不上']]

condition_4 = (two_yes[['近三年交易_項目成交', '訂單重點人物', '業務主管指定名單', '服務費設計師']].any(axis=1))

# Condition 2
condition_5 = (two_yes[['近一年完成K大或拜訪', '近一年送樣', '項目_詢價']].any(axis=1))

# Condition 3
condition_6 = (~two_yes[['近三年交易_項目成交', '訂單重點人物', '業務主管指定名單', '服務費設計師', '近一年完成K大或拜訪', '近一年送樣', '項目_詢價']].any(axis=1))

# Assign values based on conditions
two_yes.loc[condition_4, 'mark'] = '經營客戶'
two_yes.loc[condition_5 & two_yes['mark'].isna(), 'mark'] = '開發中'
two_yes.loc[condition_6 & two_yes['mark'].isna(), 'mark'] = '開發客戶'

#condition_silence = (two_yes['mark'] == '開發客戶') & (two_yes['近半年聯繫不上'] == True)
#two_yes.loc[condition_silence, 'mark'] = '沉默客戶'
final_2 = two_yes[['COMPANYID', '近半年聯繫不上']]
final_2 = final_2.drop_duplicates(subset=['COMPANYID'], keep='first')
final_2 = final_2.rename(columns={'COMPANYID': 'accountCode__c'})
final_df = pd.concat([final_1, final_2], ignore_index=True)
final_df = final_df.drop_duplicates(subset='accountCode__c', keep='first')
condition_silence = (final_df['mark'] == '開發客戶') & (final_df['近半年聯繫不上'] == True)
final_df.loc[condition_silence, 'mark'] = '沉默客戶'
final_df.to_excel('0219經營開發客戶標籤.xlsx', index = False)

