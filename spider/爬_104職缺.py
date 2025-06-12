import requests
from bs4 import BeautifulSoup
import time
import random
from datetime import datetime
import os
import pandas as pd

keyword = '數據分析'
area = 6001001000  # 初始地區
job_data = []

while area <= 6001003000:  
    page = 1  

    while True:
        time.sleep(random.uniform(0.5, 1))  
        url = f'https://www.104.com.tw/jobs/search/?keyword={keyword}&isJobList=1&jobsource=joblist_search&order=15&area={area}&page={page}'

        try:
            res = requests.get(url)
            res.raise_for_status()
        except requests.exceptions.RequestException as e:
            print(f"⚠️ 爬取第 {page} 頁時發生錯誤: {e}")
            break  

        soup = BeautifulSoup(res.text, 'html.parser')

        # **檢查是否有錯誤訊息**
        error_box = soup.find('div', class_="popup-box")
        if error_box or "無法正常提供資訊" in soup.text or "異動資料失敗" in soup.text:
            print(f"⚠️ 地區 {area} 無法取得資料，跳過...")
            area += 1000  
            break

        # 獲取最大頁數
        page_options = soup.select('select.form-control.p-0.h3.h-auto.font-weight-bold option')
        max_page = max(int(option['value']) for option in page_options if option['value'].isdigit()) if page_options else 1
        print(f"📍 目前爬取區域: {area}, 最大頁數: {max_page}")

        # 爬取職缺列表
        articles = soup.find_all('div', class_="container-fluid job-list-container py-4 bg-white")  
        time.sleep(random.uniform(0.5, 1))  

        if not articles:
            print(f"✅ 區域 {area} 無職缺，跳過...")
            area += 1000  
            break  

        for article in articles:
            job_tag = article.find('a', class_="info-job__text jb-link jb-link-blue jb-link-blue--visited h2")
            job_title = job_tag.text.strip() if job_tag else "未知職缺"
            job_url = job_tag.get('href', '#') if job_tag and job_tag.get('href') else "#"

            company_tag = article.find('a', class_="info-company__text jb-link jb-link-blue jb-link-blue--visited h4")
            company_name = company_tag.text.strip() if company_tag else "未知公司"
            company_url = company_tag.get('href', '#') if company_tag and company_tag.get('href') else "#"

            info_tags = article.find_all('div', class_="info-tags gray-deep-dark")
            if info_tags:
                job_info = [tag.text.strip() for tag in info_tags[0].find_all('span')]
                company_area = job_info[0] if len(job_info) > 0 else "未知地區"
                experience = job_info[1] if len(job_info) > 1 else "未知經歷"
                educational = job_info[2] if len(job_info) > 2 else "未知學歷"
                salary = job_info[3] if len(job_info) > 3 else "不明薪資"
            else:
                company_area, experience, educational, salary = "未知", "未知", "未知", "未知"

            description_tag = article.find('div', class_="info-description text-gray-darker t4 text-break mt-2 position-relative info-description__line2")
            company_info = description_tag.text.strip() if description_tag else "無描述"

            job_data.append({
                "關鍵字": keyword,
                "職缺名稱": job_title,
                "職缺連結": job_url,
                "公司名稱": company_name,
                "公司地區": company_area,
                "經歷": experience,
                "學歷": educational,
                "薪資": salary,
                "公司資訊": company_info
            })
            time.sleep(random.uniform(0.5, 1))  
            print(f"📌 {job_title} | {company_name} | {company_area} | {experience} | {educational} | {salary} | {company_info}")

        page += 1  # 翻頁
        if page > max_page:
            break  

    area += 1000  

# **最後存 Excel**
today_date = datetime.today().strftime('%Y-%m-%d')
output_folder = r'C:\Users\11020964.TPTWKD\Desktop\python\python\104'
os.makedirs(output_folder, exist_ok=True)
output_file = os.path.join(output_folder, f'{keyword}_{today_date}.xlsx')
df = pd.DataFrame(job_data)
df.to_excel(output_file, index=False)
date_0  = datetime.now().strftime('%Y-%m-%d')
df.to_csv(fr'C:\Users\11020964.TPTWKD\Desktop\python\python\104\{date_0}_HRS_visit.txt', index=False, header=False, encoding='utf-8')
print(f"📄 已儲存所有資料到 {output_file}")
