import os
import time
import requests
import pandas as pd
from datetime import datetime
from bs4 import BeautifulSoup

keyword = '室內設計'
area = 6001001000  # 初始地區

job_data = []

while area <= 6001017000:  # 你可以根據需要調整範圍
    page = 1  # 每次換區域時，頁數要重設為 1

    """取得最大頁數"""
    url = f'https://www.104.com.tw/company/search/?area={area}&keyword={keyword}&jobsource=n_my104_search&mode=s&page=1'
    res = requests.get(url)
    res.raise_for_status()
    soup = BeautifulSoup(res.text, 'html.parser')

    # **檢查是否有錯誤訊息**
    error_box = soup.find('div', class_="popup-box")
    if error_box or "無法正常提供資訊" in soup.text or "異動資料失敗" in soup.text:
        print(f"⚠️ 地區 {area} 無法取得資料，跳過...")
        area += 1000  # 直接跳下一個區域
        continue

    page_options = soup.select('select.form-control.p-0.h3.h-auto.font-weight-bold option')
    max_page = max(int(option['value']) for option in page_options if option['value'].isdigit()) if page_options else 1
    print(f"📍 目前爬取區域: {area}, 最大頁數: {max_page}")

    while page <= max_page:
        time.sleep(0.08)
        url = f'https://www.104.com.tw/company/search/?area={area}&keyword={keyword}&jobsource=n_my104_search&mode=s&page={page}'

        try:
            res = requests.get(url)
            res.raise_for_status()
            soup = BeautifulSoup(res.text, 'html.parser')

            articles = soup.find_all('div', class_="company-list px-4 px-md-6 py-4 border-bottom company-lists__item")


            if not articles:
                print(f"✅ {area} 沒有更多職缺，結束該區域爬取。")
                break  

            for article in articles:
                company_tag = article.find('a', class_=["company-name-link--pc", "company-name-link--mobile"])
                company_name = company_tag.text.strip() if company_tag else "未知公司"
                company_url = company_tag.get('href', '#') if company_tag and company_tag.get('href') else "#"

                # 嘗試取得所有 `span` 資訊
                info_tags = article.find_all('span')
                company_area = info_tags[0].text.strip() if len(info_tags) > 0 else "未知地區"
                company_type = info_tags[1].text.strip() if len(info_tags) > 1 else "未知產業"
                company_capital = info_tags[2].text.strip() if len(info_tags) > 2 else "未知資本額"
                company_number = info_tags[3].text.strip() if len(info_tags) > 3 else "未知人數"

                description_tag = article.find('div', class_="company-list__description")
                company_info = description_tag.text.strip() if description_tag else "無描述"

                job_data.append({
                    "關鍵字": keyword,
                    "公司名稱": company_name,
                    "公司網址": company_url,
                    "公司地區": company_area,
                    "產業類別": company_type,
                    "資本額": company_capital,
                    "員工人數": company_number,
                    "公司資訊": company_info
                })

                print(f"📌 {area} | 第 {page} / {max_page} 頁 ... {company_name} | {company_area} | {company_type} | {company_capital} | {company_number} | {company_url}")
                time.sleep(0.08)

            page += 1  # 只有頁數達到 max_page 才換區域

        except Exception as e:
            print(f"⚠️ 爬取 {area} 的第 {page} 頁時發生錯誤: {e}")
            break

    # **當某個區域的所有頁數完成後，才增加 area**
    area += 1000

# **最後存 Excel**
today_date = datetime.today().strftime('%Y-%m-%d')
os.makedirs(output_folder, exist_ok=True)
output_file = os.path.join(output_folder, f'{keyword}_{today_date}.xlsx')
df = pd.DataFrame(job_data)
df.to_excel(output_file, index=False)
print(f"📄 已儲存所有資料到 {output_file}")
