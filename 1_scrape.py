import os
import csv
import time
import urllib.parse
import urllib.request
import hashlib

from time import sleep
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options

# ChromeDriverのパスを指定
driver_path = "./chromedriver"

# ヘッドレスモードの設定
options = Options()
# options.add_argument("--headless")
options.add_argument('window-size=1200x600')

# Chromeを起動
service = Service(executable_path=driver_path)
driver = webdriver.Chrome(service=service, options=options)

driver.get("https://jinzai.hellowork.mhlw.go.jp/JinzaiWeb/GICB101010.do?action=initDisp&screenId=GICB101010")

time.sleep(2)

driver.execute_script("doPostAction('transition', 1)")

time.sleep(2)

driver.find_element(By.ID, "ID_cbZenkoku1").click()
driver.find_element(By.ID, "ID_cbJigyoshoKbnYu1").click()

driver.execute_script("ScrollEvent()")

time.sleep(6)

headers = ["ID", "許可・届出受理番号", "許可届出受理年月日", "事業主名称", "事業所名称", "事業所所在地", "電話番号", "取扱職種の範囲等-取扱職種", "取扱職種の範囲等-取扱地域", "取扱職種の範囲等-その他", "得意とする職種", "参考情報（得意職種等）", "手数料", "返戻金制度", "備考"]

def write_data_to_csv(driver):
    table = driver.find_element(By.ID, "search_area").find_element(By.TAG_NAME, "table")
    data_elements = table.find_elements(By.CLASS_NAME, "searchDet_data")
    data = [element.text.replace('\n', ' ') for element in data_elements]  # 改行をスペースに置換

    # ユニークなIDを生成
    unique_id_string = data[0] + data[2] + data[3]  # "許可・届出受理番号", "事業主名称", "事業所名称"
    unique_id = hashlib.md5(unique_id_string.encode()).hexdigest()

    # ユニークなIDをデータに追加
    data.insert(0, unique_id)

    # PDFファイルのダウンロード
    pdf_links = table.find_elements(By.XPATH, ".//a[contains(@href, '.pdf')]")
    for i, link in enumerate(pdf_links):
        try:
            pdf_url = link.get_attribute('href')
            pdf_file = os.path.join('downloads', f"{unique_id}-{i+1}.pdf")
            urllib.request.urlretrieve(pdf_url, pdf_file)
        except Exception as e:
            print(f"Failed to download file from {pdf_url}. Error: {e}")

    with open('output.csv', 'a', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(data)

# CSVファイルに見出し行を書き込む
with open('output.csv', 'w', newline='') as f:
    writer = csv.writer(f)
    writer.writerow(headers)

total_count = int(driver.find_element(By.ID, "ID_lbSearchCount").text)
print(f"Total count: {total_count}")
visited_links = set()
page_number = 1

while total_count > 0:
    try:
        link_elements = driver.find_elements(By.XPATH, "//a[@name='linkKyokatodokedeNo']")
        visited_on_page = 0
        for link in link_elements:
            href = link.get_attribute('href')
            visited_on_page += 1
            if href not in visited_links:
                visited_links.add(href)
                main_window = driver.window_handles[0]
                link.click()
                time.sleep(2)
                new_window = driver.window_handles[1]
                driver.switch_to.window(new_window)
                write_data_to_csv(driver)
                driver.close()
                driver.switch_to.window(main_window)
                time.sleep(2)
                total_count -= 1
                print(f"Remaining count: {total_count}")
        if len(link_elements) == visited_on_page:
            print(f"Moving to page {page_number}")
            driver.execute_script(f"doPostAction('page', '{page_number}')")
            page_number += 1
            time.sleep(2)
    except:
        break