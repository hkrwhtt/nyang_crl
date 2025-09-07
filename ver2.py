

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
import time
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import datetime
import os

options = Options()
#크롬 드라이버 실행
options.add_argument("--headless")
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

#데이터 리스트 만들기
data = []

#웹페이지 열기
driver.get("https://apply.jinhakapply.com/SmartRatio/PastRatioUniv?univid=1140&year=2025&category=1")
time.sleep(10)

rows = driver.find_elements(By.CSS_SELECTOR, 'table tbody tr')
for row in rows:
    '''cols = row.find_elements(By.TAG_NAME, 'p.detail')
    cols += row.find_elements(By.TAG_NAME, 'td.rate')
    print([col.text for col in cols])'''
    
    #학과, 전형
    dept = row.find_element(By.CSS_SELECTOR, "p.detail").text
    #경쟁률
    rate = row.find_element(By.CSS_SELECTOR, "td.rate").text
    data.append([dept, rate])

#pandas DataFrame으로 변환
df_old = pd.DataFrame(data, columns=['학과/전형', '경쟁률'])

# 현재 시각을 열 이름으로
now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
df_new = df_old.rename(columns={"경쟁률": now})

filename = r"C:\파이썬 아이들\scr_prc\충남대학교_2025_경쟁률_추이.xlsx"


# 기존 파일 불러와 새 데이터와 합침
if os.path.exists(filename):
    df_old = pd.read_excel(filename)
    df_old["학과/전형"] = df_old["학과/전형"].str.replace("\s+", " ", regex=True).str.strip()
    df_new["학과/전형"] = df_new["학과/전형"].str.replace("\s+", " ", regex=True).str.strip()
    df_merged = pd.merge(df_old, df_new, on="학과/전형", how="outer")
else:
    df_merged = df_new

# 위 과정 엑셀 저장
df_merged.to_excel(filename, index=False, engine="openpyxl")
print(f"{now} 시각의 데이터가 추가되었습니다.")

#브라우저 닫기
driver.quit()