

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
driver.get("https://addon.jinhakapply.com/RatioV1/RatioH/Ratio11201011.html")
time.sleep(10)

headers = driver.find_elements(By.CSS_SELECTOR, "table thead th")
col_count = len(headers)

rows = driver.find_elements(By.CSS_SELECTOR, 'table tbody tr')

current_dept = ""  # 단과대명
current_num = "" # 모집인원

for row in rows:
    tds = row.find_elements(By.TAG_NAME, "td")
    values = [td.text.strip() for td in tds]

    # 열이 부족하면 이전 값 유지
    if len(values) == 5:
        dept, major, num, applicants, ratio = values
        current_dept, current_num = dept, num
    elif len(values) == 4: # 단과대학이 rowspan 된 경우
        major, num, applicants, ratio = values
        dept = current_dept
    elif len(values) == 3: # 단과대학과 모집인원이 rowspan 된 경우인데 학과 열이 2개일 수도 있어서 확인 필요
        major, applicants, ratio = values
        dept = current_dept
        num = current_num
     # elif len(values) == 6: 열 2개를 합쳐야 함 (예: 음악대학>한국음악과>거문고)
    else:
        continue  # 총계 행 건너뛰기
        # 데이터 누적
    data.append([dept, major, num, applicants, ratio])

    '''print(dept, major, num, applicants, ratio)'''



#pandas DataFrame으로 변환
df_old = pd.DataFrame(data, columns=['단과대명', '학과명', '모집인원', '지원인원', '경쟁률'])

# 현재 시각을 열 이름으로
now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
df_new = df_old.rename(columns={"경쟁률": now})

# 단과대명과 학과명, 모집인원을 붙여 key 만들기
df_new["_key"] = df_new["단과대명"].astype(str) + " " + df_new["학과명"].astype(str) + " " + df_new["모집인원"].astype(str)

filename = r"C:\파이썬 아이들\scr_prc\이화여자대학교_2025_수시_경쟁률_추이.xlsx"


# 기존 파일 불러와 새 데이터와 합침
if os.path.exists(filename):
    df_prev = pd.read_excel(filename)
    
    # 기존 데이터도 병합용 키 만들기
    df_prev["_key"] = df_prev["단과대명"].astype(str) + " " + df_prev["학과명"].astype(str) + " " + df_new["모집인원"].astype(str)

    # 병합 
    df_merged = pd.merge(
        df_prev, 
        df_new,
        on=["_key", "단과대명", "학과명", "모집인원", "지원인원"],
        how="outer"
    )

else:
    df_merged = df_new

# 최종 결과에서 모집인원, 지원인원 제거
df_final = df_merged.drop(columns=["모집인원", "지원인원"], errors="ignore")

# 저장은 df_final로
df_final.to_excel(filename, index=False)

'''df_merged = df_merged.drop(columns="_key")'''

# 위 과정 엑셀 저장
'''df_merged.to_excel(filename, index=False, engine="openpyxl")'''

print(f"{now} 시각의 데이터가 추가되었습니다.")

#브라우저 닫기
driver.quit()