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

import tempfile

options = Options()
#크롬 드라이버 실행
options.add_argument("--headless")

options.add_argument("--no-sandbox")  # 샌드박스 끄기 (EC2에서 필수)
options.add_argument("--disable-dev-shm-usage")  # 메모리 문제 해결
options.add_argument("--disable-gpu")  # GPU 비활성화
options.add_argument(f"--user-data-dir={tempfile.mkdtemp()}")

driver = webdriver.Chrome(
    service=Service(ChromeDriverManager().install()),
    options=options
)

#데이터 리스트 만들기
data = []

#웹페이지 열기 (수정)
driver.get("https://addon.jinhakapply.com/RatioV1/RatioH/Ratio10080311.html")
time.sleep(10)

headers = driver.find_elements(By.CSS_SELECTOR, "table thead th")
col_count = len(headers)

rows = driver.find_elements(By.CSS_SELECTOR, 'table tbody tr')

data = []
current_jh = ""
current_dept = ""
current_major = ""
current_num = ""
current_ratio = ""

for row in rows:
    tds = row.find_elements(By.TAG_NAME, "td")
    if not tds:
        continue
    values = [td.text.strip() for td in tds]

    # 맨 첫 줄은 6개
    if len(values) == 6:
        jh, dept, major, num, applicants, ratio = values
        current_jh, current_dept, current_major, current_num = jh, dept, major, num
    elif len(values) == 4: # 전형과 단과대 로우스팬
        major, num, applicants, ratio = values
        dept = current_dept
        jh = current_jh
    elif len(values) == 5: # 전형명 로우스팬
        dept, major, num, applicants, ratio = values
        jh = current_jh
    elif len(values) == 3: # 전형명, 모집인원, 경쟁률 로우스팬
        major, dept, applicants = values
        jh = current_jh
        num = current_num
        ratio = current_ratio
    elif len(values) == 2: # 전형, 단과대, 모집인원, 경쟁률 로우스팬
        major, applicants = values
        jh = current_jh
        dept = current_dept
        num = current_num
        ratio = current_ratio

    
    else:
        # 특수 케이스(열 병합 등): 음수를 통해 오른쪽부터 집어옴, 누락되어있으면 이전값 사용, 학교마다 수정 필요
        ratio = values[-1] if len(values) >= 1 else ""
        applicants = values[-2] if len(values) >= 2 else ""
        num = values[-3] if len(values) >= 3 else current_num
        # 만약 제일 왼쪽에 학과/전공 이름이 있으면 잡아주기
        major = values[0] if len(values) >= 3 else current_major
        dept = current_dept

    # 항상 6개 항목으로 append
    data.append([jh, dept, major, num, applicants, ratio])

# pandas DataFrame으로 변환 
# 일단 요소는 5개 다 챙기고 나중에 모집인원과 지원인원을 드랍할 거임
df_old = pd.DataFrame(data, columns=['전형명', '단과대명', '학과명', '모집인원', '지원인원', '경쟁률'])

# 현재 시각을 열 이름으로
# df_new는 df_old에서 경쟁률을 현재시각으로 바꾼 버전
now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
df_new = df_old.rename(columns={"경쟁률": now})

df_new["모집인원"] = df_new["모집인원"].fillna("").astype(str).str.strip()

# 병합용 키: 단과대 + 학과 + 모집인원 (모집인원 포함 안 하면 같은 학과 다른 전형끼리 묶여서 순서가 깨짐)
df_new["_key"] = df_new["전형명"].astype(str).str.strip() + "_" + \
                 df_new["단과대명"].astype(str).str.strip() + "_" + \
                 df_new["학과명"]
# df_new["_key"] = df_new["단과대명"].astype(str).str.strip() + " " + df_new["학과명"].astype(str).str.strip() + " " + df_new["모집인원"].astype(str).str.strip()
df_new["_row_order"] = range(len(df_new)) # 얘는 뭔 순서 유지용이라는데 ㅁㄹ

# filename = "/home/ubuntu/scr_prc/건국대학교_2026_수시_경쟁률_추이.xlsx" # 수정
filename = r"C:\파이썬 아이들\scr_prc\건국대학교_2025_수시_경쟁률_추이.xlsx"


if os.path.exists(filename):
    df_prev = pd.read_excel(filename) #df_prev가 df_new랑 뭔 차인지는 잘 모르겠음

    # 만약 이전 파일에 '모집인원'이 없으면 빈 컬럼 만들기(안정성)
    if "모집인원" not in df_prev.columns:
        df_prev["모집인원"] = ""


    df_prev["_key"] = df_prev["전형명"].astype(str).str.strip() + "_" + \
                  df_prev["단과대명"].astype(str).str.strip() + "_" + \
                  df_prev["학과명"]
    
    df_prev["모집인원"] = df_prev["모집인원"].fillna("").astype(str).str.strip()


    # 병합 (key 기준으로 가로 합치기 -> 시간 컬럼들이 옆으로 붙음)
    df_merged = pd.merge(
        df_prev,
        df_new,
        on=["_key", "전형명", "단과대명", "학과명", "모집인원"],
        suffixes=("_old", "_new"),
        how="outer"   # (수정?)
    )

    # 공통 컬럼 처리 (여기부터 이해 못 함)
    common_cols = set(df_prev.columns) & set(df_new.columns)
    for col in common_cols:
        if col in ["_key", "전형명", "단과대명", "학과명", "모집인원"]:  # key 역할인 건 제외
            continue
        df_merged[col] = df_merged[col + "_old"].combine_first(df_merged[col + "_new"])
        df_merged = df_merged.drop(columns=[col + "_old", col + "_new"], errors="ignore")

    df_final = (
        df_merged
        .sort_values(by="_row_order")
        .drop(columns=["_row_order", "_key", "모집인원", "지원인원"], errors="ignore")
    )

else:
    # 최초 실행일 때는 그냥 저장 (merge 불필요)
    df_new = df_old.rename(columns={"경쟁률": now})
    df_new["_key"] = (
        df_new["전형명"].astype(str) + " " +
        df_new["단과대명"].astype(str) + " " +
        df_new["학과명"].astype(str)
    )
    df_new["_row_order"] = range(len(df_new))
    df_final = (
        df_new
        .drop(columns=["_row_order", "_key", "모집인원", "지원인원"], errors="ignore")
    )



# 열 순서: 단과대명, 학과명 먼저, 그 다음 시간열들(나머지)
cols = [c for c in df_final.columns if c not in ("전형명", "단과대명", "학과명")]
final_cols = ["전형명", "단과대명", "학과명"] + cols
df_final = df_final[final_cols]

# 저장
df_final.to_excel(filename, index=False, engine="openpyxl")
print(f"{now} 시각의 데이터가 추가되었습니다.")