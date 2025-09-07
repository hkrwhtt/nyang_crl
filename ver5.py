from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
import time
import pandas as pd
import datetime
import os

# 크롬 드라이버 실행 (headless 모드)
options = Options()
options.add_argument("--headless")
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

# 웹페이지 열기
driver.get("https://addon.jinhakapply.com/RatioV1/RatioH/Ratio11201011.html")
time.sleep(10)

# 데이터 수집
rows = driver.find_elements(By.CSS_SELECTOR, 'table tbody tr')
data = []
current_dept, current_major, current_num = "", "", ""

for row in rows:
    tds = row.find_elements(By.TAG_NAME, "td")
    if not tds:
        continue
    values = [td.text.strip() for td in tds]

    if len(values) == 5:
        dept, major, num, applicants, ratio = values
        current_dept, current_major, current_num = dept, major, num
    elif len(values) == 4:
        major, num, applicants, ratio = values
        dept = current_dept
    else:  # 특수 케이스
        ratio = values[-1] if len(values) >= 1 else ""
        applicants = values[-2] if len(values) >= 2 else ""
        num = values[-3] if len(values) >= 3 else current_num
        major = values[0] if len(values) >= 3 else current_major
        dept = current_dept

    data.append([dept, major, num, applicants, ratio])

# pandas DataFrame 변환
df_old = pd.DataFrame(data, columns=['단과대명', '학과명', '모집인원', '지원인원', '경쟁률'])

# 모집인원/지원인원 정규화 (숫자만 추출)
df_old["모집인원"] = df_old["모집인원"].str.extract(r"(\d+)").fillna("")
df_old["지원인원"] = df_old["지원인원"].str.extract(r"(\d+)").fillna("")

# 현재 시각을 열 이름으로
now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
df_new = df_old.rename(columns={"경쟁률": now})

# 병합용 키 생성
df_new["_key"] = (
    df_new["단과대명"].astype(str).str.strip() + " " +
    df_new["학과명"].astype(str).str.strip() + " " +
    df_new["모집인원"].astype(str).str.strip()
)

filename = r"C:\파이썬 아이들\scr_prc\이화여자대학교_2025_수시_경쟁률_추이.xlsx"

if os.path.exists(filename):
    df_prev = pd.read_excel(filename)

    if "모집인원" not in df_prev.columns:
        df_prev["모집인원"] = ""

    df_prev["_key"] = (
        df_prev["단과대명"].astype(str).str.strip() + " " +
        df_prev["학과명"].astype(str).str.strip() + " " +
        df_prev["모집인원"].astype(str).str.strip()
    )

    if df_prev.duplicated(subset=["_key"]).any():
        df_prev = (
            df_prev.groupby("_key", as_index=False)
            .agg(lambda s: s.dropna().iloc[0] if s.dropna().size > 0 else "")
        )

    # 병합
    df_merged = pd.merge(
        df_prev,
        df_new,
        on=["_key", "단과대명", "학과명", "모집인원"],
        suffixes=("_old", "_new"),
        how="outer"   # inner → outer로 수정 (데이터 손실 방지)
    )

    # _row_order는 merge 후에 새로 만들기
    df_merged["_row_order"] = df_merged.groupby("단과대명").cumcount()

    # 중복 컬럼 처리
    common_cols = set(df_prev.columns) & set(df_new.columns)
    for col in common_cols:
        if col in ["_key", "단과대명", "학과명", "모집인원"]:
            continue
        if col + "_old" in df_merged and col + "_new" in df_merged:
            df_merged[col] = df_merged[col + "_old"].combine_first(df_merged[col + "_new"])
            df_merged = df_merged.drop(columns=[col + "_old", col + "_new"], errors="ignore")

    df_final = (
        df_merged
        .sort_values(by="_row_order")
        .drop(columns=["_row_order", "_key", "모집인원", "지원인원"], errors="ignore")
    )

else:
    # 최초 실행
    df_final = (
        df_new
        .drop(columns=["_row_order", "_key", "모집인원", "지원인원"], errors="ignore")
    )

# 열 순서 정리
cols = [c for c in df_final.columns if c not in ("단과대명", "학과명")]
df_final = df_final[["단과대명", "학과명"] + cols]

# 저장
df_final.to_excel(filename, index=False, engine="openpyxl")
print(f"{now} 시각의 데이터가 추가되었습니다.")
