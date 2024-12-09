from time import sleep
from selenium import webdriver
from datetime import datetime
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import os
import util as ut
import pandas as pd
import openpyxl


#1. 작업준비 - 필요한 내용은 빼서 새로운 파일로 저장
# 1) 작업할 폴더 경로
folder_path = r"C:\\IACFPYTHON\\myproject\\work"
file_list = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]

# 2) RawData 파일에서 필요한 데이터만 꺼내 새로운 파일로 저장
if file_list:
    for file in file_list:
        file_path = os.path.join(folder_path, file)
        df = pd.read_excel(file_path, engine='openpyxl')

        # ① 사용할 data열만 선택
        selected_columns_indices = [2, 3, 12, 15, 16, 20, 21, 52, 53, 54]
        selected_columns = df.iloc[:, selected_columns_indices]

        # ③ 새로운 엑셀 파일로 저장
        today_date = datetime.now().strftime('%Y-%m-%d')  # 오늘 날짜
        new_file_path = os.path.join(folder_path, f"{today_date}_{file}")  # 새로운 엑셀 파일 이름

        # ③ 새로운 파일로 데이터 저장
        with pd.ExcelWriter(new_file_path, engine='openpyxl') as writer:
            selected_columns.to_excel(writer, sheet_name=today_date, index=False)

        print(f"작업 완료: '{new_file_path}'에 데이터가 저장되었습니다.")

else:
    print("엑셀 파일이 없습니다.")
    exit()

#2. 데이터 전처리 - 필요없는 데이터 삭제하기(행 삭제)
