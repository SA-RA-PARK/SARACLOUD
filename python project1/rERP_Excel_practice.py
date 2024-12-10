import os
import pandas as pd
from datetime import datetime

# 작업할 폴더 경로
folder_path = r"C:\\IACFPYTHON\\myproject\\work"
file_list = [f for f in os.listdir(folder_path) if f.endswith('.xlsx') and not f.startswith('~$')]

if file_list:
    for file in file_list:
        file_path = os.path.join(folder_path, file)
        df = pd.read_excel(file_path, engine='openpyxl')

        # D열, M열, U열 필터링
        df_filtered = df[
            (~df.iloc[:,3].astype(str).str.contains('세종대학교', na=False)) &  # D열 필터링
            (~df.iloc[:,12].astype(str).str.contains(r'간접비|_대응|\(대응|\[대응', regex=True, na=False)) &  # M열 필터링 
            (~df.iloc[:,20].astype(str).str.contains('연구산학협력처|산학협력단', na=False))  # U열 필터링
        ]

        # 3번째 열(인덱스 2)의 헤드명을 '담당자'에서 '날짜'로 변경
        df_filtered.rename(columns={df_filtered.columns[2]: '날짜'}, inplace=True)

        # '날짜' 열의 값을 오늘 날짜로 설정
        df_filtered['날짜'] = pd.to_datetime(datetime.today().strftime('%Y-%m-%d')).date()

        # 사용할 열만 선택 & 추출
        selected_columns_indices = [2, 3, 12, 15, 16, 20, 21, 52, 53, 54]
        selected_columns = df_filtered.iloc[:, selected_columns_indices]

        # 오늘 날짜로 새로운 엑셀 파일 저장
        today_date = datetime.now().strftime('%Y-%m-%d')
        new_file_name = f"{today_date}_{file}"
        new_file_path = os.path.join(folder_path, new_file_name)

        with pd.ExcelWriter(new_file_path, engine='openpyxl') as writer:
            selected_columns.to_excel(writer, sheet_name=today_date, index=False)

        print(f"작업 완료: '{new_file_name}'에 데이터가 저장되었습니다.")

else:
    print("작업할 엑셀 파일이 없습니다.")
