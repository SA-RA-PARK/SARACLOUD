import pandas as pd
from datetime import datetime
import os
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from openpyxl.styles import NamedStyle


# 작업할 폴더 경로
folder_path = r"C:\\IACFPYTHON\\myproject\\work"
file_list = [f for f in os.listdir(folder_path) if f.endswith('.xlsx') and not f.startswith('~$')]

if file_list:
    for file in file_list:
        file_path = os.path.join(folder_path, file)
        df = pd.read_excel(file_path, engine='openpyxl')

        # 필요없는 행 삭제 : D열, M열, U열 필터링
        df_filtered = df[
            (~df.iloc[:,3].astype(str).str.contains('세종대학교|건수', na=False)) &  # D열 필터링
            (~df.iloc[:,12].astype(str).str.contains(r'간접비|_대응|\(대응|\[대응', regex=True, na=False)) &  # M열 필터링 
            (~df.iloc[:,20].astype(str).str.contains('연구산학협력처|산학협력단', na=False))  # U열 필터링
        ]
        
        # 다운로드 날짜 입력하기
        # ①3번째 열(인덱스 2)의 헤드명을 '담당자'에서 '날짜'로 변경
        df_filtered.rename(columns={df_filtered.columns[2]: '날짜'}, inplace=True)

        # ②'날짜' 열의 값(value)을 오늘 날짜로 설정
        df_filtered['날짜'] = pd.to_datetime(datetime.today().strftime('%Y-%m-%d')).date()

        # 오름차순 정렬 : 과제명 > 총연구기간시작일 > 총괄과제총연구비
        # M열(인덱스 12), BB열(인덱스 53), BA열(인덱스 52) 순으로 오름차순 정렬
        df_filtered_sorted = df_filtered.sort_values(by=[df_filtered.columns[12], df_filtered.columns[53], df_filtered.columns[52]], ascending=True)


        # 사용할 열만 선택 & 추출
        selected_columns_indices = [2, 3, 12, 15, 16, 20, 21, 52, 53, 54]
        selected_columns = df_filtered.iloc[:, selected_columns_indices]

        # 오늘 날짜로 새로운 엑셀 파일 저장
        today_date = datetime.now().strftime('%Y-%m-%d')
        new_file_name = f"{today_date}_{file}"
        new_file_path = os.path.join(folder_path, new_file_name)

        # 엑셀 파일 저장
        with pd.ExcelWriter(new_file_path, engine='openpyxl') as writer:
            selected_columns.to_excel(writer, sheet_name=today_date, index=False)

            # 열 너비 수동 설정
            workbook = writer.book
            worksheet = workbook[today_date]

            # 열 너비 설정 
            column_widths = {
                'A': 11,   # A열
                'B': 33,   # B열
                'C': 67,   # C열
                'D': 10,   # D열
                'E': 8,    # E열
                'F': 17,   # F열
                'G': 17,   # G열
                'H': 17,   # H열
                'I': 11,   # I열
                'J': 11,   # J열
            }

            # 지정된 열 너비 적용
            for col_letter, col_width in column_widths.items():
                worksheet.column_dimensions[col_letter].width = col_width

            # 첫 번째 행의 높이를 32로 설정
            worksheet.row_dimensions[1].height = 32


            # 모든 셀을 '가운데 맞춤' 및 '가운데 정렬'로 설정 (B, C 열 제외)
            for row in worksheet.iter_rows():
                for cell in row:
                    # B열과 C열을 제외한 모든 열에 대해 가운데 정렬
                    if cell.column_letter not in ['B', 'C']:
                        cell.alignment = Alignment(horizontal='center', vertical='center')

            # H열(인덱스 8)의 모든 셀에 '쉼표 스타일' 적용
            for row in worksheet.iter_rows(min_col=8, max_col=8):
                for cell in row:
                    cell.number_format = '#,##0'  # 쉼표 스타일 적용


        print(f"작업 완료: '{new_file_name}'에 데이터가 저장되었습니다.")

else:
    print("작업할 엑셀 파일이 없습니다.")
