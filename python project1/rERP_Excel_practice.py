import os
import re
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font

# M열 데이터 정제 함수
def clean_value(value):
    """
    M열의 값을 정제하는 함수:
    1. 값이 "["로 시작하면 "[~~~]" 패턴 제거.
    2. 값이 "("로 시작하면 "(~~~)" 패턴 제거.
    3. 값이 "차년도)"로 끝나면 "(차년도)" 패턴 제거.
    4. 조건에 맞지 않으면 원래 값을 반환.
    """
    if isinstance(value, str):  # 값이 문자열인 경우에만 처리
        if value.startswith("["):  # "["로 시작하는 경우
            value = re.sub(r'\[.*?\]', '', value).strip()
        if value.startswith("("):  # "("로 시작하는 경우
            value = re.sub(r'\(.*?\)', '', value).strip()
        if value.endswith("차년도)"):  # "차년도)"로 끝나는 경우
            value = re.sub(r'\(.*?차년도\)', '', value).strip()
    return value

# 작업할 폴더 경로
folder_path = r"C:\\IACFPYTHON\\myproject\\work"
file_list = [f for f in os.listdir(folder_path) if f.endswith('.xlsx') and not f.startswith('~$')]

if file_list:
    for file in file_list:
        file_path = os.path.join(folder_path, file)
        df = pd.read_excel(file_path, engine='openpyxl')

        # 필요없는 행 삭제 : D열, M열, U열 필터링
        df_filtered = df[
            (~df.iloc[:, 3].astype(str).str.contains('세종대학교|건수', na=False)) &  # D열 필터링
            (~df.iloc[:, 12].astype(str).str.contains(r'간접비|_대응|\(대응|\[대응', regex=True, na=False)) &  # M열 필터링 
            (~df.iloc[:, 20].astype(str).str.contains('연구산학협력처|산학협력단', na=False))  # U열 필터링
        ]

        # M열(13번째 열) 데이터 정제
        df_filtered.iloc[:, 12] = df_filtered.iloc[:, 12].apply(clean_value)

        # 날짜 열 처리
        df_filtered.rename(columns={df_filtered.columns[2]: '날짜'}, inplace=True)
        df_filtered['날짜'] = pd.to_datetime(datetime.today().strftime('%Y-%m-%d')).date()

        # 정렬: M열(12), BB열(53), BA열(52)
        df_filtered_sorted = df_filtered.sort_values(
            by=[df_filtered.columns[12], df_filtered.columns[53], df_filtered.columns[52]],
            ascending=True
        )

        # 사용할 열만 선택 & 추출
        selected_columns_indices = [2, 3, 12, 15, 16, 20, 21, 52, 53, 54]
        selected_columns = df_filtered_sorted.iloc[:, selected_columns_indices]


        # 중복 제거: 중복 항목 제거 후 새로운 DataFrame 생성
        df_no_duplicates = selected_columns.drop_duplicates()

        # 엑셀파일 저장
        today_date = datetime.now().strftime('%Y-%m-%d')
        new_file_name = f"{today_date}_{file}"
        new_file_path = os.path.join(folder_path, new_file_name)


        with pd.ExcelWriter(new_file_path, engine='openpyxl') as writer:
            # "오늘날짜" 시트에 데이터를 추가
            selected_columns.to_excel(writer, sheet_name=today_date, index=False)
            # "중복제거(오늘날짜)" 시트에 중복 제거된 데이터를 추가
            df_no_duplicates.to_excel(writer, sheet_name=f"중복제거({today_date})", index=False)

            # 워크북 및 워크시트 설정
            workbook = writer.book
            worksheet_today = workbook[today_date]
            worksheet_no_duplicates = workbook[f"중복제거({today_date})"]


            # 열 너비 수동 설정
            workbook = writer.book
            worksheet = workbook[today_date]
            worksheet_no_duplicates = workbook[f"중복제거({today_date})"]

            # 열 너비 설정 
            column_widths = {
                'A': 11,   # A열
                'B': 33,   # B열
                'C': 67,   # C열
                'D': 12,   # D열
                'E': 9,    # E열
                'F': 17,   # F열
                'G': 17,   # G열
                'H': 17,   # H열
                'I': 13,   # I열
                'J': 13,   # J열
            }

            # 지정된 열 너비 적용
            for col_letter, col_width in column_widths.items():
                worksheet.column_dimensions[col_letter].width = col_width
                worksheet_no_duplicates.column_dimensions[col_letter].width = col_width

            # 첫 번째 행의 높이를 32로 설정
            worksheet.row_dimensions[1].height = 32
            worksheet_no_duplicates.row_dimensions[1].height = 32

            # 첫 번째 행 가운데 정렬
            for cell in worksheet[1]:
                cell.alignment = Alignment(horizontal='center', vertical='center')

            for cell in worksheet_no_duplicates[1]:
                cell.alignment = Alignment(horizontal='center', vertical='center')

            # 글씨 크기 10으로 설정
            font_style = Font(size=10)  # 글씨 크기를 10으로 설정


            # 모든 셀을 '가운데 맞춤'으로 설정
            for row in worksheet.iter_rows():
                for cell in row:
                    cell.alignment = Alignment(horizontal='center', vertical='center')

                    # 모든 셀에 글씨 크기 10 적용
                    cell.font = Font(size=10)


            for row in worksheet_no_duplicates.iter_rows():
                for cell in row:
                    cell.alignment = Alignment(horizontal='center', vertical='center')        

                    # 모든 셀에 글씨 크기 10 적용
                    cell.font = Font(size=10)


            # B열의 두 번째 행부터 왼쪽 정렬 설정
            for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=2, max_col=2):
                for cell in row:
                    cell.alignment = Alignment(horizontal='left', vertical='center')

            for row in worksheet_no_duplicates.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=2, max_col=2):
                for cell in row:
                    cell.alignment = Alignment(horizontal='left', vertical='center')        

            # C열의 두 번째 행부터 왼쪽 정렬 설정
            for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=3, max_col=3):
                for cell in row:
                    cell.alignment = Alignment(horizontal='left', vertical='center')

            for row in worksheet_no_duplicates.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=3, max_col=3):
                for cell in row:
                    cell.alignment = Alignment(horizontal='left', vertical='center')

            # H열의 두 번째 행부터 오른쪽 정렬 설정
            for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=8, max_col=8):
                for cell in row:
                    cell.alignment = Alignment(horizontal='right', vertical='center')

            for row in worksheet_no_duplicates.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=8, max_col=8):
                for cell in row:
                    cell.alignment = Alignment(horizontal='right', vertical='center')


            # H열(인덱스 8)의 모든 셀에 '쉼표 스타일' 적용
            for row in worksheet.iter_rows(min_col=8, max_col=8):
                for cell in row:
                    cell.number_format = '#,##0'  # 쉼표 스타일 적용

                
            for row in worksheet_no_duplicates.iter_rows(min_col=8, max_col=8):
                for cell in row:
                    cell.number_format = '#,##0'  # 쉼표 스타일 적용 

        print(f"작업 완료: '{new_file_name}'에 데이터가 저장되었습니다.")

else:
    print("작업할 엑셀 파일이 없습니다.")

