import os
import re
from datetime import datetime
import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl import load_workbook

def clean_value(value):
    if isinstance(value, str):  # 값이 문자열인 경우에만 처리
        if value.startswith("["):  # "["로 시작하는 경우
            value = re.sub(r'\[.*?\]', '', value).strip()
        if value.startswith("("):  # "("로 시작하는 경우
            value = re.sub(r'\(.*?\)', '', value).strip()
        if value.endswith("차년도)"):
            value = re.sub(r'\(.*?차년도\)', '', value).strip()
        if value.endswith("년차)"):
            value = re.sub(r'\(.*?년차\)', '', value).strip()
        if value.endswith("차)"):
            value = re.sub(r'\(.*?차\)', '', value).strip()
    return value

# 작업할 폴더 경로
folder_path = r"C:\\IACFPYTHON\\myproject\\RnD_list"
file_list = [f for f in os.listdir(folder_path) if f.endswith('.xlsx') and not f.startswith('~$')]

# "result" 폴더 경로 지정 및 폴더 없으면 생성
result_folder = os.path.join(folder_path, "result")
if not os.path.exists(result_folder):
    os.makedirs(result_folder)

if file_list:
    for file in file_list:
        file_path = os.path.join(folder_path, file)
        df = pd.read_excel(file_path, engine='openpyxl')

        # 1) 필요없는 행 삭제 : D열, M열, U열 필터링
        df_filtered = df[
            (~df.iloc[:, 3].astype(str).str.contains('세종대학교|건수', na=False)) &  # D열 필터링
            (~df.iloc[:, 12].astype(str).str.contains(r'간접비|_대응|\(대응|\[대응', regex=True, na=False)) &  # M열 필터링
            (~df.iloc[:, 20].astype(str).str.contains('연구산학협력처|산학협력단', na=False))  # U열 필터링
        ]

        # 2) 날짜 열 처리 (3번째 열 -> '날짜')
        df_filtered.rename(columns={df_filtered.columns[2]: '날짜'}, inplace=True)
        df_filtered['날짜'] = pd.to_datetime(datetime.today().strftime('%Y-%m-%d')).date()

        # 3) 정렬: M열(12), BB열(53), BA열(52)
        df_filtered_sorted = df_filtered.sort_values(
            by=[df_filtered.columns[12], df_filtered.columns[53], df_filtered.columns[52]],
            ascending=True
        )

        # 사용할 열만 선택 & 추출
        selected_columns_indices = [2, 3, 12, 15, 16, 20, 21, 52, 53, 54]
        selected_columns = df_filtered_sorted.iloc[:, selected_columns_indices]

        # "오늘날짜" 시트에는 D,M,U열 필터링, 날짜열 처리, 정렬 단계까지만 반영
        today_date = datetime.now().strftime('%Y-%m-%d')
        new_file_name = f"{today_date}_{file}"
        new_file_path = os.path.join(result_folder, new_file_name)  # result 폴더에 저장

        with pd.ExcelWriter(new_file_path, engine='openpyxl') as writer:
            # "오늘날짜" 시트 저장 (여기까지만: 필터링, 날짜열 처리, 정렬)
            selected_columns.to_excel(writer, sheet_name=today_date, index=False)

            # 여기서부터 "중복제거(오늘날짜)" 시트를 위해 추가 작업
            # 4) M열(13번째 열) 데이터 정제
            df_filtered_sorted.iloc[:, 12] = df_filtered_sorted.iloc[:, 12].apply(clean_value)

            # 7) 문자열 중간에 "(X차년도 ... )" 패턴 제거
            df_filtered_sorted.iloc[:, 12] = df_filtered_sorted.iloc[:, 12].str.replace(r'\(\d+차년도[^\)]*\)', '', regex=True)

            # 8) 문자열 끝에 "_X차년도" 패턴 제거
            df_filtered_sorted.iloc[:, 12] = df_filtered_sorted.iloc[:, 12].str.replace(r'_\d+차년도$', '', regex=True)

            # (X단계-Y차년도)나 (X단계_Y차년도) -> (X단계)
            df_filtered_sorted.iloc[:, 12] = df_filtered_sorted.iloc[:, 12].str.replace(r'\((\d+단계)[-_]\d+차년도\)', r'(\1)', regex=True)

            # (X단계(차년도)) -> X단계
            df_filtered_sorted.iloc[:, 12] = df_filtered_sorted.iloc[:, 12].str.replace(r'(\d+단계)\(\d+차년도\)', r'\1', regex=True)

            # 10) 필요시 양쪽 공백 제거
            df_filtered_sorted.iloc[:, 12] = df_filtered_sorted.iloc[:, 12].str.strip()

            # 11) 재정렬: BA열(52)-내림차순, M열(12)-오름차순, BB열(53)-오름차순
            df_filtered_sorted = df_filtered_sorted.sort_values(
               by=[df_filtered_sorted.columns[52], df_filtered_sorted.columns[12], df_filtered_sorted.columns[53]],
               ascending=[False, True, True]
            )
            
            # M열 정제 이후의 selected_columns (중복제거 시트용)
            selected_columns_cleaned = df_filtered_sorted.iloc[:, selected_columns_indices]

            # 11) 중복 제거: 13번째 열(index 12)만을 기준으로 중복 제거
            df_no_duplicates = selected_columns_cleaned.drop_duplicates(subset=[selected_columns_cleaned.columns[2]])

            # "중복제거(오늘날짜)" 시트 저장
            df_no_duplicates.to_excel(writer, sheet_name=f"중복제거({today_date})", index=False)

            # 서식 적용
            workbook = writer.book
            worksheet_today = workbook[today_date]
            worksheet_no_duplicates = workbook[f"중복제거({today_date})"]

            # 첫 번째 행을 고정하기 위해 A2 셀을 freeze_panes로 설정
            worksheet_today.freeze_panes = "A2"
            worksheet_no_duplicates.freeze_panes = "A2"

            # 열 너비 설정 
            column_widths = {
                'A': 11,   # A열
                'B': 33,   # B열
                'C': 95,   # C열
                'D': 13,   # D열
                'E': 10,    # E열
                'F': 17,   # F열
                'G': 17,   # G열
                'H': 17,   # H열
                'I': 14,   # I열
                'J': 14,   # J열
            }

            for col_letter, col_width in column_widths.items():
                worksheet_today.column_dimensions[col_letter].width = col_width
                worksheet_no_duplicates.column_dimensions[col_letter].width = col_width


            # 글씨 크기 10으로 설정 및 모든 셀 가운데 정렬
            for row in worksheet_today.iter_rows():
                for cell in row:
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.font = Font(size=10)

            for row in worksheet_no_duplicates.iter_rows():
                for cell in row:
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.font = Font(size=10)
            
            # 첫 번째 행 높이 설정
            worksheet_today.row_dimensions[1].height = 32
            worksheet_no_duplicates.row_dimensions[1].height = 32

            # 첫 번째 행 가운데 정렬 + 굵게 + 셀색상 지정
            for cell in worksheet_today[1]:
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.font = Font(name="맑은 고딕", size=10.5, bold=True)
                cell.fill = PatternFill(start_color='BFBFBF', end_color='BFBFBF', fill_type='solid')

            for cell in worksheet_no_duplicates[1]:
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.font = Font(name="맑은 고딕", size=10.5, bold=True)
                cell.fill = PatternFill(start_color='BFBFBF', end_color='BFBFBF', fill_type='solid')

            # B열(2번 컬럼) 2행부터 왼쪽 정렬
            for row in worksheet_today.iter_rows(min_row=2, max_row=worksheet_today.max_row, min_col=2, max_col=2):
                for cell in row:
                    cell.alignment = Alignment(horizontal='left', vertical='center')

            for row in worksheet_no_duplicates.iter_rows(min_row=2, max_row=worksheet_no_duplicates.max_row, min_col=2, max_col=2):
                for cell in row:
                    cell.alignment = Alignment(horizontal='left', vertical='center')

            # C열(3번 컬럼) 2행부터 왼쪽 정렬
            for row in worksheet_today.iter_rows(min_row=2, max_row=worksheet_today.max_row, min_col=3, max_col=3):
                for cell in row:
                    cell.alignment = Alignment(horizontal='left', vertical='center')

            for row in worksheet_no_duplicates.iter_rows(min_row=2, max_row=worksheet_no_duplicates.max_row, min_col=3, max_col=3):
                for cell in row:
                    cell.alignment = Alignment(horizontal='left', vertical='center')

            # H열(8번 컬럼) 2행부터 오른쪽 정렬
            for row in worksheet_today.iter_rows(min_row=2, max_row=worksheet_today.max_row, min_col=8, max_col=8):
                for cell in row:
                    cell.alignment = Alignment(horizontal='right', vertical='center')

            for row in worksheet_no_duplicates.iter_rows(min_row=2, max_row=worksheet_no_duplicates.max_row, min_col=8, max_col=8):
                for cell in row:
                    cell.alignment = Alignment(horizontal='right', vertical='center')

            # H열 쉼표 스타일 적용
            for row in worksheet_today.iter_rows(min_col=8, max_col=8):
                for cell in row:
                    cell.number_format = '#,##0'  # 쉼표 스타일 적용

            for row in worksheet_no_duplicates.iter_rows(min_col=8, max_col=8):
                for cell in row:
                    cell.number_format = '#,##0'  # 쉼표 스타일 적용

            # (2) 전체 셀에 테두리 적용
            thin_side = Side(border_style="thin", color="000000")  
            thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

            # 오늘날짜 시트 테두리 적용
            for row in worksheet_today.iter_rows():
                for cell in row:
                    cell.border = thin_border

            # 중복제거(오늘날짜) 시트 테두리 적용
            for row in worksheet_no_duplicates.iter_rows():
                for cell in row:
                    cell.border = thin_border

        print(f"작업 완료: '{new_file_name}'에 데이터가 저장되었습니다.")

else:
    print("작업할 엑셀 파일이 없습니다.")
