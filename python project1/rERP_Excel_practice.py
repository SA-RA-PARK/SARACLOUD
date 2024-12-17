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

if file_list:
    # 수정 시간 기준으로 내림차순 정렬(가장 최근 파일이 첫 번째)
    file_list = sorted(file_list, key=lambda x: os.path.getmtime(os.path.join(folder_path, x)), reverse=True)

    # 가장 최신 파일 선택
    latest_file = file_list[0]
    file_path = os.path.join(folder_path, latest_file)

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
    new_file_name = f"{today_date}_{latest_file}"
    new_file_path = os.path.join(folder_path, new_file_name)

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

        # 이하 서식 적용 코드는 기존과 동일
        workbook = writer.book
        worksheet_today = workbook[today_date]
        worksheet_no_duplicates = workbook[f"중복제거({today_date})"]

        worksheet_today.freeze_panes = "A2"
        worksheet_no_duplicates.freeze_panes = "A2"

        column_widths = {
            'A': 11,
            'B': 33,
            'C': 95,
            'D': 13,
            'E': 10,
            'F': 17,
            'G': 17,
            'H': 17,
            'I': 14,
            'J': 14,
        }

        for col_letter, col_width in column_widths.items():
            worksheet_today.column_dimensions[col_letter].width = col_width
            worksheet_no_duplicates.column_dimensions[col_letter].width = col_width

        for row in worksheet_today.iter_rows():
            for cell in row:
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.font = Font(size=10)

        for row in worksheet_no_duplicates.iter_rows():
            for cell in row:
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.font = Font(size=10)

        worksheet_today.row_dimensions[1].height = 32
        worksheet_no_duplicates.row_dimensions[1].height = 32

        for cell in worksheet_today[1]:
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.font = Font(name="맑은 고딕", size=10.5, bold=True)
            cell.fill = PatternFill(start_color='BFBFBF', end_color='BFBFBF', fill_type='solid')

        for cell in worksheet_no_duplicates[1]:
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.font = Font(name="맑은 고딕", size=10.5, bold=True)
            cell.fill = PatternFill(start_color='BFBFBF', end_color='BFBFBF', fill_type='solid')

        for row in worksheet_today.iter_rows(min_row=2, max_row=worksheet_today.max_row, min_col=2, max_col=2):
            for cell in row:
                cell.alignment = Alignment(horizontal='left', vertical='center')

        for row in worksheet_no_duplicates.iter_rows(min_row=2, max_row=worksheet_no_duplicates.max_row, min_col=2, max_col=2):
            for cell in row:
                cell.alignment = Alignment(horizontal='left', vertical='center')

        for row in worksheet_today.iter_rows(min_row=2, max_row=worksheet_today.max_row, min_col=3, max_col=3):
            for cell in row:
                cell.alignment = Alignment(horizontal='left', vertical='center')

        for row in worksheet_no_duplicates.iter_rows(min_row=2, max_row=worksheet_no_duplicates.max_row, min_col=3, max_col=3):
            for cell in row:
                cell.alignment = Alignment(horizontal='left', vertical='center')

        for row in worksheet_today.iter_rows(min_row=2, max_row=worksheet_today.max_row, min_col=8, max_col=8):
            for cell in row:
                cell.alignment = Alignment(horizontal='right', vertical='center')

        for row in worksheet_no_duplicates.iter_rows(min_row=2, max_row=worksheet_no_duplicates.max_row, min_col=8, max_col=8):
            for cell in row:
                cell.alignment = Alignment(horizontal='right', vertical='center')

        for row in worksheet_today.iter_rows(min_col=8, max_col=8):
            for cell in row:
                cell.number_format = '#,##0'

        for row in worksheet_no_duplicates.iter_rows(min_col=8, max_col=8):
            for cell in row:
                cell.number_format = '#,##0'

        thin_side = Side(border_style="thin", color="000000")  
        thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

        for row in worksheet_today.iter_rows():
            for cell in row:
                cell.border = thin_border

        for row in worksheet_no_duplicates.iter_rows():
            for cell in row:
                cell.border = thin_border

    print(f"작업 완료: '{new_file_name}'에 데이터가 저장되었습니다.")

else:
    print("작업할 엑셀 파일이 없습니다.")
