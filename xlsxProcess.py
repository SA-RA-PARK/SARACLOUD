import pandas as pd
import util

log_bool= True

# 시트 이름 설정
SHEET_NAMES = {
    "bank": "은행자료",
    "saer": "회계자료",
    "combined": "시트합치기",
    "pivot_out": "피봇출금",
    "pivot_in": "피봇입금",
    "error_check": "오류확인대상"
}

def toExcelErp(directory: SystemError, filename):
    
    # 계좌 번호 추출 및 파일 경로 설정
    filename2 = filename.split('_')[1][-6:]
    account_num = filename.split('_')[1]
    file_paths = {
        "bank": f"{directory}/workF/{filename}.xls",
        "saer": f"{directory}/workF/거래처원장 {filename2}.xls"
    }
    output_file_path = f"{directory}/resultF/{account_num}.xlsx"
    output_file_path = output_file_path.replace("workF","resultF")
    output_file_path = util.save_excel_with_seq(output_file_path)

    # 데이터프레임 읽기 및 전처리
    df_bank = pd.read_excel(file_paths["bank"], header=None).iloc[6:]
    df_saer = pd.read_excel(file_paths["saer"], header=None).iloc[7:]
    
    # 초기 저장 (은행 자료, 회계 자료)
    with pd.ExcelWriter(output_file_path) as writer:
        df_bank.to_excel(writer, sheet_name=SHEET_NAMES["bank"], index=False, header=False)
        df_saer.to_excel(writer, sheet_name=SHEET_NAMES["saer"], index=False, header=False)
    
    # 5. 특정 시트 읽기    
    df_bank = pd.read_excel(output_file_path, sheet_name=SHEET_NAMES["bank"])
    df_saer = pd.read_excel(output_file_path, sheet_name=SHEET_NAMES["saer"])

    # 데이터 가공 및 정리
    df_bank = preprocess_bank_data(df_bank)
    df_saer = preprocess_saer_data(df_saer)
    
    # 데이터 병합
    df_combined = combine_df_data(df_bank, df_saer)

    # 피봇 테이블 생성
    df_pivot_out, df_pivot_in = create_pivot_tables(df_combined, SHEET_NAMES["bank"], SHEET_NAMES["saer"])
 
    # 중간저장
    with pd.ExcelWriter(output_file_path) as writer:
        df_bank.to_excel(writer, sheet_name=SHEET_NAMES["bank"], index=False)
        df_saer.to_excel(writer, sheet_name=SHEET_NAMES["saer"], index=False)
        df_combined.to_excel(writer, sheet_name=SHEET_NAMES["combined"], index=False)
        df_pivot_out.to_excel(writer, sheet_name=SHEET_NAMES["pivot_out"])
        df_pivot_in.to_excel(writer, sheet_name=SHEET_NAMES["pivot_in"])

    # 피봇입금+피봇출금 자료 읽어오기
    df_pivot_out = pd.read_excel(output_file_path, sheet_name=SHEET_NAMES["pivot_out"])
    df_pivot_in = pd.read_excel(output_file_path, sheet_name=SHEET_NAMES["pivot_in"])

    # 피봇입금 + 피봇출금 데이터 가공 및 정리 / #컬럼명 변경하고 오류값만 추출하기
    df_pivot_comb = combine_df_pivot_data(df_pivot_out, df_pivot_in)

    # 최종 저장 (->앞에서 한 저장을 한번더 반복해줘야함. 안그러면 컴바인피봇 표 시트만 생김)
    with pd.ExcelWriter(output_file_path) as writer:
        df_bank.to_excel(writer, sheet_name=SHEET_NAMES["bank"], index=False)
        df_saer.to_excel(writer, sheet_name=SHEET_NAMES["saer"], index=False)
        df_combined.to_excel(writer, sheet_name=SHEET_NAMES["combined"], index=False)
        df_pivot_out.to_excel(writer, sheet_name=SHEET_NAMES["pivot_out"], index=False)
        df_pivot_in.to_excel(writer, sheet_name=SHEET_NAMES["pivot_in"], index=False)
        df_pivot_comb.to_excel(writer, sheet_name=SHEET_NAMES["error_check"], index=False)



##--------------------------------데이터 프로세스----------------------------------------##


#피봇입금+ 피봇출금 데이터 가공 및 정리(0번에 구분 넣기)
def combine_df_pivot_data(df_pivot_out, df_pivot_in):    
    df_pivot_out = df_pivot_out[['거래일시', '은행자료', '회계자료', '출금차액','상태']]
    df_pivot_out.insert(0, '구분', SHEET_NAMES["pivot_out"])  # '구분' 컬럼 추가하기(컬럼 이름은 '시트이름'으로)
    df_pivot_out = df_pivot_out[~df_pivot_in['상태'].str.contains('오류', na=False)]    #'상태'열에서 특정 단어 포함한 행 삭제

    df_pivot_in = df_pivot_in[['거래일시', '은행자료', '회계자료', '입금차액', '상태']]
    df_pivot_in.columns = ['거래일시', '은행자료','회계자료', '출금차액', '상태']  # 컬럼명 변경(입금차액->출금차액)
    df_pivot_in.insert(0, '구분', SHEET_NAMES["pivot_in"])  # '구분' 컬럼 추가하기(컬럼 이름은 '시트이름'으로)
    df_pivot_in = df_pivot_in[~df_pivot_in['상태'].str.contains('오류', na=False)]   #'상태'열에서 특정 단어 포함한 행 삭제

    df_pivot_comb = pd.concat([df_pivot_out, df_pivot_in], ignore_index=True)
    return df_pivot_comb

def preprocess_bank_data(df_bank):
    """은행 데이터를 정리합니다."""
    df_bank['거래일시'] = df_bank['거래일시'].str.replace('.', '', regex=False).str.slice(0, 8)
    df_bank = df_bank.sort_values(by='거래일시')
    df_bank = df_bank.drop(columns=['잔액(원)','내 통장 표시','적요','처리점','구분'])
    return df_bank

def preprocess_saer_data(df_saer):
    """회계 데이터를 정리합니다."""
    df_saer['일자'] = df_saer['일자'].str.replace('-', '', regex=False).str.slice(0, 8)
    df_saer = df_saer.sort_values(by='일자')
    df_saer = df_saer[~df_saer['일자'].str.contains('전기 이월|월계|누계', na=False)]
    df_saer = df_saer[~df_saer['적요'].str.contains('합계', na=False)]
    df_saer = df_saer.drop(columns=['전표번호','계정명','잔액','회계단위명'])
    return df_saer

def combine_df_data(df_bank, df_saer):
    df_bank = df_bank[['거래일시', '출금액(원)','입금액(원)']]
    df_bank.columns = ['거래일시', '출금','입금']  # 컬럼명 변경
    df_bank.insert(0, '구분', SHEET_NAMES["bank"])  # 신규 구분 컬럼을 첫 번째 위치에 추가

    df_saer = df_saer[['일자', '대변','차변']]
    df_saer.columns = ['거래일시', '출금','입금']  # 컬럼명 변경
    df_saer.insert(0, '구분', SHEET_NAMES["saer"])  # 신규 구분 컬럼을 첫 번째 위치에 추가

    df_comb = pd.concat([df_bank, df_saer], ignore_index=True)
    return df_comb

def create_pivot_tables(df_combined, bank_label, saer_label):
    """출금 및 입금 피봇 테이블을 생성합니다."""
    df_pivot_out = df_combined.pivot_table(index="거래일시", columns="구분", values="출금", aggfunc="sum", fill_value=0)
    df_pivot_out["출금차액"] = df_pivot_out[bank_label] - df_pivot_out.get(saer_label, 0)
    df_pivot_out["상태"] = df_pivot_out["출금차액"].apply(lambda x: "정상" if x == 0 else "오류")

    df_pivot_in = df_combined.pivot_table(index="거래일시", columns="구분", values="입금", aggfunc="sum", fill_value=0)
    df_pivot_in["입금차액"] = df_pivot_in[bank_label] - df_pivot_in.get(saer_label, 0)
    df_pivot_in["상태"] = df_pivot_in["입금차액"].apply(lambda x: "정상" if x == 0 else "오류")

    return df_pivot_out, df_pivot_in

def get_worker(directory: SystemError, search_item):
    workers_file_path = f"{directory}\\workers.xlsx"
    df = pd.read_excel(workers_file_path,header=0)
    df.columns = df.columns.str.strip()

    # 조건에 맞는 행 필터링
    matched_row = df[df['계좌번호'].astype(str) == search_item]
    # 계좌번호가 있는지 확인 후, 해당 행의 이름과 메일 가져오기
    if not matched_row.empty:
        name = matched_row.iloc[0]['이름']
        email = matched_row.iloc[0]['메일']
    else:
        name = None
        email = None
        util.debug_print(f"{search_item}의 정보를 찾을 수 없습니다.", display=log_bool)
    
    result = {"Name" : name, "Email" : email}
    return result
