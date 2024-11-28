import re, os, inspect, logging, datetime, sys, configparser
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from openpyxl import load_workbook
# from PyQt5.QtWidgets import QTextEdit, QProgressBar, QFileDialog

log_bool= True

def exedir(chk) :                           #chk = 변수이름, check의 줄임말
    if os.path.splitext(sys.executable)[1] == ".exe" and "python" not in sys.executable.lower():
        execute_dir = os.path.dirname(sys.executable)
    elif chk == "dir" : # 현재 디렉토리 위치 얻기
        execute_dir = os.getcwd()
    elif chk == "py"  : # 현재 파이썬 파일 디렉토리 위치 얻기
        execute_dir = os.path.dirname(os.path.abspath(__file__))
    else:
        execute_dir = "디렉토리 오류"
    
    return execute_dir

today = datetime.datetime.now().strftime("%Y%m%d")  # 'YYYYMMDD' 형식
log_filename = os.path.expandvars(f"{exedir("py")}\\log_{today}.txt")

# 로그 설정 (파일 저장)
logging.basicConfig(
    filename=log_filename,              # 로그 파일명
    level=logging.INFO,                # 로그 레벨
    format="%(asctime)s - %(message)s", # 로그 포맷
    filemode="a"                        #a = append, w = 덮어쓰기
)

def debug_print(*args, **kwargs):
    display = kwargs.get('display', False)
    # 현재 파일명과 줄 번호 가져오기
    frame = inspect.currentframe().f_back
    file_name = frame.f_code.co_filename
    line_number = frame.f_lineno

    # 파일명과 줄 번호를 포함하여 로그 메시지 작성
    message = f"[{file_name}:{line_number}] - " + " ".join(map(str, args))

    # 로그 메시지 파일에 기록
    logging.info(message)
    # display=True일 때만 콘솔 출력
    if display:
        print(message)

def acctonum(account_number) :
    # 정규식 패턴
    pattern = r'(\d{6}-?\d{2}-?\d{6})|(\d{12,})'
    # 계좌번호 추출
    result = re.search(pattern, account_number)
    account_number = result.group().replace('-', '') if result else None
    return str(account_number)

#def select_folder(widget):
#    folder_path = QFileDialog.getExistingDirectory(widget, "폴더 선택")
#    return folder_path

def create_folder(folder_path):
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
    return folder_path

def save_excel_with_seq(file_path):
    base, ext = os.path.splitext(file_path)
    sequence = 1
    new_file_path = file_path

    # 파일명이 이미 존재하면 시퀀스를 붙여서 새 파일명을 만듬
    while os.path.exists(new_file_path):
        new_file_path = f"{base}_{sequence}{ext}"
        sequence += 1
    
    return new_file_path

def select_folder_console(display=False):
    """
    사용자가 폴더를 선택할 수 있게 하고, 선택된 폴더의 파일 목록을 반환합니다.
    """
    # Tkinter 창을 숨기기
    root = tk.Tk()
    root.withdraw()
    
    # 폴더 선택 대화상자 열기
    folder_path = filedialog.askdirectory(title="폴더를 선택하세요")
    
    # 선택된 폴더가 없으면 종료
    if not folder_path:
        debug_print("폴더가 선택되지 않았습니다. Download 폴더로 기본지정됩니다.",display=log_bool)
        folder_path = os.path.expandvars(r"%UserProfile%\\Downloads")
    else :
        # 폴더의 파일 목록 가져오기
        file_list = os.listdir(folder_path)
        debug_print(f"선택된 폴더: {folder_path}",display=log_bool)
    
    return folder_path

def get_login_info(display=False):
    # config_file_path = f"{crdir("py", False)}/config.txt"
    config_file_path = f"{exedir("py")}\\config.txt"
    debug_print(f"msg:[config_file_path ==> {config_file_path}",display=log_bool)
    if os.path.exists(config_file_path):
        # 파일 읽기
        with open(config_file_path, 'r', encoding="utf-8") as file:
            # 파일의 각 줄을 읽어서 리스트로 저장
            lines = file.readlines()

        # 변수 초기화
        id_value = password_value = recipient_email = None

        for i, line in enumerate(lines):
            if i == 0:
                id_value = line.strip()         # 첫 번째 줄: ID
            elif i == 1:
                password_value = line.strip()   # 두 번째 줄: Password
        
        logininfo={"ID" : id_value,"PW" : password_value}
    else :
        logininfo={"ID" : None,"PW" : None}
    return logininfo

def load_config(file_path=f"{exedir("py")}\\config.txt"):
    config = configparser.ConfigParser()
    config.optionxform = str  # 키 이름의 대소문자 구분 유지
    # config.txt 파일이 없으면 생성
    if not os.path.exists(file_path):
        debug_print(f"{file_path} does not exist. Creating the file with default values.",log_bool)
        # 기본 섹션과 값 추가
        config["login"] = {
            "ID": "your_username",
            "PW": "your_password",
            "EMAIL": "your_email"
        }
        # 파일 작성
        with open(file_path, "w", encoding="utf-8") as configfile:
            config.write(configfile)

    # UTF-8 인코딩으로 파일 읽기
    config.read(file_path, encoding="utf-8")
    # 모든 섹션과 항목을 동적으로 읽어서 딕셔너리로 변환
    config_data = {}
    for section in config.sections():
        config_data[section] = {}
        for key in config[section]:
            # Boolean 값인 경우 처리
            if config[section][key].lower() in ['true', 'false']:
                config_data[section][key] = config.getboolean(section, key)
            else:
                config_data[section][key] = config.get(section, key)
    return config_data
