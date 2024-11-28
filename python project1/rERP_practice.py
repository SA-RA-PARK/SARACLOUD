from time import sleep
from selenium import webdriver
from datetime import datetime
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import os
import util as ut

config_info = ut.load_config()       
current_path = ut.exedir('py')
#print(config_info) 
config_info["folder"] = {
    "current" : current_path,
    "work" : os.path.join(current_path, "work\\"),
    "result" : os.path.join(current_path, "result\\")
}
#print("================") 
#print(config_info["login"]["NAME"])                      


# 폴더 새로 만들기
ut.create_folder(config_info["folder"]["work"])      #파일이 없으면 만들고, 있으면 pass다 / work 폴더 만듦
ut.create_folder(config_info["folder"]["result"])    #파일이 없으면 만들고, 있으면 pass다  / result 폴더 만듦


# 브라우저 꺼짐 방지 옵션
chrome_options = Options()
# chrome_options.add_argument("headless")  # 헤드리스 모드 설정
chrome_options.add_experimental_option("detach", True)
chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])

download_path = config_info["folder"]["work"]   # 저장할 위치를 지정하는 거다 / 위치는 위에서 새로만든 work

chrome_options.add_experimental_option("prefs", {
"download.default_directory": download_path,  # 다운로드 폴더 지정
"download.prompt_for_download": False,  # 다운로드 시 사용자에게 묻지 않음
"download.directory_upgrade": True,
"safebrowsing.enabled": True
})

driver = webdriver.Chrome(options=chrome_options)

# 1. 웹페이지 해당 주소(iacf.sejong.ac.kr)로 이동
driver.get(config_info["iacf"]["URL"])       #config.txt에 있는 url을 가지고 오는거다.
driver.maximize_window()

main_window = driver.current_window_handle     #떠있는 화면을 handling하는거다.

# 2. iacf 로그인하기
driver.find_element(By.CSS_SELECTOR,config_info["iacf"]["ID_CSS"]).send_keys(config_info["iacf"]["ID"])  #rerp ID입력
driver.find_element(By.CSS_SELECTOR, config_info["iacf"]["PW_CSS"]).send_keys(config_info["iacf"]["PW"]) #rerp pw입력
driver.find_element(By.CSS_SELECTOR,config_info["iacf"]["LOGIN_BT_CSS"]).click()    #로그인 버튼 누르기


# 3. 로그인 후 뜨는 모든 파업창 닫기
sleep(5)                                  #팝업이 뜨는 속도와 닫는 속도가 다를 수 있으니까 앞에서 sleep을 주는거다. 보통 5초 정도 준다.
all_windows = driver.window_handles      #뜨는 모든 팝업창을 다 가지고 오기 / 새로운 창이 떴다면, 새로운 창으로 전환
for window in all_windows :
    if window != main_window:             #메인창은 닫지마
        driver.switch_to.window(window)   #모든 iframe을 닫으면 안되니까 iframe들을 들리는거다.
        driver.close()
driver.switch_to.window(main_window)      #팝업창 닫기


# 4. [통계]메뉴 버튼 누르기
driver.find_element(By.CSS_SELECTOR, "#menu_id_6").click() 
sleep(3)     #다음 html이 다 뜰때까지 기다려주는거다


# 5. [연구비수주현황whghl] iframe으로 이동   # iframe 찾는 방법은 9번 이하 참고
iframes = driver.find_elements(By.TAG_NAME, 'iframe')   
driver.switch_to.frame(iframes[1])  


# 6. [연구비수주현황조회]-[연구기간 : 진행기준] 클릭
button = driver.find_element(By.CSS_SELECTOR, "#DATE_GB2")
button.click()
button.is_selected         #체크박스, 라디오버튼은 is_selected사용하기 / 일반 click 누르는거랑 다르다


# 7. [연구비수주현황조회]-[과제분류] 체크박스 풀기
checkbox1 = driver.find_element(By.CSS_SELECTOR, "#PRJ_CATE_CD_02")    #연구시설장비비통합과제 체크박스 풀기
checkbox2 = driver.find_element(By.CSS_SELECTOR, "#PRJ_CATE_CD_04")    #인건비풀링과제 체크박스 풀기          
if checkbox1.is_selected(): checkbox1.click()     # 체크풀기(체크박스1)
if checkbox2.is_selected(): checkbox2.click()     # 체크풀기(체크박스2)
#if not checkbox.is_selected(): checkbox.click()  #체크안된거면 체크하기


# 8. [연구비수주현황조회]-[조회] 클릭
driver.find_element(By.CSS_SELECTOR, "#btnSearch").click()   
sleep(10)

# 9. [파일저장] 클릭
driver.find_element(By.CSS_SELECTOR, "#gridDonw > img").click()   
sleep(3)

# 10. 사유 적기
# ①사유창 열기
sleep(5)                                  #팝업이 뜨는 속도와 닫는 속도가 다를 수 있으니까 앞에서 sleep을 주는거다. 보통 5초 정도 준다.
all_windows = driver.window_handles       #뜨는 모든 팝업창을 다 가지고 오기 / 새로운 창이 떴다면, 새로운 창으로 전환
for window in all_windows :
    if window != main_window:             #메인창은 닫지마
        driver.switch_to.window(window)   #모든 iframe을 닫으면 안되니까 iframe들을 들리는거다.

# ② 사유 적고 확인 누르기
driver.find_element(By.CSS_SELECTOR,"#FILE_DOWN_DESC").send_keys("실적조회")
driver.find_element(By.CSS_SELECTOR, "#btnSaveDesc > a").click()   #확인버튼 클릭
sleep(3)

# 11. 다시 원래 프레임으로 돌아오기
driver.switch_to.default_content()

# 12. 창 닫기
driver.close()

# 13. 다운로드파일 이름 변경(오늘일자로)--안먹힘 ㅠㅠ
# ①다운로드된 파일 확인 및 이름 변경 
downloaded_files = os.listdir(config_info["folder"]["work"])

# ②파일이 다운로드되었는지 확인하고 파일명 변경 
today_date = datetime.today().strftime('%Y-%m-%d')

for file in downloaded_files:
    if file.endswith(".xlsx"):  # 파일 확장자에 맞게 수정
        old_file_path = os.path.join(config_info["folder"]["work"], file)
        new_file_name = f"{today_date}.xlsx"  # 오늘 날짜로 이름 변경
        new_file_path = os.path.join(config_info["folder"]["work"], new_file_name)
        os.rename(old_file_path, new_file_path)
        print(f"파일이 '{new_file_name}'로 저장되었습니다.")
        break


"""
######################  iframe 찾기  ###################################################
# 9. [연구비수주현황(연구자)] 들어가기
# ①연구비수주현황(연구자) 클릭을 위한 iframe찾기
iframes = driver.find_elements(By.TAG_NAME, 'iframe')           # 모든 iframe을 찾고, 올바른 iframe으로 전환
for i, iframe in enumerate(iframes):
    print(f"iframe {i}: {iframe.get_attribute('id')}, {iframe.get_attribute('id')}, {iframe.get_attribute('src')}")

# ②찾은 iframe으로 이동하기 (연구비수주현황(연구자) iframe으로 이동)
driver.switch_to.frame(iframes[1])    #/rstat_0001_01.act?MENU_SEQ=205 (이 번호는 검사해서 찾아보면 확인가능) → iframe 1번이랑 번호가 같음(205)

# ③연구비수주현황(연구자) 클릭
driver.find_element(By.CSS_SELECTOR, "#menu_id_680 > a").click()   

# 10. 다시 원래 프레임으로 돌아오기
driver.switch_to.default_content()

# 11. 컨텐츠 iframe 찾기
#①
iframes = driver.find_elements(By.TAG_NAME, 'iframe')           # 모든 iframe을 찾고, 올바른 iframe으로 전환
for i, iframe in enumerate(iframes):
    print(f"iframe {i}: {iframe.get_attribute('id')}, {iframe.get_attribute('id')}, {iframe.get_attribute('src')}")

# ②찾은 iframe으로 이동하기 (연구비수주현황(연구자) iframe으로 이동)
driver.switch_to.frame(iframes[2])    #/rstat_0040_01.act?MENU_SEQ=680 (이 번호는 검사해서 찾아보면 확인가능) → iframe 2번이랑 번호가 같음(680)

# ③연구비수주현황(연구자) 조회버튼 클릭
driver.find_element(By.CSS_SELECTOR, "#btnSearch").click()   """





