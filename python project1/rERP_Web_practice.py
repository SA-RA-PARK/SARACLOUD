from time import sleep
from selenium import webdriver
from datetime import datetime
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import os
import util as ut
import pandas as pd


config_info = ut.load_config()       
current_path = ut.exedir('py')
#print(config_info) 
config_info["folder"] = {
    "current" : current_path,
    "RnD" : os.path.join(current_path, "RnD_list\\")
} 

# 폴더 새로 만들기
ut.create_folder(config_info["folder"]["RnD"])      #파일이 없으면 만들고, 있으면 pass다 / RnD_list 폴더 만듦


# 브라우저 꺼짐 방지 옵션
chrome_options = Options()
# chrome_options.add_argument("headless")  # 헤드리스 모드 설정
chrome_options.add_experimental_option("detach", True)
chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])

# 저장할 위치 지정
download_path = config_info["folder"]["RnD"] 

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


# 5. [연구비수주현황조회] iframe으로 이동      # iframe 찾는 방법은 9번 이하 참고
iframes = driver.find_elements(By.TAG_NAME, 'iframe')   
driver.switch_to.frame(iframes[1])  


# 6. [연구비수주현황조회]-[연구기간] 변경
input_element = driver.find_element(By.CSS_SELECTOR,config_info["period"]["startdate_CSS"])       #시작날짜박스 찾기
input_element.click()                                                                             #찾은 날짜 박스 클릭
input_element.clear()                                                                             #디폴트 날짜 지우기 (방법1)
input_element.send_keys(config_info["period"]["start_dt"])                                        #원하는 끝나는 일자 입력
sleep(2)

input_element_e = driver.find_element(By.CSS_SELECTOR,config_info["period"]["enddate_CSS"])       #끝나는 날짜박스 찾기
input_element_e.send_keys(Keys.BACKSPACE * 10)                                                    #디폴트 날짜 지우기(백스페이스 10번) (방법2)
input_element_e.send_keys(config_info["period"]["end_dt"] + Keys.ENTER)                           #원하는 끝나는 일자 입력


# 7. [연구비수주현황조회]-[연구기간 : 진행기준] 클릭
button = driver.find_element(By.CSS_SELECTOR, "#DATE_GB2")
button.click()
button.is_selected         #체크박스, 라디오버튼은 is_selected사용하기 / 일반 click 누르는거랑 다르다
sleep(3)


# 8. [연구비수주현황조회]-[과제분류] 체크박스 풀기
checkbox1 = driver.find_element(By.CSS_SELECTOR, "#PRJ_CATE_CD_02")    #연구시설장비비통합과제 체크박스 풀기
checkbox2 = driver.find_element(By.CSS_SELECTOR, "#PRJ_CATE_CD_03")    #간접비과제 체크박스 풀기  
checkbox3 = driver.find_element(By.CSS_SELECTOR, "#PRJ_CATE_CD_04")    #인건비풀링과제 체크박스 풀기            
checkbox4 = driver.find_element(By.CSS_SELECTOR, "#PRJ_CATE_CD_05")    #교내과제 체크박스 풀기     
if checkbox1.is_selected(): checkbox1.click()     # 체크풀기(체크박스1)
if checkbox2.is_selected(): checkbox2.click()     # 체크풀기(체크박스2)
if checkbox3.is_selected(): checkbox3.click()     # 체크풀기(체크박스3)
if checkbox4.is_selected(): checkbox4.click()     # 체크풀기(체크박스4)
#if not checkbox.is_selected(): checkbox.click()  #체크안된거면 체크하기


# 9. [연구비수주현황조회]-[조회] 클릭
driver.find_element(By.CSS_SELECTOR, "#btnSearch").click()   
sleep(10)

# 10. [파일저장] 클릭
driver.find_element(By.CSS_SELECTOR, "#gridDonw > img").click()   
sleep(3)

# 11. 사유 적기
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

# 12. 다시 원래 프레임으로 돌아오기
driver.switch_to.default_content()

# 13. 창 닫기
driver.close() 

print("다운로드 완료")
