from time import sleep
from selenium import webdriver
from datetime import datetime
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import os
import util as ut
import pandas as pd

def Scival_Web_main() :
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


    # 1. 웹페이지 해당 주소(scival.com)로 이동
    driver.get(config_info["scival"]["URL"])       #config.txt에 있는 url을 가지고 오는거다.
    driver.maximize_window()
    #sleep(2)
  
    main_window = driver.current_window_handle     #떠있는 화면을 handling하는거다.
     
    # 2. scival 로그인하기
    driver.find_element(By.CSS_SELECTOR,"#gh-main-cnt > div:nth-child(4) > a").click()                           #Sign in 버튼 누르기
    driver.find_element(By.CSS_SELECTOR,config_info["scival"]["ID_CSS"]).send_keys(config_info["scival"]["ID"])  #ID입력
    driver.find_element(By.CSS_SELECTOR,config_info["scival"]["CONTINUE_BT_CSS"]).click()                        #Continue 버튼 누르기
    driver.find_element(By.CSS_SELECTOR, config_info["scival"]["PW_CSS"]).send_keys(config_info["scival"]["PW"]) #pw입력
    driver.find_element(By.CSS_SELECTOR,config_info["scival"]["LOGIN_BT_CSS"]).click()                           #Sign in 버튼 누르기

    # 3. Compare-Benchmark All metrics 메뉴 들어가기
    driver.find_element(By.CSS_SELECTOR,"#aa-globalheader-navigation-link2").click()                             #compare 메뉴 클릭
    driver.find_element(By.CSS_SELECTOR,"##aa-globalheader-navigation-link2-1 > span").click()                   #All metrics 메뉴 클릭
    
    # 4. Year Range 설정
    driver.find_element(By.CSS_SELECTOR,"#yearRangeSelector > div.sliderToggle > span.sliderHandle > svg:nth-child(1)").click()   #Year Range 클릭
    input_element_e.send_keys(Keys.BACKSPACE)                                                                      #기존에 적혀있던 연도 backspace로 삭제
    #input_element = driver.find_element(By.CSS_SELECTOR,config_info["period"]["startyr_CSS"])                     #시작연도 박스 찾기
    #input_element.click()                                                                                         #찾은 연도 박스 클릭
    #input_element.clear()                                                                                         #디폴트 연도 지우기 (방법1)
    input_element.send_keys(config_info["period"]["start_yr"])                                                     #원하는 시작연도 입력
    
    input_element_e = driver.find_element(By.CSS_SELECTOR,config_info["period"]["endyr_CSS"])                      #끝나는 날짜박스 찾기
    input_element_e.send_keys(Keys.BACKSPACE)                                                                      #디폴트 날짜 지우기(백스페이스) (방법2)
    input_element_e.send_keys(config_info["period"]["end_yr"] + Keys.ENTER)                                        #원하는 끝나는 일자 입력
 
    driver.find_element(By.CSS_SELECTOR,"#commitYearRangeUpdate").click()                                          #Apply 버튼 누르기


    # 5. 비교대학 선택
    input_element = driver.find_element(By.CSS_SELECTOR,"#addMoreText_institution")                               #대학명 검색 박스 찾기
    input_element.click()                                                                                         #대학명 검색 박스 클릭

    input_element.send_keys(config_info["university"]["CA_univ"])                                                 #대학1 - 중앙대학교 입력
    driver.find_element(By.CSS_SELECTOR,"#addMoreText_institution").click()                                       #체크하기
    driver.find_element(By.CSS_SELECTOR,config_info["university"]["add_CSS"]).click()                             #add to selection

    input_element.send_keys(config_info["university"]["DGIST_univ"])                                              #대학2 - 대구경북과기원 입력
    driver.find_element(By.CSS_SELECTOR,"#addMoreText_institution").click()                                       #체크하기
    driver.find_element(By.CSS_SELECTOR,config_info["university"]["add_CSS"]).click()                             #add to selection

    input_element.send_keys(config_info["university"]["HY_univ"])                                                 #대학3 - 한양대 입력
    driver.find_element(By.CSS_SELECTOR,"#addMoreText_institution").click()                                       #체크하기
    driver.find_element(By.CSS_SELECTOR,config_info["university"]["add_CSS"]).click()                             #add to selection

    input_element.send_keys(config_info["university"]["KAIST_univ"])                                              #대학 4 - 한국과기원 입력
    driver.find_element(By.CSS_SELECTOR,"#addMoreText_institution").click()                                       #체크하기
    driver.find_element(By.CSS_SELECTOR,config_info["university"]["add_CSS"]).click()                             #add to selection

    input_element.send_keys(config_info["university"]["KR_univ"])                                                 #대학 5 - 고려대 입력
    driver.find_element(By.CSS_SELECTOR,"#addMoreText_institution").click()                                       #체크하기
    driver.find_element(By.CSS_SELECTOR,config_info["university"]["add_CSS"]).click()                             #add to selection

    input_element.send_keys(config_info["university"]["KH_univ"])                                                 #대학 6 - 경희대 입력
    driver.find_element(By.CSS_SELECTOR,"#addMoreText_institution").click()                                       #체크하기
    driver.find_element(By.CSS_SELECTOR,config_info["university"]["add_CSS"]).click()                             #add to selection

    input_element.send_keys(config_info["university"]["POSTECH_univ"])                                            #대학 7 - 포항공대 입력
    driver.find_element(By.CSS_SELECTOR,"#addMoreText_institution").click()                                       #체크하기
    driver.find_element(By.CSS_SELECTOR,config_info["university"]["add_CSS"]).click()                             #add to selection    

    input_element.send_keys(config_info["university"]["SN_univ"])                                                 #대학 8 - 서울대 입력
    driver.find_element(By.CSS_SELECTOR,"#addMoreText_institution").click()                                       #체크하기
    driver.find_element(By.CSS_SELECTOR,config_info["university"]["add_CSS"]).click()                             #add to selection     

    input_element.send_keys(config_info["university"]["SKK_univ"])                                                #대학 9 - 성균관대 입력
    driver.find_element(By.CSS_SELECTOR,"#addMoreText_institution").click()                                       #체크하기
    driver.find_element(By.CSS_SELECTOR,config_info["university"]["add_CSS"]).click()                             #add to selection     

    input_element.send_keys(config_info["university"]["UNIST_univ"])                                              #대학 10 - 울산과기원 입력
    driver.find_element(By.CSS_SELECTOR,"#addMoreText_institution").click()                                       #체크하기
    driver.find_element(By.CSS_SELECTOR,config_info["university"]["add_CSS"]).click()                             #add to selection   

    input_element.send_keys(config_info["university"]["YS_univ"])                                                 #대학 11 - 연세대 입력
    driver.find_element(By.CSS_SELECTOR,"#addMoreText_institution").click()                                       #체크하기
    driver.find_element(By.CSS_SELECTOR,config_info["university"]["add_CSS"]).click()                             #add to selection      


    # 6. 지표 Table로 보기 선택
    driver.find_element(By.CSS_SELECTOR,"#tableSelection > a").click()                                            #table  클릭

    # 7. 지표 선택하기
    driver.find_element(By.CSS_SELECTOR,"#yAxisBenchmarkHeaderMetric > button").click()                           #지표 1 - Scholarly Output(One metric over time)
    driver.find_element(By.CSS_SELECTOR,"#yAxisBenchmarkHeaderMetric_accordion_Published > li.metricItemOuter.Selected.Focused.snowball > button").click()
