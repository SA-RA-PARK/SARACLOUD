import os
import util as ut
import xlsxProcess as xlp

filename = "20241016_89900100056880_154004"   #filename은 key값이다!
filename = "20241016_89900100056881_154004"

directory = ut.exedir('py')

#directory_workF = ut.select_folder_console(display=False)
#directory_workF = r'C:\IACFPYTHON\workF'   ##챗GPT한테 물어본 경로로, 배운내용으로 다시 생각해보기
#file_list = os.listdir(directory_workF)
#print(file_list)                           ##결과값으로 파일명 4개가 나열된다


xlp.toExcelErp(directory, filename)

#search_item = filename.split('_')[1]
#worker = get_worker(search_item)
#print(worker) #  {"Name" : "홍길동", "Email" : "hong@abcd.co.kr"}


