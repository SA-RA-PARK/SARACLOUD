import util as ut
import xlsxProcess as xlp

filename = "20241016_89900100056880_154004"
directory = ut.exedir('py')

xlp.toExcelErp(directory, filename)

#search_item = filename.split('_')[1]
#worker = get_worker(search_item)
#print(worker) #  {"Name" : "홍길동", "Email" : "hong@abcd.co.kr"}
