import xlrd

import xlwt
from datetime import datetime
from xlrd import xldate_as_tuple
#数据源头
# book = xlrd.open_workbook('./保险公司提供数据格式-中华、信美.xlsx')
def readExcel(excelItem):
    print('excelIndex=',excelItem)
    book = xlrd.open_workbook(file_contents=excelItem.read())
    return book.sheet_names()
