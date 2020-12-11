from django.http import HttpResponse, StreamingHttpResponse
from django.shortcuts import render
import sys
sys.path.append(r'myExcel')
from excel.excelIndex import readExcel
from excel.excelFunc import mainExcel
from excel.time import getTime
from excel.tableMaker import tableMakerMain
import xlwt
import xlrd
from datetime import datetime, timedelta, timezone
import os
import requests
import chardet
import json
from django.core.files.storage import default_storage
from django.core.files.base import ContentFile
from django.conf import settings
# 表单
def search_form(request):
    # return render(request, 'index.html')
    return request
def file_iterator(file_name, chunk_size=512):
    '''
    # 用于形成二进制数据
    :return:
    '''
    print('file_name===>',file_name)
    with open(file_name, 'rb') as f:
        while True:
            c = f.read(chunk_size)
            if c:
                yield c
            else:
                break
def tamaker(request):
    bookruleFile = request.FILES.get('fileRule')
    bookrule = xlrd.open_workbook(file_contents=bookruleFile.read())#对照字段
    bookFile = request.FILES.get('fileTa')
    book = xlrd.open_workbook(file_contents=bookFile.read())#数据源头
    results = tableMakerMain(book,bookrule)
    print('results====>>>>',results)
    return HttpResponse(json.dumps(results))
def searchTest(request):
    bookruleFile = request.FILES.get('fileRule')
    bookrule = xlrd.open_workbook(file_contents=bookruleFile.read())#对照字段
    bookFile = request.FILES.get('fileDate')
    book = xlrd.open_workbook(file_contents=bookFile.read())#数据源头
    saveExcel(bookrule)
    results = mainExcel(book,bookrule)
    print('results====>>>>',results)
    return HttpResponse(json.dumps(results))
def saveExcel(fileData):
    # 将xlrd格式的文件存储到本地
    workbook = xlwt.Workbook(encoding='utf-8')
    workbookWrite = workbook.add_sheet('字段对照表',cell_overwrite_ok=True)
    for i in range(fileData.sheet_by_name('字段对照表').ncols):
        for j in range(fileData.sheet_by_name('字段对照表').nrows):
            workbookWrite.write(j,i,fileData.sheet_by_name('字段对照表').cell_value(j,i))
    workbook.save('./myExcel/excel/投保数据导入模板.xlsx')
def saveFile(fileData):
    file_handle=open('excelLib/1.xls',mode='wb+')
    file_handle.write(fileData.read())
    file_handle.close()
    # print('settings.MEDIA_ROOT=======>',settings.MEDIA_ROOT)
    # print('fileData=======>',fileData)
    # path = default_storage.save('excelLib/somename.xls', ContentFile(fileData.read()))
    # tmp_file = os.path.join('', path)
    # print('path=======>',path)
def searchTestforA(request):
    #保存用
    response = HttpResponse(content_type='application/vnd.ms-excel') 
    response['Content-Disposition'] = 'attachment; filename=DEMO.xls' 
    workbook = xlwt.Workbook(encoding='utf-8')
    savesheet = workbook.add_sheet('newsheet',cell_overwrite_ok=True)
    savesheet.write(0,0,123456)
    workbook.save(response) 
    return response
# 接收请求数据
def search(request):

    bookruleFile = request.FILES.get('fileRule')
    bookrule = xlrd.open_workbook(file_contents=bookruleFile.read())#对照字段
    bookFile = request.FILES.get('fileDate')
    book = xlrd.open_workbook(file_contents=bookFile.read())#数据源头
    workbook = mainExcel(book,bookrule)
    timeStr = getTime()
    workbook.save('./excelLib/导出数据'+timeStr+'.xls') 
    return HttpResponse("true")

    
    # return HttpResponse(readExcel(obj))
def searchDir(request):
    file_dir = os.path.join(os.path.abspath(__file__).split('myExcel')[0], 'myExcel/excelLib')
    for root,dirs,files in os.walk(file_dir):  
        # print('root',root) #当前目录路径  
        # print('dirs',dirs) #当前路径下所有子目录  
        print('files',files) #当前路径下所有非目录子文件  
    return HttpResponse(';'.join(files))
def getOthersApi(request):
    headers = {"Cookie" : "pgv_pvi=2230124544; RK=Y0ixOeLINK; ptcz=89bd93e62490fc6f6ae649e3960a2137be49087e191f3b64c0c5b31089d93a75; pgv_pvid=4539032033; o_cookie=1973231806; pac_uid=1_1973231806; ts_uid=9753643830; aics=1Q2hVh8VE36y6dqPZLbGlSOXkStxtcN6RR1hvzDt; ts_refer=www.baidu.com/link; pgv_si=s8489264128; tvfe_boss_uuid=3513dfed55bfe87f; pgv_info=ssid=s7588147420; sd_userid=84791599529161440; sd_cookie_crttime=1599529161440; PHPSESSID=isdnvuk3g2dceks5ota1h9n8u1; Tk_EbR=CPBHeQjJhhwfQ4gUr5EWOqn6bKj2mxyJx%2Bjez7s2xZ18SIqG8cUmT%2F3gSmtu6FsOC7vkDUXCitO5aerewgddcYrQKZjPg%2FiSM7pAcwnMhsVaAUszHxs%2B4Bh82s4ZL%2Fx8SL37STl5AgCR9HqdJFX8%2FKR%2FhKk7yJMo54ai5%2FMllo9SixNJVkk0%2BuKY8JEbgR4wFGc2lJjrKzIwLa%2BwKf3EY%2Fugs5bGNDFYdhRPld%2BfHz4XaIhqdKuXrHuVryQ0j3g%2B31u1mahnbck3g8uYI3ZrGTrZ%2Bilr8v4N0YutQd%2BfllfAKM1skj8a956nKQ60EPhfPuLDA76A9VFmHDisbYlaKQ%3D%3D; Tk_EbR_api=ySQJ2MO%2Fep2d7Zfcyo4HyKTI7ThRnoRLLbu8pXtXB6ThZu0Lhl1i%2FcMTz6r9BDSqK3RDiB5jtrzVTlVB8khZTrL0yoRsTY1Da9x%2FZ0OCJQlTmE4ljgjdtu%2B8vAGsWzuc3t2eXMWD5gFEGRCVzbS7yIjJjWS3SLEQ8aD5I7R8U5Mr3yJdpairoY9rZEBcF9z%2FvbyuIhnGtdKu1qKZw69SoQ%3D%3D; m_check=c5ce94ab; ts_last=mta.qq.com/h5/base/ctr_core_data"}
    url = 'https://mta.qq.com/h5/base/ctr_core_data/get_table_realtime?need_compare=0&start_compare_date=&end_compare_date=&rnd=1601032889128&ajax=1'
    app_id = request.GET.get("app_id")
    print("app_id====>",request.GET.get("app_id"))
    params = {'app_id':app_id,'start_date':'2020-07-01','end_date':'2020-09-28'}
    get_data = requests.get(url,headers=headers,params=params)
    return HttpResponse(get_data)