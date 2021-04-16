import xlrd

import xlwt
import time
from datetime import datetime
from xlrd import xldate_as_tuple
from django.conf import settings
import os
from collections import Counter 
taikangCompany = {
'R'	:'泰康人寿保险有限责任公司山西分公司',
'5'	:'泰康人寿保险有限责任公司四川分公司',
'F'	:'泰康人寿保险有限责任公司重庆分公司',
'T'	:'泰康人寿保险有限责任公司江西分公司',
'1'	:'泰康人寿保险有限责任公司北京分公司',
'O'	:'泰康人寿保险有限责任公司黑龙江分公司',
'I'	:'泰康人寿保险有限责任公司深圳分公司',
'8'	:'泰康人寿保险有限责任公司海南分公司',
'X'	:'泰康人寿保险有限责任公司宁夏分公司',
'4'	:'泰康人寿保险有限责任公司上海分公司',
'A'	:'泰康人寿保险有限责任公司江苏分公司',
'B'	:'泰康人寿保险有限责任公司浙江分公司',
'N'	:'泰康人寿保险有限责任公司河北分公司',
'6'	:'泰康人寿保险有限责任公司辽宁分公司',
'H'	:'泰康人寿保险有限责任公司湖南分公司',
'E'	:'泰康人寿保险有限责任公司天津分公司',
'2'	:'泰康人寿保险有限责任公司湖北分公司',
'W'	:'泰康人寿保险有限责任公司甘肃分公司',
'D'	:'泰康人寿保险有限责任公司河南分公司',
'C'	:'泰康人寿保险有限责任公司山东分公司'
}


gongyinCompany = {
'御享人生重大疾病保险'	:['工银安盛人寿御享人生重大疾病保险','保至105周岁'],
'附加安心意外伤害医疗保险（B款）'	:['工银安盛人寿附加安心意外伤害医疗保险（B款）','1年'],
'御享颐生重大疾病保险'	:['工银安盛人寿御享颐生重大疾病保险','保至105周岁'],
'附加综合意外伤害保险'	:['附加综合意外伤害保险-工银安盛','1年'],
'康至优选医疗保险'	:['工银安盛人寿康至优选医疗保险','1年'],
'附加豁免保险费重大疾病保险'	:['附加豁免保险费重大疾病保险-工银安盛','保至105周岁'],
'e+保医疗保险-转保版'	:['工银安盛人寿e+保医疗保险','1年'],
'e+保医疗保险'	:['工银安盛人寿e+保医疗保险','1年'],
'附加安心住院费用医疗保险'	:['工银安盛人寿附加安心住院费用医疗保险','1年'],
'质子/重离子医疗补偿金'	:['质子/重离子医疗补偿金',''],
'附加豁免保险费定期寿险（2016）'	:['附加豁免保险费定期寿险（2016）','1年'],
'康至优选医疗保险-转保版'	:['工银安盛人寿康至优选医疗保险','1年'],
'御享颐生重大疾病保险（尊享版）'	:['工银安盛人寿御享颐生重大疾病保险（尊享版）','保至105周岁'],
'附加住院津贴医疗保险'	:['附加住院津贴医疗保险-工银安盛','1年'],
}
bookmodel = xlrd.open_workbook(os.path.join(settings.BASE_DIR, 'myExcel/excel/投保数据导入模板.xlsx')) # 投保数据导入模板 (5).xlsx
sheetmodel = bookmodel.sheet_by_name('字段对照表')
titlemodel = [elem for elem in sheetmodel.row_values(2)[1:] if elem != '']
sheetmodelNcols = len(titlemodel)
def mainExcel(book,bookrule):
    #数据源头book;对照字段bookrule
    global bookmodel
    global sheetmodel
    global titlemodel
    global sheetmodelNcols
    bookmodel = xlrd.open_workbook(os.path.join(settings.BASE_DIR, 'myExcel/excel/投保数据导入模板.xlsx')) # 投保数据导入模板 (5).xlsx
    sheetmodel = bookmodel.sheet_by_name('字段对照表')
    titlemodel = [elem for elem in sheetmodel.row_values(2)[1:] if elem != '']
    sheetmodelNcols = len(titlemodel)

    sheetrule = bookrule.sheet_by_name('字段对照表')
    sheetRuleCol = sheetrule.col_values(0)
    jtrow = sheetrule.row_values(sheetRuleCol.index('江泰'))
    #保存用
    workbook = xlwt.Workbook(encoding='utf-8')

    sheetList = [] #表数据列表
    titleDictList = [] #表数据title字典列表
    sheetDictList = [] #表数据字典列表
    ruleDictList = [] #表头对照字典列表
    resultDict = {} #excel结果
    for i in book.sheet_names():
        sheet = book.sheet_by_name(i)#数据源sheet数据
        sheetList.append(sheet)

        titleDict = makeTitleDict(sheet,i)
        titleDictList.append(titleDict)

        listSheet = listSheetMaker(sheet,titleDict,i)# 数据源处理成【字典...】
        sheetDictList.append(listSheet)

        ruleDict = ruleDictMaker(i,sheetrule)
        ruleDictList.append(ruleDict)

        results = resultsFactory(listSheet,ruleDict,i)
        
        resultsToWorksheet(workbook.add_sheet(i,cell_overwrite_ok=True),results)
        
        resultDict[i] = results
    return resultDict

def makeTitleDict(sheet,sheetName):
    #单个数据源title字典
    titleDict = {}
    titleList = repeatIndexMaker(sheet.row_values(0),sheetName)
    for name in titleList:
        titleDict[name] = None
    return titleDict

def listSheetMaker(sheet,titleDict,sheetName):
    # 数据源处理成【字典...】,sheet.nrows行数56,sheet.ncols列数66,sheet.col_values(0)列数组
    listSheet = [titleDict.copy() for i in range(sheet.nrows-1)]
    titleList = repeatIndexMaker(sheet.row_values(0),sheetName)
    for i in range(sheet.nrows):
        for j in range(sheet.ncols):
            if i!=0:
                if sheet.cell(i,j).ctype==3:
                    #日期处理
                    date = datetime(*xldate_as_tuple(sheet.row_values(i)[j], 0))
                    listSheet[i-1][titleList[j]] = date.strftime('%Y-%m-%d')
                elif sheet.cell(i,j).ctype==2 and int(sheet.cell_value(i,j))==sheet.cell_value(i,j):
                    #数字类型，且整型等于原有数据，直接转化为整型
                    listSheet[i-1][titleList[j]] = int(sheet.row_values(i)[j])
                else:
                    listSheet[i-1][titleList[j]] = sheet.row_values(i)[j]
    return listSheet
def repeatIndexMaker(item,name):
    #重复字段加上数字
    repeatDict = {
        '中华': ['受益人顺序','受益人比例'],
        '天安': ['受益人姓名','受益人关系','受益人证件号码','受益比例','受益人顺序','邮编','投保人邮箱','投保人国籍','投保人联系电话']
    }
    if name in repeatDict:
        repeatItem = {key:value for key,value in dict(Counter(item)).items()if key in repeatDict[name]}
        for i in repeatItem:
            index = item.index(i)
            countNum = 0
            while countNum < repeatItem[i]:
                item[index] += str(countNum)
                countNum = countNum + 1
                if countNum < repeatItem[i]:
                    index = item.index(i,index)
    return item
def ruleDictMaker(cell_value,sheetrule):
    #表头对照字典(key=导出表格title;value=数据源title)
    ruleDict = {}
    sheetRuleCol = sheetrule.col_values(0)
    jtrow = sheetrule.row_values(sheetRuleCol.index('江泰'))
    zhrow = repeatIndexMaker(sheetrule.row_values(sheetRuleCol.index(cell_value)),cell_value)
    zhrowDef = sheetrule.row_values(sheetRuleCol.index(cell_value)+1)
    for i in range(sheetrule.ncols):
        if i!=0:
            if len(jtrow[i])!=0:
                if len(zhrow[i])!=0:
                    ruleDict[jtrow[i]] = {'value':zhrow[i],'defValue':zhrowDef[i]}
                elif  zhrowDef[i]!='':
                    if isinstance(zhrowDef[i],float) and int(zhrowDef[i])==zhrowDef[i]:
                        ruleDict[jtrow[i]] = {'value':int(zhrowDef[i]),'defValue':int(zhrowDef[i])}
                    else:
                        ruleDict[jtrow[i]] = {'value':zhrowDef[i],'defValue':zhrowDef[i]}
            else:#合并单元，将上一个数据改为列表
                if(len(zhrow[i])!=0):
                    ruleDict[jtrow[i-1]] = [{'value':ruleDict[jtrow[i-1]]['value'],'defValue':ruleDict[jtrow[i-1]]['defValue']},{'value':zhrow[i],'defValue':zhrowDef[i]}]
    return ruleDict
def getListSheetKey(rd,listSheet,i,defValue):
    if rd in listSheet[i].keys():
        ls = listSheet[i][rd]
    else:
        #默认数据
        ls = defValue
    return str(ls)
def resultsFactory(listSheet,ruleDict,sheetName):
    #拼凑数据，过滤数据
    global titlemodel
    global sheetmodelNcols
    reduceRowIndex = []#需要被删除数据的index
    gongyinObj = {'reduceRowIndex':[],'saveArr':[]}#需要被删除数据的index
    results = [['' for j in range(sheetmodelNcols+1)] for i in range(len(listSheet)+1)]
    for i in range(len(listSheet)):
        lsItemList = []
        for j in range(sheetmodelNcols):
            ls = ''
            rd = ''
            tm = titlemodel[j]
            if tm in ruleDict.keys():
                rd = ruleDict[tm]
            else:
                #表头对照字典不存在直接跳过
                continue
            if isinstance(rd,list):
                #合并单元格
                for rd in rd:
                    ls += getListSheetKey(rd['value'],listSheet,i,rd['defValue'])
            elif isinstance(rd['value'],str):
                #普通单元格
                ls = getListSheetKey(rd['value'],listSheet,i,rd['defValue'])
            else:
                ls = rd['value']
            lsItem = lsChange(ls,j,sheetName,i,reduceRowIndex,results,gongyinObj)
            ls = lsItem['ls']
            results[i+1][j] = ls
            if lsItem['lsx']!=None:
                #记录需要处理的数据
                lsItemList.append(lsItem)
        if len(lsItemList)>0:
            #前置数据确定后置数据，在一条完整数据生成后进行处理
            for lsi in lsItemList:
                for index in range(len(lsi['lsJudge'])):
                    if results[lsi['lsx']][lsi['lsy'][0]]==lsi['lsJudge'][index] or lsi['lsJudge'][index]=='other':
                        for indexlsy in lsi['lsy']:
                            results[lsi['lsx']][indexlsy] = lsi['lsExpect'][index]
                        break
    for i in range(sheetmodelNcols):
        results[0][i] = titlemodel[i]
    resNumArr = [item[titlemodel.index('投保单号')] for item in results][1:]
    saveNumArr = [item[titlemodel.index('投保单号')] for item in gongyinObj['saveArr']]
    baofeiIndex = titlemodel.index('保费')
    policyIndex = titlemodel.index('投保单号')
    riskIndex = titlemodel.index('险种名称')
    for indexNum in range(len(gongyinObj['reduceRowIndex'])):
        currentIndex = resNumArr.index(results[gongyinObj['reduceRowIndex'][indexNum]][policyIndex])#等待加保费的index
        results[currentIndex][baofeiIndex] = str(float(results[currentIndex][baofeiIndex]) + float(gongyinObj['saveArr'][indexNum][baofeiIndex]))
    countRed = 0
    for i in gongyinObj['reduceRowIndex']:
        del results[i-countRed]
        countRed += 1
    countRed = 0
    for i in reduceRowIndex:
        del results[i-countRed]
        countRed += 1
    return results
def lsChange(ls,j,sheetName,i,reduceRowIndex,results,gongyinObj):
    lsx = None # 第几条数据
    lsy = None # 数据第几个值
    lsName = None # 数据title
    lsJudge = None # 需要修改位置上的数据判断
    lsExpect = None # 需要修改位置上的数据值
    #写入数据前，个性化修改数据
    if(sheetName=='中华'):
        if(j==titlemodel.index('供应商出单公司')):
            ls = '中华联合人寿保险股份有限公司'+ls+'分公司'
        if(j==titlemodel.index('江泰出单机构') and ls=='江泰保险经纪股份有限公司'):
            ls = ls+'北京分公司'
        if(j==titlemodel.index('保单状态')):
            if(ls=='已终止'):
                ls='终止'
                lsx = i+1
                lsy = [titlemodel.index('退保类型'),titlemodel.index('退保原因')]
                lsName = '退保类型'
                lsJudge = ['犹退终止','other']
                lsExpect = ['犹豫期内退保','犹豫期外退保']
            elif(ls=='投保中止' or ls=='核保不通过'):
                ls='删除本条记录'
                reduceRowIndex.append(i+1)
            else:
                ls='有效'
        if(j==titlemodel.index('退保类型')):
            if '犹' not in ls:
                ls = ''
        if(j==titlemodel.index('投保单号')):
            if not ls:
                ls='删除本条记录'
                reduceRowIndex.append(i+1)
    elif(sheetName=='信美'):
        if(j==titlemodel.index('生效时间')):
            ls = ls.split('生效')[0]
        if(j==titlemodel.index('退保时间')):
            ls = ls.split('生效')[0]
        if(j==titlemodel.index('主附险标识')):
            ls = '主险'
        if(j==titlemodel.index('险种名称')):
            if('-' in ls):
                ls = ls.split('-')[0]
        if(j==titlemodel.index('保单状态')):
            if(ls=='有效'):
                lsx = i+1
                lsy = [titlemodel.index('退保时间')]
                lsName = '退保时间'
                lsJudge = ['other']
                lsExpect = ['']
    elif(sheetName=='天安'):
        if(j==titlemodel.index('供应商出单公司')):
            if ls=='北京营业总部':
                ls = '天安人寿保险股份有限公司北京分公司'
            else:
                ls = '天安人寿保险股份有限公司'+ls
        if(j==titlemodel.index('主附险标识')):
            if ls=='是':
                ls = '主险'
            elif ls=='否':
                ls = '附加险'
        if(j==titlemodel.index('第二受益人顺序')):
            if ls=='1':
                ls = '2'
        if(j==titlemodel.index('第三受益人顺序')):
            if ls=='1':
                ls = '3'
        if(j==titlemodel.index('险种名称')):
            if ls=='天安人寿附加住院费用医疗保险':
                ls = '附加住院费用医疗保险-天安'
            if ls=='天安人寿附加住院津贴医疗保险':
                ls = '附加住院津贴医疗保险-天安'
    elif(sheetName=='君康'):
        if(j==titlemodel.index('供应商出单公司')):
            ls = '君康人寿保险股份有限公司'+ls+"分公司"
        if(j==titlemodel.index('江泰出单机构')):
            ls = ls.split('营业本部')[0].split('本部')[0]
        if(j==titlemodel.index('退保类型')):
            if ls=='犹豫期退保':
                ls = '犹豫期内退保'
            if ls=='退保':
                ls = '犹豫期外退保'
        if(j==titlemodel.index('主附险标识')):
            if ls=='Y':
                ls = '主险'
            if ls=='N':
                ls = '附加险'
        if(j==titlemodel.index('险种名称')):
            # 去修改其他位置的数据
            lsx = i+1
            lsy = [titlemodel.index('保险期间')]
            lsName = '保险期间'
            lsJudge = ['other']
            if ls.find('医疗')!=-1:
                lsExpect = ['1年']
            else:
                lsExpect = ['保至105周岁']
    elif(sheetName=='弘康'):
        if(j==titlemodel.index('供应商出单公司')):
            ls = '弘康人寿保险股份有限公司'
        if(j==titlemodel.index('保单状态')):
            if ls == '承保':
                ls = '有效'
            if ls == '退保终止' or ls == '犹退终止':
                ls = '终止'
        if(j==titlemodel.index('退保类型')):
            if ls == '犹豫期退保':
                ls = '犹豫期内退保'
            if ls == '退保':
                ls = '犹豫期外退保'
        if(j==titlemodel.index('与投保人关系') or j==titlemodel.index('第一受益人与被保人关系')):
            if ls == '丈夫' or ls == '妻子':
                ls = '配偶'
            if ls == '儿子' or ls == '女儿':
                ls = '子女'
            if ls == '孙女' or ls == '孙子' or ls == '外孙' or ls == '外孙女':
                ls = '其他'
        if(j==titlemodel.index('受益类型')):
            if ls == '身故受益人':
                ls = '指定'
            if ls == 'null':
                ls = '法定'
    elif(sheetName=='工银'):
        if(j==titlemodel.index('保单状态')):
            if ls == '承保未生效':
                ls = '有效'
        if(j==titlemodel.index('缴费方式')):
            if ls == '1':
                ls = '一次交清'
            else:
                ls = '年交'
        if(j==titlemodel.index('险种名称')):
            if ls in gongyinCompany:
                print(ls in gongyinCompany,ls)
                print(gongyinCompany[ls])
                if gongyinCompany[ls][0]=='质子/重离子医疗补偿金':
                    gongyinObj['reduceRowIndex'].append(i+1)
                    gongyinObj['saveArr'].append(results[i+1])
                ls = gongyinCompany[ls][0]
        if(j==titlemodel.index('保险期间')):
            if ls in gongyinCompany:
                ls = gongyinCompany[ls][1]
        if(j==titlemodel.index('投保时间')):
            # 时间减一天
            if ls:
                timeStruct = time.strptime(ls, "%Y-%m-%d") 
                timeStamp = int(time.mktime(timeStruct)) - 60*60*24
                localTime = time.localtime(timeStamp) 
                strTime = time.strftime("%Y-%m-%d", localTime) 
                ls = strTime
    elif(sheetName=='泰康'):
        if(j==titlemodel.index('供应商出单公司')):
            ls = taikangCompany[ls]
        if(j==titlemodel.index('保单状态')):
            if ls != '有效':
                ls = '终止'
        if j == titlemodel.index('投保时间') or j == titlemodel.index('交费时间') or j == titlemodel.index('承保时间') or j == titlemodel.index('生效时间') or j == titlemodel.index('保单签发时间') or j == titlemodel.index('回执时间') or j == titlemodel.index('回访时间') or j == titlemodel.index('退保时间') or j == titlemodel.index('终止时间'):
            if ls:
                ls = ls[:4] + '-' + ls[4:6] + '-' + ls[6:8]
        if(j==titlemodel.index('险种名称')):
            if ls=='岁月有约养老年金保险产品计划':
                ls = '泰康岁月有约养老年金保险（分红型）'
        if(j==titlemodel.index('缴费方式')):
            if ls=='1' or ls=='1年缴清' or '缴至' in ls:
                ls = '一次交清'
            if '年缴清' in ls:
                ls = '年交'
    elif(sheetName=='信泰'):
        if(j==titlemodel.index('供应商出单公司')):
            ls = ls.replace('信泰保险','信泰人寿保险股份有限公司')
            ls = ls.replace('本部销售','')
    #所有保险公司字段的特殊处理
    if j == titlemodel.index('江泰出单机构'):
        if '有限公司' in ls:
            ls = ls.split('有限公司')[1]
    if(j==titlemodel.index('保单状态')):
        if ls=='生效' or ls=='待生效':
            ls = '有效'
    if(j==titlemodel.index('险种名称')):
        if ls.find('(')!=-1:
            ls = ls.replace('(','（')
            ls = ls.replace(')','）')
    if j == titlemodel.index('保险期间'):
        lsNum = getNumberFromString(ls)
        if len(lsNum)>0:
            if lsNum != '105':
                if int(lsNum)>30:
                    if int(lsNum)>=999:
                        ls = '保至105周岁'
                    else:
                        ls = '保至'+lsNum+'周岁'
                else:
                    ls = lsNum+'年'
            else:
                ls = '保至105周岁'
        if ls == '终身':
            ls = '保至105周岁'
    if j == titlemodel.index('缴费方式'):
        if ls=='趸交' or ls=='趸缴':
            ls = '一次交清'
        if ls=='按年交':
            ls = '年交'
    if j == titlemodel.index('缴费期间'):
        if ls=='0' or ls=='1' or '周岁' in ls:
            ls = '1'
        if '年' in ls:
            ls = ls.split('年')[0]
    if j == titlemodel.index('与投保人关系'):
        if ls=='夫妻':
            ls = '配偶'
    if j == titlemodel.index('投保人证件类型') or j == titlemodel.index('被保险人证件类型'):
        if ls=='居民身份证':
            ls = '身份证'
        if ls=='户口簿':
            ls = '户口本'
        if ls=='出生医学证明':
            ls = '出生证'
    if j == titlemodel.index('投保时间') or j == titlemodel.index('交费时间') or j == titlemodel.index('承保时间') or j == titlemodel.index('生效时间') or j == titlemodel.index('保单签发时间') or j == titlemodel.index('回执时间') or j == titlemodel.index('回访时间') or j == titlemodel.index('退保时间') or j == titlemodel.index('终止时间'):
        if ls:
            ls = ls.replace('/','-').split(' ')[0]
    if ls == 'null':
        ls = ''
    if j == titlemodel.index('被保险人性别'):
        if '性' in ls:
            ls = ls.split(' ')[0]
    return {
        'ls': ls,
        'lsx': lsx,
        'lsy': lsy,
        'lsName': lsName,
        'lsJudge': lsJudge,
        'lsExpect': lsExpect
    }
def getNumberFromString(ls):
    #从字符串中获取数字
    num = ''
    numIterator = filter(str.isdigit,ls)
    try:
        while True:
            num += next(numIterator)
    except StopIteration:
        pass
    return num
def resultsToWorksheet(savesheet,results):
    #将数据写入excel中
    for i in range(len(results)):
        for j in range(len(results[0])):
            savesheet.write(i,j,results[i][j])

