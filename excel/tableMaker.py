import xlrd

import xlwt
from collections import Counter 
def filterBookListMaker(book,bookRule):
    sheetrule = bookRule.sheet_by_name('字段对照表')
    sheetRuleCol = sheetrule.col_values(0)
    taDef = sheetrule.row_values(sheetRuleCol.index('天安'))#天安表
    taofDef = sheetrule.row_values(sheetRuleCol.index('天安')+2)#天安表所属表
    filterBookList={}
    for index in range(len(taofDef)):
        if taofDef[index] in book.sheet_names():
            taSheetRow = book.sheet_by_name(taofDef[index]).row_values(0)
            if taofDef[index] in filterBookList:
                filterBookList[taofDef[index]].append(taSheetRow.index(taDef[index]))
            else:
                filterBookList[taofDef[index]] = [taSheetRow.index(taDef[index])]
    return filterBookList
def sheetToList(sheet,sheetName, filterBookList):
    #excel表转List
    ##所有sheet转list都在这里，利用rule中的对应关系，只保留rule中的字段，减少无用字段
    myList = [[] for i in range(sheet.nrows)]
    for i in range(sheet.nrows):
        for j in range(sheet.ncols):
            if j in filterBookList[sheetName] or sheet.cell_value(0,j) in ['险种代码','被保人编码','保单号码']:
                #'险种代码','被保人编码'是多表合并的关键数据，保留
                myList[i].append(sheet.cell_value(i,j))
    return myList
def listMerge(origList,list):
    # 简单合并sheet，不进行对比，默认对的上
    myList = []
    print(len(origList), len(list))
    for i in range(len(origList)):
        myList.append(origList[i]+list[i])
    return myList
def lackListMerge(origList,lackList,arr):
    # 缺省表合并sheet，不进行对比，缺省的后延
    myList = []
    print(len(origList), len(lackList))
    for i in range(len(origList)):
        if i==0:
            # title直接合并
            myList.append(origList[i]+lackList[i])
        elif i>0 and arr[i] != None:
            myList.append(origList[i]+lackList[arr[i]])
        else:
            # 缺省情况直接用空字符串补齐
            myList.append(origList[i]+['' for index in range(len(lackList[0]))])
    return myList
def resultsToWorksheet(savesheet,results):
    #将数据写入excel中
    for i in range(len(results)):
        for j in range(len(results[0])):
            try:
                res = results[i][j]
            except BaseException:
                res = ''
            savesheet.write(i,j,res)
    workbook.save('天安导出数据.xlsx')
def mergeElementForList(origList,mergeWhich):
    listArr = []
    indexList = []#险种代码被保人编码
    listX = origList[0].index(mergeWhich[0])
    listY = origList[0].index(mergeWhich[1])
    for item in origList:
        listArr.append([item[listX]+item[listY]]+item)
        indexList.append(item[listX]+item[listY])
    return {'list':listArr,'indexList':indexList}

def beneficiaryPushInsured(diffDict):
    #受益人表并入被保人
    #受益人表
    diffDictSheet = mergeElementForList(diffDict['受益人表']['sheetList'],['险种代码', '被保人编码'])
    #被保人表
    insuredSheet = mergeElementForList(diffDict['被保人表']['sheetList'],['险种代码', '被保人编码'])
    isi = insuredSheet['list'][0].index('险种代码')
    for index in range(len(insuredSheet['indexList'])):
        if insuredSheet['indexList'][index] in diffDictSheet['indexList']:
            #存在受益人
            dsi = diffDictSheet['indexList'].index(insuredSheet['indexList'][index])
            if dsi==0:
                pushKey = '受益类型'
            else:
                pushKey = '指定'
            insuredSheet['list'][index] += [pushKey] + diffDictSheet['list'][dsi]
        else:
            #不存在受益人
            insuredSheet['list'][index] += ['法定'] + ['' for i in range(len(diffDictSheet['list'][0]))]
    return insuredSheet['list']
def balaPushRisk(diffDict,beneficiaryAndInsured):
    #被保人表（+受益人表）并入险种表，并入过程扩充险种表
    #险种表的compareList并入后失效
    baiList = beneficiaryAndInsured#被保人表（+受益人表）
    baiIndexList = [item[beneficiaryAndInsured[0].index('险种代码')]+item[beneficiaryAndInsured[0].index('保单号码')] for item in beneficiaryAndInsured]
    riskSheetList = diffDict['险种表']['sheetList']
    riskSheetListForCheck = [item[riskSheetList[0].index('险种代码')]+item[riskSheetList[0].index('保单号码')] for item in riskSheetList]
    rsiOri = riskSheetList[0].index('险种代码')
    policyOri = riskSheetList[0].index('保单号码')
    riskList = []
    for index in range(len(riskSheetListForCheck)):
        #险种表的险种代码字段，在被保人表（+受益人表）的出现次数
        # 保单号与险种编码合并到一起校验两条数据是否需要合并
        riskCode = riskSheetListForCheck[index] #险种表当前循环中的 险种代码+保单号码
        repeatNum = dict( Counter(baiIndexList) )[riskCode]
        startIndex = 0
        if riskCode in baiIndexList:
            #险种表的险种代码 在 被保人表中存在
            if repeatNum == 1:
                # riskCode相同且唯一，合并为一条数据
                riskList.append(riskSheetList[index] + baiList[baiIndexList.index(riskCode)])
            else:
                #多条数据，根据被保人数据分成多条数据
                while repeatNum>0:
                    thisIndex = baiIndexList.index(riskCode,startIndex)
                    if riskSheetListForCheck[index] == baiIndexList[thisIndex]:
                        # riskCode相同才可以合并为一条数据
                        riskList.append(riskSheetList[index] + baiList[thisIndex])
                    else:
                        print('118=====>',index,thisIndex,riskCode,startIndex,repeatNum,riskSheetList[index][riskSheetList[0].index('保单号码')], baiList[thisIndex][baiList[0].index('保单号码')],baiList[thisIndex])
                        pass
                    startIndex = thisIndex+1
                    repeatNum -= 1
        else:
            #暂时不存在这种情况
            print('没有匹配到---balaPushRisk')
    return riskList
def riskMaker(riskList,sheetOrigList):
    #险种表，多险种处理成多条数据
    arr = []
    origPolicyList = [item[sheetOrigList[0].index('保单号码')] for item in sheetOrigList]
    riskPolicyList = [item[riskList[0].index('保单号码')] for item in riskList]
    riskCompareList = compareListMaker(arr,origPolicyList,riskPolicyList)
    print(riskCompareList,len(origPolicyList),len(riskPolicyList))
    sheetOrigListAddRisk = []
    for index in range(len(riskCompareList)):
        if isinstance(riskCompareList[index],list):
            #多险种数据
            for i in riskCompareList[index]:
                sheetOrigListAddRisk.append(sheetOrigList[index]+riskList[i])
        else:
            #单险种
            if riskCompareList[index]!=None:
                sheetOrigListAddRisk.append(sheetOrigList[index]+riskList[riskCompareList[index]])
    return sheetOrigListAddRisk
def compareListMaker(arr,origPolicyList,itemPolicyList):
    ##对比基础表的保单号码，生成每个表的index对照列表，提出独立函数
    for policyValue in origPolicyList:
        findStart = 0
        flag = True
        indexArr = []
        while flag:
            try:
                #循环找每个能够匹配上的数据
                index = itemPolicyList.index(policyValue,findStart)
                indexArr.append(index)
                findStart = index+1
            except BaseException:
                flag = False
        if len(indexArr)==1:
            arr.append(indexArr[0])
        elif len(indexArr)>1:
            arr.append(indexArr)
        else:
            #查找不到的数据
            arr.append(None)
    return arr
def isComplex(arr):
    #判断list中是否包含list
    for i in arr:
        if isinstance(i, list):
            #包含
            return 1
    #不包含
    return 0
def duplicateRemoval(comList):
    # 删除重复表头
    dupDict = {key:value for key,value in dict(Counter(comList[0])).items()if value>1 }
    arr = []
    for i in dupDict:
        count = dupDict[i]
        indexLast = 0
        while count>0:
            currentIndex = comList[0].index(i,indexLast)
            # 天安重复数据无需删除
            if indexLast != 0 and i not in ['受益人姓名','受益人关系','受益人证件号码','受益比例','受益人顺序','邮编','投保人邮箱','投保人国籍','投保人联系电话']:
                arr.append(currentIndex)
            indexLast = currentIndex + 1
            count -= 1
    arr.sort()
    arr.reverse()
    #至此，arr包含了所有要删除的表头index，从大到小排列
    for i in arr:
        for j in range(len(comList)):
            del comList[j][i]
    return comList
def tableMakerMain(book, bookRule):
    #数据源头 book
    # book = xlrd.open_workbook('./副本保险公司提供数据格式-天安20201012.xlsx')
    #数据规则 bookRule
    # bookRule = xlrd.open_workbook('./fileRule20200925.xlsx')
    filterBookList = filterBookListMaker(book,bookRule)
    #保存用
    workbook = xlwt.Workbook()
    sheetOrigList = []
    diffDict = {}
    for sheetIndex in range(len(book.sheet_names())):
        arr = []
        sheetName = book.sheet_names()[sheetIndex]
        if sheetName not in filterBookList:
            continue
        if sheetIndex==0:
            sheetOrig = book.sheet_by_name(sheetName)#基础表
            sheetOrigList = sheetToList(sheetOrig,sheetName,filterBookList)
            continue
        else:
            sheetItem = book.sheet_by_name(sheetName)#变化表
            sheetItemList = sheetToList(sheetItem,sheetName,filterBookList)
        arr = compareListMaker(arr,sheetOrig.col_values(1),sheetItem.col_values(1))
        whichFunc = isComplex(arr)
        #受益人表必须是复杂表
        if whichFunc == 0 and sheetName != '受益人表':
            # 简单表直接数据合并
            # print(arr,len(arr),sheetName,'简单',len(sheetItemList))
            if len(sheetOrigList) == len(sheetItemList):
                # 基础表数据与简单表数据个数一样
                sheetOrigList = listMerge(sheetOrigList, sheetItemList)
            elif len(sheetOrigList) > len(sheetItemList):
                # 基础表数据比简单表数据个数多
                sheetOrigList = lackListMerge(sheetOrigList, sheetItemList, arr)
            elif len(sheetOrigList) < len(sheetItemList):
                print('此情况还未出现')
        elif whichFunc==1 or sheetName == '受益人表':
            # 复杂表记录数据后，在之后流程中转门处理
            # print(arr,len(arr),sheetName,'复杂')
            diffDict[sheetName]={
                'compareList': arr,
                'sheetList': sheetItemList.copy()
            }
    if len(diffDict)>0:
        #复杂表有数据
        #受益人表并入被保人
        beneficiaryAndInsured = beneficiaryPushInsured(diffDict)
        #被保人表（+受益人表）并入险种表，并入过程扩充险种表
        riskList = balaPushRisk(diffDict,beneficiaryAndInsured)
        #险种表并入基本表，多险种处理成多条数据
        completeList = riskMaker(riskList,sheetOrigList)
        completeList = duplicateRemoval(completeList)
        # resultsToWorksheet(workbook.add_sheet('天安',cell_overwrite_ok=True),completeList)
    else:
        #复杂表无数据
        completeList = sheetOrigList

    resultDict = {'天安':completeList} #excel结果
    return resultDict

