import datetime

import pandas as pd
from shutil import copyfile
from util import mkdir, zip_file_path,set_back_color
# import matplotlib.pyplot as plt
# 处理原始工单

def connact_sourceorder_Sheet(fileName):
    """将原始工单中的三个sheet合并，并提取有效值"""
    shengaoData = pd.DataFrame(pd.read_excel(fileName,sheet_name='申告'))
    shengaoData["来源"] = '申告'
    shuqiuData  = pd.DataFrame(pd.read_excel(fileName,sheet_name='诉求'))
    shuqiuData['来源'] = '诉求'
    xianchangData = pd.DataFrame(pd.read_excel(fileName,sheet_name='现场'))
    xianchangData['来源'] = '现场'

    shengaoData = shengaoData[['受理时间','工单编号','来源','投诉类型','故障地点']]
    xianchangData = xianchangData[['受理时间','工单编号','来源','投诉类型','故障地址']]
    shuqiuData = shuqiuData[['COMPLAINTTIME','工单编号','来源','投诉类型','故障地址']]
    new_col =['受理时间','工单编号','来源','投诉类型','故障地点']
    xianchangData.columns = new_col
    shuqiuData.columns = new_col
    sheetList = [shengaoData,shuqiuData,xianchangData]
    data = pd.concat(sheetList)

    finalName ='./files/sourceComplaintOrder/'+ '原始投诉工单('+fileName.split('/')[-1].split('移')[0] +').xlsx'
    print( finalName)
    writer = pd.ExcelWriter(finalName)
    data.to_excel(writer, index=None)
    writer.save()
    return finalName


def cut_address(fileName):

    """将地址数据按|切割出临时地市和临时区县"""
    # 读取工单数据
    filename = fileName.split('/')[-1]
    desfileName = './files/sourceComplaintOrder/' + '已切'+ filename
    data = pd.read_excel(fileName)
    data = pd.DataFrame(data)
    print(data.tail(5))
    data['故障地市（临时提取）'] =data.apply(lambda data: str(data['故障地点']).split("|")[0] if '|' in str(data['故障地点']) else None, axis=1)
    data['故障区县（临时提取）'] = data.apply(lambda data: str(data['故障地点']).split("|")[1] if '|' in str(data['故障地点']) else None, axis=1)
    print(data.tail(5))

    # 保存为新的文件
    writer = pd.ExcelWriter(desfileName)
    data.to_excel(writer, index=None)
    writer.save()
    writer.close()
    return desfileName

def match_rightcountry(previousData,rightData):
    """根据映射表匹配故障区县（省）,故障地市（省）"""
    #print(str(previousData['故障区县（临时提取）']) in str(rightData["故障区县（工单）"].values))
    a = str(previousData['故障地市（临时提取）']) + str(previousData['故障区县（临时提取）'])
    b = rightData["故障地市（工单）"] + rightData["故障区县（工单）"]
    if a in b.values:

        try:
            return[rightData.loc[b == a]["故障地市（规整后）"].iloc[0], rightData.loc[b == a]["故障区县（规整后）"].iloc[0]]
        except:
            return [None, None]
    else:
        return [previousData['故障地市（临时提取）'], previousData['故障区县（临时提取）']]



def mark_address_exception(citydata,regiondata,rightData ):

    """对异常地市进行说明"""

    if citydata == None and regiondata == None:
        print('地址为空')
        return '地址为空'
    if str(citydata) not in rightData['故障地市（规整后）'].values:

        return "省外地址"
    if (citydata != None and regiondata == None) or  regiondata == '用户不知晓或不提供':
        return '地址不详'

def check_complaint_address_order(addressMappingFileName,complaintAddressOrderFileName):

    """
    核查投诉工单地址地市名，区县名进行规范并添加异常说明
    """

    """
    addressMappingFileName:地市，区县规整映射表"
    complaintAddressOrderFileName:切割地址后的原始投诉地址表
    """
    # 读取地市区县标准表
    rightData = pd.read_excel(addressMappingFileName)
    rightData = pd.DataFrame(rightData)

    data = pd.read_excel(complaintAddressOrderFileName)
    data = pd.DataFrame(data)
    data = data .where(data .notnull(), None)
    #data['故障地市（省）'] = data['故障地市（临时提取）']
    data['故障地市（省）']= data.apply(lambda data: match_rightcountry(data, rightData)[0], axis=1)
    data['故障区县（省）'] = data.apply(lambda data: match_rightcountry(data, rightData)[1], axis=1)
    data['地址异常说明(省)'] =data.apply(lambda data: mark_address_exception(data['故障地市（省）'], data['故障区县（省）'],rightData), axis=1)
    print(data.columns.values.tolist())
    data = data[['受理时间', '工单编号', '来源', '投诉类型', '故障地点', '故障地市（省）', '故障区县（省）']]
    filename = complaintAddressOrderFileName.split('/')[-1]
    desfileName = './files/checkedComplainOrder/' +'已规整' + filename
    writer = pd.ExcelWriter(desfileName)
    data.to_excel(writer, index=None)
    writer.save()
    return desfileName


def cutorderfile_tocity(oderfileName,desFolder):

    """将文件按名字切分成22个地市文件"""
    """oderfileName:工单文件名，desFolder处理文件后存放路径"""
    timeName = oderfileName.split('单')[1].split('.x')[0]
    orderData = pd.read_excel(oderfileName)
    orderData = pd.DataFrame(orderData)
    # 将工单按照四川省地市名切割成各地市的文件并打包文件
    # 读取四川省地市名
    citystanderData= pd.DataFrame(pd.read_excel('./files/mappingFiles/地市区县标准表.xlsx'))
    citynameList= citystanderData.columns.values.tolist()[:-2]
    orderData['故障地市（省）'] = orderData['故障地市（省）'].apply(str)
    condition = ''
    toSendFilesList = []
    for cityName in citynameList:
        condition = condition + cityName + "|"
        cityData = orderData[orderData['故障地市（省）'].str.contains(cityName)].copy()
        cityData['故障地市（市）'] = None
        cityData['故障区县（市）'] = None
        cityData['原因说明（市）'] = None
        #cityData.plot(colors={'故障地市（市）': '#EAEAAE', '故障区县（市）': '#EAEAAE','原因说明（市）':'#EAEAAE'})
        cityData = cityData.style.applymap(set_back_color, subset=['故障地市（市）', '故障区县（市）','原因说明（市）'])

        #path = desFolder +'/'+ cityName + '春旺地址核查' + timeName
        path = desFolder +'/' + timeName
        mkdir(path)
        #copyfile('./files/mappingFiles/地市区县标准表.xlsx', path + '/' '地市区县标准表.xlsx')
        cityOrderName = path + '/' + cityName + timeName +'分公司春旺地址核查.xlsx'
        writer = pd.ExcelWriter(cityOrderName)
        cityData.to_excel(writer, index=None)
        toSendFilesList.append(cityOrderName)
        writer.save()
        #zip_file_path(path,  './files/cityfiles/zipfile/', cityName + '分公司春旺地址核查' + timeName + '.zip')        # 压缩每个地市的文件

    print(condition[:-1])

    # 保存未下发文件
    data =orderData[~orderData['故障地市（省）'].str.contains(condition[:-1])].copy()
    data['故障地市（市）'] = None
    data['故障区县（市）'] = None
    data['原因说明（市）'] = None
    # cityData.plot(colors={'故障地市（市）': '#EAEAAE', '故障区县（市）': '#EAEAAE','原因说明（市）':'#EAEAAE'})
    data = data.style.applymap(set_back_color, subset=['故障地市（市）', '故障区县（市）', '原因说明（市）'])
    notSendFinalName = './files/cityfiles/sendEmail/' + '未下发核查'+ timeName +'.xlsx'
    writer = pd.ExcelWriter(notSendFinalName)
    data.to_excel(writer, index=None)
    writer.save()
    reslut = {'toSendFilesList':toSendFilesList,'notSendFinalName':notSendFinalName}
    return reslut




#print(datetime.datetime.strptime(str((datetime.datetime.now()).date()) + '10:30','%Y-%m-%d%H:%M'))