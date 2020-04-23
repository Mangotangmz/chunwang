import pandas as pd



def check_cityorigon(cityData,rightData):
    """
    检查分公司核查的地市和区县是否合格
    :return:
    """

    print(cityData)
    a = cityData["故障地市（市）"]
    b = rightData["故障地市（规整后）"]

    c = cityData["故障区县（市）"]
    d = rightData["故障区县（规整后）"]


    if str(cityData['故障地市（市）']) + str(cityData['故障区县（市）']) != str(cityData['故障地市（省）']) + str(cityData['故障区县（省）']) :
        if  cityData['原因说明（市）'] == None:

            return '标注：未说明更正原因'
        else:
            if a not in b.values:
                return '标注：非标准地市'

            elif c not in d.values:

                return '标注：非标准区县'
            if a in b.values and c in d.values:
                return  '内容符合规范'

    else:
        return '内容符合规范'


def check_cityorder(cityOderFileName):
    """
    对分公司反馈的工单进行验证

    """
    rightData = pd.DataFrame(pd.read_excel('./files/mappingFiles/地市，区县规整映射表.xlsx'))
    cityData = pd.DataFrame(pd.read_excel(cityOderFileName))
    cityData = cityData.where(cityData.notnull(), None)
    # 核查分公司反馈的“故障地市（市）”、“故障区县（市）”分别与省网优处理的“故障地市（省）”、“故障区县（省）”是否一致，如一致则不进行操作

    cityData['省网优核查结果（省）'] = cityData.apply(lambda cityData: check_cityorigon(cityData, rightData), axis=1)
     # 统计是否合格和不合格数

    checkResult = pd.DataFrame([cityData['省网优核查结果（省）'].value_counts()])
    if '内容符合规范' in checkResult.columns.values.tolist():
        sumQualified = checkResult['内容符合规范'].values[0]
        if sumQualified == len(cityData):
            re = '是'
        else:
            re = '否'

    else:
        re = '否'
        sumQualified = 0
    # 不合格数
    print(len(cityData))
    sumNotQualified = len(cityData) - sumQualified
    # 得到地市名列表，根据文件名判断是哪一个地市的文件
    cityName = ''
    cityNameList = pd.DataFrame(pd.read_excel('./files/mappingFiles/地市区县标准表.xlsx')).columns.values.tolist()
    for cityname in cityNameList:
        if cityname in cityOderFileName:
            cityName = cityname
            break
    desFileName = './files/cityfiles/receivedEmail/checkedorder/' + '省网优已核查' + cityOderFileName.split('/')[-1]
    writer = pd.ExcelWriter(desFileName)
    cityData.to_excel(writer, index=None)
    writer.save()
    msg = {'cityName':cityName, 'fileName':desFileName, 'result':[re, sumNotQualified]}
    #print(msg)
    return msg


def get_city_addres():
    """
    读取分公司邮箱文件返回一个地市邮箱字典{'阿坝[..,..],...'}
    :return:
    """

    adressData = pd.DataFrame(pd.read_excel('./files/mappingFiles/分公司邮箱（接收春旺数据）.xlsx'))
    adressData = adressData.where(adressData.notnull(), None)
    dict_address = adressData.set_index('地市').T.to_dict('list')
    #print(dict_address.get('阿坝'))

    return dict_address

