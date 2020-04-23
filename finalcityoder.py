import pandas as pd


class FinalCityOrder():
    """
    最后各地市的反馈文件字典
    """
    def __init__(self):  # 构造函数带参数
        citystanderData = pd.DataFrame(pd.read_excel('./files/mappingFiles/地市区县标准表.xlsx'))
        citynameList = citystanderData.columns.values.tolist()
        self.cityDict ={}
        for city in citynameList:
            self.cityDict[city]= ''


    def combineFinalOrder(self,finalName):
        """

        :param finalName: 传入最后合并文件的名称
        :return:
        """
        cityList = []
        for key, value in self.cityDict.items():
            if value !='':
                citydata = pd.DataFrame(pd.read_excel(value))

                cityList.append(citydata)
        data = pd.concat(cityList)
        writer = pd.ExcelWriter(finalName)
        data.to_excel(writer, index=None)
        writer.save()


