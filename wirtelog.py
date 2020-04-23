import pandas as pd
from datetime import datetime
import time

class OperationLog():
    def __init__(self, name):  # 构造函数带参数
        self.name = name

    def create(self):
        """
        创建一个操作日志文件
        :return:
        """
        writer = pd.ExcelWriter(self.name)

        data = pd.DataFrame(columns=('分公司','春旺周期',	'省网优下发时间',	'分公司反馈时间',	'是否合格','不合格工单数量'))
        data.to_excel(writer, index = None)
        writer.save()


    def addSendLog(self,msg):
        """
        添加省网优下发邮件的日志
        :param msg:{‘cityname’: ,'time': ,'emailadr':,}
        :return:
        """
        data = pd.read_excel(self.name)
        data = data.append({'分公司': msg.get('cityname'), '春旺周期': msg.get('circle'), '省网优下发时间': msg.get('time')}, ignore_index=True)
        writer = pd.ExcelWriter(self.name)
        data.to_excel(writer, index = None)
        writer.save()


    def addReciveLog(self,msg):
        """
        添加地市反馈核查文件，并记录省网优再次核查的结果的日志
        :param msg: [{cityname:,time:,result:[是否合格，不合格工单数]}]
        :return:
        """
        data = pd.read_excel(self.name)
        #data.append({'分公司': msg.cityname, '春旺周期': self.name.split('志')[-1], '省网优下发时间': msg.time}, ignore_index=True)
        print(data.loc[data['分公司'] == msg.get('cityname')])
        print(data.loc[data['分公司'] == msg.get('cityname')][-1:])
        data.loc[data['分公司'] == msg.get('cityname'),'分公司反馈时间'] = msg.get('time')
        data.loc[data['分公司'] == msg.get('cityname'),'是否合格'] = msg.get('result')[0]
        data.loc[data['分公司'] == msg.get('cityname'),'不合格工单数量'] = msg.get('result')[1]
        writer = pd.ExcelWriter(self.name)
        data.to_excel(writer, index=None)
        writer.save()

