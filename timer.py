# 邮件反馈计时器
import threading
import time
import datetime

import pandas as pd

from email139 import send_emai
from handlecityorder import get_city_addres


def timer(operLogFileName):
    """
    判断地市有无反馈邮件，没有反馈发邮件催
    :param operLogFileName: 操作日志文件名
    :return:
    """

    # 判断五小时后有没有回复邮件
    time.sleep(1800)

    emailAdressDict = get_city_addres()
    logdata= pd.DataFrame(pd.read_excel(operLogFileName))
    notResponeCity = logdata[logdata.isnull().T.any()]['分公司'].values
    print(notResponeCity)
    # 给没有反馈的地市发送邮件
    endtime = datetime.datetime.strptime(str((datetime.datetime.now() + datetime.timedelta(days=+1)).date()) + '10:30','%Y-%m-%d%H:%M')
    #endtime = datetime.datetime.strptime(str((datetime.datetime.now()).date()) + '17:45', '%Y-%m-%d%H:%M')
    for city in notResponeCity:
        try:
            print(emailAdressDict .get(city)[0])
            #send_emai(emailAdressDict .get(city)[0], '您好：\n\n\b \b请及时处理春旺地址核查。\r\n  \b \b \b \b \b \b\b \b \b \b \b \b四川移动\r\n  \b \b \b \b \b \b\b \b \b \b \b \b网优中心机器人', '')
            #time.sleep(10)
        except:
            pass
    # 每两小时催没有反馈的地市，发送邮件

    while 1:
        time.sleep(7200)
        print(int(datetime.datetime.now().hour))
        n_time = datetime.datetime.now()
        if 8 <= int(datetime.datetime.now().hour) <= 18:
            print('在工作时间内')

            # 当时间大于截至时间时，结束循环

            logdata = pd.DataFrame(pd.read_excel(operLogFileName))
            notResponeCity = logdata[logdata.isnull().T.any()]['分公司'].values
            print(notResponeCity)
            for city in notResponeCity:
                try:
                    print('--')
                    #print(emailAdressDict.get(city)[0])
                    #send_emai(emailAdressDict.get(city)[0],
                              #'您好：\n\n\b \b请及时处理春旺地址核查。\r\n  \b \b \b \b \b \b\b \b \b \b \b \b四川移动\r\n  \b \b \b \b \b \b\b \b \b \b \b \b网优中心机器人', '')
                    #time.sleep(10)
                except:

                    pass
        #print(endtime-n_time)
        if n_time > endtime:
            print('等待计时器')
            break
    print('====等待结束===')
