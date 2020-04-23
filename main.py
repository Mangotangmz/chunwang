import datetime
import time

from email139 import receive_email, send_emai
from finalcityoder import FinalCityOrder
from handlecityorder import get_city_addres, check_cityorder
from handlesourceorder import cut_address, check_complaint_address_order, cutorderfile_tocity, connact_sourceorder_Sheet
from timer import timer
from util import searchFile
from wirtelog import OperationLog
from time import sleep, ctime
import threading



# 发送压缩包至各地市
def sendCityFiles(toSendFilesList):

    '''将压缩文件下发至各个地市'''
    # 获取压缩文件列表
    #zipList = searchFile('.zip', startPath='./files/cityfiles/zipfile')
    # 获取各个地市的邮箱地址
    emailAdressDict = get_city_addres()
    for key in emailAdressDict:
        print(key)
        """
        for zipfile in zipList:
            if key in zipfile:
                #print(emailAdressDict.get(key)[0] + ':' + zipfile.split('\\')[-1])
            #try:
                send_emai(emailAdressDict.get(key)[0], '网优机器人', zipfile)

                time.sleep(10)
                logmsg = {'cityname':key, 'time':datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),'circle':revolution}
                # 添加一条日志
                operationlog.addSendLog(logmsg)

            #except:
                    #pass  """
        text  ='您好，\n\b\b\b附件为本周需要查验地址的投诉工单，请查收，并按要求整改反馈至本邮箱。\n\b\b\b\b反馈截止日期为本周五上午10:00\n\n\b\b\b\b~~~~~~~~~~~~~~~~~~~~~~~~~\n四川移动\n网优机器人'
        for cityOrderFile in toSendFilesList:
            if key in cityOrderFile:
                #print(emailAdressDict.get(key)[0] + ':' + zipfile.split('\\')[-1])
            #try:
                standerCityFile ='./files/mappingFiles/地市区县标准表.xlsx'
                send_emai(emailAdressDict.get(key)[0], text, [cityOrderFile, standerCityFile])


                logmsg = {'cityname':key, 'time':datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),'circle':revolution}
                # 添加一条日志
                operationlog.addSendLog(logmsg)
                time.sleep(3)
            #except:
                    #pass

def recivecityFiles(startEmailNum):
    """收取地市邮件"""

    # 获取初始邮件数
    #startEmailNum = receive_email(0, 1)
    print(startEmailNum)
    # 获取有附件信息的邮件信息列表
    startTime = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    # 获取最新的邮件封数
    # 获取截至时间
    endtime = datetime.datetime.strptime(str((datetime.datetime.now() + datetime.timedelta(days=+1)).date()) + '10:30','%Y-%m-%d%H:%M')
    #endtime =datetime.datetime.strptime(str((datetime.datetime.now()).date()) + '17:45', '%Y-%m-%d%H:%M')
    while 1:
        n_time = datetime.datetime.now()

        # 只在8：00至18：00 时间段收取邮件
        if 8<= int(datetime.datetime.now().hour) <= 18:
            print("处于工作时间内")
            emailList = receive_email(startEmailNum, 0)[0]
            print(emailList)
            startEmailNum = receive_email(1, 0)[1]

            emailAdressDict = get_city_addres()

            # 处理邮件中的附件文件
            for emailItem in emailList:
                emailTime = emailItem.get('emailTime')
                emailAdr = emailItem.get('senderAdr')
                emailFile = emailItem.get('localFileName')
                # 对收取的附件进行核查,并获取核查后的结果
                try:
                    result = check_cityorder(emailFile)
                    print(result)
                    checkedCityName = result.get('cityName')
                    checkedfileName = result.get('fileName')
                    checkedResult = result.get('result')
                    logmsg = {'cityname':checkedCityName,'time':emailTime,'result':checkedResult }
                    operationlog.addReciveLog(logmsg)

                    """ 根据结果判断是否合格，合格则无需下发附件，不合格则重新发送邮件下发附件"""
                    print(checkedResult[0] == '是')
                    if checkedResult[0] == '是':
                        send_emai(emailAdressDict.get(checkedCityName)[0], '您好：\n核查内容合格', '')

                    else:
                        print(checkedfileName)
                        send_emai(emailAdressDict.get(checkedCityName)[0], '您好：\n附件核查内容不合格，详情见附件！\n\b\b\b\b\b\b\b\b网优中心机器人', checkedfileName)
                    # 保存地市核查过的文件至最终文件字典
                    finacityorder.cityDict[checkedCityName] = checkedfileName
                except:
                    print('附件格式有误')
                    send_emai(emailAdr, '您好：<br>附件格式有误，请检查收重新发送\n\b\b\b\b\b\b\b\b网优中心机器人', '')
            if n_time >= endtime:
                print('结束收发邮件')
                break
            time.sleep(30)




if __name__ == '__main__':
    """ 
    处理原始工单
    """

    sourecefileName = './files/sourceComplaintOrder/3.8-4.11移动业务详情.xlsx'

    addressMappingFileName = './files/mappingFiles/地市，区县规整映射表.xlsx'
    # 实例化一个处理日志
    fileName = connact_sourceorder_Sheet(sourecefileName)
    operationlogName = './files./operationlog/' + "全省春旺地址核查处理日志" + fileName.split('单')[-1]
    operationlog = OperationLog(operationlogName)
    operationlog.create()

    # 获取周期
    revolution = fileName.split('单')[-1].split('.x')[0]
    startEmailNum = receive_email(0, 1)

    # 实例化一个最终地市文件字典
    finacityorder = FinalCityOrder()

    # 将原始工单地址切割为临时地市，临时区县
    complaintOrderFileName = cut_address(fileName)

    # 核查原始工单地市和区县名，并规整

    CheckedcomplainOrderFileName = check_complaint_address_order(addressMappingFileName, complaintOrderFileName)

    # 将已规整的工单按地市切割并打包
    desFolder = './files/cityfiles/sendEmail'
    sendResultDict = cutorderfile_tocity(CheckedcomplainOrderFileName, desFolder)
    notSendFinalName = sendResultDict.get("notSendFinalName")
    toSendFilesList = sendResultDict.get("toSendFilesList")
    # 保存未下发工单至最终核查工单字典
    finacityorder.cityDict['未下发'] = notSendFinalName

    """下发所有压缩包至22个地市"""
    sendCityFiles(toSendFilesList)
    #recivecityFiles()

    """启动两个线程，一个收取邮件，另一个子线程监控地市反馈情况，并对为反馈地市发送邮件"""
    lock = threading.Lock()
    threads = []
    # 循环收取邮件
    t1 = threading.Thread(target=recivecityFiles,args=(startEmailNum,))
    threads.append(t1)
    # 循环检查是否有地市未反馈
    t2 = threading.Thread(target=timer, args=(operationlogName,))
    threads.append(t2)
    # 执行线程列表中的线程
    for t in threads:
        #t.setDaemon(True)
        t.start()

    for p in threads:
        p.join()
    finalComorderFileName = './files/finalorderfiles/全省春旺地址核查'+ revolution + '.xlsx'
    finacityorder.combineFinalOrder(finalComorderFileName)