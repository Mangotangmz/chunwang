from datetime import datetime
import email
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
import pandas as pd
from email import encoders, parser
from email.header import Header, decode_header
from email.mime.text import MIMEText
from email.utils import parseaddr, formataddr

import poplib
import smtplib

def _format_addr(s):
    name, addr = parseaddr(s)
    return formataddr((Header(name, 'utf-8').encode(), addr))

def send_emai(addres,msgs,filename):



    msg = MIMEMultipart()
    msg.attach(MIMEText(msgs, 'plain', 'utf-8'))
    #发送邮箱地址
    from_addr = 'wyzx_rpa_robot_a1@163.com'
    password = 'HPWDGVCWLYWKAUXW'
    #收件箱地址
    to_addr =addres
    #smtp服务器
    smtp_server = 'smtp.163.com'
    #发送邮箱地址
    msg['From'] =_format_addr('网优中心机器人 <%s>' % from_addr)
    #收件箱地址
    msg['To'] = to_addr
    #主题
    msg['Subject'] = '春旺地址核查'

    if filename != '':
        # 添加附件就是加上一个MIMEBase，从本地读取一个图片:
        if isinstance(filename, list):
            for itemName in filename:
                with open(itemName, 'rb') as f:
                    # 设置附件的MIME和文件名，这里是png类型:
                    mime = MIMEBase('xlsx', 'xlsx', filename=itemName.split('/')[-1])
                    # 加上必要的头信息:
                    mime.add_header('Content-Disposition', 'attachment', filename=itemName.split('/')[-1])
                    mime.add_header('Content-ID', '<0>')
                    mime.add_header('X-Attachment-Id', '0')
                    # 把附件的内容读进来:
                    mime.set_payload(f.read())
                    # 用Base64编码:
                    encoders.encode_base64(mime)
                    # 添加到MIMEMultipart:
                    msg.attach(mime)
        else:
            with open(filename, 'rb') as f:
            # 设置附件的MIME和文件名，这里是png类型:
                mime = MIMEBase('xlsx', 'xlsx', filename=filename.split('/')[-1])
                # 加上必要的头信息:
                mime.add_header('Content-Disposition', 'attachment', filename=filename.split('/')[-1])
                mime.add_header('Content-ID', '<0>')
                mime.add_header('X-Attachment-Id', '0')
                # 把附件的内容读进来:
                mime.set_payload(f.read())
                # 用Base64编码:
                encoders.encode_base64(mime)

                # 添加到MIMEMultipart:
                msg.attach(mime)
    import smtplib
    server = smtplib.SMTP_SSL(smtp_server,465)
    server.set_debuglevel(1)
    print(from_addr)
    print(password)
    server.login(from_addr,password)
    server.sendmail(from_addr,[to_addr],msg.as_string())
    server.quit()




def guess_charset(msg):
    charset = msg.get_charset()
    if charset is None:
        content_type = msg.get('Content-Type', '').lower()
        pos = content_type.find('charset=')
        if pos >= 0:
            charset = content_type[pos + 8:].strip()
    return charset

def decode_str(s):
    value, charset = decode_header(s)[0]
    if charset:
        value = value.decode(charset)
    return value


#send_emai('18280207065@139.com','test','D:\RPArobotData\阿坝.csv')
def decode_mime_words(s):
    return u''.join(
        word.decode(encoding or 'utf8') if isinstance(word, bytes) else word
        for word, encoding in email.header.decode_header(s))

def decodeMsgHeader(header):
    """
    解码头文件
    :param header: 需解码的内容
    :return:
    """
    value, charset = decode_header(header)[0]
    if charset:
        value = value.decode(charset)
    return value
def receive_email(emai_number,a = 0):
    """

    :param 上一次邮件数
    :param a: 当为1是获取当前邮件数
    :return: [邮件信息列表，当前邮件数]
    """
    email = '18280207065@139.com'
    password = 'auto666666'
    email = 'wyzx_rpa_robot_a1@163.com'
    password = 'HPWDGVCWLYWKAUXW'

    pop3_server = 'pop.163.com'
    #pop3_server = 'pop.139.com'

    # 连接到POP3服务器:
    server = poplib.POP3_SSL(pop3_server, 995)

    # 可以打开或关闭调试信息:
    server.set_debuglevel(1)
    # 可选:打印POP3服务器的欢迎文字:
    print(server.getwelcome().decode('utf-8'))

    # 身份认证:
    server.user(email)
    server.pass_(password)

    # 获取当前邮件数
    resp, mails, octets = server.list()

    index = len(mails)
    print(index)
    if a == 1:
        return index
    # 保存最新的i封邮件
    messages = []
    for i in range(emai_number, index):
        resp, message, octets = server.retr(i+1)
        print(i)
        messages.append(message)
        print(messages)
    #print(messages[1])
    messages = [b'\r\n'.join(mssg).decode('utf-8') for mssg in messages]
    messages = [parser.Parser().parsestr(mssg) for mssg in messages]


    i = 0
    emailList = []


    for message in messages:
        i = i + 1
        j = 0
        for part in message.walk():
            j = j + 1
            fileName = part.get_filename()
            contentType = part.get_content_type()
            if fileName:
                # 获取邮件收取时间
                emailTime = message['date']
                senderContent = message["From"]
                # parseaddr()函数返回的是一个元组(realname, emailAddress)
                senderRealName, senderAdr = parseaddr(senderContent)
                # 将加密的名称进行解码
                senderRealName = decodeMsgHeader(senderRealName)
                senderAdr = decodeMsgHeader(senderAdr)

                file_name = decode_mime_words(fileName)
                #print(file_name)
                data = part.get_payload(decode=True)
                localFileName = './files/cityfiles/receivedEmail/' +datetime.now().strftime('%Y%m%d%H%M')+ file_name
                f = open(localFileName, 'wb')
                # 组装邮件信息
                emailList.append({'emailTime':emailTime, 'senderAdr':senderAdr, 'localFileName':localFileName})
                f.write(data)        # 保存附件到本地
                f.close()
            elif contentType == 'text/plain' or contentType == 'text/html':
                data = part.get_payload(decode=True)
                #print(data)

    server.quit()

    return [emailList, index]  # 返回邮件列表信息


#receive_email()
#print(datetime.now().strftime('%Y%m%d%H%M'))