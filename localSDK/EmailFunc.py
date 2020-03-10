#!python3
# -*- coding: utf-8 -*-

# import poplib
# import smtplib
# from email.mime.text import MIMEText
# from email.mime.multipart import MIMEMultipart
# from pathlib import Path
# from email.header import Header
# from email.parser import Parser
# from email.parser import BytesParser
# from email.header import decode_header
# from email.utils import parseaddr
#
#
# class EmailUtil(object):
#     # 初始化数据
#     def __init__(self, sender="", password="", stmp_server="smtp.partner.outlook.cn",
#                  stmp_port=587, pop_server="partner.outlook.cn", pop_port=995):
#         self.charset = "UTF-8"
#         self.sender = usr_pwd[2]
#         self.password = usr_pwd[3]
#         self.stmp_server = stmp_server
#         self.stmp_port = stmp_port
#         self.pop_server = pop_server
#         self.pop_port = pop_port
#
#
#     def send_mail(self, mainTitle, subject, mainText, attachs=None, toReceiver=None, ccReciver=None):
#         # mainTitle:主题
#         # subject:发件箱标题
#         # mainText:z主要内容
#         # attachs:附件数组
#         # toReceiver:收件人数据中
#         # ccReciver:抄送人数组
#         try:
#             receivers = []
#             # 创建一个带附件的实例
#             message = MIMEMultipart()
#             message["From"] = self.sender
#             message["To"] = ";".join(toReceiver)
#             if ccReciver is not None:
#                 receivers = toReceiver + ccReciver
#                 message["Cc"] = ";".join(ccReciver)
#             else:
#                 receivers = toReceiver
#             message["Subject"] = subject
#             print(message)
#             # 邮件正文内容
#             message.attach(MIMEText(mainText, "plain", self.charset))
#             # 构造附件，传送当前目录下文件
#             if attachs is not None and len(attachs) > 0:
#                 for filePath in attachs:
#                     att = MIMEText(open(filePath, "rb").read(), "base64", self.charset)
#                     att["Content-Type"] = "application/octet-stream"
#                     # att["Accept-Language"]="zh-CN"
#                     # att["Accept-Charset"]="ISO-8859-1,UTF-8,GBK"
#                     # 这里的filename可以任意写，写什么名字，邮件中显示什么名字
#                     tmpFileName = os.path.basename(filePath)
#                     att.add_header("Content-Disposition", "attachment", filename=("GBK", "", tmpFileName))
#                     message.attach(att)
#             smtpObj = smtplib.SMTP(self.stmp_server, self.stmp_port)
#             '''------------加密，安全传输----------------'''
#             smtpObj.starttls()  # 开启安全传输模式  -》 如果个人邮箱，未加密情况下，不需要这句
#             smtpObj.ehlo()  # 向Gamil发送SMTP "ehlo" 命令
#             # smtpObj.starttls()
#             smtpObj.login(self.sender, self.password)
#             smtpObj.sendmail(self.sender, receivers, message.as_string())
#             rpa.logger.warn("邮件发送成功")
#             sleep(0.5)
#             smtpObj.quit()
#             return True
#         except smtplib.SMTPException as e:
#             rpa.logger.warn("Error: 无法发送邮件1: %s" % e)
#             return False
#         except Exception as e2:
#             rpa.logger.warn("Error: 无法发送邮件2: %s" % e2)
#             return False
#
#
# class Email1(object):
#
#     # 初始化数据
#     def __init__(self):
#         pass
#
#     # 获取参数配置
#     def readMailPara(self):
#         # 邮箱账号
#         self.sender = usr_pwd[2]
#         # 邮箱密码
#         self.password = usr_pwd[3]
#         self.pop_server = "partner.outlook.cn"
#         self.pop_port = 995
#
#     # 收取邮件,,,
#     def getResult(self):
#         down_path = ""
#         self.readMailPara()
#         # 连接到POP3服务器:
#         server = poplib.POP3_SSL(self.pop_server, port=self.pop_port)
#         # server = poplib.POP3(self.pop_server, self.pop_port)
#         # 可以打开或关闭调试信息:
#         server.set_debuglevel(1)
#         # 可选:打印POP3服务器的欢迎文字:
#         # print(server.getwelcome().decode("utf-8"))
#         # 身份认证:
#         server.user(self.sender)
#         server.pass_(self.password)
#         server.rset()
#         print('Messages: %s. Size: %s' % server.stat())
#         resp, mails, octets = server.list()
#         index = len(mails)
#         du_email_count = index - int(before_index)
#         for i in range(du_email_count, -1, -1):  # 记录上次邮件数，用当前邮件数减去上次邮件数=本次要执行的邮件数
#             global send_email_list
#             send_email_list = []
#             print(index - i)
#             resp, lines, octets = server.retr(index - i)
#             try:
#                 msg_content = b'\r\n'.join(lines).decode("UTF-8")
#             except Exception as e:
#                 print("跳过")
#                 continue
#             msg = Parser().parsestr(msg_content)
#             subject = ""
#             # 获取邮件标题和发件人
#             for header in ['From', 'Subject']:
#                 value = msg.get(header, '')
#                 if value:
#                     if header == 'Subject':
#                         subject = self.decode_str(value)  # 邮件标题
#                     else:
#                         hdr, addr = parseaddr(value)
#                         # name = self.decode_str(hdr)
#                         # fromname = u'%s' % (name)  # 发件人名称
#                         fromaddr = u'%s' % (addr)  # 发件人邮箱
#             send_email_list.append(fromaddr)
#             print("收件人列表=", send_email_list)
#             # 判断有效邮件
#             # tmpSubject = '申请台账核对'
#             if "分销待建采购合同" in subject:
#                 email_download_path = base_path
#                 print(email_download_path)
#             else:
#                 print("无需要下载的附件")
#                 continue
#             for part in msg.walk():
#                 file_name = part.get_filename()
#                 # 不知为何excel附件格式为application/octet-stream
#                 if file_name:
#                     h = Header(file_name)
#                     # 对附件名称进行解码
#                     dh = decode_header(h)
#                     filename = dh[0][0]
#                     if dh[0][1]:
#                         # 将附件名称可读化
#                         filename = self.decode_str(str(filename, dh[0][1]))
#                         print(filename)
#                     data = part.get_payload(decode=True)
#                     if not os.path.exists(email_download_path):
#                         os.makedirs(email_download_path)
#                     # os.chdir(email_download_path)  # 切换到这个目录下，
#                     down_path = email_download_path + "\\" + filename
#                     print('---email_download_path+"\\"+filename', down_path)
#                     # fEx = open("%s.xlsx" % (filename), 'wb')
#                     # fEx = open(filename, 'wb')
#                     fEx = open(down_path, 'wb')
#                     fEx.write(data)
#                     fEx.close()
#                     print("文件下载成功")
#
#         server.quit()
#         Util.write_to_excel(os.path.join(base_path, "数据文件.xlsx"), "C2", index, sheet_name="账号密码")
#         if down_path:
#             return down_path

import smtplib
from email.mime.text import MIMEText
from email.header import Header

# class Email:

sender = 'from@runoob.com'
receivers = ['429240967@qq.com']  # 接收邮件，可设置为你的QQ邮箱或者其他邮箱

# 三个参数：第一个为文本内容，第二个 plain 设置文本格式，第三个 utf-8 设置编码
message = MIMEText('Python 邮件发送测试...', 'plain', 'utf-8')
message['From'] = Header("菜鸟教程", 'utf-8')  # 发送者
message['To'] = Header("测试", 'utf-8')  # 接收者

subject = 'Python SMTP 邮件测试'
message['Subject'] = Header(subject, 'utf-8')

try:
    smtpObj = smtplib.SMTP('localhost')
    smtpObj.sendmail(sender, receivers, message.as_string())
    print
    "邮件发送成功"
except smtplib.SMTPException:
    print
    "Error: 无法发送邮件"