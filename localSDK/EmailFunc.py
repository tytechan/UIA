import os
import time
import smtplib
import poplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.parser import Parser
from email.header import decode_header, Header
from email.utils import parseaddr
from config.ErrConfig import CNBMException, handleErr


class EmailUtil(object):

    def __init__(self, user, password,
                 pop_server=None,
                 pop_port=None,
                 smtp_server=None,
                 smtp_port=None):
        '''
        收邮件初始化,邮箱配置可使用代码中默认配置,也可传入其他邮箱信息,通过邮件标题来判读有效邮件
        :param user: 邮箱账号                                                               --<str>
        :param password: 邮箱密码                                                           --<str>
        :param pop_server: pop服务器地址,收邮件需传入                                       --<str>
        :param pop_port: pop端口,收邮件需传入                                               --<int>
        :param smtp_server: smtp服务器地址,发邮件需传入                                     --<str>
        :param smtp_port: smtp端口,发邮件需传入                                             --<int>
        '''
        self.sender = user
        self.password = password
        self.pop_server = pop_server
        self.pop_port = pop_port
        self.smtp_server = smtp_server
        self.smtp_port = smtp_port
        self.charset = "UTF-8"

    @CNBMException
    def get_result(self, before_index=1, email_title="", save_path=None):
        '''
        读取邮件,通过邮件标题判断有效邮件进行下载附件
        :param before_index: 从那一封邮件开始读                                      --<int>
        :param email_title: 要读取的邮件标题,部分匹配                                --<str>
        :param save_path: 如果需要下载附件,传入下载附件的保存路径,不传不下载         --<str>
        :return: 返回邮箱里最新的邮件数
        '''
        try:
            server = self.connect_server()
            resp, mails, octets = server.list()
            index = len(mails)
            du_email_count = index - int(before_index)
            for i in range(du_email_count, -1, -1):
                print("第%s封邮件" % (index - i))
                resp, lines, octets = server.retr(index - i)
                try:
                    msg_content = b'\r\n'.join(lines).decode("UTF-8")
                except Exception as e:
                    continue
                msg = Parser().parsestr(msg_content)
                subject = self.get_title(msg)
                print("邮件标题为:", subject)
                if email_title in subject:
                    for part in msg.walk():
                        file_name = part.get_filename()
                        charset = self.guess_charset(part)
                        if file_name:
                            data = part.get_payload(decode=True)
                            if save_path:
                                self.save_file(file_name, data, save_path)
                        else:
                            if charset:
                                content = part.get_payload(decode=True).decode(charset)
            server.quit()
            return index
        except Exception as e:
            handleErr(e)
            raise e

    @CNBMException
    def get_title(self, msg):
        '''
        获取邮件标题
        :param msg: email对象
        :return: 返回邮件标题
        '''
        try:
            subject = ""
            for header in ['From', 'Subject']:
                value = msg.get(header, '')
                if value:
                    if header == 'Subject':
                        subject = self.decode_str(value)
                    else:
                        hdr, addr = parseaddr(value)
                        fromaddr = u'%s' % (addr)
            return subject
        except Exception as e:
            handleErr(e)
            raise e

    @CNBMException
    def save_file(self, filename, data, save_path):
        '''
        保存附件
        :param filename: 附件名
        :param data: 附件二进制数据
        :param save_path: 保存路径
        :return:
        '''
        try:
            h = Header(filename)
            dh = decode_header(h)
            filename = dh[0][0]
            if dh[0][1]:
                filename = self.decode_str(str(filename, dh[0][1]))
            if not os.path.exists(save_path):
                os.makedirs(save_path)
            os.chdir(save_path)
            f = open(filename, 'wb')
            f.write(data)
            f.close()
            print("{}文件下载成功".format(filename))
        except Exception as e:
            handleErr(e)
            raise e

    @CNBMException
    def connect_server(self):
        '''连接到POP3服务器'''
        try:
            for i in range(3):
                try:
                    server = poplib.POP3_SSL(self.pop_server, port=self.pop_port)
                    server.user(self.sender)
                    server.pass_(self.password)
                    server.rset()
                    return server
                except Exception as e:
                    time.sleep(1)
        except Exception as e:
            handleErr(e)
            raise e

    @CNBMException
    def decode_str(self, s):
        '''字符解码'''
        try:
            value, charset = decode_header(s)[0]
            if charset:
                value = value.decode(charset)
            return value
        except Exception as e:
            handleErr(e)
            raise e

    @CNBMException
    def guess_charset(self, msg):
        '''获取邮件的字符编码，首先在message中寻找编码，如果没有，就在header的Content-Type中寻找'''
        try:
            charset = msg.get_charset()
            if charset is None:
                content_type = msg.get('Content-Type', '').lower()
                pos = content_type.find('charset=')
                if pos >= 0:
                    charset = content_type[pos + 8:].strip()
            return charset
        except Exception as e:
            handleErr(e)
            raise e

    @CNBMException
    def send_mail(self, subject, mainText, attachs=None, toReceiver=None, ccReciver=None):
        '''
        发邮件
        :param subject: 标题   --<str>
        :param mainText: 正文   --<str>
        :param attachs: 附件   --<list>
        :param toReceiver: 收件人 --<list>
        :param ccReciver: 抄送人 --<list>
        :return:是否发送成功  --- <bool>
        '''
        try:
            message = MIMEMultipart()
            message["From"] = self.sender
            message["To"] = ";".join(toReceiver)
            if ccReciver:
                receivers = toReceiver + ccReciver
                message["Cc"] = ";".join(ccReciver)
            else:
                receivers = toReceiver
            message["Subject"] = subject
            message.attach(MIMEText(mainText, "plain", self.charset))
            if attachs is not None and len(attachs) > 0:
                for filePath in attachs:
                    att = MIMEText(open(filePath, "rb").read(), "base64", self.charset)
                    att["Content-Type"] = "application/octet-stream"
                    tmpFileName = os.path.basename(filePath)
                    att.add_header("Content-Disposition", "attachment", filename=("GBK", "", tmpFileName))
                    message.attach(att)
            self.create_connect(receivers, message)
            return True
        except smtplib.SMTPException as e:
            print("Error: 无法发送邮件: %s" % e)
            return False
        except Exception as e2:
            print("Error: 无法发送邮件2: %s" % e2)
            return False

    @CNBMException
    def create_connect(self, receivers, message):
        '''
        建立smtp连接,将创建好的邮件实例发送
        :param receivers: 邮件接收人(收件人和抄送人)
        :param message: 邮件实例
        :return: 是否发送成功
        '''
        try:
            smtpObj = smtplib.SMTP(self.smtp_server, self.smtp_port)
            smtpObj.starttls()
            smtpObj.ehlo()
            smtpObj.login(self.sender, self.password)
            smtpObj.sendmail(self.sender, receivers, message.as_string())
            print("邮件发送成功")
            time.sleep(0.2)
            smtpObj.quit()
        except smtplib.SMTPException as e:
            print("Error: 无法发送邮件: %s" % e)
            handleErr(e)
            raise e



if __name__ == "__main__":
    senderInit = "yyxz_cnbm@cnbmtech.com"
    passwordInit = "rpa_123456"
    smtpServerInit = "smtp.partner.outlook.cn"
    popServerInit = "partner.outlook.cn"
    smtp_portInit = 587
    pop_portInit = 995
    read_index = 7166
    em = EmailUtil(senderInit, passwordInit, popServerInit, pop_portInit, smtpServerInit, smtp_portInit)
    # em.get_result(7200)
    em.send_mail("这是一个测试邮件", "testgfhjkl", [r"C:\Users\网讯达\Desktop\验证码.png"], toReceiver=["1162197278@qq.com"])
