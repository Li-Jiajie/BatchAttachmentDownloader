import imaplib
import poplib
import re


# IMAP4协议 邮件接收类
class ImapReceiver:
    def __init__(self, host: str, email_address: str, email_password: str):
        # 连接IMAP4服务器(SSL):
        try:
            self.__connection = imaplib.IMAP4_SSL(host)
        except OSError as e:
            print('连接服务器失败，请检查服务器地址或网络连接。')
            self.close()
            return

        # 登录:
        try:
            s = self.__connection.login(email_address, email_password)
        except Exception as e:
            print(e.args)
            print('登陆失败，请检查用户名/密码。并确保您的邮箱已开启IMAP服务。')
            self.close()
            return

    def get_email_status(self):
        response, data = self.__connection.status('INBOX', '(MESSAGES)')
        quantity = int(re.findall(r'\d+', data[0].decode())[0])
        return quantity, -1

    def get_mail_list(self, condition: dict = None):
        self.__connection.select()
        # TODO IMAP协议支持搜索，可以把筛选条件放到这里。不过国内大多邮箱服务器不支持搜索操作。
        # NOTE 有些邮箱SEARCH的结果不一定是有序的
        response, mail_list = self.__connection.search(None, '(ALL)')
        return [x.decode() for x in reversed(mail_list[0].split())]

    def get_mail_header_bytes(self, mail_number: str):
        response, data = self.__connection.fetch(mail_number, '(BODY[HEADER])')
        if data[0] is None:
            # 极少数邮件无法获取到内容，一般是系统发送的邮件
            raise ValueError('邮件解析失败')
        return data[0][1]

    def get_full_mail_bytes(self, mail_number: str):
        response, data = self.__connection.fetch(mail_number, '(RFC822)')
        size = int(re.findall(r'\d+', data[0][0].split()[2].decode())[0])
        return data[0][1], size

    def close(self):
        if self.__connection is not None:
            try:
                self.__connection.close()
            except OSError as e:
                print('断开时发生错误')
            self.__connection = None


# POP3协议 邮件接收类
class Pop3Receiver:
    def __init__(self, host: str, email_address: str, email_password: str):
        # 连接POP3服务器(SSL):
        try:
            self.__connection = poplib.POP3_SSL(host)
        except OSError as e:
            print('连接服务器失败，请检查服务器地址或网络连接。')
            self.close()
            return

        self.__connection.set_debuglevel(False)
        poplib._MAXLINE = 32768  # POP3数据单行最长长度，在有些邮件中，该长度会超出协议建议值，所以适当调高

        # 服务器欢迎文字:
        print(self.__connection.getwelcome().decode())

        # 登录:
        self.__connection.user(email_address)
        try:
            self.__connection.pass_(email_password)
        except Exception as e:
            print(e.args)
            print('登陆失败，请检查用户名/密码。并确保您的邮箱已开启POP3服务。')
            self.close()
            return

    def get_mail_list(self, condition: dict = None):
        # mail_list中是各邮件信息，格式['number octets'] (1 octet = 8 bits)
        response, mail_list, octets = self.__connection.list()
        return [x.split()[0].decode() for x in reversed(mail_list)]

    def get_email_status(self):
        return self.__connection.stat()

    def get_mail_header_bytes(self, mail_number: str):
        # TOP命令接收前n行，此处仅读取邮件属性，读部分数据可加快速度。TOP非所有服务器支持，若不支持请使用RETR。
        response, content_byte, octets = self.__connection.top(mail_number, 40)
        # 第一个空行前之是头部信息 RFC822
        try:
            mail_header_end = content_byte.index(b'')
        except ValueError as e:
            mail_header_end = len(content_byte)
        return self.__merge_bytes_list(content_byte[:mail_header_end])

    def get_full_mail_bytes(self, mail_number: str):
        response, content_byte, size = self.__connection.retr(mail_number)
        return self.__merge_bytes_list(content_byte), size

    @staticmethod
    def __merge_bytes_list(bytes_list):
        # 注：极个别邮件中，同一封邮件存在多种编码，那么就不要join后整体解码，而是每一行单独解码。情况少见，暂时忽略。
        return b'\n'.join(bytes_list)

    def close(self):
        if self.__connection is not None:
            try:
                self.__connection.close()
            except OSError as e:
                print('断开时发生错误')
            self.__connection = None
