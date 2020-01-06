import abc
import os
import re
from email.parser import Parser
from email.header import decode_header
from email.utils import parseaddr
from email.message import Message
import poplib
import datetime


# 邮件信息类
class EmailInfo(object):
    def __init__(self):
        self.date = None
        self.subject = None
        self.from_address = None
        self.from_name = None
        self.size = None
        self.attachments_name = []

    # 返回易阅读的文件大小字符串（两位小数），如 12345678 bytes返回'11.77MB'
    @staticmethod
    def bytes_to_readable(bytes_size: int):
        size_unit = [' Bytes', ' KB', ' MB', ' GB', ' TB', ' PB']
        unit_index = 0
        easy_read_size = bytes_size
        while easy_read_size >= 1024:
            easy_read_size /= 1024
            unit_index += 1
        if unit_index == 0:
            return str(easy_read_size) + size_unit[unit_index]
        else:
            return '{:.2f}'.format(easy_read_size) + size_unit[min(len(size_unit), unit_index)]

    def print_info(self):
        print('subject:', self.subject)
        print('from_address:', self.from_address)
        print('from_name:', self.from_name)
        print('date:', self.date)
        print('attachments:', self.attachments_name)
        print('total size:', self.bytes_to_readable(self.size))
        print('-----------------------------')

    def add_attachment_name(self, attachment_name):
        self.attachments_name.append(attachment_name)


# 附件储存类_基类
class Saver(metaclass=abc.ABCMeta):
    __SUBJECT_MAX_LENGTH = 51

    @abc.abstractmethod
    def __init__(self, root_path, file_name, file_data):
        self._root_path = root_path
        self._file_name = file_name
        self._file_data = file_data

    def _save_file(self, directory_path):
        # 储存文件，directory_path是绝对路径，不包含文件名
        if not os.path.exists(directory_path):
            os.makedirs(directory_path)
        self._file_name = Saver.file_name_check_and_update(directory_path, self._file_name)
        file = open(os.path.join(directory_path, self._file_name), 'wb')
        file.write(self._file_data)
        file.close()

    @staticmethod
    # 检查文件名，如果相同则自动递增编号。返回文件名，不包含路径。
    def file_name_check_and_update(path, file_name):
        file_number = 2
        exist_file_list = os.listdir(path)
        pure_name, extension = os.path.splitext(file_name)
        while file_name in exist_file_list:
            file_name = pure_name + '_' + str(file_number) + extension
            file_number += 1
        return file_name

    @staticmethod
    # 检查文件夹名词，去除非法字符并控制长度
    def normalize_directory_name(directory_name):
        normalized_name = re.sub('[*"/:?|<>\n]', '', directory_name, 0)
        normalized_name = normalized_name[0:min(Saver.__SUBJECT_MAX_LENGTH, len(normalized_name))].strip()
        return normalized_name

    @abc.abstractmethod
    def save(self):
        pass


# 模式0：所有附件存入一个文件夹
class MergeSaver(Saver):
    def __init__(self, root_path, file_name, file_data):
        super().__init__(root_path, file_name, file_data)

    def save(self):
        self._save_file(self._root_path)


# 模式1：每个邮箱地址一个文件夹
class AddressClassifySaver(Saver):
    def __init__(self, root_path, file_name, file_data, email_address):
        super().__init__(root_path, file_name, file_data)
        self._email_address = self.normalize_directory_name(email_address)

    def save(self):
        self._save_file(os.path.join(self._root_path, self._email_address))


# 模式2：每个邮件主题一个文件夹
class SubjectClassifySaver(Saver):
    def __init__(self, root_path, file_name, file_data, email_subject):
        super().__init__(root_path, file_name, file_data)
        self._email_subject = self.normalize_directory_name(email_subject)

    def save(self):
        self._save_file(os.path.join(self._root_path, self._email_subject))


# 模式3：每个发件人的每个邮件主题一个文件夹
class AddressSubjectClassifySaver(Saver):
    def __init__(self, root_path, file_name, file_data, email_address, email_subject):
        super().__init__(root_path, file_name, file_data)
        self._email_address = self.normalize_directory_name(email_address)
        self._email_subject = self.normalize_directory_name(email_subject)

    def save(self):
        self._save_file(os.path.join(self._root_path, self._email_address, self._email_subject))


# 模式4：每个发件人昵称一个文件夹
class AliasClassifySaver(Saver):
    def __init__(self, root_path, file_name, file_data, from_alias):
        super().__init__(root_path, file_name, file_data)
        self._from_alias = self.normalize_directory_name(from_alias)

    def save(self):
        self._save_file(os.path.join(self._root_path, self._from_alias))


# 储存器工厂
class SaverFactor:
    def __init__(self, mode: int):
        self.__mode = mode

    def __call__(self, root_path, file_name, file_data, email_info: EmailInfo):
        """
            保存模式    SAVE_MODE
        【0：所有附件存入一个文件夹】
        【1：每个邮箱地址一个文件夹】
        【2：每个邮件主题一个文件夹】
        【3：每个发件人的每个邮件主题一个文件夹】
        【4：每个发件人昵称一个文件夹】
        """
        if self.__mode == 0:
            return MergeSaver(root_path, file_name, file_data)
        elif self.__mode == 1:
            return AddressClassifySaver(root_path, file_name, file_data, email_info.from_address)
        elif self.__mode == 2:
            return SubjectClassifySaver(root_path, file_name, file_data, email_info.subject)
        elif self.__mode == 3:
            return AddressSubjectClassifySaver(root_path, file_name, file_data, email_info.from_address, email_info.subject)
        elif self.__mode == 4:
            return AliasClassifySaver(root_path, file_name, file_data, email_info.from_name)
        else:
            return None


# 邮件属性判断_基类
class EmailJudge:
    @abc.abstractmethod
    def judge(self):
        pass


# 日期判断
class DateJudge(EmailJudge):
    def __init__(self, date_begin, date_end,  time_zone, email_date):
        self.__date_begin = date_begin
        self.__date_end = date_end
        self.__time_zone = time_zone
        self.__email_date = email_date

    def judge(self):
        # Date格式'4 Jan 2020 11:59:25 +0800'
        date_mail = datetime.datetime.strptime(self.__email_date, '%d %b %Y %H:%M:%S %z')
        date_begin = datetime.datetime.strptime((self.__date_begin + self.__time_zone), '%Y-%m-%d %H:%M%z')
        date_end = datetime.datetime.strptime((self.__date_end + self.__time_zone), '%Y-%m-%d %H:%M%z')
        return date_begin < date_mail < date_end


# 邮件主题判断
class SubjectJudge(EmailJudge):
    def __init__(self, subject_include, email_subject):
        self.__subject_include = subject_include
        self.__email_subject = email_subject

    def judge(self):
        return self.__subject_include in self.__email_subject


# 邮件发件人地址判断
class AddressJudge(EmailJudge):
    def __init__(self, address_include, email_from_address):
        self.__address_include = address_include
        self.__email_from_address = email_from_address

    def judge(self):
        return self.__address_include in self.__email_from_address


# 邮件发件人姓名判断
class NameJudge(EmailJudge):
    def __init__(self, name_include, email_from_name):
        self.__name_include = name_include
        self.__email_from_name = email_from_name

    def judge(self):
        return self.__name_include in self.__email_from_name


# 邮件筛选器
class EmailFilter:
    def __init__(self):
        self.__judges = []

    def add_judge(self, judge: EmailJudge):
        self.__judges.append(judge)

    def judge_conditions(self):
        for condition_judge in self.__judges:
            if not condition_judge.judge():
                return False
        return True


# 批量邮件下载类
class BatchEmail:
    def __init__(self, pop3_server, email_address, email_password):
        self.__pop3_server = pop3_server
        self.__email_address = email_address
        self.__email_password = email_password
        self.__save_mode = 0        # 附件保存模式
        self.save_path = 'Email-Attachments'        # 附件保存位置

        # 筛选属性
        self.date_begin, self.date_end = '2020-1-1 00:00', '2020-1-4 20:00'     # 筛选属性：起止时间
        self.time_zone = '+0800'     # 筛选属性：时区
        self.from_address = ''     # 筛选属性：发件人地址
        self.from_name = ''     # 筛选属性：发件人姓名
        self.subject = ''     # 筛选属性：邮件主题

        self.__connection = None
        self.__saver_factor = None

    def set_save_mode(self, save_mode):
        self.__save_mode = save_mode
        self.__saver_factor = SaverFactor(self.__save_mode)

    def connect(self):
        # 连接POP3服务器(SSL):
        try:
            self.__connection = poplib.POP3_SSL(self.__pop3_server)
        except OSError as e:
            print('连接服务器失败，请检查服务器地址或网络连接。')
            self.close()
            return

        self.__connection.set_debuglevel(False)
        poplib._MAXLINE = 32768     # POP3数据单行最长长度，在有些邮件中，该长度会超出协议建议值，所以适当调高

        # 服务器欢迎文字:
        print(self.__connection.getwelcome().decode())

        # 登录:
        self.__connection.user(self.__email_address)
        try:
            self.__connection.pass_(self.__email_password)
        except Exception as e:
            print(e.args)
            print('登陆失败，请检查用户名/密码。并确保您的邮箱已开启POP3服务。')
            self.close()
            return

    def download_attachments(self):
        if self.__connection is None:
            return

        # 邮件数量和总大小:
        mail_quantity, mail_total_size = self.__connection.stat()
        print('邮件总数:', mail_quantity)
        print('邮件总大小:', EmailInfo.bytes_to_readable(mail_total_size), end='\n\n')

        # mail_list中是各邮件信息，格式['number octets'] (1 octet = 8 bits)
        response, mail_list, octets = self.__connection.list()
        error_count = 0

        # 倒序读取（从最新的开始）
        for mail_info in reversed(mail_list):
            mail_number = int(bytes.decode(mail_info).split()[0])  # 从mail_list解析邮件编号
            # mail_number = 590     # debug

            try:
                # TOP命令接收前n行，此处仅读取邮件属性，读部分数据可加快速度。TOP非所有服务器支持，若不支持请使用RETR。
                response, content_byte, octets = self.__connection.top(mail_number, 40)
                mail_header_end = content_byte.index(b'')
                mail_message = self.parse_mail_byte_content(content_byte[:mail_header_end])     # 第一个空行前之是头部信息 RFC822
                message_info = self.__get_email_info(mail_message)
            except Exception as e:
                print('邮件接收或解码失败，邮件编号：[', mail_number, ']  错误信息：', e)
                error_count += 1
                continue

            email_filter = EmailFilter()
            email_filter.add_judge(DateJudge(self.date_begin, self.date_end, self.time_zone, message_info.date))
            email_filter.add_judge(SubjectJudge(self.subject, message_info.subject))
            email_filter.add_judge(AddressJudge(self.from_address, message_info.from_address))
            email_filter.add_judge(NameJudge(self.from_name, message_info.from_name))

            if email_filter.judge_conditions():
                response, content_byte, message_info.size = self.__connection.retr(mail_number)     # 接收完整邮件
                mail_message = self.parse_mail_byte_content(content_byte)
                self.__save_email_attachments(mail_message, message_info)
                print('( %d / %d )【%s】已保存，下一封' % (len(mail_list) - mail_number + 1, len(mail_list), message_info.subject))
                message_info.print_info()
            else:
                print('( %d / %d )【%s】不符合筛选条件，下一封' % (len(mail_list) - mail_number + 1, len(mail_list), message_info.subject))
        print('处理完成')
        if error_count > 0:
            print('有 %d 个邮件发生错误，请手动检查' % error_count)

    def close(self):
        if self.__connection is not None:
            try:
                self.__connection.close()
            except OSError as e:
                print('断开时发生错误')
            self.__connection = None

    @staticmethod
    # 将邮件中的bytes数据转为字符串
    def decode_mail_info_str(content):
        result_content = []
        for value, charset in decode_header(content):
            if type(value) != str:
                if charset is None:
                    value = value.decode(errors='replace')
                elif charset.lower() in ['gbk', 'gb2312', 'gb18030']:
                    # 一些特殊符号标着gbk，但编码可能是gb18030中的。gb18030向下兼容gbk、gb2312，所以一律用gb18030。
                    value = value.decode(encoding='gb18030', errors='replace')
                else:
                    value = value.decode(charset, errors='replace')

            result_content.append(value)
        return ' '.join(result_content)

    @staticmethod
    # 把邮件内容解析为Message对象
    def parse_mail_byte_content(content_byte):
        # 极个别邮件中，同一封邮件存在多种编码，那么就不要join后整体解码，而是每一行单独解码。情况少见，暂时忽略。
        mail_content = b'\n'.join(content_byte)
        try:
            mail_content = mail_content.decode()
        except UnicodeDecodeError as e:
            mail_content = mail_content.decode(encoding='GB18030', errors='replace')   # GB18030兼容GB231、GBK

        return Parser().parsestr(mail_content)

    def __save_email_attachments(self, message: Message, email_info):
        for part in message.walk():
            file_name = part.get_filename()
            if file_name:
                file_name = self.decode_mail_info_str(file_name)
                email_info.add_attachment_name(file_name)
                data = part.get_payload(decode=True)
                self.__saver_factor(self.save_path, file_name, data, email_info).save()

    def __get_email_info(self, message: Message):
        email_info = EmailInfo()

        name, address = parseaddr(message.get('From'))
        email_info.from_address = address
        email_info.from_name = self.decode_mail_info_str(name)

        # Date格式'Sat, 4 Jan 2020 11:59:25 +0800', 也有可能是'4 Jul 2019 21:37:08 +0800'
        # 开头星期去除，部分数据末尾有附加信息，因此以首个冒号后+12截取
        date = message.get('Date')
        date_begin_index = 0
        for date_begin_index in range(len(date)):
            if '0' <= date[date_begin_index] <= '9':
                break
        email_info.date = date[date_begin_index:date.find(':') + 12]

        try:
            email_info.subject = self.decode_mail_info_str(message.get('Subject'))
        except TypeError as e:
            email_info.subject = '无主题'

        return email_info
