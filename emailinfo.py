import abc
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
        print('date:', datetime.datetime.fromtimestamp(self.date))
        print('attachments:', self.attachments_name)
        print('total size:', self.bytes_to_readable(self.size))
        print('-----------------------------')

    def add_attachment_name(self, attachment_name):
        self.attachments_name.append(attachment_name)


# 邮件属性判断_基类
class EmailJudge:
    @abc.abstractmethod
    def judge(self):
        pass


# 日期判断
class DateJudge(EmailJudge):
    def __init__(self, date_begin, date_end, time_zone, email_date):
        self.__date_begin = date_begin
        self.__date_end = date_end
        self.__time_zone = time_zone
        self.__email_date = email_date

    def judge(self):
        # Date格式'4 Jan 2020 11:59:25 +0800'
        date_mail = self.__email_date
        date_begin = datetime.datetime.strptime((self.__date_begin + self.__time_zone), '%Y-%m-%d %H:%M%z').timestamp()
        date_end = datetime.datetime.strptime((self.__date_end + self.__time_zone), '%Y-%m-%d %H:%M%z').timestamp()
        return date_begin < date_mail < date_end

    @staticmethod
    # 比较是否比Target时间更早，用于结束邮件遍历的循环。包含时区。
    def is_earlier(email_time, target_time):
        return email_time < datetime.datetime.strptime(target_time, '%Y-%m-%d %H:%M%z').timestamp()


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
