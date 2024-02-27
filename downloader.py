import re
from email.header import decode_header
from email.message import Message
from email.parser import Parser
from email.utils import parseaddr

from emailinfo import *
from receiver import ImapReceiver, Pop3Receiver
from saver import SaverFactor


# 批量邮件下载类
class BatchEmail:
    def __init__(self, mode, email_server, email_address, email_password):
        self.__save_mode = 0  # 附件保存模式
        self.save_path = 'Email-Attachments'  # 附件保存位置

        # 筛选属性
        self.date_begin, self.date_end = '2020-1-1 00:00', '2020-1-4 20:00'  # 筛选属性：起止时间
        self.time_zone = '+0800'  # 筛选属性：时区
        self.from_address = ''  # 筛选属性：发件人地址
        self.from_name = ''  # 筛选属性：发件人姓名
        self.subject = ''  # 筛选属性：邮件主题
        self.to_address = ''  # 筛选属性：收件人地址
        self.to_name = ''  # 筛选属性：收件人姓名

        self.__saver_factor = None

        if mode.lower().find('pop') != -1:
            self.__receiver = Pop3Receiver(email_server, email_address, email_password)
        elif mode.lower().find('imap') != -1:
            self.__receiver = ImapReceiver(email_server, email_address, email_password)
        else:
            print('请选择邮件协议，POP3或IMAP。')
            return

    def set_save_mode(self, save_mode):
        self.__save_mode = save_mode
        self.__saver_factor = SaverFactor(self.__save_mode)

    def download_attachments(self):
        if self.__receiver is None:
            return

        # 邮件数量和总大小:
        mail_quantity, mail_total_size = self.__receiver.get_email_status()
        print('邮件总数:', mail_quantity)
        if mail_total_size > 0:
            print('邮件总大小:', EmailInfo.bytes_to_readable(mail_total_size), end='\n\n')

        # mail_list中是各邮件信息，格式['number octets'] (1 octet = 8 bits)
        mail_list = self.__receiver.get_mail_list()
        error_count = 0

        # 倒序读取（从最新的开始）
        for mail_number in mail_list:
            # mail_number = '2075'     # debug
            try:
                content_byte = self.__receiver.get_mail_header_bytes(mail_number)
                mail_message = self.parse_mail_byte_content(content_byte)
                message_info = self.__get_email_info(mail_message)
            except Exception as e:
                print('邮件接收或解码失败，邮件编号：[%s]  错误信息：%s' % (mail_number, e))
                error_count += 1
                continue

            email_filter = EmailFilter()
            email_filter.add_judge(DateJudge(self.date_begin, self.date_end, self.time_zone, message_info.date))
            email_filter.add_judge(SubjectJudge(self.subject, message_info.subject))
            email_filter.add_judge(AddressJudge(self.from_address, message_info.from_address))
            email_filter.add_judge(NameJudge(self.from_name, message_info.from_name))
            if self.to_address:
                email_filter.add_judge(RecipientAddressJudge(self.to_address, message_info.to_addresses))
            if self.to_name:
                email_filter.add_judge(RecipientNameJudge(self.to_name, message_info.to_names))

            # 超出设定的最早时间则结束循环
            if DateJudge.is_earlier(message_info.date, self.date_begin + self.time_zone):
                break

            if email_filter.judge_conditions():
                content_byte, message_info.size = self.__receiver.get_full_mail_bytes(mail_number)  # 接收完整邮件
                mail_message = self.parse_mail_byte_content(content_byte)
                file_number = self.__save_email_attachments(mail_message, message_info)

                print('( %d / %d )【%s】' % (
                    len(mail_list) - int(mail_number) + 1, len(mail_list), message_info.subject), end='')
                print('已保存，下一封') if file_number != 0 else print('无附件')
                message_info.print_info()
            else:
                print( datetime.datetime.fromtimestamp(message_info.date), '( %d / %d )【%s】不符合筛选条件，下一封' % (
                    len(mail_list) - int(mail_number) + 1, len(mail_list), message_info.subject))
        print('处理完成')
        if error_count > 0:
            print('有 %d 个邮件发生错误，请手动检查' % error_count)

    def close(self):
        self.__receiver.close()

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
        try:
            mail_content = content_byte.decode()
        except UnicodeDecodeError as e:
            mail_content = content_byte.decode(encoding='GB18030', errors='replace')  # GB18030兼容GB231、GBK

        return Parser().parsestr(mail_content)

    # 解析收件人地址名称
    def __parse_mail_reciver_info(self, to_address_list):
        to_names = list()
        to_addresses = list()
        if to_address_list[0].split(","):
            for address in to_address_list[0].split(","):
                name, email = parseaddr(address)
                decoded_string = self.decode_mail_info_str(name)
                if decoded_string:
                    to_names.append(decoded_string)
                if email:
                    to_addresses.append(email)
        return to_names, to_addresses
    
    # 附件解析与保存，返回附件数量
    def __save_email_attachments(self, message: Message, email_info):
        file_count = 0
        for part in message.walk():
            file_name = part.get_filename()
            if file_name:
                file_name = self.decode_mail_info_str(file_name)
                email_info.add_attachment_name(file_name)
                data = part.get_payload(decode=True)
                self.__saver_factor(self.save_path, file_name, data, email_info).save()
                file_count += 1
        return file_count

    def __get_email_info(self, message: Message):
        email_info = EmailInfo()

        try:
            email_info.subject = self.decode_mail_info_str(message.get('Subject'))
        except TypeError as e:
            email_info.subject = '无主题'

        name, address = parseaddr(message.get('From'))
        email_info.from_address = address
        email_info.from_name = self.decode_mail_info_str(name)
        email_info.to_names, email_info.to_addresses = self.__parse_mail_reciver_info(
            message.get_all("To")
        )

        date = message.get('Date')
        # 少数情况下无Data字段，尝试从其他字段中获取时间
        if date is None:
            if date is None:
                date = decode_time_from_received(message)
            if date is None:
                date = decode_time_from_x_qq_mid(message)
            # 如果还不能获取到Date则报错。极少数邮件信息头内没有时间信息，偶发于一些系统发送的邮件
            if date is None:
                raise ValueError('该邮件收件时间解析失败，邮件主题：【%s】' % email_info.subject)

        # Date格式统一，最终格式：'4 Jan 2020 11:59:25 +0800' (%d %b %Y %H:%M:%S %z)
        # 通常源数据格式为'Sat, 4 Jan 2020 11:59:25 +0800', 也有可能是'4 Jul 2019 21:37:08 +0800'
        # 开头星期去除，部分数据末尾有附加信息，因此以首个冒号后+12截取
        date_begin_index = 0
        for date_begin_index in range(len(date)):
            if '0' <= date[date_begin_index] <= '9':
                break
        date = date.replace('GMT', '+0000')  # 部分邮件用GMT表示
        date = date[date_begin_index:date.find(':') + 12]
        # 最后以时间戳保存
        email_info.date = datetime.datetime.strptime(date, '%d %b %Y %H:%M:%S %z').timestamp()

        return email_info


# 尝试从X-QQ-mid字段解析邮件时间
def decode_time_from_x_qq_mid(message: Message):
    field_str = message.get('X-QQ-mid')
    if field_str is None:
        return None

    # X-QQ-mid格式例如：newapiserver5t1618419145t10192，时间是长度为10的时间戳，头尾加上t来分割
    time_str = re.findall('t[0-9]{10}t', field_str)
    if len(time_str) == 0:
        return None
    time = datetime.datetime.fromtimestamp(float(time_str[0][1:-1]), datetime.timezone.utc)
    return time.strftime('%d %b %Y %H:%M:%S %z')


# 尝试从Received字段解析邮件时间
def decode_time_from_received(message: Message):
    field_str = message.get('Received')
    if field_str is None:
        return None
    return field_str[field_str.rfind(';') + 1:]
