import abc
import os
import re

from emailinfo import EmailInfo


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
            return AddressSubjectClassifySaver(root_path, file_name, file_data, email_info.from_address,
                                               email_info.subject)
        elif self.__mode == 4:
            return AliasClassifySaver(root_path, file_name, file_data, email_info.from_name)
        else:
            return None
