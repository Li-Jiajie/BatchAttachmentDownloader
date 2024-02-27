"""
https://github.com/Li-Jiajie/BatchAttachmentDownloader

BatchAttachmentDownloader   v1.4.0
邮件附件批量下载
Python 3开发，支持IMAP4与POP3协议

支持多种附件保存模式、筛选模式

使用场景：通过邮箱收作业、调查等，批量下载附件    等

2021.04.23
Jiajie Li
"""

import downloader

# ************************请设置以下参数************************

# 邮箱地址  （必填）
EMAIL_ADDRESS = '*****your email address*****'
# 邮箱密码  （必填）
EMAIL_PASSWORD = '*****your email password*****'

# 邮件协议  （必填，POP3或IMAP4）
EMAIL_PROTOCOL = 'POP3'
# 服务器地址(SSL)    （必填，请根据协议填入合适的地址）
SERVER_ADDRESS = 'pop.qq.com'

# 附件保存位置
SAVE_PATH = r'F:\Email-Attachments'
# 筛选起止时间    yyyy-MM-dd HH:mm:ss
DATE_BEGIN, DATE_END = '2020-10-20 00:00', '2020-11-5 18:00'  # 筛选起止时间（包含此时间）
# 时区 默认东八区北京时间，如需更改请按如下格式更改
TIME_ZONE = '+0800'
# 筛选包含此内容的邮件地址，''表示全部邮件地址
FROM_ADDRESS = ''
# 筛选包含此内容的发件人昵称，''表示全部发件人昵称
FROM_NAME = ''
# 筛选包含此内容的邮件主题，''表示全部邮件主题
SUBJECT = ''

# 重要说明：收件人名称和收件人地址目前只能二选一设置，同时设置将会失效
# 筛选包含此收件人地址，''表示全部邮件主题 xxxxx@xxxxx.com
TO_ADDRESS = ''
# 筛选包含此收件人名称，''表示全部邮件主题
TO_NAME = ''

"""
    保存模式    SAVE_MODE
【0：所有附件存入一个文件夹】
【1：每个邮箱地址一个文件夹】
【2：每个邮件主题一个文件夹】
【3：每个发件人的每个邮件主题一个文件夹】
【4：每个发件人昵称一个文件夹】
【5：每个邮件主题+日期前缀一个文件夹】
"""
SAVE_MODE = 1

# ************************请设置以上参数************************


if __name__ == '__main__':
    # 服务器连接与邮箱登录
    downloader = downloader.BatchEmail(EMAIL_PROTOCOL, SERVER_ADDRESS, EMAIL_ADDRESS, EMAIL_PASSWORD)

    # 选项设置
    downloader.set_save_mode(SAVE_MODE)
    downloader.save_path = SAVE_PATH
    downloader.date_begin = DATE_BEGIN
    downloader.date_end = DATE_END
    downloader.time_zone = TIME_ZONE
    downloader.from_name = FROM_NAME
    downloader.from_address = FROM_ADDRESS
    downloader.subject = SUBJECT
    downloader.to_address = TO_ADDRESS
    downloader.to_name = TO_NAME

    # 下载附件
    downloader.download_attachments()
    downloader.close()
