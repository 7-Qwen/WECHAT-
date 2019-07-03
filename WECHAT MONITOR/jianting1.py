#Designer:7-Qwen
# -*- coding: utf-8 -*-
import itchat
from itchat.content import *
import os
import time
import json
import xlsxwriter
import xlrd
import xlwt
import sys
import pymongo
from pymongo import MongoClient



# 文件临时存储页
rec_tmp_dir = os.path.join(os.getcwd(), 'tmp/','')
print("当前工作目录 : %s" % os.getcwd())

# 存储数据的字典
rec_msg_dict = {}




def run():
    itchat.auto_login(hotReload=True)
    itchat.run()
    if not is_online(auto_login=True):
        return
    print('登录成功')

def is_online(auto_login=False):
        """
        判断是否还在线。
        :param auto_login: bool,当为 Ture 则自动重连(默认为 False)。
        :return: bool,当返回为 True 时，在线；False 已断开连接。
        """

        def _online():
            """
            通过获取好友信息，判断用户是否还在线。
            :return: bool,当返回为 True 时，在线；False 已断开连接。
            """
            try:
                if itchat.search_friends():
                    return True
            except IndexError:
                return False
            return True

        if _online(): return True  # 如果在线，则直接返回 True
        if not auto_login:  # 不自动登录，则直接返回 False
            print('微信已离线..')
            return False


# 好友信息监听
@itchat.msg_register([TEXT, PICTURE, RECORDING, ATTACHMENT, VIDEO], isFriendChat=True)
def handle_friend_msg(msg):
    fengefu = '\n'
    msg_jieshoufang_id = msg['FromUserName'][0:4]
    msg_fasongfang_id = msg['ToUserName'][0:4]
    msg_id = msg['MsgId']
    msg_from_user =msg['User']['RemarkName']
    # 收到信息的时间
    msg_time_rec = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
    msg_create_time = msg['CreateTime']
    msg_type = msg['Type']

    if msg['Type'] == 'Text':
        msg_content = msg['Content']
    elif msg['Type'] == 'Picture' \
            or msg['Type'] == 'Recording' \
            or msg['Type'] == 'Video' \
            or msg['Type'] == 'Attachment':
        #msg_content = r"" + msg['FileName']
        msg['Text'](rec_tmp_dir + msg['FileName'])

    #s = str(rec_msg_dict)
    #file_pathF = open('D:\\记录\\F.txt', 'w',encoding="utf-8")
    #file_pathF.writelines(s)
#   file_pathF.close()
#   lines = file_pathF.readline()
#   file_pathO = open('D:\\记录\\O.txt', 'w', encoding='utf-8')
#    for word in lines.split(',', lines.count(',')):
#        file_pathO.writelines(word)
#        file_pathO.write('\n')
#    shutil.copy('D:\\记录\\O.txt','D:\\记录\\提取结果.txt')
#    print("监听成功...")

    rec_msg_dict.update({
         msg_id :
         {
            '监听接收方ID': msg_fasongfang_id,
            '监听发送方ID': msg_jieshoufang_id,
            '聊天对象': msg_from_user,
            '时间': msg_time_rec,
            '消息': msg_content
         }

    })
    generate_excel(rec_msg_dict)
    print("监控成功...")
    print(msg)


# 生成excel文件
def generate_excel(expenses):
    workbook = xlsxwriter.Workbook('./监控.xlsx')
    worksheet = workbook.add_worksheet()

    # 设定格式，等号左边格式名称自定义，字典中格式为指定选项
    # bold：加粗，num_format:数字格式
    bold_format = workbook.add_format({'bold': True})
    # money_format = workbook.add_format({'num_format': '$#,##0'})
    # date_format = workbook.add_format({'num_format': 'mmmm d yyyy'})

    # 将二行二列设置宽度为15(从0开始)
    worksheet.set_column(1, 1, 15)

    # 用符号标记位置，例如：A列1行
    #worksheet.write('A1', 'id', bold_format)
    worksheet.write('B1', '监听接收方ID', bold_format)
    worksheet.write('C1', '监听发送方ID', bold_format)
    worksheet.write('D1', '聊天对象', bold_format)
    worksheet.write('E1', '时间', bold_format)
    worksheet.write('F1', '消息', bold_format)
    row = 1
    col = 0
    key_1 = []
    for key in expenses:
     for key_1 in expenses[key]:
      #for item in expenses[key][key_1]:
        #print(key)
        #print(key_1)
        #print(item)
        if key_1 == '监听接收方ID':
         worksheet.write_string(row, col + 1, expenses[key][key_1])
        if key_1 == '监听发送方ID':
         worksheet.write_string(row, col + 2, expenses[key][key_1])
        if key_1 == '聊天对象':
         worksheet.write_string(row, col + 3, expenses[key][key_1])
        if key_1 == '时间':
         worksheet.write_string(row, col + 4, expenses[key][key_1])
        if key_1 == '消息':
         worksheet.write_string(row, col + 5, expenses[key][key_1])
         row += 1
    workbook.close()

     # 使用write_string方法，指定数据格式写入数据
        #worksheet.write_number(row, col, str(item[]))

#         worksheet.write_number(row, col + 1, str(item['监听接收方ID']))
#         worksheet.write_string(row, col + 2, str(item['监听发送方ID']))
#         worksheet.write_string(row, col + 3, item['聊天对象'])
#         worksheet.write_datetime(row, col + 4, str(item['时间']))
#         worksheet.write_string(row, col + 5, item['消息'])
#         row += 1
#    workbook.close()
def mongo():
    # 连接数据库
    client = MongoClient('localhost', 27017)
    db = client.jiankong
    account = db.kim

    data = xlrd.open_workbook('./监控.xlsx')
    table = data.sheets()[0]
    # 读取excel第一行数据作为存入mongodb的字段名
    rowstag = table.row_values(0)
    nrows = table.nrows
    # ncols=table.ncols
    # print rows
    returnData = {}
    for i in range(1, nrows):
        # 将字段名和excel数据存储为字典形式，并转换为json格式
        returnData[i] = json.dumps(dict(zip(rowstag, table.row_values(i))))
        # 通过编解码还原数据
        returnData[i] = json.loads(returnData[i])
        # print returnData[i]
        account.insert(returnData[i])


if __name__ == '__main__':
    run()
    mongo()
    if not os.path.exists(rec_tmp_dir):
        os.mkdir(rec_tmp_dir)