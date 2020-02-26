import requests
import xlrd
import json
import unittest
import time
from dingtalkchatbot.chatbot import DingtalkChatbot
import os

class Path:
    project_path = os.path.dirname(os.path.realpath(__file__))
    data_path = os.path.join(project_path,"data.xlsx")

start=time.time()
col_url_id = 1
col_data_id = 3
book = xlrd.open_workbook(Path.data_path)
sheet = book.sheets()[0]
rows = sheet.nrows

class TestMethod(unittest.TestCase):

    def setUp(self):
        print("开始")

    def tearDown(self):
        print("结束")

    def test01(self):
        row = 1
        url = sheet.cell_value(row, col_url_id)
        para = sheet.cell_value(row, col_data_id)
        res = requests.post(url,data=json.loads(para)).json()
        if res['reason'] == "success":
            result = 1
        else:
            result = 0
        return result


    def test02(self):
        row = 2
        url = sheet.cell_value(row, col_url_id)
        para = sheet.cell_value(row, col_data_id)
        res = requests.post(url,data=json.loads(para)).json()
        if res['reason'] == "success":
            result = 1
        else:
            result = 0
        return result

    def test03(self):
        row = 3
        url = sheet.cell_value(row, col_url_id)
        para = sheet.cell_value(row, col_data_id)
        res = requests.get(url, params=json.loads(para)).json()
        if res['reason'] == "successed!":
            result = 1
        else:
            result = 0
        return result

    def test04(self):
        row = 4
        url = sheet.cell_value(row, col_url_id)
        para = sheet.cell_value(row, col_data_id)
        res = requests.get(url, params=json.loads(para)).json()
        if res['reason'] == "successed!":
            result = 1
        else:
            result = 0
        return result

a = TestMethod()
b = int(a.test01())
c = int(a.test02())
e = int(a.test03())
f = int(a.test04())

h = ["不通过","通过"]
end=time.time()
m = time.strftime("%Y-%m-%d",time.localtime(time.time()))

w = time.strftime("%H:%M:%S",time.localtime(time.time()))

webhook = 'https://oapi.dingtalk.com/robot/send?access_token=770b9692bbf6cd9aa45b064c841fc6a63e706695f9363f76a0dd39e9ff25a17e'

xiaoding = DingtalkChatbot(webhook)

xiaoding.send_text(msg="test1测试结果为:%s"
                       "\ntest2测试结果为:%s"
                       "\ntest3测试结果为:%s"
                       "\ntest4测试结果为:%s"
                       "\n通过率为%d %%"
                       "\n接口响应时间: %.2f 秒"
                       "\n日期:%s"
                       "\n时间:%s"%(h[b],h[c],h[e],h[f],(((b+c+e+f)/4)*100),(end-start),m,w))


