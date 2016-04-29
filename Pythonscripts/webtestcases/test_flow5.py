#coding=utf-8
from selenium import webdriver
import time
import unittest
import ConfigParser
import  StringIO
import traceback
from classmethod import findStr
import csv
import os
from classmethod import login
cf = ConfigParser.ConfigParser()
cf.read(r"D:\Workspace\Pythonscripts\environment\env.conf")
host=cf.get('service','host')
method=cf.get('dir','method')
data=cf.get('dir','data')
#读取截图存放路径
shot_path=cf.get('shotpath','path')
#读取product_id
csvpaths=file(''+data+'product_id.csv', 'r') #读取 产品名 以及模式
product_id=csvpaths.readline()
print product_id

class Core_Enterprise(unittest.TestCase):
    (u"核心模块")
    @classmethod
    def setUpClass(cls):
        cls.browser = webdriver.Firefox()
        cls.browser.maximize_window()
    @classmethod
    def tearDownClass(cls):
        cls.browser.close()
        cls.browser.quit()
    def test(self):
        (u"新建方案")
        browser = self.browser
        try:

           login.operate_login(self,'operation_login.csv') #登陆
           time.sleep(2)
           browser.find_element_by_link_text(u"产品配置").click()
           time.sleep(2)
           browser.find_element_by_link_text(u"方案配置").click()
           time.sleep(2)
           browser.find_element_by_id('new-program').click()
           browser.find_element_by_id('product').click()
           time.sleep(2)
           browser.find_element_by_xpath("//select[@id='product']/option[@value="+product_id+"]").click()
           #browser.find_element_by_xpath("")













        except:
            fp = StringIO.StringIO()  # 创建内存文件对象
            traceback.print_exc(file=fp)
            message = fp.getvalue()
            index = findStr.findStr(message, "File", 2)
            message = message[0:index]
            #message = message + e.msg
            browser.get_screenshot_as_file(shot_path + browser.title + ".png")
            self.assertTrue(False, message)

