#coding=utf-8
from selenium import webdriver
import time
import unittest
import ConfigParser
from classmethod.getdate import get_date
import os
from classmethod import login
cf = ConfigParser.ConfigParser()
cf.read(r"D:\Workspace\Pythonscripts\environment\env.conf")
host=cf.get('service','host')
method=cf.get('dir','method')
data=cf.get('dir','data')

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
        # (u"新建方案")
        # browser = self.browser
        # login.operate_login(self,'operation_login.csv') #登陆
        # time.sleep(2)
        # browser.find_element_by_link_text(u"产品配置").click()
        # time.sleep(2)
        # browser.find_element_by_link_text(u"方案配置").click()
        # time.sleep(2)
        # browser.find_element_by_id('new-program').click()
        # browser.find_element_by_id('product').click()
        num=get_date.date
        print( num)
