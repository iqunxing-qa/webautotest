# coding:utf-8

from selenium import webdriver
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
# 引入ActionChains鼠标操作类
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
# 引入keys类操作
import time
import unittest
# import HTMLTestRunner
import ConfigParser
cf = ConfigParser.ConfigParser()
cf.read(r"E:\Workspace\Pythonscripts\environment\env.conf")
host=cf.get('service','host')

def edittextclear(self,text):
    self.driver.keyevent(123)
    for i in range(0,len(text)):
        self.driver.keyevent(67)

def verifycase(a,b):
    if a==b:
        print ("is pass.")
        return True
    else:
        print ("is fail.")
        return False

class department_register(unittest.TestCase):
    def setUp(self):
        self.browser=webdriver.Firefox()

    def test_invite(self):
        self.browser.get('http://'+host+'.dcfservice.com/loginop.jsp')

    def tearDown(self):
        self.browser.close()