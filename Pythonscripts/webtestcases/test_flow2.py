# coding:utf-8
<<<<<<< HEAD
import random
from classmethod import getprofile
from classmethod import login
=======

>>>>>>> 6b0d8150ae4267a2864647f93ef9b5f65d480b17
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
import csv
import ConfigParser
import os
cf = ConfigParser.ConfigParser()
cf.read(r"D:\Workspace\Pythonscripts\environment\env.conf")
host=cf.get('service','host')
method=cf.get('dir','method')
data=cf.get('dir','data')
csvfile = file(data+r'\depart_login.csv', 'rb')
reader = csv.reader(csvfile)
for line in reader:
    username=line[0].decode('utf-8')
    password=line[1].decode('utf-8')

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
    u"机构注册验证"
    def setUp(self):
        self.browser=webdriver.Firefox()
        self.browser.maximize_window()

    def test_invite(self):
        u"平台邀请注册"
        browser=self.browser
        #admin账户登录
        login.operate_login(slef,operation_login.csv)
        time.sleep(3)
        #客户邀请
        browser.find_element_by_link_text(u'客户邀请').click()
        time.sleep(3)
        browser.find_element_by_id('inviteCustomer').click()
        time.sleep(2)
        browser.find_element_by_id('customerFullName').send_keys(u'平安保险')
        browser.find_element_by_xpath(".//*[@id='inviteForm']/div[2]/div/div/div[1]/button[2]").click()
        time.sleep(2)
        browser.find_element_by_link_text(u'农、林、牧、渔业').click()
        browser.find_element_by_xpath(".//*[@id='inviteForm']/div[3]/div/div/div[1]/button[2]").click()
        time.sleep(2)
        sh=browser.find_element_by_css_selector("#province>li>a[value='310000']")
        browser.execute_script("arguments[0].scrollIntoView()",sh)
        sh.click()
        browser.find_element_by_id('optionsRadios2').click()#选择机构



    # def tearDown(self):
    #     self.browser.close()