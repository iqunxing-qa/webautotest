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
    def setUp(self):
        self.browser=webdriver.Firefox()
        self.browser.maximize_window()

    def test_invite(self):
        browser=self.browser
        #admin账户登录
        browser.get('http://'+host+'.dcfservice.com/loginop.jsp')
        browser.find_element_by_id('j_user_name').send_keys(username)
        browser.find_element_by_id('j_password').send_keys(password)
        browser.find_element_by_id('reg-btn').click()
        time.sleep(3)
        #客户邀请
        browser.find_element_by_link_text(u'客户邀请').click()
        time.sleep(3)
        browser.find_element_by_id('inviteCustomer').click()
        time.sleep(2)
        browser.find_element_by_id('customerFullName').send_keys(u'平安保险')


    # def tearDown(self):
    #     self.browser.close()