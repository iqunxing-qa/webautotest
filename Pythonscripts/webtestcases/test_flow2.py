# coding:utf-8
import random
from classmethod import clipboard
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
csvfile = file(data+r'\operation_login.csv', 'rb')
reader = csv.reader(csvfile)
for line in reader:
    username=line[0].decode('utf-8')
    password=line[1].decode('utf-8')
csvfile = file(data+r'\depart_login.csv', 'rb')
reader = csv.reader(csvfile)
for line in reader:
    depart_name=line[0].decode('utf-8')
    depart_mail=line[1].decode('utf-8')
    depart_mobile=line[2].decode('utf-8')

class department_register(unittest.TestCase):
    u"机构注册验证"
    @classmethod
    def setUpClass(cls):
        cls.browser=webdriver.Firefox()
        cls.browser.maximize_window()

    def test_1_invite(self):
        u"平台邀请注册"
        browser=self.browser
        #admin账户登录
        browser.get('http://'+host+'.dcfservice.com/loginop.jsp')
        time.sleep(3)
        browser.find_element_by_id('j_user_name').send_keys(username)
        browser.find_element_by_id('j_password').send_keys(password)
        browser.find_element_by_id('reg-btn').click()
        time.sleep(3)
        #客户邀请
        browser.find_element_by_link_text(u'客户邀请').click()
        time.sleep(3)
        browser.find_element_by_id('inviteCustomer').click()
        time.sleep(3)
        browser.find_element_by_id('customerFullName').send_keys(u'平安保险'+str(random.randrange(1,100000)))
        browser.find_element_by_xpath(".//*[@id='inviteForm']/div[2]/div/div/div[1]/button[2]").click()
        time.sleep(3)
        browser.find_element_by_link_text(u'农、林、牧、渔业').click()
        browser.find_element_by_xpath(".//*[@id='inviteForm']/div[3]/div/div/div[1]/button[2]").click()
        time.sleep(3)
        sh=browser.find_element_by_css_selector("#province>li>a[value='310000']")
        browser.execute_script("arguments[0].scrollIntoView()",sh)
        sh.click()
        browser.find_element_by_id('optionsRadios2').click()#选择机构
        time.sleep(2)
        invitedUser=browser.find_element_by_id("invitedUser")
        invitedUser.clear()
        invitedUser.send_keys(depart_name)
        time.sleep(2)
        invitedEmail=browser.find_element_by_id("invitedEmail")
        invitedEmail.clear()
        invitedEmail.send_keys(depart_mail)
        time.sleep(2)
        invitedMobile=browser.find_element_by_id("invitedMobile")
        invitedMobile.clear()
        invitedMobile.send_keys(depart_mobile)
        time.sleep(3)
        browser.find_element_by_css_selector(".btn.btn-danger.createInviteBtn").click()
        time.sleep(8)
        #新建邀请
        department_register=browser.execute_script("return document.getElementById('inviteUrl-core').value")
        browser.get(department_register)
        time.sleep(3)
    def test_2_department_register(self):
        u"机构客户注册"
        browser=self.browser
        #browser.get(department_register.url)
        browser.implicitly_wait(3)
        browser.find_element_by_id('inputPassword').send_keys('iqunxing1234')
        browser.find_element_by_id('inputRePassword').send_keys('iqunxing1234')
        browser.find_element_by_id('getDynamic').click()
        time.sleep(3)
        now_handle = browser.current_window_handle
        #获取验证码
        vcode_url=host+'.dcfservice.com/v1/public/sms/get?cellphone='+depart_mobile
        js_script='window.open("'+vcode_url+'")'
        print js_script
        browser.execute_script('''window.open("t6.dcfservice.com/v1/public/sms/get?cellphone=18621982600")''')
        time.sleep(2)
        vcode=browser.find_element_by_css_selector("html>body>pre").text
        print vcode









    # @classmethod
    # def tearDownClass(cls):
    #     cls.browser.close()