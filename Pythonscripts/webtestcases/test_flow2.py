# coding:utf-8
import random
from classmethod import getprofile
from classmethod import login
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
#获取主要配置
cf = ConfigParser.ConfigParser()
cf.read(r"D:\Workspace\Pythonscripts\environment\env.conf")
host=cf.get('service','host')
method=cf.get('dir','method')
data=cf.get('dir','data')
#获取Firefox的profile
propath=getprofile.get_profile()
profile=webdriver.FirefoxProfile(propath)
#读取运营账户密码
csvfile = file(data+r'\operation_login.csv', 'rb')
reader = csv.reader(csvfile)
for line in reader:
    username=line[0].decode('utf-8')
    password=line[1].decode('utf-8')
#读取部门注册信息
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
        cls.browser=webdriver.Firefox(profile)
        cls.browser.maximize_window()

    def test_1_invite(self):
        u"平台邀请注册"
        browser=self.browser
        #admin账户登录
        login.operate_login(self,'operation_login.csv')
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

    def test_2_department_register(self):
        u"机构客户注册"
        browser=self.browser
        department_register.url=browser.execute_script("return document.getElementById('inviteUrl-core').value")
        time.sleep(3)
        browser.get(department_register.url)
        browser.implicitly_wait(3)
        time.sleep(2)
        browser.find_element_by_id('inputPassword').send_keys('iqunxing1234')
        browser.find_element_by_id('inputRePassword').send_keys('iqunxing1234')
        browser.find_element_by_id('getDynamic').click()
        time.sleep(3)
        now_handle = browser.current_window_handle
        #获取验证码
        vcode_url="http://"+host+'.dcfservice.com/v1/public/sms/get?cellphone='+depart_mobile
        js_script='window.open("'+vcode_url+'")'
        browser.execute_script(js_script)
        time.sleep(2)
        all_handles=browser.window_handles
        for handle in all_handles:
            if handle != now_handle:
                print"Switched window is %s" % handle  # 输出待选择的窗口句柄
                browser.switch_to_window(handle)
                time.sleep(3)
        Dynamic_code=browser.find_element_by_css_selector("html>body>pre").text
        Dynamic_code=Dynamic_code[1:7]
        print  Dynamic_code #取得验证码
        #填写验证码
        browser.switch_to_window(now_handle)
        browser.find_element_by_id('validateCode').send_keys(Dynamic_code)
        browser.find_element_by_id('registerbtn').click()
        time.sleep(5)











    # @classmethod
    # def tearDownClass(cls):
    #     cls.browser.close()