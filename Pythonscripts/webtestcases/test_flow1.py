#coding=utf-8
from unittest.test import test_suite
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
import  time
import unittest
import HTMLTestRunner
import sys
import os
reload(sys)
sys.setdefaultencoding('utf8')
import csv
import ConfigParser
import mysql.connector
#dcf_user数据库配置
cf = ConfigParser.ConfigParser()
cf.read(r"D:\Workspace\Pythonscripts\environment\env.conf")
host=cf.get('service','host')
method=cf.get('dir','method')
data=cf.get('dir','data')
#读取数据库文件
USER=cf.get('dcf_user','user')
HOST=cf.get('dcf_user','host')
PASSWORD=cf.get('dcf_user','password')
PORT=cf.get('dcf_user','port')
DATABASE=cf.get('dcf_user','database')
#读取登录数据
csvfile = file(data+'\operation_login.csv', 'rb')
reader = csv.reader(csvfile)
for line in reader:
    username = line[0]
    password = line[1]
csvfile.close()
#读取核心客户注册信息
csvfile = file(data+'\core_enterprise_customer.csv', 'rb')
reader = csv.reader(csvfile)
for line in reader:
    enterprise_name = line[0].decode('utf-8')
    customer_name = line[1].decode('utf-8')
    customer_email=line[2].decode('utf-8')
    customer_phone=line[3].decode('utf-8')
#获取核心企业登录密码
csvfile = file(data+'\core_enterprise_password.csv', 'rb')
reader = csv.reader(csvfile)
for line in reader:
    enterprise_password = line[0]

class Core_Enterprise(unittest.TestCase):
    customer_url=''
    (u"核心模块")
    @classmethod
    def setUpClass(cls):
        cls.browser = webdriver.Firefox()
        cls.browser.maximize_window()
    @classmethod
    def tearDownClass(cls):
        cls.browser.close()
        cls.browser.quit()

    def test_1(self):
        (u"平台邀请注册")
        browser = self.browser
        browser.implicitly_wait(10)
        browser.get("http://"+host+".dcfservice.com/loginop.jsp")
        try:
            # 运营平台登录
            browser.find_element_by_id("j_user_name").send_keys(username)
            browser.find_element_by_id("j_password").send_keys(password)
            browser.find_element_by_id("reg-btn").click()
            time.sleep(2)
            #客户邀请
            browser.find_element_by_link_text("客户邀请").click()
            time.sleep(4)
            browser.find_element_by_id("inviteCustomer").click()
            time.sleep(1)
            #客户信息填写
            browser.find_element_by_id("customerFullName").send_keys(enterprise_name)
            browser.find_element_by_xpath(".//*[@id='inviteForm']/div[2]/div/div/div[1]/button[2]").click()
            time.sleep(2)
            browser.find_element_by_link_text(u"水利、环境和公共设施管理业").click()
            #browser.execute_script("arguments[0].scrollIntoView()",we)
            browser.find_element_by_xpath(".//*[@id='inviteForm']/div[3]/div/div/div[1]/button[2]").click()
            time.sleep(2)
            we=browser.find_element_by_css_selector("#province>li>a[value='310000']")
            browser.execute_script("arguments[0].scrollIntoView()",we)
            we.click()
            time.sleep(2)
            invitedUser=browser.find_element_by_id("invitedUser")
            invitedUser.clear()
            invitedUser.send_keys(customer_name)
            time.sleep(2)
            invitedEmail=browser.find_element_by_id("invitedEmail")
            invitedEmail.clear()
            invitedEmail.send_keys(customer_email)
            time.sleep(2)
            invitedMobile=browser.find_element_by_id("invitedMobile")
            invitedMobile.clear()
            invitedMobile.send_keys(customer_phone)
            time.sleep(2)
            #新建邀请并注册
            browser.find_element_by_css_selector(".btn.btn-danger.createInviteBtn").click()
            time.sleep(5)
            self.assertEqual(browser.find_element_by_id("invite-email-core").text,u"发送邀请","没有新建成功")
            Core_Enterprise.customer_url=browser.execute_script("return document.getElementById('inviteUrl-core').value")
            time.sleep(5)
        except NoSuchElementException:
            self.assertTrue(False,"元素未找到")
    def test_2(self):
        (u"核心企业注册")
        browser=self.browser
        browser.implicitly_wait(10)
        browser.get(Core_Enterprise.customer_url)
        time.sleep(2)
        browser.find_element_by_id("jiaru").click()
        time.sleep(2)
        try:
            #填写注册信息
            browser.find_element_by_id("inputPassword").send_keys(enterprise_password)
            time.sleep(1)
            browser.find_element_by_id("inputRePassword").send_keys(enterprise_password)
            time.sleep(1)
            browser.find_element_by_id("getDynamic").click()
            time.sleep(5)
            #获取验证码
            now_handle = browser.current_window_handle
            Dynamic_url="http://"+host+".dcfservice.com/v1/public/sms/get?cellphone="+customer_phone
            js_script='window.open('+'"'+Dynamic_url+'"'+')'
            browser.execute_script(js_script)
            time.sleep(2)
            all_handles=browser.window_handles
            for handle in all_handles:
                if handle != now_handle:
                    browser.switch_to_window(handle)
            Dynamic_code=browser.find_element_by_css_selector("html>body>pre").text
            Dynamic_code=Dynamic_code[1:7]
            browser.switch_to_window(now_handle)
            #填写验证码
            browser.find_element_by_id("validateCode").send_keys(Dynamic_code)
            time.sleep(1)
            browser.find_element_by_id("registerbtn").click()
            #等待5秒进入主页面后，关闭导航页面
            time.sleep(5)
            browser.find_element_by_css_selector(".aknowledge").click()
            time.sleep(5)
            browser.find_element_by_xpath(".//*[@id='zhongjin-banner']/div[1]").click()
            time.sleep(1)
            #获取登录的用户名
            login_name=browser.find_element_by_css_selector("#logoutDiv>a").text
            if login_name==customer_name:
                self.assertTrue(True,"客户注册成功")
            else:
                self.assertTrue(False,"客户注册失败")
        except NoSuchElementException:
            self.assertTrue(False,"元素未找到")
    def test_3(self):
        (u"核心企业认证")
        browser=self.browser
        browser.implicitly_wait(10)
        browser.get("http://"+host+".dcfservice.com/loginop.jsp")
        time.sleep(2)
        try:
            # 运营平台登录
            browser.find_element_by_id("j_user_name").send_keys(username)
            browser.find_element_by_id("j_password").send_keys(password)
            browser.find_element_by_id("reg-btn").click()
            time.sleep(2)
            #客户认证
            browser.find_element_by_link_text(u"客户管理")
            time.sleep(1)
            browser.find_element_by_link_text(u"客户认证")
            time.sleep(2)
            try:
                #数据库连接
                conn = mysql.connector.connect(host=HOST, user=USER, passwd=PASSWORD, db=DATABASE, port=PORT)
                # 创建游标
                cur = conn.cursor()
                # customername_id查询
                sql = 'select customer_id from user where customer_name="'+ customer_name + '"'
                cur.execute(sql)
                # 获取查询结果
                result_set = cur.fetchall()
                if result_set:
                    for row in result_set:
                        customername_id = row[0]
                # 关闭游标和连接
                cur.close()
                conn.close()
            except mysql.connector.Error,e:
                self.assertTrue(False,"数据库连接异常")
            customername_id = str(customername_id)
            path="//*[@id='table-account']/tbody/tr[@data-customerid="+'"'+customername_id+'"]/td[6]/div/a[2]'
            browser.find_element_by_xpath(path).click()
            browser.execute_script('''alert("还在继续编写")''')
        except NoSuchElementException:
            self.assertTrue(False,"元素未找到")


if __name__ =='__main__':
    testsuite=unittest.TestSuite()
    testsuite.addTest(Core_Enterprise("test_1"))
    testsuite.addTest(Core_Enterprise("test_2"))
    filename = "d:\\result.html"
    fp = file(filename, 'wb')
    runner = HTMLTestRunner.HTMLTestRunner(stream=fp, title='Result', description='Test_Report')
    runner.run(testsuite)
