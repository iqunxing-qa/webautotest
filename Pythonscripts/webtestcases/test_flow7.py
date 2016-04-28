#coding=utf-8
from unittest.test import test_suite
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from classmethod import getprofile
from classmethod import login
from classmethod import findStr
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
import  time
import unittest
import  random
import HTMLTestRunner
import sys
import os
import StringIO
import traceback
reload(sys)
sys.setdefaultencoding('utf8')
import csv
import ConfigParser
import mysql.connector
#读取全局配置文件
cf = ConfigParser.ConfigParser()
cf.read(r"D:\Workspace\Pythonscripts\environment\env.conf")
host=cf.get('service','host')
method=cf.get('dir','method')
data=cf.get('dir','data')
#获取Firefox的profile
propath=getprofile.get_profile()
profile=webdriver.FirefoxProfile(propath)
#读取截图存放路径
shot_path=cf.get('shotpath','path')
print shot_path
class core_contract(unittest.TestCase):
    (u"新建流水模块")
    @classmethod
    def setUpClass(cls):
        cls.browser=webdriver.Firefox(profile)
        cls.browser.maximize_window()
    def test_1_contract_allocation(self):
        (u"上传流水")
        browser=self.browser
        browser.implicitly_wait(10)
        try:
            # 登录运营平台
            login.corp_login(self, "core_enterprise_login.csv")
            time.sleep(1)
            #针对第一次登录要
            try:
                if browser.find_element_by_xpath("html/body/div[2]/div[1]").is_displayed():
                    browser.find_element_by_xpath("html/body/div[2]/div[1]").click()
            except NoSuchElementException,e:
                print ""
            browser.find_element_by_id("addDashBtn").click()#点击新建流水
            time.sleep(1)
            ###########################################
            #            第一次使用安装数字证书       #
            ###########################################
            browser.implicitly_wait(5)
            try:
                browser.find_element_by_id("getDy").click()#获取验证码
                time.sleep(1)
                now_handle=browser.current_window_handle#获取当前的handle
                Dynamic_url = "http://" + host + ".dcfservice.com/v1/public/sms/get?cellphone=18701762172"#获取验证码路径
                js_script = 'window.open(' + '"' + Dynamic_url + '"' + ')'
                browser.execute_script(js_script)
                time.sleep(2)
                all_handles = browser.window_handles
                for handle in all_handles:
                    if handle != now_handle:
                        browser.switch_to_window(handle)
                Dynamic_code = browser.find_element_by_css_selector("html>body>pre").text
                Dynamic_code = Dynamic_code[1:7]#截取字符串获取验证码
                print  Dynamic_code
                browser.switch_to_window(now_handle)#切换回以前handle
                browser.find_element_by_id("dyCode").send_keys(Dynamic_code)
                time.sleep(1)
                browser.find_element_by_id("validateDy").click()
                time.sleep(1)
                browser.find_element_by_id("installCfca").click()#点击立即安装
                browser.implicitly_wait(30)
                ###########未写完

            except NoSuchElementException,e:
                print u"该客户已经安装好安全控件"
             #######################################################################
            browser.implicitly_wait(10)#恢复隐式查找10S时间
            browser.find_element_by_xpath(".//*[@id='uploadArea']/div[1]/div[1]/span[1]").click()#点击上传文件
            time.sleep(2)
            os.system("D:\\workspace\\Pythonscripts\\classmethod\\upload_transaction_flow.exe")
            time.sleep(2)
            browser.find_element_by_id("submit-now").click()
            ############


































        except NoSuchElementException,e:
            fp = StringIO.StringIO()  # 创建内存文件对象
            traceback.print_exc(file=fp)
            message = fp.getvalue()
            index = findStr.findStr(message, "File", 2)
            message = message[0:index]
            message=message+e.msg
            browser.get_screenshot_as_file(shot_path + browser.title + ".png")
            self.assertTrue(False, message)
    # def test_2_contract_awaken(self):
    #     (u"产看流水是否新建成功")
    #     browser=self.browser
    #     browser.implicitly_wait(10)
    #     try:
    #         print ""
    #
    #
    #
    #
    #
    #
    #
    #
    #
    #
    #
    #
    #
    #
    #
    #
    #
    #     except NoSuchElementException,e:
    #         fp = StringIO.StringIO()  # 创建内存文件对象
    #         traceback.print_exc(file=fp)
    #         message = fp.getvalue()
    #         index = findStr.findStr(message, "File", 2)
    #         message = message[0:index]
    #         browser.get_screenshot_as_file("D：/" + browser.title + ".png")
    #         self.assertTrue(False, message)




































    # @classmethod
    # def tearDownClass(cls):
    #     cls.browser.close()
    #     cls.browser.quit()

