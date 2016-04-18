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

    def test_platform_invitation(self):
        (u"平台邀请注册")
        browser = self.browser
        path=os.path.abspath('..')
        print path
        csvfile = file('D:\\longin2.csv', 'rb')
        reader = csv.reader(csvfile)
        for line in reader:
            username_value=line[0]
            password_value=line[1]
        csvfile.close()
        username_value=username_value.decode("utf-8")
        password_value=password_value.decode("utf-8")
        print username_value
        print  password_value

        browser.get("http://t6.dcfservice.com/loginop.jsp")
        try:
            # 运营平台登录
            browser.find_element_by_id("j_user_name").send_keys(username_value)
            browser.find_element_by_id("j_password").send_keys(password_value)
            browser.find_element_by_id("reg-btn").click()
            time.sleep(2)
            #客户邀请
            browser.find_element_by_link_text("客户邀请").click()
            time.sleep(4)
            browser.find_element_by_id("inviteCustomer").click()
            time.sleep(1)
            #客户信息填写
            browser.find_element_by_id("customerFullName").send_keys(u"世纪佳缘")
            browser.find_element_by_xpath(".//*[@id='inviteForm']/div[2]/div/div/div[1]/button[2]").click()
            time.sleep(2)
            we=browser.find_element_by_link_text("水利、环境和公共设施管理业")
            #browser.execute_script("arguments[0].scrollIntoView()",we)
            we.click()
            time.sleep(5)














        except NoSuchElementException:
            self.assertTrue(False,"元素未找到")

if __name__ =='__main__':
    testsuite=unittest.TestSuite()
    testsuite.addTest(Core_Enterprise("test_platform_invitation"))
    filename = "d:\\result.html"
    fp = file(filename, 'wb')
    runner = HTMLTestRunner.HTMLTestRunner(stream=fp, title='Result', description='Test_Report')
    runner.run(testsuite)
