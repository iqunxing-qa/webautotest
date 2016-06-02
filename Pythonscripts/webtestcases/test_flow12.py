#coding=utf-8
from unittest.test import test_suite
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from classmethod import getprofile
from classmethod import *
from classmethod import findStr
import  time
import unittest
import sys
import StringIO
import traceback
import csv
import  os
import win32com.client
reload(sys)
sys.setdefaultencoding('utf8')
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
# 读取loan_document_id.csv中的单据编号
csvfile = file(data + 'loan_document_id.csv', 'rb')
reader = csv.reader(csvfile)
for line in reader:
    loan_document_id = line[0].decode('utf-8')
# 读取core_enterprise_login.csv
csvfile = file(data + 'core_enterprise_login.csv', 'rb')
reader = csv.reader(csvfile)
for line in reader:
    core_customer = line[0].decode('utf-8')
#读取合同配置的银行信息
csvfile = file(data + 'core_bank_number.csv', 'rb')
reader = csv.reader(csvfile)
for line in reader:
    financing_bank_no = line[0].decode('utf-8')#融资银行卡卡号
    back_bank_no= line[1].decode('utf-8')  # 回款银行卡卡号

#读取截图存放路径
shot_path=cf.get('shotpath','path')
class core_contract(unittest.TestCase):
    (u"链属企业融资")
    @classmethod
    def setUpClass(cls):
        cls.browser=webdriver.Firefox(profile)
        cls.browser.maximize_window()
    def test_1_contract_allocation(self):
        (u"申请融资")
        browser=self.browser
        browser.implicitly_wait(10)
        try:
            login.corp_login(self,"chain_enterprise_customer.csv")
            time.sleep(1)
            #对于第一次登陆，要关闭引导页
            try:
                browser.find_element_by_xpath(".//*[@id='zhongjin-banner']/div[1]").click()
                time.sleep(1)
            except NoSuchElementException,e:
                print ""
            click_xpath='//*[@id="'+loan_document_id+'"]/td[9]/button'
            browser.find_element_by_xpath(click_xpath).click()#点击融资
            time.sleep(1)
            if (browser.find_element_by_xpath(".//*[@id='loan-account']/table/tbody/tr[2]/td[2]").text).replace(" ", "")!=financing_bank_no:  #验证融资账户
                self.assertFalse(True,"zhang hu  bu pi pei")
            if (browser.find_element_by_xpath(".//*[@id='return-account']/table/tbody/tr[2]/td[2]").text).replace(" ", "") != back_bank_no:  # 验证融资账户
                self.assertFalse(True, "zhang hu  bu pi pei")
            time.sleep(1)
            ####################
            #其他验证信息？
            ####################
            browser.find_element_by_id("financing-apply-now").click()#点击融资申请
            time.sleep(1)
            browser.find_element_by_xpath(".//*[@id='depositNotRequired']/div/div/div[1]/button").click()
            time.sleep(1)
            browser.refresh()
            #点击融资申请后，查看已融资模块
            xpath='.//*[@id="'+loan_document_id+'"]'
            if not browser.find_element_by_xpath(xpath).is_displayed():
                title_index = browser.title.find("-")
                title = browser.title[0:title_index]
                browser.get_screenshot_as_file(shot_path + title + ".png")  # 对错误增加截图
                self.assertFalse(True,"已点击申请融资的信息没有出现在已融入区域")
        except Exception, e:
            fp = StringIO.StringIO()  # 创建内存文件对象
            traceback.print_exc(file=fp)
            message = fp.getvalue()
            index_file = findStr.findStr(message, "File", 2)
            index_Exception = message.find("Message")
            print_message = message[0:index_file] + message[index_Exception:]
            time.sleep(1)
            title_index=browser.title.find("-")
            title=browser.title[0:title_index]
            # im = ImageGrab.grab()
            # im.save(shot_path +title + ".png")
            browser.get_screenshot_as_file(shot_path +title + ".png")
            self.assertTrue(False, print_message)
    def test_2_contract_awaken(self):
        (u"查询下载资产包，点击提交")
        browser=self.browser
        browser.implicitly_wait(10)
        try:
           login.operate_login(self,"operation_login.csv")
           time.sleep(1)
           browser.find_element_by_link_text("融资管理").click()
           time.sleep(1)
           browser.find_element_by_id("searchTxt").clear()
           browser.find_element_by_id("searchTxt").send_keys(core_customer)#在搜索框里搜索融资的核心客户名称
           time.sleep(1)
           browser.find_element_by_xpath("html/body/div[1]/div[3]/div/div[5]/table/tbody/tr[1]/td[11]/a[4]").click()#资产包提交
           time.sleep(1)
           # browser.find_element_by_id("waitForDownload").click()#点击下载资产包
           # time.sleep(10)
           # return os.path.join(os.path.expanduser("~"), 'Desktop')






















        except Exception, e:
            fp = StringIO.StringIO()  # 创建内存文件对象
            traceback.print_exc(file=fp)
            message = fp.getvalue()
            index_file = findStr.findStr(message, "File", 2)
            index_Exception = message.find("Message")
            print_message = message[0:index_file] + message[index_Exception:]
            time.sleep(1)
            title_index = browser.title.find("-")
            title = browser.title[0:title_index]
            browser.get_screenshot_as_file(shot_path + title + ".png")
            self.assertTrue(False, print_message)
    # @classmethod
    # def tearDownClass(cls):
    #     cls.browser.close()
    #     cls.browser.quit()

