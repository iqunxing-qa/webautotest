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
#读取核心客户注册信息
csvfile = file(data+'\core_random_customer.csv', 'rb')
reader = csv.reader(csvfile)
for line in reader:
    if reader.line_num==1:
        enterprise_name = line[0].decode('utf-8')
#读取合同配置银行卡卡号
csvfile = file(data+'\core_bank_number.csv', 'rb')
reader = csv.reader(csvfile)
for line in reader:
    if reader.line_num==1:
        financing_bank_number = line[0]
        loan_bank_number=line[0]
        received_bank_number=line[0]
        recovery_bank_number=line[0]
class core_contract(unittest.TestCase):
    (u"核心企业合同登记启用")
    @classmethod
    def setUpClass(cls):
        cls.browser=webdriver.Firefox(profile)
        cls.browser.maximize_window()
    def test_1_contract_allocation(self):
        (u"配置核心企业合同")
        browser=self.browser
        browser.implicitly_wait(10)
        try:
            # 登录运营平台
            login.operate_login(self, "operation_login.csv")
            time.sleep(2)
            browser.find_element_by_link_text(u"产品配置").click()
            time.sleep(1)
            browser.find_element_by_link_text(u"合同配置").click()
            time.sleep(1)
            browser.execute_script("arguments[0].click()",browser.find_element_by_xpath("html/body/div[1]/div[3]/div/div[1]/div[2]/a"))
            time.sleep(3)
            #####################################################################
            # 登记合同信息                                                      #
            #####################################################################
            #合同信息
            Schema_name=u"天玺-1+N-旭景"
            Select(browser.find_element_by_id("solution_id")).select_by_visible_text(Schema_name)
            time.sleep(1)
            #填写客户名称
            ct_select=browser.find_element_by_css_selector(".select2-selection__arrow")
            browser.execute_script("arguments[0].scrollIntoView()",ct_select)
            time.sleep(1)
            ct_select.click()
            time.sleep(2)
            #在搜索框输入企业名称
            browser.find_element_by_css_selector(".select2-search__field").send_keys(enterprise_name)
            time.sleep(3)
            #按enter键输入
            browser.find_element_by_css_selector(".select2-search__field").send_keys(Keys.ENTER)
            time.sleep(1)
            #配置授信参数
            financing_title=browser.find_element_by_xpath("html/body/div[1]/div[4]/div/div/form/div[2]/div[1]")
            browser.execute_script("arguments[0].scrollIntoView()",financing_title)#配置授信参数区域可视
            time.sleep(1)
            browser.find_element_by_id("quota").send_keys("10")#配置额度
            time.sleep(1)
            browser.find_element_by_id("validity").clear()
            browser.find_element_by_id("validity").send_keys("12")#合同有效期限
            browser.find_element_by_id("loan-account-account-name").send_keys(enterprise_name)#融资账户名
            time.sleep(1)
            browser.find_element_by_id("loan-account-bank-account-no").send_keys(financing_bank_number)#融资银行卡卡号
            time.sleep(1)
            browser.find_element_by_id("loan-account-branch-bank").click()#选择融资银行
            time.sleep(1)
            #创建融资银行账户
            browser.find_element_by_xpath(".//*[@id='form']/div[1]/div[2]/span/span[1]/span/span[2]").click()#点击融资开户行下拉箭头
            time.sleep(1)
            browser.find_element_by_xpath("html/body/span[6]/span/span[1]/input").send_keys(u"中国光大银行")
            time.sleep(2)
            browser.find_element_by_xpath("html/body/span[6]/span/span[1]/input").send_keys(Keys.ENTER)#融资银行选择中国光大银行
            time.sleep(1)
            Select(browser.find_element_by_id("province")).select_by_value(u"上海市")#选择上海市
            time.sleep(1)
            Select(browser.find_element_by_id("city")).select_by_value(u"上海市")
            time.sleep(1)
            Select(browser.find_element_by_id("child-bank")).select_by_visible_text("中国光大银行上海浦东支行")#选择光大银行上海浦东支行
            time.sleep(1)
            browser.find_element_by_id("bank-selecting").click()#确定选择
            time.sleep(1)
            #配置标准收费规则
            charge_rule=browser.find_element_by_xpath("html/body/div[1]/div[4]/div/div/form/div[3]/div[1]")
            browser.execute_script("arguments[0].scrollIntoView()",charge_rule)#使收费规则区域可视
            time.sleep(1)
            browser.find_element_by_css_selector(".fee-group-add").click()#点击新增收费规则按钮
            time.sleep(1)
            Select(browser.find_element_by_id("period")).select_by_value("4")#期限区间选择全期限
            time.sleep(1)
            browser.find_element_by_css_selector('.radio-inlne>input[name="payer_id"][value="1"]').click()#付款方为买家
            time.sleep(1)
            Select(browser.find_element_by_id("charge_type")).select_by_value("1")#收费方式选择按笔固定比例
            time.sleep(1)
            browser.find_element_by_id("charge_type_value").clear()
            browser.find_element_by_id("charge_type_value").send_keys("10.00")#收费方式的比例为10.00
            time.sleep(1)
            browser.find_element_by_id("add-fee-tr").click()
            time.sleep(1)
            #配置托管方式

            #放款资金托管方式
            browser.execute_script("arguments[0].scrollIntoView()",browser.find_element_by_xpath("html/body/div[1]/div[4]/div/div/form/div[4]/div[2]/div[1]"))#使放款资金托管区域可见
            time.sleep(1)
            browser.find_element_by_id("loan_deposit_account_name").send_keys(enterprise_name)#填写放款资金账户名
            time.sleep(1)
            browser.find_element_by_id("loan_deposit_account_account_no").send_keys(loan_bank_number)#填写放款银行卡账号
            time.sleep(1)
            browser.find_element_by_id("loan_deposit_account_branch_bank").click()#选择放款银行
            time.sleep(1)
            #创建放款银行账户
            browser.find_element_by_xpath(".//*[@id='form']/div[1]/div[2]/span/span[1]/span/span[2]").click()#点击放款开户行下拉箭头
            browser.find_element_by_xpath("html/body/span[6]/span/span[1]/input").send_keys(u"中国光大银行")
            time.sleep(2)
            browser.find_element_by_xpath("html/body/span[6]/span/span[1]/input").send_keys(Keys.ENTER)#放款银行选择中国光大银行
            time.sleep(1)
            Select(browser.find_element_by_id("province")).select_by_value(u"上海市")  # 选择上海市
            time.sleep(1)
            Select(browser.find_element_by_id("city")).select_by_value(u"上海市")
            time.sleep(1)
            Select(browser.find_element_by_id("child-bank")).select_by_visible_text("中国光大银行上海浦东支行")  # 选择光大银行上海浦东支行
            time.sleep(1)
            browser.find_element_by_id("bank-selecting").click()
            time.sleep(1)
            #回款资金托管方式
            browser.execute_script("arguments[0].scrollIntoView()", browser.find_element_by_xpath("html/body/div[1]/div[4]/div/div/form/div[4]/div[2]/div[3]"))  # 使回款资金托管区域可见
            browser.find_element_by_id("received_payment_account_account_name").send_keys(enterprise_name)  # 填写回款资金账户名
            time.sleep(1)
            browser.find_element_by_id("received_payment_account_account_no").send_keys(received_bank_number)  # 填写回款银行卡账号
            time.sleep(1)
            browser.find_element_by_id("received_payment_account_branch_bank").click()  # 选择回款银行
            time.sleep(1)
            # 创建回款银行账户
            browser.find_element_by_xpath(".//*[@id='form']/div[1]/div[2]/span/span[1]/span/span[2]").click()  # 点击回款开户行下拉箭头
            browser.find_element_by_xpath("html/body/span[6]/span/span[1]/input").send_keys(u"中国光大银行")
            time.sleep(2)
            browser.find_element_by_xpath("html/body/span[6]/span/span[1]/input").send_keys(Keys.ENTER)  # 放款银行选择中国光大银行
            time.sleep(1)
            Select(browser.find_element_by_id("province")).select_by_value(u"上海市")  # 选择上海市
            time.sleep(1)
            Select(browser.find_element_by_id("city")).select_by_value(u"上海市")
            time.sleep(1)
            Select(browser.find_element_by_id("child-bank")).select_by_visible_text("中国光大银行上海浦东支行")  # 选择光大银行上海浦东支行
            time.sleep(1)
            browser.find_element_by_id("bank-selecting").click()
            time.sleep(1)

            # 回款资金托管方式
            browser.execute_script("arguments[0].scrollIntoView()", browser.find_element_by_xpath("html/body/div[1]/div[4]/div/div/form/div[4]/div[2]/div[7]"))  # 使回收本金资金托管区域可见
            browser.find_element_by_id("recovery_principal_account_account_name").send_keys(enterprise_name)  # 填写回收本金账户名
            time.sleep(1)
            browser.find_element_by_id("recovery_principal_account_account_no").send_keys(recovery_bank_number)  # 填写回收本金银行卡账号
            time.sleep(1)
            browser.find_element_by_id("recovery_principal_account_branch_bank").click()  # 选择回收本金银行
            time.sleep(1)
            # 创建回收本金银行账户
            browser.find_element_by_xpath(".//*[@id='form']/div[1]/div[2]/span/span[1]/span/span[2]").click()  # 点击回收本金开户行下拉箭头
            browser.find_element_by_xpath("html/body/span[6]/span/span[1]/input").send_keys(u"中国光大银行")
            time.sleep(2)
            browser.find_element_by_xpath("html/body/span[6]/span/span[1]/input").send_keys(Keys.ENTER)  # 回收本金银行选择中国光大银行
            time.sleep(1)
            Select(browser.find_element_by_id("province")).select_by_value(u"上海市")  # 选择上海市
            time.sleep(1)
            Select(browser.find_element_by_id("city")).select_by_value(u"上海市")
            time.sleep(1)
            Select(browser.find_element_by_id("child-bank")).select_by_value("64122")  # 选择光大银行上海浦东支行
            time.sleep(1)
            browser.find_element_by_id("bank-selecting").click()
            time.sleep(1)
            #最后点击新建合同
            browser.execute_script("document.documentElement.scrollTop=document.body.scrollHeight")#滑动滚动条至底部使新建合同按钮可视
            time.sleep(1)
            browser.find_element_by_id("registration").click()
        except NoSuchElementException,e:
            fp = StringIO.StringIO()  # 创建内存文件对象
            traceback.print_exc(file=fp)
            message = fp.getvalue()
            index = findStr.findStr(message, "File", 2)
            message = message[0:index]
            browser.get_screenshot_as_file("D：/" + browser.title + ".png")
            self.assertTrue(False, message)
    def test_2_contract_awaken(self):
        (u"核心合同启用")
        browser=self.browser
        browser.implicitly_wait(10)
        try:
            print ""


















        except NoSuchElementException,e:
            fp = StringIO.StringIO()  # 创建内存文件对象
            traceback.print_exc(file=fp)
            message = fp.getvalue()
            index = findStr.findStr(message, "File", 2)
            message = message[0:index]
            browser.get_screenshot_as_file("D：/" + browser.title + ".png")
            self.assertTrue(False, message)



































    #
    # @classmethod
    # def tearDownClass(cls):
    #     cls.browser.close()
    #     cls.browser.quit()

