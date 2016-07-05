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
import re
import win32con
import  random
import HTMLTestRunner
import sys
import os
import win32com.client
import win32com
import StringIO
import traceback
reload(sys)
sys.setdefaultencoding('utf8')
import csv
import  win32clipboard as w
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
xlApp = win32com.client.Dispatch('Excel.Application')  # 打开EXCEL
xlBook = xlApp.Workbooks.Open( r'D:\\Workspace\\Pythonscripts\\testdatas\\core_customer.xlsx')
xlSht = xlBook.Worksheets('Sheet1')
enterprise_name=xlSht.Cells(2, 1).Value
xlBook.Close(SaveChanges=1)
del xlApp
#读取方案名称
xlApp = win32com.client.Dispatch('Excel.Application')  # 打开EXCEL
xlBook = xlApp.Workbooks.Open( r'D:\\Workspace\\Pythonscripts\\testdatas\\product_configuration.xlsx')
xlSht = xlBook.Worksheets('Sheet3')
Schema_name=xlSht.Cells(2, 1).Value
xlBook.Close(SaveChanges=1)
del xlApp
#读取机构信息
xlApp = win32com.client.Dispatch('Excel.Application')  # 打开EXCEL
xlBook = xlApp.Workbooks.Open( r'D:\\Workspace\\Pythonscripts\\testdatas\\institution_data.xlsx')
xlSht = xlBook.Worksheets('Sheet1')
institution_name=xlSht.Cells(2, 1).Value
institution_bank_no=str(xlSht.Cells(2, 2).Value).replace(" ","")
#读取合同配置信息
xlApp = win32com.client.Dispatch('Excel.Application')  # 打开EXCEL
xlBook = xlApp.Workbooks.Open(r'D:\\Workspace\\Pythonscripts\\testdatas\\contract_information.xlsx')
xlSht = xlBook.Worksheets('Sheet1')
pattern = re.compile(r'\d*')
core_in_supplychain = xlSht.Cells(3, 1).Value  # 供应链买卖方
limit_monkey = re.search(pattern,str(xlSht.Cells(3, 2).Value)).group()  # 额度
recourse_or_not =str(xlSht.Cells(3, 3).Value)  # 有无追索
contract_Validityperiod = re.search(pattern,str(xlSht.Cells(3, 4).Value)).group() # 合同有效期
xlBook.Close(SaveChanges=1)
del xlApp
#读取合同保存路径
assert_path=cf.get('assert_package','assert_path')
#读取截图存放路径
shot_path=cf.get('shotpath','path')
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
            time.sleep(2)
            browser.find_element_by_link_text(u"合同配置").click()
            time.sleep(2)
            browser.execute_script("arguments[0].click()",browser.find_element_by_xpath("html/body/div[1]/div[3]/div/div[1]/div[2]/a"))
            time.sleep(3)
            #####################################################################
            # 登记合同信息                                                      #
            #####################################################################
            Select(browser.find_element_by_id("solution_id")).select_by_visible_text(Schema_name)#选择产品方案
            time.sleep(1)
            ct_select=browser.find_element_by_css_selector(".select2-selection__arrow")
            browser.execute_script("arguments[0].scrollIntoView()",ct_select)
            time.sleep(1)
            ct_select.click()
            time.sleep(2)
            browser.find_element_by_css_selector(".select2-search__field").send_keys(enterprise_name)#在搜索框输入企业名称
            time.sleep(3)
            browser.find_element_by_css_selector(".select2-search__field").send_keys(Keys.ENTER)# 按enter键输入
            time.sleep(1)
            browser.find_element_by_css_selector(".radio-inlne>input[value='0']").click()#对于N+1模式核心企业是供应链的买方
            time.sleep(1)
            ###############################
            #配置授信参数                #
            ###############################
            financing_title=browser.find_element_by_xpath("html/body/div[1]/div[4]/div/div/form/div[2]/div[1]")
            browser.execute_script("arguments[0].scrollIntoView()",financing_title)#配置授信参数区域可视
            time.sleep(2)
            print  limit_monkey
            if browser.find_element_by_id("quota").is_displayed():
                print "find it"
                browser.find_element_by_id("quota").send_keys(limit_monkey)#配置额度
            time.sleep(1)
            if u"有" in recourse_or_not:
                browser.find_element_by_xpath("html/body/div[1]/div[4]/div/div/form/div[2]/div[2]/div[2]/div/label[1]/input").click()#有追索
            else:
                browser.find_element_by_xpath("html/body/div[1]/div[4]/div/div/form/div[2]/div[2]/div[2]/div/label[2]/input").click()#无追索
            time.sleep(1)
            browser.find_element_by_id("validity").clear()
            browser.find_element_by_id("validity").send_keys(contract_Validityperiod)#合同有效期限
            # browser.find_element_by_id("loan-account-account-name").send_keys(enterprise_name)#融资账户名
            # time.sleep(1)
            # browser.find_element_by_id("loan-account-bank-account-no").send_keys(financing_bank_number)#融资银行卡卡号
            # time.sleep(1)
            # browser.find_element_by_id("loan-account-branch-bank").click()#选择融资银行
            # time.sleep(1)
            # #创建融资银行账户
            # browser.find_element_by_xpath(".//*[@id='form']/div[1]/div[2]/span/span[1]/span/span[2]").click()#点击融资开户行下拉箭头
            # time.sleep(1)
            # browser.find_element_by_xpath("html/body/span[6]/span/span[1]/input").send_keys(u"中国光大银行")
            # time.sleep(2)
            # browser.find_element_by_xpath("html/body/span[6]/span/span[1]/input").send_keys(Keys.ENTER)#融资银行选择中国光大银行
            # time.sleep(1)
            # Select(browser.find_element_by_id("province")).select_by_value(u"上海市")#选择上海市
            # time.sleep(1)
            # Select(browser.find_element_by_id("city")).select_by_value(u"上海市")
            # time.sleep(1)
            # Select(browser.find_element_by_id("child-bank")).select_by_visible_text("中国光大银行上海浦东支行")#选择光大银行上海浦东支行
            # time.sleep(1)
            # browser.find_element_by_id("bank-selecting").click()#确定选择
            # time.sleep(1)
            #############################
            #配置帐期内标准收费规则     #
            #############################
            xlApp = win32com.client.Dispatch('Excel.Application')  # 打开EXCEL
            xlBook = xlApp.Workbooks.Open(r'D:\\Workspace\\Pythonscripts\\testdatas\\contract_information.xlsx')
            xlSht = xlBook.Worksheets('Sheet2')#读取收费规则表单
            for i in range(3,xlSht.UsedRange.Rows.Count+1):
                term_interval = xlSht.Cells(i, 1).Value#期限区间
                fee_type = xlSht.Cells(i, 2).Value  #费用类型
                charge_party = xlSht.Cells(i, 3).Value  #收费方
                payer_party=xlSht.Cells(i, 4).Value  # 付款方
                charge_base=xlSht.Cells(i, 5).Value  # 收费基数
                charge_type=xlSht.Cells(i, 6).Value  # 收费方式
                amount_or_proportion=str(xlSht.Cells(i, 7).Value)  # 比例/金额
                charge_rule=browser.find_element_by_xpath("html/body/div[1]/div[4]/div/div/form/div[3]/div[1]")
                browser.execute_script("arguments[0].scrollIntoView()",charge_rule)#使收费规则区域可视
                time.sleep(1)
                browser.find_element_by_css_selector(".fee-group-add").click()#点击新增收费规则按钮
                time.sleep(1)
                Select(browser.find_element_by_id("period")).select_by_visible_text(term_interval)#期限区间
                time.sleep(1)
                if u"利息" in fee_type:
                    browser.find_element_by_css_selector('.radio-inlne>input[name="charge_category"][value="1"]').click()#费用类型为利息
                else:
                    browser.find_element_by_css_selector('.radio-inlne>input[name="charge_category"][value="0"]').click()#费用类型为手续费
                time.sleep(1)
                if u"机构" in charge_party:
                    browser.find_element_by_css_selector('.radio-inlne>input[name="charge_id"][value="2"]').click()#收费为机构
                elif u"平台" in charge_party:
                    browser.find_element_by_css_selector('.radio-inlne>input[name="charge_id"][value="3"]').click()#收费方为平台
                else:
                    browser.find_element_by_css_selector('.radio-inlne>input[name="charge_id"][value="0"]').click()#收费方为卖家
                if u"卖家" in payer_party:
                    browser.find_element_by_css_selector('.radio-inlne>input[name="payer_id"][value="0"]').click()#付款方为卖家
                else:
                    browser.find_element_by_css_selector('.radio-inlne>input[name="payer_id"][value="1"]').click()#付款方为买家
                time.sleep(1)
                Select(browser.find_element_by_id("payment_base")).select_by_visible_text(charge_base)#收费基数
                time.sleep(1)
                Select(browser.find_element_by_id("charge_type")).select_by_visible_text(charge_type)#收费方式
                time.sleep(1)
                browser.find_element_by_id("charge_type_value").clear()
                browser.find_element_by_id("charge_type_value").send_keys(amount_or_proportion)#收费方式的比例为10.00
                Select(browser.find_element_by_id("charge_point")).select_by_value("0")  # 费用放款时扣除
                time.sleep(1)
                browser.find_element_by_id("add-fee-tr").click()
                time.sleep(1)
            #配置托管方式
            ###############################
            #放款账户配置(机构)           #
            ###############################
            browser.execute_script("arguments[0].scrollIntoView()",browser.find_element_by_xpath("html/body/div[1]/div[4]/div/div/form/div[4]/div[2]/div[1]"))#使放款资金托管区域可见
            time.sleep(1)
            browser.find_element_by_id("loan_deposit_account_name").send_keys(institution_name)#填写放款资金账户名(机构名)
            time.sleep(1)
            browser.find_element_by_id("loan_deposit_account_account_no").send_keys(institution_bank_no)#填写放款银行卡账号(机构卡号)
            time.sleep(1)
            browser.find_element_by_id("loan_deposit_account_branch_bank").click()#选择放款银行
            time.sleep(1)
            browser.find_element_by_xpath(".//*[@id='form']/div[1]/div[2]/span/span[1]/span/span[2]").click()#点击放款开户行下拉箭头
            time.sleep(1)
            browser.find_element_by_css_selector(".select2-search__field").send_keys(u"中信银行")
            time.sleep(2)
            browser.find_element_by_css_selector(".select2-search__field").send_keys(Keys.ENTER)#放款银行选择中信银行
            time.sleep(1)
            Select(browser.find_element_by_id("province")).select_by_value(u"上海市")  # 选择上海市
            time.sleep(1)
            Select(browser.find_element_by_id("city")).select_by_value(u"上海市")
            time.sleep(1)
            Select(browser.find_element_by_id("child-bank")).select_by_visible_text(u"中信银行股份有限公司上海徐家汇支行") #选择中信银行股份有限公司上海徐家汇支行
            time.sleep(1)
            browser.find_element_by_id("bank-selecting").click()
            time.sleep(1)
            ##############################
            # 回款账户配置(机构)         #
            ##############################
            browser.execute_script("arguments[0].scrollIntoView()", browser.find_element_by_xpath("html/body/div[1]/div[4]/div/div/form/div[4]/div[2]/div[3]"))  # 使回款资金托管区域可见
            browser.find_element_by_id("received_payment_account_account_name").send_keys(institution_name)# 填写回款资金账户名(机构)
            time.sleep(1)
            browser.find_element_by_id("received_payment_account_account_no").send_keys(institution_bank_no) # 填写回款银行卡账号(机构)
            time.sleep(1)
            browser.find_element_by_id("received_payment_account_branch_bank").click() # 选择回款银行
            time.sleep(1)
            # 创建回款银行账户
            browser.find_element_by_xpath(".//*[@id='form']/div[1]/div[2]/span/span[1]/span/span[2]").click() #点击回款开户行下拉箭头
            browser.find_element_by_css_selector(".select2-search__field").send_keys(u"中信银行")
            time.sleep(2)
            browser.find_element_by_css_selector(".select2-search__field").send_keys(Keys.ENTER)  #放款银行选择中信银行
            time.sleep(1)
            Select(browser.find_element_by_id("province")).select_by_value(u"上海市")  # 选择上海市
            time.sleep(1)
            Select(browser.find_element_by_id("city")).select_by_value(u"上海市")
            time.sleep(1)
            Select(browser.find_element_by_id("child-bank")).select_by_visible_text(u"中信银行股份有限公司上海徐家汇支行")  # 选择中信银行股份有限公司上海徐家汇支行
            time.sleep(1)
            browser.find_element_by_id("bank-selecting").click()
            time.sleep(1)
            ############################
            #回收本金户配置(机构)      #
            ############################
            browser.execute_script("arguments[0].scrollIntoView()", browser.find_element_by_xpath("html/body/div[1]/div[4]/div/div/form/div[4]/div[2]/div[7]"))  # 使回收本金资金托管区域可见
            browser.find_element_by_id("recovery_principal_account_account_name").send_keys(institution_name)  # 填写回收本金账户名(机构)
            time.sleep(1)
            browser.find_element_by_id("recovery_principal_account_account_no").send_keys(institution_bank_no)  # 填写回收本金银行卡账号(机构)
            time.sleep(1)
            browser.find_element_by_id("recovery_principal_account_branch_bank").click()#选择回收本金银行
            time.sleep(1)
            # 创建回收本金银行账户
            browser.find_element_by_xpath(".//*[@id='form']/div[1]/div[2]/span/span[1]/span/span[2]").click()  # 点击回收本金开户行下拉箭头
            browser.find_element_by_css_selector(".select2-search__field").send_keys(u"中信银行")
            time.sleep(2)
            browser.find_element_by_css_selector(".select2-search__field").send_keys(Keys.ENTER)  # 回收本金银行选择中国光大银行
            time.sleep(1)
            Select(browser.find_element_by_id("province")).select_by_value(u"上海市")  # 选择上海市
            time.sleep(1)
            Select(browser.find_element_by_id("city")).select_by_value(u"上海市")
            time.sleep(1)
            Select(browser.find_element_by_id("child-bank")).select_by_visible_text(u"中信银行股份有限公司上海徐家汇支行")  # 选择中信银行股份有限公司上海徐家汇支行
            time.sleep(1)
            browser.find_element_by_id("bank-selecting").click()
            time.sleep(1)
            ############################
            #配置商务合同/协议
            ############################
            browser.find_element_by_xpath(".//span[@class='btn btn-primary fileinput-button']").click()#点击上传合同/协议图片
            time.sleep(2)
            # upload_file路径，上传图片
            upload_file = method + "\\upload.exe " + data + "test_picture.jpg"
            os.system(upload_file)
            time.sleep(2)
            browser.find_element_by_css_selector(".form-control.fileName").send_keys(u"新建协议")
            time.sleep(2)
            #最后点击新建合同
            browser.execute_script("document.documentElement.scrollTop=document.body.scrollHeight")#滑动滚动条至底部使新建合同按钮可视
            time.sleep(1)
            browser.find_element_by_id("registration").click()
            time.sleep(5)
            browser.refresh()
            ###############################
            # 查看合同是否新建成功        #
            ###############################
            time.sleep(2)
            browser.find_element_by_xpath("html/body/div[1]/div[3]/div/div[2]/form/div/div/span/span[1]/span/span[2]").click()#点击核心企业的下拉箭头
            time.sleep(2)
            browser.find_element_by_css_selector(".select2-search__field").send_keys(enterprise_name)#输入核心企业名称
            time.sleep(2)
            browser.find_element_by_css_selector(".select2-search__field").send_keys(Keys.ENTER)  # 按enter键输入
            time.sleep(2)
            if not browser.find_element_by_xpath(".//*[@id='results_list']/tbody/tr/td[2]").text==enterprise_name:
                self.assertFalse(True,u"新建合同没有出现在列表中")
            browser.find_element_by_xpath(".//*[@id='results_list']/tbody/tr/td[8]/a[1]").click()#点击启用
            time.sleep(2)
            browser.find_element_by_id("btn-enable").click()#点击启用合同
            time.sleep(2)
            #####################
            browser.find_element_by_xpath("html/body/div[1]/div[3]/div/div[2]/form/div/div/span/span[1]/span/span[2]").click()  # 点击核心企业的下拉箭头
            time.sleep(2)
            browser.find_element_by_css_selector(".select2-search__field").send_keys(enterprise_name)  # 输入核心企业名称
            time.sleep(2)
            browser.find_element_by_css_selector(".select2-search__field").send_keys(Keys.ENTER)  # 按enter键输入
            time.sleep(2)
            if not browser.find_element_by_xpath(".//*[@id='results_list']/tbody/tr/td[7]/span").text==u"正常":
                self.assertFalse(True,u"合同启动失败")#断言点击合同启动时，状态是否成功

        except Exception,e:
            fp = StringIO.StringIO()  # 创建内存文件对象
            traceback.print_exc(file=fp)
            message = fp.getvalue()
            index_file = findStr.findStr(message, "File", 2)
            index_Exception=message.find("Message")
            index_Stacktrace = message.find("Stacktrace:")
            print_message = message[0:index_file] + message[index_Exception:index_Stacktrace]
            time.sleep(1)
            title_index = browser.title.find("-")
            title = browser.title[0:title_index]
            browser.get_screenshot_as_file(shot_path + title + ".png")
            self.assertTrue(False, print_message)
    def test_2_contract_awaken(self):
        (u"登记链属企业合同")
        browser = self.browser
        browser.implicitly_wait(10)
        login.operate_login(self, "operation_login.csv")
        time.sleep(2)
        browser.find_element_by_link_text(u"产品配置").click()
        time.sleep(1)
        browser.find_element_by_link_text(u"合同配置").click()
        time.sleep(1)
        try:
            browser.find_element_by_xpath(
                "html/body/div[1]/div[3]/div/div[2]/form/div/div/span/span[1]/span/span[2]").click()  # 点击核心企业的下拉箭头
            time.sleep(2)
            browser.find_element_by_css_selector(".select2-search__field").send_keys(enterprise_name)  # 输入核心企业名称
            time.sleep(2)
            browser.find_element_by_css_selector(".select2-search__field").send_keys(Keys.ENTER)  # 按enter键输入
            time.sleep(2)
            browser.find_element_by_xpath(".//*[@id='results_list']/tbody/tr/td[8]/a[4]").click()#点击关联方额度
            time.sleep(2)
            browser.find_element_by_css_selector(".btn.btn-success.pull-right.register-sub").click()#登记关联方合同
            time.sleep(2)
            browser.find_element_by_xpath(".//*[@id='register_form']/div[1]/div[1]/div/span").click()#点击选择合同文件
            time.sleep(2)
            ##################################
            #修改链属合同                    #
            ##################################
            chain_customer=u"测试链属企业"+str(time.strftime("%m%d%H%M%S", time.localtime()))#随机生成链属企业名称
            limit_num=500000#额度大小
            finance_scale=80 #融资比例
            approved_period=30#核对账期
            repayment_period=7#还款期
            buyback_period=0#回购期
            accept_business_period=61#接受的商务账期
            open_account=1#开账日
            postpone_or_not="Y"#是否顺延
            xlApp = win32com.client.Dispatch('Excel.Application')  # 打开EXCEL
            xlBook = xlApp.Workbooks.Open(
                "D:\\Workspace\\Pythonscripts\\testdatas\\"+u"hetong.xlsx")  # 将D:\\1.xls改为要处理的excel文件路径
            xlSht = xlBook.Worksheets(u'额度账期参数')
            xlSht.Cells(2, 1).Value =enterprise_name  # 修改合同核心企业名称
            xlSht.Cells(2, 2).Value =chain_customer  # 修改关联方客户名称
            xlSht.Cells(2, 3).Value =limit_num  # 修改合同核心企业名称
            xlSht.Cells(2, 4).Value =finance_scale  # 修改融资比例
            xlSht.Cells(2, 5).Value = approved_period  # 修改核对账期
            xlSht.Cells(2, 6).Value = repayment_period # 修改还款期
            xlSht.Cells(2, 7).Value =buyback_period  #回购期
            xlSht.Cells(2, 8).Value = accept_business_period  # 接受的商务账期
            xlSht.Cells(2, 11).Value =open_account# 修改开账日
            xlSht.Cells(2, 13).Value =postpone_or_not  # 修改开账日
            xlBook.Close(SaveChanges=1)  # 完成 关闭保存文件
            del xlApp
            os.system( method + "\\upload.exe " + data +u"hetong.xlsx")#上传合同模板
            time.sleep(3)
            browser.find_element_by_id("regis_contract_btn").click()#确认按钮

            ######################################################
            # 链属登记合同后，将随机生成的链属写入chain_customer  #
            ######################################################

            xlApp = win32com.client.Dispatch('Excel.Application')  # 打开EXCEL
            xlBook = xlApp.Workbooks.Open(
                r'D:\\Workspace\\Pythonscripts\\testdatas\\chain_customer.xlsx')  # 将随机生成的名称写入链属企业
            xlSht = xlBook.Worksheets('Sheet1')
            xlSht.Cells(2, 1).Value =chain_customer  # 修改合同核心企业名称
            xlBook.Close(SaveChanges=1)  # 完成 关闭保存文件
            del xlApp
        except Exception, e:
            fp = StringIO.StringIO()  # 创建内存文件对象
            traceback.print_exc(file=fp)
            message = fp.getvalue()
            index_file = findStr.findStr(message, "File", 2)
            index_Exception = message.find("Message")
            index_Stacktrace = message.find("Stacktrace:")
            print_message = message[0:index_file] + message[index_Exception:index_Stacktrace]
            time.sleep(1)
            title_index = browser.title.find("-")
            title = browser.title[0:title_index]
            browser.get_screenshot_as_file(shot_path + title + ".png")
            self.assertTrue(False, print_message)

    def test_3_invite_chain_register(self):
        (u"登记核心邀请链属注册")
        browser = self.browser
        browser.implicitly_wait(10)
        try:
            xlApp = win32com.client.Dispatch('Excel.Application')  # 打开EXCEL
            xlBook = xlApp.Workbooks.Open(r'D:\\Workspace\\Pythonscripts\\testdatas\\chain_customer.xlsx')
            xlSht = xlBook.Worksheets('Sheet1')
            chain_customer = xlSht.Cells(2, 1).Value
            chain_name = xlSht.Cells(2, 3).Value
            chain_email = xlSht.Cells(2, 5).Value
            pattern = re.compile(r'\d*')
            chain_password=str(xlSht.Cells(2, 4).Value)
            chain_cellphone= re.search(pattern, str(xlSht.Cells(2, 6).Value)).group()
            xlBook.Close(SaveChanges=1)
            del xlApp
            login.corp_login(self,"core_customer.xlsx")#核心企业登录
            time.sleep(2)
            browser.find_element_by_link_text(u"企业管理").click()#点击企业管理
            time.sleep(2)
            browser.find_element_by_id("input-company-name").send_keys(chain_customer)
            time.sleep(1)
            browser.find_element_by_id("search").click()#点击搜索按钮
            time.sleep(2)
            browser.find_element_by_xpath(".//*[@id='inviteTable']/tbody/tr[3]/td[7]/input").clear()
            browser.find_element_by_xpath(".//*[@id='inviteTable']/tbody/tr[3]/td[7]/input").send_keys(chain_name)#填入联系人
            time.sleep(1)
            browser.find_element_by_xpath(".//*[@id='inviteTable']/tbody/tr[3]/td[8]/input").clear()
            browser.find_element_by_xpath(".//*[@id='inviteTable']/tbody/tr[3]/td[8]/input").send_keys(chain_cellphone)#填入手机号
            time.sleep(1)
            browser.find_element_by_xpath(".//*[@id='inviteTable']/tbody/tr[3]/td[9]/input").clear()
            browser.find_element_by_xpath(".//*[@id='inviteTable']/tbody/tr[3]/td[9]/input").send_keys(chain_email)#填入邮箱
            time.sleep(1)
            browser.find_element_by_xpath(".//a[@class='inviteId']").click()#点击邀请按钮
            time.sleep(2)
            browser.find_element_by_id("invite-next-1").click()#确认邀请
            time.sleep(2)
            browser.find_element_by_id("invite-next-2").click()#确认邀请后点击完成
            time.sleep(2)
            if not browser.find_element_by_css_selector(".assetStatus-4").text==u"已邀请":
                self.assertFalse(True,u"已发出邀请，但是邀请状态没有改变")
            # #############################
            # 发出邀请后登录邮箱等操作
            ############################
            for i in range(1,2):
                regist_cmd = method + "\\chain_regist.exe " +str(i)
                os.system(regist_cmd)
                w.OpenClipboard()
                wds= w.GetClipboardData(win32con.CF_TEXT)
                w.CloseClipboard()
                if "http:" in wds:
                    pattern = re.compile(r'http.*\s')
                    chain_url = re.search(pattern,wds).group()
                    print  chain_url
                    break
            browser.get(chain_url)#登录链属注册
            time.sleep(2)
            browser.find_element_by_id("jiaru").click()#加入平台
            time.sleep(2)
            browser.find_element_by_id("inputPassword").send_keys(chain_password)#输入密码
            time.sleep(2)
            browser.find_element_by_id("inputRePassword").send_keys(chain_password)#确认密码
            time.sleep(2)
            browser.find_element_by_id("getDynamic").click()#点击获取验证码
            time.sleep(5)
            # 获取验证码
            now_handle = browser.current_window_handle
            Dynamic_url = "http://" + host + ".dcfservice.com/v1/public/sms/get?cellphone=" +chain_cellphone
            js_script = 'window.open(' + '"' + Dynamic_url + '"' + ')'
            browser.execute_script(js_script)
            time.sleep(2)
            all_handles = browser.window_handles
            for handle in all_handles:
                if handle != now_handle:
                    browser.switch_to_window(handle)
            Dynamic_code = browser.find_element_by_css_selector("html>body>pre").text
            Dynamic_code = Dynamic_code[1:7]
            browser.close()#关闭验证码页面
            browser.switch_to_window(now_handle)
            # 填写验证码
            browser.find_element_by_id("validateCode").send_keys(Dynamic_code)
            time.sleep(1)
            browser.find_element_by_id("registerbtn").click()
            # 等待8秒进入链属企业在线签约页面
            time.sleep(8)
            browser.execute_script("document.documentElement.scrollTop=document.body.scrollHeight")  # 滑动滚动条至底部
            time.sleep(8)
            os.system(method + "\\click_flash.exe")
            time.sleep(2)
            browser.find_element_by_xpath("//button[@class='btn btn-success'][@type='button']").click()  # 点击立即申请
            time.sleep(2)
            browser.find_element_by_id("submitBtn").click()#上传普通版本
            time.sleep(2)
            ################################################################################
            # 以下为链属企业上传认证资料流程
            ###############################################################################
            browser.find_elements_by_xpath("//span[@class='btn btn-success fileinput-button fileinputBusiness']")[1].click()  # 上传营业执照
            time.sleep(2)
            # upload_file路径，上传图片
            upload_file = method + "\\upload.exe " + data + "test_picture.jpg"
            os.system(upload_file)
            time.sleep(5)
            browser.find_element_by_xpath(".//span[@class='btn btn-success fileinput-button fileinputOrganization']").click()#上传组织机构代码
            time.sleep(2)
            # upload_file路径，上传图片
            upload_file = method + "\\upload.exe " + data + "test_picture.jpg"
            os.system(upload_file)
            time.sleep(5)
            browser.find_elements_by_xpath(
                ".//span[@class='btn btn-success fileinput-button fileinputIdcard']")[2].click()  # 上传身份证正面
            time.sleep(2)
            # upload_file路径，上传图片
            upload_file = method + "\\upload.exe " + data + "test_picture.jpg"
            os.system(upload_file)
            time.sleep(5)
            browser.find_elements_by_xpath(
                ".//span[@class='btn btn-success fileinput-button fileinputIdcard']")[3].click()  # 上传身份证反面
            time.sleep(2)
            # upload_file路径，上传图片
            upload_file = method + "\\upload.exe " + data + "test_picture.jpg"
            os.system(upload_file)
            time.sleep(5)
            browser.find_elements_by_xpath("//span[@class='btn btn-success fileinput-button fileinputBusiness']")[2].click() # 手持身份证上传
            time.sleep(2)
            # upload_file路径，上传图片
            upload_file = method + "\\upload.exe " + data + "test_picture.jpg"
            os.system(upload_file)
            time.sleep(5)
            browser.find_element_by_xpath("//button[@class='btn btn-primary']").click()#点击确认
            time.sleep(2)
            ##################################
            #链属上传服务资料
            ##################################
            time.sleep(2)
            browser.find_element_by_xpath(".//span[@class='btn btn-success fileinput-button fileinputContract']").click()#上传新建协议
            time.sleep(2)
            # upload_file路径，上传图片
            upload_file = method + "\\upload.exe " + data + "test_picture.jpg"
            os.system(upload_file)
            time.sleep(5)
            browser.find_element_by_id("confirm").click()#点击确认按钮
            time.sleep(2)
            browser.find_element_by_xpath(".//button[@class='btn btn-primary']").click()#点击提交资料
            time.sleep(2)
            browser.find_element_by_xpath(".//a[@class='btn btn-success']").click()#点击我知道了
            time.sleep(2)
            browser.implicitly_wait(5)
            try:
                browser.find_element_by_xpath("//div[@class='full-screen-mask-close']").click()#如果有图标的话就点击
            except NoSuchElementException,e:
                print ""
        except Exception, e:
            fp = StringIO.StringIO()  # 创建内存文件对象
            traceback.print_exc(file=fp)
            message = fp.getvalue()
            index_file = findStr.findStr(message, "File", 2)
            index_Exception = message.find("Message")
            index_Stacktrace = message.find("Stacktrace:")
            print_message = message[0:index_file] + message[index_Exception:index_Stacktrace]
            time.sleep(1)
            title_index = browser.title.find("-")
            title = browser.title[0:title_index]
            browser.get_screenshot_as_file(shot_path + title + ".png")
            self.assertTrue(False, print_message)
    def test_4(self):
        (u"运营端认证链属企业")
        browser = self.browser
        browser.implicitly_wait(10)
        try:

            login.operate_login(self, "operation_login.csv")
            time.sleep(2)
            browser.find_element_by_link_text(u"客户管理")
            time.sleep(2)
            browser.find_element_by_link_text(u"客户认证")
            time.sleep(2)
            xlApp = win32com.client.Dispatch('Excel.Application')  # 打开EXCEL
            xlBook = xlApp.Workbooks.Open(r'D:\\Workspace\\Pythonscripts\\testdatas\\chain_customer.xlsx')
            xlSht = xlBook.Worksheets('Sheet1')
            chain_customer = xlSht.Cells(2, 1).Value
            xlBook.Close(SaveChanges=1)
            del xlApp
            browser.find_element_by_id("search-text").send_keys(chain_customer)
            time.sleep(2)
            browser.find_element_by_id("btn-search").click()
            time.sleep(2)
            browser.find_element_by_xpath(".//*[@id='table-account']/tbody/tr/td[6]/div/a[2]").click()#点击认证按钮
            time.sleep(2)
            browser.find_element_by_xpath(".//input[@name='commitMethod'][@value='1']").click()  # 选择普通版本
            time.sleep(3)
            ####################################################################################################
            #                               以下区域填写营业执照相关信息                                       #
            ####################################################################################################
            browser.execute_script("arguments[0].scrollIntoView()",
                                   browser.find_elements_by_css_selector(".crumbs-default")[
                                       1])  # 使审核认证资料信息可视，从而使营业执照右侧通过按钮可视
            time.sleep(2)
            browser.find_element_by_css_selector(
                ".radio.col-xs-6>input[value='2'][name='businessLicense_Pass']").click()  # 点击营业执照的通过按钮
            time.sleep(2)
            browser.find_element_by_css_selector(".form-control[name='enterprise_no']").send_keys("111111")  # 填写营业执照号
            time.sleep(2)
            browser.find_element_by_xpath(
                './/input[@name="businessLicenseNeverExpireFlag"][@value="1"]').click()  # 营业执照无期限
            time.sleep(2)
            browser.find_element_by_xpath(".//input[@name='enterprise_money']").send_keys("10000")  # 填写注册资本
            time.sleep(2)
            browser.execute_script("arguments[0].scrollIntoView()",
                                   browser.find_elements_by_xpath('.//div[@class="f12"]')[1])
            time.sleep(1)
            browser.find_element_by_xpath(".//input[@name='organizationNo_Pass'][@value='2']").click()  # 点击组织机构代码通过
            time.sleep(2)
            browser.find_element_by_xpath(".//input[@class='form-control'][@name='organization_no']").send_keys(
                "111111")  # 填写组织机构代码
            time.sleep(2)
            browser.find_element_by_xpath(".//input[@class='form-control'][@name='organization_no_regist']").send_keys(
                "111111")  # 填写登记号
            time.sleep(2)
            Select(browser.find_element_by_id("province")).select_by_visible_text(u"上海")  # 选择上海市
            time.sleep(2)
            Select(browser.find_element_by_id("city")).select_by_visible_text(u"上海市")  # 选择上海市
            time.sleep(2)
            browser.execute_script('''arguments[0].value="2020-2-21"''',
                                   browser.find_element_by_xpath(".//input[@class='form-control datepicker']"))
            time.sleep(1)
            browser.execute_script("arguments[0].scrollIntoView()",
                                   browser.find_elements_by_xpath('.//div[@class="f12"]')[2])  # 上移组织机构代码，使操作者身份证区域可见
            time.sleep(1)
            browser.find_element_by_xpath(".//input[@name='operatorId_Pass'][@value='2']").click()  # 点击操作者身份证通过
            time.sleep(2)
            browser.find_element_by_xpath(".//input[@name='operator_user_name']").send_keys(u"周大强")  # 输入身份证名称
            time.sleep(2)
            browser.find_element_by_xpath(".//input[@name='operator_ID']").send_keys("222222222222222222")  # 输入身份证号码
            time.sleep(2)
            browser.find_element_by_xpath(".//input[@name='operator_ID_never_expire']").click()  # 长期
            time.sleep(2)
            browser.execute_script("arguments[0].scrollIntoView()",
                                   browser.find_elements_by_xpath('.//div[@class="f12"]')[3])  # 上移身份证操作者区域，使操作者手持身份证照片
            time.sleep(2)
            browser.find_element_by_xpath(".//input[@name='picHandling_Pass'][@value='2']").click()  # 点击操作者手持身份证照片通过按钮
            time.sleep(2)
            browser.execute_script("document.documentElement.scrollTop=0")  # 滑动滚动条至顶部
            time.sleep(2)
            browser.find_element_by_xpath(".//input[@name='commitMethod'][@value='1']").click()  # 再次选择普通版本
            time.sleep(2)
            browser.find_element_by_id("btnSubmit").click()  # 点击提交
            time.sleep(2)
            browser.find_element_by_css_selector("#modalFooter>a").click()  # 点击返回列表
            time.sleep(2)
            if not browser.find_element_by_xpath(".//*[@id='table-account']/tbody/tr/td[4]/span").text==u"已认证":
                self.assertFalse(True,u"已认证，但是认证状态没有改变")
            time.sleep(2)
            #################################
            #服务审核
            #################################
            browser.find_element_by_link_text(u"服务审核").click()#点击服务审核
            time.sleep(2)
            browser.find_element_by_id("search-text").send_keys(chain_customer)
            time.sleep(2)
            browser.find_element_by_id("btn-search").click()
            time.sleep(2)
            browser.find_element_by_xpath(".//a[@class='nowrap']").click()#点击审核按钮
            time.sleep(2)
            browser.execute_script('''arguments[0].value="2016-06-21"''',browser.find_element_by_id("sign_date"))#合同签署日
            time.sleep(2)
            browser.execute_script('''arguments[0].value="2020-02-21"''',browser.find_element_by_id("expire_date"))  # 合同到期日
            time.sleep(2)
            browser.find_element_by_xpath(".//input[@type='checkbox'][@name='have_contract_account_period']").click()#合同账期填写无
            time.sleep(2)
            browser.find_element_by_xpath(".//input[@type='checkbox'][@name='have_final_expire_date']").click()#最终到期日填写长期
            time.sleep(2)
            browser.execute_script("arguments[0].scrollIntoView()",browser.find_element_by_xpath("//div[@class='crumbs']"))
            time.sleep(2)
            browser.find_element_by_xpath("//input[@type='radio'][@value='2']").click()#点击通过按钮
            time.sleep(2)
            browser.find_element_by_id("btnSubmit").click()#审核成功后点击提交
            time.sleep(2)
            browser.find_element_by_xpath(".//*[@id='modalFooter']/a").click()#点击返回列表
            time.sleep(4)
            browser.find_element_by_link_text(u"服务审核")
            time.sleep(2)
            browser.find_element_by_id("search-text").send_keys(chain_customer)
            time.sleep(2)
            browser.find_element_by_id("btn-search").click()
            time.sleep(2)
            if not browser.find_element_by_xpath(".//span[@class='status certified']").text==u"已通过":
                self.assertFalse(False,u"认证已提交但是认证状态没有改变")
            ####################################
            #认证服务审核成功后获取通用结算户
            ####################################
            browser.find_element_by_link_text(u"群星支付").click()
            time.sleep(2)
            browser.find_element_by_id("search-text").send_keys(chain_customer)
            i = 0
            General_account = True
            while True:
                time.sleep(10)
                browser.find_element_by_id("btn-search").click()  # 在群星支付界面搜索链属企业账户
                time.sleep(2)
                elements = browser.find_elements_by_xpath(".//*[@id='table-account']/tbody/tr")
                if len(elements) > 3:
                    self.assertFalse(False, u"该账户存在多余3个账户")
                try:
                    if browser.find_element_by_xpath(
                            ".//*[@id='table-account']/tbody/tr[2]/td[4]/div[2]").is_displayed():
                        chain_General_account = str(browser.find_element_by_xpath(
                            ".//*[@id='table-account']/tbody/tr[2]/td[4]/div[2]").text)  # 获取通用结算户
                        break
                except NoSuchElementException, e:
                    print ""
                i = i + 1
                if i == 10:
                    General_account = False
                    break
            if not General_account:
                self.assertFalse(True, "该账户已认证成功，但是没有创建通用结算户")
            time.sleep(2)
            chain_account_id = str(browser.find_element_by_css_selector(".odd.grouped>td[data-name='accountId']").text)  # 获取群星id号
            time.sleep(2)
            xlApp = win32com.client.Dispatch('Excel.Application')  # 打开EXCEL
            xlBook = xlApp.Workbooks.Open(r'D:\\Workspace\\Pythonscripts\\testdatas\\chain_customer.xlsx')
            xlSht = xlBook.Worksheets('Sheet1')
            xlSht.Cells(2, 2).Value = chain_General_account
            xlBook.Close(SaveChanges=1)  # 完成 关闭保存文件
            del xlApp
            #############################################
            # 记账方式充值                               #
            #############################################
            browser.find_element_by_link_text(u"群星支付").click()
            time.sleep(2)
            browser.find_element_by_css_selector(".nav-list.account-list>div").click()  # 点击账务管理
            time.sleep(2)
            browser.find_elements_by_css_selector(".nav-list.account-list>ul>li>a")[1].click()  # 点击手工记账
            time.sleep(5)
            browser.find_elements_by_css_selector(".select2-selection__arrow")[0].click()
            # browser.find_element_by_xpath("html/body/div[1]/div[3]/div[2]/form/div[1]/div/span/span[1]/span/span[2]").click()
            time.sleep(2)
            browser.find_element_by_css_selector(".select2-search__field").send_keys(u"一般户充值")  # 选择一般户充值
            time.sleep(2)
            browser.find_element_by_css_selector(".select2-search__field").send_keys(Keys.ENTER)  # 按enter键输入
            time.sleep(2)
            browser.find_elements_by_css_selector(".select2-selection__arrow")[3].click()  # 点击收款方
            time.sleep(2)
            browser.find_element_by_css_selector(".select2-search__field").send_keys(chain_customer)  # 输入收款方
            time.sleep(2)
            browser.find_element_by_css_selector(".select2-search__field").send_keys(Keys.ENTER)  # 按enter键输入
            time.sleep(2)
            Select(browser.find_element_by_css_selector(
                ".form-control[data-type='receiveAccount'][name='selectReceiveAccount']")).select_by_value(
                chain_account_id)  # 选择通用结算户
            time.sleep(2)
            browser.find_element_by_css_selector(".form-control.num").send_keys(10000000)  # 充值完成
            time.sleep(2)
            browser.find_element_by_id("btn-save").click()  # 保存按钮
            time.sleep(2)
            browser.find_element_by_id("close").click()
        except Exception, e:
            fp = StringIO.StringIO()  # 创建内存文件对象
            traceback.print_exc(file=fp)
            message = fp.getvalue()
            index_file = findStr.findStr(message, "File", 2)
            index_Exception = message.find("Message")
            index_Stacktrace = message.find("Stacktrace:")
            print_message = message[0:index_file] + message[index_Exception:index_Stacktrace]
            time.sleep(1)
            title_index = browser.title.find("-")
            title = browser.title[0:title_index]
            browser.get_screenshot_as_file(shot_path + title + ".png")
            self.assertTrue(False, print_message)
    def test_5_chain_sign_online(self):
        (u"新建链属在线签约")
        browser = self.browser
        browser.implicitly_wait(10)
        try:
            login.corp_login(self, "chain_customer.xlsx")
            #############################
            #针对第一次登录             #
            #############################
            browser.implicitly_wait(5)
            try:
                if browser.find_element_by_xpath("html/body/div[2]/div[1]").is_displayed():
                    browser.find_element_by_xpath("html/body/div[2]/div[1]").click()
            except NoSuchElementException, e:
                print ""
            try:
                if browser.find_element_by_xpath(".//*[@id='zhongjin-banner']/div[1]").is_displayed():
                    browser.find_element_by_xpath(".//*[@id='zhongjin-banner']/div[1]").click()
            except NoSuchElementException, e:
                print ""
            browser.implicitly_wait(10)
            browser.execute_script("arguments[0].click()",browser.find_element_by_xpath("html/body/div[1]/div[2]/div/div[2]/div/a"))#点击立即签约
            time.sleep(1)
            ###########################################
            #            第一次使用安装数字证书       #
            ###########################################
            try:
                xlApp = win32com.client.Dispatch('Excel.Application')  # 打开EXCEL
                xlBook = xlApp.Workbooks.Open(r'D:\\Workspace\\Pythonscripts\\testdatas\\chain_customer.xlsx')
                xlSht = xlBook.Worksheets('Sheet1')
                pattern = re.compile(r'\d*')
                chain_cellphone= re.search(pattern, str(xlSht.Cells(2, 6).Value)).group()
                chain_bank_no=str(xlSht.Cells(2, 2).Value).replace(" ","")
                xlBook.Close(SaveChanges=1)
                del xlApp
                browser.execute_script("arguments[0].click()", browser.find_element_by_id("getDy"))  # 获取验证码
                time.sleep(4)
                now_handle = browser.current_window_handle  # 获取当前的handle
                Dynamic_url = "http://" + host + ".dcfservice.com/v1/public/sms/get?cellphone="+chain_cellphone # 获取验证码路径
                js_script = 'window.open(' + '"' + Dynamic_url + '"' + ')'
                browser.execute_script(js_script)
                time.sleep(2)
                all_handles = browser.window_handles
                for handle in all_handles:
                    if handle != now_handle:
                        browser.switch_to_window(handle)
                Dynamic_code = browser.find_element_by_css_selector("html>body>pre").text
                Dynamic_code = Dynamic_code[1:7]  # 截取字符串获取验证码
                browser.close()
                browser.switch_to_window(now_handle)  # 切换回以前handle
                browser.find_element_by_id("dyCode").send_keys(Dynamic_code)
                time.sleep(1)
                browser.find_element_by_id("validateDy").click()
                time.sleep(1)
                browser.find_element_by_id("installCfca").click()  # 点击立即安装
                time.sleep(10)#等待页面刷新
                browser.execute_script("arguments[0].click()", browser.find_element_by_xpath("html/body/div[1]/div[2]/div/div[2]/div/a"))  # 点击立即签约
            except NoSuchElementException, e:
                print "The customer has installed security controls "

            browser.find_element_by_css_selector('''.show-form-account[href="javascript:void(0)"]''').click()#点击创建融资账户‘
            time.sleep(2)
            browser.find_element_by_id("account").send_keys(chain_bank_no)#填写融资账户
            time.sleep(2)
            Select(browser.find_element_by_id("bank")).select_by_visible_text(u"中信银行")#填写开户行
            time.sleep(2)
            Select(browser.find_element_by_id("province")).select_by_visible_text(u"上海市")#开户省上海市
            time.sleep(2)
            Select(browser.find_element_by_id("city")).select_by_visible_text(u"上海市")#城市为上海市
            time.sleep(2)
            Select(browser.find_element_by_id("child-bank")).select_by_visible_text(u"中信银行股份有限公司上海徐汇支行")
            time.sleep(2)
            browser.find_element_by_id("bankCreate").click()#点击创建融资账户
            time.sleep(2)
            browser.find_element_by_css_selector(".btn.btn-primary.next-step").click()#点击下一步
            time.sleep(5)#等待供应链邀请函
            browser.find_element_by_xpath("//input[@type='checkbox']").click()#仔细阅读
            time.sleep(5)
            browser.find_element_by_xpath("//button[@class='btn btn-success download-btn']").click()#同意并下载
            time.sleep(4)  # d等待生成excel
            cmd = method + "downloads.exe " + assert_path
            os.system(cmd)
            time.sleep(5)  # 等待下载完成
        except Exception, e:
            fp = StringIO.StringIO()  # 创建内存文件对象
            traceback.print_exc(file=fp)
            message = fp.getvalue()
            index_file = findStr.findStr(message, "File", 2)
            index_Exception = message.find("Message")
            index_Stacktrace = message.find("Stacktrace:")
            print_message = message[0:index_file] + message[index_Exception:index_Stacktrace]
            time.sleep(1)
            title_index = browser.title.find("-")
            title = browser.title[0:title_index]
            browser.get_screenshot_as_file(shot_path + title + ".png")
            self.assertTrue(False, print_message)
    @classmethod
    def tearDownClass(cls):
        cls.browser.close()
        cls.browser.quit()

