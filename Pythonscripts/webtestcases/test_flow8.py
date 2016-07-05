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
import  datetime
import StringIO
import traceback
import csv
from selenium.webdriver.common.keys import Keys
import  os
import calendar
import win32com.client
reload(sys)
sys.setdefaultencoding('utf8')
import ConfigParser
from selenium.webdriver.support.ui import Select
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
#读取数据库文件
USER=cf.get('database','user')
HOST=cf.get('database','host')
PASSWORD=cf.get('database','password')
PORT=cf.get('database','port')
DATABASE=cf.get('database','dcf_user')
DATABASE1=cf.get('database','dcf_settlement')
DATABASE2=cf.get('database','dcf_payment')
DATABASE3=cf.get('database','dcf_loan')
#读取链属企业
xlApp = win32com.client.Dispatch('Excel.Application')  # 打开EXCEL
xlBook = xlApp.Workbooks.Open(r'D:\\Workspace\\Pythonscripts\\testdatas\\chain_customer.xlsx')
xlSht = xlBook.Worksheets('Sheet1')
chain_customer=xlSht.Cells(2, 1).Value  # 读取链属企业名称
financing_bank_no=str(xlSht.Cells(2, 2).Value).replace(" ","")
#读取核心企业
xlApp = win32com.client.Dispatch('Excel.Application')  # 打开EXCEL
xlBook = xlApp.Workbooks.Open(r'D:\\Workspace\\Pythonscripts\\testdatas\\core_customer.xlsx')  # 将随机生成的名称写入链属企业
xlSht = xlBook.Worksheets('Sheet1')
core_customer=xlSht.Cells(2, 1).Value  # 读取核心企业名称
xlBook.Close(SaveChanges=1)
del xlApp
#读取机构
xlApp = win32com.client.Dispatch('Excel.Application')  # 打开EXCEL
xlBook = xlApp.Workbooks.Open(r'D:\\Workspace\\Pythonscripts\\testdatas\\institution_data.xlsx')
xlSht = xlBook.Worksheets('Sheet1')
return_bank_no=str(xlSht.Cells(2, 2).Value).replace(" ","")
institution_name=xlSht.Cells(2,1).Value
xlBook.Close(SaveChanges=1)
del xlApp
#读取核心企业
xlApp = win32com.client.Dispatch('Excel.Application')  # 打开EXCEL
xlBook = xlApp.Workbooks.Open(r'D:\\Workspace\\Pythonscripts\\testdatas\\core_customer.xlsx')  # 将随机生成的名称写入链属企业
xlSht = xlBook.Worksheets('Sheet1')
core_customer=xlSht.Cells(2, 1).Value  # 读取核心企业名称
xlBook.Close(SaveChanges=1)
del  xlApp
#读取截图存放路径
shot_path=cf.get('shotpath','path')
#读取下载包存放路径
assert_path=cf.get('assert_package','assert_path')
class loan_flow(unittest.TestCase):
    (u"放款流程")
    @classmethod
    def setUpClass(cls):
        cls.browser=webdriver.Firefox(profile)
        cls.browser.maximize_window()
        lcoal_time = str(time.strftime("%Y/%m/%d", time.localtime()))
        cls.start_time = lcoal_time
    def test_1_apply_finance(self):
        (u"链属企业申请融资")
        browser=self.browser
        browser.implicitly_wait(10)
        financing_amount=0
        financing_cost=0
        try:
            login.corp_login(self,"chain_customer.xlsx")
            time.sleep(1)
            #对于第一次登陆，要关闭引导页
            try:
                browser.find_element_by_xpath(".//*[@id='zhongjin-banner']/div[1]").click()
                time.sleep(1)
            except NoSuchElementException,e:
                print ""
            try:
                browser.execute_script("arguments[0].click()",browser.find_element_by_id("today"))
            except Exception,e:
                print ""
            time.sleep(5)
            xlApp = win32com.client.Dispatch('Excel.Application')  # 打开EXCEL
            xlBook = xlApp.Workbooks.Open(
                r'D:\\Workspace\\Pythonscripts\\testdatas\\transaction_flow.xlsx')  # 读取交易流水
            xlSht = xlBook.Worksheets('Sheet2')
            for i in range(2, xlSht.UsedRange.Rows.Count + 1):#对交易流水中的单据进行全部融资
                click_xpath='//*[@id="'+str(xlSht.Cells(i,1).Value)+'"]/td[1]/input'
                financing_amount=financing_amount+float(xlSht.Cells(i,4).Value)
                financing_cost=financing_cost+float(xlSht.Cells(i,5).Value)
                browser.execute_script("arguments[0].click()",browser.find_element_by_xpath(click_xpath))#点击融资
            time.sleep(1)
            browser.find_element_by_xpath(".//button[@class='btn btn-success loanMore']").click()
            time.sleep(2)
            ###########################################
            #            第一次使用安装数字证书       #
            ###########################################
            browser.implicitly_wait(5)
            try:
                browser.execute_script("arguments[0].click()", browser.find_element_by_id("getDy"))  # 获取验证码
                time.sleep(2)
                now_handle = browser.current_window_handle  # 获取当前的handle
                Dynamic_url = "http://" + host + ".dcfservice.com/v1/public/sms/get?cellphone=18751986831"  # 获取验证码路径
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
                print  Dynamic_code
                browser.switch_to_window(now_handle)  # 切换回以前handle
                browser.find_element_by_id("dyCode").send_keys(Dynamic_code)
                time.sleep(1)
                browser.find_element_by_id("validateDy").click()
                time.sleep(1)
                browser.find_element_by_id("installCfca").click()  # 点击立即安装
                time.sleep(10)
                for i in range(2, xlSht.UsedRange.Rows.Count + 1):  # 对交易流水中的单据进行全部融资
                    click_xpath = '//*[@id="' + str(xlSht.Cells(i, 1).Value) + '"]/td[1]/input'
                    browser.find_element_by_xpath(click_xpath).click()  # 点击融资
                browser.find_element_by_xpath(".//button[@class='btn btn-success loanMore']").click()
                time.sleep(2)
                ###########未写完
            except NoSuchElementException, e:
                print "The customer has installed security controls "
            #############################################
            #N+1链属点击融资申请后断言验证              #
            #############################################
            if (browser.find_element_by_xpath(".//*[@id='loan-account']//td[@class='number']").text).replace(" ", "")!=financing_bank_no:  #验证融资账户
                self.assertFalse(False,"account do not match")
            if float(str(browser.find_element_by_xpath(".//span[@class='legendApplyAmount']/b").text).replace(",",""))-financing_amount!=0: #断言融资总金额
                self.assertFalse(True, "financing_amount is worng")
            time.sleep(1)
            if float(str(browser.find_element_by_xpath(".//span[@class='legendApplyCost']/b[@class='applyPrice']").text).replace(",",""))-financing_cost!=0:#断言融资总成本
                self.assertFalse(True, "financing_cost is worng")
            time.sleep(1)
            #######应收账款表断言后续补上######################
            if (browser.find_element_by_xpath(".//div[@id='return-account']//td[@class='number']").text).replace(" ","") != return_bank_no:  # 验证汇款账户
                self.assertFalse(True, "return account do not match")
            time.sleep(1)
            browser.find_element_by_id("financing-apply-now").click()#点击融资申请
            ####大数据处理####
            time_flag=0
            while True:
                time.sleep(5)
                try:
                    if browser.find_element_by_xpath(".//*[@id='depositNotRequired']/div/div/div[1]/button").is_displayed():
                        time.sleep(2)
                        browser.find_element_by_xpath(".//*[@id='depositNotRequired']/div/div/div[1]/button").click()
                        break
                except NoSuchElementException,e:
                    time_flag=time_flag+1
                if time_flag==20:
                    break
            time.sleep(2)
            try:
                browser.execute_script("arguments[0].click()", browser.find_element_by_id("today"))
            except Exception, e:
                print ""
            browser.execute_script("arguments[0].click()",browser.find_element_by_id("today"))
            time.sleep(5)
            #################################
            #点击融资申请后，查看已融资模块
            #################################
            for i in range(2, xlSht.UsedRange.Rows.Count + 1):
                amount = float(xlSht.Cells(i, 3).Value)
                buyer_name=core_customer
                start_time=self.start_time
                loan_start_time=self.start_time#对于实时放款的放款日就是当天
                financing_amount=float(xlSht.Cells(i, 4).Value)#申请融资金额
                financing_cost=float(xlSht.Cells(i, 5).Value)#申请融资的成本
                loan_document_no_xpath = '//div[@class="listDiv"][@data-type="financinging"]//tbody//tr[@id="' + str(xlSht.Cells(i, 1).Value) + '"]/td[1]' #在已融资模块查找单据号
                seller_name_xpath = '//div[@class="listDiv"][@data-type="financinging"]//tbody//tr[@id="' + str(xlSht.Cells(i, 1).Value) + '"]/td[2]'#在已融资模块查找核心企业
                amount_xpath = '//div[@class="listDiv"][@data-type="financinging"]//tbody//tr[@id="' + str(xlSht.Cells(i, 1).Value) + '"]/td[3]'#在已申请融资模块查找单据金额
                start_time_xpath = '//div[@class="listDiv"][@data-type="financinging"]//tbody//tr[@id="' + str(xlSht.Cells(i, 1).Value) + '"]/td[4]'#在已申请融资模块查找起始日期
                loan_start_xpath= '//div[@class="listDiv"][@data-type="financinging"]//tbody//tr[@id="' + str(xlSht.Cells(i, 1).Value) + '"]/td[5]'#预计放款日
                financing_amount_xpath = '//div[@class="listDiv"][@data-type="financinging"]//tbody//tr[@id="' + str(xlSht.Cells(i, 1).Value) + '"]/td[6]'#融资金额
                financing_cost_xpath = '//div[@class="listDiv"][@data-type="financinging"]//tbody//tr[@id="' + str(xlSht.Cells(i, 1).Value) + '"]/td[7]'#融资成本
                ################断言填写#################################
                try:
                    if browser.find_element_by_xpath(loan_document_no_xpath).text != str(xlSht.Cells(i, 2).Value):  # 单据号断言
                        browser.get_screenshot_as_file(shot_path + u"单据号不一致" + ".png")  # 对错误增加截图
                        self.assertFalse(True, "Transaction document No. is inconsistent with EXCEL")
                    if browser.find_element_by_xpath(seller_name_xpath).text != buyer_name:  # 客户名称断言
                        browser.get_screenshot_as_file(shot_path + u"买家名称不一致" + ".png")  # 对错误增加截图
                        self.assertFalse(True, "customer_name is inconsistent with EXCEL")
                    if float(str(browser.find_element_by_xpath(amount_xpath).text).replace(",","")) - amount != 0:  # 单据金额断言
                        browser.get_screenshot_as_file(shot_path + u"上传金额不一致" + ".png")  # 对错误增加截图
                        self.assertFalse(True, "amount is inconsistent with EXCEL")
                    if browser.find_element_by_xpath(start_time_xpath).text != start_time:  # 单据起始日期断言
                        browser.get_screenshot_as_file(shot_path + u"起始日不一致" + ".png")  # 对错误增加截图
                        self.assertFalse(True, "the start_time of document is inconsistent with EXCEL")
                    if browser.find_element_by_xpath(loan_start_xpath).text != loan_start_time:  # 放款日期断言
                        browser.get_screenshot_as_file(shot_path + u"放款日不正确" + ".png")  # 对错误增加截图
                        self.assertFalse(True, "the loan_start_time of document is wrong")
                    if float(str(browser.find_element_by_xpath(financing_amount_xpath).text).replace(",","")) - financing_amount != 0:  # 融资金额断言
                        browser.get_screenshot_as_file(shot_path + u"融资金额计算不正确" + ".png")  # 对错误增加截图
                        self.assertFalse(True, "the financing_days is wrong")
                    if float(str(browser.find_element_by_xpath(financing_cost_xpath).text).replace(",","")) - financing_cost != 0:  # 融资成本断言
                        browser.get_screenshot_as_file(shot_path + u"融资成本不正确" + ".png")  # 对错误增加截图
                        self.assertFalse(True, "the financing_cost is wrong")
                    time.sleep(0.1)
                except NoSuchElementException,e:
                    print   loan_document_no_xpath
                    print str(xlSht.Cells(i, 1).Value)+" load_documnet_id is not find"
            xlBook.Close(SaveChanges=1)
            del xlApp #关闭excel
        except Exception, e:
            fp = StringIO.StringIO()  # 创建内存文件对象
            traceback.print_exc(file=fp)
            message = fp.getvalue()
            index_file = findStr.findStr(message, "File", 2)
            index_Exception = message.find("Message")
            index_Stacktrace = message.find("Stacktrace:")
            print_message = message[0:index_file] + message[index_Exception:index_Stacktrace]
            time.sleep(1)
            title_index=browser.title.find("-")
            title=browser.title[0:title_index]
            # im = ImageGrab.grab()
            # im.save(shot_path +title + ".png")
            browser.get_screenshot_as_file(shot_path +title + ".png")
            self.assertTrue(False, print_message)
    def test_2_institution_approve(self):
        (u"机构审批")
        build_time =str(time.strftime("%Y-%m-%d", time.localtime()))
        id_str=""
        order_by_str=''
        cloum=2
        browser=self.browser
        browser.implicitly_wait(10)
        xlApp = win32com.client.Dispatch('Excel.Application')  # 打开EXCEL
        xlBook = xlApp.Workbooks.Open(r'D:\\Workspace\\Pythonscripts\\testdatas\\transaction_flow.xlsx')  # 读取交易流水
        xlSht = xlBook.Worksheets('Sheet2')
        for i in range(2, xlSht.UsedRange.Rows.Count + 1):
            id_str = id_str + "'" + str(xlSht.Cells(i,1).Value) + "'" + ","
            order_by_str=order_by_str+str(xlSht.Cells(i,1).Value)+","
        id_str = id_str[:-1]
        order_by_str = order_by_str[:-1]  # 去掉最后一个逗号
        order_by_str = "'" + order_by_str + "'"
        try:
            login.corp_login(self,'institution_data.xlsx')#结构登录
            browser.find_element_by_css_selector("#tab-loanTip>.operatorName").click()#点击融资管理
            time.sleep(2)
            browser.find_element_by_xpath(".//*[@id='loanBtn']/a[2]").click()#点击贷中管理
            time.sleep(2)
            browser.find_element_by_xpath(".//*[@id='topTab']/ul/li[2]/a").click()#点击融资申请审批
            time.sleep(4)
            try:
                browser.find_element_by_xpath(".//span[@class='core-name'][text()='"+core_customer+"']").click()#选择核心企业来进行审批
            except Exception,e:
                print ""
            time.sleep(2)
            ######################################
            # 查询资产包编号                     #
            ######################################
            try:
                # 数据库连接
                conn = mysql.connector.connect(host=HOST, user=USER, passwd=PASSWORD, db=DATABASE3, port=PORT)
                # 创建游标
                cur = conn.cursor()
                # assert_package_id查询
                sql = 'SELECT a.asset_package_id  from t_asset_package_loan_application_association a LEFT JOIN t_loan_application b on a.loan_application_id=b.loan_application_id where b.loan_document_id IN (' + id_str + ')'+'ORDER BY FIND_IN_SET (b.loan_document_id,'+order_by_str+')'
                print sql
                cur.execute(sql)
                # 获取查询结果
                result_set = cur.fetchall()
                if result_set:
                    for row in result_set:
                        assert_package_id = row[0]  # 从数据库取得id号
                        xlSht.Cells(cloum, 6).Value=str(assert_package_id)#将资产包编号写入excel
                        cloum=cloum+1
                else:
                    self.assertTrue(False, "the loan_document_id do not exsit in database!")
                # 关闭游标和连接
                cur.close()
                conn.close()
                xlBook.Close(SaveChanges=1)  # 完成 关闭保存文件
                del xlApp
            except mysql.connector.Error, e:
                print e.message
            time.sleep(2)
            ##########################################
            #根据资产包编号勾选所有融资单据
            ##########################################
            xlApp = win32com.client.Dispatch('Excel.Application')  # 打开EXCEL
            xlBook = xlApp.Workbooks.Open(r'D:\\Workspace\\Pythonscripts\\testdatas\\transaction_flow.xlsx')  # 读取交易流水
            xlSht = xlBook.Worksheets('Sheet2')
            for i in range(2,xlSht.UsedRange.Rows.Count+1):
                package_id=xlSht.Cells(i, 6).Value
                try:
                    browser.find_element_by_xpath(".//td[text()='"+package_id+"']/preceding-sibling::td[1]/input").click()#勾选当前单据
                    ##################
                    #断言每一条编号
                    ##################
                    core_customer_xpath=".//td[text()='"+package_id+"']/following-sibling::td[1]"
                    assert_package_status=".//td[text()='"+package_id+"']/following-sibling::td[2]/span"
                    apply_amount = ".//td[text()='" + package_id + "']/following-sibling::td[3]"
                    assert_amount=".//td[text()='" + package_id + "']/following-sibling::td[4]"
                    assert_build_time=".//td[text()='" + package_id + "']/following-sibling::td[5]"
                    if browser.find_element_by_xpath(core_customer_xpath).text!=core_customer:
                        self.assertTrue(True,u"核心企业不一致")
                    if browser.find_element_by_xpath(assert_package_status).text!=u"待审批":
                        self.assertTrue(True,u"审批状态不是待审批状态")
                    if float(str(browser.find_element_by_xpath(apply_amount).text).replace(",",""))-float(xlSht.Cells(i, 4).Value)!=0:
                        self.assertTrue(True, u"申请金额不正确")
                    if float(str(browser.find_element_by_xpath(assert_amount).text).replace(",",""))-float(xlSht.Cells(i, 3).Value)!=0:
                        self.assertTrue(True, u"资产包金额不正确")
                    if browser.find_element_by_xpath(assert_build_time).text!=build_time:
                        self.assertTrue(True, u"生成日期不正确")
                    time.sleep(0.1)
                except NoSuchElementException,e:
                    ################################
                    #如果第一次没有知啊到再一次找
                    #################################
                    wait_time=0
                    while True:
                        time.sleep(3)
                        try:
                            browser.find_element_by_xpath(
                                ".//td[text()='" + package_id + "']/preceding-sibling::td[1]/input").click()  # 勾选当前单据
                            ##################
                            # 断言每一条编号
                            ##################
                            if browser.find_element_by_xpath(core_customer_xpath).text != core_customer:
                                self.assertTrue(True, u"核心企业不一致")
                            if browser.find_element_by_xpath(assert_package_status).text != u"待审批":
                                self.assertTrue(True, u"审批状态不是待审批状态")
                            if float(str(browser.find_element_by_xpath(apply_amount).text).replace(",", "")) - float(
                                    xlSht.Cells(i, 4).Value) != 0:
                                self.assertTrue(True, u"申请金额不正确")
                            if float(str(browser.find_element_by_xpath(assert_amount).text).replace(",", "")) - float(
                                    xlSht.Cells(i, 3).Value) != 0:
                                self.assertTrue(True, u"资产包金额不正确")
                            if browser.find_element_by_xpath(assert_build_time).text != build_time:
                                self.assertTrue(True, u"生成日期不正确")
                            time.sleep(0.1)
                        except NoSuchElementException,e:
                            print ""
                        wait_time=wait_time+1
                        if wait_time==3:
                            print  u"资产包编号为:"+package_id
                            self.assertTrue(False,u"没有找到资产包编号")
                            break
            xlBook.Close(SaveChanges=1)
            del xlApp

            ####################
            # 下载资产包       #
            ####################
            if not os.path.exists(assert_path):
                os.makedirs(assert_path)#不存在创建
            browser.find_element_by_xpath(".//button[@class='btn btn-primary head-download-btn']").click()#点击下载资产包
            wait_time = 0
            while True:  #用循环方法等待生成文件完成
                time.sleep(5)
                try:
                    if browser.find_element_by_xpath("//div[text()='正在生成文件，请稍后……']").is_displayed():
                        wait_time = wait_time + 1
                    else:
                        break
                except NoSuchElementException,e:
                    break
                if wait_time == 50:
                    break
            time.sleep(1)
            cmd=method+"downloads.exe "+assert_path
            os.system(cmd)
            time.sleep(5)#等待下载完成
            browser.find_element_by_xpath(".//button[@class='btn btn-success head-success-btn']").click()#下载资产包后同意
            time.sleep(2)
            browser.find_element_by_id("approvalAgree").click()#同意放款
            wait_time = 0
            while True:#用循环等等
                time.sleep(5)
                try:
                    if browser.find_element_by_id("approvalAgree").is_displayed():
                        wait_time = wait_time + 1
                    else:
                        break
                except NoSuchElementException,e:
                    break
                if wait_time == 50:
                    break
            time.sleep(2)
            ##########################################
            #审批通过后产看已审批单据里是否含有该单据#
            ##########################################
            browser.execute_script("arguments[0].click()",browser.find_element_by_id("asset_package_state_wrapper_drop"))
            time.sleep(1)
            browser.find_element_by_xpath(".//*[@id='asset_package_state']/li[3]/a").click()
            time.sleep(2)
            xlApp = win32com.client.Dispatch('Excel.Application')  # 打开EXCEL
            xlBook = xlApp.Workbooks.Open(r'D:\\Workspace\\Pythonscripts\\testdatas\\transaction_flow.xlsx')  # 读取交易流水
            xlSht = xlBook.Worksheets('Sheet2')
            for i in range(2, xlSht.UsedRange.Rows.Count + 1):
                package_id = xlSht.Cells(i, 6).Value
                try:
                    ##################
                    # 断言每一条编号
                    ##################
                    core_customer_xpath = ".//td[text()='" + package_id + "']/following-sibling::td[1]"
                    assert_package_status = ".//td[text()='" + package_id + "']/following-sibling::td[2]/span"
                    apply_amount = ".//td[text()='" + package_id + "']/following-sibling::td[3]"
                    assert_amount = ".//td[text()='" + package_id + "']/following-sibling::td[4]"
                    assert_build_time = ".//td[text()='" + package_id + "']/following-sibling::td[5]"
                    assert_approve_time = ".//td[text()='" + package_id + "']/following-sibling::td[8]"
                    assert_approve_pepople = ".//td[text()='" + package_id + "']/following-sibling::td[9]"
                    if browser.find_element_by_xpath(core_customer_xpath).text != core_customer:
                        self.assertTrue(True, u"核心企业不一致")
                    if browser.find_element_by_xpath(assert_package_status).text != u"待审批":
                        self.assertTrue(True, u"审批状态不是待审批状态")
                    if float(str(browser.find_element_by_xpath(apply_amount).text).replace(",","")) - float(xlSht.Cells(i, 4).Value) != 0:
                        self.assertTrue(True, u"申请金额不正确")
                    if float(str(browser.find_element_by_xpath(assert_amount).text).replace(",","")) - float(xlSht.Cells(i, 3).Value) != 0:
                        self.assertTrue(True, u"资产包金额不正确")
                    if browser.find_element_by_xpath(assert_build_time).text != build_time:
                        self.assertTrue(True, u"生成日期不正确")
                    if browser.find_element_by_xpath(assert_approve_time).text != build_time:
                        self.assertTrue(True, u"审批日期不正确")
                    if browser.find_element_by_xpath(assert_approve_pepople).text !=institution_name:
                        self.assertTrue(True, u"审批人不正确")
                    time.sleep(0.1)
                except NoSuchElementException, e:
                    self.assertTrue(False, u"资产包编号没有找到")
            xlBook.Close(SaveChanges=1)
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
    def test_3_operation_approve(self):
        (u"运营端审批通过")
        browser=self.browser
        browser.implicitly_wait(10)
        try:
            login.operate_login(self,"operation_login.csv")
            time.sleep(1)
            browser.find_element_by_link_text("融资管理").click()
            time.sleep(3)
            browser.find_element_by_xpath("//li[@class='nav-list financePage']/div").click()#点击融资按钮
            time.sleep(2)
            browser.find_element_by_xpath("//a[text()='融资交易核对']").click()#点击交易校验核对
            time.sleep(3)
            browser.find_element_by_xpath("//span[text()='请选择机构']/following-sibling::span").click()
            time.sleep(3)
            browser.find_element_by_css_selector(".select2-search__field").send_keys(institution_name)#搜索框输入机构名称
            time.sleep(1)
            browser.find_element_by_css_selector(".select2-search__field").send_keys(Keys.ENTER)#按enter输入
            time.sleep(1)
            Select(browser.find_element_by_id("checkStatus")).select_by_visible_text(u"未进行检验")
            time.sleep(1)
            browser.find_element_by_id("searchBtn").click()#点击搜索按钮
            time.sleep(3)
            browser.find_element_by_xpath("//button[@id='pageSizeWraper']/following-sibling::button").click()#点击分页按钮
            time.sleep(1)
            browser.find_element_by_xpath(".//*[@id='pageSizeName']//a[text()='500']").click()#分页选择500页
            time.sleep(2)
            page_index=2
            ############################################
            #通过excel中数据循环查找验证各单据状态
            ############################################
            xlApp = win32com.client.Dispatch('Excel.Application')  # 打开EXCEL
            xlBook = xlApp.Workbooks.Open(r'D:\\Workspace\\Pythonscripts\\testdatas\\transaction_flow.xlsx')  # 读取交易流水
            xlSht = xlBook.Worksheets('Sheet2')
            for i in range(2, xlSht.UsedRange.Rows.Count + 1):
                package_id =str(xlSht.Cells(i, 6).Value)
                try:
                    browser.find_element_by_xpath(".//td[text()='" + package_id + "']/preceding-sibling::td[7]/input").click()  # 勾选当前单据

                    ##################
                    # 断言每一条编号
                    ##################
                    # balance_build_time = ".//td[text()='" + package_id + "']/following-sibling::td[5]"
                    payment_customer_xpath = ".//td[text()='" + package_id + "']/preceding-sibling::td[4]"
                    gather_customer_status = ".//td[text()='" + package_id + "']/preceding-sibling::td[3]"
                    loan_amount_xpath = ".//td[text()='" + package_id + "']/preceding-sibling::td[2]"
                    institution_name_xpath = ".//td[text()='" + package_id + "']/preceding-sibling::td[1]"
                    check_status_xpath = ".//td[text()='" + package_id + "']/following-sibling::td[1]"
                    if browser.find_element_by_xpath(payment_customer_xpath).text !=institution_name:
                        self.assertTrue(False, u"付款方显示的不是机构")
                    if browser.find_element_by_xpath(gather_customer_status).text !=chain_customer:
                        self.assertTrue(False, u"收款方显示的不是链属企业")
                    if float(str(browser.find_element_by_xpath(loan_amount_xpath).text).replace(",","")) - float(xlSht.Cells(i, 4).Value-xlSht.Cells(i, 5).Value) != 0:
                        self.assertTrue(False, u"放款的金额不正确")
                    if browser.find_element_by_xpath(institution_name_xpath).text != institution_name:
                        self.assertTrue(False, u"机构名称不正确")
                    if browser.find_element_by_xpath(check_status_xpath).text != u"未进行检验":
                        self.assertTrue(False, u"交易校验核对状态不正确")
                    time.sleep(0.1)
                except NoSuchElementException, e:
                    browser.find_element_by_xpath(".//*[@id='pagebar']//a[text()='"+str(page_index)+"']").click()#点击下一页
                    time.sleep(2)
                    page_index=page_index+1
                    try:
                        browser.find_element_by_xpath(".//td[text()='" + package_id + "']/preceding-sibling::td[7]/input").click()  # 勾选当前单据
                        ##################
                        # 断言每一条编号
                        ##################
                        # balance_build_time = ".//td[text()='" + package_id + "']/following-sibling::td[5]"
                        payment_customer_xpath = ".//td[text()='" + package_id + "']/preceding-sibling::td[4]"
                        gather_customer_status = ".//td[text()='" + package_id + "']/preceding-sibling::td[3]"
                        loan_amount_xpath = ".//td[text()='" + package_id + "']/preceding-sibling::td[2]"
                        institution_name_xpath = ".//td[text()='" + package_id + "']/preceding-sibling::td[1]"
                        check_status_xpath = ".//td[text()='" + package_id + "']/following-sibling::td[1]"
                        if browser.find_element_by_xpath(payment_customer_xpath).text != institution_name:
                            self.assertTrue(False, u"付款方显示的不是机构")
                        if browser.find_element_by_xpath(gather_customer_status).text != chain_customer:
                            self.assertTrue(False, u"收款方显示的不是链属企业")
                        if float(str(browser.find_element_by_xpath(loan_amount_xpath).text).replace(",","")) - float(
                                        xlSht.Cells(i, 4).Value - xlSht.Cells(i, 5).Value) != 0:
                            self.assertTrue(False, u"放款的金额不正确")
                        if browser.find_element_by_xpath(institution_name_xpath).text != institution_name:
                            self.assertTrue(False, u"机构名称不正确")
                        if browser.find_element_by_xpath(check_status_xpath).text != u"未进行检验":
                            self.assertTrue(False, u"交易校验核对状态不正确")
                        time.sleep(0.1)
                    except NoSuchElementException,e:
                        self.assertTrue(False,package_id+"is not find")
            xlBook.Close(SaveChanges=1)
            del xlApp
            ########################
            #交易校验核对选择下载单据
            #########################
            if not os.path.exists(assert_path):
                os.makedirs(assert_path)  # 不存在创建
            time.sleep(2)
            browser.find_element_by_id("downloadChecked").click()  # 勾选后选择下载
            wait_time = 0
            while True:  # 用循环方法等待生成文件完成
                time.sleep(5)
                try:
                    if browser.find_element_by_xpath("//div[text()='正在生成文件，请稍后……']").is_displayed():
                        wait_time = wait_time + 1
                    else:
                        break
                except NoSuchElementException,e:
                    break
                if wait_time == 50:
                    break
            time.sleep(2)
            cmd = method + "downloads.exe " + assert_path
            os.system(cmd)
            time.sleep(5)  # 等待下载完成
            browser.execute_script("arguments[0].click()",browser.find_element_by_id("allPass"))#点击通过按钮
            print "执行了通过按钮"
            wait_time=0
            while True:
                print "进入了循环"
                time.sleep(5)
                try:
                    if browser.find_element_by_xpath(".//*[@id='modalFooter']/button").is_displayed():
                        print "出现"
                        break
                except NoSuchElementException,e:
                    wait_time=wait_time+1
                if wait_time==20:
                    self.assertTrue(False,u"交易校验通过失败")
                    break

            #################################
            #产看交易校验结算指令是否发起成功#
            #################################
            time.sleep(2)
            browser.find_element_by_xpath(".//*[@id='modalFooter']/button").click()
            print "执行后续操作"
            time.sleep(5)
            Select(browser.find_element_by_id("checkStatus")).select_by_visible_text(u"全部")
            time.sleep(1)
            browser.find_element_by_id("searchBtn").click()  # 点击搜索按钮
            wait_time = 0
            while True:
                time.sleep(5)
                try:
                    if browser.find_element_by_xpath("//div[@class='loading']").is_displayed():
                        wait_time = wait_time + 1
                    else:
                        break
                except NoSuchElementException, e:
                    break
                if wait_time == 50:
                    break
            browser.find_element_by_xpath("//button[@id='pageSizeWraper']/following-sibling::button").click()  # 点击分页按钮
            time.sleep(1)
            browser.find_element_by_xpath(".//*[@id='pageSizeName']//a[text()='500']").click()  # 分页选择500页
            time.sleep(2)
            xlApp = win32com.client.Dispatch('Excel.Application')  # 打开EXCEL
            xlBook = xlApp.Workbooks.Open(r'D:\\Workspace\\Pythonscripts\\testdatas\\transaction_flow.xlsx')  # 读取交易流水
            xlSht = xlBook.Worksheets('Sheet2')
            for i in range(2, xlSht.UsedRange.Rows.Count + 1):
                package_id = str(xlSht.Cells(i, 6).Value)
                try:
                    browser.implicitly_wait(10)
                    ##################
                    # 断言每一条编号
                    ##################
                    # balance_build_time = ".//td[text()='" + package_id + "']/following-sibling::td[5]"
                    settle_id_xpath = ".//td[text()='" + package_id + "']/following-sibling::td[2]"
                    xlSht.Cells(i, 7).Value = browser.find_element_by_xpath(settle_id_xpath).text
                    payment_customer_xpath = ".//td[text()='" + package_id + "']/preceding-sibling::td[4]"
                    gather_customer_status = ".//td[text()='" + package_id + "']/preceding-sibling::td[3]"
                    loan_amount_xpath = ".//td[text()='" + package_id + "']/preceding-sibling::td[2]"
                    institution_name_xpath = ".//td[text()='" + package_id + "']/preceding-sibling::td[1]"
                    check_status_xpath = ".//td[text()='" + package_id + "']/following-sibling::td[1]"
                    check_operation_xpath=".//td[text()='" + package_id + "']/following-sibling::td[3]"
                    if browser.find_element_by_xpath(payment_customer_xpath).text != institution_name:
                        self.assertTrue(False, u"付款方显示的不是机构")
                    if browser.find_element_by_xpath(gather_customer_status).text != chain_customer:
                        self.assertTrue(False, u"收款方显示的不是链属企业")
                    if float(str(browser.find_element_by_xpath(loan_amount_xpath).text).replace(",","")) - float(
                                    xlSht.Cells(i, 4).Value - xlSht.Cells(i, 5).Value) != 0:
                        self.assertTrue(False, u"放款的金额不正确")
                    if browser.find_element_by_xpath(institution_name_xpath).text != institution_name:
                        self.assertTrue(False, u"机构名称不正确")
                    if browser.find_element_by_xpath(check_status_xpath).text != u"发起成功":
                        self.assertTrue(False, u"未发起成功")
                    if browser.find_element_by_xpath(check_operation_xpath).text!=u"通过":
                        self.assertTrue(False,u"操作状态显示未通过")
                    time.sleep(0.1)
                except NoSuchElementException, e:
                    browser.find_element_by_xpath(
                        ".//*[@id='pagebar']//a[text()='" + str(page_index) + "']").click()  # 点击下一页
                    time.sleep(2)
                    page_index = page_index + 1
                    browser.implicitly_wait(80)
                    ##################
                    # 断言每一条编号
                    ##################
                    payment_customer_xpath = ".//td[text()='" + package_id + "']/preceding-sibling::td[4]"
                    gather_customer_status = ".//td[text()='" + package_id + "']/preceding-sibling::td[3]"
                    loan_amount_xpath = ".//td[text()='" + package_id + "']/preceding-sibling::td[2]"
                    institution_name_xpath = ".//td[text()='" + package_id + "']/preceding-sibling::td[1]"
                    check_status_xpath = ".//td[text()='" + package_id + "']/following-sibling::td[1]"
                    check_operation_xpath = ".//td[text()='" + package_id + "']/following-sibling::td[3]"
                    try:
                        if browser.find_element_by_xpath(payment_customer_xpath).text != institution_name:
                            self.assertTrue(True, u"付款方显示的不是机构")
                        if browser.find_element_by_xpath(gather_customer_status).text != chain_customer:
                            self.assertTrue(True, u"收款方显示的不是链属企业")
                        if float(str(browser.find_element_by_xpath(loan_amount_xpath).text).replace(",","")) - float(
                                        xlSht.Cells(i, 4).Value - xlSht.Cells(i, 5).Value) != 0:
                            print "jinge bu dui"
                            self.assertTrue(True, u"放款的金额不正确")
                        if browser.find_element_by_xpath(institution_name_xpath).text != institution_name:
                            self.assertTrue(True, u"机构名称不正确")
                        if browser.find_element_by_xpath(check_status_xpath).text != u"发起成功":
                            self.assertTrue(False, u"未发起成功")
                        if browser.find_element_by_xpath(check_operation_xpath).text != u"通过":
                            self.assertTrue(False, u"操作状态显示未通过")
                        time.sleep(0.1)
                    except NoSuchElementException,e:
                        print package_id+u"这条已通过交易校验核对的单据号未找到"

            xlBook.Close(SaveChanges=1)
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
    def test_4_verify_loan(self):
        (u"验证放款")
        browser = self.browser
        browser.implicitly_wait(10)
        in_str=''
        in_str1=''
        cloum=2
        cloum_1=2
        cloum_2=2
        cloum_3=2
        page_index=1
        amount=0
        #####################################
        # 根据settle_id查询payment_id
        ####################################
        xlApp = win32com.client.Dispatch('Excel.Application')  # 打开EXCEL
        xlBook = xlApp.Workbooks.Open(r'D:\\Workspace\\Pythonscripts\\testdatas\\transaction_flow.xlsx')  # 读取交易流水
        xlSht = xlBook.Worksheets('Sheet2')
        # for i in range(2, xlSht.UsedRange.Rows.Count + 1):
        #     in_str = in_str + "'" + str(xlSht.Cells(cloum, 7)) + "'" + ","
        #     cloum = cloum + 1
        in_str ="'"+str(xlSht.Cells(cloum, 7))+"'"
        try:
            # 数据库连接
            conn = mysql.connector.connect(host=HOST, user=USER, passwd=PASSWORD, db=DATABASE1, port=PORT)
            # 创建游标
            cur = conn.cursor()
            # 根据settle_id查询流水
            sql = 'select payment_detail_id from t_settlement_order  where id IN (' + in_str + ')'
            print  sql
            cur.execute(sql)
            # 获取查询结果
            result_set = cur.fetchall()
            if result_set:
                for row in result_set:
                    payment_id = row[0]  # 从数据库取得id号
                    xlSht.Cells(cloum_1, 8).Value = str(payment_id)
                    cloum_1= cloum_1 + 1
            else:
                self.assertTrue(False, "the so_no do not exsit in database!")
                # 关闭游标和连接
            cur.close()
            conn.close()
        except mysql.connector.Error, e:
            print e.message
        xlBook.Close(SaveChanges=1)
        del xlApp
        time.sleep(1)
        ####################################################
        #根据payment_id查询BP
        ###################################################
        xlApp = win32com.client.Dispatch('Excel.Application')  # 打开EXCEL
        xlBook = xlApp.Workbooks.Open(r'D:\\Workspace\\Pythonscripts\\testdatas\\transaction_flow.xlsx')  # 读取交易流水
        xlSht = xlBook.Worksheets('Sheet2')
        # for i in range(2, xlSht.UsedRange.Rows.Count + 1):
        #     in_str1 = in_str1 + "'" + str(xlSht.Cells(cloum_2, 8)) + "'" + ","
        #     cloum_2 = cloum_2 + 1
        # in_str1 = in_str1[:-1]
        in_str1="'"+str(xlSht.Cells(cloum_2, 8))+"'"
        try:
            # 数据库连接
            conn = mysql.connector.connect(host=HOST, user=USER, passwd=PASSWORD, db=DATABASE2, port=PORT)
            # 创建游标
            cur = conn.cursor()
            # 根据settle_id查询流水
            sql = 'select id from t_payment_bank_order  where invokeid  IN (' + in_str1 + ')'
            print  sql
            cur.execute(sql)
            # 获取查询结果
            result_set = cur.fetchall()
            if result_set:
                for row in result_set:
                    bp_id = row[0]  # 从数据库取得id号
                    xlSht.Cells(cloum_3, 9).Value = str(bp_id)
                    cloum_3 = cloum_3 + 1
            else:
                self.assertTrue(False, "the so_no do not exsit in database!")
                # 关闭游标和连接
            cur.close()
            conn.close()
        except mysql.connector.Error, e:
            print e.message
        xlBook.Close(SaveChanges=1)
        del xlApp
        time.sleep(1)

        try:
            login.operate_login(self, "operation_login.csv")
            time.sleep(1)
            browser.find_element_by_link_text(u"群星支付").click()
            time.sleep(3)
            browser.find_element_by_xpath("//div[text()='凭证查询']").click()#凭证查询
            time.sleep(2)
            #########################################
            # 根据查询到的BP验证付款凭证
            #########################################
            browser.find_element_by_xpath("//li[@class='nav-list voucherPage']//a[text()='付款凭证查询']").click()#凭证查询
            time.sleep(2)
            print  str(bp_id)
            browser.find_element_by_id("payNum").send_keys(str(bp_id))
            time.sleep(2)
            browser.find_element_by_id("searchBtn").click()
            time.sleep(2)
            payment_status = ".//td[text()='" + str(bp_id) + "']/following-sibling::td[9]/span"
            wait_time=0
            while True:
                time.sleep(20)
                try:
                    if browser.find_element_by_xpath(payment_status).text==u"已付款":
                        break
                    else:
                        browser.refresh()
                        browser.find_element_by_id("payNum").clear()
                        browser.find_element_by_id("payNum").send_keys(str(bp_id))
                        browser.find_element_by_id("searchBtn").click()
                except NoSuchElementException,e:
                    self.assertTrue(False,u"付款编号为找到")
                wait_time=wait_time+1
                if wait_time==50:
                    self.assertTrue(False,u"未付款成功")
                    break

            xlApp = win32com.client.Dispatch('Excel.Application')  # 打开EXCEL
            xlBook = xlApp.Workbooks.Open(r'D:\\Workspace\\Pythonscripts\\testdatas\\transaction_flow.xlsx')  # 读取交易流水
            xlSht = xlBook.Worksheets('Sheet2')
            for i in range(2, xlSht.UsedRange.Rows.Count + 1):
                amount=amount+float(xlSht.Cells(i,4).Value)-float(xlSht.Cells(i,5).Value)#取得放款的总额
            xlBook.Close(SaveChanges=1)
            del  xlApp

            try:
                payment_xpath=".//td[text()='" +str(bp_id)+ "']/following-sibling::td[1]"
                receipt_xpath=".//td[text()='" +str(bp_id)+ "']/following-sibling::td[3]"
                amount_xpath=".//td[text()='" +str(bp_id)+ "']/following-sibling::td[5]"
                if browser.find_element_by_xpath(payment_xpath).text!=institution_name:
                    self.assertTrue(True,u"机构名称不正确")
                if browser.find_element_by_xpath(receipt_xpath).text!=chain_customer:
                    self.assertTrue(True,u"到款方不正确")
                if float(str(browser.find_element_by_xpath(amount_xpath).text).replace(",",""))-amount!=0:
                    self.assertTrue(True,u"结算凭证的金额不正确")
                time.sleep(0.1)
            except NoSuchElementException,e:
                self.assertTrue(False,u"付款编号未找到")

            ########################################
            #验证结算凭证
            ########################################
            time.sleep(2)
            browser.find_element_by_xpath("//div[text()='凭证查询']").click()  # 凭证查询
            time.sleep(2)
            browser.find_element_by_xpath("//li[@class='nav-list voucherPage']//a[text()='结算凭证查询']").click()  # 凭证查询
            time.sleep(2)
            Select(browser.find_element_by_xpath("//*[@class='settlementType form-control']")).select_by_value("2")#点击融资放款
            time.sleep(2)
            browser.find_element_by_xpath("//input[@placeholder='请输入付款方']").send_keys(institution_name)
            time.sleep(2)
            browser.find_element_by_id("searchBtn").click()#点击搜索
            wait_time=0

            while True:
                time.sleep(5)
                try:
                    if browser.find_element_by_xpath("//div[@class='loading']").is_displayed():
                        wait_time=wait_time+1
                    else:
                        break
                except NoSuchElementException,e:
                    break
                if wait_time==50:
                    break
            browser.find_element_by_xpath("//button[@id='pageSizeWraper']/following-sibling::button").click()  # 点击分页按钮
            time.sleep(1)
            browser.find_element_by_xpath(".//*[@id='pageSizeName']//a[text()='500']").click()  # 分页选择500页
            wait_time = 0
            while True:
                time.sleep(5)
                try:
                    if browser.find_element_by_xpath("//div[@class='loading']").is_displayed():
                        wait_time = wait_time + 1
                    else:
                        break
                except NoSuchElementException, e:
                    break
                if wait_time == 50:
                    break

            xlApp = win32com.client.Dispatch('Excel.Application')  # 打开EXCEL
            xlBook = xlApp.Workbooks.Open(r'D:\\Workspace\\Pythonscripts\\testdatas\\transaction_flow.xlsx')  # 读取交易流水
            xlSht = xlBook.Worksheets('Sheet2')
            for i in range(2, xlSht.UsedRange.Rows.Count + 1):
                settle_id = str(xlSht.Cells(i, 7).Value)
                try:
                    browser.implicitly_wait(10)
                    ##################
                    # 断言每一条编号
                    ##################
                    payment_customer_xpath = ".//td[text()='" + settle_id + "']/following-sibling::td[4]"
                    gather_customer_status = ".//td[text()='" + settle_id + "']/following-sibling::td[6]"
                    loan_amount_xpath = ".//td[text()='" + settle_id + "']/following-sibling::td[8]"
                    check_status_xpath = ".//td[text()='" + settle_id + "']/following-sibling::td[11]"
                    if browser.find_element_by_xpath(check_status_xpath).text != u"通过":
                        wait_time = 0
                        while True:
                            print "进入循环"
                            time.sleep(30)
                            browser.find_element_by_xpath("//input[@placeholder='请输入付款方']").clear()
                            browser.find_element_by_xpath("//input[@placeholder='请输入付款方']").send_keys(institution_name)
                            time.sleep(2)
                            browser.find_element_by_id("searchBtn").click()  # 点击搜索
                            ######################
                            #等待loading消失
                            #####################
                            wait_times=0
                            while True:
                                time.sleep(5)
                                try:
                                    if browser.find_element_by_xpath("//div[@class='loading']").is_displayed():
                                        wait_times = wait_times + 1
                                    else:
                                        break
                                except NoSuchElementException, e:
                                    break
                                if wait_times == 50:
                                    break
                            if browser.find_element_by_xpath(check_status_xpath).text == u"通过":
                                break
                            wait_time=wait_time+1
                            if wait_time==20:
                                self.assertTrue(False,u"一直处于结算中")
                                break
                    if browser.find_element_by_xpath(payment_customer_xpath).text != institution_name:
                        self.assertTrue(True, u"付款方显示的不是机构")
                    if browser.find_element_by_xpath(gather_customer_status).text != chain_customer:
                        self.assertTrue(True, u"收款方显示的不是链属企业")
                    if float(str(browser.find_element_by_xpath(loan_amount_xpath).text).replace(",","")) - float(
                                    xlSht.Cells(i, 4).Value - xlSht.Cells(i, 5).Value) != 0:
                        self.assertTrue(True, u"结算的金额不正确")

                except NoSuchElementException, e:
                    browser.find_element_by_xpath(
                        ".//*[@id='pagebar']//a[text()='" + str(page_index) + "']").click()  # 点击下一页
                    time.sleep(2)
                    page_index = page_index + 1
                    browser.implicitly_wait(80)
                    ##################
                    # 断言每一条编号
                    ##################
                    payment_customer_xpath = ".//td[text()='" + settle_id + "']/following-sibling::td[4]"
                    gather_customer_status = ".//td[text()='" + settle_id + "']/following-sibling::td[6]"
                    loan_amount_xpath = ".//td[text()='" + settle_id + "']/following-sibling::td[8]"
                    check_status_xpath = ".//td[text()='" + settle_id + "']/following-sibling::td[11]"
                    try:
                        if browser.find_element_by_xpath(check_status_xpath).text != u"通过":
                            wait_time = 0
                            while True:
                                time.sleep(30)
                                browser.find_element_by_xpath("//input[@placeholder='请输入付款方']").clear()
                                browser.find_element_by_xpath("//input[@placeholder='请输入付款方']").send_keys(institution_name)
                                time.sleep(2)
                                browser.find_element_by_id("searchBtn").click()  # 点击搜索
                                ######################
                                # 等待loading消失
                                #####################
                                wait_times = 0
                                while True:
                                    time.sleep(5)
                                    try:
                                        if browser.find_element_by_xpath("//div[@class='loading']").is_displayed():
                                            wait_times = wait_times + 1
                                        else:
                                            break
                                    except NoSuchElementException, e:
                                        break
                                    if wait_times == 50:
                                        break
                                if browser.find_element_by_xpath(check_status_xpath).text == u"通过":
                                    break
                                wait_time = wait_time + 1
                                if wait_time == 20:
                                    self.assertTrue(False, u"一直处于结算中")
                                    break
                        if browser.find_element_by_xpath(payment_customer_xpath).text != institution_name:
                            self.assertTrue(True, u"付款方显示的不是机构")
                        if browser.find_element_by_xpath(gather_customer_status).text != chain_customer:
                            self.assertTrue(True, u"收款方显示的不是链属企业")
                        if float(str(browser.find_element_by_xpath(loan_amount_xpath).text).replace(",","")) - float(
                                        xlSht.Cells(i, 4).Value - xlSht.Cells(i, 5).Value) != 0:
                            self.assertTrue(True, u"结算的金额不正确")
                    except NoSuchElementException,e:
                        print u"已通过交易校验核对的单据号未找到"
            xlBook.Close(SaveChanges=1)
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

    def test_5_login_chain_verify(self):
        (u"登录链属验证是否放款成功")
        browser = self.browser
        browser.implicitly_wait(10)
        login.corp_login(self,"chain_customer.xlsx")
        ##############
        #
        ##############
        start_time=str(time.strftime("%Y/%m/%d", time.localtime()))
        year = int(time.strftime("%Y", time.localtime()))
        mount = int(time.strftime("%m", time.localtime()))
        day = int(time.strftime("%d", time.localtime()))
        monthRange = calendar.monthrange(year, mount)
        financing_days = monthRange[1] - day + 31
        d1 = datetime.datetime.now()
        d3 = d1 + datetime.timedelta(days=financing_days)#获取financing_days后的日期
        end_time = str(d3.strftime("%Y/%m/%d"))
        browser.find_element_by_id("today").click()#点击今天按钮
        wait_time=0
        while True:
            time.sleep(5)
            try:
                if browser.find_element_by_xpath("//div[@class='loading']").is_displayed():
                    wait_time=wait_time+1
                else:
                    break
            except NoSuchElementException,e:
                break
            if wait_time==50:
                break
        browser.find_element_by_xpath("//a[text()='已融资']").click()#点击已申请融资
        while True:
            time.sleep(5)
            try:
                if browser.find_element_by_xpath("//div[@class='loading']").is_displayed():
                    wait_time = wait_time + 1
                else:
                    break
            except NoSuchElementException, e:
                break
            if wait_time == 50:
                break

        xlApp = win32com.client.Dispatch('Excel.Application')  # 打开EXCEL
        xlBook = xlApp.Workbooks.Open(r'D:\\Workspace\\Pythonscripts\\testdatas\\transaction_flow.xlsx')  # 读取交易流水
        xlSht = xlBook.Worksheets('Sheet2')
        for i in range(2, xlSht.UsedRange.Rows.Count + 1):
            loan_document_no=str(xlSht.Cells(i, 2))
            payment_customer_xpath = ".//td[text()='" +loan_document_no+ "']/following-sibling::td[1]"
            amount_xpath=".//td[text()='" +loan_document_no+ "']/following-sibling::td[2]"
            start_time_xpath=".//td[text()='" +loan_document_no+ "']/following-sibling::td[3]"
            end_time_xpath=".//td[text()='" +loan_document_no+ "']/following-sibling::td[4]"
            financing_amount_xpath=".//td[text()='" +loan_document_no+ "']/following-sibling::td[5]"
            try:
                if browser.find_element_by_xpath(payment_customer_xpath).text!=core_customer:
                    self.assertTrue(True,u"已融资单据的客户名称显示不正确")
                if float(str(browser.find_element_by_xpath(amount_xpath).text).replace(",",""))-float(xlSht.Cells(i, 3))!=0:
                    self.assertTrue(True,u"已融资单据金额不正确")
                if browser.find_element_by_xpath(start_time_xpath).text!=start_time:
                    self.assertTrue(True,u"融资发放日不正确")
                if browser.find_element_by_xpath(end_time_xpath).text!=end_time:
                    self.assertTrue(True,u"到期日不正确")
                if float(str(browser.find_element_by_xpath(financing_amount_xpath).text).replace(",",""))-float(xlSht.Cells(i, 4))!=0:
                    self.assertTrue(True,u"融资金额不正确")
                time.sleep(0.1)
            except NoSuchElementException,e:
                self.assertTrue(False,u"放款成功后，单据号没有出现在已放款模块")



















                # @classmethod
    # def tearDownClass(cls):
    #     cls.browser.close()
    #     cls.browser.quit()

