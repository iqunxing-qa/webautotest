#coding=utf-8
from unittest.test import test_suite
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import ElementNotVisibleException
from classmethod import getprofile
from classmethod import login
from classmethod import findStr
from selenium.webdriver.common.action_chains import ActionChains
import  datetime
import  time
import unittest
import HTMLTestRunner
import sys
import os
import StringIO
import traceback
reload(sys)
sys.setdefaultencoding('utf8')
import csv
import win32com.client
import ConfigParser
import mysql.connector
import re
import calendar
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
DATABASE=cf.get('database','dcf_loan')
#读取合同配置信息
xlApp = win32com.client.Dispatch('Excel.Application')  # 打开EXCEL
xlBook = xlApp.Workbooks.Open(r'D:\\workspace\\Pythonscripts\\testdatas\\contract_information.xlsx')
xlSht = xlBook.Worksheets('Sheet2')
loan_proportion=float(xlSht.Cells(3, 7).Value)/100
xlBook.Close(SaveChanges=1)
del xlApp
#读取合同信息
xlApp = win32com.client.Dispatch('Excel.Application')  # 打开EXCEL
xlBook = xlApp.Workbooks.Open(r'D:\\workspace\\Pythonscripts\\testdatas\\hetong.xlsx')
xlSht = xlBook.Worksheets(u'额度账期参数')
financing_proportion=float(xlSht.Cells(2, 4).Value)/100  # 读取链属企业名称
xlBook.Close(SaveChanges=1)
del xlApp
#读取链属企业
xlApp = win32com.client.Dispatch('Excel.Application')  # 打开EXCEL
xlBook = xlApp.Workbooks.Open(r'D:\\workspace\\Pythonscripts\\testdatas\\chain_customer.xlsx')  # 将随机生成的名称写入链属企业
xlSht = xlBook.Worksheets('Sheet1')
pattern = re.compile(r'\d*')
chain_customer=xlSht.Cells(2, 1).Value  # 读取链属企业名称
cell_phone_no=re.search(pattern,str(xlSht.Cells(2, 6).Value)).group()
xlBook.Close(SaveChanges=1)
del xlApp
#读取核心企业
xlApp = win32com.client.Dispatch('Excel.Application')  # 打开EXCEL
xlBook = xlApp.Workbooks.Open(r'D:\\workspace\\Pythonscripts\\testdatas\\core_customer.xlsx')  # 将随机生成的名称写入链属企业
xlSht = xlBook.Worksheets('Sheet1')
core_customer=xlSht.Cells(2, 1).Value  # 读取核心企业名称
xlBook.Close(SaveChanges=1)
del xlApp
#读取截图存放路径
shot_path=cf.get('shotpath','path')
class build_transaction_flow(unittest.TestCase):
    (u"新建流水模块")
    loan_no_list = []
    monkey_list=[]
    @classmethod
    def setUpClass(cls):
        xlApp = win32com.client.Dispatch('Excel.Application')  # 打开EXCEL
        xlBook = xlApp.Workbooks.Open(r'D:\\workspace\\Pythonscripts\\testdatas\\transaction_flow.xlsx')  # 将D:\\1.xls改为要处理的excel文件路径
        xlSht = xlBook.Worksheets('sheet1')
        lcoal_time = str(time.strftime("%Y/%m/%d", time.localtime()))
        cls.start_time=lcoal_time
        for i in range(2,20):
            loan_document_no = "aaRYX"+str(time.strftime("%m%d%H%M%S", time.localtime()))+str(i)
            xlSht.Cells(i, 1).Value = loan_document_no  # 随机生成单据编号
            xlSht.Cells(i, 2).Value = chain_customer  # 随机生成单据编号
            xlSht.Cells(i, 3).Value =u"应付" # 应付
            xlSht.Cells(i, 4).Value =150  # 应付
            build_transaction_flow.monkey_list.append(150)
            xlSht.Cells(i, 5).Value = lcoal_time  # 修改单据起始时间
            xlSht.Cells(i, 7).Value =u"测试流水"  # 修改单据起始时间
            build_transaction_flow.loan_no_list.append(loan_document_no)
            time.sleep(1)
        xlBook.Close(SaveChanges=1)  # 完成 关闭保存文件
        del xlApp
        cls.browser=webdriver.Firefox(profile)
        cls.browser.maximize_window()
    def test_1_upload_transaction(self):
        (u"上传交易流水")
        browser=self.browser
        browser.implicitly_wait(10)
        try:
            # 登录运营平台
            login.corp_login(self, "core_customer.xlsx")
            time.sleep(1)
            #针对第一次登录要
            try:
                if browser.find_element_by_xpath("html/body/div[2]/div[1]").is_displayed():
                    browser.find_element_by_xpath("html/body/div[2]/div[1]").click()
            except NoSuchElementException,e:
                print ""
            # if browser.find_element_by_id("addDashBtn").is_displayed():
            browser.execute_script("arguments[0].click()",browser.find_element_by_id("addDashBtn"))#点击新建流水
            time.sleep(1)
            ###########################################
            #            第一次使用安装数字证书       #
            ###########################################
            browser.implicitly_wait(5)
            try:
                browser.execute_script("arguments[0].click()",browser.find_element_by_id("getDy"))#获取验证码
                time.sleep(5)
                now_handle=browser.current_window_handle#获取当前的handle
                Dynamic_url = "http://" + host + ".dcfservice.com/v1/public/sms/get?cellphone="+cell_phone_no#获取验证码路径
                js_script = 'window.open(' + '"' + Dynamic_url + '"' + ')'
                browser.execute_script(js_script)
                time.sleep(2)
                all_handles = browser.window_handles
                for handle in all_handles:
                    if handle != now_handle:
                        browser.switch_to_window(handle)
                Dynamic_code = browser.find_element_by_css_selector("html>body>pre").text
                Dynamic_code = Dynamic_code[1:7]#截取字符串获取验证码
                browser.close()
                browser.switch_to_window(now_handle)#切换回以前handle
                browser.find_element_by_id("dyCode").send_keys(Dynamic_code)
                time.sleep(2)
                browser.find_element_by_id("validateDy").click()
                time.sleep(2)
                browser.find_element_by_id("installCfca").click()#点击立即安装
                time.sleep(10)
                browser.execute_script("arguments[0].click()", browser.find_element_by_id("addDashBtn"))  # 点击新建流水
                ###########未写完
            except NoSuchElementException,e:
                print "The customer has installed security controls "
             #######################################################################
            browser.implicitly_wait(10)#恢复隐式查找10S时间
            time.sleep(2)
            browser.find_element_by_xpath(".//*[@id='uploadArea']/div[1]/div[1]/span[1]").click()#点击上传文件
            time.sleep(1)
            upload_file = method + "\\upload.exe " + data + "transaction_flow.xlsx"
            os.system( upload_file)
            time.sleep(2)
            browser.execute_script("arguments[0].click()",browser.find_element_by_id("submit-now"))
            upload_flag=True
            wait_num=0
            while upload_flag:
                time.sleep(10)
                try:
                    if browser.find_element_by_css_selector("#fileuploadError>b").text==u"验签失败 ,不是证书用户，请先申请证书":
                        browser.get_screenshot_as_file(shot_path + u"验证证书失败.png")  # 对错误增加截图
                        upload_flag=False
                        self.assertFalse(True,"验签失败 ,不是证书用户，请先申请证书")
                except NoSuchElementException,e:
                    self.assertTrue(True)
                try:
                    browser.execute_script("arguments[0].click()",browser.find_element_by_xpath(".//*[@id='uploadArea']/div[3]/a"))#解析成功后点击取消
                    upload_flag=False
                except NoSuchElementException,e:
                    print  ""
                wait_num=wait_num+1
                if wait_num==20:
                    upload_flag=False
                    self.assertFalse(True, "上传交易流水时间过长")
            browser.refresh()#上传流水完成后刷新网页
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
    def test_2_check_core_transaction(self):
        (u"产看核心页面流水是否新建成功")
        browser=self.browser
        browser.implicitly_wait(10)
        seller_name=chain_customer #获取上传交易流水excel的卖家名称
        year = int(time.strftime("%Y", time.localtime()))
        mount = int(time.strftime("%m", time.localtime()))
        day= int(time.strftime("%d", time.localtime()))
        monthRange = calendar.monthrange(year, mount)
        financing_days=monthRange[1]-day+31
        start_time=self.start_time#获取交易流水鈤日期
        cloum=2
        cloum_1=2
        cloum_2 = 2
        try:
            ##############################################################
            #                 根据生产的单据号查询id                     #
            ##############################################################
            try:
                # 数据库连接
                conn = mysql.connector.connect(host=HOST, user=USER, passwd=PASSWORD, db=DATABASE, port=PORT)
                # 创建游标
                cur = conn.cursor()
                #产生的单据号写入transcation_flow.xlsx
                xlApp = win32com.client.Dispatch('Excel.Application')  # 打开EXCEL
                xlBook = xlApp.Workbooks.Open(
                    r'D:\\workspace\\Pythonscripts\\testdatas\\transaction_flow.xlsx')  # 将随机生成的名称写入链属企业
                xlSht = xlBook.Worksheets('Sheet2')
                for monkey in build_transaction_flow.monkey_list:
                    xlSht.Cells(cloum_2, 3).Value =monkey
                    cloum_2 = cloum_2 + 1
                in_str=''
                order_by_str=''
                for loan_document_no in build_transaction_flow.loan_no_list:
                    in_str=in_str+"'"+str(loan_document_no)+"'"+","
                    order_by_str=order_by_str+str(loan_document_no)+","
                    xlSht.Cells(cloum_1, 2).Value = str(loan_document_no)
                    cloum_1=cloum_1+1
                in_str=in_str[:-1]
                order_by_str=order_by_str[:-1]#去掉最后一个逗号
                order_by_str="'"+order_by_str+"'"
                # customername_id查询
                sql = 'select loan_document_id from t_loan_document where loan_document_no IN (' +in_str+')'+'ORDER BY FIND_IN_SET (loan_document_no,'+order_by_str+')'
                print  sql
                cur.execute(sql)
                # 获取查询结果
                result_set = cur.fetchall()
                if result_set:
                    for row in result_set:
                        loan_document_id = row[0]#从数据库取得id号
                        xlSht.Cells(cloum,1).Value = str(loan_document_id)
                        cloum = cloum + 1
                else:
                    self.assertTrue(False, "the loan_document_id do not exsit in database!")
                # 关闭游标和连接
                cur.close()
                conn.close()
                xlBook.Close(SaveChanges=1)  # 完成 关闭保存文件
                del xlApp
            except mysql.connector.Error, e:
                print e.message
            ##############################
            #对xlsx单据号进行循环断言查询
            #############################
            xlApp = win32com.client.Dispatch('Excel.Application')  # 打开EXCEL
            xlBook = xlApp.Workbooks.Open(
                r'D:\\workspace\\Pythonscripts\\testdatas\\transaction_flow.xlsx')
            xlSht = xlBook.Worksheets('Sheet2')
            for i in range(2,xlSht.UsedRange.Rows.Count+1):
                amount=float(xlSht.Cells(i,3).Value)
                loan_document_no_xpath='//*[@id="'+ str(xlSht.Cells(i,1).Value)+'"]/td[2]'
                seller_name_xpath='//*[@id="'+str(xlSht.Cells(i,1).Value)+'"]/td[3]'
                amount_xpath='//*[@id="'+str(xlSht.Cells(i,1).Value)+'"]/td[4]'
                start_time_xpath = '//*[@id="' + str(xlSht.Cells(i,1).Value) + '"]/td[5]'
                financing_days_xpath='//*[@id="' + str(xlSht.Cells(i,1).Value) + '"]/td[7]'
                financing_cost_xpath='//*[@id="' + str(xlSht.Cells(i,1).Value) + '"]/td[8]'
                ################断言填写#################################
                if browser.find_element_by_xpath(loan_document_no_xpath).text!=str(xlSht.Cells(i,2).Value):#单据号断言
                    print browser.find_element_by_xpath(loan_document_no_xpath).text
                    print str(xlSht.Cells(i,2).Value)
                    browser.get_screenshot_as_file(shot_path + u"单据号不一致" + ".png")#对错误增加截图
                    self.assertFalse(True,"Transaction document No. is inconsistent with EXCEL")
                if browser.find_element_by_xpath(seller_name_xpath).text!=seller_name:#客户名称断言
                    print browser.find_element_by_xpath(seller_name_xpath).text
                    print  seller_name
                    browser.get_screenshot_as_file(shot_path + u"卖家名称不一致"+ ".png")#对错误增加截图
                    self.assertFalse(True,"customer_name is inconsistent with EXCEL")
                if float(str(browser.find_element_by_xpath(amount_xpath).text).replace(",",""))-amount!=0:#单据金额断言
                    print float(browser.find_element_by_xpath(amount_xpath).text)
                    print amount
                    browser.get_screenshot_as_file(shot_path + u"上传金额不一致" + ".png")#对错误增加截图
                    self.assertFalse(True,"amount is inconsistent with EXCEL")
                if browser.find_element_by_xpath(start_time_xpath).text!=start_time:#单据起始日期断言
                    print browser.find_element_by_xpath(start_time_xpath).text
                    print  start_time
                    browser.get_screenshot_as_file(shot_path + u"起始日不一致" + ".png")#对错误增加截图
                    self.assertFalse(True, "the start_time of document is inconsistent with EXCEL")
                if float(str(browser.find_element_by_xpath(financing_days_xpath).text).replace(",",""))-financing_days!=0:#可融资天数断言
                    print float(browser.find_element_by_xpath(financing_days_xpath).text)
                    print  financing_days
                    browser.get_screenshot_as_file(shot_path + u"可融资天数计算不正确" + ".png")  # 对错误增加截图
                    self.assertFalse(True, "the financing_days is wrong")
                if float(browser.find_element_by_xpath(financing_cost_xpath).text)!= 0:  # 可融资天数断言
                    print float(browser.find_element_by_xpath(financing_cost_xpath).text)
                    browser.get_screenshot_as_file(shot_path + u"融资成本不正确" + ".png")  # 对错误增加截图
                    self.assertFalse(True, "the financing_cost is wrong")
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
    def test_3_check_chain_transaction(self):
        (u"查看链属企业融资单据")
        browser = self.browser
        browser.implicitly_wait(10)
        year = int(time.strftime("%Y", time.localtime()))
        mount = int(time.strftime("%m", time.localtime()))
        day= int(time.strftime("%d", time.localtime()))
        monthRange = calendar.monthrange(year, mount)
        financing_days=monthRange[1]-day+31
        start_time=self.start_time
        buyer_name =core_customer
        try:
            login.corp_login(self,"chain_customer.xlsx")#链属企业登录
            time.sleep(1)
            try:
                browser.find_element_by_xpath("//div[@class='index-bg-1-close']").click()
                time.sleep(2)
                browser.find_element_by_xpath(".//div[@class='index-bg-2-close']").click()
                time.sleep(2)
            except NoSuchElementException,e:
                print ""
            xlApp = win32com.client.Dispatch('Excel.Application')  # 打开EXCEL
            xlBook = xlApp.Workbooks.Open(r'D:\\workspace\\Pythonscripts\\testdatas\\transaction_flow.xlsx')  # 将随机生成的名称写入链属企业
            xlSht = xlBook.Worksheets('Sheet2')
            for i in range(2,xlSht.UsedRange.Rows.Count+1):
                amount=float(xlSht.Cells(i,3).Value)
                financing_amount = round(amount * financing_proportion, 2)  # 融资金额
                financing_cost = round(financing_amount * loan_proportion, 2)  # 预计融资成本
                xlSht.Cells(i, 4).Value=financing_amount
                xlSht.Cells(i, 5).Value = financing_cost
                loan_document_no_xpath='//*[@id="'+ str(xlSht.Cells(i,1).Value)+'"]/td[2]'
                seller_name_xpath='//*[@id="'+str(xlSht.Cells(i,1).Value)+'"]/td[3]'
                amount_xpath='//*[@id="'+str(xlSht.Cells(i,1).Value)+'"]/td[4]'
                start_time_xpath = '//*[@id="' + str(xlSht.Cells(i,1).Value) + '"]/td[5]'
                financing_days_xpath='//*[@id="' + str(xlSht.Cells(i,1).Value) + '"]/td[7]'
                financing_cost_xpath='//*[@id="' + str(xlSht.Cells(i,1).Value) + '"]/td[8]'
                ################断言填写#################################
                if browser.find_element_by_xpath(loan_document_no_xpath).text!=str(xlSht.Cells(i,2).Value):#单据号断言
                    browser.get_screenshot_as_file(shot_path + u"单据号不一致" + ".png")#对错误增加截图
                    self.assertFalse(True,"Transaction document No. is inconsistent with EXCEL")
                if browser.find_element_by_xpath(seller_name_xpath).text!=buyer_name:#客户名称断言
                    browser.get_screenshot_as_file(shot_path + u"买家名称不一致"+ ".png")#对错误增加截图
                    self.assertFalse(True,"customer_name is inconsistent with EXCEL")
                if float(str(browser.find_element_by_xpath(amount_xpath).text).replace(",",""))-amount!=0:#单据金额断言
                    browser.get_screenshot_as_file(shot_path + u"上传金额不一致" + ".png")#对错误增加截图
                    self.assertFalse(True,"amount is inconsistent with EXCEL")
                if browser.find_element_by_xpath(start_time_xpath).text!=start_time:#单据起始日期断言
                    browser.get_screenshot_as_file(shot_path + u"起始日不一致" + ".png")#对错误增加截图
                    self.assertFalse(True, "the start_time of document is inconsistent with EXCEL")
                if float(browser.find_element_by_xpath(financing_days_xpath).text)-financing_days!=0:#可融资天数断言
                    browser.get_screenshot_as_file(shot_path + u"可融资天数计算不正确" + ".png")  # 对错误增加截图
                    self.assertFalse(True, "the financing_days is wrong")
                if float(str(browser.find_element_by_xpath(financing_cost_xpath).text).replace(",",""))-financing_cost!= 0:  # 可融资天数断言
                    browser.get_screenshot_as_file(shot_path + u"融资成本不正确" + ".png")  # 对错误增加截图
                    self.assertFalse(True, "the financing_cost is wrong")
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
    # @classmethod
    # def tearDownClass(cls):
    #     cls.browser.close()
    #     cls.browser.quit()

