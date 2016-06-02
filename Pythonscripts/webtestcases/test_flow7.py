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
import  random
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
#读取截图存放路径
shot_path=cf.get('shotpath','path')
class core_contract(unittest.TestCase):
    (u"新建流水模块")
    @classmethod
    def setUpClass(cls):
        xlApp = win32com.client.Dispatch('Excel.Application')  # 打开EXCEL
        xlBook = xlApp.Workbooks.Open(
            r'D:\\workspace\\Pythonscripts\\testdatas\\transaction_flow.xlsx')  # 将D:\\1.xls改为要处理的excel文件路径
        xlSht = xlBook.Worksheets('sheet1')
        lcoal_time = str(time.strftime("%Y/%m/%d", time.localtime()))
        loan_document_no = "aaRYX" + str(random.randrange(1, 100000))
        # 将随机生成的单据编号写入random_loan_no.csv中
        csv_random_loan = file(data + 'random_loan_no.csv', 'wb')
        writer = csv.writer(csv_random_loan)
        writer.writerow([loan_document_no])
        csv_random_loan.close()
        xlSht.Cells(2, 5).Value = lcoal_time  # 修改单据起始时间
        xlSht.Cells(2, 1).Value = loan_document_no  # 随机生成单据编号
        cls.seller_name = xlSht.Cells(2, 2).Value
        cls.loan_document_no=xlSht.Cells(2, 1).Value
        cls.amount=xlSht.Cells(2, 4).Value
        cls.start_time=lcoal_time
        xlBook.Close(SaveChanges=1)  # 完成 关闭保存文件
        del xlApp
        cls.browser=webdriver.Firefox(profile)
        cls.browser.maximize_window()
    def test_1_upload_transaction(self):
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
            # if browser.find_element_by_id("addDashBtn").is_displayed():
            browser.execute_script("arguments[0].click()",browser.find_element_by_id("addDashBtn"))#点击新建流水
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
                print "The customer has installed security controls "
             #######################################################################
            browser.implicitly_wait(10)#恢复隐式查找10S时间
            browser.find_element_by_xpath(".//*[@id='uploadArea']/div[1]/div[1]/span[1]").click()#点击上传文件
            time.sleep(1)
            upload_file = method + "\\upload.exe " + data + "transaction_flow.xlsx"
            os.system( upload_file)
            time.sleep(2)
            browser.execute_script("arguments[0].click()",browser.find_element_by_id("submit-now"))
            time.sleep(10)
            browser.execute_script("arguments[0].click()",browser.find_element_by_xpath(".//*[@id='uploadArea']/div[3]/a"))#解析成功后点击取消
            time.sleep(1)
            browser.refresh()#上传流水完成后刷新网页
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
    def test_2_check_core_transaction(self):
        (u"产看核心页面流水是否新建成功")
        browser=self.browser
        browser.implicitly_wait(10)
        seller_name=self.seller_name.strip() #获取上传交易流水excel的卖家名称
        loan_document_no=self.loan_document_no.strip() #获取交易流水的单据号
        amount=self.amount #获取交易流水的金额
        start_time=self.start_time#获取交易流水起始日期

        try:
            ##############################################################
            #                 根据生产的单据号查询id                     #
            ##############################################################
            try:
                # 数据库连接
                conn = mysql.connector.connect(host=HOST, user=USER, passwd=PASSWORD, db=DATABASE, port=PORT)
                # 创建游标
                cur = conn.cursor()
                # customername_id查询
                sql = 'select loan_document_id from t_loan_document where loan_document_no="' +loan_document_no+ '"'
                cur.execute(sql)
                # 获取查询结果
                result_set = cur.fetchall()
                if result_set:
                    for row in result_set:
                        loan_document_id = row[0]#从数据库取得id号

                else:
                    self.assertTrue(False, "the customer_id do not exsit in database!")
                # 关闭游标和连接
                cur.close()
                conn.close()
            except mysql.connector.Error, e:
                print e.message
            loan_document_id=str(loan_document_id)
            #产生的单据号写入loan_document_id.csv
            loan_document = file(data + 'loan_document_id.csv', 'wb')
            writer = csv.writer(loan_document)
            writer.writerow([loan_document_id])
            loan_document.close()#结束写入
            loan_document_no_xpath='//*[@id="'+loan_document_id+'"]/td[2]'
            seller_name_xpath='//*[@id="'+loan_document_id+'"]/td[3]'
            amount_xpath='//*[@id="'+loan_document_id+'"]/td[4]'
            start_time_xpath = '//*[@id="' + loan_document_id + '"]/td[5]'
            if browser.find_element_by_xpath(loan_document_no_xpath).text!=loan_document_no:
                title_index = browser.title.find("-")
                title = browser.title[0:title_index]
                browser.get_screenshot_as_file(shot_path + title + ".png")#对错误增加截图
                self.assertFalse(True,"Transaction document No. is inconsistent with EXCEL")
            time.sleep(1)
            if browser.find_element_by_xpath(seller_name_xpath).text!=seller_name:
                title_index = browser.title.find("-")
                title = browser.title[0:title_index]
                browser.get_screenshot_as_file(shot_path + title + ".png")#对错误增加截图
                self.assertFalse(True,"customer_name is inconsistent with EXCEL")
            time.sleep(1)
            if float(browser.find_element_by_xpath(amount_xpath).text)!=amount:
                title_index = browser.title.find("-")
                title = browser.title[0:title_index]
                browser.get_screenshot_as_file(shot_path + title + ".png")#对错误增加截图
                self.assertFalse(True,"amount is inconsistent with EXCEL")
            time.sleep(1)
            if browser.find_element_by_xpath(start_time_xpath).text!=start_time:
                title_index = browser.title.find("-")
                title = browser.title[0:title_index]
                browser.get_screenshot_as_file(shot_path + title + ".png")#对错误增加截图
                self.assertFalse(True, "the start_time of document is inconsistent with EXCEL")
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

    def test_3_check_chain_transaction(self):
        (u"查看链属企业融资单据")
        browser = self.browser
        browser.implicitly_wait(10)
        loan_document_no=self.loan_document_no
        amount=self.amount
        start_time=self.start_time
        financing_cost=0.88
        try:
            login.corp_login(self,"chain_enterprise_customer.csv")#链属企业登录
            time.sleep(1)
            #读取loan_document_id.csv中的单据编号
            csvfile = file(data + 'loan_document_id.csv', 'rb')
            reader = csv.reader(csvfile)
            for line in reader:
                loan_document_id=line[0].decode('utf-8')
            # 读取产生的核心企业
            csvfile = file(data + 'core_enterprise_login.csv', 'rb')
            reader = csv.reader(csvfile)
            for line in reader:
                customer_name = line[0].decode('utf-8')
            loan_document_no_xpath = '//*[@id="' + loan_document_id + '"]/td[2]'
            customer_name_xpath='//*[@id="'+loan_document_id+'"]/td[3]'
            amount_xpath='//*[@id="'+loan_document_id+'"]/td[4]'
            start_time_xpath = '//*[@id="' + loan_document_id + '"]/td[5]'
            financing_cost_xpath='//*[@id="' + loan_document_id + '"]/td[8]'
            if browser.find_element_by_xpath(loan_document_no_xpath).text!=loan_document_no:#判断单据号是否相等
                title_index = browser.title.find("-")
                title = browser.title[0:title_index]
                browser.get_screenshot_as_file(shot_path + title + ".png")#对错误增加截图
                self.assertFalse(True,"Transaction document No. is inconsistent with EXCEL")
            time.sleep(1)
            if browser.find_element_by_xpath(customer_name_xpath).text!=customer_name:#判断客户名称是否相同
                title_index = browser.title.find("-")
                title = browser.title[0:title_index]
                browser.get_screenshot_as_file(shot_path + title + ".png")#对错误增加截图
                self.assertFalse(True,"customer_name is inconsistent with EXCEL")
            time.sleep(1)
            if float(browser.find_element_by_xpath(amount_xpath).text)!=amount:#判断融资金额是否相等
                title_index = browser.title.find("-")
                title = browser.title[0:title_index]
                browser.get_screenshot_as_file(shot_path + title + ".png")#对错误增加截图
                self.assertFalse(True,"amount is inconsistent with EXCEL")
            time.sleep(1)
            if browser.find_element_by_xpath(start_time_xpath).text!=start_time:#判断单据开始时间是否相同
                title_index = browser.title.find("-")
                title = browser.title[0:title_index]
                browser.get_screenshot_as_file(shot_path + title + ".png")#对错误增加截图
                self.assertFalse(True, "the start_time of document is inconsistent with EXCEL")
            time.sleep(1)
            if float(browser.find_element_by_xpath(financing_cost_xpath).text)!=financing_cost:  # 判断单据的融资成本是否计算正确
                title_index = browser.title.find("-")
                title = browser.title[0:title_index]
                browser.get_screenshot_as_file(shot_path + title + ".png")  # 对错误增加截图
                self.assertFalse(True, "the calculate of financing_cost is wrong!")
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
    @classmethod
    def tearDownClass(cls):
        cls.browser.close()
        cls.browser.quit()

