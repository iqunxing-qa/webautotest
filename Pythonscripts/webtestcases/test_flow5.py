#coding=utf-8
from selenium import webdriver
import time
import unittest
import ConfigParser
import  StringIO
import traceback
from selenium.webdriver.support.ui import Select
from classmethod import *
from selenium.common.exceptions import NoSuchElementException
from classmethod import login
import mysql.connector
import win32com.client
import sys
reload(sys)
import re
sys.setdefaultencoding('utf8')
cf = ConfigParser.ConfigParser()
cf.read(r"D:\Workspace\Pythonscripts\environment\env.conf")
host=cf.get('service','host')
method=cf.get('dir','method')
data=cf.get('dir','data')
#读取数据库文件
USER=cf.get('database','user')
HOST=cf.get('database','host')
PASSWORD=cf.get('database','password')
PORT=cf.get('database','port')
DATABASE=cf.get('database','dcf_contract')
#读取截图存放路径
shot_path=cf.get('shotpath','path')

#读取机构名和id
xlxApp = win32com.client.Dispatch('Excel.Application')  # 打开EXCEL
xlxBook=xlxApp.Workbooks.Open(r'D:\\Workspace\\Pythonscripts\\testdatas\\institution_data.xlsx')
xlSht = xlxBook.Worksheets('sheet1')
jihou_name = xlSht.Cells(2, 1).Value
customername_id = xlSht.Cells(2,7).Value
xlxBook.Close(SaveChanges=1)
del xlxApp

#定义方案名称并写入product_configuration.xlsx
a=str(time.strftime("%m%d%H%M%S", time.localtime()))
solution_name=((u'群星测试')+a)
xlxApp = win32com.client.Dispatch('Excel.Application')  # 打开EXCEL
xlxBook=xlxApp.Workbooks.Open(r'D:\\Workspace\\Pythonscripts\\testdatas\\product_configuration.xlsx')
xlSht = xlxBook.Worksheets('sheet3')
xlSht.Cells(2,1).Value=solution_name
time.sleep(2)

#读取product_id,product_type
xlSht = xlxBook.Worksheets('sheet1')
product_type = xlSht.Cells(2, 2).Value
product_id = xlSht.Cells(2,3).Value
pattern = re.compile(r'\d*')
product_id= re.search(pattern, str(xlSht.Cells(2, 3).Value)).group()

#读取模板文件名称
xlSht = xlxBook.Worksheets('sheet2')
protocol_document= xlSht.Cells(2, 3).Value
control_document= xlSht.Cells(2,4).Value
contract_document= xlSht.Cells(2,5).Value
agency_id=xlSht.Cells(2,2).Value
pattern = re.compile(r'\d*')
agency_id= re.search(pattern, str(xlSht.Cells(2, 2).Value)).group()
xlxBook.Close(SaveChanges=1)
del xlxApp
#获取Firefox的profile
propath=getprofile.get_profile()
profile=webdriver.FirefoxProfile(propath)
class Core_Enterprise(unittest.TestCase):
    (u"核心模块")
    @classmethod
    def setUpClass(cls):
        cls.browser = webdriver.Firefox(profile)
        cls.browser.maximize_window()

    def test_Create_program(self):
        (u"新建方案")
        browser = self.browser
        browser.implicitly_wait(10)
        try:
            login.operate_login(self,'operation_login.csv') #登陆
            time.sleep(2)
            browser.find_element_by_link_text(u"产品配置").click()
            time.sleep(2)
            browser.find_element_by_link_text(u"方案配置").click()
            time.sleep(2)
            browser.find_element_by_id('new-program').click()
            time.sleep(2)
            browser.find_element_by_id('product').click()
            time.sleep(2)
            browser.find_element_by_xpath("//select[@id='product']/option[@value='"+product_id+"']").click()#产品名
            browser.find_element_by_id('institution-process').click()
            time.sleep(2)
            browser.find_element_by_xpath("//select[@id='institution-process']/option[@value="+agency_id+"]").click()#机构工作方式
            browser.find_element_by_id('org').click()
            time.sleep(2)
            browser.find_element_by_xpath("//select[@id='org']/option[@value='"+customername_id+"']").click()#确定机构
            time.sleep(2)
            browser.find_element_by_name('program').send_keys(solution_name)
            browser.find_element_by_id('first').click()
            time.sleep(2)
            browser.find_element_by_xpath("//select[@id='first']/option[@value=2]").click()#机构对卖家
            browser.find_element_by_id('second').click()
            time.sleep(2)
            browser.find_element_by_xpath("//select[@id='second']/option[@value=0]").click()
            time.sleep(2)
            Select(browser.find_element_by_id("templates")).select_by_visible_text(protocol_document)
            time.sleep(2)
            browser.find_element_by_xpath("//a[@id='new']").click()
            browser.find_element_by_id('second').click()
            time.sleep(2)
            browser.find_element_by_xpath("//select[@id='second']/option[@value=1]").click()
            browser.find_element_by_id("new").click()
            time.sleep(2)
            Select(browser.find_element_by_id("operation")).select_by_visible_text(control_document)
            time.sleep(2)
            Select(browser.find_element_by_id("contract")).select_by_visible_text(contract_document)
            time.sleep(2)
            browser.find_element_by_id('create-program').click()
            time.sleep(2)
            browser.find_element_by_id('back-go').click()
            time.sleep(2)
            browser.refresh()
            #检验是否新建成功
            try:
               conn = mysql.connector.connect(host=HOST,user=USER,passwd=PASSWORD,db=DATABASE,port=PORT)
               # 创建游标
               cur = conn.cursor()
               # institution_process_model_pkey查询
               sql='select solution_name from t_solution where solution_name="'+ solution_name +'"'
               cur.execute(sql)
               # 获取查询结果
               result_set = cur.fetchall()
               if result_set:
                  for row in result_set:
                      solution_name1 = row[0]
               else:
                  print "No date"
               # 关闭游标和连接
               cur.close()
               conn.close()
            except mysql.connector.Error, e:
                print e.message
            time.sleep(4)
            path="//tr/td[text()='"+ solution_name+"']"
            print path
            if self.browser.find_element_by_xpath(path).is_displayed():
                self.assertTrue(True,"方案新建成功")
            else:
                self.assertFalse(True,"方案新建失败")
        except Exception, e:
            fp = StringIO.StringIO()  # 创建内存文件对象
            traceback.print_exc(file=fp)
            message = fp.getvalue()
            index_file = findStr.findStr(message, "File", 2)
            index_Exception = message.find("message")
            print_message = message[0:index_file] + message[index_Exception:]
            time.sleep(1)
            browser.get_screenshot_as_file(shot_path + browser.title + ".png")
            self.assertTrue(False, print_message)
    def test_Enable_program(self):
        (u"启用方案")
        browser = self.browser
        browser.implicitly_wait(10)
        try:
            time.sleep(3)
            path1="//tr/td[text()='"+solution_name+"']/following::td[5]/a[3]"
            time.sleep(2)
            browser.find_element_by_xpath(path1).click() #点击启用
            time.sleep(3)
            browser.find_element_by_id('start').click() # 确认启用
            time.sleep(1)
            #检验是否启用成功
            status=browser.find_element_by_xpath("//tr/td[text()='"+solution_name+"']/following::td[4]").text
            print  status
            if status==u'已启用':
                self.assertTrue(True,'方案启用成功')
            else:
                self.assertFalse(True,'方案启用失败')
            time.sleep(3)
            #查看详情页面
            browser.find_element_by_xpath("//tr/td[text()='"+solution_name+"']/following::td[5]/a[1]").click()#点击查看
            time.sleep(3)
            name1=browser.find_element_by_xpath('''.//*[@id='program-form']/div/div[2]/div[3]/div[2]''').text
            if name1==solution_name:
                self.assertTrue(True,'新建方案后详情页面显示正常')
            else:
                self.assertFalse(True,'新建方案后详情页面显示异常')
        except Exception, e:
            fp = StringIO.StringIO()  # 创建内存文件对象
            traceback.print_exc(file=fp)
            message = fp.getvalue()
            index_file = findStr.findStr(message, "File", 2)
            index_Exception = message.find("message")
            print_message = message[0:index_file] + message[index_Exception:]
            time.sleep(1)
            browser.get_screenshot_as_file(shot_path + browser.title + ".png")
            self.assertTrue(False, print_message)
    # @classmethod
    # def tearDownClass(cls):
    #     cls.browser.close()
    #     cls.browser.quit()
