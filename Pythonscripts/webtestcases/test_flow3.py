#coding=utf-8
from selenium import webdriver
import time
import csv
import unittest
import ConfigParser
import mysql.connector
from selenium.common.exceptions import NoSuchElementException
import  StringIO
import traceback
from classmethod import findStr
from classmethod import login
import random
import win32com.client
import sys
reload(sys)
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
print shot_path
#读取 产品名 ,模式
xlxApp = win32com.client.Dispatch('Excel.Application')  # 打开EXCEL
xlxBook=xlxApp.Workbooks.Open(r'D:\\workspace\\Pythonscripts\\testdatas\\product_configuration.xlsx')
xlSht = xlxBook.Worksheets('sheet1')
name = str(xlSht.Cells(2, 1).Value)
b=str(time.strftime("%m%d%H%M%S", time.localtime()))
product_name=(name+b).decode('utf-8')
product_type= xlSht.Cells(2, 2).Value
product_id=xlSht.Cells(2,3).Value
xlxBook.Close(SaveChanges=1)
del xlxApp
class Core_Enterprise(unittest.TestCase):
    (u"核心模块")
    @classmethod
    def setUpClass(cls):
        cls.browser = webdriver.Firefox()
        cls.browser.maximize_window()
    def test1_Create_product(self):
        (u"新建产品")
        browser=self.browser
        browser.implicitly_wait(10)
        try:
            login.operate_login(self,'operation_login.csv') #登陆
            time.sleep(2)
            #新建产品
            self.browser.find_element_by_link_text(u"产品配置").click()
            time.sleep(4)
            #product_name=line.decode('utf-8')
            self.browser.find_element_by_id('new-product').click()
            time.sleep(2)
            self.browser.find_element_by_id('product-name').clear()
            self.browser.find_element_by_id('product-name').send_keys(product_name)
            time.sleep(3)
            if product_type=='N+1':
                self.browser.find_element_by_xpath("//span[text()='账单']").click()#贸易结算方式
            else:
                self.browser.find_element_by_xpath("//span[text()='逐笔']").click()#贸易结算方式
            time.sleep(10)
            self.browser.find_element_by_id('button-next-1').click()
            self.browser.find_element_by_id('loanPrincipalCredit1').click()
            self.browser.find_element_by_id('button-next-2').click()
            if product_type=='N+1':
                self. browser.find_element_by_xpath("//input[@name='lendingTarget']/ancestor::label[1]/span[text()='卖家']").click()#放款对象
            else:
                self.browser.find_element_by_xpath("//input[@name='lendingTarget']/ancestor::label[1]/span[text()='卖家']").click()
            self.browser.find_element_by_name('loanApplicant').click()
            if product_type=='N+1':
               self.browser.find_element_by_xpath("//input[@name='dataProvider']/ancestor::label[1]/span[text()='买方']").click()#数据提交方式
            else:
              self.browser.find_element_by_xpath("//input[@name='dataProvider']/ancestor::label[1]/span[text()='卖方']").click()
            if product_type=='B2G':
               self.browser.find_element_by_xpath("//input[@name='dataConfirmMethod']/ancestor::label[1]/span[text()='一次性电子确认']").click()#数据确认方式
            else:
                self.browser.find_element_by_xpath("//input[@name='dataConfirmMethod']/ancestor::label[1]/span[text()='逐笔电子确认']").click()
            self.browser.find_element_by_id('button-create').click()
            time.sleep(2)
            browser.find_element_by_id('back-go').click()
            time.sleep(2)
            # #检验是否新建成功
            try:
            # 数据库连接
                conn = mysql.connector.connect(host=HOST,user=USER,passwd=PASSWORD,db=DATABASE,port=PORT)
            # 创建游标
                cur = conn.cursor()
            # product_name查询
                sql='select product_pkey from t_product where product_name="'+ product_name + '"'
                cur.execute(sql)
            # 获取查询结果
                result_set = cur.fetchall()
                if result_set:
                    for row in result_set:
                     product_id = row[0]
                else:
                 print "No date"
            # 关闭游标和连接
                cur.close()
                conn.close()
            except mysql.connector.Error, e:
                print e.message
            product_id = str(product_id)
            #将product_id 写入product_configuration.xlsx
            xlxApp = win32com.client.Dispatch('Excel.Application')  # 打开EXCEL
            xlxBook=xlxApp.Workbooks.Open(r'D:\\workspace\\Pythonscripts\\testdatas\\product_configuration.xlsx')
            xlSht = xlxBook.Worksheets('sheet1')
            xlSht.Cells(2,3).Value=product_id
            xlxBook.Close(SaveChanges=1)
            del xlxApp
            path="//tr/td[text()="+ product_id +"]"
            if self.browser.find_element_by_xpath(path).is_displayed():
                self.assertTrue(True,"产品新建成功")
            else:
                self.assertFalse(True,"产品新建失败")
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
    def test2_Enable_product(self):
        (u"启用产品")
        browser=self.browser
        browser.implicitly_wait(10)
        try:
            time.sleep(2)
            path="//tr/td[text()='"+product_name+"']/following::td[4]/a[3]"
            # 点击启用
            browser.find_element_by_xpath(path).click()
            time.sleep(2)
            browser.find_element_by_id('start').click()
            time.sleep(2)
            #检验是否启用成功
            status=browser.find_element_by_xpath("//tr/td[text()='"+ product_name +"']/following::td[3]").text
            time.sleep(2)
            if status==u'已启用':
                self.assertTrue(True,'产品启用成功')
            else:
                self.assertFalse(True,'产品启用失败')
            time.sleep(2)
            #查看详情页面
            browser.find_element_by_xpath("//tr/td[text()='"+ product_name +"']/following::td[4]/a[1]").click()#点击查看
            time.sleep(3)
            name1=browser.find_element_by_xpath('''.//label[@class="col-sm-2 control-label"][contains(@for,'product-name')]''').text
            if name1==product_name:
                self.assertTrue(True,'新建产品后详情页面显示正常')
            else:
                self.assertFalse(True,'新建产品后详情页面显示异常')
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
    @classmethod
    def tearDownClass(cls):
        cls.browser.close()
        cls.browser.quit()