#coding=utf-8
from selenium import webdriver
import time
import unittest
import ConfigParser
import  StringIO
import traceback
from selenium.webdriver.support.ui import Select
from classmethod import findStr
import csv
from selenium.common.exceptions import NoSuchElementException
from classmethod import login
import mysql.connector
import random
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
#读取product_id,product_type
csvpaths=file(''+data+'middle_product.csv', 'r') #读取 产品id以及模式
f=csv.reader(csvpaths)
for line in f:
    product_type=line[0].decode('utf-8')
    product_id=line[1]
    print product_type,product_id
#定义方案名称
a=str(random.randint(100, 1000))
solution_name=((u'群星测试')+a)
#读取模板文件名称
csvfile =file(''+data +'middle_agency.csv','rb')
f=csv.reader(csvfile)
for line in f:
    protocol_document=line[0].decode('utf-8')
    control_document=line[1].decode('utf-8')
    contract_document=line[2].decode('utf-8')
    agency_id=line[3].decode('utf-8')
    print protocol_document,control_document,contract_document,agency_id
csvfile.close()
class Core_Enterprise(unittest.TestCase):
    (u"核心模块")
    @classmethod
    def setUpClass(cls):
        cls.browser = webdriver.Firefox()
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
            browser.find_element_by_xpath("//select[@id='org']/option[2]").click()#确定机构
            time.sleep(2)
            browser.find_element_by_name('program').send_keys(solution_name)
            #机构
            browser.find_element_by_id('first').click()
            time.sleep(2)
            browser.find_element_by_xpath("//select[@id='first']/option[@value=2]").click()
            browser.find_element_by_id('second').click()
            time.sleep(2)
            #如果product_type 是N+1 则为机构对卖方，否则为机构对买方
            if product_type=='N+1':

                browser.find_element_by_xpath("//select[@id='second']/option[@value=0]").click()
            else:
                browser.find_element_by_xpath("//select[@id='second']/option[@value=0]").click()
            time.sleep(2)
            Select(browser.find_element_by_id("templates")).select_by_visible_text(protocol_document)
            time.sleep(2)
            browser.find_element_by_xpath("//a[@id='new']").click()
            time.sleep(2)
            Select(browser.find_element_by_id("operation")).select_by_visible_text(control_document)
            time.sleep(2)
            Select(browser.find_element_by_id("contract")).select_by_visible_text(contract_document)
            time.sleep(2)
            browser.find_element_by_id('create-program').click()
            time.sleep(2)
            browser.find_element_by_id('back-go').click()
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
                      print solution_name1
               else:
                  print "No date"
               # 关闭游标和连接
               cur.close()
               conn.close()
            except mysql.connector.Error, e:
                print e.message
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
       try:
          time.sleep(3)
          path1="//tr/td[text()='"+solution_name+"']/following::td[5]/a[3]"
          time.sleep(2)
          browser.find_element_by_xpath(path1).click() #点击启用
          time.sleep(3)
          browser.find_element_by_id('start').click() # 确认启用
          time.sleep(1)
          #检验是否启用成功
          path2="//tr/td[text()='"+solution_name+"']/following::td[4]"
          if browser.find_element_by_xpath(path2).is_displayed():
              print 'ok'
              self.assertTrue(True)
          else:
              self.assertFalse(False)
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
