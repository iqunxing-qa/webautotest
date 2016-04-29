__author__ = 'MP-4DRH8'
#coding=utf-8
from selenium import webdriver
import time
import unittest
import ConfigParser
import os
import mysql.connector
from classmethod import login
cf = ConfigParser.ConfigParser()
cf.read(r"D:\Workspace\Pythonscripts\environment\env.conf")
host=cf.get('service','host')
method=cf.get('dir','method')
data=cf.get('dir','data')
#读取数据库文件
USER=cf.get('dcf_contract','user')
HOST=cf.get('dcf_contract','host')
PASSWORD=cf.get('dcf_contract','password')
PORT=cf.get('dcf_contract','port')
DATABASE=cf.get('dcf_contract','database')
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

    def test_1(self):
        (u"新建机构工作方式")
        browser = self.browser
        agencyName='test9'
        login.operate_login(self,'operation_login.csv') #登陆
        time.sleep(2)
        browser.find_element_by_link_text(u"产品配置").click()
        time.sleep(2)
        browser.find_element_by_link_text(u"机构工作方式").click()
        time.sleep(2)
        browser.find_element_by_id('new-orgMethod').click()
        browser.find_element_by_id('agencyName').send_keys(agencyName)#融资申请模板
        we=browser.find_element_by_xpath("//div[@id='fileuploadApplyAreaDiv']/div/div/span")
        browser.execute_script("arguments[0].scrollIntoView()",we)
        time.sleep(2)
        we.click()
        time.sleep(2)
        os.system("D:\workspace\Pythonscripts\classmethod\upload_excel.exe")
        browser.find_element_by_id('fileuploadModelName').send_keys('protocol_document')#协议文档
        time.sleep(2)
        browser.find_element_by_xpath("//input[@id='fileuploadModel']/ancestor::div[1]/span").click()
        time.sleep(2)
        os.system("D:\workspace\Pythonscripts\classmethod\upload_word.exe")
        time.sleep(2)
        browser.find_element_by_xpath("//span[text()='+创建模板']").click()
        browser.find_element_by_id('fileuploadControlName').send_keys('control_document')#操作说明文档
        time.sleep(2)
        browser.find_element_by_xpath("//input[@id='fileuploadControl']/ancestor::div[1]/span").click()
        os.system("D:\workspace\Pythonscripts\classmethod\upload_PDF.exe")
        time.sleep(2)
        browser.find_element_by_xpath("//span[text()='+创建操作说明']").click()
        browser.find_element_by_id('fileuploadContractName').send_keys('contract_document')#微合同文档
        time.sleep(2)
        browser.find_element_by_xpath("//input[@id='fileuploadContract']/ancestor::div[1]/span").click()
        os.system("D:\workspace\Pythonscripts\classmethod\upload_TXT.exe")
        time.sleep(2)
        browser.find_element_by_xpath("//span[text()='+创建微合同']").click()
        browser.find_element_by_id('creatAgency').click()
        time.sleep(2)
        browser.find_element_by_link_text(u'返回列表').click()
        time.sleep(2)
        browser.find_element_by_id('search-button').click()
        time.sleep(2)
        #检验是否新建成功
        try:
           # 数据库连接
           conn = mysql.connector.connect(host=HOST,user=USER,passwd=PASSWORD,db=DATABASE,port=PORT)
           # 创建游标
           cur = conn.cursor()
           # institution_process_model_pkey查询
           sql='select institution_process_model_pkey from t_institution_process_model where institution_process_model_name="' + agencyName + '"'
           cur.execute(sql)
           # 获取查询结果
           result_set = cur.fetchall()
           if result_set:
              for row in result_set:
                 institution_id = row[0]
                 print institution_id
           else:
              print "No date"
           # 关闭游标和连接
           cur.close()
           conn.close()
        except mysql.connector.Error, e:
            print e.message
        institution_id = str(institution_id)
        path="//tr[@id=" + institution_id + "]/td[text()="+agencyName+"]"
        if browser.find_element_by_xpath(path).is_displayed():
            print 'New Success ！'
        else:
            print 'New Fail ！'

    def test_2(self):
        (u"启用机构工作方式")
        browser = self.browser
        agencyName='test9'
        #login.operate_login(self,'operation_login.csv') #登陆
        browser.find_element_by_link_text(u"产品配置").click()
        time.sleep(2)
        browser.find_element_by_link_text(u"机构工作方式").click()
        time.sleep(2)
        path="//tr/td[text()="+agencyName+"]/following::td[2]/a[3]"
        browser.find_element_by_xpath(path).click() #点击启用
        time.sleep(3)
        browser.find_element_by_id('modalBtn').click() # 确认启用
        time.sleep(1)
        #检验是否启用成功
        if browser.find_element_by_xpath("").is_displayed():
            print 'Start Success！'
        else:
            print 'Start Fail ！'