#coding=utf-8
import csv
import ConfigParser
from selenium import webdriver
import  time
cf = ConfigParser.ConfigParser()
cf.read(r"D:\Workspace\Pythonscripts\environment\env.conf")
host=cf.get('service','host')
method=cf.get('dir','method')
data=cf.get('dir','data')
import win32com.client
import sys
import os
reload(sys)
sys.setdefaultencoding('utf8')
def corp_login(self,file):
     browser=self.browser
     paths=(''+data+file)
     xlxApp = win32com.client.Dispatch('Excel.Application')  # 打开EXCEL
     xlxBook=xlxApp.Workbooks.Open(paths)
     xlSht = xlxBook.Worksheets('sheet1')
     corname = xlSht.Cells(2, 1).Value
     username= xlSht.Cells(2, 3).Value
     password=xlSht.Cells(2, 4).Value
     browser.get('http://'+host+'.dcfservice.com/login.jsp')
     browser.find_element_by_id('corp_name').clear()
     time.sleep(1)
     browser.find_element_by_id('corp_name').send_keys(corname)
     browser.find_element_by_id('j_user_name').clear()
     time.sleep(1)
     browser.find_element_by_id('j_user_name').send_keys(username)
     browser.find_element_by_id('j_password').clear()
     time.sleep(1)
     browser.find_element_by_id('j_password').send_keys(password)
     browser.find_element_by_id("reg-btn").click()
     time.sleep(5)
     if file=='core_customer.xlsx':
         core_login_status=browser.find_element_by_xpath('''.//*[@id='logoutDiv']/a''').text
         if core_login_status==username:
             self.assertTrue(True,'登陆成功')
         else:
             self.assertFalse(True,'登录失败')
     elif file=='institution_data.xlsx':
         institution_login_status=browser.find_element_by_xpath(".//*[@id='personTip']/span").text
         if institution_login_status==username:
             self.assertTrue(True,'登陆成功')
         else:
             self.assertFalse(True,'登录失败')
     elif file=='chain_customer.xlsx':
         chain_login_status=browser.find_element_by_xpath('''.//*[@id='logoutDiv']/a''').text
         if chain_login_status==username:
             self.assertTrue(True,'登陆成功')
         else:
             self.assertFalse(True,'登录失败')

def operate_login(self,csvfile):
     browser=self.browser
     csvpaths=file(''+data+csvfile, 'rb')
     f = csv.reader(csvpaths)
     browser.get('http://'+host+'.dcfservice.com/loginop.jsp')
     for line in f:
         #list=line.replace("\n","").split(",")
         #print list
         username=line[0].decode('utf-8')
         password=line[1].decode('utf-8')
         browser.find_element_by_id('j_user_name').clear()
         browser.find_element_by_id('j_user_name').send_keys(username)
         browser.find_element_by_id('j_password').clear()
         browser.find_element_by_id('j_password').send_keys(password)
         browser.find_element_by_id('reg-btn').click()
         time.sleep(3)
         op_login_status=browser.find_element_by_id("profile")
         if op_login_status.is_displayed():
             self.assertTrue(True,'登陆成功')
         else:
             self.assertFalse(True,'登录失败')

