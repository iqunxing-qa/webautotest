#coding=utf-8
import csv
import ConfigParser
from selenium import webdriver
import  time
import win32com.client
cf = ConfigParser.ConfigParser()
cf.read(r"D:\Workspace\Pythonscripts\environment\env.conf")
host=cf.get('service','host')
method=cf.get('dir','method')
data=cf.get('dir','data')
def corp_login(self, file):
    browser = self.browser
    paths = ('' + data + file)
    xlxApp = win32com.client.Dispatch('Excel.Application')  # 打开EXCEL
    xlxBook = xlxApp.Workbooks.Open(paths)
    xlSht = xlxBook.Worksheets('Sheet1')
    corname = xlSht.Cells(2, 1).Value
    username = xlSht.Cells(2, 3).Value
    password = xlSht.Cells(2, 4).Value
    browser.get('http://' + host + '.dcfservice.com/login.jsp')
    browser.find_element_by_id('corp_name').clear()
    time.sleep(1)
    browser.find_element_by_id('corp_name').send_keys(corname)
    browser.find_element_by_id('j_user_name').clear()
    time.sleep(1)
    browser.find_element_by_id('j_user_name').send_keys(username)
    browser.find_element_by_id('j_password').clear()
    time.sleep(1)
    browser.find_element_by_id('j_password').send_keys(password)
    time.sleep(1)
    browser.find_element_by_id("reg-btn").click()
    xlxBook.Close(SaveChanges=1)
    del xlxApp

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
         time.sleep(1)
         browser.find_element_by_id('j_user_name').send_keys(username)
         time.sleep(2)
         browser.find_element_by_id('j_password').clear()
         time.sleep(1)
         browser.find_element_by_id('j_password').send_keys(password)
         time.sleep(2)
         browser.find_element_by_id('reg-btn').click()

