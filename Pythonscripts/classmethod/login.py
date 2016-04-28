#coding=utf-8
import csv
import ConfigParser
from selenium import webdriver
cf = ConfigParser.ConfigParser()
cf.read(r"D:\Workspace\Pythonscripts\environment\env.conf")
host=cf.get('service','host')
method=cf.get('dir','method')
data=cf.get('dir','data')

def corp_login(self,csvfile):
     browser=self.browser
     csvpaths=file(''+data+csvfile, 'rb')
     f = csv.reader(csvpaths)
     browser.get('http://'+host+'.dcfservice.com/login.jsp')
     for line in f:
         #list=line.replace("\n","").split(",")
         #print list
         corname=line[0].decode('utf-8')
         username=line[1].decode('utf-8')
         password=line[2].decode('utf-8')

         browser.find_element_by_id('corp_name').clear()
         browser.find_element_by_id('corp_name').send_keys(corname)
         browser.find_element_by_id('j_user_name').clear()
         browser.find_element_by_id('j_user_name').send_keys(username)
         browser.find_element_by_id('j_password').clear()
         browser.find_element_by_id('j_password').send_keys(password)
         browser.find_element_by_id("reg-btn").click()


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

