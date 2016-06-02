#coding=utf-8
from selenium import webdriver
import time
import unittest
import csv
import ConfigParser
from selenium.common.exceptions import NoSuchElementException
import  StringIO
import traceback
from classmethod import findStr
from classmethod import login
import win32com.client
import sys
import os
reload(sys)
sys.setdefaultencoding('utf8')
import mysql.connector
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
#receipt_id='e1013'
#读取单据编号receipt_id,单据金额receipt_money
#os.system('D:\\workspace\\Pythonscripts\\testdatas\\transaction_flow.xlsx')
'''xlxApp = win32com.client.Dispatch('Excel.Application')  # 打开EXCEL
xlxBook=xlxApp.Workbooks.Open(r'D:\\workspace\\Pythonscripts\\testdatas\\transaction_flow.xlsx')
xlSht = xlxBook.Worksheets('sheet1')
receipt_id = xlSht.Cells(2, 1).Value
receipt_money= xlSht.Cells(2, 4).Value
print receipt_id,receipt_money
#os.system('taskkill /f /im "EXCEL.exe"')
del xlxApp'''

#读取核心企业名称core_enterprise_name
csvpaths2=file(''+data+'core_enterprise_login.csv', 'rb')
f = csv.reader(csvpaths2)
for line in f:
    core_enterprise_name=line[0].decode('utf-8')
    print core_enterprise_name
csvpaths2.close()
#读取截图存放路径
shot_path=cf.get('shotpath','path')
class Core_Enterprise(unittest.TestCase):
    (u"核心模块")
    @classmethod
    def setUpClass(cls):
        cls.browser = webdriver.Firefox()
        cls.browser.maximize_window()
    '''def test_Apply_repayment(self):
        (u"申请还款")
        browser = self.browser
        browser.implicitly_wait(10)
        try:
            login.corp_login(self,'core_enterprise_login.csv') #核心企业登陆
            time.sleep(2)
            browser.find_element_by_id('today').click() #切换到当天页面
            time.sleep(2)
            browser.find_element_by_xpath("//div[@id='anchorPoint']/following::div/ul/li[6]").click()#点击已融资
            time.sleep(2)
            #根据单据号定位需要还款的融资款
            we=browser.find_element_by_xpath("//div[@id='anchorPoint']/following::div[1]/ul")
            browser.execute_script("arguments[0].scrollIntoView()",we)
            time.sleep(2)
            path="//tr/td[2][text()='"+ receipt_id +"']/following::td[7]/div/button[1]"
            print path
            browser.find_element_by_xpath(path).click()
            time.sleep(2)
            now_handle = browser.current_window_handle#获取当前窗口句柄
            browser.find_element_by_xpath("//a[@id='cancelButton']/following::button[1]").click()#下一步
            time.sleep(2)
            all_handles=browser.window_handles
            for handle in all_handles:
                if handle != now_handle:
                    print"Switched window is %s" % handle  # 输出待选择的窗口句柄
                    browser.switch_to_window(handle)
                    time.sleep(2)
                    browser.find_element_by_xpath("//button[text()='立即支付']").click()   #立即支付
                    time.sleep(2)
                    browser.close()
            #################################################################################################
            #                          以下校验是否还款成功                                                 #
            #################################################################################################
            time.sleep(2)
            browser.switch_to_window(now_handle)  #返回主窗口
            time.sleep(2)
            browser.find_element_by_xpath("//div[@id='paymentInfoApplication']/div/div/div/button[text()='×']").click()
            time.sleep(90)
            browser.find_element_by_id('today').click() #切换到当天页面
            time.sleep(2)
            browser.find_element_by_xpath("//div[@id='anchorPoint']/following::div/ul/li[8]").click()#到已还款页面
            time.sleep(5)
            count=0
        #循环查询5-8min,等待还款状态改变
            while count<50:
                browser.find_element_by_id('today').click() #切换到当天页面
                time.sleep(3)
                browser.find_element_by_xpath("//div[@id='anchorPoint']/following::div/ul/li[8]").click()#到已还款页面
                time.sleep(10)
                count+=1
            path="//tr/td[1][text()='"+ receipt_id +"']"#//tr/td[1][text()=234]
            if browser.find_element_by_xpath(path).is_displayed():
                print 'ok'
                self.assertTrue(True,'Apply repayment success!')
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
        browser.close()'''

    def test_Settlement_documents(self):
        (u"检查凭证")
        browser = self.browser
        browser.implicitly_wait(10)
        try:
            login.operate_login(self,'operation_login.csv') #核心企业登陆
            time.sleep(2)
            browser.find_element_by_id('qunxingPay-nav').click()
            browser.find_element_by_xpath("//div[@id='header']//following::div/ul/li[3]/div").click()
            time.sleep(2)
            browser.find_element_by_xpath("//div[@id='header']//following::div/ul/li[3]/div/following::ul/li[3]/a[text()='结算凭证查询']").click()
            time.sleep(4)
            browser.find_element_by_xpath("//input[@id='createTime']/ancestor::div[1]/span/span/span").click()#选择时间
            time.sleep(2)
            browser.find_element_by_xpath("/html/body/div[2]/div/div/div/div/button[1][text()=' 今 天']").click()#选择今天
            time.sleep(2)
            #path1="//tr/td[4][text()='支付']/following::td[1][text()='"+ core_enterprise_name +"']/following::td[4][text()='"+ receipt_id +"']"
            '''try:
                benjin_no=""
                lixi_no=""
                zhifu_no=""
                conn = mysql.connector.connect(host=HOST,user=USER,passwd=PASSWORD,db=DATABASE,port=PORT)
            # 创建游标
                cur = conn.cursor()
                sql='SELECT d.settle_id,c.memo FROM dcf_loan.t_loan_document a LEFT JOIN dcf_loan.t_repayment_detail b on a.loan_document_id=b.document_id LEFT JOIN dcf_loan.t_repayment_settlement_order c on c.repayment_id=b.repayment_id LEFT JOIN dcf_loan.t_repayment_settlement_notify_log d ON d.so_no=c.so_no where a.loan_document_no="1005"'
                cur.execute(sql)
            # 获取查询结果
                result_set = cur.fetchall()
                if result_set:
                    for row in result_set:
                        print row[1]
                        if u"本金" in str(row[1]):
                            benjin_no=str(row[0])
                        elif u"利息" in str(row[1]):
                            lixi_no=str(row[0])
                        elif u"融资" in str(row[1]):
                            zhifu_no=str(row[0])
                        print  benjin_no ,zhifu_no,lixi_no
                else:
                    print 'fail'
            # 关闭游标和连接
                cur.close()
                conn.close()
            except mysql.connector.Error, e:
                print e.message
            benjin_no=str(benjin_no)
            zhifu_no=str(zhifu_no)
            lixi_no=str(lixi_no)
            time.sleep(2)'''
            #校验还款单据金额和结算状态
            browser.find_element_by_xpath("//div/label[text()='付款方']/following::input[1]").clear()
            browser.find_element_by_xpath("//div/label[text()='结算流水号']/following::input[1]").send_keys('STL201605250004166')#输入本金的结算编号
            time.sleep(2)
            browser.find_element_by_xpath("//button[@id='searchBtn']").click()
            time.sleep(2)
            money=browser.find_element_by_xpath("//tr/td[9]").text
            money=money.replace(',','')
            money=money.replace('.00','')
            money=int(money)
            print money
            if money==200:
                self.assertTrue(True,'还款金额正确')
            else:
                self.assertFalse(True,'还款金额错误')
            count2=1
            status=browser.find_element_by_xpath("//tr/td[12]").text
            while status!="结算成功":
                #status=browser.find_element_by_xpath("//tr/td[12]").text
                print status
                if status=="结算成功":
                    self.assertTrue(True,"结算成功")
                break

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







