# coding:utf-8
import random
from classmethod import *
from selenium import webdriver
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
# 引入ActionChains鼠标操作类
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
# 引入keys类操作
import time
import unittest
import csv
import ConfigParser
import StringIO
import traceback
import mysql.connector
import win32com.client
import sys
import os
import re
reload(sys)
sys.setdefaultencoding('utf8')
#获取主要配置
cf = ConfigParser.ConfigParser()
cf.read(r"D:\Workspace\Pythonscripts\environment\env.conf")
host=cf.get('service','host')
method=cf.get('dir','method')
data=cf.get('dir','data')
#获取Firefox的profile
propath=getprofile.get_profile()
profile=webdriver.FirefoxProfile(propath)
print propath
#读取数据库文件
USER=cf.get('database','user')
HOST=cf.get('database','host')
PASSWORD=cf.get('database','password')
PORT=cf.get('database','port')
DATABASE=cf.get('database','dcf_user')
#读取截图存放路径
shot_path=cf.get('shotpath','path')
#读取机构注册信息
xlxApp = win32com.client.Dispatch('Excel.Application')  # 打开EXCEL
xlxBook=xlxApp.Workbooks.Open(r'D:\\workspace\\Pythonscripts\\testdatas\\institution_data.xlsx')
xlSht = xlxBook.Worksheets('sheet1')
depart_name = xlSht.Cells(2, 3).Value
depart_mail= xlSht.Cells(2, 5).Value
depart_mobile= xlSht.Cells(2,6).Value
pattern = re.compile(r'\d*')
depart_mobile= re.search(pattern, str(xlSht.Cells(2, 6).Value)).group()
xlxBook.Close(SaveChanges=1)
del xlxApp
class department_register(unittest.TestCase):
    u"机构注册验证"
    jigou_name=""
    @classmethod
    def setUpClass(cls):
        cls.browser=webdriver.Firefox(profile)
        cls.browser.maximize_window()

    def test_1_invite(self):
        u"平台邀请注册"
        browser=self.browser
        try:
            #admin账户登录
            login.operate_login(self,'operation_login.csv')
            time.sleep(3)
            #客户邀请
            browser.find_element_by_link_text(u'客户邀请').click()
            time.sleep(3)
            browser.find_element_by_id('inviteCustomer').click()
            time.sleep(6)
            a=str(time.strftime("%m%d%H%M%S", time.localtime()))
            department_register.jigou_name=((u'机构测试')+a)#随机生成机构名
            #将机构名写入institution_data.xlsx
            xlxApp = win32com.client.Dispatch('Excel.Application')  # 打开EXCEL
            xlxBook=xlxApp.Workbooks.Open(r'D:\\workspace\\Pythonscripts\\testdatas\\institution_data.xlsx')
            xlSht = xlxBook.Worksheets('sheet1')
            xlSht.Cells(2,1).Value=department_register.jigou_name
            xlxBook.Close(SaveChanges=1)
            del xlxApp
            time.sleep(3)
            browser.find_element_by_id('customerFullName').send_keys(department_register.jigou_name)
            browser.find_element_by_xpath(".//*[@id='inviteForm']/div[2]/div/div/div[1]/button[2]").click()
            time.sleep(3)
            browser.find_element_by_link_text(u'农、林、牧、渔业').click()
            browser.find_element_by_xpath(".//*[@id='inviteForm']/div[3]/div/div/div[1]/button[2]").click()
            time.sleep(3)
            sh=browser.find_element_by_css_selector("#province>li>a[value='310000']")
            browser.execute_script("arguments[0].scrollIntoView()",sh)
            sh.click()
            browser.find_element_by_id('optionsRadios2').click()#选择机构
            time.sleep(2)
            invitedUser=browser.find_element_by_id("invitedUser")
            invitedUser.clear()
            invitedUser.send_keys(depart_name)
            time.sleep(2)
            invitedEmail=browser.find_element_by_id("invitedEmail")
            invitedEmail.clear()
            invitedEmail.send_keys(depart_mail)
            time.sleep(2)
            invitedMobile=browser.find_element_by_id("invitedMobile")
            invitedMobile.clear()
            invitedMobile.send_keys(depart_mobile)
            time.sleep(3)
            browser.find_element_by_css_selector(".btn.btn-danger.createInviteBtn").click()
            time.sleep(10)
        except NoSuchElementException,e:
            fp = StringIO.StringIO()  # 创建内存文件对象
            traceback.print_exc(file=fp)
            message = fp.getvalue()
            index = findStr.findStr(message, "File", 2)
            message = message[0:index]
            message = message + e.msg
            browser.get_screenshot_as_file(shot_path + browser.title + ".png")
            self.assertTrue(False, message)

    def test_2_department_register(self):
        u"机构客户注册"
        browser=self.browser
        try:
            department_register.url=browser.execute_script("return document.getElementById('inviteUrl-core').value")
            print department_register.url
            time.sleep(3)
            browser.get(department_register.url)
            browser.implicitly_wait(3)
            time.sleep(2)
            browser.find_element_by_id('inputPassword').send_keys('iqunxing1234')
            browser.find_element_by_id('inputRePassword').send_keys('iqunxing1234')
            browser.find_element_by_id('getDynamic').click()
            time.sleep(3)
            now_handle = browser.current_window_handle
            #获取验证码
            vcode_url="http://"+host+'.dcfservice.com/v1/public/sms/get?cellphone='+depart_mobile
            js_script='window.open("'+vcode_url+'")'
            browser.execute_script(js_script)
            time.sleep(2)
            all_handles=browser.window_handles
            for handle in all_handles:
                if handle != now_handle:
                    print"Switched window is %s" % handle  # 输出待选择的窗口句柄
                    browser.switch_to_window(handle)
                    time.sleep(3)
            Dynamic_code=browser.find_element_by_css_selector("html>body>pre").text
            Dynamic_code=Dynamic_code[1:7]
            print  Dynamic_code #取得验证码
            #填写验证码
            browser.switch_to_window(now_handle)
            browser.find_element_by_id('validateCode').send_keys(Dynamic_code)
            browser.find_element_by_id('registerbtn').click()
            time.sleep(10)
            # if browser.find_element_by_id('login').is_displayed():
            #     self.assertTrue(True,"返回首页成功")
            # else:
            #     self.assertFalse(True,"返回首页失败")
            browser.find_element_by_xpath('html/body/div[1]/div[1]/div[2]/button').click()#提交资料
            time.sleep(3)
    ###############################################################################################################
    #                                       以下区域为上传照片                                                    #
    ###############################################################################################################
            # 上传营业执照
            browser.find_element_by_xpath(".//*[@id='index-area']/div/div[1]/div[3]/div[2]/div[2]/div[1]/div[1]/div[2]/span").click()
            upload_file = method + "\\upload.exe "+data+"test_picture.jpg"
            os.system(upload_file)
            time.sleep(2)
            # 上传组织机构代码
            organizationNoFileupload = browser.find_element_by_xpath(".//*[@id='index-area']/div/div[1]/div[4]/div[2]/div[2]/div[1]/div[1]/div[2]/span")
            browser.execute_script("arguments[0].scrollIntoView()", organizationNoFileupload)
            time.sleep(2)
            browser.find_element_by_xpath(".//*[@id='index-area']/div/div[1]/div[4]/div[2]/div[2]/div[1]/div[1]/div[2]/span").click()
            time.sleep(2)
            upload_file = method + "\\upload.exe " + data + "test_picture.jpg"
            os.system(upload_file)
            time.sleep(2)
            #上传身份证正面
            operatorIdFrontFileupload = browser.find_element_by_xpath(".//*[@id='index-area']/div/div[1]/div[5]/div[2]/div[2]/div[1]/div[1]/div[1]/div[2]/span")
            browser.execute_script("arguments[0].scrollIntoView()", operatorIdFrontFileupload)
            time.sleep(2)
            browser.find_element_by_xpath(".//*[@id='index-area']/div/div[1]/div[5]/div[2]/div[2]/div[1]/div[1]/div[1]/div[2]/span").click()
            time.sleep(2)
            upload_file = method + "\\upload.exe " + data + "test_picture.jpg"
            os.system(upload_file)
            time.sleep(2)
            #上传身份证反面
            operatorIdFrontFileupload = browser.find_element_by_xpath(".//*[@id='index-area']/div/div[1]/div[5]/div[2]/div[2]/div[1]/div[2]/div[1]/div[2]/span")
            browser.execute_script("arguments[0].scrollIntoView()", operatorIdFrontFileupload)
            time.sleep(2)
            browser.find_element_by_xpath(".//*[@id='index-area']/div/div[1]/div[5]/div[2]/div[2]/div[1]/div[2]/div[1]/div[2]/span").click()
            time.sleep(2)
            upload_file = method + "\\upload.exe " + data + "test_picture.jpg"
            os.system(upload_file)
            time.sleep(2)
            #上传手持身份证照片
            operatorIdFrontFileupload = browser.find_element_by_xpath(".//*[@id='index-area']/div/div[1]/div[6]/div[2]/div[2]/div[1]/div[1]/div[2]/span")
            browser.execute_script("arguments[0].scrollIntoView()", operatorIdFrontFileupload)
            time.sleep(2)
            browser.find_element_by_xpath(".//*[@id='index-area']/div/div[1]/div[6]/div[2]/div[2]/div[1]/div[1]/div[2]/span").click()
            time.sleep(2)
            upload_file = method + "\\upload.exe " + data + "test_picture.jpg"
            os.system(upload_file)
            time.sleep(2)
            #上传完成点击提交资料
            browser.execute_script("arguments[0].click()", browser.find_element_by_xpath(".//*[@id='index-area']/div/div[1]/div[7]/button"))
            time.sleep(5)
        except NoSuchElementException,e:
            fp = StringIO.StringIO()  # 创建内存文件对象
            traceback.print_exc(file=fp)
            message = fp.getvalue()
            index = findStr.findStr(message, "File", 2)
            message = message[0:index]
            message = message + e.msg
            browser.get_screenshot_as_file(shot_path + browser.title + ".png")
            self.assertTrue(False, message)
    def test_3_department_authentication(self):
        (u"核心企业认证")
        browser = self.browser
        browser.implicitly_wait(10)
        try:
            login.operate_login(self,'operation_login.csv')
            time.sleep(2)
            conn = mysql.connector.connect(host=HOST, user=USER, passwd=PASSWORD, db=DATABASE, port=PORT)
            # 创建游标
            cur = conn.cursor()
            # customername_id查询
            sql = 'select customer_id from user where customer_name="'+department_register.jigou_name+'"'
            cur.execute(sql)
            # 获取查询结果
            result_set = cur.fetchall()
            if result_set:
                for row in result_set:
                    customername_id = row[0]
                    print customername_id
            else:
                self.assertTrue(False,"the customer_id do not exsit in database!")
            # 关闭游标和连接
            cur.close()
            conn.close()
        except mysql.connector.Error, e:
            print e.message
        customername_id = str(customername_id)
        #将customername_id写入institution_data
        xlxApp = win32com.client.Dispatch('Excel.Application')  # 打开EXCEL
        xlxBook=xlxApp.Workbooks.Open(r'D:\\workspace\\Pythonscripts\\testdatas\\institution_data.xlsx')
        xlSht = xlxBook.Worksheets('sheet1')
        xlSht.Cells(2,7).Value=customername_id
        xlxBook.Close(SaveChanges=1)
        del xlxApp
        time.sleep(4)
        browser.find_element_by_xpath("//span[@id='searchFilter']/input").send_keys(department_register.jigou_name)
        browser.find_element_by_xpath("//button[@id='btn-search']").click()
        time.sleep(2)
        path = "//*[@id='table-account']/tbody/tr[@data-customerid='"+customername_id+"']/td[6]/div/a[2]"
        browser.find_element_by_xpath(path).click()
        time.sleep(2)
        ####################################################################################################
        #                               以下区域填写营业执照相关信息                                       #
        ####################################################################################################
        # 点击业执照区域的通过按钮，弹出编辑区域
        browser.find_element_by_xpath('''.//input[@name="businessLicense_Pass"][@value='2']''').click()
        time.sleep(4)
        # 编辑营业执照
        browser.find_element_by_xpath('''.//input[@name="enterprise_no"]''').send_keys("1111112")#填写营业执照注册号
        time.sleep(1)
        #browser.find_element_by_css_selector('''.radio.col-xs-6>label>input[value="0"][2]''').click()
        time.sleep(1)
        start_date=browser.find_element_by_xpath('''.//input[@name="business_license_start_date"]''')
        browser.execute_script('''arguments[0].value="2012-2-21"''',start_date)
        time.sleep(3)
        end_date=browser.find_element_by_xpath('''.//input[@name="business_license_end_date"]''')
        browser.execute_script('''arguments[0].value="2020-2-21"''',end_date)
        time.sleep(3)
        browser.find_element_by_xpath('''.//input[@name="enterprise_money"]''').send_keys("10000")
        time.sleep(4)

        # 上移营业执照区域使组织机构代码通过按钮可见
        businessLicense_move = browser.find_element_by_xpath('''.//div[@for="businessLicense_Pass"][contains(@style,'block')]''')
        browser.execute_script("arguments[0].scrollIntoView()", businessLicense_move)
        time.sleep(2)
        # 点击组织机构区域的通过按钮，弹出编辑区域
        browser.find_element_by_xpath('''.//input[@name="organizationNo_Pass"][@value="2"]''').click()
        time.sleep(4)
        # 编辑组织机构
        browser.find_element_by_xpath('''.//input[@name="organization_no"]''').send_keys("111111")
        time.sleep(1)
        browser.find_element_by_xpath('''.//input[@name="organization_no_regist"]''').send_keys("121212")
        time.sleep(1)
        Select(browser.find_element_by_id("province")).select_by_value("310000")#选择上海市
        time.sleep(1)
        Select(browser.find_element_by_id("city")).select_by_value("310100")
        time.sleep(2)
        validity_date=browser.find_element_by_xpath('''.//input[@name="organization_deadline"]''')
        browser.execute_script('''arguments[0].value="2020-2-21"''',validity_date)
        time.sleep(2)
        # 点击操作身份证区域的通过按钮，弹出编辑区域
        browser.find_element_by_xpath('''.//input[@name="operatorId_Pass"][@value="2"]''').click()
        time.sleep(1)
        # 编辑操作身份证区域
        browser.find_element_by_xpath('''.//input[@name="operator_user_name"]''').send_keys(u"周大强")
        time.sleep(2)
        browser.find_element_by_xpath('''.//input[@name="operator_ID"]''').send_keys("340823198876611221")
        time.sleep(2)
        idcard_startdate=browser.find_element_by_xpath('''.//input[@name="operator_ID_start_date"]''')
        browser.execute_script('''arguments[0].value="2016-4-1"''', idcard_startdate)
        time.sleep(2)
        idcard_enddate=browser.find_element_by_xpath('''.//input[@name="operator_ID_end_date"]''')
        browser.execute_script('''arguments[0].value="2020-4-1"''', idcard_enddate)
        time.sleep(2)
        # 上移身份证操作区域使操作者手持身份证区域的通过按钮可视
        operator_move = browser.find_element_by_xpath('''.//div[@for="operatorId_Pass"][contains(@style,'block')]''')
        browser.execute_script("arguments[0].scrollIntoView()", operator_move)
        time.sleep(2)
        # 点击手持身份证区域的通过按钮
        browser.find_element_by_xpath('''.//input[@name="picHandling_Pass"][@value="2"]''').click()
        time.sleep(2)
         #填写完所有资料保存
        browser.find_element_by_id("btnSubmit").click()
        time.sleep(3)
        # browser.find_element_by_xpath("//button[@id='choseadmin-confrim']").click()
        # time.sleep(2)
        browser.find_element_by_link_text(u"返回列表").click()
        time.sleep(2)
        #查看是否被认证
        ver_xpath='''.//*[@id='table-account']/tbody/tr[@data-customerid="'''+customername_id+'''"]/td[4]/span'''
        ver_text=browser.find_element_by_xpath(ver_xpath).text
        if ver_text==u"已认证":
            self.assertTrue(True,"客户认证成功")
        else:
            self.assertTrue(False,"客户资料已填写，但是提交后客户认证失败")
        time.sleep(5)
        #检验查看详情
        browser.find_element_by_xpath('''.//*[@id='table-account']/tbody/tr[@data-customerid="'''+customername_id+'''"]/td[6]/div/a[1]''').click()#点击查看详情
        time.sleep(3)
        service_status=browser.find_element_by_xpath('''.//*[@id='tbl_services']/tbody/tr/td[2]/span''').text
        if service_status==u'已认证':
            self.assertTrue(True,"详情页面显示客户认证成功")
        else:
            self.assertTrue(False,"详情页显示异常")
    def test_4_department_bank_no(self):
        (u'获取银行信息')
        browser = self.browser
        browser.implicitly_wait(10)
        try:
            login.corp_login(self,'institution_data.xlsx')
            browser.find_element_by_link_text(u"账户管理").click()
            time.sleep(5)
            browser.find_element_by_xpath("/html/body/div/div[2]/div[2]/div/ul/li/div/div/span[text()='尾号：']/following::span[1]").click()
            time.sleep(2)
            bank_no=browser.find_element_by_xpath('''.//span[@class="account-info-num"]''').text
            #将银行账号写入institution_data.xlsx
            xlxApp = win32com.client.Dispatch('Excel.Application')  # 打开EXCEL
            xlxBook=xlxApp.Workbooks.Open(r'D:\\workspace\\Pythonscripts\\testdatas\\institution_data.xlsx')
            xlSht = xlxBook.Worksheets('sheet1')
            xlSht.Cells(2,2).Value=bank_no
            xlxBook.Close(SaveChanges=1)
            del xlxApp
        except NoSuchElementException,e:
            fp = StringIO.StringIO()  # 创建内存文件对象
            traceback.print_exc(file=fp)
            message = fp.getvalue()
            index = findStr.findStr(message, "File", 2)
            message = message[0:index]
            message = message + e.msg
            browser.get_screenshot_as_file(shot_path + browser.title + ".png")
            self.assertTrue(False, message)

    def test_5_bank_recharge(self):
        (u'账户记账充值')
        browser = self.browser
        browser.implicitly_wait(10)
        try:
            login.operate_login(self,'operation_login.csv')
            #获取银行账号id
            browser.find_element_by_id('qunxingPay-nav').click()
            time.sleep(5)
            browser.find_element_by_xpath("//input[@id='search-text']").send_keys(department_register.jigou_name)#输入机构名称
            time.sleep(2)
            browser.find_element_by_id("btn-search").click()
            time.sleep(2)
            bank_id=browser.find_element_by_xpath(".//*[@class='odd grouped']/td[text()='通用结算户']/following::td[1]").text
            wait_time=0
            while True:
                time.sleep(5)
                browser.find_element_by_id("search-text").clear()
                browser.find_element_by_id("search-text").send_keys(department_register.jigou_name)
                browser.find_element_by_id("btn-search").click()  # 在群星支付界面搜索核心企业账户
                time.sleep(10)
                try:
                    if browser.find_element_by_xpath("//div[@class='loading']").is_displayed():
                        browser.refresh()
                        time.sleep(5)
                        ###########################s
                        # 查找通用结算户
                        ###########################
                    try:
                        elements = browser.find_elements_by_xpath(".//*[@id='table-account']/tbody/tr")
                        if browser.find_element_by_xpath( ".//*[@id='table-account']/tbody/tr[2]/td[4]/div[2]").is_displayed():
                            bank_id = str(browser.find_element_by_xpath(".//*[@class='odd grouped']/td[text()='通用结算户']/following::td[1]").text)  # 获取通用结算户id
                            if len(elements) > 3:
                                self.assertFalse(False, u"该账户存在多余3个账户")
                            break
                    except NoSuchElementException, e:
                        print ""
                except NoSuchElementException, e:
                    ###########################
                    # 查找通用结算户
                    ###########################
                    elements = browser.find_elements_by_xpath(".//*[@id='table-account']/tbody/tr")
                    try:
                        if browser.find_element_by_xpath(".//*[@id='table-account']/tbody/tr[2]/td[4]/div[2]").is_displayed():
                            bank_id = str(browser.find_element_by_xpath(".//*[@class='odd grouped']/td[text()='通用结算户']/following::td[1]").text)  # 获取通用结算户
                            if len(elements) > 3:
                                self.assertFalse(False, u"该账户存在多余3个账户")
                            break
                    except NoSuchElementException, e:
                        print ""
                wait_time = wait_time + 1
                if wait_time == 50:
                    General_account = False
                    break
            #充值
            time.sleep(2)
            browser.find_element_by_xpath('''.//li[@class="nav-list account-list"]/div''').click() # 点击账务管理
            time.sleep(2)
            browser.find_element_by_xpath('''.//li[@class="nav-list account-list"]/ul/li[2]/a''').click()#点击手工记账
            time.sleep(4)
            browser.find_element_by_xpath('''.//*[@class="select2-selection__placeholder"][contains(text(),'交易场景')]''').click()
            time.sleep(4)
            browser.find_element_by_xpath('''.//input[@class="select2-search__field"]''').send_keys(u'一般户充值')
            time.sleep(2)
            browser.find_element_by_xpath('''.//li[@class="select2-results__option select2-results__option--highlighted"][contains(text(),'一般户充值')]''').click()
            time.sleep(3)
            browser.find_element_by_xpath('''.//label[@class="control-label"][contains(text(),'收款方')]/following::div/span''').click()
            time.sleep(2)
            browser.find_element_by_xpath('''.//input[@class="select2-search__field"]''').send_keys(department_register.jigou_name)#输入收款方名称
            time.sleep(4)
            path='.//li[@class="select2-results__option select2-results__option--highlighted"][contains(text(),"'''+department_register.jigou_name+'")]'#选择收款方
            browser.find_element_by_xpath(path).click()
            time.sleep(2)
            #browser.find_element_by_xpath('''.//select[@name="selectReceiveAccount"]''').click()
            time.sleep(2)
            Select(browser.find_element_by_css_selector(".form-control[data-type='receiveAccount']")).select_by_value(bank_id)
            time.sleep(3)
            browser.find_element_by_xpath('''.//input[@name="amount"]''').send_keys('100000000')
            time.sleep(2)
            browser.find_element_by_id("btn-save").click()
            time.sleep(2)
            #检验是否充值成功
            login.corp_login(self,'institution_data.xlsx')
            time.sleep(2)
            browser.find_element_by_link_text(u"账户管理").click()
            time.sleep(2)
            recharge_money=browser.find_element_by_xpath('''.//*[@class="crad-money"]/span''').text
            recharge_money=recharge_money.replace(',','')
            recharge_money=recharge_money.replace('.00','')
            if recharge_money=='1000000' :
                self.assertTrue(True,'充值成功')
            else:
                self.assertTrue(False,'充值失败')
            time.sleep(4)
        except NoSuchElementException,e:
            fp = StringIO.StringIO()  # 创建内存文件对象
            traceback.print_exc(file=fp)
            message = fp.getvalue()
            index = findStr.findStr(message, "File", 2)
            message = message[0:index]
            message = message + e.msg
            browser.get_screenshot_as_file(shot_path + browser.title + ".png")
            self.assertTrue(False, message)

    @classmethod
    def tearDownClass(cls):
        cls.browser.close()
        cls.browser.quit()