#coding=utf-8
from unittest.test import test_suite
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from classmethod import *
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
import  time
import unittest
import  re
from selenium.webdriver.common.keys import Keys
import StringIO
import traceback
import sys
import win32con
import win32com.client
import os
reload(sys)
sys.setdefaultencoding('utf8')
import csv
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
DATABASE=cf.get('database','dcf_user')
#读取核心客户注册信息
xlApp = win32com.client.Dispatch('Excel.Application')  # 打开EXCEL
xlBook = xlApp.Workbooks.Open( r'D:\\workspace\\Pythonscripts\\testdatas\\core_customer.xlsx')
xlSht = xlBook.Worksheets('Sheet1')
customer_email=xlSht.Cells(2, 5).Value
customer_name=xlSht.Cells(2, 3).Value
pattern = re.compile(r'\d*')
customer_phone=re.search(pattern, str(xlSht.Cells(2, 6).Value)).group()
customer_password=xlSht.Cells(2, 4).Value
del xlApp
#读取截图存放路径
shot_path=cf.get('shotpath','path')

class Core_Enterprise(unittest.TestCase):
    (u"核心模块")
    enterprise_ranname=""
    @classmethod
    def setUpClass(cls):
        cls.browser = webdriver.Firefox(profile)
        cls.browser.maximize_window()

    def test_1_invitation_register(self):
        (u"平台邀请注册")
        browser = self.browser
        browser.implicitly_wait(10)
        try:
            # 运营平台登录
            login.operate_login(self,"operation_login.csv")
            # browser.find_element_by_id("j_user_name").send_keys(username)
            # browser.find_element_by_id("j_password").send_keys(password)
            # browser.find_element_by_id("reg-btn").click()
            time.sleep(2)
            # 客户邀请
            browser.find_element_by_link_text("客户邀请").click()
            time.sleep(4)
            browser.find_element_by_id("inviteCustomer").click()
            time.sleep(1)
            # 客户信息填写
            Core_Enterprise.enterprise_ranname=u"测试核心企业"+str(time.strftime("%m%d%H%M%S", time.localtime()))#随机生成客户信息
            #将随机生成的客户名称写入core_customer.xlsx中
            xlApp = win32com.client.Dispatch('Excel.Application')  # 打开EXCEL
            xlBook = xlApp.Workbooks.Open(r'D:\\workspace\\Pythonscripts\\testdatas\\core_customer.xlsx')
            xlSht = xlBook.Worksheets('Sheet1')
            xlSht.Cells(2, 1).Value =Core_Enterprise.enterprise_ranname
            xlBook.Close(SaveChanges=1)  # 完成 关闭保存文件
            del xlApp
            browser.find_element_by_id("customerFullName").send_keys(Core_Enterprise.enterprise_ranname)
            #browser.find_element_by_id("customerFullName").send_keys(enterprise_name)
            browser.find_element_by_xpath(".//*[@id='inviteForm']/div[2]/div/div/div[1]/button[2]").click()
            time.sleep(2)
            browser.find_element_by_link_text(u"水利、环境和公共设施管理业").click()
            # browser.execute_script("arguments[0].scrollIntoView()",we)
            browser.find_element_by_xpath(".//*[@id='inviteForm']/div[3]/div/div/div[1]/button[2]").click()
            time.sleep(2)
            we = browser.find_element_by_css_selector("#province>li>a[value='310000']")
            browser.execute_script("arguments[0].scrollIntoView()", we)
            we.click()
            time.sleep(2)
            invitedUser = browser.find_element_by_id("invitedUser")
            invitedUser.clear()
            invitedUser.send_keys(customer_name)
            time.sleep(2)
            invitedEmail = browser.find_element_by_id("invitedEmail")
            invitedEmail.clear()
            invitedEmail.send_keys(customer_email)
            time.sleep(2)
            invitedMobile = browser.find_element_by_id("invitedMobile")
            invitedMobile.clear()
            invitedMobile.send_keys(customer_phone)
            time.sleep(2)
            # 新建邀请并注册
            browser.find_element_by_css_selector(".btn.btn-danger.createInviteBtn").click()
            time.sleep(5)
            # 防止出现JS弹出框
            try:
                invite_core = browser.find_element_by_id("invite-email-core")
            except NoSuchElementException:
                time.sleep(2)
                browser.switch_to_alert().accept()
                browser.find_element_by_css_selector(".btn.btn-danger.createInviteBtn").click()
            time.sleep(2)
            self.assertEqual(invite_core.text, u"发送邀请", "Customers can not build success")
            customer_url = browser.execute_script("return document.getElementById('inviteUrl-core').value")
            browser.get(customer_url)
            time.sleep(5)
            jiaru = browser.find_element_by_id("jiaru")
            if jiaru.is_displayed():
                self.assertTrue(True, "Generate invitation connection, invite Success!")
                jiaru.click()
            else:
                self.assertTrue(False, "Generated invitation connection,but unable to enter the registration page!")
        except Exception, e:
            fp = StringIO.StringIO()  # 创建内存文件对象
            traceback.print_exc(file=fp)
            message = fp.getvalue()
            index_file = findStr.findStr(message, "File", 2)
            index_Exception = message.find("Message")
            index_Stacktrace=message.find("Stacktrace:")
            print_message = message[0:index_file] + message[index_Exception:index_Stacktrace]
            time.sleep(1)
            title_index = browser.title.find("-")
            title = browser.title[0:title_index]
            browser.get_screenshot_as_file(shot_path + title + ".png")
            self.assertTrue(False, print_message)
    def test_2_core_register(self):
        (u"核心企业注册")
        browser = self.browser
        browser.implicitly_wait(10)
        time.sleep(4)
        try:
            # 填写注册信息
            browser.find_element_by_id("inputPassword").send_keys(customer_password)
            time.sleep(1)
            browser.find_element_by_id("inputRePassword").send_keys(customer_password)
            time.sleep(1)
            browser.find_element_by_id("getDynamic").click()
            time.sleep(5)
            # 获取验证码
            now_handle = browser.current_window_handle
            Dynamic_url = "http://" + host + ".dcfservice.com/v1/public/sms/get?cellphone=" + customer_phone
            js_script = 'window.open(' + '"' + Dynamic_url + '"' + ')'
            browser.execute_script(js_script)
            time.sleep(2)
            all_handles = browser.window_handles
            for handle in all_handles:
                if handle != now_handle:
                    browser.switch_to_window(handle)
            Dynamic_code = browser.find_element_by_css_selector("html>body>pre").text
            Dynamic_code = Dynamic_code[1:7]
            browser.switch_to_window(now_handle)
            # 填写验证码
            browser.find_element_by_id("validateCode").send_keys(Dynamic_code)
            time.sleep(1)
            browser.find_element_by_id("registerbtn").click()
            # 等待5秒进入主页面后，关闭导航页面
            time.sleep(5)
            try:
                browser.find_element_by_css_selector(".aknowledge").click()
                time.sleep(5)
                browser.find_element_by_xpath(".//*[@id='zhongjin-banner']/div[1]").click()
            except NoSuchElementException,e:
                print ""
            time.sleep(1)
            # 获取登录的用户名
            login_name = browser.find_element_by_css_selector("#logoutDiv>a").text
            if login_name == customer_name:
                self.assertTrue(True, "客户注册成功")
            else:
                self.assertTrue(False, "客户注册失败")
        except Exception, e:
            fp = StringIO.StringIO()  # 创建内存文件对象
            traceback.print_exc(file=fp)
            message = fp.getvalue()
            index_file = findStr.findStr(message, "File", 2)
            index_Exception = message.find("Message")
            index_Stacktrace = message.find("Stacktrace:")
            print_message = message[0:index_file] + message[index_Exception:index_Stacktrace]
            time.sleep(1)
            title_index = browser.title.find("-")
            title = browser.title[0:title_index]
            browser.get_screenshot_as_file(shot_path + title + ".png")
            self.assertTrue(False, print_message)
    def test_3_core_authentication(self):
        (u"核心企业认证")
        browser = self.browser
        browser.implicitly_wait(10)
        time.sleep(2)
        try:
            # 运营平台登录
            login.operate_login(self,"operation_login.csv")
            time.sleep(2)
            # 客户认证
            browser.find_element_by_link_text(u"客户管理")
            time.sleep(1)
            browser.find_element_by_link_text(u"客户认证")
            time.sleep(2)
            try:
                # 数据库连接
                conn = mysql.connector.connect(host=HOST, user=USER, passwd=PASSWORD, db=DATABASE, port=PORT)
                # 创建游标
                cur = conn.cursor()
                # customername_id查询
                sql = 'select customer_id from user where customer_name="' +Core_Enterprise.enterprise_ranname+ '"'
                cur.execute(sql)
                # 获取查询结果
                result_set = cur.fetchall()
                if result_set:
                    for row in result_set:
                        customername_id = row[0]
                else:
                    self.assertTrue(False,"the customer_id do not exsit in database!")
                # 关闭游标和连接
                cur.close()
                conn.close()
            except mysql.connector.Error, e:
                print e.message
            customername_id = str(customername_id)
            browser.find_element_by_id("search-text").send_keys(Core_Enterprise.enterprise_ranname)#输入核心企业以便查找
            time.sleep(2)
            browser.find_element_by_id("btn-search").click()
            time.sleep(2)
            browser.find_element_by_xpath(".//a[@class='nowrap'][contains(@href,'action')]").click()#点击认证
            time.sleep(4)
            browser.find_element_by_xpath(".//input[@name='commitMethod'][@value='1']").click()#选择普通版本
            time.sleep(3)
            ###############################################################################################################
            #                                       以下区域为上传照片                                                    #
            ###############################################################################################################

            web_elements=browser.find_elements_by_css_selector(".btn.btn-primary.fileinput-button")#找出所有上传文件的按钮
            time.sleep(2)
            browser.execute_script("arguments[0].scrollIntoView()",browser.find_elements_by_css_selector(".upload-box.license-box")[1])#使营业执照区域可视
            web_elements[1].click()#点击营业执照上传按钮
            time.sleep(2)
            # upload_file路径，上传图片
            upload_file = method + "\\upload.exe " + data + "test_picture.jpg"
            os.system(upload_file)
            time.sleep(2)
            browser.execute_script("arguments[0].scrollIntoView()",browser.find_elements_by_css_selector(".upload-box.license-box")[2])  # 使组织机构代码区域可视
            web_elements[2].click()  # 点击组织机构代码上传按钮
            time.sleep(1)
            # upload_file路径，上传图片
            upload_file = method + "\\upload.exe " + data + "test_picture.jpg"
            os.system(upload_file)
            time.sleep(2)
            browser.execute_script("arguments[0].scrollIntoView()", browser.find_elements_by_css_selector(".upload-box.sub-box")[2])  # 使操作者身份证正面区域可视
            web_elements[5].click()  # 点击操作者正面上传按钮
            time.sleep(2)
            # upload_file路径，上传图片
            upload_file = method + "\\upload.exe " + data + "test_picture.jpg"
            os.system(upload_file)
            time.sleep(2)
            browser.execute_script("arguments[0].scrollIntoView()", browser.find_elements_by_css_selector(".upload-box.sub-box")[3])  # 使操作者身份证反面区域可视
            web_elements[6].click()  # 点击操作者反面上传按钮
            time.sleep(2)
            # upload_file路径，上传图片
            upload_file = method + "\\upload.exe " + data + "test_picture.jpg"
            os.system(upload_file)
            time.sleep(2)
            browser.execute_script("arguments[0].scrollIntoView()", browser.find_elements_by_css_selector(".upload-box.license-box")[3])  # 使操作者手持身份证照片
            web_elements[7].click()  # 点击操作者手持身份证照片上传按钮
            time.sleep(2)
            # upload_file路径，上传图片
            upload_file = method + "\\upload.exe " + data + "test_picture.jpg"
            os.system(upload_file)
            ####################################################################################################
            #                               以下区域填写营业执照相关信息                                       #
            ####################################################################################################
            browser.execute_script("arguments[0].scrollIntoView()",browser.find_elements_by_css_selector(".crumbs-default")[1])#使审核认证资料信息可视，从而使营业执照右侧通过按钮可视
            time.sleep(2)
            browser.find_element_by_css_selector(".radio.col-xs-6>input[value='2'][name='businessLicense_Pass']").click()#点击营业执照的通过按钮
            time.sleep(2)
            browser.find_element_by_css_selector(".form-control[name='enterprise_no']").send_keys("111111")#填写营业执照号
            time.sleep(2)
            browser.find_element_by_xpath('.//input[@name="businessLicenseNeverExpireFlag"][@value="1"]').click()#营业执照无期限
            time.sleep(2)
            browser.find_element_by_xpath(".//input[@name='enterprise_money']").send_keys("10000")#填写注册资本
            time.sleep(2)
            browser.execute_script("arguments[0].scrollIntoView()",browser.find_elements_by_xpath('.//div[@class="f12"]')[1])
            time.sleep(1)
            browser.find_element_by_xpath(".//input[@name='organizationNo_Pass'][@value='2']").click()#点击组织机构代码通过
            time.sleep(2)
            browser.find_element_by_xpath(".//input[@class='form-control'][@name='organization_no']").send_keys("111111")#填写组织机构代码
            time.sleep(2)
            browser.find_element_by_xpath(".//input[@class='form-control'][@name='organization_no_regist']").send_keys("111111")#填写登记号
            time.sleep(2)
            Select(browser.find_element_by_id("province")).select_by_visible_text(u"上海")#选择上海市
            time.sleep(2)
            Select(browser.find_element_by_id("city")).select_by_visible_text(u"上海市")  # 选择上海市
            time.sleep(2)
            browser.execute_script('''arguments[0].value="2020-2-21"''', browser.find_element_by_xpath(".//input[@class='form-control datepicker']"))
            time.sleep(1)
            browser.execute_script("arguments[0].scrollIntoView()",browser.find_elements_by_xpath('.//div[@class="f12"]')[2])#上移组织机构代码，使操作者身份证区域可见
            time.sleep(1)
            browser.find_element_by_xpath(".//input[@name='operatorId_Pass'][@value='2']").click()#点击操作者身份证通过
            time.sleep(2)
            browser.find_element_by_xpath(".//input[@name='operator_user_name']").send_keys(u"周大强")#输入身份证名称
            time.sleep(2)
            browser.find_element_by_xpath(".//input[@name='operator_ID']").send_keys("222222222222222222")#输入身份证号码
            time.sleep(2)
            browser.find_element_by_xpath(".//input[@name='operator_ID_never_expire']").click()#长期
            time.sleep(2)
            browser.execute_script("arguments[0].scrollIntoView()",browser.find_elements_by_xpath('.//div[@class="f12"]')[3])  # 上移身份证操作者区域，使操作者手持身份证照片
            time.sleep(2)
            browser.find_element_by_xpath(".//input[@name='picHandling_Pass'][@value='2']").click()#点击操作者手持身份证照片通过按钮
            time.sleep(2)
            browser.execute_script("document.documentElement.scrollTop=0")  # 滑动滚动条至顶部
            time.sleep(2)
            browser.find_element_by_xpath(".//input[@name='commitMethod'][@value='1']").click()  # 再次选择普通版本
            time.sleep(2)
            browser.find_element_by_id("btnSubmit").click()  # 点击提交
            time.sleep(2)
            browser.find_element_by_css_selector("#modalFooter>a").click()#点击返回列表
            #################################
            #验证认证状态
            #################################
            browser.find_element_by_id("search-text").send_keys(Core_Enterprise.enterprise_ranname)  # 输入核心企业以便查找
            time.sleep(2)
            browser.find_element_by_id("btn-search").click()
            time.sleep(2)
            if not browser.find_element_by_xpath(".//span[@class='status certified']").text==u"已认证":
                self.assertTrue(False,"已认证成功但是状态没有改变")
            ############################################################
            #查找新建用户通用结算户，把customer_id和通用结算户传入excel#
            ############################################################
            browser.find_element_by_link_text(u"群星支付").click()
            time.sleep(2)
            ##########################
            # 等待一直loading的按钮消失
            ###########################
            wait_time = 0
            General_account = True
            while True:
                browser.find_element_by_id("search-text").clear()
                browser.find_element_by_id("search-text").send_keys(Core_Enterprise.enterprise_ranname)
                browser.find_element_by_id("btn-search").click()  # 在群星支付界面搜索核心企业账户
                time.sleep(5)
                try:
                    if browser.find_element_by_xpath("//div[@class='loading']").is_displayed():
                        browser.refresh()
                        time.sleep(5)
                    ###########################
                    # 查找通用结算户
                    ###########################
                    try:
                        elements = browser.find_elements_by_xpath(".//*[@id='table-account']/tbody/tr")
                        if browser.find_element_by_xpath( ".//*[@id='table-account']/tbody/tr[2]/td[4]/div[2]").is_displayed():
                            core_General_account = str(browser.find_element_by_xpath(".//*[@id='table-account']/tbody/tr[2]/td[4]/div[2]").text)  # 获取通用结算户
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
                            core_General_account = str(browser.find_element_by_xpath(".//*[@id='table-account']/tbody/tr[2]/td[4]/div[2]").text)  # 获取通用结算户
                            if len(elements) > 3:
                                self.assertFalse(False, u"该账户存在多余3个账户")
                            break
                    except NoSuchElementException, e:
                        print ""
                wait_time = wait_time + 1
                if wait_time == 50:
                    General_account = False
                    break
            if not General_account:
                self.assertFalse(True, "该账户已认证成功，但是没有创建通用结算户")
            core_account_id=str(browser.find_element_by_css_selector(".odd.grouped>td[data-name='accountId']").text)#获取群星id号
            xlApp = win32com.client.Dispatch('Excel.Application')  # 打开EXCEL
            xlBook = xlApp.Workbooks.Open(r'D:\\workspace\\Pythonscripts\\testdatas\\core_customer.xlsx')
            xlSht = xlBook.Worksheets('Sheet1')
            xlSht.Cells(2, 2).Value =core_General_account
            # xlSht.Cells(2,7).Value=customername_id
            xlBook.Close(SaveChanges=1)  # 完成 关闭保存文件
            del xlApp
            #############################################
            #记账方式充值                               #
            #############################################
            browser.find_element_by_link_text(u"群星支付").click()
            time.sleep(2)
            browser.find_element_by_css_selector(".nav-list.account-list>div").click()#点击账务管理
            time.sleep(2)
            browser.find_elements_by_css_selector(".nav-list.account-list>ul>li>a")[1].click()#点击手工记账
            time.sleep(5)
            wait_time=0
            while True:
                try:
                    browser.find_elements_by_css_selector(".select2-selection__arrow")[0].click()
                    #browser.find_element_by_xpath("html/body/div[1]/div[3]/div[2]/form/div[1]/div/span/span[1]/span/span[2]").click()
                    time.sleep(2)
                    browser.find_element_by_css_selector(".select2-search__field").send_keys(u"一般户充值")#选择一般户充值
                    time.sleep(2)
                    browser.find_element_by_css_selector(".select2-search__field").send_keys(Keys.ENTER)  # 按enter键输入
                    time.sleep(2)
                    browser.find_elements_by_css_selector(".select2-selection__arrow")[3].click()#点击收款方
                    time.sleep(2)
                    browser.find_element_by_css_selector(".select2-search__field").send_keys(Core_Enterprise.enterprise_ranname)  # 输入收款方
                    time.sleep(2)
                    browser.find_element_by_css_selector(".select2-search__field").send_keys(Keys.ENTER)  # 按enter键输入
                    time.sleep(2)
                    Select(browser.find_element_by_css_selector(".form-control[data-type='receiveAccount'][name='selectReceiveAccount']")).select_by_value(core_account_id)#选择通用结算户
                    time.sleep(2)
                    browser.find_element_by_css_selector(".form-control.num").send_keys(10000000)#充值完成
                    time.sleep(2)
                    browser.find_element_by_id("btn-save").click()#保存按钮
                    time.sleep(2)
                    break
                except NoSuchElementException,e:
                    browser.refresh()
                    time.sleep(5)
                wait_time=wait_time+1
                if wait_time==20:
                    break
        except Exception, e:
            fp = StringIO.StringIO()  # 创建内存文件对象
            traceback.print_exc(file=fp)
            message = fp.getvalue()
            index_file = findStr.findStr(message, "File", 2)
            index_Exception = message.find("Message")
            index_Stacktrace = message.find("Stacktrace:")
            print_message = message[0:index_file] + message[index_Exception:index_Stacktrace]
            time.sleep(1)
            title_index = browser.title.find("-")
            title = browser.title[0:title_index]
            browser.get_screenshot_as_file(shot_path + title + ".png")
            self.assertTrue(False, print_message)
    @classmethod
    def tearDownClass(cls):
        # comand = "mysql -h t6.db.dcfservice.com -uroot -pdcf2014<\"D:\workspace\Pythonscripts\classmethod\delete.sql\""
        # os.system(comand)
        cls.browser.close()
        cls.browser.quit()

# if __name__ == "__main__":
#     testsuite=unittest.TestSuite()
#     testsuite.addTest(Core_Enterprise("test_1_invitation_register"))
#     testsuite.addTest(Core_Enterprise("test_2_core_register"))
#     testsuite.addTest(Core_Enterprise("test_3_core_authentication"))
#     filename = "d:\\result.html"
#     fp = file(filename, 'wb')
#     runner = HTMLTestRunner.HTMLTestRunner(stream=fp, title='Result', description='Test_Report')
#     runner.run(testsuite)