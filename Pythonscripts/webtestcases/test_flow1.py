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
import  random
import HTMLTestRunner
import StringIO
import traceback
import sys
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
USER=cf.get('dcf_user','user')
HOST=cf.get('dcf_user','host')
PASSWORD=cf.get('dcf_user','password')
PORT=cf.get('dcf_user','port')
DATABASE=cf.get('dcf_user','database')
#读取核心客户注册信息
csvfile = file(data+'\core_enterprise_customer.csv', 'rb')
reader = csv.reader(csvfile)
for line in reader:
    if reader.line_num==1:
        enterprise_name = line[0].decode('utf-8')
        customer_name = line[1].decode('utf-8')
        customer_email=line[2].decode('utf-8')
        customer_phone=line[3].decode('utf-8')
#获取核心企业登录密码
csvfile = file(data+'\core_enterprise_password.csv', 'rb')
reader = csv.reader(csvfile)
for line in reader:
    enterprise_password = line[0]
#读取截图存放路径
shot_path=cf.get('shotpath','path')

class Core_Enterprise(unittest.TestCase):
    (u"核心模块")
    enterprise_ranname=""
    @classmethod
    def setUpClass(cls):
        cls.browser = webdriver.Firefox()
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
            Core_Enterprise.enterprise_ranname=enterprise_name+str(random.randrange(1,100000))#随机生成客户信息
            #将随机生成的客户名称写入core_random_customer中
            csv_random_customer = file(data+'core_random_customer.csv', 'wb')
            writer=csv.writer(csv_random_customer)
            writer.writerow([Core_Enterprise.enterprise_ranname,customer_name,enterprise_password,customer_phone])
            csv_random_customer.close()
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
        except NoSuchElementException,e:
            fp = StringIO.StringIO()  # 创建内存文件对象
            traceback.print_exc(file=fp)
            message = fp.getvalue()
            index = findStr.findStr(message, "File", 2)
            message = message[0:index]
            message = message + e.msg
            browser.get_screenshot_as_file(shot_path + browser.title + ".png")
            self.assertTrue(False, message)
    def test_2_core_register(self):
        (u"核心企业注册")
        browser = self.browser
        browser.implicitly_wait(10)
        time.sleep(4)
        try:
            # 填写注册信息
            browser.find_element_by_id("inputPassword").send_keys(enterprise_password)
            time.sleep(1)
            browser.find_element_by_id("inputRePassword").send_keys(enterprise_password)
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
            browser.find_element_by_css_selector(".aknowledge").click()
            time.sleep(5)
            browser.find_element_by_xpath(".//*[@id='zhongjin-banner']/div[1]").click()
            time.sleep(1)
            # 获取登录的用户名
            login_name = browser.find_element_by_css_selector("#logoutDiv>a").text
            if login_name == customer_name:
                self.assertTrue(True, "客户注册成功")
            else:
                self.assertTrue(False, "客户注册失败")
        except NoSuchElementException,e:
            fp = StringIO.StringIO()  # 创建内存文件对象
            traceback.print_exc(file=fp)
            message = fp.getvalue()
            index = findStr.findStr(message, "File", 2)
            message = message[0:index]
            message = message + e.msg
            browser.get_screenshot_as_file(shot_path + browser.title + ".png")
            self.assertTrue(False, message)
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
            path = "//*[@id='table-account']/tbody/tr[@data-customerid=" + '"' + customername_id + '"]/td[6]/div/a[2]'
            browser.find_element_by_xpath(path).click()
            time.sleep(2)
            ###############################################################################################################
            #                                       以下区域为上传照片                                                    #
            ###############################################################################################################
            # 通过滑动使上传营业执照按钮可视
            businessLicenseFileupload = browser.find_element_by_xpath(".//*[@id='content-form']/div[2]/div[2]/div[4]/div[2]")
            browser.execute_script("arguments[0].scrollIntoView()", businessLicenseFileupload)
            time.sleep(2)
            browser.find_element_by_xpath(".//*[@id='content-form']/div[2]/div[1]/div/div[1]/div[1]").click()
            time.sleep(2)
            # upload_file路径，上传图片
            upload_file = method + "\\upload.exe "+data+"test_picture.jpg"
            os.system(upload_file)
            time.sleep(2)
            # 通过滑动使上传组织机构代码按钮可视
            organizationNoFileupload = browser.find_element_by_xpath(".//*[@id='orgPriImgContainer']/div[1]/div[1]")
            browser.execute_script("arguments[0].scrollIntoView()", organizationNoFileupload)
            time.sleep(2)
            browser.find_element_by_xpath(".//*[@id='orgPriImgContainer']/div[1]/div[1]/div[1]").click()
            time.sleep(2)
            # upload_file路径，上传图片
            upload_file = method + "\\upload.exe " + data + "test_picture.jpg"
            os.system(upload_file)
            time.sleep(2)
            # 通过滑动使上传身份证正面按钮可视
            operatorIdFrontFileupload = browser.find_element_by_xpath(
                ".//*[@id='content-form']/div[4]/div[1]/div[1]/div[1]")
            browser.execute_script("arguments[0].scrollIntoView()", operatorIdFrontFileupload)
            time.sleep(2)
            browser.find_element_by_xpath(".//*[@id='content-form']/div[4]/div[1]/div[1]/div[1]/div[1]").click()
            time.sleep(2)
            # upload_file路径，上传图片
            upload_file = method + "\\upload.exe " + data + "test_picture.jpg"
            os.system(upload_file)
            time.sleep(2)
            # 通过滑动使上传身份证反面按钮可见
            operatorIdBackFileupload = browser.find_element_by_xpath(".//*[@id='content-form']/div[4]/div[1]/div[2]/div[1]")
            browser.execute_script("arguments[0].scrollIntoView()", operatorIdBackFileupload)
            time.sleep(2)
            browser.find_element_by_xpath(".//*[@id='content-form']/div[4]/div[1]/div[2]/div[1]/div[1]").click()
            time.sleep(2)
            # upload_file路径，上传图片
            upload_file = method + "\\upload.exe " + data + "test_picture.jpg"
            os.system(upload_file)
            time.sleep(2)
            # 通过滑动使操作者手持身份证照片按钮可见
            picHandlingFileupload = browser.find_element_by_xpath(".//*[@id='content-form']/div[5]/div[1]/div/div[1]")
            browser.execute_script("arguments[0].scrollIntoView()", picHandlingFileupload)
            time.sleep(2)
            browser.find_element_by_xpath(".//*[@id='content-form']/div[5]/div[1]/div/div[1]/div[1]").click()
            time.sleep(2)
            # upload_file路径，上传图片
            upload_file = method + "\\upload.exe " + data + "test_picture.jpg"
            os.system(upload_file)
            #上传完所有图片后，通过控制人员信息栏让营业执照的通过按钮可见
            member=browser.find_element_by_xpath("html/body/div[1]/div[2]/div[2]")
            browser.execute_script("arguments[0].scrollIntoView()", member)
            time.sleep(3)
            ####################################################################################################
            #                               以下区域填写营业执照相关信息                                       #
            ####################################################################################################
            # 点击业执照区域的通过按钮，弹出编辑区域
            browser.find_element_by_xpath(".//*[@id='content-form']/div[2]/div[2]/div[1]/div[2]/div/label[1]/input").click()
            time.sleep(1)
            # 使营业执照区域可视
            businessLicenseFileupload = browser.find_element_by_xpath(".//*[@id='content-form']/div[2]/div[2]")
            browser.execute_script("arguments[0].scrollIntoView()", businessLicenseFileupload)
            time.sleep(4)
            # 编辑营业执照
            browser.find_element_by_xpath(".//*[@id='content-form']/div[2]/div[2]/div[2]/div[2]/div/input").send_keys("1111112")#填写营业执照注册号
            time.sleep(1)
            browser.find_element_by_css_selector('''.radio.col-xs-6>label>input[value="0"]''').click()
            time.sleep(1)
            start_date=browser.find_element_by_xpath(".//*[@id='content-form']/div[2]/div[2]/div[2]/div[3]/div/div[2]/input[1]")
            browser.execute_script('''arguments[0].value="2012-2-21"''',start_date)
            time.sleep(3)
            end_date=browser.find_element_by_xpath(".//*[@id='content-form']/div[2]/div[2]/div[2]/div[3]/div/div[2]/input[2]")
            browser.execute_script('''arguments[0].value="2020-2-21"''',end_date)
            time.sleep(3)
            browser.find_element_by_xpath(".//*[@id='content-form']/div[2]/div[2]/div[2]/div[4]/div/div/div[2]/input").send_keys("10000")
            time.sleep(2)


            # 上移营业执照区域使组织机构代码通过按钮可见
            businessLicense_move = browser.find_element_by_xpath(".//*[@id='content-form']/div[2]/div[2]/div[4]/div[2]")
            browser.execute_script("arguments[0].scrollIntoView()", businessLicense_move)
            time.sleep(2)
            # 点击组织机构区域的通过按钮，弹出编辑区域
            browser.find_element_by_xpath(".//*[@id='content-form']/div[3]/div[2]/div[1]/div[2]/div/label[1]/input").click()
            time.sleep(1)
            # 使组织机构区域可视
            organizationNoFileupload = browser.find_element_by_xpath(".//*[@id='content-form']/div[3]/div[2]")
            browser.execute_script("arguments[0].scrollIntoView()", organizationNoFileupload)
            time.sleep(4)
            # 编辑组织机构
            browser.find_element_by_xpath(".//*[@id='content-form']/div[3]/div[2]/div[2]/div[1]/div[2]/div/input").send_keys("111111")
            time.sleep(1)
            browser.find_element_by_xpath(".//*[@id='content-form']/div[3]/div[2]/div[2]/div[1]/div[3]/div/input").send_keys("121212")
            time.sleep(1)
            Select(browser.find_element_by_id("province")).select_by_value("310000")#选择上海市
            time.sleep(1)
            Select(browser.find_element_by_id("city")).select_by_value("310100")
            time.sleep(2)
            validity_date=browser.find_element_by_xpath(".//*[@id='content-form']/div[3]/div[2]/div[2]/div[1]/div[6]/div/input")
            browser.execute_script('''arguments[0].value="2020-2-21"''',validity_date)


            # 上移组织机构区域使操作者身份证区域的通过按钮可见
            organization_move = browser.find_element_by_xpath(".//*[@id='content-form']/div[3]/div[2]/div[4]/div[2]")
            browser.execute_script("arguments[0].scrollIntoView()", organization_move)
            time.sleep(2)
            # 点击操作身份证区域的通过按钮，弹出编辑区域
            browser.find_element_by_xpath(".//*[@id='content-form']/div[4]/div[2]/div[1]/div[2]/div/label[1]/input").click()
            time.sleep(1)
            # 编辑操作身份证区域
            browser.find_element_by_xpath(".//*[@id='content-form']/div[4]/div[2]/div[2]/div[1]/div/input").send_keys(u"周大强")
            time.sleep(2)
            browser.find_element_by_xpath(".//*[@id='content-form']/div[4]/div[2]/div[2]/div[2]/div/input").send_keys("340823198876611221")
            time.sleep(2)
            idcard_startdate=browser.find_element_by_xpath(".//*[@id='content-form']/div[4]/div[2]/div[2]/div[3]/div/input[1]")
            browser.execute_script('''arguments[0].value="2016-4-1"''', idcard_startdate)
            time.sleep(2)
            idcard_enddate=browser.find_element_by_xpath(".//*[@id='content-form']/div[4]/div[2]/div[2]/div[3]/div/input[2]")
            browser.execute_script('''arguments[0].value="2020-4-1"''', idcard_enddate)
            time.sleep(2)

            # 上移身份证操作区域使操作者手持身份证区域的通过按钮可视
            operator_move = browser.find_element_by_xpath(".//*[@id='content-form']/div[4]/div[2]/div[4]/div[2]")
            browser.execute_script("arguments[0].scrollIntoView()", operator_move)
            time.sleep(2)
            # 点击手持身份证区域的通过按钮
            browser.find_element_by_xpath(".//*[@id='content-form']/div[5]/div[2]/div[1]/div[2]/div/label[1]/input").click()
            time.sleep(2)
             #填写完所有资料保存
            browser.find_element_by_id("btnSubmit").click()
            time.sleep(3)
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