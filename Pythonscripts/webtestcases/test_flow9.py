#coding=utf-8
from selenium import webdriver
import time
import unittest
import ConfigParser
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
import  StringIO
import traceback
from classmethod import findStr
from classmethod import login
import win32com.client
import sys
import exceptions
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

#读取单据编号receipt_id,单据金额receipt_money
#os.system('D:\\Workspace\\Pythonscripts\\testdatas\\transaction_flow.xlsx')
xlxApp = win32com.client.Dispatch('Excel.Application')  # 打开EXCEL
xlxBook=xlxApp.Workbooks.Open(r'D:\\Workspace\\Pythonscripts\\testdatas\\transaction_flow.xlsx')
xlSht = xlxBook.Worksheets('sheet1')
list_receipt_id=[]#存放单据编号
list_receipt_money=[]#存放单据金额
list_financing_money=[]#存放融资金额
list_retainage_money=[]#存放尾款
list_loan_document_id=[]
for i in range(2,5):    #获取需要还款的单据编号和金额
    receipt_id = xlSht.Cells(i, 1).Value
    receipt_money= str(xlSht.Cells(i, 4).Value)
    receipt_money=receipt_money+'0'
    list_receipt_id.append(receipt_id)
    list_receipt_money.append(receipt_money)
#读取融资金额和loan_document_id
xlSht2 = xlxBook.Worksheets('Sheet2')
for i in range(2,5):    #获取需要还款的单据对应的融资金额
    financing_money= str(xlSht2.Cells(i, 4).Value)
    financing_money=financing_money+'0'
    list_financing_money.append(financing_money)
    loan_document_id=str(xlSht2.Cells(i, 1).Value)
    list_loan_document_id.append(loan_document_id)
xlxBook.Close(SaveChanges=1)
del xlxApp
#计算尾款
for a in range(len(list_receipt_money)):
    for b in range(len(list_financing_money)):
        if a==b:
            retainage_money=float(list_receipt_money[a])-float(list_financing_money[b])
            list_retainage_money.append( retainage_money)
#读取机构名和银行账号
xlxApp = win32com.client.Dispatch('Excel.Application')  # 打开EXCEL
xlxBook=xlxApp.Workbooks.Open(r'D:\\Workspace\\Pythonscripts\\testdatas\\institution_data.xlsx')
xlSht = xlxBook.Worksheets('sheet1')
jigou_name= xlSht.Cells(2, 1).Value
jigou_bank_no=xlSht.Cells(2, 2).Value
xlxBook.Close(SaveChanges=1)
del xlxApp
#读取买方企业名称
xlxApp = win32com.client.Dispatch('Excel.Application')  # 打开EXCEL
xlxBook=xlxApp.Workbooks.Open(r'D:\\Workspace\\Pythonscripts\\testdatas\\core_customer.xlsx')
xlSht = xlxBook.Worksheets('Sheet1')
core_enterprise_name = xlSht.Cells(2, 1).Value
core_bank_no=xlSht.Cells(2,2).Value
xlxBook.Close(SaveChanges=1)
del xlxApp
#读取链属企业名称和银行账号
xlxApp = win32com.client.Dispatch('Excel.Application')  # 打开EXCEL
xlxBook=xlxApp.Workbooks.Open(r'D:\\Workspace\\Pythonscripts\\testdatas\\chain_customer.xlsx')
xlSht = xlxBook.Worksheets('Sheet1')
chain_enterprise_name = xlSht.Cells(2, 1).Value
chain_bank_no=xlSht.Cells(2,2).Value
xlxBook.Close(SaveChanges=1)
del xlxApp
#读取截图存放路径
shot_path=cf.get('shotpath','path')
class Core_Enterprise(unittest.TestCase):
    (u"核心模块")
    @classmethod
    def setUpClass(cls):
        cls.browser = webdriver.Firefox()
        cls.browser.maximize_window()
    def test1_Apply_repayment(self):
        (u"申请还款")
        browser = self.browser
        browser.implicitly_wait(10)
        try:
            login.corp_login(self,'core_customer.xlsx') #买家企业登陆
            time.sleep(2)
            browser.find_element_by_id('today').click() #切换到当天页面
            browser.find_element_by_xpath('''.//*[@class="daySelectArea"]/span[18]''').click()
            time.sleep(2)
            browser.find_element_by_xpath("//div[@id='anchorPoint']/following::div/ul/li[6]").click()#点击已融资页签
            time.sleep(2)
            #根据单据号定位需要还款的融资款
            we=browser.find_element_by_xpath("//div[@id='anchorPoint']/following::div[1]/ul")
            browser.execute_script("arguments[0].scrollIntoView()",we)
            time.sleep(2)
            for i in range(len(list_receipt_id)):
                path="//tr/td[2][text()='"+list_receipt_id[i]+"']/preceding-sibling::td"
                browser.find_element_by_xpath(path).click()
                time.sleep(2)
            time.sleep(2)
            browser.find_element_by_xpath('''.//*[@id='tableList']/div[5]/div[2]/table/tbody/tr[1]/td[2]/button''').click() #点击还款
            time.sleep(3)
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
            #循环查询,等待还款状态改变
            while count<20:
                browser.find_element_by_id('today').click() #切换到当天页面
                time.sleep(3)
                browser.find_element_by_xpath("//div[@id='anchorPoint']/following::div/ul/li[8]").click()#到已还款页面
                time.sleep(2)
                a=browser.find_element_by_xpath("//div[@id='anchorPoint']/following::div/ul/li[8]/a/span").text#获取已还款单据的总数
                if a=='('+str(len(list_receipt_id))+')':
                    break
                time.sleep(60)
                count+=1
                if count==19:
                    self.assertFalse(True,'单据状态转换失败')
            for i in range(len(list_loan_document_id)):
                path=".//*[@id='"+list_loan_document_id[i]+"']/td[2]"
                if browser.find_element_by_xpath(path).is_displayed():
                    self.assertTrue(True,list_loan_document_id[i]+'申请还款成功')
                else:
                    self.assertFalse(True,list_loan_document_id[i]+'申请还款失败')
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
    def test2_Settlement_documents(self):
        (u"检查凭证")
        browser = self.browser
        browser.implicitly_wait(10)
        try:
            login.operate_login(self,'operation_login.csv') #运营端登陆
            time.sleep(2)
            browser.find_element_by_id('qunxingPay-nav').click()
            browser.find_element_by_xpath("//div[@id='header']//following::div/ul/li[3]/div").click()  #点击凭证查询
            ###########################################################################################################
            #                              结算凭证校验                                                               #
            ###########################################################################################################
            #计算还款总额sum_money
            sum_money=[]
            sum=0
            for i in list_receipt_money:
                sum+=float(i)
            sum=str(sum)+'0'
            sum_money.append(sum)
            time.sleep(4)
            browser.find_element_by_xpath("//div[@id='header']//following::div/ul/li[3]/div/following::ul/li[3]/a[text()='结算凭证查询']").click()
            time.sleep(4)
            browser.find_element_by_xpath("//input[@id='createTime']/ancestor::div[1]/span/span/span").click()#选择时间
            time.sleep(2)
            browser.find_element_by_xpath("/html/body/div[2]/div/div/div/div/button[1][text()=' 今 天']").click()#选择今天
            time.sleep(2)
            try:
                in_str=''
                for i in list_receipt_id:
                    in_str=in_str+"'"+str(i)+"'"+","
                in_str=in_str[:-1]
                #定义存储本金、支付、尾款对应结算凭证号和付款编号的二维数组
                list_benjin=[[0]*2  for i in range (len(list_receipt_id))]
                list_weikuan=[[0]*2  for i in range (len(list_receipt_id))]
                list_zhifu=[[0,0]]
                conn = mysql.connector.connect(host=HOST,user=USER,passwd=PASSWORD,db=DATABASE,port=PORT)
                # 创建游标
                cur = conn.cursor()
                sql='SELECT b.repayment_id FROM dcf_loan.t_loan_document a LEFT JOIN dcf_loan.t_repayment_detail b on a.loan_document_id=b.document_id where a.loan_document_no IN (' +in_str+')'
                cur.execute(sql)
                #得到单据编号对应的repayment_id
                result_set = cur.fetchall()
                if result_set:
                    for row in result_set:
                        repayment_id = row[0]
                else:
                    print "No date"
                repayment_id=str(repayment_id)
            # 获取查询结果,将本金和融资对应结算流水号存到对应列表
                cur = conn.cursor()
                sql='select DISTINCT c.so_no  ,d.settle_id,c.memo   from dcf_loan.t_repayment_detail b LEFT JOIN dcf_loan.t_repayment_settlement_order c on c.repayment_id=b.repayment_id LEFT JOIN dcf_loan.t_repayment_settlement_notify_log d ON d.so_no=c.so_no where b.repayment_id="'+ repayment_id+'"'
                cur.execute(sql)
                result_set = cur.fetchall()
                if result_set:
                    a=0
                    b=0
                    for row in result_set:
                        if u"本金" in str(row[2]):
                            list_benjin[a][0]=(str(row[1]))
                            a+=1
                        elif u"尾款" in str(row[2]):
                            list_weikuan[b][0]=(str(row[1]))
                            b+=1
                        elif u"融资" in str(row[2]):
                            list_zhifu[0][0]=(str(row[1]))
                else:
                    print 'fail'
            # 关闭游标和连接
                cur.close()
                conn.close()
            except mysql.connector.Error, e:
                print e.message
            time.sleep(15)
            #######################################################
            #定义校验结算凭证支付，本金，尾款金额和结算状态 的函数#
            #######################################################
            def Query__settlement_voucher(list_settlement_no,list_money,type,payer,payee):
                for i in range(0,len(list_settlement_no)):
                    settlement_no=list_settlement_no[i][0]
                    browser.find_element_by_xpath("//div/label[text()='结算流水号']/following::input[1]").clear()
                    browser.find_element_by_xpath("//div/label[text()='结算流水号']/following::input[1]").send_keys(settlement_no)   #输入结算编号
                    time.sleep(2)
                    browser.find_element_by_xpath("//button[@id='searchBtn']").click()#查询
                    time.sleep(2)
                    money2=browser.find_element_by_xpath('''.//*[@id='settlement-result']/tbody/tr/td[9]''').text
                    money2=money2.replace(',','')
                    if money2 in list_money:
                        self.assertTrue(True,settlement_no+'结算凭证还款'+type+'金额正确')
                    else:
                        self.assertFalse(True,settlement_no+'结算凭证还款'+type+'金额错误')
                    time.sleep(2)
                    try:
                        #付款方校验
                        payer2=browser.find_element_by_xpath(".//*[@id='settlement-result']/tbody/tr[1]/td[5]").text#获取付款方
                        if payer2==payer :
                            self.assertTrue(True,settlement_no+'结算页面，支付'+type+'付款方名称正确')
                        else:
                            self.assertFalse(True,settlement_no+'结算页面，支付'+type+'付款方名称错误')
                        #收款方校验
                        payee2=browser.find_element_by_xpath(".//*[@id='settlement-result']/tbody/tr[1]/td[7]").text#获取收款方
                        if payee2==jigou_name :
                            self.assertTrue(True,settlement_no+'结算页面，支付'+type+'收款方名称正确')
                        else:
                            self.assertFalse(True,settlement_no+'结算页面，支付'+type+'收款方名称错误')
                        time.sleep(2)
                    except Exception, e:
                        fp = StringIO.StringIO()  # 创建内存文件对象
                        traceback.print_exc(file=fp)
                        message = fp.getvalue()
                        print message
                    #结算状态校验
                    for a in range(10):
                        status=browser.find_element_by_xpath("//tr/td[12]").text
                        if status==u'结算成功':
                            self.assertTrue(True,settlement_no+'结算凭证'+type+'结算成功')
                            break
                        else:
                            self.assertFalse(True,settlement_no+'结算凭证'+type+'结算失败')
                        time.sleep(10)
                    #点击详情按钮，校验详情页面
                    browser.find_element_by_xpath('''.//*[@class="showDetail"]''').click()
                    time.sleep(5)
                    detail_status=browser.find_element_by_xpath('''.//*[@class="flowTypeIcon settlementIcon"]/following::span[1]''').text
                    time.sleep(2)
                    if detail_status==u'结算成功':
                        self.assertTrue(True,settlement_no+'结算凭证详情页面显示正确')
                    else:
                        self.assertFalse(True,settlement_no+'结算凭证详情页面显示异常')
                    browser.find_element_by_xpath('''.//*[@class="close"]''').click()
                    time.sleep(2)

            Query__settlement_voucher(list_zhifu,sum_money,u'单据金额',core_enterprise_name,jigou_name)
            Query__settlement_voucher(list_benjin,list_financing_money,u'本金',jigou_name,jigou_name)

            ###########################################################################################################
            #                             付款凭证校验                                                                #
            ###########################################################################################################
            conn = mysql.connector.connect(host=HOST,user=USER,passwd=PASSWORD,db=DATABASE,port=PORT)
            # 创建游标
            cur = conn.cursor()
            sql='SELECT a.so_no,b.settle_id,c.id,b.payment_orderId from dcf_loan.t_repayment_settlement_order a LEFT JOIN dcf_loan.t_repayment_settlement_notify_log b ON a.so_no=b.so_no LEFT JOIN dcf_payment.t_payment_bank_order c ON c.invokeId=b.payment_orderId where a.repayment_id="'+ repayment_id+'"'
            cur.execute(sql)
            #得到单据编号对应的付款编号并存入二维数组
            result_set = cur.fetchall()
            if result_set:
                for row in result_set:
                    for i in range(0,len(list_benjin)):
                         benjin_STL=list_benjin[i][0]
                         if benjin_STL in str(row[1]):
                             list_benjin[i][1]=(str(row[2]))
                         elif list_zhifu[0][0] in str(row[1]):
                             list_zhifu[0][1]=(str(row[2]))
                    for j in range(0,len(list_weikuan)):
                         weikuan_STL=list_weikuan[j][0]
                         if weikuan_STL in str(row[1]):
                             list_weikuan[j][1]=(str(row[2]))
                print  list_zhifu,list_benjin,list_weikuan
            else:
                 print 'fail'
            browser.find_element_by_xpath("//div[@id='header']//following::div/ul/li[3]/div").click()  #点击凭证查询
            time.sleep(2)
            browser.find_element_by_xpath("//div[@id='header']//following::div/ul/li[3]/div/following::ul/li[2]/a[text()='付款凭证查询']").click()
            time.sleep(4)
            #######################################################
            #定义校验付款凭证支付，本金，尾款金额和付款状态 的函数#
            #######################################################
            #参数为存付款编号的数组，存金额的数组，金额类型，付款方，收款方
            def Query_pay_voucher(list_pay_no,list_money,type,payer,payee):
                for k in range(0,len(list_pay_no)):
                    pay_no=list_pay_no[k][1]
                    browser.find_element_by_id("payNum").clear()
                    browser.find_element_by_id("payNum").send_keys(pay_no)      #输入金额对应付款编号
                    browser.find_element_by_xpath("//button[@id='searchBtn']").click()      #搜索
                    time.sleep(2)
                    money4=browser.find_element_by_xpath(".//*[@id='"+pay_no+"']/td[7]").text
                    money4=money4.replace(',','')
                    if money4  in  list_money:
                        self.assertTrue(True,pay_no+'付款凭证'+type+'金额正确')
                    else:
                        self.assertTrue(False,pay_no+'付款凭证'+type+'金额错误')
                    time.sleep(2)
                    try:
                        #付款方校验
                        payer4=browser.find_element_by_xpath(".//*[@id='"+pay_no+"']/td[3]").text#获取付款方
                        if payer4==payer :
                            self.assertTrue(True,pay_no+'付款凭证页面，支付'+type+'付款方名称正确')
                        else:
                            self.assertFalse(True,pay_no+'付款凭证页面，支付'+type+'付款方名称错误')
                        #收款方校验
                        payee4=browser.find_element_by_xpath(".//*[@id='"+pay_no+"']/td[5]").text#获取付款方
                        if payee4==payee:
                            self.assertTrue(True,pay_no+'付款凭证页面，支付'+type+'收款方名称正确')
                        else:
                            self.assertFalse(True,pay_no+'付款凭证页面，支付'+type+'收款方名称错误')
                        time.sleep(2)
                    except Exception, e:
                        fp = StringIO.StringIO()  # 创建内存文件对象
                        traceback.print_exc(file=fp)
                        message = fp.getvalue()
                        print message
                    #付款状态校验
                    for a in range(100):
                        status=browser.find_element_by_xpath("//tr/td[11]").text
                        if status==u'已付款'or status==u'无需付款' :
                            self.assertTrue(True,pay_no+'付款凭证'+type+'付款状态成功')
                            break
                        else:
                            self.assertFalse(True,pay_no+'付款凭证'+type+'付款状态失败')
                        time.sleep(10)
                    #校验详情页面显示是否正确
                    browser.find_element_by_xpath('''.//*[@class="showDetail"]''').click()
                    time.sleep(5)
                    detail_status=browser.find_element_by_xpath('''.//*[@class="pay-seal"]''')
                    try:
                        if detail_status.is_displayed():
                            self.assertTrue(True,pay_no+'付款凭证查看详情页面显示'+type+'已付款')
                        else:
                            self.assertFalse(True,pay_no+'付款凭证查看详情页面显示'+type+'无需付款')
                    except Exception, e:
                        fp = StringIO.StringIO()  # 创建内存文件对象
                        traceback.print_exc(file=fp)
                        message = fp.getvalue()
                        print message
                    browser.find_element_by_xpath('''.//*[@class="close"]''').click()
                    time.sleep(2)

            Query_pay_voucher(list_zhifu,sum_money,u'支付',core_enterprise_name,jigou_name)
            Query_pay_voucher(list_benjin,list_financing_money,u'本金',jigou_name,jigou_name)
            Query_pay_voucher(list_weikuan,list_retainage_money,u'尾款',jigou_name,chain_enterprise_name)

            ###########################################################################################################
            #                             账户明细 查看                                                               #
            ###########################################################################################################
            #定义检查账户明细函数
            # 参数为：企业名称，银行账号，交易类型（收入/支出），记账类型，存放金额的数组
            def Bank_detail(enterprise_name,bank_no,trade_type,type,list_money):
                browser.find_element_by_xpath('''.//li[@class="nav-list account-list"]/div''').click()
                time.sleep(2)
                browser.find_element_by_link_text(u"账户总览").click()
                time.sleep(5)
                browser.find_element_by_xpath("//input[@id='search-text']").send_keys(enterprise_name)#输入还款企业名称
                browser.find_element_by_id("btn-search").click()
                time.sleep(2)
                browser.find_element_by_xpath(".//*[@class='nowrap'][text()='通用结算户']/following::td[2]/div[text()='"+bank_no+"']/ancestor::td/following::td[5]/div/a[1]").click()#点击收支明细
                time.sleep(10)
                browser.find_element_by_xpath("//span[text()='记账时间：']/ancestor::span/div/span/span/span/span/span[2]").click()#选择记账时间
                browser.find_element_by_xpath("//button[text()=' 今 天']").click()
                time.sleep(2)
                browser.find_element_by_id("select_capital").click() #选择交易方向
                time.sleep(2)
                if trade_type==u'收入':
                    browser.find_element_by_xpath(".//button[@id='select_capital']/following::ul/li[3]/a[contains(text(),'"+trade_type+"')]").click()
                elif trade_type==u'支出':
                    browser.find_element_by_xpath(".//button[@id='select_capital']/following::ul/li[4]/a[contains(text(),'"+trade_type+"')]").click()
                time.sleep(4)
                browser.find_element_by_xpath("//button[@id='select_trade']").click()#选择记账类型
                time.sleep(4)
                if type==u'支付':
                    browser.find_element_by_xpath("//button[@id='select_trade']/following::ul[1]/li[4]/a[text()='"+type+"']").click()#筛选出金额类型
                elif type==u'支付（本金）':
                    browser.find_element_by_xpath("//button[@id='select_trade']/following::ul[1]/li[7]/a[text()='"+type+"']").click()#筛选出金额类型
                elif type==u'支付（尾款）':
                    browser.find_element_by_xpath("//button[@id='select_trade']/following::ul[1]/li[9]/a[text()='"+type+"']").click()#筛选出金额类型
                #获取金额并校验
                for i in range(1,len(list_money)):
                    i=str(i)
                    payment=browser.find_element_by_xpath('''.//*[@id='table-details']/tbody/tr['''+i+''']/td[4]''').text
                    payment=payment.replace(',','')
                    try:
                        if payment in list_money:
                            self.assertTrue(True,'账户收支明细'+type+'金额正确')
                        else:
                            self.assertFalse(True,'账户收支明细'+type+'金额错误')
                    except Exception, e:
                        fp = StringIO.StringIO()  # 创建内存文件对象
                        traceback.print_exc(file=fp)
                        message = fp.getvalue()
                        print message
                    time.sleep(2)

            #检查还款账户支出金额明细
            Bank_detail(core_enterprise_name,core_bank_no,u'支出',u'支付',sum_money)
            #检查机构账户收入本金明细
            Bank_detail(jigou_name,jigou_bank_no,u'收入',u'支付（本金）',list_financing_money)
            #检查回款账户收入金额明细

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







