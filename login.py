#参考文档https://zhuanlan.zhihu.com/p/27115580  如何使用selenium
#https://blog.csdn.net/jojoy_tester/article/details/53453888 定位方式
#https://blog.csdn.net/qq_36962569/article/details/77283063
#https://www.zhihu.com/question/34506326
#https://blog.csdn.net/huilan_same/article/details/52200586
from selenium import webdriver
import time
from selenium.webdriver.common.action_chains import ActionChains
import win32api
from selenium.common.exceptions import NoSuchElementException
class Login:
    def __init__(self):
        self.driver = webdriver.Chrome()
    def login(self):
        self.driver.get('http://mail.163.com/')
        self.driver.maximize_window()
        # 设置等待时间，等待页面加载之后再操作
        time.sleep(10)
        # frame中实际上是嵌入了另一个页面，而webdriver每次只能在一个页面识别,需先跳转到iframe框架
        self.driver.switch_to_frame('x-URS-iframe')
        # 设置等待时间，等待页面加载之后再操作
        time.sleep(5)
        elem_user = self.driver.find_element_by_name('email')
        elem_user.clear()
        time.sleep(1)
        elem_user.send_keys('liutengyu1989')
        elem_pwd = self.driver.find_element_by_name('password')
        elem_pwd.clear()
        elem_pwd.send_keys('lty12280827')
        elem=self.driver.find_element_by_id("dologin")
        elem.click()
        #self.driver.minimize_window()
        time.sleep(10)
    def switch_to_addressee(self):
        # 退出frame，没有这一句后续的元素定位会出错,找不到oz0的元素
        self.driver.switch_to.default_content()
        if  self.driver.find_element_by_class_name("oz0"):#收件箱
            self.driver.find_element_by_class_name("oz0").click()
        time.sleep(3)
        #只显示未读邮件,若存在未读邮件可以找到元素
        if self.driver.find_element_by_xpath("/html/body/div[2]/div[1]/div[2]/div/div/div[3]/div[1]/span[1]/a"):
            ele=self.driver.find_element_by_xpath("/html/body/div[2]/div[1]/div[2]/div/div/div[3]/div[1]/span[1]/a")
            unread_count=ele.text[2]
            ele.click()
            time.sleep(8)
            return int(unread_count)
        else:
            return 0
    def dowlond_attach(self,unread_count,key_words):
        i=0
        while 1:
            try:
                ele=self.driver.find_element_by_xpath('/html/body/div[2]/div[1]/div[2]/div/div/div[4]/div['+str(i+2)+']')
            except NoSuchElementException as msg:
                print(msg)
                break
            i=i+1
            print(i+1)
            flag=0
            flag=ele.text.find(key_words)
            time.sleep(1)
            if flag is not -1:
                ele.click()
                #time.sleep(1)
                time.sleep(10)
                #找到附件
                ele=self.driver.find_element_by_link_text("查看附件")
                chain = ActionChains(self.driver)
                chain.move_to_element(ele)
                chain.pause(2)
                chain.click()
                time.sleep(2)
                self.driver.execute_script("arguments[0].scrollIntoView(true);",ele)
                time.sleep(2)
                ele.click()
                #ele=self.driver.find_element_by_css_selector(u".hx0>span:nth-child(1)>span:nth-child(3)")
                #ele.click()
                time.sleep(5)
                elem=self.driver.find_element_by_class_name("cf0")
                chain = ActionChains(self.driver)
                chain.move_to_element(ele)
                chain.pause(2)
                chain.click()
                time.sleep(2)
                self.driver.execute_script("arguments[0].scrollIntoView(true);",ele)
                #模拟鼠标悬停显示隐藏元素
                time.sleep(3)
                chain = ActionChains(self.driver)
                chain.move_to_element(elem).perform()
#                ActionChains(self.driver).move_to_element(elem).perform()
                time.sleep(5)
                #找到附件进行下载
                self.driver.find_element_by_css_selector("div.bU1>a.js-component-link.cK0").click()
                #使用pywin32操作键盘,执行下载
                time.sleep(2)
                win32api.keybd_event(0x28,0,0,0)
                time.sleep(1)
                win32api.keybd_event(0x0D,0,0,0)
                #返回邮件列表
                self.driver.switch_to.default_content()
                time.sleep(5)
                self.driver.find_element_by_xpath("/html/body/div[2]/div[1]/div[3]/header/div/div[1]/div/span").click()
                time.sleep(5)
    def quit(self):
        self.driver.quit()
        
        

