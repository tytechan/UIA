# encoding = utf - 8
from localSDK.BasicFunc import PublicFunc
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
# from config.ErrConfig import CNBMException, handleErr, capture_screen
from config.ErrConfig import *
from time import sleep
import time

PF = PublicFunc()

class ObjectMap:
    ''' 用于存放定位元素及操作的基本方法 '''
    def __init__(self, driver):
        self.driver = driver
        self.dict = {
            'id': ['ID', 'locatorExpression'],
            'name': ['NAME', 'locatorExpression'],
            'classname': ['CLASS_NAME', 'locatorExpression'],
            'link_text': ['LINK_TEXT', 'locatorExpression'],
            'xpath': ['XPATH', 'locatorExpression'],
            'css_selector': ['CSS_SELECTOR', 'locatorExpression'],
            'partial_link_text': ['PARTIAL_LINK_TEXT', 'locatorExpression'],
            'value': ['XPATH', '\"//*[contains(@value,\'\"+locatorExpression+\"\')]\"'],
            'text': ['XPATH', '\"//*[text()=\'\"+locatorExpression+\"\']\"'],
        }

    def findElebyMethod(self, locateType, locatorExpression, errInfo="", timeout=10):
        errInfo = "%s 秒内未找到 %s 为 %s 的元素！" %(timeout, locateType, locatorExpression) + errInfo
        try:
            locateType = locateType.lower()
            objStr = 'WebDriverWait(self.driver, timeout).until(' \
                     'EC.visibility_of_element_located((By.%s, ' \
                     '%s)), errInfo)' %(self.dict[locateType][0], self.dict[locateType][1])
            specify_element = eval(objStr)

            return specify_element
        except Exception as e:
            raise e


    def findElesbyMethod(self, locateType, locatorExpression, errInfo="", timeout=10):
        errInfo = "%s 秒内未找到 %s 为 %s 的元素！" %(timeout, locateType, locatorExpression) + errInfo
        try:
            locateType = locateType.lower()
            objStr = 'WebDriverWait(self.driver, timeout).until(' \
                     'EC.presence_of_all_elements_located((By.%s, ' \
                     'locatorExpression)), errInfo)' %self.dict[locateType][0]
            specify_elements = eval(objStr)

            return specify_elements
        except Exception as e:
            raise e


    def findEleByDetail(self, locateType, locatorExpression, timeout=10):
        errInfo = "%s 秒内未找到 %s 为 %s 的元素！" %(timeout, locateType, locatorExpression)
        try:
            element = WebDriverWait(self.driver, timeout).until(lambda x: x.find_element(by = locateType, value = locatorExpression), errInfo)
            return element
        except Exception as e:
            raise e

    def findElesByDetail(self, locateType, locatorExpression, timeout=10):
        errInfo = "%s 秒内未找到 %s 为 %s 的元素！" %(timeout, locateType, locatorExpression)
        try:
            elements = WebDriverWait(self.driver, timeout).until(lambda x: x.find_elements(by = locateType, value = locatorExpression), errInfo)
            return elements
        except Exception as e:
            raise e

    def moveMouse(self, myX,myY):
        try:
            from pymouse import PyMouse
            m = PyMouse()
            m.position()
            m.move(myX, myY)
        except Exception as e:
            raise e

    def highlight(self, element, timeToDisappear=None):
        style = element.get_attribute("style")
        self.driver.execute_script("arguments[0].setAttribute('style',arguments[1]);",
                              element,"border:2px solid red;")
                              # element,"background:green;border:2px solid red;")
        if timeToDisappear:
            while timeToDisappear:
                time.sleep(1)
                timeToDisappear -= 1
            self.setAttribute(element, "style", style)

    def addAttribute(self, element, attributeName, value):
        self.driver.execute_script("arguments[0].%s = arguments[1]" %attributeName, element, value)

    def setAttribute(self, element, attributeName, value):
        self.driver.execute_script("arguments[0].setAttribute(arguments[1],arguments[2])",
                              element, attributeName, value)

    def removeAttribute(self, element, attributeName):
        self.driver.execute_script("arguments[0].removeAttribute(arguments[1])",
                              element, attributeName)


class KeyboardKeys:
    ''' 模拟键盘按键类
    （瞬时按键未封装，可直接调用）如：
    obj.send_keys(Keys.CONTROL, Keys.SHIFT, "x")
    '''
    VK_CODE = {
        'enter': 0x0D,
        'shift': 0x10,
        'ctrl': 0x11,
        'a': 0x41,
        'c': 0x43,
        'v': 0x56,
        'x': 0x58,
        'F1': 112,
        'F2': 113,
        'F3': 114,
        'F4': 115,
        'F5': 116,
        'F6': 117,
        'F7': 118,
        'F8': 119,
        'F9': 120,
        'F10': 121,
        'F11': 122,
        'F12': 123,
        "Tab": 9,
    }

    @staticmethod
    def keyDown(keyName):
        # 按下按键（长按）
        win32api.keybd_event(KeyboardKeys.VK_CODE[keyName],0,0,0)

    @staticmethod
    def keyUp(keyName):
        # 释放按键
        win32api.keybd_event(KeyboardKeys.VK_CODE[keyName],0,win32con.KEYEVENTF_KEYUP,0)

    @staticmethod
    def oneKey(key):
        # 模拟单个按键
        KeyboardKeys.keyDown(key)
        KeyboardKeys.keyUp(key)

    @staticmethod
    def twoKeys(key1,key2):
        # 模拟两个组合按键
        KeyboardKeys.keyDown(key1)
        KeyboardKeys.keyDown(key2)
        KeyboardKeys.keyUp(key2)
        KeyboardKeys.keyUp(key1)


class WaitUtil:
    # 映射定位方式字典对象
    def __init__(self,driver):
        self.locationTyoeDict = {
            "xpath":By.XPATH,
            "id":By.ID,
            "name":By.NAME,
            "css_selector":By.CSS_SELECTOR,
            "class_name":By.CLASS_NAME,
            "tag_name":By.TAG_NAME,
            "link_text":By.LINK_TEXT,
            "partial_link_text":By.PARTIAL_LINK_TEXT
        }
        # 初始化driver
        self.driver = driver
        # 创建显式等待实例对象,等待时间待全局化处理
        self.wait = WebDriverWait(self.driver, 10)

    def presenceOfElementLocated(self,locatorMethod, locatorExpression):
        '''
        显式等待页面元素出现在DOM中，但不一定可见，存在则返回元素对象
        :param locatorMethod: 定位方法
        :param locatorExpression: 定位表达式
        :param args:
        :return: 页面元素对象
        '''
        errInfo = "未找到 '%s' 为 '%s' 的元素！" %(locatorMethod, locatorExpression)
        try:
            if locatorMethod.lower() in self.locationTyoeDict:
                element = self.wait.until(
                    EC.presence_of_element_located((self.locationTyoeDict[locatorMethod.lower()], locatorExpression)), errInfo)
                return element
            else:
                raise TypeError(u'未找到定位方式，请确认定位方法使用是否正确')
        except Exception as e:
            raise e

    def visibilityOfElementLocated(self,locationType, locationExpression):
        '''
        显式等待页面元素出现在DOM中，并且可见，存在则返回该页面元素对象
        :param locationType: 定位方法
        :param locationExpression: 定位表达式
        :param args:
        '''
        errInfo = "未找到 '%s' 为 '%s' 的元素！" %(locationType, locationExpression)
        try:
            el = self.wait.until(EC.visibility_of_element_located((self.locationTyoeDict[locationType.lower()],
                                                                   locationExpression)), errInfo)
            print('********** 元素是否可操作：',el.is_enabled()," **********")
            sleep(0.5)
            # return el
        except Exception as e:
            raise e

    def frameToBeAvailableAndSwitchToIt(self,locationType, locationExpression):
        '''
        检查frame是否存在，存在则切换到frame控件中
        :param locationType: 定位方法
        :param LocationExpression: 定位表达式
        :param args:
        :return: None
        '''
        errInfo = "未找到 '%s' 为 '%s' 的元素！" %(locationType, locationExpression)
        try:
            self.wait.until(
                EC.frame_to_be_available_and_switch_to_it((self.locationTyoeDict[locationType.lower()],locationExpression)), errInfo)
        except Exception as e:
            # 抛出异常给上层调用者
            raise e


class PageAction:
    '''
    1、浏览器操作:open_browser、visit_url、close_browser、close_page、switch_to_frame、switch_to_default_content、
                maximize_browser、switch_to_now_window、refresh_page、scroll_slide_field；
    2、常规操作：clear、specObjClear、click_Obj、click_SpecObj、sendkeys_To_Obj、sendkeys_To_SpecObj、sendkeys_to_elements、SelectValues、
        xpath_combination_click、xpath_combination_click_loop、xpath_combination_send_keys、xpath_combination_click_send_keys_loop、
        xpath_combination_send_keys_click_loop、menu_select、
        capture_screen（setValueByTextAside、selectValueByTextAside,capture_screen_old）；
    3、辅助定位：highlightElement、highlightElements、whichIsEnabled、whichIsDisplayed；
    4、获取信息：getTitle、getPageSource、getAttribute、getDate_Now、getDateCalcuated、getTextInTable；
    5、断言及判断：assert_string_in_pagesourse、assert_title、assert_list；
    6、剪贴板操作：paste_string、press_key；
    7、等待：loadPage、sleep、waitPresenceOfElementLocated、waitVisibilityOfElementLocated、wait_elements_vanish
            waitFrameToBeAvailableAndSwitchToIt；
    8、鼠标键盘模拟：moveToElement、init_Mouse、pageKeySimulate、get_clipboard_return；
    9、外部程序调用：runProcessFile、page_upload_file；
    10、字符串操作：randomNum、pinyinTransform、compose_JSON；
    11、带判断关键字：ifExistThenClick、ifExistThenSendkeys、BoxHandler、ifExistThenSelect、ifExistThenSetData、ifExistThenReturnAttribute_pinyin、
        ifExistThenReturnOperateValue、ifExistThenChooseOperateValue、ifExistThenChooseOperateValue_diffPosition、
        ifExistThenPass_xpath_combination
    12、JS相关：setDataByJS；
    '''

    _driver = None
    def __init__(self):
        # self.driver = None
        self.timeout = 10

    @CNBMException
    def open_browser(self, browserName):
        ''' 打开指定类型浏览器
        :param browserName: 浏览器类型（ie/chrome/edge/firefox）
        :return: 全局变量driver
        '''
        try:
            global waitUtil
            if browserName.lower() == 'ie':
                driver = webdriver.Ie()
            elif browserName.lower() == 'chrome':
                chrome_options = webdriver.ChromeOptions()
                # 用于控制进程是否在后台执行
                # chrome_options.add_argument('--headless')
                # chrome_options.add_argument('--disable-gpu')
                driver = webdriver.Chrome(chrome_options=chrome_options)
            elif browserName.lower() == 'edge':
                driver = webdriver.Edge()
            elif browserName.lower() == 'firefox':
                driver = webdriver.Firefox()
            # driver对象创建成功，创建等待类实例对象
            waitUtil = WaitUtil(driver)
            self._driver = driver
        except Exception as e:
            handleErr(e)
            raise e


    @CNBMException
    def visit_url(self, url):
        ''' 访问网址 '''
        try:
            self._driver.get(url)
        except Exception as e:
            handleErr(e)
            raise e

    @CNBMException
    def close_browser(self):
        ''' 关闭浏览器 '''
        try:
            if self._driver:
                self._driver.quit()
                self._driver = None
        except Exception as e:
            handleErr(e)
            raise e

    @CNBMException
    def maximize_browser(self):
        ''' 窗口最大化 '''
        try:
            self._driver.maximize_window(self)
        except Exception as e:
            handleErr(e)
            raise e

    @CNBMException
    def close_page(self):
        ''' 关闭标签页 '''
        try:
            self._driver.close()
        except Exception as e:
            handleErr(e)
            raise e

    @CNBMException
    def refresh_page(self):
        ''' 刷新网页 '''
        try:
            self._driver.refresh()
        except Exception as e:
            handleErr(e)
            raise e

    @CNBMException
    def findElement(self, locationType, locatorExpression, timeout=None):
        ''' 通过属性值查找控件
        :param locationType: 下拉框控件定位类型
        :param locatorExpression: 下拉框控件定位属性值
        '''
        try:
            OM = ObjectMap(self._driver)
            time = timeout if timeout else self.timeout

            element = OM.findElebyMethod(locationType, locatorExpression, timeout=time)
            return WebElement(self._driver, element)
        except Exception as e:
            handleErr(e)
            raise e

    @CNBMException
    def localElement(self, name, timeout=None):
        ''' 从本地库查找控件
        :param name: 本地工程对象库中控件名称
        '''
        try:
            OM = ObjectMap(self._driver)
            time = timeout if timeout else self.timeout

            locatorExpression = PF.getObjFromLog("Chrome", name)
            locatorExpression = locatorExpression["xpath"]
            element = OM.findElebyMethod("xpath", locatorExpression, timeout=time)
            return WebElement(self._driver, element)
        except Exception as e:
            handleErr(e)
            raise e

    @CNBMException
    def highlight(self, element, timeToDisappear=None):
        style = element.get_attribute("style")
        self._driver.execute_script("arguments[0].setAttribute('style',arguments[1]);",
                              element,"border:2px solid red;")
                              # element,"background:green;border:2px solid red;")
        if timeToDisappear:
            while timeToDisappear:
                time.sleep(1)
                timeToDisappear -= 1
            self.setAttribute(element, "style", style)

    @CNBMException
    def captureScreen(self):
        try:
            projectName = cf.get_value("projectName")
            picPath = "%s\\pictures\\%s" \
                      %(frozen_dir.app_path(), projectName)
            dirName = createCurrentDateDir(picPath)
            capture_screen(dirName)
        except Exception as e:
            handleErr(e)
            raise e

    @CNBMException
    def addAttribute(self, element, attributeName, value):
        self._driver.execute_script("arguments[0].%s = arguments[1]" %attributeName, element, value)

    @CNBMException
    def setAttribute(self, element, attributeName, value):
        self._driver.execute_script("arguments[0].setAttribute(arguments[1],arguments[2])",
                              element, attributeName, value)

    @CNBMException
    def removeAttribute(self, element, attributeName):
        self._driver.execute_script("arguments[0].removeAttribute(arguments[1])",
                              element, attributeName)


class WebElement:
    def __init__(self, driver, element):
        self.driver = driver
        self.element = element

    @CNBMException
    def click(self):
        ''' 点击控件 '''
        try:
            self.element.click()
        except Exception as e:
            handleErr(e)
            raise e

    @CNBMException
    def sendkeys(self, inputContent):
        ''' 输入框输值 '''
        try:
            self.element.clear()
            self.element.send_keys(inputContent)
        except Exception as e:
            handleErr(e)
            raise e

    @CNBMException
    def select(self, inputContent):
        ''' 下拉框选择 '''
        try:
            element = Select(self.element)
            element.select_by_visible_text(inputContent)
        except Exception as e:
            handleErr(e)
            raise e

    @CNBMException
    def clear(self):
        ''' 清除下拉框默认值 '''
        try:
            self.element.clear()
        except Exception as e:
            handleErr(e)
            raise e

    @CNBMException
    def get_attribute(self, attrName):
        ''' 获取控件属性值
        :param attrName: 属性名
        '''
        try:
            attribute = self.element.get_attribute(attrName)
            return attribute
        except Exception as e:
            handleErr(e)
            raise e

    @CNBMException
    def scroll(self, position="up"):
        ''' 滚动页面至元素可见
        :param position: 操作类型（top：滑动后元素置顶；bottom：滑动后元素置底）
        '''
        try:
            if position == "top":
                self.driver.execute_script("arguments[0].scrollIntoView();", self.element)
            elif position == "bottom":
                self.driver.execute_script("arguments[0].scrollIntoView(false);", self.element)
        except Exception as e:
            handleErr(e)
            raise e


if __name__ == "__main__":
    # driver = webdriver.Chrome()
    # print('打开浏览器')
    # driver.maximize_window()
    #
    # driver.get('http://cdwp.cnbmxinyun.com')
    # print('打开网页')

    PA = PageAction()
    PA.open_browser("chrome")
    PA.visit_url("http://cdwp.cnbmxinyun.com")


    # el = OM.findElebyMethod("xpath", '//input[@ng-model="user_name"]')
    # print(el.get_attribute("placeholder"))
    # el.send_keys("abc")
    #
    # el1 = OM.findElebyMethod("xpath", '//input[@ng-model="password"]')
    # print(el1.get_attribute("placeholder"))
    # el1.send_keys("123456")
    #
    # driver.quit()

    PA.findElement("xpath", '//input[@ng-model="password"]').sendkeys("123456")