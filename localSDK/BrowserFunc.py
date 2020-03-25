# encoding = utf - 8
from localSDK.BasicFunc import PublicFunc
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
# from config.ErrConfig import CNBMException, handleErr, capture_screen
from config.ErrConfig import *
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
        ''' 在 timeout 内，找到元素（显式等待页面元素出现在DOM中，并且可见，存在则返回该页面元素对象） '''
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

    def findElePresented(self, locateType, locatorExpression, errInfo="", timeout=10):
        ''' 在 timeout 内，找到元素（显式等待页面元素出现在DOM中，但不一定可见，存在则返回元素对象） '''
        errInfo = "%s 秒内未找到 %s 为 %s 的元素！" %(timeout, locateType, locatorExpression) + errInfo
        try:
            locateType = locateType.lower()
            objStr = 'WebDriverWait(self.driver, timeout).until(' \
                     'EC.presence_of_element_located((By.%s, ' \
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


class PageAction:
    _driver = None
    def __init__(self):
        self.type = None            # 控件类型，从对象库取值前需定义
        self.timeout = 10           # 默认等待控件的最长时间
        self.loadTime = 60          # 默认等待页面加载的最长时间
        self.pageSource = None
        self.title = None

    @CNBMException
    def open_browser(self, dataPath):
        ''' 打开指定类型浏览器
        :param dataPath: 固定配置文件路径（可与本地已有配置相同），暂只用作chrome
        :return: 全局变量driver
        '''
        try:
            ''' 
            配置无需等driver下页面加载完毕再执行下一步
            （此操作会去掉所有等待操作，弃用）            
            '''
            # desired_capabilities = DesiredCapabilities.CHROME
            # desired_capabilities["pageLoadStrategy"] = "none"

            if self.type.lower() == 'ie':
                driver = webdriver.Ie()
            elif self.type.lower() == 'chrome':
                chrome_options = webdriver.ChromeOptions()
                # 最大化打开
                chrome_options.add_argument('--start-maximized')

                # 设置配置数据路径
                chrome_options.add_argument(
                    '--user-data-dir=C:%s' %dataPath)

                # 让Chrome在root权限下执行
                chrome_options.add_argument('-–no-sandbox')

                # 用于控制进程是否在后台执行
                # chrome_options.add_argument('--headless')
                # chrome_options.add_argument('--disable-gpu')

                driver = webdriver.Chrome(chrome_options=chrome_options)
            elif self.type.lower() == 'edge':
                driver = webdriver.Edge()
            elif self.type.lower() == 'firefox':
                driver = webdriver.Firefox()
            # driver对象创建成功，创建等待类实例对象
            self._driver = driver
            self.load_page()
        except Exception as e:
            handleErr(e)
            raise e

    @CNBMException
    def load_page(self, loopTime=None):
        '''
        修改全局等待页面加载的最长时间
        :param loopTime: 默认为self.loadTime
        '''
        try:
            loopTime = self.loadTime if not loopTime else loopTime
            self._driver.set_page_load_timeout(loopTime)
            self._driver.set_script_timeout(loopTime)
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
            self._driver.maximize_window()
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
    def update_page_source(self):
        ''' 重新获取页面资源 '''
        try:
            self.pageSource = self._driver.page_source
            self.title = self._driver.title
        except Exception as e:
            handleErr(e)
            raise e

    @CNBMException
    def findElement(self, locationType, locatorExpression, elements=None, index=0, timeout=None):
        ''' 通过属性值查找控件（找到即代表元素可见）
        :param locationType: 下拉框控件定位类型
        :param locatorExpression: 下拉框控件定位属性值
        '''
        try:
            OM = ObjectMap(self._driver)
            time = timeout if timeout else self.timeout

            if not elements:
                element = OM.findElebyMethod(locationType, locatorExpression, timeout=time)
            else:
                element = elements[index]
            return WebElement(self._driver, element)
        except Exception as e:
            handleErr(e)
            raise e

    @CNBMException
    def findElements(self, locationType, locatorExpression, timeout=None):
        ''' 通过属性值查找多个控件（找到即代表元素存在，不一定可见）
        :param locationType: 下拉框控件定位类型
        :param locatorExpression: 下拉框控件定位属性值
        '''
        try:
            OM = ObjectMap(self._driver)
            time = timeout if timeout else self.timeout

            elements = OM.findElesbyMethod(locationType, locatorExpression, timeout=time)
            return elements
        except Exception as e:
            handleErr(e)
            raise e

    @CNBMException
    def wait_for_presence(self, locationType, locatorExpression, timeout=None):
        ''' 通过属性值查找控件
        （找到即代表元素存在于DOM中，并不一定可见！无特殊要求，直接用 findElemrnt 即可）
        :param locationType: 下拉框控件定位类型
        :param locatorExpression: 下拉框控件定位属性值
        '''
        try:
            OM = ObjectMap(self._driver)
            time = timeout if timeout else self.timeout

            element = OM.findElePresented(locationType, locatorExpression, timeout=time)
            return WebElement(self._driver, element)
        except Exception as e:
            handleErr(e)
            raise e

    @CNBMException
    def localElement(self, name, index=0, timeout=None):
        ''' 从本地库查找控件
        :param name: 本地工程对象库中控件名称
        :param index: 当对象库中xpath对应多个元素时，通过index定位
        '''
        try:
            OM = ObjectMap(self._driver)
            time = timeout if timeout else self.timeout

            locatorExpression = PF.getObjFromLog(self.type, name)
            xpath = locatorExpression["xpath"]
            locatorExpression = xpath if index == 0 else (xpath %index)
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

    @CNBMException
    def switch_to_frame(self, locationType, frameLocatorExpression, timeout=None):
        ''' 切换到指定frame
        :param locationType: frame定位类型
        :param frameLocatorExpression: frame定位值
        '''
        try:
            frame = self.findElement(locationType, frameLocatorExpression, timeout=timeout)
            self._driver.switch_to.frame(frame)
        except Exception as e:
            handleErr(e)
            raise e

    @CNBMException
    def switch_to_default_content(self):
        ''' 切换到主文档（切到frame中之后，便不能继续操作主文档的元素，若想操作主文档内容，则需切回主文档） '''
        try:
            self._driver.switch_to.default_content()
        except Exception as e:
            handleErr(e)
            raise e

    @CNBMException
    def switch_to_window(self, name=None, num=0, loopTime=None):
        ''' 切换到指定的页签（若传入页签名，则依据名称查找；否则依据索引查找，默认切换到第一个页签）
        :param name: 页签名称
        :param num: 页签索引号
        :param loopTime: 此次切换需等待页面加载的最长时间（切换完成后自动恢复为默认时间）
        '''
        try:
            if loopTime:
                self.load_page(loopTime)
            if name:
                self._driver.switch_to.window(name)
            else:
                all_handles = self._driver.window_handles
                self._driver.switch_to.window(all_handles[num])
            # 重置等待页面加载的最长时间
            self.load_page(self.loadTime)
        except Exception as e:
            handleErr(e)
            raise e

    @CNBMException
    def set_date(self, locationType, locatorExpression, inputTime):
        ''' 通过js修改控件“readonly”属性并设置日期
        :param locationType: 控件定位类型
        :param locatorExpression: 控件定位值
        :param inputTime: 日期值（格式须与系统对应）
        '''
        try:
            element = self.findElement(locationType, locatorExpression)
            self.removeAttribute(element, "readonly")
            element.sendkeys(inputTime)
        except Exception as e:
            handleErr(e)
            raise e


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
    def scroll(self, position="top"):
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

    @CNBMException
    def move(self):
        ''' 移动鼠标光标至元素上 '''
        try:
            ActionChains(self.driver).move_to_element(self.element).perform()
        except Exception as e:
            handleErr(e)
            raise e

    @CNBMException
    def js_click(self):
        ''' js点击 '''
        try:
            self.driver.execute_script("arguments[0].click();", self.element)
        except Exception as e:
            handleErr(e)
            raise e

    @CNBMException
    def drag_and_drop(self, x_index, y_index):
        ''' 模拟鼠标拖拽 '''
        try:
            ActionChains(self.driver).drag_and_drop_by_offset(self.element, x_index, y_index).perform()
        except Exception as e:
            handleErr(e)
            raise e

    @CNBMException
    def simulate_keys(self, keyType):
        ''' 模拟键盘向控件发送指定键
        :param keyType: 键类型，须与 “selenium.webdriver.common.keys” 中的键对应
        '''
        try:
            keyStr = "self.element.send_keys(Keys.%s)" %keyType
            eval(keyStr)
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