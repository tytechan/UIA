#!python3
# -*- coding: utf-8 -*-
from localSDK import *
from tkinter import ttk
from config.ErrConfig import CNBMException, handleErr
from config.DirAndTime import getCurrentDate, getCurrentTime
from selenium.webdriver.common.keys import Keys
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
import pickle
import pythoncom
import time
import PyHook3
import tkinter as tk
import copy
import uiautomation as auto
import config.Globals as cf
import win32api, win32con
import threading
import re

class AppControl:
    ''' Windows控件基本处理方法 '''
    def __init__(self):
        self.appName = None                             # 应用名称（用于写日志）
        self.waitTime = 2                               # 最长等待时间
        self.intervalTime = 1                           # 查找控件间隔时间
        self.dict = dict()                              # 工程下“log.txt”中的所有控件信息
        self.maxNum = 9999                              # 相同搜索条件下需遍历的控件量最大值
        self.projectName = None                         # 工程名

    def openApp(self, path):
        ''' 打开应用程序
        :param path: 程序安装路径
        '''
        import win32api
        win32api.ShellExecute(0, 'open', path, '', '', 1)

    def killApp(self, appName):
        ''' 关闭应用程序
        :param appName:程序名
        '''

        # TODO:下述方法为关闭该应用所有进程
        # import os
        # cmd = r"taskkill /F /IM " + appName
        # os.system(cmd)

    def ShowDesktop(self):
        ''' 显示桌面 '''
        auto.ShowDesktop()

    def searchFromChildren(self, parentObj, infoWithIndex, index):
        ControlType = infoWithIndex["ControlType"]
        AutomationId = infoWithIndex["AutomationId"]
        ClassName = infoWithIndex["ClassName"]
        Name = infoWithIndex["Name"]
        # SupportedPattern = infoWithIndex["SupportedPattern"]

        keyList = list()
        childrenList = parentObj.GetChildren()
        for obj in childrenList:
            try:
                assert ClassName == obj["ClassName"]
            except AssertionError as e:
                raise e

    def checkBottom(self, keyObj, flag=False):
        ''' （Windows）用户选择该控件需保存时，进行提示（凭获取信息无法唯一识别）
        :param keyObj: 控件树信息
        :param flag: True（进行唯一性校验，只用在录制时）/False（反之）
        :return: True：通过已获取的层级结构及属性信息可唯一定位控件；False：反之
        '''
        import uiautomation as auto

        bottomObj = keyObj[str(len(keyObj) - 1)]
        x = bottomObj["x"]
        y = bottomObj["y"]

        searchStr = None
        searchObj = None
        """ 
        最上层（桌面）及最下层（目标）除外，每层遍历一次 
        先通过属性直接查找下一层，若坐标不匹配，再通过GetChildren结合属性+坐标查找下一层
        """
        for i in range(len(keyObj) - 1):
            obj = keyObj[str(i + 1)]
            ControlType = obj.get("ControlType", "None")
            AutomationId = obj.get("AutomationId", "None")
            ClassName = obj.get("ClassName", "None")
            Name = obj.get("Name", "None")
            SupportedPattern = obj.get("SupportedPattern", "None")

            # parentObj = auto if i == 0 else searchObj
            searchStr = "auto" if i == 0 else "searchObj"
            searchStr = searchStr + "." + str(ControlType) + \
                        "(searchDepth=1, ClassName='" + ClassName + "'"

            RegexName = ""
            for myStr in list(Name):
                RegexName += (".*" + myStr)
            RegexName += ".*"
            RegexNameStr = ", RegexName='" + RegexName + "'"
            searchStr += RegexNameStr if Name != "None" else ""

            AutomationIdStr = ", AutomationId='" + AutomationId + "'"
            searchStr += AutomationIdStr if (AutomationId != "" and not AutomationId.isdigit()) else ""

            SupportedPatternStr = ", SupportedPattern='" + SupportedPattern + "'"
            searchStr += SupportedPatternStr if SupportedPattern != "None" else ""

            DescStr = ", Desc='" + Name + "')"
            searchStr += DescStr if (i == 0 and Name != "None") else ")"

            print(searchStr)
            try:
                searchObj = eval(searchStr)
                searchObj.Refind(maxSearchSeconds=self.waitTime, searchIntervalSeconds=self.intervalTime)

                # 置顶应用程序
                if i == 0:
                    # searchObj.SetTopmost()
                    searchObj.SetActive()
                    # searchObj.SetFocus()

                if flag:
                    # 检查坐标是否匹配
                    rect = searchObj.BoundingRectangle
                    # print(rect)

                    assert x <= rect.right and x >= rect.left \
                           and y <= rect.bottom and y >= rect.top, \
                        "第 %s 层控件坐标不匹配！" %(i + 1)
            except AssertionError as e:
                print(e)
                # raise e
                return False
            except Exception as e:
                print("未找到第 %s 层控件！" %(i+ 2))
                # raise e
                return False

        # print("唯一性校验完毕！")
        return searchObj

    def checkObjFromLog(self, conductType, name):
        ''' 对本地控件库中控件进行唯一性校验
        :param conductType: 控件类型
        :param name: 控件名
        :param projectName: 所在工程名
        '''
        import win32api
        import win32con

        # TODO：根据判断识别出结果但应用程序实际不可识别的处理方法
        try:
            filePath = r"%s\%s\log.txt" %(parentDirPath, self.projectName)
            with open(filePath, "r+") as f:
                rawData = f.read()
                if not rawData:
                    rawData = "{}"

            rawDict = eval(rawData)
            objDict = rawDict.get(conductType)
            assert objDict is not None, \
                "本地控件库无 [%s] 类型控件，请检查！" %conductType
            objDict = rawDict[conductType].get(name)
            assert objDict is not None, \
                "[%s] 库中未找到控件 [%s] ，请检查控件名称！" %(conductType, name)

            if conductType == "Windows":
                keyObj = objDict.get("Depth")
                flag = self.checkBottom(keyObj)
            else:
                CH = ChromeHooker()
                eleInfo = objDict["xpath"]
                # flag = self.checkBrowserElement(eleInfo)
                try:
                    flag = CH.checkElement(eleInfo)
                except:
                    flag = False

            if not flag:
                win32api.MessageBox(0, "依据获取信息无法在识别 [%s] 库中控件 [%s] ，请检查控件状态或重新识别！"
                                    %(conductType, name), "提示", win32con.MB_OK)
            else:
                print("本地库 [%s] 下 [%s] 控件校验通过！" %(conductType, name))
        except AssertionError as e:
            raise e
        except Exception as e:
            raise e

    def deleteObjFromLog(self, conductType, name):
        ''' 删除本地控件
        :param conductType: 控件类型
        :param name: 控件名
        '''
        try:
            filePath = r"%s\%s\log.txt" %(parentDirPath, self.projectName)
            with open(filePath, "r+") as f:
                rawData = f.read()
                if not rawData:
                    rawData = "{}"

            rawDict = eval(rawData)
            objDict = rawDict.get(conductType)
            assert objDict is not None, \
                "本地控件库无 [%s] 类型控件，请检查！" %conductType
            objDict = rawDict[conductType].get(name)
            assert objDict is not None, \
                "[%s] 库中未找到控件 [%s] ，请检查控件名称！" %(conductType, name)

            del rawDict[conductType][name]

            with open(filePath, "w") as f:
                f.write(str(rawDict))
        except AssertionError as e:
            raise e
        except Exception as e:
            raise e

    def insertIntoLog(self, autoType, objName, xpath, projectPath):
        ''' 向本地对象库插入控件信息（只限用在浏览器控件！！！）
        :param autoType: 控件类型
        :param objName: 控件名
        :param xpath: 手写xpath
        :param projectPath: 插入工程路径
        '''
        try:
            filePath = projectPath + r"\log.txt"
            # print(filePath)
            with open(filePath, "a+"):
                pass
            with open(filePath, "r+") as f:
                rawData = f.read()
                if not rawData:
                    rawData = "{'%s': {}}" %autoType
                rawDict = eval(rawData)

                if autoType not in rawDict:
                    rawDict[autoType] = dict()
                rawDict[autoType][objName] = dict()
                rawDict[autoType][objName]["xpath"] = xpath
                rawDict[autoType][objName]["time"] = "%s %s" %(getCurrentDate(), getCurrentTime())
                rawDict[autoType][objName]["info"] = "手工插入"
            with open(filePath, "w") as f:
                f.write(str(rawDict))
                # f.close()
        except Exception as e:
            raise e


    def checkBottom_EX(self, keyObj):
        ''' 用户选择该控件需保存时，进行提示（凭获取信息无法唯一识别）
        :param keyObj: 控件树信息
        :return:
        '''
        import uiautomation as auto

        bottomObj = keyObj[str(len(keyObj) - 1)]
        appInfo = keyObj.get("1")  # info为桌面信息时，只有一层“0”

        """ Step 1：定位APP（第一层） """
        self.appName = appInfo["Name"]
        appWindow = auto.WindowControl(searchDepth=1, ClassName=appInfo["ClassName"],
                                       RegexName=".*" + appInfo["Name"] + ".*", Desc=self.appName)
        # 置顶应用程序
        appWindow.SetTopmost()
        # appWindow.SetActive()
        # appWindow.SetFocus()
        try:
            assert appWindow.Exists(maxSearchSeconds=self.waitTime, searchIntervalSeconds=self.intervalTime), \
                "请检查待操作应用程序状态！"
            appWindow.Refind()
        except AssertionError as e:
            raise e

        """ Step 2：定位目标层（最下层） """
        if len(keyObj) > 2:
            # 两层时，最下层即为app本身；大于两层时，在app下查找
            ControlType = bottomObj["ControlType"]
            Name = bottomObj["Name"]
            ClassName = bottomObj["ClassName"]
            AutomationId = bottomObj["AutomationId"]
            SupportedPattern = bottomObj["SupportedPattern"]

            searchStr = "appWindow" + "." + str(ControlType) + \
                        "(searchDepth=" + str(len(keyObj) - 1) \
                        + ", ClassName='" + ClassName \
                        + "', SupportedPattern='" + SupportedPattern + "'"

            RegexNameStr = ", RegexName='.*" + Name + ".*'"
            searchStr += RegexNameStr if Name else ""

            AutomationIdStr = "', AutomationId='" + AutomationId + "'"
            searchStr += AutomationIdStr if (AutomationId and not AutomationId.isdigit()) else ""

            # 假设最底层存在多个相同属性控件，遍历并获取集合
            container = []
            for i in range(self.maxNum):
                copyStr = copy.deepcopy(searchStr)
                foundIndexStr = ", foundIndex=" + str(i + 1) + ")"
                copyStr += foundIndexStr
                print(copyStr)

                try:
                    searchObj = eval(copyStr)
                    searchObj.Refind()
                    container.append(searchObj)
                except:
                    try:
                        assert i > 0, "根据获取属性未查找到最底层（目标）控件！"
                        # 查找到最底层有多个控件，保存并进行下一步判断
                        break
                    except AssertionError as e:
                        raise e
            print(container)


    @CNBMException
    def objControl(self, name):
    # def objControl(self, name, conductType, inputStr=None):
        sucFlag = False
        errInfo = ""
        try:
            assert self.dict["Windows"].get(name) is not None, \
                "本地库中 [Windows] 类型下未找到名为 [%s] 的控件，请检查“log.txt”文件！" %name
            info = self.dict["Windows"].get(name).get("Depth")
            obj = self.checkBottom(info)
            assert obj, "根据本地控件信息未定位到目标控件 [%s]！" %name

            # if conductType == "点击":
            #     obj.Click()
            #
            # elif conductType == "输入":
            #     obj.SendKeys(inputStr)

            sucFlag = True
            return WinElement(obj)
        except AssertionError as e:
            errInfo = e
            raise e
        except Exception as e:
            errInfo = e
            raise e
        finally:
            if not sucFlag:
                handleErr(errInfo)
                # cf.set_value("err", errInfo)
            # auto.Logger.Write(name)
            # logInfo = "[节点日志] 控件 [%s]，操作类型 [%s]，执行结束" %(name, conductType)
            # auto.Logger.WriteLine(logInfo, auto.ConsoleColor.Cyan, writeToFile = True)

    def objControl_EX(self, name, conductType, string=None, **kwargs):
        import win32api
        # 每层遍历，直接通过属性查询
        '''
        info = self.dict.get(name)["Depth"]
        if len(info) == 2:
            searchObj = auto
        elif len(info) > 2:
            searchStr = None
            """ 
            最上层（桌面）及最下层（目标）除外，每层遍历一次 
            """
            for i in range(len(info) - 2):
                obj = info[str(i + 1)]
                ControlType = obj["ControlType"]
                AutomationId = obj["AutomationId"]
                ClassName = obj["ClassName"]
                Name = obj["Name"]
                SupportedPattern = obj["SupportedPattern"]

                searchStr = "auto" if i == 0 else "searchObj"
                searchStr = searchStr + "." + str(ControlType) + \
                            "(searchDepth=1, ClassName='" + ClassName \
                            + "', SupportedPattern='" + SupportedPattern + "'"

                RegexNameStr = ", RegexName='.*" + Name + ".*'"
                searchStr += RegexNameStr if Name else ""

                AutomationIdStr = "', AutomationId='" + AutomationId + "'"
                searchStr += AutomationIdStr if (AutomationId and not AutomationId.isdigit())  else ""

                DescStr = ", Desc='" + self.appName + "')"
                searchStr += DescStr if i == 0 else ")"

                print(searchStr)
                try:
                    searchObj = eval(searchStr)
                    searchObj.Refind()
                except Exception as e:
                    print("未找到第 %s 层控件！" %(i+ 2))
                    raise e

            """
            在倒数第二层所有子空间中查找符合条件情况
            """
        '''

        # 每层遍历，在上一层的所有子控件中查找
        info = self.dict.get(name)["Depth"]


class WinElement:
    def __init__(self, obj):
        self.obj = obj

    @CNBMException
    def click(self):
        ''' 点击控件 '''
        try:
            self.obj.Click()
        except Exception as e:
            handleErr(e)
            raise e

    @CNBMException
    def sendkeys(self, inputContent):
        ''' 输入框输值 '''
        try:
            self.obj.clear()
            self.obj.SendKeys(inputContent)
        except Exception as e:
            handleErr(e)
            raise e


class PublicFunc:
    ''' （录制、回放）功能函数 '''
    def readFromLog(self):
        ''' 读取log文件信息 '''
        objDict = None
        try:
            filePath = "log.txt"
            with open(filePath, "r+") as f:
                rawData = f.read()
            assert rawData and type(rawData) == str, "本地数据为空，请先维护本地控件库！"
            objDict = eval(rawData)
            # print(rawData)
        except AssertionError as e:
            win32api.MessageBox(0, "请先维护本地控件库！", "提示", win32con.MB_OK)
            raise e
        except Exception as e:
            raise e
        finally:
            return objDict

    def getObjFromLog(self, elementType, name):
        try:
            rawDict = self.readFromLog()
            assert elementType in rawDict.keys(), \
                "本地库中无 [%s] 类型控件，请检查！" %elementType
            assert name in rawDict[elementType].keys(), \
                      "本地库中 [%s] 类型下未找到名为 [%s] 的控件，请检查！" %(elementType, name)
            objDict = rawDict[elementType][name]
            return objDict
        except AssertionError as e:
            raise e
        except KeyError as e:
            raise e
        except Exception as e:
            raise e

    def openChrome(self, path):
        ''' 调起chrome进程，可用于后续流程识别
        :param path: 存放chrome配置文件的路径（自定义）
        '''
        import subprocess
        import ctypes, sys
        cmd = 'chrome.exe --remote-debugging-port=9222 --user-data-dir="%s"' %path

        # 判断是否管理员权限
        if not ctypes.windll.shell32.IsUserAnAdmin():
            ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)
        popen = subprocess.Popen(cmd, shell=True, stdout=subprocess.PIPE)
        # popen.wait()

    def openFirefox(self, profileDir):
        ''' 先启动 geckodriver.exe（确保已配置环境变量），再调起firefox服务 '''
        from selenium import webdriver
        from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
        import pickle
        import subprocess
        import psutil

        # 确保杀掉geckodriver进程
        killGecko = "taskkill /F /IM geckodriver.exe"
        popen = subprocess.Popen(killGecko, shell=True, stdout=subprocess.PIPE)
        popen.wait()

        # 打开geckodriver进程
        cmd = 'geckodriver.exe'
        subprocess.Popen(cmd, shell=True, stdout=subprocess.PIPE)
        # os.system(cmd)

        # 调起Firefox前须确保geckodriver进程已启动
        processList = list()
        while "geckodriver.exe" not in processList:
            pids = psutil.pids()
            for pid in pids:
                # print(pid)
                p = psutil.Process(pid)
                # print(p.name)
                processList.append(p.name())

        # import time
        # while "geckodriver.exe" not in [psutil.Process(i).name() for i in psutil.pids()]:
        #     time.sleep(0.2)

        # profile = webdriver.FirefoxProfile(profileDir)
        #         # profile.update_preferences()

        profile = webdriver.FirefoxProfile()
        profile.add_extension()
        profile.add_extension(r"C:\Python35\Lib\site-packages\selenium\webdriver\firefox\xpath finder.xpi")
        # from selenium.webdriver.firefox.firefox_profile import FirefoxProfile
        # profile = FirefoxProfile()r
        # profile.accept_untrusted_certs = True

        # driver = webdriver.Remote(
        #     command_executor="http://127.0.0.1:4444",
        #                                               desired_capabilities=DesiredCapabilities.FIREFOX,
        #                                               browser_profile=profile)
        # driver.implicitly_wait(10)



        # driver = webdriver.remote.webdriver.WebDriver(command_executor="http://127.0.0.1:4444",
        #                                               desired_carpabilities=DesiredCapabilities.FIREFOX)

        driver = webdriver.remote.webdriver.WebDriver(command_executor="http://127.0.0.1:4444",
                                                      desired_capabilities={"browserName": "firefox",
                                                                            "marionette": True,
                                                                            "acceptInsecureCerts": True,
                                                                            "javascriptEnabled": True,
                                                                            # "platform": "ANY",
                                                                            # "moz:firefoxOptions":{},
                                                                            },
                                                      browser_profile=profile
                                                      )

        # driver = webdriver.remote.webdriver.WebDriver(command_executor="http://127.0.0.1:4444",
        #                                               desired_capabilities=DesiredCapabilities.FIREFOX,
        #                                               browser_profile=profile)

        # driver = webdriver.remote.webdriver.WebDriver(command_executor="http://127.0.0.1:4444",
        #                                               desired_capabilities={"browserName": "firefox",
        #                                                                     "marionette": True,
        #                                                                     "acceptInsecureCerts": True,
        #                                                                     "moz:firefoxOptions":{}},
        #                                               browser_profile=profile)
        driver.get('http://www.baidu.com/')
        print(driver.capabilities)
        print(driver.command_executor.keep_alive)
        print(driver.command_executor._url)
        print(driver.session_id)

        params = {}
        params["session_id"] = driver.session_id
        params["server_url"] = driver.command_executor._url

        dataPath = parentDirPath + r"\hooker\params.data"
        f = open(dataPath, 'wb')
        # 转储对象至文件
        pickle.dump(params, f)
        f.close()

    def clickMe(self):
        # button被点击之后会被执行
        global autoType
        win.destroy()
        autoType = browser.get()
        cf.set_value("autoType", autoType)
        # print(result)

    def initProcess(self):
        global win, browser
        # 初始化全局变量
        cf._init()
        path = os.getcwd()
        cf.set_value("path", path)

        win = tk.Tk()
        # win.geometry("300x150+500+200")  # 大小和位置
        win.title("请确认")
        ttk.Label(win, text="请选择浏览器类型：").grid(column=1, row=0)  # 添加一个标签，并将其列设置为1，行设置为0

        # 按钮
        action = ttk.Button(win, text="选择", command=self.clickMe)
        action.grid(column=2, row=1)

        # 创建一个下拉列表
        browser = tk.StringVar()
        browserChosen = ttk.Combobox(win, width=12, textvariable=browser)
        browserChosen['values'] = ("", "Chrome", "IE", "Firefox", "Windows")
        browserChosen.grid(column=1, row=1)
        # 设置下拉列表默认显示的值，0为 numberChosen['values'] 的下标值
        browserChosen.current(0)
        # 当调用mainloop()时,窗口才会显示出来
        win.mainloop()

    def startHook(self):
        ''' 开始录制流程 '''
        import hooker.Hook as H
        self.initProcess()
        autoType = cf.get_value("autoType")
        print("本次录制类型:", autoType)
        cf.set_value("autoType", autoType)

        if autoType == "Windows":
            HK = H.Hooker()
            HK.hooks()
        elif autoType == "IE":
            IEXPath = parentDirPath + r"\tools\IEXPath.exe"
            win32api.ShellExecute(0, 'open', IEXPath, '', '', 1)

            HK = H.Hooker()
            HK.hooks()
        elif autoType == "Chrome" or autoType == "Firefox":
            Init().hooks()
            from hooker.Extension import run
            run()
        else:
            pass

        # os._exit(0)


hm = PyHook3.HookManager()
class Init:
    def __init__(self):
        self.pause = False
        self.flag = False
        self.autoType = cf.get_value("autoType")

    #键盘事件处理函数
    def OnKeyboardEvent(self, event):
        keyType = event.Key
        print('Key:', keyType)
        if keyType == "Return":
            # print(self.pause, type(self.pause))
            self.pause = bool(1 - self.pause)

            if self.autoType == "Chrome" and self.pause:
                # 类型选择Chrome，且开始录制后两次点击快捷开关（先暂停，后开启录制）
                # 默认此时焦点在driver对应Chrome浏览器上
                CH = ChromeHooker()

                html = WebDriverWait(CH.driver, 0.5).until(
                    EC.visibility_of_element_located((By.XPATH, "/html")),  "请检查Chrome状态！")
                time.sleep(0.5)

                for i in range(3):
                    try:
                        WebDriverWait(CH.driver, 0.5).until(
                            EC.visibility_of_element_located((By.ID, "xpath_inspector_toolbar")),
                            "请检查xpath_inspector状态！")
                        break
                    except:
                        if i != 2:
                            time.sleep(0.3)
                        else:
                            html.send_keys(Keys.CONTROL, Keys.ENTER)
                self.flag = True
            elif self.autoType == "Firefox" and self.pause:
                import hooker.LocalFirefox as LF

                dataFile = parentDirPath + "\hooker\params.data"
                f = open(dataFile, 'rb')
                params = pickle.load(f)
                print(params)

                driver = LF.myWebDriver(service_url=params["server_url"],
                    session_id=params["session_id"])

                for i in range(3):
                    try:
                        driver.find_element_by_id("run_xpath")
                        break
                    except:
                        if i != 2:
                            time.sleep(0.3)
                        else:
                            KK = KeyboardKeys()
                            KK.keyDown("ctrl")
                            KK.keyDown("enter")
                            KK.keyUp("enter")
                            KK.keyUp("ctrl")
                self.flag = True
        return True

    def main(self):
        hm.KeyDown = self.OnKeyboardEvent
        hm.HookKeyboard()
        pythoncom.PumpMessages(10000)


    def hooks(self):
        while True:
            try:
                t = threading.Thread(target=self.main, args=())
                t.setDaemon(True)
                t.start()
            except:
                print("Error")
            if self.flag:
                print("[开始录制]")
                hm.UnhookMouse()
                hm.UnhookKeyboard()
                return
            time.sleep(0.5)


class ChromeHooker:
    def __init__(self):
        self.chrome_driver = path
        self.chrome_options = Options()
        self.chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")

        self.driver = webdriver.Chrome(self.chrome_driver, chrome_options=self.chrome_options)

        try:
            cf._global_dict
        except:
            # checkFunc中调用时，此前未定义 cf._global_dict
            cf._init()
        cf.set_value("driver", self.driver)
        print("[title] ", self.driver.title)

        self.page_source = self.driver.page_source
        self.url = self.driver.current_url

    def refreshDriver(self):
        # while True:
        if self.url != self.driver.current_url:
            # 网页有变动，更新driver
            self.url = self.driver.current_url
            self.driver = webdriver.Chrome(self.chrome_driver, chrome_options=self.chrome_options)

            self.page_source = self.driver.page_source
            # print("刷新后 title 为：", self.driver.title, "\n")
            # print(self.page_source)

    def checkElement(self, eleInfo):
        ''' 唯一性校验
        :param eleInfo: 目标元素xpath
        :return: 目标元素
        '''
        from localSDK.BrowserFunc import ObjectMap
        OM = ObjectMap(self.driver)
        try:
            element = OM.findElebyMethod("xpath", eleInfo, timeout=0.1)
            OM.highlight(element, 1)
            return element
        except Exception as e:
            raise e


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
    def twoKeys(key1, key2):
        # 模拟两个组合按键
        KeyboardKeys.keyDown(key1)
        KeyboardKeys.keyDown(key2)
        KeyboardKeys.keyUp(key2)
        KeyboardKeys.keyUp(key1)

if __name__ == "__main__":
    AC = AppControl()
    PF = PublicFunc()

    profileDir = r"C:\Users\47612\AppData\Roaming\Mozilla\Firefox\Profiles\14q6qdug.default-release-1"
    PF.openFirefox(profileDir)