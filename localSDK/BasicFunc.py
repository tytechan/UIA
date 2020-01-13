#!python3
# -*- coding: utf-8 -*-
from localSDK import *
from tkinter import ttk
from config.ErrConfig import CNBMException, handleErr
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

    def checkBrowserElement(self, eleInfo):
        ''' （Chrome/Firefox/IE）用户选择该控件需保存时，进行提示（凭获取信息无法唯一识别）
        :param eleInfo: 网页元素
        :return: True：通过已获取的层级结构及属性信息可唯一定位控件；False：反之
        '''
        from hooker.Hook import ChromeHooker

        # if not driver:
        #     driver = cf.get_value("driver")
        # OM = ObjectMap(driver)

        CH = ChromeHooker()
        try:
            # element = OM.findElebyMethod("xpath", eleInfo, timeout=0.1)
            element = CH.checkElement(eleInfo)
            # OM.highlight(element)
            return eleInfo
        except:
            splitList = re.findall(r"\[@class=(.+?)]", eleInfo)
            for i in range(len(splitList)-1, -1, -1):
                flag = False
                # 通过eleInfo定位失败时，从后往前删除@class属性，直至定位成功
                s = "[@class=%s]" %splitList[i]
                eleInfo = "".join(eleInfo.rsplit(s, 1))
                eleInfo = eleInfo.replace("\'", '\"')
                # mySlice = eleInfo.rfind(s)
                # eleInfo = eleInfo[:mySlice] + eleInfo[mySlice+1:]

                try:
                    # element = OM.findElebyMethod("xpath", eleInfo, timeout=0.1)
                    element = CH.checkElement(eleInfo)
                    flag = True
                except:
                    continue

                if flag:
                    # OM.highlight(element)
                    return eleInfo

            print("唯一性校验失败！")
            return False

    def checkObjFromLog(self, conductType, name):
        ''' 对本地控件库中控件进行唯一性校验
        :param conductType: 控件类型
        :param name: 控件名
        :param projectName: 所在工程名
        '''
        from hooker.Hook import ChromeHooker
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
                print("本地库 [%s] 下 [%s] 控件校验通过！")
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
    def objControl(self, name, conductType, inputStr=None):
        sucFlag = False
        errInfo = ""
        try:
            assert self.dict["Windows"].get(name) is not None, \
                "本地库中未找到名称为 [%s] 的控件，请检查“log.txt”文件！" %name
            info = self.dict["Windows"].get(name).get("Depth")
            obj = self.checkBottom(info)
            assert obj, "根据本地控件信息未定位到目标控件 [%s]！" %name

            if conductType == "点击":
                obj.Click()

            elif conductType == "输入":
                obj.SendKeys(inputStr)

            sucFlag = True
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
                "本地库中无 %s 类型控件，请检查！" %elementType
            assert name in rawDict[elementType].keys(), \
                      "本地库中 %s 类型下未找到名为 %s 控件，请检查！" %(elementType, name)
            objDict = rawDict[elementType][name]
            return objDict
        except AssertionError as e:
            raise e
        except KeyError as e:
            raise e
        except Exception as e:
            raise e

    def openChrome(self, path=r"C:\Users\47612\AppData\Local\Google\Chrome\Applicationpath"):
        ''' 调起chrome进程，可用于后续流程识别
        :param path: chrome安装路径
        '''
        import subprocess
        cmd = 'chrome.exe --remote-debugging-port=9222 --user-data-dir="%s"' %path
        # cmd = r'chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\Users\47612\AppData\Local\Google\Chrome\Applicationpath"'
        popen = subprocess.Popen(cmd, shell=True, stdout=subprocess.PIPE)
        # popen.wait()


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
        elif autoType == "Chrome":
            HK = H.Hooker()
            CH = H.ChromeHooker()

            # CH.keyUp("ctrl")
            # CH.keyUp("shift")

            # 模拟点击“ctrl+shift+x”，并长按“shift”，激活“xpath helper”识别功能
            # CH.keyDown("shift")

            t = threading.Thread(target=CH.refreshDriver, args=[])
            t.setDaemon(True)
            t.start()

            HK.hooks()

            CH.keyUp("shift")
        elif autoType == "IE":
            pass
        elif autoType == "Firefox":
            pass
        else:
            pass

        # os._exit(0)

if __name__ == "__main__":
    AC = AppControl()