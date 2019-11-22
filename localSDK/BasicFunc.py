#!python3
# -*- coding: utf-8 -*-
import copy
import uiautomation as auto

class AppControl:
    def __init__(self):
        self.waitTime = 5                # 最长等待时间
        self.intervalTime = 1            # 查找控件间隔时间
        self.dict = None                 # 工程下“log.txt”中的所有控件信息
        self.appName = None              # 应用名称（用于写日志）
        self.maxNum = 9999               # 相同搜索条件下需遍历的控件量最大值

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

    def checkBottom(self, keyObj):
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


    def objControl(self, name, conductType, string=None, **kwargs):
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

    def objControl_EX(self, name, conductType, string=None, **kwargs):
        import win32api
        info = self.dict.get(name)["Depth"]
        objInfo = info[str(len(info) - 1)]
        appInfo = info.get("1")             # info为桌面信息时，只有一层“0”

        # Step 1：定位APP
        appWindow = auto.WindowControl(searchDepth=1, ClassName=appInfo["ClassName"],
                                       RegexName=".*" + appInfo["Name"] + ".*", Desc=self.appName + " window")
        # appWindow.SetActive()
        appWindow.SetTopmost()
        # appWindow.SetFocus()
        try:
            assert appWindow.Exists(maxSearchSeconds=self.waitTime, searchIntervalSeconds=self.intervalTime), \
                "请检查待操作应用程序状态！"


            # Step 2：获取倒数第二层


            ct = objInfo["ControlType"]
            cn = objInfo["ClassName"]
            ai = objInfo["AutomationId"]
            dt = objInfo["Depth"]

            # 获取最下一层相同“ControlType”的控件集合
            objs = appWindow.ButtonControl(searchDepth=dt, foundIndex=None)






            if ct == "TextControl":     # TODO：待补充控件类型
                obj = appWindow.TextControl(searchDepth=int(objInfo["Depth"]), ClassName=objInfo["ClassName"],
                                      AutomationId=objInfo["AutomationId"], RegexName=".*" + objInfo["Name"] + ".*")
            elif ct == "ButtonControl":
                obj = appWindow.ButtonControl(searchDepth=int(objInfo["Depth"]), ClassName=objInfo["ClassName"],
                                      AutomationId=objInfo["AutomationId"], RegexName=".*" + objInfo["Name"] + ".*")
            elif ct == "EditControl":
                obj = appWindow.EditControl(searchDepth=int(objInfo["Depth"]), ClassName=objInfo["ClassName"],
                                      AutomationId=objInfo["AutomationId"], RegexName=".*" + objInfo["Name"] + ".*")
            elif ct == "ComboBoxControl":
                obj = appWindow.ComboBoxControl(searchDepth=int(objInfo["Depth"]), ClassName=objInfo["ClassName"],
                                      AutomationId=objInfo["AutomationId"], RegexName=".*" + objInfo["Name"] + ".*")
            elif ct == "MenuItemControl":
                obj = appWindow.MenuItemControl(searchDepth=int(objInfo["Depth"]), ClassName=objInfo["ClassName"],
                                      AutomationId=objInfo["AutomationId"], RegexName=".*" + objInfo["Name"] + ".*")

            if conductType == "点击":
                obj.Click()
            elif conductType == "双击":
                obj.DoubleClick()
            elif conductType == "输入":
                if ct == "EditControl":
                    obj.GetValuePattern().SetValue(string)
                else:
                    pass
        except AssertionError as e:
            raise e
        except Exception as e:
            raise e



if __name__ == "__main__":
    AC = AppControl()