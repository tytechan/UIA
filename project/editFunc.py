#!python3
# -*- coding: utf-8 -*-
from localSDK.BasicFunc import *
from project import parentDirPath

auto.uiautomation.DEBUG_EXIST_DISAPPEAR = True
auto.uiautomation.DEBUG_SEARCH_TIME = True
auto.uiautomation.TIME_OUT_SECOND = 10
auto.Logger.SetLogFile(parentDirPath + "\log\ExecuteLog.txt")

def readFromLog():
    ''' 读取log文件信息 '''
    import win32api, win32con

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

# 读取log内容
filePath = "log.txt"

if __name__ == "__main__":
    import time
    AC = AppControl()
    AC.dict = readFromLog()
    AC.appName = "计算器"

    # AC.openApp("D:\SAP\SAPgui\saplogon.exe")
    # AC.objControl("SAP-登陆环境", "点击")
    # AC.objControl("SAP-登陆按钮", "点击")
    # AC.objControl("SAP-登陆按钮", "点击")
    # AC.objControl("SAP-登陆密码输入框", "输入", "1234qwer")

    # AC.killApp("saplogon.exe")


    AC.openApp("Calc.exe")
    # AC.objControl("计算器-侧边栏", "点击")
    # AC.objControl("计算器-科学", "点击")
    AC.objControl("计算器-5", "点击")
    AC.objControl("计算器-×", "点击")
    AC.objControl("计算器-7", "点击")
    AC.objControl("计算器-等于", "点击")


    AC.openApp("notepad.exe")
    AC.objControl("记事本-格式", "点击")
    AC.objControl("记事本-字体", "点击")
    AC.objControl("记事本-字体-倾斜", "点击")
    AC.objControl("记事本-字体-确认按钮", "点击")
    # time.sleep(2)
    AC.objControl("计算器-编辑框", "点击")
    AC.objControl("计算器-编辑框", "输入", "test")

