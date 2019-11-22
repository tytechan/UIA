#!python3
# -*- coding: utf-8 -*-
from localSDK.BasicFunc import *


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
    AC = AppControl()
    AC.dict = readFromLog()
    AC.appName = "SAP"

    # AC.openApp("D:\SAP\SAPgui\saplogon.exe")
    # AC.objControl("SAP-登陆环境", "点击")
    # AC.objControl("SAP-登陆按钮", "点击")
    AC.objControl("SAP-登陆按钮", "点击")
    # AC.objControl("SAP-登陆密码输入框", "输入", "1234qwer")

    # AC.killApp("saplogon.exe")