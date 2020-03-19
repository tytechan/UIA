#!python3
# -*- coding: utf-8 -*-
from localSDK.BasicFunc import *
from project import *

AC = AppControl()
PF = PublicFunc()
AC.projectName = projectName
''' 检查环节不写日志 '''
if __name__ == "__main__":
    # ----- Windows -----
    # AC.checkObjFromLog("Windows", "abc")               # 本地库无此控件
    # AC.checkObjFromLog("Windows", "计算器-5")          # 检测成功
    # AC.checkObjFromLog("Windows", "test-Null")         # 本地库不存在x

    # AC.deleteObjFromLog("Windows", "记事本-确认按钮")            # 成功删除本地控件
    # AC.deleteObjFromLog("Windows", "test2")            # 删除本地库没有的控件

    AC.deleteObjFromLog("Windows", "iepath")
    AC.deleteObjFromLog("Windows", "test")
    AC.deleteObjFromLog("Windows", "xpath1")
    AC.deleteObjFromLog("Windows", "xpath2")
    AC.deleteObjFromLog("Windows", "xpath3")
    AC.deleteObjFromLog("Windows", "任务栏")
    AC.deleteObjFromLog("Windows", "微信2")
    AC.deleteObjFromLog("Windows", "微信输入框")

    # ----- Chrome -----
    # AC.checkObjFromLog("Chrome", "abc")               # 本地库无此控件

    # AC.deleteObjFromLog("Chrome", "首页-单据类型")            # 删除本地库没有的控件

    ''' chrome在已启动driver中调试 '''
    # PF.openChrome()
    # AC.checkObjFromLog("Chrome", "登陆-用户名")          # 检测成功
    # AC.checkObjFromLog("Chrome", "登陆-密码")              # 检测成功
    # AC.insertIntoLog("Chrome", "test-插入", "//abc[@id='123123']", projectPath)       # 插入控件及相关信息

    # AC.deleteObjFromLog("IE", "test")            # 删除本地库没有的控件
    # AC.insertIntoLog("IE", "test-插入", "//abc[@id='123123']", projectPath)       # 插入控件及相关信息