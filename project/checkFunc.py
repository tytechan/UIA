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
    # AC.checkObjFromLog("Windows", "test-Null")         # 本地库不存在

    # AC.deleteObjFromLog("Windows", "记事本-确认按钮")            # 成功删除本地控件
    # AC.deleteObjFromLog("Windows", "test2")            # 删除本地库没有的控件

    # AC.deleteObjFromLog("Windows", "记事本-编辑框")

    # ----- Chrome -----
    # AC.checkObjFromLog("Chrome", "abc")               # 本地库无此控件

    # AC.deleteObjFromLog("Chrome", "首页-单据类型")            # 删除本地库没有的控件

    ''' 在已启动driver中调试 '''
    # PF.openChrome()
    AC.checkObjFromLog("Chrome", "登陆-姓名")

