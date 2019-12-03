#!python3
# -*- coding: utf-8 -*-
from localSDK.BasicFunc import *
from project import *

AC = AppControl()
AC.projectName = projectName

# AC.checkObjFromLog("abc")               # 本地库无此控件
# AC.checkObjFromLog("计算器-5")          # 检测成功
# AC.checkObjFromLog("test-Null")         # 本地库不存在

# AC.deleteObjFromLog("记事本-确认按钮")            # 成功删除本地控件
# AC.deleteObjFromLog("test2")            # 删除本地库没有的控件




AC.deleteObjFromLog("记事本-编辑框")