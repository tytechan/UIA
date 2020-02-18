#!python3
# -*- coding: utf-8 -*-
from localSDK.BasicFunc import *
from localSDK.BrowserFunc import *
from config.DirAndTime import *
import config.Globals as cf
import uiautomation as auto
import frozen_dir

auto.uiautomation.DEBUG_EXIST_DISAPPEAR = True
auto.uiautomation.DEBUG_SEARCH_TIME = True
auto.uiautomation.TIME_OUT_SECOND = 10

# dateDir = createCurrentDateDir("%s\log" %parentDirPath)
dateDir = createCurrentDateDir("%s\log" %frozen_dir.app_path())
auto.Logger.SetLogFile("%s\ExecuteLog.txt" %dateDir)
# print(dateDir)

# 保存全局变量”工程名“，并写入日志
cf._init()
projectName = os.getcwd().split("\\")[-1]
cf.set_value("projectName", projectName)
auto.Logger.WriteLine("-----\n[%s %s] 开始执行工程 '%s'"
                      %(getCurrentDate(), getCurrentTime(), projectName))
# 读取工程下日志文件
PF = PublicFunc()
logDict = PF.readFromLog()

AC = AppControl()
AC.dict = logDict

# ---------- 以上为公共流程 ----------
if __name__ == "__main__":
    # ----- Windows -----

    # # AC.openApp("D:\SAP\SAPgui\saplogon.exe")
    # # AC.objControl("SAP-登陆环境", "点击")
    # # AC.objControl("SAP-登陆按钮", "点击")
    # # AC.objControl("SAP-登陆按钮", "点击")
    # # AC.objControl("SAP-登陆密码输入框", "输入", "1234qwer")
    #
    # # AC.killApp("saplogon.exe")


    AC.appName = "计算器"
    AC.openApp("Calc.exe")
    # AC.objControl("计算器-侧边栏").click()
    # # AC.objControl("计算器-科学", "点击")
    AC.objControl("计算器-5").click()
    AC.objControl("计算器-×").click()
    AC.objControl("计算器-8").click()
    AC.objControl("计算器-等于").click()


    # AC.openApp("notepad.exe")
    # AC.objControl("记事本-格式", "点击")
    # AC.objControl("记事本-字体", "点击")
    # AC.objControl("记事本-字体-倾斜", "点击")
    # AC.objControl("记事本-字体-确认按钮", "点击")
    # # time.sleep(2)
    # AC.objControl("计算器-编辑框", "点击")
    # AC.objControl("计算器-编辑框", "输入", "test")


    # ----- Chrome -----
    # PA = PageAction()
    # PA.open_browser("chrome", capture=True)
    # PA.visit_url("http://cdwp.cnbmxinyun.com")
    #
    # # OM = ObjectMap(PA.driver)
    #
    # # el = OM.findElebyMethod("xpath", '//input[@ng-model="user_name"]')
    # # print(el.get_attribute("placeholder"))
    # # el.send_keys("abc")
    # #
    # # el1 = OM.findElebyMethod("xpath", '//input[@ng-model="password"]')
    # # print(el1.get_attribute("placeholder"))
    # # el1.send_keys("123456")
    #
    #
    # # PA.sendkeys("abc", "xpath", '//input[@ng-model="user_name"]')
    # # PA.sendkeys("123456", "xpath", '//input[@ng-model="password3"]', 0.5)
    #
    #
    # # PA.findElement("xpath", '//input[@ng-model="user_name"]').sendkeys("abc")
    # # PA.findElement("xpath", '//input[@ng-model="password"]', 0.5).sendkeys("123456")
    # # PA.findElement("text", '登录').sendkeys("abc")
    #
    # PA.localElement("登陆-用户名", 1).sendkeys("abc")
    # PA.captureScreen()
    # PA.localElement("登陆-密码1", 1).sendkeys("123456")
    # PA.localElement("登陆-按钮").click()
