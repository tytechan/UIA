#!python3
# -*- coding: utf-8 -*-
from localSDK.BasicFunc import *
from localSDK.BrowserFunc import *
from localSDK.ParseExcel import *
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

browser = PageAction()
win = AppControl()
key = KeyboardKeys()
win.dict = logDict
excel = ParseExcel()

# ---------- 以上为公共流程 ----------
if __name__ == "__main__":
    # ----- Windows -----

    # # win.openApp("D:\SAP\SAPgui\saplogon.exe")
    # # win.objControl("SAP-登陆环境", "点击")
    # # win.objControl("SAP-登陆按钮", "点击")
    # # win.objControl("SAP-登陆按钮", "点击")
    # # win.objControl("SAP-登陆密码输入框", "输入", "1234qwer")
    #
    # # win.killApp("saplogon.exe")


    # win.appName = "计算器"
    # win.openApp("Calc.exe")
    # # win.objControl("计算器-侧边栏").click()
    # # # win.objControl("计算器-科学", "点击")
    # win.objControl("计算器-5").click()
    # win.objControl("计算器-×").click()
    # win.objControl("计算器-8").click()
    # win.objControl("计算器-等于").click()


    # win.openApp("notepad.exe")
    # win.objControl("记事本-格式", "点击")
    # win.objControl("记事本-字体", "点击")
    # win.objControl("记事本-字体-倾斜", "点击")
    # win.objControl("记事本-字体-确认按钮", "点击")
    # # time.sleep(2)
    # win.objControl("计算器-编辑框", "点击")
    # win.objControl("计算器-编辑框", "输入", "test")


    # ----- Chrome -----
    browser.type = "Chrome"
    # browser.open_browser(r"\Users\47612\AppData\Local\Google\Chrome\User Data\Default")
    browser.open_browser(r"\Users\47612\AppData\Local\Google\Chrome\User Data")
    browser.visit_url("http://cdwp.cnbmxinyun.com")

    browser.localElement("登陆-用户名", timeout=1).sendkeys("abc")
    attr = browser.localElement("登陆-用户名", timeout=1).get_attribute("class")
    print(attr)
    browser.captureScreen()
    browser.localElement("登陆-密码1", timeout=1).sendkeys("123456")
    browser.localElement("登陆-按钮").click()
