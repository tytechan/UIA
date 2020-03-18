# encoding = utf - 8

import time, datetime
import os
import win32api, win32con

import subprocess
import win32com.client
import psutil
from config.ErrConfig import *

class SAP:
    def __init__(self):
        self.waitTime = 10          # 默认最长等待时间为10s
        self.session = None
        self.connection = None
        self.activewindow = None
        self.element = None
        self.shell = None

    def wait(self, objStr, time=3):
        for i in range(time):
            try:
                obj = eval(objStr)
                print("SAP成功启动！")
                assert type(obj) == win32com.client.CDispatch
                return obj
            except Exception as e:
                if i + 1 < time:
                    continue
                else:
                    raise e

    @CNBMException
    def open(self, path, env):
        ''' 调起SAP服务
        :param path: SAP执行文件路径
        :param env: 本地登陆环境名
        :return: 全局变量session
        '''
        try:
            subprocess.Popen(path)
            time.sleep(1)

            SapGuiAuto = self.wait("win32com.client.GetObject('SAPGUI')")
            application = self.wait("SapGuiAuto.GetScriptingEngine")
            self.connection = self.wait("application.OpenConnection(env, True)")
            self.session = self.wait("connection.Children(0)")
        except Exception as e:
            handleErr(e)
            raise e

    @CNBMException
    def login(self, userName, passWord):
        ''' 登陆SAP '''
        try:
            self.getObj("id", "wnd[0]/usr/txtRSYST-BNAME").sendkeys(userName)
            self.getObj("id", "wnd[0]/usr/pwdRSYST-BCODE").sendkeys(passWord)
            self.getObj("id", "wnd[0]").keys("Enter")
            try:
                # 等待“多次登陆”弹出框2s
                self.waitUntil(session.Children.count == 2, "False", valueReturned="pass", maxSec=2)
                # 父对象，'/app/con[i]'
                ''' 方法一 .Parent '''
                # p = self.session.Parent
                # try:
                #     if p.findById("ses[0]/wnd[1]"):
                #         self.session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").select()
                #         self.session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").setFocus()
                #         self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                # except:
                #     pass

                ''' 方法二 .ActiveWindow.FindAllByName '''
                self.updateActiveWindow()
                if self.activewindow.FindAllByName("wnd[1]", "GuiModalWindow").count == 1:
                    self.getObj("id", "usr/radMULTI_LOGON_OPT2").set()
                    self.getObj("id", "usr/radMULTI_LOGON_OPT2").focus()
                    self.getObj("id", "tbar[0]/btn[0]").click()
            except:
                pass

            # 等待“信息”弹出框1s
            try:
                self.getObj("name", "wnd[1]", "GuiModalWindow",
                            text="信息", valueReturned="pass", maxTime=1)
                self.getObj("id", "wnd[1]/tbar[0]/btn[0]").click()
            except:
                pass

            self.getObj("name", "wnd[0]", "GuiMainWindow", text="SAP 轻松访问")
            # print("SAP成功登陆！")
        except Exception as e:
            handleErr(e)
            raise e

    @CNBMException
    def createNewSession(self):
        '''
        创建新session（窗口）
        '''
        try:
            count = self.connection.children.count
            # session数+1
            self.session.createSession()
            time.sleep(1)
            self.session = self.connection.children(count)
            for i in range(2):
                self.updateActiveWindow()
                try:
                    self.getObj("name", "titl", "GuiTitlebar",
                                text="SAP 轻松访问 中建信息", valueReturned="pass", maxTime=3)
                    return
                except:
                    pass
            self.updateActiveWindow()
            self.getObj("name", "titl", "GuiTitlebar",
                        text="SAP 轻松访问 中建信息", maxTime=3)
        except Exception as e:
            handleErr(e)
            raise e

    @CNBMException
    def updateActiveWindow(self):
        '''
        更新ActiveWindow
        '''
        try:
            self.activewindow = self.session.ActiveWindow
        except Exception as e:
            handleErr(e)
            raise e

    @CNBMException
    def getObj(self, case, infoA, infoB=None, text=None, valueReturned="err", maxTime=None):
        '''
        循环等待对象
        :param case: 属性（id/name）|属性值
        :param infoA: id值/name值
        :param infoB: 选填，case为name时，对应控件type属性值
        :param text: 选填，对象text属性（有值校验，不填不检验）
        :param valueReturned: 未等到结果时结束方式（err：报错/pass：通过）
        :param maxTime: 选填，最长等待时间（默认为self.waitTime）
        '''
        # 数据初始化
        try:
            self.updateActiveWindow()
            maxSec = self.waitTime if not maxTime else int(maxTime)

            for i in range(maxSec):
                try:
                    if case == "id":
                        self.element = self.activewindow.findById(infoA)
                    elif case == "name":
                        if text:
                            objs = self.activewindow.FindAllByName(infoA, infoB)
                            for j in range(objs.count):
                                if objs[j].text == text:
                                    self.element = objs[j]
                                    break
                                assert objs.count != j + 1
                        else:
                            self.element = self.activewindow.FindByName(infoB, infoB)
                    return SAPElement(self.activewindow, self.element)
                except:
                    try:
                        assert maxSec != i + 1, \
                             "%d s内未找到 '%s' 为 '%s' 的对象！" %(maxSec, case, infoA)
                        time.sleep(1)
                    except AssertionError as e:
                        if valueReturned == "err":
                            raise AssertionError(e)
                        elif valueReturned == "pass":
                            pass
                    except Exception as e:
                        raise e
        except Exception as e:
            handleErr(e)
            raise e

    @CNBMException
    def getShell(self, case, infoA, infoB=None, text=None):
        '''
        获取table类型对象（分type为GuiShell及常规table两种类型）
        :param case: id/name
        :param infoA: id属性值/name属性值
        :param infoB: 选填，type属性值（case为name时必填）
        '''
        try:
            self.updateActiveWindow()
            self.shell = self.getObj(case, infoA, infoB=infoB, text=text)
            return SAPGuiShell(self.shell)
        except Exception as e:
            handleErr(e)
            raise e

    @CNBMException
    def waitUntil(self, event, condition, valueReturned="err", maxTime=None):
        '''
        循环等待（event是否发生与condition一致时，等待，否则跳出）
        :param event: 待判断主体（布尔类型）
        :param condition: True/False
        :param valueReturned: 未等到结果时结束方式（err：报错/pass：通过）
        :param maxTime: 选填，最长等待时间（默认为self.waitTime）
        '''
        try:
            maxSec = self.waitTime if not maxTime else int(maxTime)
            for i in range(maxSec):
                try:
                    assert maxSec != i + 1, \
                        str(maxSec) + " s内均未等到指定事件发生！"
                    if bool(event) == eval(condition):
                        time.sleep(1)
                    else:
                        return
                except AssertionError as e:
                    if valueReturned == "err":
                        raise AssertionError(e)
                    elif valueReturned == "pass":
                        pass
                except Exception as e:
                    raise e
        except Exception as e:
            handleErr(e)
            raise e

    @CNBMException
    def closeSessions(self):
        try:
            if not self.connection:
                return

            sessions = self.connection.children
            for i in range(sessions.count):
                # 遍历 sessions.count 个session
                for j in range(10):
                    try:
                        wndId = "wnd[" + str(j) + "]"
                        window = sessions[i].findById(wndId)
                        window.close()
                        try:
                            # 存在“注销”弹出框，则为最后一个window
                            msgBox = sessions[i].findById("wnd[" + str(j + 1) + "]")
                            assert msgBox.text == "注销"
                            msgBox.findById("usr/btnSPOP-OPTION1").press()
                            self.session = None
                            return
                        except:
                            pass
                    except:
                        break
        except Exception as e:
            handleErr(e)
            raise e

    @CNBMException
    def kill(self):
        '''
        结束SAP（所有）进程
        '''
        try:
            pids = psutil.pids()
            for pid in pids:
                p = psutil.Process(pid)
                # print('pid-%s,pname-%s' % (pid, p.name()))
                if p.name() == 'saplogon.exe':
                    cmd = 'taskkill /F /IM saplogon.exe'
                    # 无需结束sap进程时，注释下行
                    os.system(cmd)
        except Exception as e:
            handleErr(e)
            raise e


class SAPElement:
    def __init__(self, activeWindow, element):
        self.activeWindow = activeWindow
        self.element = element
        # 键盘模拟对应表
        self.keysDict = {
            "Ctrl+/": 72,
            "Ctrl+C": 77,
            "Ctrl+E": 70,
            "Ctrl+F": 71,
            "Ctrl+F1": 25,
            "Ctrl+F10": 34,
            "Ctrl+F11": 35,
            "Ctrl+F12": 36,
            "Ctrl+F2": 26,
            "Ctrl+F3": 27,
            "Ctrl+F4": 28,
            "Ctrl+F5": 29,
            "Ctrl+F6": 30,
            "Ctrl+F7": 31,
            "Ctrl+F8": 32,
            "Ctrl+F9": 33,
            "Ctrl+G": 84,
            "Ctrl+N": 74,
            "Ctrl+O": 75,
            "Ctrl+P": 86,
            "Ctrl+PageDown": 83,
            "Ctrl+PageUp": 80,
            "Ctrl+R": 85,
            "Ctrl+S": 11,
            "Ctrl+Shift+F1": 37,
            "Ctrl+Shift+F10": 46,
            "Ctrl+Shift+F11": 47,
            "Ctrl+Shift+F12": 48,
            "Ctrl+Shift+F2": 38,
            "Ctrl+Shift+F3": 39,
            "Ctrl+Shift+F4": 40,
            "Ctrl+Shift+F5": 41,
            "Ctrl+Shift+F6": 42,
            "Ctrl+Shift+F7": 43,
            "Ctrl+Shift+F8": 44,
            "Ctrl+Shift+F9": 45,
            "Ctrl+V": 78,
            "Ctrl+X": 76,
            "Ctrl+Z": 79,
            "Ctrl+\\": 73,
            "Enter": 0,
            "F1": 1,
            "F10": 10,
            "F12": 12,
            "F2": 2,
            "F3": 3,
            "F4": 4,
            "F5": 5,
            "F6": 6,
            "F7": 7,
            "F8": 8,
            "F9": 9,
            "PageDown": 82,
            "PageUp": 81,
            "Shift+Ctrl+0": 22,
            "Shift+F1": 13,
            "Shift+F11": 23,
            "Shift+F12": 24,
            "Shift+F2": 14,
            "Shift+F3": 15,
            "Shift+F4": 16,
            "Shift+F5": 17,
            "Shift+F6": 18,
            "Shift+F7": 19,
            "Shift+F8": 20,
            "Shift+F9": 21
        }

    @CNBMException
    def click(self):
        '''
        左击控件
        '''
        try:
            self.element.press()
        except Exception as e:
            handleErr(e)
            raise e

    @CNBMException
    def click(self):
        '''
        双击控件
        '''
        try:
            self.element.doubleClick()
        except Exception as e:
            handleErr(e)
            raise e

    @CNBMException
    def sendkeys(self, text):
        '''
        输入框输值
        :param text: 待输入字段
        '''
        try:
            self.element.text = text
        except Exception as e:
            handleErr(e)
            raise e

    @CNBMException
    def select(self, text):
        '''
        常规下拉框选择
        :param text: 待输入字段
        '''
        try:
            self.element.key = text
        except Exception as e:
            handleErr(e)
            raise e

    @CNBMException
    def set(self):
        '''
        选择控件
        '''
        try:
            self.element.Select()
        except Exception as e:
            handleErr(e)
            raise e

    @CNBMException
    def choose(self, text):
        '''
        勾选框选择
        :param text: -1（勾选）/0（取消勾选）
        '''
        try:
            if self.element.type in ["GuiCheckBox"]:
                # 示例type：GuiCheckBox
                self.element.selected = int(text)
            elif self.element.type in ["GuiShell"]:
                # 示例type：GuiShell
                self.element.modifyCheckbox(0, "SEL", int(text))
                self.element.triggerModified()
        except Exception as e:
            handleErr(e)
            raise e

    @CNBMException
    def focus(self):
        '''
        聚焦控件
        '''
        try:
            self.element.setFocus()
        except Exception as e:
            handleErr(e)
            raise e

    @CNBMException
    def keys(self, text):
        '''
        模拟键盘输值
        :param text: 参照self.keys
        '''
        try:
            if text in self.keys.keys():
                self.element.sendVKey(self.keysDict[text])
            else:
                self.element.sendVKey(text)
        except Exception as e:
            handleErr(e)
            raise e

    @CNBMException
    def close(self):
        '''
        关闭控件
        '''
        try:
            self.element.close()
        except Exception as e:
            handleErr(e)
            raise e

    @CNBMException
    def text(self):
        '''
        获取控件文本
        '''
        try:
            return self.element.text
        except Exception as e:
            handleErr(e)
            raise e


class SAPGuiShell:
    def __init__(self, shell):
        self.shell = shell

    @CNBMException
    def sendkeys(self, row, colInfo, text):
        '''
        向GuiShell类型table的具体行列输值
        :param row: 行数，从1开始
        :param colInfo: 列信息，通过tracker获取
        :param text: 待输值
        '''
        try:
            self.shell.modifyCell(row, colInfo, text)
        except Exception as e:
            handleErr(e)
            raise e

    @CNBMException
    def pressButton(self, btnType):
        '''
        向GuiShell类型table的具体行列输值
        :param btnType: 按钮功能类型
        '''
        try:
            if btnType == "提交":
                self.shell.pressToolbarButton("ZPOST")
            elif btnType == "全部选择":
                self.shell.pressToolbarButton("ZSALL")
            elif btnType == "清空选择":
                self.shell.pressToolbarButton("ZDSAL")
        except Exception as e:
            handleErr(e)
            raise e

    @CNBMException
    def selectMenuItem(self, menuType):
        '''
        向GuiShell类型table的具体行列输值
        :param btnType: 按钮功能类型
        '''
        try:
            if menuType == "提交":
                self.shell.selectContextMenuItem("&DETAIL")
        except Exception as e:
            handleErr(e)
            raise e


if __name__ == "__main__":
    pass