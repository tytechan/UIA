# _*_ coding:utf-8 _*_

import PyHook3
import threading
import time
import pythoncom
import win32api, win32con
import win32clipboard as wc
import config.Globals as cf
import uiautomation as auto
from hooker import *

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

hm = PyHook3.HookManager()
# cf._init()

class Hooker:
    def __init__(self):
        self.pause = False              # 暂停标志位，触发键盘录制开关时使用（=True，点击不录制/=False，点击录制）
        self.output = None              # 返回结果信息
        self.autoType = cf.get_value("autoType")

    def getTimeNow(self):
        from datetime import datetime
        timeStr = datetime.now()
        timeNow = timeStr.strftime('%D %H:%M:%S.%f')
        return timeNow

    # 鼠标事件处理函数
    def OnMouseEvent(self, event):
        # print('MessageName:',event.MessageName)  # 事件名称
        # print('Message:',event.Message)          # windows消息常量
        # print('Time:',event.Time)                # 事件发生的时间戳
        # print('Window:',event.Window)            # 窗口句柄
        # print('WindowName:',event.WindowName)    # 窗口标题
        #
        # # position = event.Position              # pyhook3中获取鼠标坐标结果与uiautomation不一致，此处用后者
        x, y = auto.GetCursorPos()
        # print('Position:', "(%s, %s)" %(x, y))   # 事件发生时相对于整个屏幕的坐标
        # print('Wheel:',event.Wheel)              # 鼠标滚轮
        # print('Injected:',event.Injected)        # 判断这个事件是否由程序方式生成，而不是正常的人为触发。
        # print('---')

        if event.MessageName == "mouse left down" and not self.pause:
            fp = open(hookLogPath, "a", encoding='utf-8')
            fp.write('\n-----')
            fp.write('\n' + '[捕获时间] %s' %self.getTimeNow())
            fp.write('\n' + '[捕获对象类型] %s' %self.autoType)

            if self.autoType == "Windows":
                from hooker.Compile import myPopen
                fp.write('\n' + '[MessageName] ' + str(event.MessageName))
                fp.write('\n' + '[Message] ' + str(event.Message))
                fp.write('\n' + '[position] ' + "(%s, %s)" %(x, y))
                fp.write('\n' + '[Window] ' + str(event.Window))
                fp.write('\n' + '[WindowName] ' + str(event.WindowName) + '\n')

                # adb = "automation.py –r –d1 –t0 -n"
                # d = os.popen(adb)
                # f = d.read()
                # # print(f)
                # fp.write('\n' + '[path] ' + str(f))
                # ctr(focus=True, seconds=0)
                #X fp.write('\n-----')

                # thr = threading.Thread(ctr, args=(False, True, False, False, False, False, 0xFFFFFFFF, 0,), name="t")
                # thr.start()

                fp.write('\n' + '[层级结构]' + '\n')
                self.output = str(myPopen("automation.py -a -t0 -n"))
                # print("output:", self.output)

                cf.set_value("x", x)
                cf.set_value("y", y)
                fp.write(self.output)
            elif self.autoType == "Chrome":
                CH = ChromeHooker()
                # 释放长按的“shift”
                CH.keyUp("shift")
                self.output = CH.getTarget()

                fp.write('\n\n' + '[层级结构]' + '\n')
                fp.write(self.output)                    # 写入 HookLog，todo：self.output 可进一步处理

            return

        # 返回True代表将事件继续传给其他句柄，为False则停止传递，即被拦截
        return True

    #键盘事件处理函数
    def OnKeyboardEvent(self, event):
        # print('MessageName:',event.MessageName)          #同上，共同属性不再赘述
        # print('Message:',event.Message)
        # print('Time:',event.Time)
        # print('Window:',event.Window)
        # print('WindowName:',event.WindowName)
        # print('Ascii:', event.Ascii, chr(event.Ascii))   #按键的ASCII码
        #
        keyType = event.Key
        print('Key:', keyType)                           #按键的名称
        # print('KeyID:', event.KeyID)                     #按键的虚拟键值
        #
        # print('ScanCode:', event.ScanCode)               #按键扫描码
        # print('Extended:', event.Extended)               #判断是否为增强键盘的扩展键
        # print('Injected:', event.Injected)
        # print('Alt', event.Alt)                          #是否同时按下Alt
        # print('Transition', event.Transition)            #判断转换状态
        # print('---')

        if keyType == "Lmenu":
            # TODO：设置键盘快捷开关（此处为‘Alt’），后期可放开
            # print(self.pause, type(self.pause))
            self.pause = bool(1 - self.pause)

            if self.autoType == "Chrome" and self.pause == False:
                # 类型选择Chrome，且开始录制后两次点击快捷开关（先暂停，后开启录制）
                # 默认此时焦点在driver对应Chrome浏览器上
                CH = ChromeHooker()

                html = WebDriverWait(CH.driver, 5).until(
                    EC.visibility_of_element_located((By.XPATH, '/html')), "请检查Chrome状态！")
                time.sleep(1)
                html.send_keys(Keys.CONTROL, Keys.SHIFT, "x")

                # 模拟点击“ctrl+shift+x”，并长按“shift”，激活“xpath helper”识别功能
                CH.keyDown("shift")

        # 同上
        return True

    def main(self):
        # 创建一个“钩子”管理对象
        # hm = pyHook.HookManager()
        # 监听所有键盘事件
        hm.KeyDown = self.OnKeyboardEvent
        # 设置键盘“钩子”
        hm.HookKeyboard()
        # 监听所有鼠标事件
        hm.MouseAll = self.OnMouseEvent
        # 设置鼠标“钩子”
        hm.HookMouse()

        # 进入循环，如不手动关闭，程序将一直处于监听状态
        pythoncom.PumpMessages(10000)

    def loopToHook(self):
        from hooker.Compile import recordIntoProject_Win, recordIntoProject_Chrome
        # 打印操作轨迹监控
        # path=os.getcwd()        # 获取当前目录
        # fp=open("E:\python相关\RPA_test\log\hook_log.txt","a",encoding='utf-8')

        try:
            # _thread.start_new_thread(main, ())

            t = threading.Thread(target=self.main, args=())              # 创建线程
            t.setDaemon(True)                                       # 设置为后台线程，这里默认是False，设置为True之后则主线程不用等待子线程
            t.start()                                               # 开启线程
            # t.join()
        except:
            print("Error")

        while True:
            # 中断录制功能判断
            if self.output:
                print("[output]\n", self.output)
                hm.UnhookMouse()
                hm.UnhookKeyboard()
                break
                # os._exit(0)
            else:
                pass

        try:
            if self.autoType == "Windows":
                # 添加鼠标点击坐标信息
                # print(type(output), "output:\n", output)
                output = eval(self.output)
                md = output["Depth"]

                assert md != {}, "无法识别所选控件！"

                # 正常识别时
                md[str(len(md) - 1)]["x"] = cf.get_value("x")
                md[str(len(md) - 1)]["y"] = cf.get_value("y")

                isRecord = recordIntoProject_Win(output)
            elif self.autoType == "Chrome":
                assert self.output is not None, ""
                # isRecord = recordIntoProject_Chrome()
                isRecord = None
        except AssertionError as e:
            win32api.MessageBox(0, e, "提示", win32con.MB_OK)
            isRecord = False
        except Exception as e:
            isRecord = False
        finally:
            time.sleep(0.5)
            # cf.set_value("output", None)
            self.output = None
            return isRecord



        # if self.autoType == "Windows":
        #     # 添加鼠标点击坐标信息
        #     # print(type(output), "output:\n", output)
        #     output = eval(self.output)
        #     md = output["Depth"]
        #
        #     if md != {}:
        #         # 正常识别时
        #         md[str(len(md) - 1)]["x"] = cf.get_value("x")
        #         md[str(len(md) - 1)]["y"] = cf.get_value("y")
        #
        #         isRecord = recordIntoProject_Win(output)
        #     else:
        #         # 无法识别控件内容（uiautomation不支持）
        #         win32api.MessageBox(0, "无法识别所选控件！", "提示", win32con.MB_OK)
        #         isRecord = False
        # elif self.autoType == "Chrome":
        #     if self.output:
        #         isRecord = recordIntoProject_Chrome()
        #     else:



    def hooks(self):
        import win32api, win32con
        import easygui
        import tkinter.messagebox as msg

        imgPath = parentDirPath + "\image"
        while True:
            isRecord = self.loopToHook()
            # print(isRecord)
            time.sleep(0.5)
            # if isRecord:
            #     点击后
            #     result = win32api.MessageBox(0, "是否继续识别？", "提示", win32con.MB_OKCANCEL)
            #     result = easygui.boolbox(msg="是否继续识别？", title='提示', choices=('是', '否'), image=None)
            #     result = msg.askyesnocancel('提示', "是否继续识别？")
            #     if not result:
            #         # 不继续识别
            #         break

class ChromeHooker:
    def __init__(self):
        self.chrome_driver = path
        self.chrome_options = Options()
        self.chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")

        self.driver = webdriver.Chrome(self.chrome_driver, chrome_options=self.chrome_options)
        print("[title] ", self.driver.title)
        self.page_source = self.driver.page_source
        self.url = self.driver.current_url

        self.vk_code = {
            'shift': 0x10,
            'ctrl': 0x11,
            'enter': 0x0D,
            'a': 0x41,
            'c': 0x43,
            'v': 0x56,
            'x': 0x58,
        }

    def keyDown(self, KeyboardKeys):
        # 按下按键
        win32api.keybd_event(self.vk_code[KeyboardKeys], 0, 0, 0)

    def keyUp(self, KeyboardKeys):
        # 释放按键
        win32api.keybd_event(self.vk_code[KeyboardKeys], 0, win32con.KEYEVENTF_KEYUP, 0)

    def getCopyText(self):
        wc.OpenClipboard()
        copy_text = wc.GetClipboardData(win32con.CF_TEXT)
        wc.CloseClipboard()
        # result = copy_text.decode("UTF-8")
        result = str(copy_text, encoding="gbk")
        return result

    def refreshDriver(self):
        while True:
            if self.url != self.driver.current_url:
                # 网页有变动，更新driver
                self.url = self.driver.current_url
                self.driver = webdriver.Chrome(self.chrome_driver, chrome_options=self.chrome_options)

                self.page_source = self.driver.page_source
                # print("刷新后 title 为：", self.driver.title, "\n")
                # print(self.page_source)

    def getTarget(self):
        # 借助chrome插件“xpath helper”
        # 切换frame
        newFrame = WebDriverWait(self.driver, 5).until(
            EC.visibility_of_element_located((By.XPATH, '//iframe[@id="xh-bar"]')),
            "未定位到目标frame！")
        self.driver.switch_to.frame(newFrame)

        queryBox = WebDriverWait(self.driver, 5).until(
            EC.visibility_of_element_located((By.XPATH, '//textarea[@id="query"]')),
            "未定位到 xpath helper 扩展程序的 QUERY 文本框！")
        queryBox.click()

        queryBox.send_keys(Keys.CONTROL, "a")         # "ctrl+a"
        queryBox.send_keys(Keys.CONTROL, "c")         # "ctrl+c"

        text = self.getCopyText()
        print("[chrome_text] ", text)
        self.keyUp("shift")
        return text

    def writeFile(self):
        # 向 HookLog.txt 及 工程下 log.txt 写入日志
        pass

if __name__ == "__main__":
    # loopToHook()
    # hooks()
    pass