# _*_ coding:utf-8 _*_

import PyHook3
import threading
import pythoncom
from hooker import *
from localSDK.BrowserFunc import *
from selenium.webdriver.common.keys import Keys

hm = PyHook3.HookManager()

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

        if not self.pause:

            '''
            1、鼠标点击触发录制；
            2、键盘快捷开关已打开
            '''
            if self.autoType == "IE":
                ''' IE '''
                try:
                    # 鼠标右键触发录制
                    assert event.MessageName == "mouse right down"

                    from hooker.Compile import myPopen
                    self.output = str(myPopen("automation.py -a -t0 -n"))

                    fp = open(hookLogPath, "a", encoding='utf-8')
                    fp.write('\n-----')
                    fp.write('\n' + '[捕获时间] %s' %self.getTimeNow())
                    fp.write('\n' + '[捕获对象类型] %s' %self.autoType)
                    fp.write('\n\n' + '[层级结构]' + '\n')
                    fp.write(self.output)                    # 写入 HookLog，todo：self.output 可进一步处理
                except AssertionError:
                    pass

            else:
                ''' Windows '''
                try:
                    # 鼠标左键触发录制
                    assert event.MessageName == "mouse left down"

                    fp = open(hookLogPath, "a", encoding='utf-8')
                    fp.write('\n-----')
                    fp.write('\n' + '[捕获时间] %s' % self.getTimeNow())
                    fp.write('\n' + '[捕获对象类型] %s' % self.autoType)

                    if self.autoType == "Windows":
                        from hooker.Compile import myPopen

                        fp.write('\n' + '[层级结构]' + '\n')
                        self.output = str(myPopen("automation.py -a -t0 -n"))
                        # print("output:", self.output)

                        cf.set_value("x", x)
                        cf.set_value("y", y)
                        fp.write(self.output)
                    return
                except AssertionError:
                    pass
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


        if keyType == "Delete":
            # TODO: 退出录制快捷键“Del”（shift+esc 为打开chrome任务管理器默认快捷键），后期可放开
            hm.UnhookMouse()
            hm.UnhookKeyboard()

            if self.autoType == "IE" or self.autoType == "Windows":
                cmd = r"taskkill /F /IM IEXpath.exe"
                os.system(cmd)

            print("退出录制！")
            os._exit(0)

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
        from hooker.Compile import recordIntoProject

        try:
            t = threading.Thread(target=self.main, args=())              # 创建线程
            t.setDaemon(True)                                            # 设置为后台线程，这里默认是False，设置为True之后则主线程不用等待子线程
            t.start()                                                    # 开启线程
            # t.join()
        except:
            print("Error")

        while True:
            # 中断录制功能判断
            try:
                # assert self.output and eval(self.output)
                assert self.output
                print("[output]\n", self.output, "\n")
                hm.UnhookMouse()
                hm.UnhookKeyboard()
                break
                # os._exit(0)
            except:
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

                isRecord = recordIntoProject(output)
            elif self.autoType == "IE":
                output = eval(self.output)
                isRecord = recordIntoProject(output)

        except AssertionError as e:
            isRecord = False
            win32api.MessageBox(0, e, "提示", win32con.MB_OK)
        except Exception as e:
            isRecord = False
        finally:
            time.sleep(0.5)
            # cf.set_value("output", None)
            self.output = None
            return isRecord

    def hooks(self):
        while True:
            isRecord = self.loopToHook()
            # print(isRecord)
            time.sleep(0.5)

if __name__ == "__main__":
    # loopToHook()
    # hooks()
    pass