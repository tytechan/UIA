# _*_ coding:utf-8 _*_

import PyHook3
import threading
import time
import pythoncom
import win32api, win32con
import config.Globals as cf
import uiautomation as auto
from hooker import *


hm = PyHook3.HookManager()
output = None
cf._init()


def getTimeNow():
    from datetime import datetime
    timeStr = datetime.now()
    timeNow = timeStr.strftime('%D %H:%M:%S.%f')
    return timeNow

# 鼠标事件处理函数
def OnMouseEvent(event):
    print('MessageName:',event.MessageName)  # 事件名称
    print('Message:',event.Message)          # windows消息常量
    print('Time:',event.Time)                # 事件发生的时间戳
    print('Window:',event.Window)            # 窗口句柄
    print('WindowName:',event.WindowName)    # 窗口标题

    # position = event.Position              # pyhook3中获取鼠标坐标结果与uiautomation不一致，此处用后者
    x, y = auto.GetCursorPos()
    print('Position:', "(%s, %s)" %(x, y))   # 事件发生时相对于整个屏幕的坐标
    print('Wheel:',event.Wheel)              # 鼠标滚轮
    print('Injected:',event.Injected)        # 判断这个事件是否由程序方式生成，而不是正常的人为触发。
    print('---')

    if event.MessageName == "mouse left down":
        from hooker.Compile import myPopen, recordIntoProject
        fp = open(parentDirPath + "\log\HookLog.txt","a",encoding='utf-8')
        fp.write('\n-----')
        fp.write('\n' + '[捕获时间] %s' %getTimeNow())
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
        # fp.write('\n-----')

        # thr = threading.Thread(ctr, args=(False, True, False, False, False, False, 0xFFFFFFFF, 0,), name="t")
        # thr.start()

        fp.write('\n' + '[层级结构]' + '\n')
        output = str(myPopen("automation.py -a -t0 -n"))
        cf.set_value("output", output)

        cf.set_value("x", x)
        cf.set_value("y", y)
        fp.write(output)

        # recordIntoProject(output)
        return output

    # 返回True代表将事件继续传给其他句柄，为False则停止传递，即被拦截
    return True

#键盘事件处理函数
def OnKeyboardEvent(event):
    print('MessageName:',event.MessageName)          #同上，共同属性不再赘述
    print('Message:',event.Message)
    print('Time:',event.Time)
    print('Window:',event.Window)
    print('WindowName:',event.WindowName)
    print('Ascii:', event.Ascii, chr(event.Ascii))   #按键的ASCII码
    print('Key:', event.Key)                         #按键的名称
    print('KeyID:', event.KeyID)                     #按键的虚拟键值
    print('ScanCode:', event.ScanCode)               #按键扫描码
    print('Extended:', event.Extended)              #判断是否为增强键盘的扩展键
    print('Injected:', event.Injected)
    print('Alt', event.Alt)                          #是否同时按下Alt
    print('Transition', event.Transition)            #判断转换状态
    print('---')

    # 同上
    return True

def main_test():
    #将OnMouseEvent函数绑定到MouseAllButtonsDown事件上
    hm.MouseAllButtonsDown = OnMouseEvent
    #将OnKeyboardEvent函数绑定到KeyDown事件上
    hm.KeyDown = OnKeyboardEvent


    hm.HookMouse()        #设置鼠标钩子
    hm.HookKeyboard()   #设置键盘钩子

    time.sleep(20)

    hm.UnhookMouse()    #取消鼠标钩子
    hm.UnhookKeyboard() #取消键盘钩子


def main():
    # 创建一个“钩子”管理对象
    # hm = pyHook.HookManager()
    # 监听所有键盘事件
    hm.KeyDown = OnKeyboardEvent
    # 设置键盘“钩子”
    hm.HookKeyboard()
    # 监听所有鼠标事件
    hm.MouseAll = OnMouseEvent
    # 设置鼠标“钩子”
    hm.HookMouse()

    # 进入循环，如不手动关闭，程序将一直处于监听状态
    pythoncom.PumpMessages(10000)

def loopToHook():
    from hooker.Compile import recordIntoProject
    # 打印操作轨迹监控
    # path=os.getcwd()        # 获取当前目录
    # fp=open("E:\python相关\RPA_test\log\hook_log.txt","a",encoding='utf-8')

    try:
        # _thread.start_new_thread(main, ())

        t = threading.Thread(target=main, args=())              # 创建线程
        t.setDaemon(True)                                       # 设置为后台线程，这里默认是False，设置为True之后则主线程不用等待子线程
        t.start()                                               # 开启线程
    except:
        print("Error")

    while True:
        output = cf.get_value("output")
        if output:
            print("[output]\n", output)
            hm.UnhookMouse()
            hm.UnhookKeyboard()
            break
            # os._exit(0)
        else:
            pass

    # 添加鼠标点击坐标信息
    # print(type(output), "output:\n", output)
    output = eval(output)
    md = output["Depth"]

    if md != {}:
        # 正常识别时
        md[str(len(md) - 1)]["x"] = cf.get_value("x")
        md[str(len(md) - 1)]["y"] = cf.get_value("y")

        isRecord = recordIntoProject(output)
    else:
        # 无法识别控件内容（uiautomation不支持）
        win32api.MessageBox(0, "无法识别所选控件！", "提示", win32con.MB_OK)
        isRecord = False

    time.sleep(0.5)
    cf.set_value("output", None)
    return isRecord

def hooks():
    import win32api, win32con
    import easygui
    import tkinter.messagebox as msg

    imgPath = parentDirPath + "\image"
    while True:
        isRecord = loopToHook()
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

if __name__ == "__main__":
    loopToHook()
    # hooks()