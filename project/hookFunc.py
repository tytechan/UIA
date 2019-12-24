#!python3
# -*- coding: utf-8 -*-
import tkinter as tk
import subprocess
from tkinter import ttk
from hooker.Hook import *


def openChrome(path=r"C:\Users\47612\AppData\Local\Google\Chrome\Applicationpath"):
    ''' 调起chrome进程，可用于后续流程识别

    '''
    cmd = 'chrome.exe --remote-debugging-port=9222 --user-data-dir="%s"' %path
    # cmd = r'chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\Users\47612\AppData\Local\Google\Chrome\Applicationpath"'
    popen = subprocess.Popen(cmd, shell=True, stdout=subprocess.PIPE)
    # popen.wait()


def clickMe():
    # button被点击之后会被执行
    global autoType
    win.destroy()
    autoType = browser.get()
    # print(result)

win = tk.Tk()
# win.geometry("300x150+500+200")  # 大小和位置
win.title("请确认")
ttk.Label(win, text="请选择浏览器类型：").grid(column=1, row=0)  # 添加一个标签，并将其列设置为1，行设置为0

# 按钮
action = ttk.Button(win, text="选择", command=clickMe)
action.grid(column=2, row=1)

# 创建一个下拉列表
browser = tk.StringVar()
browserChosen = ttk.Combobox(win, width=12, textvariable=browser)
browserChosen['values'] = ("", "Chrome", "IE", "Firefox", "Windows")
browserChosen.grid(column=1, row=1)
# 设置下拉列表默认显示的值，0为 numberChosen['values'] 的下标值
browserChosen.current(0)
# 当调用mainloop()时,窗口才会显示出来
win.mainloop()


if __name__ == "__main__":
    # 打开自调chrome
    # openChrome()

    print("本次录制类型:", autoType)

    if autoType == "Windows":

        cf._init()
        path = os.getcwd()
        cf.set_value("path", path)

        HK = Hooker()
        HK.hooks()
    elif autoType == "Chrome":
        pass
    elif autoType == "IE":
        pass
    elif autoType == "Firefox":
        pass
    else:
        pass

    # os._exit(0)