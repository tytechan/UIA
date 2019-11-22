#!python3
# -*- coding: utf-8 -*-

import os
import sys
import time
import subprocess
from distinguish import *

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # not required after 'pip install uiautomation'
import uiautomation as auto

auto.uiautomation.DEBUG_EXIST_DISAPPEAR = True  # set it to False and try again, default is False
auto.uiautomation.DEBUG_SEARCH_TIME = True  # set it to False and try again, default is False
auto.uiautomation.TIME_OUT_SECOND = 10  # global time out
auto.Logger.SetLogFile(parentDirPath + "\log\ExecuteLog.txt")


def Calc(window, btns, expression):
    expression = ''.join(expression.split())
    if not expression.endswith('='):
        expression += '='
    for char in expression:
        auto.Logger.Write(char, writeToFile = True)
        btns[char].Click(waitTime=0.05)
    time.sleep(0.1)
    window.SendKeys('{Ctrl}c', waitTime = 0.1)
    result = auto.GetClipboardText()
    auto.Logger.WriteLine(result, auto.ConsoleColor.Cyan, writeToFile = True)
    time.sleep(1)

def writeLog(func):
    def wrapper(*args, **kwargs):
        try:
            obj = func(*args, **kwargs)
            auto.Logger.WriteLine(obj, auto.ConsoleColor.Cyan, writeToFile = True)
            return obj
        except Exception as e:
            raise e
    return wrapper

def findObj(ControlType, searchDepth, PControl=None, **searchProperties):
    ''' 根据属性再控件树中查找元素
    :param PControl: 为空时从桌面开始遍历，有值时从指定层向下遍历
    :return: 目标控件
    '''
    parent = PControl if PControl else auto
    if ControlType == "WindowControl":
        obj = parent.WindowControl(ControlType=ControlType, searchDepth=searchDepth, **searchProperties)
    elif ControlType == "ButtonControl":
        obj = parent.ButtonControl(ControlType=ControlType, searchDepth=searchDepth, **searchProperties)
    elif ControlType == "EditControl":
        obj = parent.EditControl(ControlType=ControlType, searchDepth=searchDepth, **searchProperties)
    elif ControlType == "TextControl":
        obj = parent.TextControl(ControlType=ControlType, searchDepth=searchDepth, **searchProperties)
    elif ControlType == "ListControl":
        obj = parent.ListControl(ControlType=ControlType, searchDepth=searchDepth, **searchProperties)
    elif ControlType == "ListItemControl":
        obj = parent.ListItemControl(ControlType=ControlType, searchDepth=searchDepth, **searchProperties)
    try:
        if obj:
            return obj
    except Exception as e:
        # c = auto.Control()
        # print(c.GetSearchPropertiesStr())
        auto.Logger.Write("控件定位失败！", auto.ConsoleColor.Cyan, writeToFile=False)
        # obj = None
        raise e
    # finally:
    #     return obj

if __name__ == '__main__':
    # CalcOnWindows10()

    # auto.Logger.Write('\nPress any key to exit', auto.ConsoleColor.Cyan)
    # import msvcrt
    # while not msvcrt.kbhit():           # 判断按键是否不在等待读取
    #     pass

    # ------------------------
    # win = findObj(ControlType="WindowControl", searchDepth=1, ClassName="#32770")
    win = findObj(ControlType="WindowControl", searchDepth=1, RegexName = '.*SAP.*')          # 正确
    # win = findObj(ControlType="WindowControl", searchDepth=1, Name="SAP")                       # 错误

    # win = findObj(ControlType="TextControl", searchDepth=8, Name="ET1")
    # win = findObj(ControlType="TextControl", searchDepth=8, SupportedPattern="LegacyIAccessiblePattern", Name="ET1")
    # win = findObj(ControlType="TextControl", searchDepth=8, Name="ET7")
    print(win)
    # obj = findChild(win, ControlType="TextControl", searchDepth=8, Name="810")
    obj = findObj(ControlType="TextControl", searchDepth=8, PControl=win)
    print(obj)

    # win.Click()