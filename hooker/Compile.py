# _*_ coding:utf-8 _*_
import subprocess
import config.Globals as cf
from hooker import *
from config.GUIdesign import *
from localSDK.BasicFunc import *


def myPopen(cmd):
    ''' 执行命令cmd '''
    try:
        # popen = subprocess.Popen(cmd, stdout=subprocess.PIPE)
        popen = subprocess.Popen(cmd, shell=True, stdout=subprocess.PIPE)
        popen.wait()
        lines = popen.stdout.readlines()
        # print(lines)
        dt = compileData(lines)
        return dt

        # return [line.decode('UTF-8') for line in lines]
        # for line in lines:
        #     print(line.decode("UTF-8"))
    except BaseException as e:
        # return -1
        return e
    except Exception as e:
        raise e

def compileData(data):
    ''' 解析 automation.py 返回数据 （list）
    :param data: 列表数据
    :return: 整理后字典
    '''
    myDict = dict()
    myDict["Depth"] = dict()
    for i in range(len(data)):
        info = data[i].decode("UTF-8")
        # print(info)
        if "Starts" in info:
            startTime = info.split("automation.py")[0].strip()
            myDict["Starts"] = startTime
            continue
        elif "Ends" in info:
            endTime = info.split("automation.py")[0].strip()
            myDict["Ends"] = endTime
            continue

        try:
            assert "ControlType" in info
            # dep = getValue(info, "Depth")
            # info = info.split(" ")
            dt = getKeyAndValue(info, ":")
            dep = dt["Depth"]
            myDict["Depth"][dep] = dt
        except:
            pass
    # print("[myDict]\n", myDict)
    return myDict

def getValue(data, key, str):
    ''' 从长字符串获取对应键的值
    :param data:数据字符串
    :param key: 需查找的键名
    :param str: 连词符
    :return: key对应的值
    '''
    s = data.split(key + str)[1].strip()
    r = s.split(" ")[0]
    return r

def getKeyAndValue(data, str):
    ''' 从长字符串自动获取所有键值
    :param data:数据字符串
    :param str:连词符
    :return:键值对dict
    '''
    myDict = dict()
    dt = data.replace("\n", "").replace("\r", "").split(" ")
    for i in range(len(dt)):
        try:
            assert str in dt[i]
            key = dt[i].replace(str, "")
            try:
                assert dt[i].endswith(str)
                myDict[key] = dt[i+1] if (i + 1) < len(dt) else ""
            except:
                s = dt[i].split(str)
                myDict[key] = s[1].strip()
        except:
            pass
    # print("[myDict]\n", myDict)
    return myDict

def compileLogFile():
    ''' 读取日志文件中最后一次录入控件属性 '''
    filePath = parentDirPath + "\log\HookLog.txt"
    file = open(filePath, "rb")
    content = file.read().decode("UTF-8")
    # print(content)
    info = content.split("[层级结构]")[-1]
    # print(info)
    return info

def recordIntoProject(eleProperties):
    ''' 点击目标控件后，判断是否保存到本地库，并定义控件名称
    :param eleProperties: 目标控件所有信息
    :return: True：成功获取并保存/False：不保存
    '''
    import tkinter.messagebox as msg
    import win32api, win32con
    import easygui

    AC = AppControl()
    try:
        keyObj = eleProperties.get("Depth")
        info = keyObj[str(len(keyObj) - 1)]
        # print(info)
        copyInfo = {
            "AutomationId": info["AutomationId"],
            "ClassName": info["ClassName"],
            "ControlType": info["ControlType"],
            "Depth": info["Depth"],
            "Name": info["Name"],
        }

        message = "是否添加控件：\n" + str(copyInfo)
        # result = easygui.boolbox(msg=message, title='提示', choices=('是', '否'), image=None)                # rasygui，可用
        result = win32api.MessageBox(0, message, "提示", win32con.MB_OKCANCEL)                                # pywin32，可用
        # result = msg.askyesnocancel('提示', message)                                                        # tk下总会有空白/多余弹框，且易卡顿，不可用
        # print(result)
        path = cf.get_value("path")
        if result == 1:
            # 保存前脚本进行唯一性校验
            # （True：可直接保存；False：人工进行index判断）
            flag = AC.checkBottom(keyObj)           # TODO：根据判断识别出结果但应用程序实际不可识别处理方法
            if not flag:
                win32api.MessageBox(0, "依据获取信息无法识别控件，请检查控件状态！", "提示", win32con.MB_OK)
                return False

            # 点击后保存
            filePath = path + r"\log.txt"
            # print(filePath)
            with open(filePath, "a+") as f:
                pass
            with open(filePath, "r+") as f:
                rawData = f.read()
                if not rawData:
                    rawData = "{}"
                rawDict = eval(rawData)

                # objName = input("定义控件名称为：")
                objName = getInput("请确认", "请定义控件名称：")

                rawDict[objName] = eleProperties
                # print(rawDict)
            with open(filePath, "w") as f:
                f.write(str(rawDict))
                # f.close()
            return True
        else:
            # 点击后不保存，默认继续识别
            return False
    except Exception as e:
        raise e

if __name__ == "__main__":
    # info = myPopen("automation.py -a -t0 -n")
    # print(info)

    # os_cmd()

    # a = "ControlType: PaneControl    ClassName: #32769    AutomationId:     Rect: (0,0,1280,720)[1280x720]    Name: 桌面 1    Handle: 0x10010(65552)    Depth: 0    SupportedPattern: LegacyIAccessiblePattern"
    # # a = "ControlType: PaneControl    asd:    zxc: 123123    zaa:234431     wer:"
    # # b = getValue(a, "Name", ": ")
    # # print(b)
    #
    # b = getKeyAndValue(a, ":")
    # print(b)

    # compileLogFile()

    recordIntoProject("test")