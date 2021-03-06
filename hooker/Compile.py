# _*_ coding:utf-8 _*_
from hooker import *
from config.GUIdesign import *
from localSDK.BasicFunc import *
import subprocess
import config.Globals as cf

autoType = cf.get_value("autoType")

def myPopen(cmd):
    ''' 执行命令cmd '''
    try:
        # popen = subprocess.Popen(cmd, stdout=subprocess.PIPE)
        popen = subprocess.Popen(cmd, shell=True, stdout=subprocess.PIPE)
        popen.wait()
        lines = popen.stdout.readlines()
        # print(["lines"], type(lines), "\n", lines)

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
    if autoType == "IE":
        for i in range(len(data)-1, -1, -1):
            # 倒序遍历
            info = data[i].decode("UTF-8")
            # print(info)

            # 在IE空间中先遍历到最下层，再把倒数第二层作为校验点
            if "ControlType" in info:
                try:
                    assert "ControlType: ListItemControl" in info
                    # 最下层控件类型符合条件
                    parentInfo = data[i - 1].decode("UTF-8")
                    assert "XPath（选择后双击复制）" in parentInfo
                    break
                except AssertionError:
                    '''
                    退出情况（未点击IEPath的xpath栏）：
                    1、最下层非“ListItemControl”；
                    2、倒数第二层不满足条件
                    '''
                    return {}
            else:
                # 继续遍历直至找到第一个控件层（最下层）
                continue

    myDict = dict()
    myDict["Depth"] = dict()
    for i in range(len(data)):
        info = data[i].decode("UTF-8")
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
    ''' 点击目标控件（IE）后，判断是否保存到本地库，并定义控件名称（不做唯一性校验）
    :param eleProperties: 目标控件所有信息
    :return: True：成功获取并保存/False：不保存
    '''
    import tkinter.messagebox as msg
    import win32api, win32con
    import easygui

    try:
        autoType = cf.get_value("autoType")
        if autoType == "IE":
            # IE类型控件
            keyObj = eleProperties.get("Depth")
            info = keyObj[str(len(keyObj) - 1)]
            # print(info)
            copyInfo = info["Name"]
        elif autoType == "Windows":
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
        else:
            # chrome或firefox类型控件
            copyInfo = eleProperties


        message = "是否添加控件：\n" + str(copyInfo)
        # result = easygui.boolbox(msg=message, title='提示', choices=('是', '否'), image=None)                # rasygui，可用
        result = win32api.MessageBox(0, message, "提示", win32con.MB_OKCANCEL)                                # pywin32，可用
        # result = message_askyesno("提示", message)                                                        # tk下总会有空白/多余弹框，且易卡顿，不可用
        # print(result)

        path = cf.get_value("path")
        if result == 1:
            if autoType == "Windows":
                # windows控件保存前脚本进行唯一性校验
                # （True：可直接保存；False：人工进行index判断）
                AC = AppControl()
                flag = AC.checkBottom(keyObj, True)           # TODO：根据判断识别出结果但应用程序实际不可识别处理方法
                if not flag:
                    # win32api.MessageBox(0, "依据获取信息无法识别控件，请检查控件状态！", "提示", win32con.MB_OK)
                    # rasygui，可用
                    confirmMsg = "请检查控件当前状态，并确认是否继续保存？"
                    confirm = win32api.MessageBox(0, confirmMsg, "请确认", win32con.MB_OKCANCEL)
                    if confirm != 1:
                        return False

            # 点击后保存
            filePath = path + r"\log.txt"
            # print(filePath)
            with open(filePath, "a+"):
                pass
            with open(filePath, "r+") as f:
                rawData = f.read()
                if not rawData:
                    rawData = "{'%s': {}}" %autoType
                rawDict = eval(rawData)

                # objName = input("定义控件名称为：")
                objName = getInput("请确认", "请定义控件名称：")
                if objName is None:
                    # 点击后不保存，默认继续识别
                    return False

                if autoType not in rawDict:
                    rawDict[autoType] = dict()
                rawDict[autoType][objName] = dict()
                if autoType == "IE":
                    rawDict[autoType][objName]["xpath"] = copyInfo
                    rawDict[autoType][objName]["Starts"] = eleProperties["Starts"]
                    rawDict[autoType][objName]["Ends"] = eleProperties["Ends"]
                elif autoType == "Windows":
                    rawDict[autoType][objName] = eleProperties
                else:
                    # chrome/firefox
                    rawDict[autoType][objName]["xpath"] = copyInfo
                    rawDict[autoType][objName]["time"] = "%s %s" %(getCurrentDate(), getCurrentTime())
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
    recordIntoProject("test")