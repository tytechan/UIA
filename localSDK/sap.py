# encoding = utf - 8

import time, datetime
import os
import win32api, win32con


# SAP功能函数函数
class CNBMError(Exception):
    def __init__(self, ErrorInfo):
        super().__init__(self)  # 初始化父类
        self.errorinfo = ErrorInfo

    def __str__(self):
        return self.errorinfo


def SAPException(func):
    def wrapper(*args, **kwargs):
        try:
            try:
                return func(*args, **kwargs)
            except Exception as e:
                # closeSAP()            # 关闭所有打开的sap进程
                # closeAllSession()       # 关闭本流程的所有sap进程
                funcName = func.__name__
                errInfo = "[关键字] " + funcName \
                          + " [异常信息] %s" % repr(e)
                print(errInfo)
                raise CNBMError(errInfo)
        except CNBMError as err:
            raise err

    return wrapper


@SAPException
def createObject(path, env):
    ''' 调起SAP服务
    :param path: SAP执行文件路径
    :param env: 本地登陆环境名
    :return: 全局变量session
    '''
    import subprocess
    import win32com.client
    global session, MS, connection

    subprocess.Popen(path)
    time.sleep(1)

    # 最长等待时间（须≥2）
    MS = 10

    loopTime = 3
    for i in range(loopTime):
        try:
            SapGuiAuto = win32com.client.GetObject('SAPGUI')
            print("SAP成功启动！")
            assert type(SapGuiAuto) == win32com.client.CDispatch
            break
        except AssertionError as e:
            return
        except Exception as e:
            if i + 1 < loopTime:
                continue
            else:
                raise e

    application = SapGuiAuto.GetScriptingEngine
    if not type(application) == win32com.client.CDispatch:
        SapGuiAuto = None
        return

    # help(application.OpenConnection)
    # Function openConnection(descriptionString As String, sync As Boolean False, raiseAsBoolean = True)

    # connection = application.OpenConnection(env, True)
    #
    # if not type(connection) == win32com.client.CDispatch:
    #     application = None
    #     SapGuiAuto = None
    #     return

    for i in range(loopTime):
        try:
            connection = application.OpenConnection(env, True)
            assert type(connection) == win32com.client.CDispatch
            break
        except AssertionError as e:
            application = None
            SapGuiAuto = None
            return
        except Exception as e:
            if i + 1 < loopTime:
                continue
            else:
                raise e

    for i in range(loopTime):
        try:
            session = connection.Children(0)
            assert type(session) == win32com.client.CDispatch
            break
        except AssertionError as e:
            connection = None
            application = None
            SapGuiAuto = None
            return
        except Exception as e:
            if i + 1 < loopTime:
                time.sleep(1)
                continue
            else:
                raise e


@SAPException
def saplogin(userName, passWord):
    ''' 登陆SAP '''
    global session, AW

    session.findById("wnd[0]/usr/txtRSYST-BNAME").text = userName
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = passWord
    session.findById("wnd[0]").sendVKey(0)

    try:
        # 等待“多次登陆”弹出框2s
        waitUntil(session.Children.count == 2, "False", maxSec="2|pass")
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
        updateActiveWindow()
        if AW.FindAllByName("wnd[1]", "GuiModalWindow").count == 1:
            AW.findById("usr/radMULTI_LOGON_OPT2").select()
            AW.findById("usr/radMULTI_LOGON_OPT2").setFocus()
            AW.findById("tbar[0]/btn[0]").press()
            # session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").select()
            # session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").setFocus()
            # session.findById("wnd[1]/tbar[0]/btn[0]").press()
    except:
        pass

    # 等待“信息”弹出框1s
    try:
        waitObj("name|wnd[1]|GuiModalWindow", "pass", "1|信息")
        btn = session.findById("wnd[1]/tbar[0]/btn[0]")
        btn.press()
    except:
        pass

    waitObj("name|wnd[0]|GuiMainWindow", "err", "SAP 轻松访问")
    print("SAP成功登陆！")


@SAPException
def updateActiveWindow():
    ''' 更新ActiveWindow，20190521 '''
    global AW, session
    AW = session.ActiveWindow
    # print("切换窗口成功！")
    return AW


def closeSAP():
    ''' 结束SAP（所有）进程 '''
    import psutil, os
    pids = psutil.pids()
    for pid in pids:
        p = psutil.Process(pid)
        # print('pid-%s,pname-%s' % (pid, p.name()))
        if p.name() == 'saplogon.exe':
            cmd = 'taskkill /F /IM saplogon.exe'
            # 无需结束sap进程时，注释下行
            # os.system(cmd)


@SAPException
def createNewSession():
    ''' 创建新session（窗口） '''
    global session, connection
    oc = connection.children.count
    session.createSession()  # session数+1
    sleep(1)
    session = connection.children(oc)
    for i in range(2):
        updateActiveWindow()
        try:
            waitObj("name|titl|GuiTitlebar", "pass", "SAP 轻松访问 中建信息|3")
            return
        except:
            pass
    updateActiveWindow()
    waitObj("name|titl|GuiTitlebar", "err", "SAP 轻松访问 中建信息|3")


def closeAllSession():
    global connection, session
    if 'connection' not in locals().keys():
        return

    sessions = connection.children
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
                    session = None
                    return
                except:
                    pass
            except:
                break


# **************************************** 基本操作 ****************************************
@SAPException
def performObj(performType, info, text=None):
    ''' 操作对象
    :param performType: 操作类型（输入/左击/勾选框/聚焦）
    :param info: 属性类型（id/name）|属性值
    :param text: 输入内容（performType为“输入”时）/对象text属性值（performType非“输入”时），选填
    '''
    global session, AW
    myInfo = info.split("|")

    if len(myInfo) == 4:
        # info: "name"|name|type|text
        obj = getObjsByNameAndText(myInfo[1] + "|" + myInfo[2], myInfo[3])
    elif performType in ["左击", "聚焦", "选择"] and text:
        obj = getObjsByNameAndText(info.split("|", 1)[1], text)
    else:
        obj = getObj(myInfo[0], info.split("|", 1)[1])

    if performType == "输入":
        obj.text = text
    elif performType == "下拉框":
        obj.key = text
    elif performType == "左击":
        obj.press()
    elif performType == "双击":
        obj.doubleClick()
    elif performType == "勾选框":
        # text = -1（勾选）/0（取消勾选）
        if obj.type in ["GuiCheckBox"]:
            # 示例type：GuiCheckBox
            obj.selected = int(text)
        elif obj.type in ["GuiShell"]:
            # 示例type：GuiShell
            obj.modifyCheckbox(0, "SEL", int(text))
            obj.triggerModified()
    elif performType == "聚焦":
        obj.setFocus()
    elif performType == "选择":
        obj.Select()
    elif performType == "模拟键盘":
        if text in keySimulated.keys():
            obj.sendVKey(keySimulated[text])
        else:
            obj.sendVKey(text)
    elif performType == "关闭":
        obj.close()


@SAPException
def performOnTable(performType, objInfo, inputInfo):
    ''' 根据同一类对象（同列）的不同index（不同行），操作table表格
    :param performType: 操作类型
    :param objInfo: name|type（表格类型为“GuiShell”时，后面还有“列信息”）
    :param inputInfo: text1%n1|text2%n2|...
                        (text为输入值；“%n”为选填项，对应表格index，从0开始)
    '''
    global AW
    II = inputInfo.split("|")
    OI = objInfo.split("|")
    # 默认从第一行开始
    initInfo = II[0].split("%")
    rowCount = 0 if len(initInfo) == 1 else int(initInfo[1])

    for i in range(len(II)):
        info = II[i].split("%")
        text = info[0]
        if i > 0:
            rowCount = (rowCount + 1) if len(info) == 1 else int(info[1])
        objs = AW.FindAllByName(OI[0], OI[1])

        if objs.count > 0 and objs[0].type == "GuiShell":
            # GuiShell类型表格
            obj = objs[0]
            if performType == "输入":
                obj.modifyCell(rowCount, OI[2], text)
        else:
            # 常规表格，通过index定位
            obj = objs[rowCount]
            if performType == "输入":
                obj.text = text
            elif performType == "下拉框":
                obj.key = text
            elif performType == "左击":
                obj.press()
            elif performType == "双击":
                obj.doubleClick()
            elif performType == "勾选框_A":
                # text = -1（勾选）/0（取消勾选）
                # 示例type：GuiCheckBox
                obj.selected = int(text)
            elif performType == "聚焦":
                obj.setFocus()
            elif performType == "选择":
                obj.Select()
            elif performType == "模拟键盘":
                obj.sendVKey(keySimulated[text])


@SAPException
def performToolbar(performType, info):
    ''' 操作type为GuiShell对象
    :param performType: 操作类型（提交/全部选择/清空选择/查看详情）
    :param info: 属性类型（id/name）|属性值
    '''
    global session, AW
    myInfo = info.split("|")

    if len(myInfo) == 4:
        # info: "name"|name|type|text
        obj = getObjsByNameAndText(myInfo[1] + "|" + myInfo[2], myInfo[3])
    else:
        obj = getObj(myInfo[0], info.split("|", 1)[1])

    if performType == "提交":
        obj.pressToolbarButton("ZPOST")
    elif performType == "全部选择":
        obj.pressToolbarButton("ZSALL")
    elif performType == "清空选择":
        obj.pressToolbarButton("ZDSAL")
    elif performType == "查看详情":
        obj.selectContextMenuItem("&DETAIL")


@SAPException
def getObj(case, info):
    ''' 获取对象
    :param case: 属性
    :param info: 属性值
    '''
    global AW
    info = info.split("|")
    typeInfo = case + "|" + info[0]
    typeInfo += ("|" + info[1]) if len(info) > 1 else ""
    waitObj(typeInfo, "err")

    if case == "id":
        obj = AW.findById(info[0])
        return obj
    elif case == "name":
        # obj = AW.FindByName(info[0], info[1])
        obj = AW.FindAllByName(info[0], info[1])[0]
        return obj


@SAPException
def getObjsByNameAndText(info, text):
    ''' 通过FindAllByName获取对象集合，再由text唯一定位
    :param info: 属性值（name|type）
    :param text: text属性值
    '''
    global AW
    typeInfo = "name|" + info
    waitObj(typeInfo, "err")

    info = info.split("|")
    objs = AW.FindAllByName(info[0], info[1])
    for i in range(objs.count):
        if objs[i].text == text:
            return objs[i]
        assert i + 1 == objs.count, \
            "该页面未找到name为“%s”、type为“%s”、text为“%s”的对象！" % (info[0], info[1], text)


@SAPException
def getText(case, info):
    ''' 获取对象text '''
    obj = getObj(case, info)
    return obj.text


@SAPException
def getNumInText(numInfo):
    import re
    num = re.findall('\d+', numInfo)
    return num[0] if len(num) == 1 else num[1]


# **************************************** 等待和校验 ****************************************
@SAPException
def waitObj(typeInfo, valueReturned, extraInfo=""):
    ''' 查找id，循环等待对象，20190522
    :param typeInfo: 属性（id/name）|属性值
    :param valueReturned: 未等到结果时结束方式（err：报错/pass：通过）
    :param extraInfo（选填）:可填写最长等待时间（默认为MS），及对象text属性（有值校验，不填不检验），顺序不限，用“|”隔开
    '''
    import re
    global AW, MS
    # 数据初始化
    updateActiveWindow()
    typeInfo = typeInfo.split("|")

    num = re.findall('\d+', extraInfo)
    num = num[0] if len(num) else ""
    maxSec = MS if not num else int(num)
    text = extraInfo.replace(num, "").replace("|", "")

    for i in range(maxSec):
        try:
            if typeInfo[0] == "id":
                obj = AW.findById(typeInfo[1])
            elif typeInfo[0] == "name":
                obj = AW.FindByName(typeInfo[1], typeInfo[2])

            # 校验text属性
            if text:
                # assert obj.text == text
                assert text in obj.text
            return
        except:
            try:
                assert maxSec != i + 1, \
                    str(maxSec) + " s内未找到 '" + typeInfo[0] + "' 为 '" + typeInfo[1] + "' 的对象！"
                time.sleep(1)
            except AssertionError as e:
                if valueReturned == "err":
                    raise AssertionError(e)
                elif valueReturned == "pass":
                    pass
            except Exception as e:
                raise e


@SAPException
def waitUntil(event, condition, maxSec):
    ''' 循环等待（event是否发生与condition一致时，等待，否则跳出），20190521
    :param event: 待判断主体（布尔类型）
    :param condition: True/False
    :param maxSec: 最长循环等待时间（≥2）|未等到结果（err：报错/pass：通过）
    '''
    ms = maxSec.split("|")
    maxSec = int(ms[0])
    result = ms[1]
    for i in range(maxSec):
        try:
            assert maxSec != i + 1, \
                str(maxSec) + " s内均未等到指定事件发生！"
            if bool(event) == eval(condition):
                time.sleep(1)
            else:
                return
        except AssertionError as e:
            if result == "err":
                raise AssertionError(e)
            elif result == "pass":
                pass
        except Exception as e:
            raise e


@SAPException
def checkText(textExpected, realText):
    ''' 校验
    :param textExpected: 预期值
    :param realText: 实际值
    '''
    assert textExpected == realText, \
        "预期值 " + textExpected + " 与实际值 " + realText + " 不符！"


if __name__ == "__main__":
    fp = "D:\SAP\SAPgui\saplogon.exe"
    env = "ET1"
    userName = "qianxin"
    password = "1234qwer"

    createObject(fp, env)
    saplogin(userName, password)

    # 进入交易
    updateActiveWindow()
    performObj("输入", "id|tbar[0]/okcd", "zfi234")
    performObj("左击", "id|tbar[0]/btn[0]")
    waitObj("name|wnd[0]|GuiMainWindow", "err", "1|销售开票及税务信息查询表")
    print("成功进入 zfi234 ！")

    # startDate = datetime.datetime.now().strftime("%Y.%m.01")
    startDate = datetime.datetime.now().strftime("01.%m.%Y")
    performObj("输入", "name|S_BUDAT-LOW|GuiCTextField", startDate)
    endDate = datetime.datetime.now().strftime("%d.%m.%Y")
    performObj("输入", "name|S_BUDAT-HIGH|GuiCTextField", endDate)
    performObj("输入", "name|S_BUKRS-LOW|GuiCTextField", "1000")
    performObj("左击", "id|wnd[0]/tbar[1]/btn[8]")
    waitObj("name|btn[43]|GuiButton", "err")
    print("成功进入查询结果界面！")