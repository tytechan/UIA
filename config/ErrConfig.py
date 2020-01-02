# _*_ coding:utf-8 _*_
import win32api
import win32con
from PIL import ImageGrab
from config.DirAndTime import *
import config.Globals as cf
import uiautomation as auto

class CNBMError(Exception):
    def __init__(self,ErrorInfo):
        # 初始化父类
        super().__init__(self)
        self.errorinfo = ErrorInfo

    def __str__(self):
        return ("\n%s" %self.errorinfo)

def CNBMException(func):
    def wrapper(*args, **kwargs):
        try:
            try:
                funcStr = "[%s %s] 执行语句 %s(" %(getCurrentDate(), getCurrentTime(), func.__name__)
                if args:
                    for count, string in enumerate(args):
                        if count == 0:
                            continue
                        if count == 1:
                            funcStr += "'%s'" %string if isinstance(string, str) else ", %s" %string
                        else:
                            funcStr += ", '%s'" %string if isinstance(string, str) else ", %s" %string
                if kwargs:
                    for key in kwargs.keys():
                        addStr = ", %s='%s'" %(key, kwargs[key])
                        funcStr += addStr
                funcStr += ")"
                auto.Logger.WriteLine(funcStr)

                return func(*args, **kwargs)
            except Exception as e:
                if "projectName" in cf._global_dict.keys():
                    projectName = cf.get_value("projectName")
                    picPath = os.path.dirname(os.path.dirname(os.path.abspath(__file__))) \
                              + "\\pictures\\%s" %projectName
                    dirName = createCurrentDateDir(picPath)
                    capture_screen(dirName)

                    # err = "\n[错误行] %s\n[报错文件] %s\n[错误信息] %s\n" \
                    #       % (e.__traceback__.tb_lineno,
                    #          e.__traceback__.tb_frame.f_globals["__file__"], e)
                    # auto.Logger.WriteLine(err, auto.ConsoleColor.Cyan, writeToFile=True)

                funcName = func.__name__
                err = cf.get_value("err")
                if projectName:
                    errInfo = "\n[工程名称] %s\n%s\n[关键字] %s\n[异常信息] %s\n" \
                              %(projectName, err, funcName, repr(e))
                else:
                    errInfo = "\n%s\n[关键字] %s\n[异常信息] %s" \
                              %(err, funcName, repr(e))
                auto.Logger.WriteLine(errInfo)
                raise CNBMError(errInfo)
        except CNBMError as err:
            raise err
    return wrapper

def capture_screen(picDir):
    ''' 截图
    :param picDir: 截图存放路径
    '''
    # 获取当前时间，精确到秒
    currentTime = getCurrentTime()
    # 拼接一场图片保存的绝对路径及名称
    picNameAndPath = str(picDir) + "\\" + str(currentTime) + ".png"
    try:
        ''' 方法一：部分截图 '''
        # im = ImageGrab.grab()
        # im.save(picNameAndPath.replace('\\',r'\\'))
        ''' 方法二：全屏截图 '''
        win32api.keybd_event(win32con.VK_SNAPSHOT, 0)
        time.sleep(0.5)
        im=ImageGrab.grabclipboard()
        im.save(picNameAndPath.replace('\\',r'\\'))
    except Exception as e:
        raise e
    else:
        return picNameAndPath

def handleErr(err):
    ''' 处理报错信息并写入日志文件
    :param err: 报错信息
    '''
    err = "[报错文件] %s\n[错误行] %s" \
              %(err.__traceback__.tb_frame.f_globals["__file__"],
                err.__traceback__.tb_lineno)
    # auto.Logger.WriteLine(err, auto.ConsoleColor.Cyan, writeToFile = True)
    cf.set_value("err", err)