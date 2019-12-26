from config.DirAndTime import *

# 获取当前文件所在目录的绝对路径
parentDirPath = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# chromedriver存放路径，依据电脑配置
path = r"C:\Python35\chromedriver.exe"

# 创建日志路径（按日期）及 HookLog.txt
logPath = createCurrentDateDir(parentDirPath + "\\log\\")

if not os.path.exists(logPath + "\\HookLog.txt"):
    createTXT(logPath, "HookLog.txt")
hookLogPath = logPath + "\\HookLog.txt"