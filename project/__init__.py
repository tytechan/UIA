import os

# 获取当前文件所在目录的绝对路径
parentDirPath = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
# 获取当前工程名
projectName = os.getcwd().split("\\")[-1]