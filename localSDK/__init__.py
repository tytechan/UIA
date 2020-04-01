import os

# 获取当前文件所在目录的绝对路径
parentDirPath = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# chromedriver存放路径，依据电脑配置
path = r"C:\Python35\chromedriver.exe"